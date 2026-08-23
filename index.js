console.log('🚀 CRM Sync Bot (EXPORT URL EXTRACTOR) starting...');

const puppeteer = require('puppeteer');
const { google } = require('googleapis');
const { parse } = require('csv-parse/sync');

// ---------- Environment ----------
const {
  CRM_USERNAME,
  CRM_PASSWORD,
  CRM_LOGIN_URL,
  CRM_LIST_PAGE_URL,
  GOOGLE_SHEET_ID,
  GOOGLE_CREDENTIALS,
} = process.env;

// ---------- Validate ----------
const required = [
  'CRM_USERNAME', 'CRM_PASSWORD', 'CRM_LOGIN_URL',
  'CRM_LIST_PAGE_URL', 'GOOGLE_SHEET_ID', 'GOOGLE_CREDENTIALS'
];
for (const v of required) {
  if (!process.env[v]) {
    console.error(`❌ Missing: ${v}`);
    process.exit(1);
  }
}

// ---------- Google Auth ----------
let auth;
try {
  const creds = JSON.parse(GOOGLE_CREDENTIALS);
  auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
} catch (e) {
  console.error('❌ Failed to parse GOOGLE_CREDENTIALS:', e.message);
  process.exit(1);
}
const sheets = google.sheets({ version: 'v4', auth });

// ---------- Config ----------
const REMINDER_WINDOW_DAYS = 3;
const UNIQUE_KEY = 'Phone';
const REQUIRED_HEADERS = [
  'Name', 'Phone', 'Address', 'Package',
  'Amount Paid', 'Payment Date', 'Activation Date',
  'Expiry Date', 'Status', 'Days Remaining', 'Last Notified'
];

// ---------- Helpers ----------
function stripHtml(str) {
  if (!str) return '';
  return str.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}

const MONTHS = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };

function parseCrmDate(text) {
  const clean = stripHtml(text);
  const match = clean.match(/(\d{1,2})\s+(\w{3})\w*\s+(\d{4})/);
  if (match) {
    const [, day, monAbbr, year] = match;
    const month = MONTHS[monAbbr.slice(0, 3)];
    if (month !== undefined) return new Date(parseInt(year), month, parseInt(day));
  }
  const isoLike = clean.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (isoLike) {
    const [, y, m, d] = isoLike;
    return new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
  }
  return null;
}

function formatDateForSheet(date) {
  if (!date) return '';
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function computeStatus(expiryDate, today) {
  if (!expiryDate) return { status: '', daysRemaining: '' };
  const diffDays = Math.round((expiryDate - today) / (1000 * 60 * 60 * 24));
  let status;
  if (diffDays < 0) status = 'Expired';
  else if (diffDays <= REMINDER_WINDOW_DAYS) status = 'Expiring Soon';
  else status = 'Active';
  return { status, daysRemaining: diffDays };
}

// ---------- Sheet functions ----------
async function getTargetSheetName() {
  try {
    const meta = await sheets.spreadsheets.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      fields: 'sheets.properties',
    });
    const sheetsList = meta.data.sheets;
    if (!sheetsList || sheetsList.length === 0) {
      console.log('📝 No sheets – creating "Sheet1"');
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: GOOGLE_SHEET_ID,
        resource: {
          requests: [{ addSheet: { properties: { title: 'Sheet1' } } }]
        }
      });
      return 'Sheet1';
    }
    const firstSheet = sheetsList[0];
    const name = firstSheet.properties.title;
    console.log(`📌 Writing to sheet: "${name}"`);
    return name;
  } catch (e) {
    console.error('❌ Failed to get sheet info:', e.message);
    return 'Sheet1';
  }
}

async function getExistingLastNotified(sheetName) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${sheetName}!A:Z`,
    });
    const rows = response.data.values;
    if (!rows || rows.length < 2) return new Map();
    const headers = rows[0];
    const phoneIdx = headers.indexOf(UNIQUE_KEY);
    const lastNotifiedIdx = headers.indexOf('Last Notified');
    if (phoneIdx === -1 || lastNotifiedIdx === -1) return new Map();
    const map = new Map();
    for (let i = 1; i < rows.length; i++) {
      const phone = (rows[i][phoneIdx] || '').trim();
      const notified = rows[i][lastNotifiedIdx] || '';
      if (phone) map.set(phone, notified);
    }
    return map;
  } catch (e) {
    console.error('❌ Failed to read Last Notified:', e.message);
    return new Map();
  }
}

async function writeFreshSheet(records) {
  const sheetName = await getTargetSheetName();
  const oldNotified = await getExistingLastNotified(sheetName);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const newRows = [];
  for (const rec of records) {
    if (!rec.phone) continue;
    const row = [
      rec.name,
      rec.phone,
      rec.address,
      rec.pkg,
      '',
      '',
      formatDateForSheet(rec.activationDate),
      formatDateForSheet(rec.expiryDate),
      '',
      '',
      ''
    ];
    const { status, daysRemaining } = computeStatus(rec.expiryDate, today);
    row[8] = status;
    row[9] = daysRemaining;
    row[10] = oldNotified.get(rec.phone) || '';
    newRows.push(row);
  }

  const allData = [REQUIRED_HEADERS, ...newRows];
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: allData },
  });
  console.log(`✅ Sheet overwritten with ${newRows.length} rows (${oldNotified.size} Last Notified preserved).`);
}

// ---------- Main sync ----------
async function syncCRM() {
  console.log(`[${new Date().toISOString()}] 🔄 Sync started...`);
  let browser = null;

  try {
    browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    console.log('📍 Logging in...');
    await page.goto(CRM_LOGIN_URL, { waitUntil: 'networkidle2', timeout: 30000 });
    await page.type('input[name="username"]', CRM_USERNAME, { delay: 30 });
    await page.type('input[name="password"]', CRM_PASSWORD, { delay: 30 });
    await Promise.all([
      page.click('button[type="submit"]'),
      page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }),
    ]);
    console.log('✅ Login successful.');

    console.log('📍 Opening customer list page...');
    await page.goto(CRM_LIST_PAGE_URL, { waitUntil: 'networkidle2', timeout: 30000 });

    // ----- Find export URL -----
    console.log('🔍 Searching for export URL...');
    const exportUrl = await page.evaluate(() => {
      // Look for links with export/csv/download in href or text
      const elements = document.querySelectorAll('a, button, input[type="button"], input[type="submit"]');
      for (const el of elements) {
        const href = el.href || '';
        const text = (el.innerText || '').toLowerCase();
        const onclick = el.getAttribute('onclick') || '';
        const dataUrl = el.getAttribute('data-url') || '';
        const className = el.className || '';
        const id = el.id || '';
        if (href.includes('export') || href.includes('csv') || href.includes('download') ||
            text.includes('export') || text.includes('csv') || text.includes('download') ||
            onclick.includes('export') || onclick.includes('csv') ||
            dataUrl.includes('export') || dataUrl.includes('csv') ||
            className.includes('export') || id.includes('export')) {
          // If it's a link, return href
          if (href && (href.includes('export') || href.includes('csv') || href.includes('download'))) {
            return href;
          }
          // If it's a button with onclick containing window.location
          if (onclick) {
            const match = onclick.match(/window\.location\.href\s*=\s*['"]([^'"]+)['"]/);
            if (match) return match[1];
            const match2 = onclick.match(/location\.href\s*=\s*['"]([^'"]+)['"]/);
            if (match2) return match2[1];
          }
          // If it has data-url
          if (dataUrl) return dataUrl;
          // If it's a button with no URL, we'll click it as fallback
          // but we return null to indicate we need to click
          return null;
        }
      }
      return null;
    });

    let csvContent = '';
    if (exportUrl) {
      console.log(`✅ Found export URL: ${exportUrl}`);
      // Navigate to it – should load CSV directly
      await page.goto(exportUrl, { waitUntil: 'networkidle2', timeout: 30000 });
      // Get page content (should be CSV)
      csvContent = await page.evaluate(() => document.body.innerText);
      console.log(`📄 Downloaded CSV via URL (${csvContent.length} characters)`);
    } else {
      // Fallback: try to click the button and intercept response
      console.log('⚠️ No export URL found. Trying to click export button and intercept...');
      const [response] = await Promise.all([
        page.waitForResponse(res => 
          res.url().includes('export') || 
          res.url().includes('csv') || 
          res.url().includes('download') ||
          res.headers()['content-type']?.includes('csv') ||
          res.headers()['content-disposition']?.includes('attachment')
        ),
        page.evaluate(() => {
          const btn = document.querySelector('a[href*="export"], button[onclick*="export"], input[value*="Export"]');
          if (btn) btn.click();
        })
      ]);
      if (response) {
        const buffer = await response.buffer();
        csvContent = buffer.toString('utf8');
        console.log(`📄 Captured CSV from network (${csvContent.length} characters)`);
      } else {
        throw new Error('Could not capture CSV export.');
      }
    }

    if (!csvContent || csvContent.length < 50) {
      throw new Error('CSV content is empty or too short.');
    }

    // ---- Validate that it's actually CSV (has commas, not HTML) ----
    if (csvContent.includes('<html') || csvContent.includes('<!DOCTYPE')) {
      console.warn('⚠️ Got HTML instead of CSV. Trying to extract table from HTML...');
      // Could attempt to scrape table from HTML as fallback, but better to error.
      throw new Error('Export returned HTML instead of CSV – the export URL may require specific parameters.');
    }

    // ----- Parse CSV -----
    const records = parse(csvContent, {
      columns: true,
      skip_empty_lines: true,
      trim: true,
    });
    console.log(`📈 Found ${records.length} customer records.`);

    // Map to our structure
    const mapped = records.map(row => {
      const name = row['Full Name'] || row['Name'] || '';
      const phone = row['Phone'] || '';
      const address = row['Address'] || '';
      const pkg = row['Package'] || '';
      const activationDate = parseCrmDate(row['Created'] || row['Activation Date'] || '');
      const expiryDate = parseCrmDate(row['Expiry'] || row['Expiry Date'] || '');
      return { name, phone, address, pkg, activationDate, expiryDate };
    });

    await writeFreshSheet(mapped);

    console.log(`[${new Date().toISOString()}] ✅ Sync completed.`);

  } catch (error) {
    console.error('❌ Sync failed:', error.message);
    if (error.stack) console.error(error.stack);
  } finally {
    if (browser) await browser.close();
  }
}

// ---------- Start ----------
syncCRM();
setInterval(syncCRM, 5 * 60 * 1000);
process.stdin.resume();
