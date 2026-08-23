console.log('🚀 CRM Sync Bot (MANUAL INDICES) starting...');

const puppeteer = require('puppeteer');
const { google } = require('googleapis');

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// ---------- Environment ----------
const {
  CRM_USERNAME,
  CRM_PASSWORD,
  CRM_LOGIN_URL,
  CRM_LIST_PAGE_URL,
  GOOGLE_SHEET_ID,
  GOOGLE_CREDENTIALS,
} = process.env;

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
  'ID No.', 'Name', 'Phone', 'Address', 'Package',
  'Amount Paid', 'Payment Date', 'Activation Date',
  'Expiry Date', 'Status', 'Days Remaining', 'Last Notified'
];

const PACKAGE_PRICES = {
  '7+7Mbps': 2200,
  '3+3Mbps': 1700,
  '4MB': 1500,
  '6MB': 1800,
  '5+5Mbps': 2000,
  '4+4Mbps': 1800,
  '10MB': 2300,
  '30 MB Day': 2500,
  '50 MB': 8500,
  'default': 0
};

// ---------- !! IMPORTANT: SET YOUR COLUMN INDICES HERE !! ----------
// Based on your table order: 
// 0=#ID, 1=Photo, 2=Username, 3=Full Name, 4=Phone, 5=Address, 6=Package, 7=Balance, 8=On/Off, 9=Expiry, 10=Created
const INDICES = {
  id: 2,       // Username → ID No.
  name: 3,     // Full Name → Name
  phone: 4,    // Phone → Phone
  address: 5,  // Address → Address
  pkg: 6,      // Package → Package
  expiry: 9,   // Expiry → Expiry Date
  created: 10  // Created → Activation Date
};
// -----------------------------------------------------------------

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
function getPriceForPackage(pkg) {
  if (!pkg) return PACKAGE_PRICES.default || 0;
  if (PACKAGE_PRICES[pkg] !== undefined) return PACKAGE_PRICES[pkg];
  const lower = pkg.toLowerCase();
  for (const [key, value] of Object.entries(PACKAGE_PRICES)) {
    if (key.toLowerCase() === lower) return value;
  }
  return PACKAGE_PRICES.default || 0;
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
    const price = getPriceForPackage(rec.pkg);
    const activationDateStr = formatDateForSheet(rec.activationDate);
    const paymentDateStr = activationDateStr;
    const row = [
      rec.id, rec.name, rec.phone, rec.address, rec.pkg,
      price, paymentDateStr, activationDateStr, formatDateForSheet(rec.expiryDate),
      '', '', ''
    ];
    const { status, daysRemaining } = computeStatus(rec.expiryDate, today);
    row[9] = status;
    row[10] = daysRemaining;
    row[11] = oldNotified.get(rec.phone) || '';
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
      protocolTimeout: 120000,
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    console.log('📍 Logging in...');
    await page.goto(CRM_LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.type('input[name="username"]', CRM_USERNAME, { delay: 30 });
    await page.type('input[name="password"]', CRM_PASSWORD, { delay: 30 });
    await Promise.all([
      page.click('button[type="submit"]'),
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 60000 }),
    ]);
    console.log('✅ Login successful.');

    console.log('📍 Opening customer list page...');
    await page.goto(CRM_LIST_PAGE_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForSelector('table', { timeout: 30000 });

    // ----- Set page size to 10 (safe) -----
    console.log('🔍 Setting page size to 10...');
    try {
      await page.evaluate(() => {
        const selects = document.querySelectorAll('select');
        for (const s of selects) {
          const opts = s.querySelectorAll('option');
          for (const opt of opts) {
            if (opt.value === '10' || opt.value === '25' || opt.value === '50') {
              s.value = opt.value;
              s.dispatchEvent(new Event('change', { bubbles: true }));
              return;
            }
          }
        }
      });
      await sleep(3000);
    } catch (e) {
      console.warn('⚠️ Could not set page size:', e.message);
    }

    // ----- Find the correct table and read headers -----
    console.log('🔍 Searching for table with Full Name, Address, Created...');
    const tableInfo = await page.evaluate(() => {
      const tables = document.querySelectorAll('table');
      let bestTable = null;
      let bestHeaders = [];
      let bestScore = -1;
      for (const table of tables) {
        const ths = table.querySelectorAll('tr:first-child th, tr:first-child td');
        const headers = Array.from(ths).map(cell => cell.innerText.trim());
        const keys = ['Full Name', 'Address', 'Created'];
        let score = 0;
        for (const k of keys) {
          if (headers.some(h => h.toLowerCase().includes(k.toLowerCase()))) score++;
        }
        if (score > bestScore) {
          bestScore = score;
          bestTable = table;
          bestHeaders = headers;
        }
      }
      if (!bestTable) return { headers: [], success: false };
      return { headers: bestHeaders, success: true };
    });

    if (!tableInfo.success || tableInfo.headers.length === 0) {
      throw new Error('Could not find table with Full Name, Address, Created.');
    }
    const headers = tableInfo.headers;
    console.log(`✅ Found table with headers: ${headers.join(' | ')}`);

    // ----- Paginate and scrape all rows from this table -----
    let allRows = [];
    let pageNum = 0;
    let maxPages = 100;

    while (pageNum < maxPages) {
      pageNum++;
      console.log(`📄 Scraping page ${pageNum}...`);

      const rows = await page.evaluate(() => {
        const tables = document.querySelectorAll('table');
        let bestTable = null;
        let bestScore = -1;
        for (const table of tables) {
          const ths = table.querySelectorAll('tr:first-child th, tr:first-child td');
          const headers = Array.from(ths).map(cell => cell.innerText.trim());
          const keys = ['Full Name', 'Address', 'Created'];
          let score = 0;
          for (const k of keys) {
            if (headers.some(h => h.toLowerCase().includes(k.toLowerCase()))) score++;
          }
          if (score > bestScore) {
            bestScore = score;
            bestTable = table;
          }
        }
        if (!bestTable) return [];
        const trs = bestTable.querySelectorAll('tr');
        const data = [];
        for (let i = 1; i < trs.length; i++) {
          const tds = trs[i].querySelectorAll('th, td');
          const rowData = Array.from(tds).map(cell => cell.innerText.trim());
          if (rowData.length > 0) data.push(rowData);
        }
        return data;
      });

      console.log(`   → Found ${rows.length} rows on page ${pageNum}`);
      allRows = allRows.concat(rows);

      // Check Next button
      const nextInfo = await page.evaluate(() => {
        const links = document.querySelectorAll('a, button');
        for (const el of links) {
          const text = (el.innerText || '').toLowerCase();
          const cls = el.className || '';
          if ((text.includes('next') || text.includes('>') || cls.includes('next')) && !el.disabled) {
            const disabled = el.disabled || el.getAttribute('aria-disabled') === 'true' || el.classList.contains('disabled');
            return { exists: true, disabled: !!disabled };
          }
        }
        return { exists: false, disabled: true };
      });

      if (!nextInfo.exists) {
        console.log('✅ No more pages (Next button not found).');
        break;
      }
      if (nextInfo.disabled) {
        console.log('✅ No more pages (Next button disabled).');
        break;
      }

      await page.evaluate(() => {
        const links = document.querySelectorAll('a, button');
        for (const el of links) {
          const text = (el.innerText || '').toLowerCase();
          const cls = el.className || '';
          if ((text.includes('next') || text.includes('>') || cls.includes('next')) && !el.disabled) {
            el.click();
            return;
          }
        }
      });

      console.log(`⏩ Clicked "Next" – loading page ${pageNum + 1}...`);
      await page.waitForSelector('table', { timeout: 30000 });
      await sleep(2000);
    }

    if (pageNum >= maxPages) {
      console.warn(`⚠️ Reached max page limit (${maxPages}). Stopping.`);
    }

    console.log(`📥 Scraped total of ${allRows.length} data rows.`);
    if (allRows.length === 0) throw new Error('No data rows found.');

    // ----- Show raw sample row to verify indices -----
    if (allRows.length > 0) {
      console.log('🔍 RAW sample row (first row):', allRows[0]);
      console.log('🔍 Headers for reference:', headers);
      console.log(`Using indices: ID=${INDICES.id}, Name=${INDICES.name}, Phone=${INDICES.phone}, Address=${INDICES.address}, Package=${INDICES.pkg}, Expiry=${INDICES.expiry}, Created=${INDICES.created}`);
    }

    // ----- Build records using the INDICES -----
    const records = allRows.map(row => {
      const id = stripHtml(row[INDICES.id] || '');
      const name = stripHtml(row[INDICES.name] || '');
      const phone = stripHtml(row[INDICES.phone] || '');
      const address = stripHtml(row[INDICES.address] || '');
      const pkg = stripHtml(row[INDICES.pkg] || '');
      const activationDate = parseCrmDate(row[INDICES.created] || '');
      const expiryDate = parseCrmDate(row[INDICES.expiry] || '');
      return { id, name, phone, address, pkg, activationDate, expiryDate };
    }).filter(r => r.phone);

    console.log(`✅ Filtered to ${records.length} valid customer records.`);

    // Show a sample mapped record
    if (records.length > 0) {
      console.log('🔍 MAPPED sample record:', records[0]);
    }

    await writeFreshSheet(records);

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
