console.log('🚀 CRM Sync Bot (TABLE SCRAPER - SHOW ALL) starting...');

const puppeteer = require('puppeteer');
const { google } = require('googleapis');

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
      rec.name, rec.phone, rec.address, rec.pkg,
      '', '', formatDateForSheet(rec.activationDate), formatDateForSheet(rec.expiryDate),
      '', '', ''
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

    // ----- Select "1000" from the "Show entries" dropdown -----
    console.log('🔍 Selecting "1000" entries from dropdown...');
    try {
      // Common selectors for DataTables "Show entries" dropdown
      const dropdownSelectors = [
        'select[name="example_length"]',
        'select[name="DataTables_Table_0_length"]',
        'select[name="user_list_length"]',
        'select[aria-controls*="DataTables"]',
        'select:has(option[value="1000"])',
        'select:has(option[value="-1"])'
      ];
      let selected = false;
      for (const selector of dropdownSelectors) {
        const dropdown = await page.$(selector);
        if (dropdown) {
          await dropdown.select('1000', '-1');
          selected = true;
          console.log(`✅ Selected "1000" using selector: ${selector}`);
          // Wait for table to reload with more rows
          await page.waitForTimeout(3000);
          break;
        }
      }
      if (!selected) {
        // Fallback: try to find any select with options containing "1000"
        await page.evaluate(() => {
          const selects = document.querySelectorAll('select');
          for (const s of selects) {
            const opts = s.querySelectorAll('option');
            for (const opt of opts) {
              if (opt.value === '1000' || opt.value === '-1' || opt.text.includes('1000') || opt.text.includes('All')) {
                s.value = opt.value;
                s.dispatchEvent(new Event('change', { bubbles: true }));
                return;
              }
            }
          }
        });
        console.log('✅ Attempted to select "1000" via fallback.');
        await page.waitForTimeout(3000);
      }
    } catch (e) {
      console.warn('⚠️ Could not set entries count – scraping default page size.');
    }

    // ----- Scrape the table -----
    console.log('📊 Scraping table...');
    const tableData = await page.evaluate(() => {
      const tables = document.querySelectorAll('table');
      if (tables.length === 0) return [];
      // Assume first table is the main data table
      const table = tables[0];
      const rows = table.querySelectorAll('tr');
      const data = [];
      rows.forEach(row => {
        const cells = row.querySelectorAll('th, td');
        const rowData = [];
        cells.forEach(cell => rowData.push(cell.innerText.trim()));
        if (rowData.length > 0) data.push(rowData);
      });
      return data;
    });

    if (tableData.length === 0) {
      throw new Error('No table data scraped.');
    }

    // Remove header if present (first row)
    let dataRows = tableData;
    if (dataRows.length > 0 && dataRows[0].some(cell => cell.match(/Name|Phone|Username|ID|Package/))) {
      dataRows = dataRows.slice(1);
    }

    console.log(`📈 Scraped ${dataRows.length} customer rows.`);

    if (dataRows.length === 0) {
      throw new Error('No data rows found.');
    }

    // ----- Map columns to our fields -----
    // Based on your CSV: columns: #ID, Photo, Username, Full Name, NID, Phone, Address, Package, Seller, Balance, Service, On/Off, On/Off Time, Usage Data/Time, Expiry, IP/MAC, SMS, Policy, Pool, NAS, Discount, Created, Action
    // We need Name (index 3), Phone (index 5), Address (index 6), Package (index 7), Activation Date (index 21), Expiry (index 14)
    const records = dataRows.map(row => {
      const name = stripHtml(row[3] || '');
      const phone = stripHtml(row[5] || '');
      const address = stripHtml(row[6] || '');
      const pkg = stripHtml(row[7] || '');
      const activationDate = parseCrmDate(row[21] || '');
      const expiryDate = parseCrmDate(row[14] || '');
      return { name, phone, address, pkg, activationDate, expiryDate };
    }).filter(r => r.phone); // remove rows without phone

    console.log(`✅ Filtered to ${records.length} valid customer records.`);

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
