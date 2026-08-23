console.log('🚀 CRM Sync Bot (FINAL: Correct mapping + auto price) starting...');

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

// ---------- Package → Price mapping (edit this to match your rates) ----------
const PACKAGE_PRICES = {
  '7+7Mbps': 2200,    // example – change to your actual prices
  '3+3Mbps': 1700,
  '4MB': 1500,
  '6MB': 1800,
  '5+5Mbps': 2000,
  '4+4Mbps': 1800,
  '10MB': 2300,
  '30 MB Day': 2500,
  '50 MB': 8500,
  // add any other packages you see
  'default': 0       // if package not found, price = 0
};

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
// Get price for a package
function getPriceForPackage(pkg) {
  if (!pkg) return PACKAGE_PRICES.default || 0;
  // Try exact match, then case‑insensitive
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
    // Amount paid from package
    const price = getPriceForPackage(rec.pkg);
    // Payment date = Activation date
    const activationDateStr = formatDateForSheet(rec.activationDate);
    const paymentDateStr = activationDateStr;  // same as activation
    // Build row: ID No., Name, Phone, Address, Package, Amount Paid, Payment Date, Activation Date, Expiry Date, Status, Days Remaining, Last Notified
    const row = [
      rec.id, rec.name, rec.phone, rec.address, rec.pkg,
      price, paymentDateStr, activationDateStr, formatDateForSheet(rec.expiryDate),
      '', '', ''
    ];
    const { status, daysRemaining } = computeStatus(rec.expiryDate, today);
    row[9] = status;      // Status
    row[10] = daysRemaining; // Days Remaining
    row[11] = oldNotified.get(rec.phone) || ''; // Last Notified
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

    // ----- Set "Show entries" to 1000 -----
    console.log('🔍 Setting dropdown to 1000...');
    const initialCount = await page.evaluate(() => document.querySelectorAll('table tr').length);
    const selected = await page.evaluate(() => {
      const selects = document.querySelectorAll('select');
      for (const s of selects) {
        const opts = s.querySelectorAll('option');
        for (const opt of opts) {
          if (opt.value === '1000' || opt.value === '-1' || opt.text.includes('1000') || opt.text.includes('All')) {
            s.value = opt.value;
            s.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          }
        }
      }
      return false;
    });
    if (selected) {
      console.log('✅ Selected 1000/All. Waiting for table update...');
      await page.waitForFunction(
        (initial) => document.querySelectorAll('table tr').length > initial,
        { timeout: 15000, args: [initialCount] }
      ).catch(() => console.warn('Row count did not increase.'));
      await sleep(2000);
    } else {
      console.warn('⚠️ Could not set dropdown, using default page.');
    }

    // ----- Scrape table -----
    console.log('📊 Scraping table...');
    const tableData = await page.evaluate(() => {
      const table = document.querySelector('table');
      if (!table) return [];
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

    if (tableData.length === 0) throw new Error('No table found.');

    // Remove header row if it contains typical headers
    let dataRows = tableData;
    if (dataRows[0] && dataRows[0].some(cell => cell.match(/Name|Phone|Username|ID|Package/))) {
      dataRows = dataRows.slice(1);
    }

    console.log(`📈 Scraped ${dataRows.length} rows (after removing header).`);

    // ----- Map using correct indices -----
    const records = dataRows.map(row => {
      const id = stripHtml(row[3] || '');      // Username -> ID No.
      const name = stripHtml(row[4] || '');    // Full Name
      const phone = stripHtml(row[6] || '');   // Phone
      const address = stripHtml(row[7] || ''); // Address
      const pkg = stripHtml(row[8] || '');     // Package
      const activationDate = parseCrmDate(row[22] || ''); // Created
      const expiryDate = parseCrmDate(row[15] || '');     // Expiry
      return { id, name, phone, address, pkg, activationDate, expiryDate };
    }).filter(r => r.phone);

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
