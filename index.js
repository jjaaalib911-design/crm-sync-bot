// ====================================================
// index.js — CRM Live Data Sync (overwrites sheet every 5 min)
// Deploy on Railway
// ====================================================

const puppeteer = require('puppeteer');
const { google } = require('googleapis');

const {
  CRM_USERNAME,
  CRM_PASSWORD,
  CRM_LOGIN_URL,
  CRM_LIST_PAGE_URL,
  CRM_DATA_API_URL,
  GOOGLE_SHEET_ID,
  GOOGLE_CREDENTIALS,
} = process.env;

const required = ['CRM_USERNAME', 'CRM_PASSWORD', 'CRM_LOGIN_URL', 'CRM_LIST_PAGE_URL', 'CRM_DATA_API_URL', 'GOOGLE_SHEET_ID', 'GOOGLE_CREDENTIALS'];
for (const v of required) {
  if (!process.env[v]) {
    console.error(`Missing environment variable: ${v}`);
    process.exit(1);
  }
}

let auth;
try {
  const creds = JSON.parse(GOOGLE_CREDENTIALS);
  auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
} catch (e) {
  console.error('Failed to parse GOOGLE_CREDENTIALS:', e.message);
  process.exit(1);
}
const sheets = google.sheets({ version: 'v4', auth });

const REMINDER_WINDOW_DAYS = 3;
const UNIQUE_KEY = 'Phone';

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
  'default': 0,
};
function getPriceForPackage(pkg) {
  if (!pkg) return PACKAGE_PRICES.default || 0;
  if (PACKAGE_PRICES[pkg] !== undefined) return PACKAGE_PRICES[pkg];
  const lower = pkg.toLowerCase().trim();
  for (const [key, value] of Object.entries(PACKAGE_PRICES)) {
    if (key.toLowerCase().trim() === lower) return value;
  }
  return PACKAGE_PRICES.default || 0;
}

function stripHtml(str) {
  if (!str) return '';
  return String(str).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
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

function extractRecord(row) {
  // Confirmed field positions from the real CRM API response:
  // 2=Username(code), 3=Name, 5=Phone, 6=Address, 7=Package, 14=Expiry, 21=Created
  const username = stripHtml(row[2]);
  const name = stripHtml(row[3]);
  const phone = stripHtml(row[5]);
  const address = stripHtml(row[6]);
  const pkg = stripHtml(row[7]);
  const activationDate = parseCrmDate(row[21]);
  const expiryDate = parseCrmDate(row[14]);
  return { username, name, phone, address, pkg, activationDate, expiryDate };
}

function buildRequestBody(start, length, filterType) {
  const orderableColumns = new Set([0, 2, 3, 4, 5, 6, 7]);
  const params = new URLSearchParams();
  params.append('draw', '1');
  for (let i = 0; i <= 22; i++) {
    params.append(`columns[${i}][data]`, String(i));
    params.append(`columns[${i}][name]`, '');
    params.append(`columns[${i}][searchable]`, 'true');
    params.append(`columns[${i}][orderable]`, orderableColumns.has(i) ? 'true' : 'false');
    params.append(`columns[${i}][search][value]`, '');
    params.append(`columns[${i}][search][regex]`, 'false');
  }
  params.append('order[0][column]', '0');
  params.append('order[0][dir]', 'desc');
  params.append('start', String(start));
  params.append('length', String(length));
  params.append('search[value]', '');
  params.append('search[regex]', 'false');
  params.append('filterType', String(filterType));
  params.append('dashboardUserTables', '1');
  return params.toString();
}

async function callDataApi(page, body) {
  return await page.evaluate(async (url, b) => {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-Requested-With': 'XMLHttpRequest',
      },
      body: b,
    });
    return res.json();
  }, CRM_DATA_API_URL, body);
}

// Try several filterType values and pick whichever returns the most records
async function findBestFilterType(page) {
  const candidates = [0, 1, 2, 3, 4, 5, ''];
  let best = { filterType: 3, recordsFiltered: 0, recordsTotal: 0 };

  for (const ft of candidates) {
    try {
      const body = buildRequestBody(0, 1, ft);
      const result = await callDataApi(page, body);
      const filtered = result.recordsFiltered || 0;
      const total = result.recordsTotal || 0;
      console.log(`  filterType=${JSON.stringify(ft)} -> recordsFiltered=${filtered}, recordsTotal=${total}`);
      if (filtered > best.recordsFiltered) {
        best = { filterType: ft, recordsFiltered: filtered, recordsTotal: total };
      }
    } catch (e) {
      console.log(`  filterType=${JSON.stringify(ft)} -> request failed: ${e.message}`);
    }
  }
  return best;
}

// ---------- OVERWRITE SHEET FUNCTION ----------
async function overwriteSheet(records) {
  // Define headers (must match your sheet columns)
  const HEADERS = [
    'ID No.', 'Name', 'Phone', 'Address', 'Package',
    'Amount Paid', 'Payment Date', 'Activation Date',
    'Expiry Date', 'Status', 'Days Remaining', 'Last Notified'
  ];

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Build rows for each record
  const rows = records.map(rec => {
    if (!rec.phone) return null;
    const price = getPriceForPackage(rec.pkg);
    const activationDateStr = formatDateForSheet(rec.activationDate);
    const paymentDateStr = activationDateStr; // Payment date = activation date
    const expiryDateStr = formatDateForSheet(rec.expiryDate);
    const { status, daysRemaining } = computeStatus(rec.expiryDate, today);
    return [
      rec.username,           // ID No.
      rec.name,               // Name
      rec.phone,              // Phone
      rec.address,            // Address
      rec.pkg,                // Package
      price,                  // Amount Paid
      paymentDateStr,         // Payment Date
      activationDateStr,      // Activation Date
      expiryDateStr,          // Expiry Date
      status,                 // Status
      daysRemaining,          // Days Remaining
      ''                      // Last Notified (blank)
    ];
  }).filter(row => row !== null);

  // Combine headers and rows
  const allData = [HEADERS, ...rows];

  // Write all data starting from A1 of the "Customers" sheet
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: 'Customers!A1',
    valueInputOption: 'USER_ENTERED',
    resource: { values: allData },
  });

  console.log(`✅ Sheet overwritten with ${rows.length} rows.`);
}

// ---------- Main sync ----------
async function syncCRM() {
  console.log(`[${new Date().toISOString()}] Sync started...`);
  let browser = null;

  try {
    browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    console.log('Logging in...');
    await page.goto(CRM_LOGIN_URL, { waitUntil: 'networkidle2', timeout: 30000 });
    await page.type('input[name="username"]', CRM_USERNAME, { delay: 30 });
    await page.type('input[name="password"]', CRM_PASSWORD, { delay: 30 });
    await Promise.all([
      page.click('button[type="submit"]'),
      page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }),
    ]);
    console.log('Login successful.');

    console.log('Opening customer list page...');
    await page.goto(CRM_LIST_PAGE_URL, { waitUntil: 'networkidle2', timeout: 30000 });

    console.log('Testing filter settings to find the one showing ALL customers...');
    const best = await findBestFilterType(page);
    console.log(`Best filterType found: ${JSON.stringify(best.filterType)} with ${best.recordsFiltered} of ${best.recordsTotal} total customers.`);

    const total = Math.max(best.recordsFiltered, best.recordsTotal, 1);
    const fullBody = buildRequestBody(0, total, best.filterType);
    const fullResult = await callDataApi(page, fullBody);

    const rawRows = fullResult.data || [];
    console.log(`Fetched ${rawRows.length} customer records.`);

    if (rawRows.length === 0) {
      throw new Error('No records returned even after testing filter types.');
    }

    // DEBUG: show the raw structure of the first record so we can verify field positions
    console.log('RAW first record (for column verification):');
    console.log(JSON.stringify(rawRows[0]));

    const records = rawRows.map(extractRecord);

    // Overwrite sheet with fresh data
    await overwriteSheet(records);

    console.log(`[${new Date().toISOString()}] Sync completed.`);

  } catch (error) {
    console.error('Sync failed:', error.message);
  } finally {
    if (browser) await browser.close();
  }
}

console.log('CRM Sync Bot starting...');
syncCRM();
setInterval(syncCRM, 5 * 60 * 1000);
