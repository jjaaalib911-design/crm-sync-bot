// ====================================================
// index.js — CRM Live Data Sync (with full pagination)
// Deploy on Railway — runs every 5 minutes, forever.
// ====================================================

const puppeteer = require('puppeteer');
const { google } = require('googleapis');

// ---------- Environment variables ----------
const {
  CRM_USERNAME,
  CRM_PASSWORD,
  CRM_LOGIN_URL,
  CRM_LIST_PAGE_URL,
  CRM_DATA_API_URL,
  GOOGLE_SHEET_ID,
  GOOGLE_CREDENTIALS,
} = process.env;

// ---------- Validate ----------
const required = [
  'CRM_USERNAME', 'CRM_PASSWORD', 'CRM_LOGIN_URL',
  'CRM_LIST_PAGE_URL', 'CRM_DATA_API_URL',
  'GOOGLE_SHEET_ID', 'GOOGLE_CREDENTIALS'
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
const PAGE_SIZE = 250;

// Required headers (exactly as they must appear in row 1)
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

function extractRecord(row) {
  const name = stripHtml(row[3]);
  const phone = stripHtml(row[5]);
  const address = stripHtml(row[6]);
  const pkg = stripHtml(row[7]);
  const activationDate = parseCrmDate(row[21]);
  const expiryDate = parseCrmDate(row[14]);
  return { name, phone, address, pkg, activationDate, expiryDate };
}

// ---------- Sheet reading with validation ----------
async function getExistingSheetData() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'A1:Z',
    });
    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log('ℹ️ Sheet is completely empty. Will create headers.');
      return { headers: [], data: [], phoneMap: new Map() };
    }
    const headers = rows[0];
    const data = rows.slice(1);
    
    const missing = REQUIRED_HEADERS.filter(h => !headers.includes(h));
    if (missing.length > 0) {
      console.error(`❌ Missing headers: ${missing.join(', ')}`);
      console.error(`⚠️ Your sheet has: ${headers.join(', ')}`);
      console.error(`✅ Required: ${REQUIRED_HEADERS.join(', ')}`);
      throw new Error('Header mismatch. Fix row 1 in your sheet.');
    }
    
    const phoneIndex = headers.indexOf(UNIQUE_KEY);
    const phoneMap = new Map();
    if (phoneIndex !== -1) {
      data.forEach((row, idx) => {
        const phone = (row[phoneIndex] || '').trim();
        if (phone) phoneMap.set(phone, idx);
      });
    }
    return { headers, data, phoneMap };
  } catch (err) {
    console.error('❌ Failed to read sheet:', err.message);
    throw err;
  }
}

// ---------- Sheet update ----------
async function updateSheet(records) {
  const { headers, data, phoneMap } = await getExistingSheetData();
  
  if (headers.length === 0) {
    console.log('📝 Creating header row...');
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'A1',
      valueInputOption: 'USER_ENTERED',
      resource: { values: [REQUIRED_HEADERS] },
    });
    const refreshed = await getExistingSheetData();
    return updateSheet(records);
  }

  const nameIdx = headers.indexOf('Name');
  const phoneIdx = headers.indexOf('Phone');
  const addressIdx = headers.indexOf('Address');
  const pkgIdx = headers.indexOf('Package');
  const activationIdx = headers.indexOf('Activation Date');
  const expiryIdx = headers.indexOf('Expiry Date');
  const statusIdx = headers.indexOf('Status');
  const daysIdx = headers.indexOf('Days Remaining');
  const lastNotifiedIdx = headers.indexOf('Last Notified');

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const updatedData = [...data];
  const newRows = [];
  let updatedCount = 0, appendedCount = 0;

  for (const rec of records) {
    if (!rec.phone) continue;

    const sheetRow = new Array(headers.length).fill('');
    if (nameIdx !== -1) sheetRow[nameIdx] = rec.name;
    if (phoneIdx !== -1) sheetRow[phoneIdx] = rec.phone;
    if (addressIdx !== -1) sheetRow[addressIdx] = rec.address;
    if (pkgIdx !== -1) sheetRow[pkgIdx] = rec.pkg;
    if (activationIdx !== -1) sheetRow[activationIdx] = formatDateForSheet(rec.activationDate);
    if (expiryIdx !== -1) sheetRow[expiryIdx] = formatDateForSheet(rec.expiryDate);

    const { status, daysRemaining } = computeStatus(rec.expiryDate, today);
    if (statusIdx !== -1) sheetRow[statusIdx] = status;
    if (daysIdx !== -1) sheetRow[daysIdx] = daysRemaining;

    const existingIndex = phoneMap.get(rec.phone);
    if (existingIndex !== undefined) {
      if (lastNotifiedIdx !== -1 && data[existingIndex] && data[existingIndex][lastNotifiedIdx]) {
        sheetRow[lastNotifiedIdx] = data[existingIndex][lastNotifiedIdx];
      }
      updatedData[existingIndex] = sheetRow;
      updatedCount++;
    } else {
      newRows.push(sheetRow);
      appendedCount++;
    }
  }

  const allRows = updatedData.concat(newRows);
  if (allRows.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `A2:Z${allRows.length + 1}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: allRows },
    });
    console.log(`✅ Sheet updated: ${updatedCount} updated, ${appendedCount} appended.`);
  } else {
    console.log('ℹ️ No changes to sheet.');
  }
}

// ---------- Build DataTables request body ----------
function buildRequestBody(start, length) {
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
  params.append('filterType', '3');
  params.append('dashboardUserTables', '1');
  return params.toString();
}

// ---------- Fetch a single page ----------
async function fetchPage(page, start, length) {
  const body = buildRequestBody(start, length);
  const result = await page.evaluate(async (url, body) => {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-Requested-With': 'XMLHttpRequest',
      },
      body,
    });
    return res.json();
  }, CRM_DATA_API_URL, body);
  return result.data || [];
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

    // --- Probe to get total count ---
    const probeBody = buildRequestBody(0, 1);
    const probeResult = await page.evaluate(async (url, body) => {
      const res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
          'X-Requested-With': 'XMLHttpRequest',
        },
        body,
      });
      return res.json();
    }, CRM_DATA_API_URL, probeBody);

    const total = probeResult.recordsTotal || 0;
    console.log(`📊 CRM reports ${total} total customers.`);

    if (total === 0) {
      throw new Error('No customers found.');
    }

    // --- Paginate: fetch all pages ---
    let allRows = [];
    let fetched = 0;
    while (fetched < total) {
      const length = Math.min(PAGE_SIZE, total - fetched);
      console.log(`📥 Fetching page from ${fetched} to ${fetched + length - 1}...`);
      const rows = await fetchPage(page, fetched, length);
      allRows = allRows.concat(rows);
      fetched += length;
      await new Promise(resolve => setTimeout(resolve, 200));
    }

    console.log(`📥 Fetched ${allRows.length} customer records total.`);

    if (allRows.length === 0) {
      throw new Error('No records returned.');
    }

    const records = allRows.map(extractRecord);
    await updateSheet(records);

    console.log(`[${new Date().toISOString()}] ✅ Sync completed.`);

  } catch (error) {
    console.error('❌ Sync failed:', error.message);
    if (error.stack) console.error(error.stack);
  } finally {
    if (browser) await browser.close();
  }
}

// ---------- Start ----------
console.log('🚀 CRM Sync Bot starting...');
syncCRM();
setInterval(syncCRM, 5 * 60 * 1000);
process.stdin.resume();
