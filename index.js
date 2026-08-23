// ====================================================
// index.js — CRM Live Data Sync (using the real data API)
// Deploy on Railway — runs every 5 minutes, forever.
// ====================================================

const puppeteer = require('puppeteer');
const { google } = require('googleapis');

// ---------- Environment variables (set in Railway) ----------
const {
  CRM_USERNAME,
  CRM_PASSWORD,
  CRM_LOGIN_URL,
  CRM_LIST_PAGE_URL,   // e.g. http://223.123.38.98/user/all
  CRM_DATA_API_URL,    // http://223.123.38.98/admin_portal/user/user/getServerSideUsers
  GOOGLE_SHEET_ID,
  GOOGLE_CREDENTIALS,
} = process.env;

// ---------- Validate required variables ----------
const required = [
  'CRM_USERNAME',
  'CRM_PASSWORD',
  'CRM_LOGIN_URL',
  'CRM_LIST_PAGE_URL',
  'CRM_DATA_API_URL',
  'GOOGLE_SHEET_ID',
  'GOOGLE_CREDENTIALS'
];
for (const v of required) {
  if (!process.env[v]) {
    console.error(`❌ Missing environment variable: ${v}`);
    process.exit(1);
  }
}

// ---------- Google Sheets Authentication ----------
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

// ---------- Configuration ----------
const REMINDER_WINDOW_DAYS = 3;   // customers with expiry within 3 days get "Expiring Soon"
const UNIQUE_KEY = 'Phone';        // column used to match rows in the sheet

// ---------- Helper: clean HTML tags from CRM fields ----------
function stripHtml(str) {
  if (!str) return '';
  return str.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}

// Month abbreviation mapping for parsing CRM dates
const MONTHS = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };

// Parse dates like "18 Aug 2026 12:00:00" or "2026-07-18 20:30:19"
function parseCrmDate(text) {
  const clean = stripHtml(text);
  // Try "DD MMM YYYY" format
  const match = clean.match(/(\d{1,2})\s+(\w{3})\w*\s+(\d{4})/);
  if (match) {
    const [, day, monAbbr, year] = match;
    const month = MONTHS[monAbbr.slice(0, 3)];
    if (month !== undefined) return new Date(parseInt(year), month, parseInt(day));
  }
  // Fallback: "YYYY-MM-DD"
  const isoLike = clean.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (isoLike) {
    const [, y, m, d] = isoLike;
    return new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
  }
  return null;
}

// Format Date object as YYYY-MM-DD for the sheet
function formatDateForSheet(date) {
  if (!date) return '';
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

// Compute Status and Days Remaining based on expiry date
function computeStatus(expiryDate, today) {
  if (!expiryDate) return { status: '', daysRemaining: '' };
  const diffDays = Math.round((expiryDate - today) / (1000 * 60 * 60 * 24));
  let status;
  if (diffDays < 0) status = 'Expired';
  else if (diffDays <= REMINDER_WINDOW_DAYS) status = 'Expiring Soon';
  else status = 'Active';
  return { status, daysRemaining: diffDays };
}

// ---------- Extract a clean record from one CRM data row ----------
// The CRM returns an array of 23 fields (0‑based)
function extractRecord(row) {
  const name = stripHtml(row[3]);       // Full Name
  const phone = stripHtml(row[5]);       // Phone
  const address = stripHtml(row[6]);     // Address
  const pkg = stripHtml(row[7]);         // Package
  const activationDate = parseCrmDate(row[21]); // Created (used as Activation Date)
  const expiryDate = parseCrmDate(row[14]);      // Expiry
  return { name, phone, address, pkg, activationDate, expiryDate };
}

// ---------- Read existing sheet data (to preserve "Last Notified") ----------
async function getExistingSheetData() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'A1:Z',
    });
    const rows = response.data.values;
    if (!rows || rows.length === 0) return { headers: [], data: [], phoneMap: new Map() };
    const headers = rows[0];
    const data = rows.slice(1);
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
    return { headers: [], data: [], phoneMap: new Map() };
  }
}

// ---------- Update the Google Sheet with the fetched records ----------
async function updateSheet(records) {
  const { headers, data, phoneMap } = await getExistingSheetData();
  if (headers.length === 0) {
    console.error('❌ Sheet headers not found. Ensure the first row has the required columns.');
    return;
  }

  // Find column indices
  const nameIdx = headers.indexOf('Name');
  const phoneIdx = headers.indexOf('Phone');
  const addressIdx = headers.indexOf('Address');
  const pkgIdx = headers.indexOf('Package');
  const activationIdx = headers.indexOf('Activation Date');
  const expiryIdx = headers.indexOf('Expiry Date');
  const statusIdx = headers.indexOf('Status');
  const daysIdx = headers.indexOf('Days Remaining');
  const lastNotifiedIdx = headers.indexOf('Last Notified');

  if (phoneIdx === -1) {
    console.error('❌ Column "Phone" not found in sheet.');
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const updatedData = [...data];
  const newRows = [];
  let updatedCount = 0, appendedCount = 0;

  for (const rec of records) {
    if (!rec.phone) continue; // skip records without phone

    // Create a new row with the same length as headers
    const sheetRow = new Array(headers.length).fill('');
    if (nameIdx !== -1) sheetRow[nameIdx] = rec.name;
    if (phoneIdx !== -1) sheetRow[phoneIdx] = rec.phone;
    if (addressIdx !== -1) sheetRow[addressIdx] = rec.address;
    if (pkgIdx !== -1) sheetRow[pkgIdx] = rec.pkg;
    if (activationIdx !== -1) sheetRow[activationIdx] = formatDateForSheet(rec.activationDate);
    if (expiryIdx !== -1) sheetRow[expiryIdx] = formatDateForSheet(rec.expiryDate);

    // Compute status & days
    const { status, daysRemaining } = computeStatus(rec.expiryDate, today);
    if (statusIdx !== -1) sheetRow[statusIdx] = status;
    if (daysIdx !== -1) sheetRow[daysIdx] = daysRemaining;

    // Check if this phone already exists
    const existingIndex = phoneMap.get(rec.phone);
    if (existingIndex !== undefined) {
      // Preserve "Last Notified" from the existing row
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

  // Write back all rows
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

// ---------- Build the POST payload for DataTables (matches the CRM's expected format) ----------
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
  params.append('filterType', '3');          // '3' = show all customers (not just active)
  params.append('dashboardUserTables', '1');
  return params.toString();
}

// ---------- Main sync function ----------
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

    // ----- LOGIN -----
    console.log('📍 Logging in...');
    await page.goto(CRM_LOGIN_URL, { waitUntil: 'networkidle2', timeout: 30000 });
    await page.type('input[name="username"]', CRM_USERNAME, { delay: 30 });
    await page.type('input[name="password"]', CRM_PASSWORD, { delay: 30 });
    await Promise.all([
      page.click('button[type="submit"]'),
      page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }),
    ]);
    console.log('✅ Login successful.');

    // ----- Navigate to the customer list page to establish session -----
    console.log('📍 Opening customer list page...');
    await page.goto(CRM_LIST_PAGE_URL, { waitUntil: 'networkidle2', timeout: 30000 });

    // ----- Probe: get total record count (fetches just 1 record) -----
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

    const total = probeResult.recordsTotal || 1000;
    console.log(`📊 CRM reports ${total} total customers.`);

    // ----- Fetch ALL records in one request -----
    const fullBody = buildRequestBody(0, total);
    const fullResult = await page.evaluate(async (url, body) => {
      const res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
          'X-Requested-With': 'XMLHttpRequest',
        },
        body,
      });
      return res.json();
    }, CRM_DATA_API_URL, fullBody);

    const rawRows = fullResult.data || [];
    console.log(`📥 Fetched ${rawRows.length} customer records.`);

    if (rawRows.length === 0) {
      throw new Error('No records returned. Check CRM_DATA_API_URL and filterType.');
    }

    // ----- Convert raw rows to clean objects and update sheet -----
    const records = rawRows.map(extractRecord);
    await updateSheet(records);

    console.log(`[${new Date().toISOString()}] ✅ Sync completed.`);

  } catch (error) {
    console.error('❌ Sync failed:', error.message);
    if (error.stack) console.error(error.stack);
  } finally {
    if (browser) await browser.close();
  }
}

// ---------- Start the bot: run once, then every 5 minutes ----------
console.log('🚀 CRM Sync Bot starting...');
syncCRM();
setInterval(syncCRM, 5 * 60 * 1000);

// Keep the process alive
process.stdin.resume();
