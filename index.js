console.log('🚀 CRM Sync Bot (DYNAMIC HEADER MAPPING) starting...');

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
  'ID No.', 'Name', 'Phone', 'Address', 'Package',
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
      rec.id, rec.name, rec.phone, rec.address, rec.pkg,
      '', '', formatDateForSheet(rec.activationDate), formatDateForSheet(rec.expiryDate),
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

    // ----- Select dropdown -----
    console.log('🔍 Searching for "Show entries" dropdown...');
    const initialRowCount = await page.evaluate(() => document.querySelectorAll('table tr').length);
    console.log(`Initial row count: ${initialRowCount}`);

    const selectSelector = await page.evaluate(() => {
      const selectors = [
        'select[name="example_length"]',
        'select[name="DataTables_Table_0_length"]',
        'select[name="user_list_length"]',
        'select[aria-controls*="DataTables"]',
        'select'
      ];
      for (const sel of selectors) {
        const el = document.querySelector(sel);
        if (el && el.tagName === 'SELECT') {
          const opts = el.querySelectorAll('option');
          for (const opt of opts) {
            if (opt.value === '1000' || opt.value === '-1' || opt.text.includes('1000') || opt.text.includes('All')) {
              return sel;
            }
          }
        }
      }
      return null;
    });

    if (selectSelector) {
      console.log(`✅ Found select: ${selectSelector}`);
      try {
        await page.select(selectSelector, '1000', '-1');
        console.log('✅ Selected value 1000 or -1');
      } catch (e) {
        await page.select(selectSelector, '1000');
      }
      await page.waitForFunction(
        (initial) => document.querySelectorAll('table tr').length > initial,
        { timeout: 15000, args: [initialRowCount] }
      ).then(() => console.log('✅ Table row count increased.')).catch(() => console.warn('⚠️ Row count did not increase.'));
      await sleep(2000);
    } else {
      console.warn('⚠️ Could not find select dropdown. Will use pagination fallback.');
    }

    // ----- Scrape table with headers -----
    console.log('📊 Scraping table with headers...');
    const { headers, rows } = await page.evaluate(() => {
      const table = document.querySelector('table');
      if (!table) return { headers: [], rows: [] };
      const allRows = table.querySelectorAll('tr');
      if (allRows.length === 0) return { headers: [], rows: [] };
      // First row as headers
      const headerCells = allRows[0].querySelectorAll('th, td');
      const headers = Array.from(headerCells).map(cell => cell.innerText.trim());
      // Data rows
      const dataRows = [];
      for (let i = 1; i < allRows.length; i++) {
        const cells = allRows[i].querySelectorAll('th, td');
        const rowData = Array.from(cells).map(cell => cell.innerText.trim());
        if (rowData.length > 0) dataRows.push(rowData);
      }
      return { headers, rows: dataRows };
    });

    if (rows.length === 0) throw new Error('No data rows found.');

    console.log(`📈 Scraped ${rows.length} data rows.`);

    // ----- Dynamic column mapping by header text -----
    const headerMap = {};
    headers.forEach((h, idx) => {
      const key = h.toLowerCase().replace(/[^a-z0-9]/g, '');
      headerMap[key] = idx;
    });

    // Map our target fields to header names
    const fieldMapping = {
      id: ['#id', 'id', 'userid', 'customerid'],
      name: ['fullname', 'name', 'full name', 'customer name'],
      phone: ['phone', 'mobile', 'contact'],
      address: ['address', 'location'],
      pkg: ['package', 'plan', 'bandwidth'],
      activationDate: ['created', 'creationdate', 'createddate', 'activationdate', 'activation'],
      expiryDate: ['expiry', 'expiration', 'expirydate', 'validuntil']
    };

    function findIndex(field) {
      const possible = fieldMapping[field];
      for (const p of possible) {
        const key = p.replace(/[^a-z0-9]/g, '');
        if (headerMap[key] !== undefined) return headerMap[key];
      }
      return -1;
    }

    const idIdx = findIndex('id');
    const nameIdx = findIndex('name');
    const phoneIdx = findIndex('phone');
    const addressIdx = findIndex('address');
    const pkgIdx = findIndex('pkg');
    const activationIdx = findIndex('activationDate');
    const expiryIdx = findIndex('expiryDate');

    console.log(`Mapped indices: ID=${idIdx}, Name=${nameIdx}, Phone=${phoneIdx}, Address=${addressIdx}, Package=${pkgIdx}, Activation=${activationIdx}, Expiry=${expiryIdx}`);

    // If any required index is missing, fallback to fixed indices from CSV structure
    if (idIdx === -1 || nameIdx === -1 || phoneIdx === -1 || addressIdx === -1 || pkgIdx === -1 || activationIdx === -1 || expiryIdx === -1) {
      console.warn('⚠️ Some headers not found – falling back to fixed indices (based on CSV order).');
      // Fallback: assume order from your CSV: 0=ID, 4=FullName, 6=Phone, 7=Address, 8=Package, 22=Created, 15=Expiry
      // But we'll use the indices that were found if any
      const fallback = {
        id: 0,
        name: 4,
        phone: 6,
        address: 7,
        pkg: 8,
        activation: 22,
        expiry: 15
      };
      // Override only if not found
      const finalIdIdx = idIdx !== -1 ? idIdx : fallback.id;
      const finalNameIdx = nameIdx !== -1 ? nameIdx : fallback.name;
      const finalPhoneIdx = phoneIdx !== -1 ? phoneIdx : fallback.phone;
      const finalAddressIdx = addressIdx !== -1 ? addressIdx : fallback.address;
      const finalPkgIdx = pkgIdx !== -1 ? pkgIdx : fallback.pkg;
      const finalActivationIdx = activationIdx !== -1 ? activationIdx : fallback.activation;
      const finalExpiryIdx = expiryIdx !== -1 ? expiryIdx : fallback.expiry;

      // Re-map
      const records = rows.map(row => {
        const id = stripHtml(row[finalIdIdx] || '');
        const name = stripHtml(row[finalNameIdx] || '');
        const phone = stripHtml(row[finalPhoneIdx] || '');
        const address = stripHtml(row[finalAddressIdx] || '');
        const pkg = stripHtml(row[finalPkgIdx] || '');
        const activationDate = parseCrmDate(row[finalActivationIdx] || '');
        const expiryDate = parseCrmDate(row[finalExpiryIdx] || '');
        return { id, name, phone, address, pkg, activationDate, expiryDate };
      }).filter(r => r.phone);

      console.log(`✅ Using fallback mapping – ${records.length} records.`);
      await writeFreshSheet(records);
      console.log(`[${new Date().toISOString()}] ✅ Sync completed.`);
      return;
    }

    // ----- Use dynamic mapping -----
    const records = rows.map(row => {
      const id = stripHtml(row[idIdx] || '');
      const name = stripHtml(row[nameIdx] || '');
      const phone = stripHtml(row[phoneIdx] || '');
      const address = stripHtml(row[addressIdx] || '');
      const pkg = stripHtml(row[pkgIdx] || '');
      const activationDate = parseCrmDate(row[activationIdx] || '');
      const expiryDate = parseCrmDate(row[expiryIdx] || '');
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
