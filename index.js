console.log('🚀 CRM Sync Bot (PAGINATION – no dropdown) starting...');

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

// ---------- Package prices ----------
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

// ---------- Main sync with pagination ----------
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

    // ----- Scrape all pages by clicking "Next" -----
    let allRows = [];
    let pageNum = 0;
    let hasNext = true;
    let headers = [];

    while (hasNext) {
      pageNum++;
      console.log(`📄 Scraping page ${pageNum}...`);

      // Extract header only on first page
      if (pageNum === 1) {
        headers = await page.evaluate(() => {
          const table = document.querySelector('table');
          if (!table) return [];
          const ths = table.querySelectorAll('tr:first-child th, tr:first-child td');
          return Array.from(ths).map(cell => cell.innerText.trim());
        });
        console.log(`🔍 Headers: ${headers.join(' | ')}`);
      }

      // Extract data rows
      const rows = await page.evaluate(() => {
        const table = document.querySelector('table');
        if (!table) return [];
        const trs = table.querySelectorAll('tr');
        const dataRows = [];
        for (let i = 1; i < trs.length; i++) {
          const tds = trs[i].querySelectorAll('th, td');
          const rowData = Array.from(tds).map(cell => cell.innerText.trim());
          if (rowData.length > 0) dataRows.push(rowData);
        }
        return dataRows;
      });

      console.log(`   → Found ${rows.length} rows on page ${pageNum}`);
      allRows = allRows.concat(rows);

      // Try to click "Next" button
      const nextClicked = await page.evaluate(() => {
        const links = document.querySelectorAll('a, button');
        for (const el of links) {
          const text = (el.innerText || '').toLowerCase();
          const cls = el.className || '';
          if ((text.includes('next') || text.includes('>') || cls.includes('next')) && !el.disabled) {
            if (el.getAttribute('aria-disabled') === 'true') return false;
            el.click();
            return true;
          }
        }
        return false;
      });

      if (!nextClicked) {
        console.log('✅ No more pages.');
        hasNext = false;
      } else {
        console.log(`⏩ Clicked "Next" – loading page ${pageNum + 1}...`);
        await page.waitForSelector('table', { timeout: 30000 });
        await sleep(2000);
      }
    }

    console.log(`📥 Scraped total of ${allRows.length} data rows.`);
    if (allRows.length === 0) throw new Error('No data rows found.');

    // ----- Map columns (using headers) -----
    let idIdx, nameIdx, phoneIdx, addressIdx, pkgIdx, expiryIdx, createdIdx;
    if (headers.length > 0) {
      function findIndex(keywords) {
        for (const kw of keywords) {
          const idx = headers.findIndex(h => h.toLowerCase().includes(kw.toLowerCase()));
          if (idx !== -1) return idx;
        }
        return -1;
      }
      idIdx = findIndex(['username', 'user id', 'id']);
      nameIdx = findIndex(['full name', 'fullname', 'name']);
      phoneIdx = findIndex(['phone', 'mobile', 'contact']);
      addressIdx = findIndex(['address', 'location']);
      pkgIdx = findIndex(['package', 'plan', 'bandwidth']);
      expiryIdx = findIndex(['expiry', 'expiration', 'expirydate', 'validuntil']);
      createdIdx = findIndex(['created', 'creationdate', 'createddate', 'activationdate']);
      console.log(`Mapped indices: ID=${idIdx}, Name=${nameIdx}, Phone=${phoneIdx}, Address=${addressIdx}, Package=${pkgIdx}, Expiry=${expiryIdx}, Created=${createdIdx}`);
    }

    const defaultIndices = { id: 3, name: 4, phone: 6, address: 7, pkg: 8, expiry: 15, created: 22 };
    const finalId = idIdx !== undefined && idIdx !== -1 ? idIdx : defaultIndices.id;
    const finalName = nameIdx !== undefined && nameIdx !== -1 ? nameIdx : defaultIndices.name;
    const finalPhone = phoneIdx !== undefined && phoneIdx !== -1 ? phoneIdx : defaultIndices.phone;
    const finalAddress = addressIdx !== undefined && addressIdx !== -1 ? addressIdx : defaultIndices.address;
    const finalPkg = pkgIdx !== undefined && pkgIdx !== -1 ? pkgIdx : defaultIndices.pkg;
    const finalExpiry = expiryIdx !== undefined && expiryIdx !== -1 ? expiryIdx : defaultIndices.expiry;
    const finalCreated = createdIdx !== undefined && createdIdx !== -1 ? createdIdx : defaultIndices.created;

    // ----- Build records -----
    const records = allRows.map(row => {
      const id = stripHtml(row[finalId] || '');
      const name = stripHtml(row[finalName] || '');
      const phone = stripHtml(row[finalPhone] || '');
      const address = stripHtml(row[finalAddress] || '');
      const pkg = stripHtml(row[finalPkg] || '');
      const activationDate = parseCrmDate(row[finalCreated] || '');
      const expiryDate = parseCrmDate(row[finalExpiry] || '');
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
