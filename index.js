// ====================================================
// index.js — CRM Table Scraper (paginates through all pages)
// Deploy on Railway
// ====================================================

const puppeteer = require('puppeteer');
const { google } = require('googleapis');

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const {
  CRM_USERNAME,
  CRM_PASSWORD,
  CRM_LOGIN_URL,
  CRM_LIST_PAGE_URL,
  GOOGLE_SHEET_ID,
  GOOGLE_CREDENTIALS,
} = process.env;

const required = ['CRM_USERNAME', 'CRM_PASSWORD', 'CRM_LOGIN_URL', 'CRM_LIST_PAGE_URL', 'GOOGLE_SHEET_ID', 'GOOGLE_CREDENTIALS'];
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
const ACTIVATION_OFFSET_DAYS = 30; // Activation/Payment Date = Expiry - 30 days, per instruction

// Edit this list to match your real package names and prices.
// The Package column text is matched against these keys (case-insensitive).
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

// ---------- Helpers ----------
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

// ---------- Find the right table AND map its columns by header text ----------
async function findTableAndColumnMap(page) {
  return await page.evaluate(() => {
    const tables = Array.from(document.querySelectorAll('table'));
    let best = null;
    let bestScore = -Infinity;
    let bestHeaders = [];

    for (const table of tables) {
      const headerCells = table.querySelectorAll('thead th, tr:first-child th, tr:first-child td');
      const headers = Array.from(headerCells).map((c) => c.innerText.trim());
      if (headers.length < 5) continue;

      const lower = headers.map((h) => h.toLowerCase());
      const hasName = lower.some((h) => h.includes('name'));
      const hasPhone = lower.some((h) => h.includes('phone') || h.includes('mobile'));
      const hasAddress = lower.some((h) => h.includes('address'));
      const hasExpiry = lower.some((h) => h.includes('expiry') || h.includes('expire'));
      const hasAction = lower.some((h) => h.includes('action'));

      let score = 0;
      if (hasName) score += 10;
      if (hasPhone) score += 10;
      if (hasAddress) score += 10;
      if (hasExpiry) score += 10;
      score += headers.length;
      if (hasAction && !hasName) score -= 30;

      if (score > bestScore) {
        bestScore = score;
        best = table;
        bestHeaders = headers;
      }
    }

    if (!best) return { success: false };

    // Assign a stable id/attribute so we can re-find this exact table on later pages
    if (!best.id) best.id = 'crm_sync_target_table';

    const lower = bestHeaders.map((h) => h.toLowerCase());
    const findCol = (keywords) => lower.findIndex((h) => keywords.some((k) => h.includes(k)));

    const columnMap = {
      name: findCol(['name']),
      phone: findCol(['phone', 'mobile']),
      address: findCol(['address']),
      pkg: findCol(['package', 'plan']),
      expiry: findCol(['expiry', 'expire']),
    };

    return { success: true, tableId: best.id, headers: bestHeaders, columnMap };
  });
}

// ---------- Set the page-length dropdown to the largest option available ----------
async function selectLargestPageLength(page) {
  return await page.evaluate(() => {
    const selects = Array.from(document.querySelectorAll('select'));
    let bestSelect = null;
    let bestValue = -1;
    for (const s of selects) {
      const opts = Array.from(s.querySelectorAll('option'));
      for (const opt of opts) {
        const num = parseInt(opt.value, 10);
        if (!isNaN(num) && num > bestValue && num <= 1000) {
          bestValue = num;
          bestSelect = s;
        }
        if (opt.value === '-1' || /all/i.test(opt.textContent)) {
          bestSelect = s;
          bestValue = 999999;
          s.value = opt.value;
          s.dispatchEvent(new Event('change', { bubbles: true }));
          return { changed: true, value: opt.value };
        }
      }
    }
    if (bestSelect && bestValue > 0) {
      bestSelect.value = String(bestValue);
      bestSelect.dispatchEvent(new Event('change', { bubbles: true }));
      return { changed: true, value: bestValue };
    }
    return { changed: false };
  });
}

// ---------- Scrape all rows of the identified table using its known column map ----------
async function scrapeTableRows(page, tableId) {
  return await page.evaluate((id) => {
    const table = document.getElementById(id) || document.querySelector(`table#${id}`);
    if (!table) return [];
    const rows = Array.from(table.querySelectorAll('tbody tr'));
    return rows
      .map((tr) => Array.from(tr.querySelectorAll('td')).map((td) => td.innerText.trim()))
      .filter((r) => r.length > 0);
  }, tableId);
}

// ---------- Click "Next" if it exists and is enabled ----------
async function clickNextIfAvailable(page) {
  return await page.evaluate(() => {
    const candidates = Array.from(document.querySelectorAll('a, button, li'));
    for (const el of candidates) {
      const text = (el.innerText || '').trim().toLowerCase();
      const cls = (el.className || '').toString().toLowerCase();
      const isNext = text === 'next' || text === '>' || cls.includes('next') || cls.includes('paginate_button next');
      if (!isNext) continue;
      const disabled =
        el.disabled ||
        el.getAttribute('aria-disabled') === 'true' ||
        cls.includes('disabled');
      if (disabled) return { clicked: false, reachedEnd: true };
      el.click();
      return { clicked: true, reachedEnd: false };
    }
    return { clicked: false, reachedEnd: true };
  });
}

// ---------- Read existing sheet, build phone -> row index map ----------
async function getExistingSheetData() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: 'Customers!A1:Z',
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
    console.error('Failed to read sheet:', err.message);
    return { headers: [], data: [], phoneMap: new Map() };
  }
}

// ---------- Update sheet, matching your existing column headers exactly ----------
async function updateSheet(records) {
  const { headers, data, phoneMap } = await getExistingSheetData();
  if (headers.length === 0) {
    console.error('Sheet headers not found. Make sure the Customers tab has header row 1 filled in.');
    return;
  }

  const nameIdx = headers.indexOf('Name');
  const phoneIdx = headers.indexOf('Phone');
  const addressIdx = headers.indexOf('Address');
  const pkgIdx = headers.indexOf('Package');
  const amountPaidIdx = headers.indexOf('Amount Paid');
  const activationIdx = headers.indexOf('Activation Date');
  const paymentDateIdx = headers.indexOf('Payment Date');
  const expiryIdx = headers.indexOf('Expiry Date');
  const statusIdx = headers.indexOf('Status');
  const daysIdx = headers.indexOf('Days Remaining');
  const lastNotifiedIdx = headers.indexOf('Last Notified');

  if (phoneIdx === -1) {
    console.error('Column "Phone" not found in sheet.');
    return;
  }

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
    if (amountPaidIdx !== -1) sheetRow[amountPaidIdx] = getPriceForPackage(rec.pkg);
    if (expiryIdx !== -1) sheetRow[expiryIdx] = formatDateForSheet(rec.expiryDate);

    // Activation Date and Payment Date = Expiry - 30 days, as instructed
    let activationDate = null;
    if (rec.expiryDate) {
      activationDate = new Date(rec.expiryDate);
      activationDate.setDate(activationDate.getDate() - ACTIVATION_OFFSET_DAYS);
    }
    if (activationIdx !== -1) sheetRow[activationIdx] = formatDateForSheet(activationDate);
    if (paymentDateIdx !== -1) sheetRow[paymentDateIdx] = formatDateForSheet(activationDate);

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
      range: `Customers!A2:Z${allRows.length + 1}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: allRows },
    });
    console.log(`Sheet updated: ${updatedCount} updated, ${appendedCount} appended.`);
  } else {
    console.log('No changes to sheet.');
  }
}

// ---------- Main sync ----------
async function syncCRM() {
  console.log(`[${new Date().toISOString()}] Sync started...`);
  let browser = null;

  try {
    browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox'],
      protocolTimeout: 180000,
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 900 });

    console.log('Logging in...');
    await page.goto(CRM_LOGIN_URL, { waitUntil: 'networkidle2', timeout: 60000 });
    await page.type('input[name="username"]', CRM_USERNAME, { delay: 30 });
    await page.type('input[name="password"]', CRM_PASSWORD, { delay: 30 });
    await Promise.all([
      page.click('button[type="submit"]'),
      page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 60000 }),
    ]);
    console.log('Login successful.');

    console.log('Opening customer list page...');
    await page.goto(CRM_LIST_PAGE_URL, { waitUntil: 'networkidle2', timeout: 60000 });
    await page.waitForSelector('table', { timeout: 30000 });
    await sleep(1500);

    console.log('Trying to select the largest page size...');
    await selectLargestPageLength(page);
    await sleep(2500);

    console.log('Identifying the correct customer table...');
    const tableInfo = await findTableAndColumnMap(page);
    if (!tableInfo.success) throw new Error('Could not find a table with Name/Phone/Address/Expiry columns.');
    console.log(`Selected table headers: ${tableInfo.headers.join(' | ')}`);
    console.log(`Column map: ${JSON.stringify(tableInfo.columnMap)}`);

    if (Object.values(tableInfo.columnMap).some((v) => v === -1)) {
      console.warn('Warning: one or more expected columns (name/phone/address/expiry) were not found by header text. Results may be incomplete.');
    }

    let allRows = [];
    let pageNum = 1;
    const maxPages = 200;

    while (pageNum <= maxPages) {
      console.log(`Scraping page ${pageNum}...`);
      const rows = await scrapeTableRows(page, tableInfo.tableId);
      console.log(`  Found ${rows.length} rows on this page.`);
      allRows = allRows.concat(rows);

      const nextResult = await clickNextIfAvailable(page);
      if (!nextResult.clicked) {
        console.log('Reached the last page.');
        break;
      }
      await sleep(1800);
      pageNum++;
    }

    console.log(`Scraped ${allRows.length} total rows across ${pageNum} page(s).`);
    if (allRows.length === 0) throw new Error('No rows found — check table detection.');

    const { name: ni, phone: pi, address: ai, pkg: pki, expiry: ei } = tableInfo.columnMap;
    const records = allRows
      .map((row) => ({
        name: ni !== -1 ? stripHtml(row[ni]) : '',
        phone: pi !== -1 ? stripHtml(row[pi]) : '',
        address: ai !== -1 ? stripHtml(row[ai]) : '',
        pkg: pki !== -1 ? stripHtml(row[pki]) : '',
        expiryDate: ei !== -1 ? parseCrmDate(row[ei]) : null,
      }))
      .filter((r) => r.phone);

    console.log(`Mapped ${records.length} valid customer records.`);
    if (records.length > 0) console.log('Sample record:', JSON.stringify(records[0]));

    await updateSheet(records);
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
