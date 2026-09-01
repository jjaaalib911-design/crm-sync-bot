// ====================================================
// index.js — CRM Live Data Sync (overwrites sheet every 30 min)
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

const sheets = google.sheets({
  version: 'v4',
  auth
});

const REMINDER_WINDOW_DAYS = 3;
const UNIQUE_KEY = 'Phone';


// ====================================================
// SYNC SCHEDULE
//
// Every 30 minutes -> 48 runs per day.
// ====================================================

const SYNC_INTERVAL_MS = 30 * 60 * 1000;

// Maximum time for one complete sync.
const SYNC_TIMEOUT_MS = 4 * 60 * 1000;

// Prevent overlapping syncs.
let isSyncing = false;


// ====================================================
// PACKAGE PRICES
// ====================================================

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
  '20+10Mbps': 3000,
  '30 MB EID': 4000,
  '10+10Mbps': 2500,
  '14+10Mbps': 2500,
  '15 MB Day': 2000,
  'default': 0,
};


function getPriceForPackage(pkg) {
  if (!pkg) return PACKAGE_PRICES.default || 0;

  if (PACKAGE_PRICES[pkg] !== undefined) {
    return PACKAGE_PRICES[pkg];
  }

  const lower = pkg.toLowerCase().trim();

  for (const [key, value] of Object.entries(PACKAGE_PRICES)) {
    if (key.toLowerCase().trim() === lower) {
      return value;
    }
  }

  return PACKAGE_PRICES.default || 0;
}


// ====================================================
// STRIP HTML
// ====================================================

function stripHtml(str) {
  if (!str) return '';

  return String(str)
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


// ====================================================
// MONTHS
// ====================================================

const MONTHS = {
  Jan: 0,
  Feb: 1,
  Mar: 2,
  Apr: 3,
  May: 4,
  Jun: 5,
  Jul: 6,
  Aug: 7,
  Sep: 8,
  Oct: 9,
  Nov: 10,
  Dec: 11
};


// ====================================================
// PARSE CRM DATE
// ====================================================

function parseCrmDate(text) {
  const clean = stripHtml(text);

  const match = clean.match(
    /(\d{1,2})\s+(\w{3})\w*\s+(\d{4})/
  );

  if (match) {
    const [, day, monAbbr, year] = match;

    const month =
      MONTHS[monAbbr.slice(0, 3)];

    if (month !== undefined) {
      return new Date(
        parseInt(year),
        month,
        parseInt(day)
      );
    }
  }

  const isoLike =
    clean.match(/(\d{4})-(\d{2})-(\d{2})/);

  if (isoLike) {
    const [, y, m, d] = isoLike;

    return new Date(
      parseInt(y),
      parseInt(m) - 1,
      parseInt(d)
    );
  }

  return null;
}


// ====================================================
// FORMAT DATE FOR GOOGLE SHEETS
// ====================================================

function formatDateForSheet(date) {
  if (!date) return '';

  const y = date.getFullYear();

  const m = String(
    date.getMonth() + 1
  ).padStart(2, '0');

  const d = String(
    date.getDate()
  ).padStart(2, '0');

  return `${y}-${m}-${d}`;
}


// ====================================================
// CALCULATE ACTIVE DATE
//
// Active Date = Expiry Date - 31 Days
// ====================================================

function calculateActiveDate(expiryDate) {
  if (!expiryDate) return null;

  const activeDate =
    new Date(expiryDate);

  activeDate.setDate(
    activeDate.getDate() - 31
  );

  return activeDate;
}


// ====================================================
// CALCULATE DAYS REMAINING
//
// Days Remaining = Expiry Date - TODAY
// ====================================================

function calculateDaysRemaining(expiryDate, today) {
  if (!expiryDate) return '';

  const diffMs =
    expiryDate - today;

  return Math.round(
    diffMs /
    (1000 * 60 * 60 * 24)
  );
}


// ====================================================
// CALCULATE STATUS
//
// Expired       = less than 0 days
// Expiring Soon = 0 to 3 days
// Active        = more than 3 days
// ====================================================

function computeStatus(expiryDate, today) {
  if (!expiryDate) {
    return {
      status: '',
      daysRemaining: ''
    };
  }

  const diffDays =
    calculateDaysRemaining(
      expiryDate,
      today
    );

  let status;

  if (diffDays < 0) {
    status = 'Expired';
  } else if (
    diffDays <= REMINDER_WINDOW_DAYS
  ) {
    status = 'Expiring Soon';
  } else {
    status = 'Active';
  }

  return {
    status,
    daysRemaining: diffDays
  };
}


// ====================================================
// EXTRACT CRM RECORD
//
// Confirmed CRM field positions:
//
// 2  = Username / ID
// 3  = Name
// 5  = Phone
// 6  = Address
// 7  = Package
// 14 = Expiry Date
//
// ====================================================

function extractRecord(row) {
  const username =
    stripHtml(row[2]);

  const name =
    stripHtml(row[3]);

  const phone =
    stripHtml(row[5]);

  const address =
    stripHtml(row[6]);

  const pkg =
    stripHtml(row[7]);

  const expiryDate =
    parseCrmDate(row[14]);

  return {
    username,
    name,
    phone,
    address,
    pkg,
    expiryDate
  };
}


// ====================================================
// BUILD REQUEST BODY
// ====================================================

function buildRequestBody(start, length, filterType) {
  const orderableColumns = new Set([
    0,
    2,
    3,
    4,
    5,
    6,
    7
  ]);

  const params =
    new URLSearchParams();

  params.append(
    'draw',
    '1'
  );

  for (let i = 0; i <= 22; i++) {
    params.append(
      `columns[${i}][data]`,
      String(i)
    );

    params.append(
      `columns[${i}][name]`,
      ''
    );

    params.append(
      `columns[${i}][searchable]`,
      'true'
    );

    params.append(
      `columns[${i}][orderable]`,
      orderableColumns.has(i)
        ? 'true'
        : 'false'
    );

    params.append(
      `columns[${i}][search][value]`,
      ''
    );

    params.append(
      `columns[${i}][search][regex]`,
      'false'
    );
  }

  params.append(
    'order[0][column]',
    '0'
  );

  params.append(
    'order[0][dir]',
    'desc'
  );

  params.append(
    'start',
    String(start)
  );

  params.append(
    'length',
    String(length)
  );

  params.append(
    'search[value]',
    ''
  );

  params.append(
    'search[regex]',
    'false'
  );

  params.append(
    'filterType',
    String(filterType)
  );

  params.append(
    'dashboardUserTables',
    '1'
  );

  return params.toString();
}


// ====================================================
// CALL CRM DATA API
// ====================================================

async function callDataApi(page, body) {
  return await page.evaluate(
    async (url, b) => {

      const res =
        await fetch(url, {
          method: 'POST',

          headers: {
            'Content-Type':
              'application/x-www-form-urlencoded; charset=UTF-8',

            'X-Requested-With':
              'XMLHttpRequest',
          },

          body: b,
        });

      if (!res.ok) {
        throw new Error(
          `CRM API returned HTTP ${res.status}`
        );
      }

      return res.json();

    },
    CRM_DATA_API_URL,
    body
  );
}


// ====================================================
// FIND BEST FILTER TYPE
// ====================================================

async function findBestFilterType(page) {
  const candidates = [
    0,
    1,
    2,
    3,
    4,
    5,
    ''
  ];

  let best = {
    filterType: 3,
    recordsFiltered: 0,
    recordsTotal: 0
  };

  for (const ft of candidates) {

    try {

      const body =
        buildRequestBody(
          0,
          1,
          ft
        );

      const result =
        await callDataApi(
          page,
          body
        );

      const filtered =
        result.recordsFiltered || 0;

      const total =
        result.recordsTotal || 0;

      console.log(
        `  filterType=${JSON.stringify(ft)} -> recordsFiltered=${filtered}, recordsTotal=${total}`
      );

      if (
        filtered >
        best.recordsFiltered
      ) {
        best = {
          filterType: ft,
          recordsFiltered: filtered,
          recordsTotal: total
        };
      }

    } catch (e) {

      console.log(
        `  filterType=${JSON.stringify(ft)} -> request failed: ${e.message}`
      );

    }
  }

  return best;
}


// ====================================================
// OVERWRITE SHEET
// ====================================================

async function overwriteSheet(records) {

  const HEADERS = [
    'ID No.',
    'Name',
    'Phone',
    'Address',
    'Package',
    'Amount Paid',
    'Payment Date',
    'Activation Date',
    'Expiry Date',
    'Status',
    'Days Remaining',
    'Last Notified'
  ];


  // ==================================================
  // TODAY
  // ==================================================

  const today =
    new Date();

  today.setHours(
    0,
    0,
    0,
    0
  );


  // ==================================================
  // BUILD ROWS
  // ==================================================

  const rows =
    records
      .map(rec => {

        if (!rec.phone) {
          return null;
        }


        // ----------------------------------------------
        // PACKAGE PRICE
        // ----------------------------------------------

        const price =
          getPriceForPackage(
            rec.pkg
          );


        // ----------------------------------------------
        // EXPIRY DATE
        // From CRM
        // ----------------------------------------------

        const expiryDateStr =
          formatDateForSheet(
            rec.expiryDate
          );


        // ----------------------------------------------
        // ACTIVE DATE
        //
        // Expiry Date - 31 Days
        // ----------------------------------------------

        const activeDate =
          calculateActiveDate(
            rec.expiryDate
          );


        const activationDateStr =
          formatDateForSheet(
            activeDate
          );


        // ----------------------------------------------
        // PAYMENT DATE
        //
        // Payment Date = Active Date
        // ----------------------------------------------

        const paymentDateStr =
          formatDateForSheet(
            activeDate
          );


        // ----------------------------------------------
        // STATUS + DAYS REMAINING
        // ----------------------------------------------

        const {
          status,
          daysRemaining
        } =
          computeStatus(
            rec.expiryDate,
            today
          );


        // ----------------------------------------------
        // FINAL ROW
        // ----------------------------------------------

        return [
          rec.username,       // ID No.
          rec.name,           // Name
          rec.phone,          // Phone
          rec.address,        // Address
          rec.pkg,            // Package
          price,              // Amount Paid

          paymentDateStr,     // Payment Date
          activationDateStr,  // Activation Date
          expiryDateStr,      // Expiry Date

          status,             // Status
          daysRemaining,      // Days Remaining

          ''                  // Last Notified
        ];

      })
      .filter(
        row => row !== null
      );


  // ==================================================
  // HEADERS + ROWS
  // ==================================================

  const allData = [
    HEADERS,
    ...rows
  ];


  // ==================================================
  // WRITE TO GOOGLE SHEET
  // ==================================================

  await sheets.spreadsheets.values.update({

    spreadsheetId:
      GOOGLE_SHEET_ID,

    range:
      'Customers!A1',

    valueInputOption:
      'USER_ENTERED',

    resource: {
      values: allData
    }

  });


  console.log(
    `✅ Sheet overwritten with ${rows.length} rows.`
  );
}


// ====================================================
// LAUNCH BROWSER
// ====================================================

async function launchBrowser() {
  return puppeteer.launch({
    headless: true,

    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--disable-software-rasterizer',
      '--disable-extensions',
      '--disable-background-networking',
      '--disable-crash-reporter',
      '--disable-breakpad',
      '--no-zygote',
      '--no-first-run',
    ],
  });
}


// ====================================================
// RUN WITH TIMEOUT
// ====================================================

function withTimeout(promise, ms, label) {
  let timer;

  const timeout = new Promise((_, reject) => {
    timer = setTimeout(
      () => reject(
        new Error(
          `${label} timed out after ${ms}ms`
        )
      ),
      ms
    );
  });

  return Promise.race([
    promise,
    timeout
  ]).finally(() => {
    clearTimeout(timer);
  });
}


// ====================================================
// ROBUST NAVIGATION
//
// IMPORTANT FIX:
//
// Old:
// waitUntil: 'networkidle2'
// timeout: 30000
//
// New:
// waitUntil: 'domcontentloaded'
// timeout: 120000
//
// Also retries navigation 3 times.
//
// This avoids the problem where CRM pages keep network
// connections open and Puppeteer waits forever for
// "networkidle2".
// ====================================================

const NAVIGATION_TIMEOUT_MS = 120 * 1000;
const NAVIGATION_RETRIES = 3;


async function safeGoto(page, url, label) {

  let lastError = null;

  for (
    let attempt = 1;
    attempt <= NAVIGATION_RETRIES;
    attempt++
  ) {

    try {

      console.log(
        `${label}: opening page (attempt ${attempt}/${NAVIGATION_RETRIES})...`
      );


      await page.goto(
        url,
        {
          waitUntil: 'domcontentloaded',
          timeout: NAVIGATION_TIMEOUT_MS
        }
      );


      console.log(
        `${label}: page loaded.`
      );


      return;

    } catch (error) {

      lastError = error;


      console.error(
        `${label}: navigation attempt ${attempt} failed: ${error.message}`
      );


      if (
        attempt < NAVIGATION_RETRIES
      ) {

        console.log(
          `${label}: waiting 5 seconds before retry...`
        );


        await new Promise(
          resolve =>
            setTimeout(
              resolve,
              5000
            )
        );

      }

    }
  }


  throw new Error(
    `${label}: unable to load after ${NAVIGATION_RETRIES} attempts. Last error: ${lastError?.message || 'unknown error'}`
  );
}


// ====================================================
// WAIT FOR LOGIN RESULT
//
// We do NOT use waitForNavigation() here.
//
// Some modern CRM websites use AJAX/SPA login and therefore
// may not trigger a normal browser navigation event.
//
// Instead, we wait until:
// 1. Login form disappears, OR
// 2. URL changes away from /login
// ====================================================

async function waitForLoginResult(page) {

  await page.waitForFunction(

    () => {

      const url =
        window.location.href.toLowerCase();


      const usernameInput =
        document.querySelector(
          'input[name="username"], input[type="email"], input[name="email"]'
        );


      const passwordInput =
        document.querySelector(
          'input[name="password"], input[type="password"]'
        );


      // Login form disappeared.
      if (
        !usernameInput &&
        !passwordInput
      ) {
        return true;
      }


      // URL changed away from login page.
      return !url.includes('/login');

    },

    {
      timeout:
        NAVIGATION_TIMEOUT_MS,

      polling:
        500
    }
  );
}


// ====================================================
// MAIN SYNC
// ====================================================

async function syncCRM() {

  // --------------------------------------------------
  // PREVENT OVERLAPPING SYNC
  // --------------------------------------------------

  if (isSyncing) {

    console.log(
      `[${new Date().toISOString()}] Previous sync still running — skipping this tick.`
    );

    return;
  }


  isSyncing = true;


  console.log(
    `[${new Date().toISOString()}] Sync started...`
  );


  let browser = null;


  try {

    await withTimeout(

      (async () => {


        // =================================================
        // LAUNCH BROWSER
        // =================================================

        browser =
          await launchBrowser();


        const page =
          await browser.newPage();


        // =================================================
        // DEFAULT PUPPETEER TIMEOUTS
        // =================================================

        page.setDefaultNavigationTimeout(
          NAVIGATION_TIMEOUT_MS
        );

        page.setDefaultTimeout(
          NAVIGATION_TIMEOUT_MS
        );


        await page.setViewport({
          width: 1280,
          height: 800
        });


        // =================================================
        // LOGIN
        // =================================================

        console.log(
          'Logging in...'
        );


        // -------------------------------------------------
        // OPEN LOGIN PAGE
        // -------------------------------------------------

        await safeGoto(
          page,
          CRM_LOGIN_URL,
          'CRM login'
        );


        // -------------------------------------------------
        // WAIT FOR USERNAME FIELD
        // -------------------------------------------------

        await page.waitForSelector(

          'input[name="username"], input[type="email"], input[name="email"]',

          {
            timeout:
              NAVIGATION_TIMEOUT_MS
          }

        );


        // -------------------------------------------------
        // FIND USERNAME SELECTOR
        // -------------------------------------------------

        const usernameSelector =
          await page.$(
            'input[name="username"]'
          )
            ? 'input[name="username"]'
            : (
                await page.$(
                  'input[type="email"]'
                )
                  ? 'input[type="email"]'
                  : 'input[name="email"]'
              );


        // -------------------------------------------------
        // FIND PASSWORD SELECTOR
        // -------------------------------------------------

        const passwordSelector =
          await page.$(
            'input[name="password"]'
          )
            ? 'input[name="password"]'
            : 'input[type="password"]';


        // -------------------------------------------------
        // TYPE USERNAME
        // -------------------------------------------------

        await page.type(
          usernameSelector,
          CRM_USERNAME,
          {
            delay: 30
          }
        );


        // -------------------------------------------------
        // TYPE PASSWORD
        // -------------------------------------------------

        await page.type(
          passwordSelector,
          CRM_PASSWORD,
          {
            delay: 30
          }
        );


        // -------------------------------------------------
        // FIND SUBMIT BUTTON
        // -------------------------------------------------

        const submitSelector =
          await page.$(
            'button[type="submit"]'
          )
            ? 'button[type="submit"]'
            : 'input[type="submit"]';


        // -------------------------------------------------
        // CLICK LOGIN
        //
        // IMPORTANT:
        // We no longer use Promise.all() with
        // page.waitForNavigation().
        // -------------------------------------------------

        await page.click(
          submitSelector
        );


        // -------------------------------------------------
        // WAIT FOR LOGIN RESULT
        // -------------------------------------------------

        await waitForLoginResult(
          page
        );


        console.log(
          'Login successful.'
        );


        // =================================================
        // OPEN CUSTOMER LIST
        // =================================================

        console.log(
          'Opening customer list page...'
        );


        await safeGoto(
          page,
          CRM_LIST_PAGE_URL,
          'Customer list'
        );


        // =================================================
        // FIND BEST FILTER
        // =================================================

        console.log(
          'Testing filter settings to find the one showing ALL customers...'
        );


        const best =
          await findBestFilterType(
            page
          );


        console.log(
          `Best filterType found: ${JSON.stringify(best.filterType)} with ${best.recordsFiltered} of ${best.recordsTotal} total customers.`
        );


        // =================================================
        // FETCH ALL CUSTOMERS
        // =================================================

        const total =
          Math.max(
            best.recordsFiltered,
            best.recordsTotal,
            1
          );


        const fullBody =
          buildRequestBody(
            0,
            total,
            best.filterType
          );


        const fullResult =
          await callDataApi(
            page,
            fullBody
          );


        const rawRows =
          fullResult.data || [];


        console.log(
          `Fetched ${rawRows.length} customer records.`
        );


        // -------------------------------------------------
        // CHECK DATA
        // -------------------------------------------------

        if (
          rawRows.length === 0
        ) {

          throw new Error(
            'No records returned even after testing filter types.'
          );

        }


        // =================================================
        // DEBUG RAW FIRST RECORD
        // =================================================

        console.log(
          'RAW first record (for column verification):'
        );


        console.log(
          JSON.stringify(
            rawRows[0]
          )
        );


        // =================================================
        // EXTRACT RECORDS
        // =================================================

        const records =
          rawRows.map(
            extractRecord
          );


        // =================================================
        // OVERWRITE SHEET
        // =================================================

        await overwriteSheet(
          records
        );


        console.log(
          `[${new Date().toISOString()}] Sync completed.`
        );


      })(),

      SYNC_TIMEOUT_MS,

      'Sync'

    );


  } catch (error) {

    console.error(
      'Sync failed:',
      error.message
    );


  } finally {


    // =================================================
    // CLOSE BROWSER
    // =================================================

    if (browser) {

      try {

        await browser.close();

      } catch (closeErr) {

        console.error(
          'Error while closing browser:',
          closeErr.message
        );

      }

    }


    isSyncing = false;

  }
}


// ====================================================
// PROCESS-LEVEL SAFETY NETS
// ====================================================

process.on(
  'unhandledRejection',
  (reason) => {

    console.error(
      'Unhandled promise rejection:',
      reason
    );

  }
);


process.on(
  'uncaughtException',
  (err) => {

    console.error(
      'Uncaught exception:',
      err.message
    );

  }
);


process.on(
  'SIGTERM',
  () => {

    console.log(
      'Received SIGTERM — shutting down gracefully.'
    );

    process.exit(0);

  }
);


// ====================================================
// START BOT
// ====================================================

console.log(
  'CRM Sync Bot starting...'
);


console.log(
  `Navigation timeout: ${NAVIGATION_TIMEOUT_MS / 1000}s | Navigation retries: ${NAVIGATION_RETRIES}`
);


syncCRM();


// ====================================================
// RUN EVERY 30 MINUTES
// ====================================================

setInterval(
  syncCRM,
  SYNC_INTERVAL_MS
);
