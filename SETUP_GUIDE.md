# CRM to Google Sheets Auto-Sync Bot — Setup Guide

This program logs into your CRM using a real browser and writes your customer
data into Google Sheets automatically, every 5 minutes, forever. You do not
need to touch it again once it's running.

There are 2 parts to set up: (A) letting the bot access your Google Sheet,
and (B) putting the bot online so it runs 24/7.

---

## PART A: Let the bot access your Google Sheet

1. Go to https://console.cloud.google.com and sign in with your Google account.
2. Click "Select a project" (top left) > "New Project". Name it anything,
   e.g. "crm-sync-bot". Click Create.
3. In the search bar at the top, search for "Google Sheets API" and click
   "Enable".
4. In the left menu, go to "IAM & Admin" > "Service Accounts".
5. Click "+ Create Service Account". Name it "sheet-writer". Click "Create
   and Continue", then "Done" (skip the optional steps).
6. Click on the service account you just created. Go to the "Keys" tab.
7. Click "Add Key" > "Create new key" > choose "JSON" > Create.
   A file will download to your computer — this is your GOOGLE_CREDENTIALS.
   Keep it safe, it's like a password.
8. Open that downloaded JSON file in Notepad, copy ALL of its content.
9. Open the JSON file again and find the line that says "client_email" —
   copy that email address (looks like
   sheet-writer@your-project.iam.gserviceaccount.com).
10. Open your Google Sheet, click "Share" (top right), paste that email
    address in, give it "Editor" access, and click Send.

Your bot can now write to your Sheet.

---

## PART B: Put the bot online (using Railway — beginner friendly)

1. Go to https://railway.app and sign up (you can sign up with your Google
   account).
2. Click "New Project" > "Empty Project".
3. You'll need these 3 files in one folder on your computer:
   package.json, index.js, and this guide (for reference).
   Zip that folder, or push it to GitHub if you know how — Railway can
   deploy from either.
4. In Railway, click "Deploy from GitHub repo" (if you used GitHub) or
   drag-and-drop your zipped folder if Railway offers an upload option
   for your account type.
5. Once deployed, go to your project's "Variables" tab in Railway and add
   these one by one:
   - CRM_USERNAME → your CRM login username
   - CRM_PASSWORD → your CRM login password
   - CRM_LOGIN_URL → http://223.123.38.98/login
   - CRM_LIST_URL → http://223.123.38.98/user/all
   - GOOGLE_SHEET_ID → the long code in your Sheet's URL, e.g. in
     https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOp/edit
     the ID is 1AbCdEfGhIjKlMnOp
   - GOOGLE_CREDENTIALS → paste the ENTIRE content of the JSON file from
     Part A, step 8, as one single line
6. Railway will automatically install everything and start running the bot.
7. Click on "Deployments" > the latest deployment > "View Logs" to watch
   it work. You should see lines like "Sync started...", "Found X rows",
   "Sheet updated."

That's it. From now on, your Google Sheet updates by itself every 5
minutes, with zero manual work.

---

## Important: this may need one adjustment

The bot currently assumes your customer list page has a simple HTML table
it can read directly. If the first run's log shows "Found 0 rows" or the
data looks jumbled, send me the log output — I'll adjust the script to
match your CRM's actual page structure.
