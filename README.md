# 💊 HSA Tracker

A personal web app for tracking HSA (Health Savings Account) expenses, receipts, and investment growth — built with vanilla HTML/CSS/JS and powered by Google APIs.

## Features

### Expense Tracking
- Log medical expenses with date, amount, and description
- Upload receipt files (PDF, JPG, PNG) to Google Drive
- Mark expenses as reimbursed with date tracking
- Filter by status (pending / reimbursed) and year
- Bulk-select and batch-reimburse

### Investment Dashboard
- Track balances across two HSA custodians (Optum Bank + Betterment)
- Combined total view with color-coded cards
- Balance history log stored in Google Sheets
- Growth projection calculator — see what your HSA could be worth at age 65 with configurable contribution, frequency, return rate, and time horizon

### Automated Receipt Processing
- Google Apps Script monitors Gmail for emails labeled **HSA**
- Extracts attachments, uploads to Drive, and uses **Gemini AI** to OCR/parse expense details
- Automatically appends rows to the spreadsheet (runs every 5 minutes)

## Tech Stack

| Layer | Tech |
|---|---|
| Frontend | Vanilla HTML, CSS, JS (no framework) |
| Auth | Google OAuth 2.0 (client-side) |
| Data | Google Sheets API v4 |
| File Storage | Google Drive API v3 |
| Automation | Google Apps Script + Gemini 2.0 Flash |

## Setup

1. **Google Cloud Project** — create an OAuth 2.0 Web Client at [console.cloud.google.com](https://console.cloud.google.com/apis/credentials) and enable the Sheets and Drive APIs.
2. **config.js** — set your `GOOGLE_CLIENT_ID`, `SPREADSHEET_ID`, `DRIVE_FOLDER_ID`, and `ALLOWED_EMAILS`.
3. **Google Sheet** — create two tabs:
   - `Sheet1` — columns: Expense | Date | Amount | Reimbursed | Reimburse Date | Receipt Link
   - `Balances` — columns: Date | Institution | Balance
4. **Apps Script** (optional) — deploy `apps-script/Code.gs` in a Google Apps Script project, set script properties (`GEMINI_API_KEY`, `SPREADSHEET_ID`, `DRIVE_FOLDER_ID`, `SHEET_NAME`), and run `setup()` to create the Gmail trigger.
5. **Serve** — open `index.html` via any static server (e.g. `npx serve` or GitHub Pages).

## Project Structure

```
index.html          Login page (Google OAuth)
app.html            Main app (all views)
app.js              Application logic
style.css           Dark-theme styles
config.js           API keys & config
apps-script/
  Code.gs           Gmail → Gemini → Sheets automation
```
