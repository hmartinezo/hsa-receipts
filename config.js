const CONFIG = {
  // ── Google Cloud OAuth 2.0 Client ID ──
  // Create at: https://console.cloud.google.com/apis/credentials
  // Type: Web application
  // Authorized JavaScript origins: http://localhost:8080, https://<your-github-pages-url>
  GOOGLE_CLIENT_ID: '176819176117-ni86t6ft79f13ennbbdkc0kvtapmt1c5.apps.googleusercontent.com',

  // ── Google Sheet ID (from the URL) ──
  SPREADSHEET_ID: '1rENdIszVSyVn-nbZxZLGYkIAVlwsXQ3x6puXYZ5mWiw',

  // ── Google Drive folder ID for receipts ──
  DRIVE_FOLDER_ID: '1kyOa6mOSJPG93OzwJgEXctHoeehVCyT3',

  // ── Sheet tab name ──
  SHEET_NAME: 'Sheet1',

  // ── Balance tracking sheet tab ──
  BALANCE_SHEET_NAME: 'Balances',

  // ── Allowed emails (lowercase) ──
  ALLOWED_EMAILS: ['hmartinezo@gmail.com'],

  // ── OAuth scopes ──
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file openid email profile',
};
