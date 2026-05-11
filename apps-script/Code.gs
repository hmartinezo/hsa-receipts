/**
 * Google Apps Script -- HSA Receipt Email Processor
 *
 * Monitors Gmail for emails labeled "HSA", extracts attachments,
 * uploads them to Google Drive, uses Gemini to OCR/extract expense
 * details, and appends rows to the HSA Expenses Tracking spreadsheet.
 *
 * SETUP:
 * 1. Set Script Properties (Project Settings > Script Properties):
 *    - GEMINI_API_KEY, SPREADSHEET_ID, DRIVE_FOLDER_ID, SHEET_NAME
 * 2. Run setup() once to create the trigger
 * 3. Authorize when prompted
 */

// =============================================
// SETUP
// =============================================

function setup() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  ScriptApp.newTrigger('processHSAEmails')
    .timeBased()
    .everyMinutes(5)
    .create();

  getOrCreateLabel('HSA/Processed');

  Logger.log('DONE: Trigger created (every 5 min). HSA/Processed label ready.');
}

// =============================================
// MAIN
// =============================================

function processHSAEmails() {
  var props = PropertiesService.getScriptProperties();
  var geminiKey = props.getProperty('GEMINI_API_KEY');
  var sheetId   = props.getProperty('SPREADSHEET_ID');
  var folderId  = props.getProperty('DRIVE_FOLDER_ID');
  var sheetName = props.getProperty('SHEET_NAME') || 'Sheet1';

  if (!geminiKey) { Logger.log('ERROR: GEMINI_API_KEY not set.'); return; }
  if (!sheetId)   { Logger.log('ERROR: SPREADSHEET_ID not set.'); return; }
  if (!folderId)  { Logger.log('ERROR: DRIVE_FOLDER_ID not set.'); return; }

  var processedLabel = getOrCreateLabel('HSA/Processed');

  // Find all threads with HSA label
  var hsaLabel = GmailApp.getUserLabelByName('HSA');
  if (!hsaLabel) {
    Logger.log('ERROR: No "HSA" label found in Gmail.');
    return;
  }

  var threads = hsaLabel.getThreads(0, 20);
  Logger.log('Found ' + threads.length + ' thread(s) with HSA label.');

  // Filter out already-processed threads
  var unprocessed = [];
  for (var i = 0; i < threads.length; i++) {
    var labels = threads[i].getLabels();
    var alreadyDone = false;
    for (var l = 0; l < labels.length; l++) {
      if (labels[l].getName() === 'HSA/Processed') {
        alreadyDone = true;
        break;
      }
    }
    if (!alreadyDone) {
      unprocessed.push(threads[i]);
    }
  }

  Logger.log('Unprocessed: ' + unprocessed.length + ' thread(s).');
  if (unprocessed.length === 0) return;

  var sheet  = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  var folder = DriveApp.getFolderById(folderId);
  if (!sheet) { Logger.log('ERROR: Sheet "' + sheetName + '" not found.'); return; }

  var totalProcessed = 0;

  for (var i = 0; i < unprocessed.length; i++) {
    var thread = unprocessed[i];
    var messages = thread.getMessages();
    Logger.log('Thread ' + (i+1) + ': "' + thread.getFirstMessageSubject() + '" (' + messages.length + ' msg)');

    // Pick the BEST single message to process (avoid duplicates)
    // Prefer the first message that has a valid attachment (PDF/image)
    var bestMsg = null;
    for (var j = 0; j < messages.length; j++) {
      var atts = messages[j].getAttachments();
      for (var a = 0; a < atts.length; a++) {
        var m = atts[a].getContentType().toLowerCase();
        if (m.indexOf('pdf') !== -1 || m.indexOf('image') !== -1) {
          bestMsg = messages[j];
          break;
        }
      }
      if (bestMsg) break;
    }
    // If no message has attachments, use the first message
    if (!bestMsg) bestMsg = messages[0];

    Logger.log('  Processing: "' + bestMsg.getSubject() + '", attachments=' + bestMsg.getAttachments().length);

    try {
      var results = processOneMessage(bestMsg, folder, geminiKey);

      for (var k = 0; k < results.length; k++) {
        var r = results[k];
        Logger.log('  -> ' + r.expense + ' | ' + r.date + ' | $' + r.amount);
        sheet.appendRow([r.expense, r.date, r.amount, 'No', '', r.receiptLink]);
        totalProcessed++;
      }
    } catch (e) {
      Logger.log('  ERROR: ' + e.toString());
    }

    thread.addLabel(processedLabel);
    Logger.log('  Tagged HSA/Processed.');
  }

  Logger.log('DONE. ' + totalProcessed + ' expense(s) added to sheet.');
}

// =============================================
// PROCESS ONE EMAIL
// =============================================

function processOneMessage(msg, folder, geminiKey) {
  var subject   = msg.getSubject() || '';
  var bodyText  = msg.getPlainBody() || stripHtml(msg.getBody());
  var emailDate = Utilities.formatDate(msg.getDate(), 'America/New_York', 'MM/dd/yyyy');

  var attachments = msg.getAttachments();
  var results = [];

  if (attachments.length > 0) {
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      var mime = att.getContentType().toLowerCase();
      Logger.log('    Attachment: "' + att.getName() + '" (' + mime + ')');

      if (mime.indexOf('pdf') === -1 &&
          mime.indexOf('image') === -1 &&
          mime.indexOf('jpg') === -1 &&
          mime.indexOf('jpeg') === -1 &&
          mime.indexOf('png') === -1) {
        Logger.log('    Skipped (unsupported type).');
        continue;
      }

      var receiptLink = uploadToDrive(att, folder);
      Logger.log('    Uploaded: ' + receiptLink);

      var extracted = extractWithGemini(geminiKey, att, subject, bodyText, emailDate);

      // Fallback: if Gemini returned 0, try regex on email body
      var amt = parseFloat(extracted.amount) || 0;
      if (amt === 0 && bodyText.length > 0) {
        var regexAmt = extractAmountFromText(bodyText);
        if (regexAmt) {
          Logger.log('    Gemini returned $0, regex fallback found: $' + regexAmt);
          extracted.amount = regexAmt;
        }
      }

      results.push({
        expense: extracted.expense,
        date: extracted.date,
        amount: extracted.amount,
        receiptLink: receiptLink
      });
    }
  }

  if (results.length === 0 && bodyText.length > 0) {
    Logger.log('    No attachments, parsing email body...');
    var extracted = extractFromTextWithGemini(geminiKey, subject, bodyText, emailDate);

    // Fallback: if Gemini returned 0, try regex
    var amt = parseFloat(extracted.amount) || 0;
    if (amt === 0) {
      var regexAmt = extractAmountFromText(bodyText);
      if (regexAmt) {
        Logger.log('    Gemini returned $0, regex fallback found: $' + regexAmt);
        extracted.amount = regexAmt;
      }
    }

    results.push({
      expense: extracted.expense,
      date: extracted.date,
      amount: extracted.amount,
      receiptLink: ''
    });
  }

  return results;
}

// =============================================
// GOOGLE DRIVE
// =============================================

function uploadToDrive(attachment, folder) {
  var blob = attachment.copyBlob();
  var fileName = attachment.getName() || ('receipt-' + new Date().getTime());
  blob.setName(fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// =============================================
// GEMINI
// =============================================

var GEMINI_MODEL = 'gemini-2.0-flash';
var GEMINI_URL = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent';

var EXTRACTION_PROMPT = [
  'You are an expense extraction assistant. Analyze the provided receipt/invoice/medical bill',
  'and extract EXACTLY three fields. Respond ONLY with valid JSON, no markdown, no explanation.',
  '',
  'Rules:',
  '- expense: Short descriptive name (e.g. "Dental Cleaning", "Eye Exam", "Prescription - Amoxicillin")',
  '- date: Service/transaction date in MM/DD/YYYY format. If unclear, use the email date provided.',
  '- amount: The dollar amount from the receipt. Look for "Amount Due", "Total", "Balance Due", "Patient Responsibility", "You Owe", or the largest dollar amount. NEVER return 0 unless the document explicitly says $0. Return as a number (no $ sign, no commas).',
  '- IMPORTANT: There is ALWAYS a non-zero amount on a medical receipt. If you cannot find it, look harder at the document.',
  '',
  'Respond with: {"expense":"...","date":"...","amount":"..."}',
  '',
  'Email subject: {{SUBJECT}}',
  'Email date: {{EMAIL_DATE}}'
].join('\n');

function extractWithGemini(apiKey, attachment, subject, bodyText, emailDate) {
  var blob = attachment.copyBlob();
  var base64 = Utilities.base64Encode(blob.getBytes());
  var mimeType = attachment.getContentType();

  var prompt = EXTRACTION_PROMPT
    .replace('{{SUBJECT}}', subject)
    .replace('{{EMAIL_DATE}}', emailDate);

  if (bodyText && bodyText.length > 0) {
    prompt += '\n\nEmail body excerpt:\n' + bodyText.substring(0, 2000);
  }

  var payload = {
    contents: [{ parts: [
      { text: prompt },
      { inline_data: { mime_type: mimeType, data: base64 } }
    ]}],
    generationConfig: { temperature: 0.1, maxOutputTokens: 256 }
  };

  return callGemini(apiKey, payload, subject, emailDate);
}

function extractFromTextWithGemini(apiKey, subject, bodyText, emailDate) {
  var prompt = EXTRACTION_PROMPT
    .replace('{{SUBJECT}}', subject)
    .replace('{{EMAIL_DATE}}', emailDate);
  prompt += '\n\nEmail body:\n' + bodyText.substring(0, 4000);

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 256 }
  };

  return callGemini(apiKey, payload, subject, emailDate);
}

function callGemini(apiKey, payload, fallbackSubject, fallbackDate) {
  var fallback = {
    expense: fallbackSubject || 'Unknown Expense',
    date: fallbackDate || Utilities.formatDate(new Date(), 'America/New_York', 'MM/dd/yyyy'),
    amount: '0.00'
  };

  try {
    var response = UrlFetchApp.fetch(GEMINI_URL + '?key=' + apiKey, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    Logger.log('    Gemini HTTP ' + code);

    if (code !== 200) {
      Logger.log('    Gemini error: ' + response.getContentText().substring(0, 500));
      return fallback;
    }

    var data = JSON.parse(response.getContentText());
    var text = '';
    var parts = data.candidates[0].content.parts;
    for (var i = 0; i < parts.length; i++) {
      if (!parts[i].thought && parts[i].text) {
        text += parts[i].text;
      }
    }

    text = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    Logger.log('    Gemini output: ' + text);

    var parsed = JSON.parse(text);
    var extractedAmount = String(parsed.amount || '0').replace(/[$,]/g, '');
    // If amount is 0 or 0.00, Gemini probably failed — log it
    if (parseFloat(extractedAmount) === 0) {
      Logger.log('    WARNING: Gemini returned $0 amount');
    }
    return {
      expense: parsed.expense || fallback.expense,
      date: parsed.date || fallback.date,
      amount: extractedAmount
    };

  } catch (e) {
    Logger.log('    Gemini error: ' + e.toString());
    return fallback;
  }
}

// =============================================
// HELPERS
// =============================================

function getOrCreateLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
    Logger.log('Created label: ' + name);
  }
  return label;
}

/**
 * Regex fallback: extract the largest dollar amount from text.
 * Looks for patterns like $123.45, $1,234.56, etc.
 */
function extractAmountFromText(text) {
  if (!text) return null;
  var matches = text.match(/\$[\d,]+\.?\d{0,2}/g);
  if (!matches || matches.length === 0) return null;

  var largest = 0;
  for (var i = 0; i < matches.length; i++) {
    var val = parseFloat(matches[i].replace(/[$,]/g, ''));
    if (val > largest) largest = val;
  }
  return largest > 0 ? largest.toFixed(2) : null;
}

function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// =============================================
// TEST FUNCTIONS
// =============================================

function testWithLatestEmail() {
  var props = PropertiesService.getScriptProperties();
  var geminiKey = props.getProperty('GEMINI_API_KEY');
  var folderId = props.getProperty('DRIVE_FOLDER_ID');

  if (!geminiKey) { Logger.log('ERROR: GEMINI_API_KEY not set.'); return; }
  if (!folderId) { Logger.log('ERROR: DRIVE_FOLDER_ID not set.'); return; }

  var hsaLabel = GmailApp.getUserLabelByName('HSA');
  if (!hsaLabel) { Logger.log('ERROR: No "HSA" label in Gmail.'); return; }

  var threads = hsaLabel.getThreads(0, 1);
  if (threads.length === 0) { Logger.log('No threads with HSA label.'); return; }

  var msg = threads[0].getMessages()[0];
  Logger.log('Subject: ' + msg.getSubject());
  Logger.log('From: ' + msg.getFrom());
  Logger.log('Attachments: ' + msg.getAttachments().length);

  var folder = DriveApp.getFolderById(folderId);
  var results = processOneMessage(msg, folder, geminiKey);

  Logger.log('=== RESULTS (not saved) ===');
  for (var i = 0; i < results.length; i++) {
    Logger.log(JSON.stringify(results[i], null, 2));
  }
}

function testGeminiConnection() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) { Logger.log('ERROR: GEMINI_API_KEY not set.'); return; }

  var payload = {
    contents: [{ parts: [{ text: 'Respond with exactly: {"status":"ok"}' }] }],
    generationConfig: { temperature: 0, maxOutputTokens: 20 }
  };

  var response = UrlFetchApp.fetch(GEMINI_URL + '?key=' + apiKey, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log('HTTP ' + response.getResponseCode());
  Logger.log(response.getContentText().substring(0, 500));
}

function checkConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  for (var key in props) {
    var val = props[key];
    if (key.indexOf('KEY') !== -1 && val.length > 8) {
      val = val.substring(0, 4) + '...' + val.substring(val.length - 4);
    }
    Logger.log(key + ' = ' + val);
  }
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Active triggers: ' + triggers.length);
}
