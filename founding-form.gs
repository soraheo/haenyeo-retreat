/**
 * Jeju, Slowly — application intake backend (Google Apps Script)
 * ------------------------------------------------------------------
 * Receives applications from the jejuslowly.com form, logs each one to
 * this spreadsheet, and emails Soraya a notification FROM her own Google
 * account (so it never lands in junk). The spreadsheet is the permanent
 * record — even if an email is ever missed, every application is a row here.
 *
 * SETUP (one time, ~10 min):
 *  1. Create a new Google Sheet (sheets.new). Name it e.g. "Jeju Slowly — Applications".
 *  2. Extensions → Apps Script. Delete any sample code, paste ALL of this file.
 *  3. In the CONFIG below, set NOTIFY_EMAIL to the inbox you want alerts in.
 *  4. Click Deploy → New deployment → gear icon → Web app.
 *       - Description: anything
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Deploy → Authorize (allow the permissions) → copy the Web app URL
 *     (it ends in /exec).
 *  5. Send Soraya that /exec URL — it gets pasted into the website form.
 *
 * To re-deploy after editing: Deploy → Manage deployments → edit (pencil)
 * → Version: New version → Deploy. (Keeps the same URL.)
 */

// ======================= CONFIG =======================
var CONFIG = {
  NOTIFY_EMAIL: 'soraya@vamosajeju.com',   // where alerts are sent
  SHEET_NAME: 'Applications',              // tab name inside the spreadsheet
  SEND_APPLICANT_AUTOREPLY: true,          // email the applicant a confirmation
  BRAND: 'Jeju, Slowly'
};
// ======================================================

// Column order in the sheet. Add fields here and they'll be captured too.
var FIELDS = [
  'retreat', 'name', 'email', 'nationality',
  'departure', 'guests', 'dietary', 'accessibility', 'motivation'
];

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};

    // Honeypot: silently accept but ignore obvious bot submissions.
    if (p._gotcha) {
      return json_({ ok: true });
    }

    var sheet = getSheet_();
    ensureHeader_(sheet);

    var now = new Date();
    var row = [now];
    for (var i = 0; i < FIELDS.length; i++) {
      row.push(p[FIELDS[i]] || '');
    }
    sheet.appendRow(row);

    notifySoraya_(p, now);

    if (CONFIG.SEND_APPLICANT_AUTOREPLY && isEmail_(p.email)) {
      sendAutoReply_(p);
    }

    return json_({ ok: true });
  } catch (err) {
    // Log the error to the sheet's "Errors" tab so nothing is lost silently.
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var errSheet = ss.getSheetByName('Errors') || ss.insertSheet('Errors');
      errSheet.appendRow([new Date(), String(err), JSON.stringify((e && e.parameter) || {})]);
    } catch (ignore) {}
    return json_({ ok: false, error: String(err) });
  }
}

// A GET on the URL shows a friendly note (useful to confirm the URL works).
function doGet() {
  return HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">Jeju, Slowly application endpoint is live. Submissions are accepted via POST.</p>'
  );
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(CONFIG.SHEET_NAME);
  }
  return sheet;
}

function ensureHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  var header = ['Timestamp'].concat(FIELDS);
  sheet.appendRow(header);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function notifySoraya_(p, when) {
  var subject = '[' + CONFIG.BRAND + '] New application — ' + (p.name || 'no name') +
                ' · ' + (p.retreat || 'unspecified');

  var lines = [
    'A new application just came in.',
    '',
    'Retreat:       ' + (p.retreat || '-'),
    'Name:          ' + (p.name || '-'),
    'Email:         ' + (p.email || '-'),
    'Nationality:   ' + (p.nationality || '-'),
    'Preferred date:' + (p.departure || '-'),
    'Guests:        ' + (p.guests || '-'),
    'Dietary:       ' + (p.dietary || '-'),
    'Accessibility: ' + (p.accessibility || '-'),
    '',
    'What they wrote:',
    (p.motivation || '(blank)'),
    '',
    'Received: ' + when,
    'Full log: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ];

  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL,
    subject: subject,
    body: lines.join('\n'),
    replyTo: isEmail_(p.email) ? p.email : CONFIG.NOTIFY_EMAIL,
    name: CONFIG.BRAND
  });
}

function sendAutoReply_(p) {
  var subject = 'We received your application — ' + CONFIG.BRAND;
  var body =
    'Hi ' + (p.name ? p.name.split(' ')[0] : 'there') + ',\n\n' +
    'Thank you for applying to ' + CONFIG.BRAND + '. Your application has reached us safely.\n\n' +
    'Soraya reads every application personally. For the Founding Guests pilot (Sep 10-12), ' +
    'applications close August 15 and everyone hears back by email by August 18. ' +
    'There is no payment now — your place is settled only after you are selected.\n\n' +
    'With warmth,\nSoraya\n' + CONFIG.BRAND + ' · Vamos Travel Korea\nsoraya@vamosajeju.com';

  MailApp.sendEmail({
    to: p.email,
    subject: subject,
    body: body,
    name: CONFIG.BRAND,
    replyTo: CONFIG.NOTIFY_EMAIL
  });
}

function isEmail_(s) {
  return typeof s === 'string' && /\S+@\S+\.\S+/.test(s);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
