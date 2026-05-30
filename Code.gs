/**
 * VENDOR PAIN SURVEY — backend (Google Apps Script)
 * ---------------------------------------------------
 * SETUP (5 steps, ~3 minutes):
 *  1. Create a new Google Sheet. Name it e.g. "Vendor Survey".
 *  2. In that Sheet: Extensions > Apps Script. (This BINDS the script to the Sheet,
 *     so you don't need to paste any Sheet ID.)
 *  3. Delete whatever is in Code.gs and paste THIS file in.
 *  4. Click the + next to "Files", choose "HTML", name it exactly: Index
 *     Then paste the Index.html contents in and save.
 *  5. Deploy > New deployment > type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *     Click Deploy, copy the Web app URL. THAT is the link you DM to vendors.
 *
 * Responses land in a tab called "Responses". Read them there. Tally column Q3.
 */

var SHEET_NAME = 'Responses';

// Column order written to the sheet. Edit labels here if you change the form.
var HEADERS = [
  'Timestamp',
  'Business / page name',
  'Sells (S1)',
  'How long selling',
  'Main income or side hustle',
  'Team size',
  'Location',
  'Orders/week (S2)',
  'Main channel (S3)',
  'Uses app now (S4)',
  'Most annoying part (Q1)',
  'Lost order / fake alert? (Q2a)',
  'What happened (Q2b)',
  'Pick ONE to fix (Q3)',
  'Would pay N1000/mo? (Q3b)',
  'Biggest business stress (Q4)',
  'Magic wand fix (Q5)',
  'Contact (email / handle)'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Quick question for vendors')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Called from the form via google.script.run. Appends one row.
 * Returns true on success so the front-end can show the thank-you screen.
 */
function submitResponse(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // avoid two writes clobbering each other
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    // Write header row once.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      data.bizname || '',
      data.s1 || '',
      data.tenure || '',
      data.fulltime || '',
      data.team || '',
      data.location || '',
      data.s2 || '',
      data.s3 || '',
      data.s4 || '',
      data.q1 || '',
      data.q2a || '',
      data.q2b || '',
      data.q3 || '',
      data.q3b || '',
      data.q4 || '',
      data.q5 || '',
      data.contact || ''
    ]);
    return true;
  } finally {
    lock.releaseLock();
  }
}
