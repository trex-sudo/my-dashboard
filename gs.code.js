// ─────────────────────────────────────────────────────────────
// DMCI POWER - MAINTENANCE FORMS PORTAL
// Google Apps Script — Code.gs
// ─────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Maintenance Forms')
    .addItem('🏠 Open Forms Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('I&C Routine Checklist — Feedwater', 'showFeedwaterForm')
    // .addItem('I&C Routine Checklist — Steam', 'showSteamForm')
    .addToUi();
}

function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, 'DMCI Power — Maintenance Forms Portal');
}

function showFeedwaterForm() {
  const html = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, 'I&C Routine Checklist — Feedwater System');
}

// ── Add future form openers below ──
// function showSteamForm() {
//   const html = HtmlService.createHtmlOutputFromFile('SteamForm')
//     .setWidth(1200).setHeight(800);
//   SpreadsheetApp.getUi().showModelessDialog(html, 'I&C Routine Checklist — Steam System');
// }

/**
 * Fetches technician names from Sheet5 A1:A7.
 */
function getTechnicians() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet5');
  if (!sheet) return [];
  return sheet.getRange('A1:A7').getValues()
    .map(r => r[0]).filter(n => n !== '' && n !== null);
}

/**
 * Fetches instrument list from a named sheet tab.
 */
function getInstruments(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({ item: r[0], device: r[1], serial: r[2], meas: r[3], qty: r[4] }));
}

/**
 * Saves the filled checklist as a PDF in Google Drive.
 * Creates a folder "DMCI Maintenance Forms" if it doesn't exist.
 * Returns the PDF file URL so the form can show a link.
 *
 * @param {Object} data - All form field values + rows from Form.html
 * @returns {Object} { url: '...', filename: '...' }
 */
function saveAsPDF(data) {
  if (!data || !data.rows) {
    throw new Error('No form data received.');
  }

  function chk(val) { return val ? '\u2713' : ''; }

  // ── 1. Create a temporary Google Doc ──
  const docName = 'DMCI_TEMP_' + new Date().getTime();
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  body.setPageWidth(720).setPageHeight(540); // landscape-ish
  body.setMarginTop(36).setMarginBottom(36)
      .setMarginLeft(36).setMarginRight(36);
  body.clear();

  // ── HEADER ──
  const titleStyle = {};
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  titleStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

  body.appendParagraph('DMCI POWER CORPORATION').setAttributes(titleStyle);

  const subStyle = {};
  subStyle[DocumentApp.Attribute.BOLD] = true;
  subStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  subStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  body.appendParagraph('MAINTENANCE DATA SHEET - ROUTINE CHECKLIST').setAttributes(subStyle);

  const metaStyle = {};
  metaStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
  metaStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  body.appendParagraph('Document ID: DPCP-MAINT-I&C-RC-FEEDWATER SYSTEM   |   Doc No.: ' + (data.docNo || '')).setAttributes(metaStyle);
  body.appendParagraph('─────────────────────────────────────────────────────────────────────────────────').setAttributes(metaStyle);

  // ── INFO ROW ──
  const infoStyle = {};
  infoStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
  body.appendParagraph(
    'System: ' + (data.system || '') +
    '   |   Date Scheduled: ' + (data.dateScheduled || '') +
    '   |   Date Performed: ' + (data.datePerformed || '') +
    '   |   Site: ' + (data.siteLocation || '') +
    '   |   Manpower: ' + (data.manpower || '')
  ).setAttributes(infoStyle);
  body.appendParagraph('─────────────────────────────────────────────────────────────────────────────────').setAttributes(metaStyle);

  // ── LIST OF INSTRUMENTS HEADING ──
  const secStyle = {};
  secStyle[DocumentApp.Attribute.BOLD] = true;
  secStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
  secStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  body.appendParagraph('LIST OF INSTRUMENTS').setAttributes(secStyle);

  // ── INSTRUMENTS TABLE ──
  const headers = [
    'Item\nNo.', 'Device Name', 'Serial / ID No.',
    'Measuring / Sensing Point', 'Qty &\nUnit',
    'Inspect\nTerm.', 'Retighten\n/ Thread', 'Cleaning\nDisplay',
    'Local\nDisplay\nvs CCR', 'Calibration', 'Photo\nof Device', 'Remarks'
  ];

  const numCols = headers.length;
  const table = body.appendTable();
  table.setBorderWidth(1);

  // Header row
  const hRow = table.appendTableRow();
  headers.forEach(function(h) {
    const cell = hRow.appendTableCell(h);
    cell.setBackgroundColor('#1a3a6b');
    const txt = cell.editAsText();
    txt.setForegroundColor('#ffffff');
    txt.setBold(true);
    txt.setFontSize(7);
    cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(3).setPaddingRight(3);
  });

  // Data rows
  data.rows.forEach(function(row, idx) {
    const tr = table.appendTableRow();
    const bg = (idx % 2 === 0) ? '#ffffff' : '#f5f7fb';
    const cells = [
      String(row.itemNo || ''),
      row.device || '',
      row.serial || '',
      row.meas || '',
      row.qty || '',
      chk(row.inspectTermination),
      chk(row.retightening),
      chk(row.cleaningDisplay),
      chk(row.localDisplay),
      chk(row.calibration),
      chk(row.photoDevice),
      row.remarks || ''
    ];
    cells.forEach(function(val) {
      const cell = tr.appendTableCell(val);
      cell.setBackgroundColor(bg);
      cell.editAsText().setFontSize(7);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(3).setPaddingRight(3);
    });
  });

  // ── SIGN OFF ──
  body.appendParagraph('─────────────────────────────────────────────────────────────────────────────────').setAttributes(metaStyle);
  body.appendParagraph('CERTIFICATION / SIGN-OFF').setAttributes(secStyle);

  const sigTable = body.appendTable();
  sigTable.setBorderWidth(1);
  const sigRow = sigTable.appendTableRow();

  ['Prepared By', 'Checked By', 'Approved By'].forEach(function(label, i) {
    const names = [data.preparedBy, data.checkedBy, data.approvedBy];
    const dates = [data.preparedDate, data.checkedDate, data.approvedDate];
    const cell = sigRow.appendTableCell(
      label + ':\n' + (names[i] || '________________________') +
      '\n\nDate: ' + (dates[i] || '________________________')
    );
    cell.editAsText().setFontSize(8);
    cell.setPaddingTop(6).setPaddingBottom(10).setPaddingLeft(6).setPaddingRight(6);
  });

  doc.saveAndClose();

  // ── 2. Export the Doc as PDF ──
  const docId = doc.getId();
  const token = ScriptApp.getOAuthToken();
  const exportUrl = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  // Delete temp doc
  DriveApp.getFileById(docId).setTrashed(true);

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: HTTP ' + response.getResponseCode());
  }

  // ── 3. Save PDF to Drive folder ──
  const folderName = 'DMCI Maintenance Forms';
  let folder;
  const folderSearch = DriveApp.getFoldersByName(folderName);
  if (folderSearch.hasNext()) {
    folder = folderSearch.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  const date = data.datePerformed || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd-yyyy');
  const filename = 'DMCI-IC-RC-Feedwater-' + date.replace(/\//g, '-') + '.pdf';
  const pdfFile = folder.createFile(response.getBlob().setName(filename));

  // Make it accessible to anyone with the link
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    url: pdfFile.getUrl(),
    filename: filename,
    folderName: folderName
  };
}

/**
 * Run this directly from Apps Script editor to test Drive + Docs access.
 * Check execution log for results.
 */
function testPDF() {
  try {
    Logger.log('Step 1: Creating test doc...');
    const doc = DocumentApp.create('DMCI_TEST_' + new Date().getTime());
    const body = doc.getBody();
    body.appendParagraph('DMCI POWER CORPORATION - TEST PDF');
    body.appendParagraph('If you can see this, the PDF generation works!');
    doc.saveAndClose();
    Logger.log('Step 2: Doc created. ID = ' + doc.getId());

    const docId = doc.getId();
    const token = ScriptApp.getOAuthToken();
    const url = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';
    Logger.log('Step 3: Exporting PDF from ' + url);

    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    Logger.log('Step 4: Export response code = ' + response.getResponseCode());

    DriveApp.getFileById(docId).setTrashed(true);
    Logger.log('Step 5: Temp doc deleted.');

    if (response.getResponseCode() !== 200) {
      Logger.log('FAILED: ' + response.getContentText().substring(0, 300));
      return;
    }

    Logger.log('Step 6: Saving PDF to Drive folder...');
    let folder;
    const search = DriveApp.getFoldersByName('DMCI Maintenance Forms');
    folder = search.hasNext() ? search.next() : DriveApp.createFolder('DMCI Maintenance Forms');

    const pdfFile = folder.createFile(response.getBlob().setName('DMCI-TEST.pdf'));
    Logger.log('Step 7: SUCCESS! File saved: ' + pdfFile.getUrl());

  } catch(e) {
    Logger.log('ERROR at step: ' + e.message);
  }
}
