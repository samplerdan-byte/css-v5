// ============================================================
// Config.gs — Global config, column map helper, shared utilities
// ============================================================

// 🛡️ COLUMN MAP HELPER
function _getColumnMap(sheet) {
  if (!sheet) throw new Error('_getColumnMap: sheet is null');
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });
  return col;
}

// NOTE: getOriginFromMark() and getCountryFromICOMark() are defined in
// EmailExtraction_v3.gs (more complete version with ICO_COUNTRIES table).
// NOTE: isExcludedSender() is defined in EmailExtraction_v3.gs
// (checks EXCLUDED_DOMAINS list).
// Do NOT re-define them here — duplicates cause project-wide compile failure.

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  senderEmails: ['samplerdan@gmail.com'],
  subjectKeywords: ['Sampling Order', 'sampling order', 'Samples Please'],
  mainSheetName: 'All Orders',
  scanSheetName: 'Scan Log',
  liveOrdersSheetName: 'Live Orders',
  completedOrdersSheetName: 'Completed Orders',
  dailyReportLogSheetName: 'Daily Report Log',
  labelWidth: 4,
  labelHeight: 6,
  statusValues: {
    RECEIVED: 'Received',
    SCANNED: 'Scanned',
    SHIPPED: 'Shipped',
    REPORTED: 'Reported',
    ARCHIVED: 'Archived'
  }
};

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
function buildSearchQuery() {
  let searchParts = [];
  if (CONFIG.senderEmails.length > 0) {
    CONFIG.senderEmails.forEach(email => searchParts.push(`from:${email}`));
  }
  if (CONFIG.subjectKeywords.length > 0) {
    CONFIG.subjectKeywords.forEach(kw => searchParts.push(`"${kw}"`));
  }
  let query = '';
  if (searchParts.length > 0) query = `(${searchParts.join(' OR ')})`;
  query += ' -label:PDF_Processed';
  return query;
}

function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);
  return label;
}

// ============================================================
// ERROR LOGGING
// ============================================================
function logError(functionName, errorMessage, context) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Error Log');

  if (!logSheet) {
    logSheet = ss.insertSheet('Error Log');
    logSheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Function', 'Error', 'Context', 'User']]);
    logSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f4cccc');
  }

  logSheet.appendRow([
    new Date(),
    functionName,
    errorMessage,
    JSON.stringify(context || {}),
    Session.getActiveUser().getEmail()
  ]);

  Logger.log('ERROR in ' + functionName + ': ' + errorMessage);
}

function logWarning(functionName, message, context) {
  Logger.log('WARNING in ' + functionName + ': ' + message + ' | Context: ' + JSON.stringify(context || {}));
}
