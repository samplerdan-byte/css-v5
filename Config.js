// ============================================================
// Config.gs — Global config, column map helper, shared utilities
// ============================================================

// 🛡️ COLUMN MAP HELPER (cached per execution — GAS resets globals between runs)
var _colMapCache = {};

function _getColumnMap(sheet) {
  if (!sheet) throw new Error('_getColumnMap: sheet is null');
  var name = sheet.getName();
  if (_colMapCache[name]) return _colMapCache[name];
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });
  _colMapCache[name] = col;
  return col;
}

function _clearColumnMapCache() {
  _colMapCache = {};
}

// 🛡️ COLUMN LETTER HELPER (handles past Z: AA, AB, etc.)
function _colToLetter(col) {
  var letter = '';
  var c = col;
  while (c >= 0) {
    letter = String.fromCharCode(65 + (c % 26)) + letter;
    c = Math.floor(c / 26) - 1;
  }
  return letter;
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
  var searchParts = [];

  // Include all sender emails from CONFIG
  if (CONFIG.senderEmails && CONFIG.senderEmails.length > 0) {
    CONFIG.senderEmails.forEach(function(email) { searchParts.push('from:' + email); });
  }

  // Include all subject keywords from CONFIG
  if (CONFIG.subjectKeywords && CONFIG.subjectKeywords.length > 0) {
    CONFIG.subjectKeywords.forEach(function(kw) { searchParts.push('"' + kw + '"'); });
  }

  // Include all KNOWN_CLIENTS domains (the big list in Emailextraction_v3)
  if (typeof KNOWN_CLIENTS !== 'undefined') {
    var seenDomains = {};
    var clientKeys = Object.keys(KNOWN_CLIENTS);
    for (var i = 0; i < clientKeys.length; i++) {
      var client = KNOWN_CLIENTS[clientKeys[i]];
      if (client.domains) {
        for (var d = 0; d < client.domains.length; d++) {
          var domain = client.domains[d].replace('@', '');
          if (!seenDomains[domain]) {
            seenDomains[domain] = true;
            searchParts.push('from:' + domain);
          }
        }
      }
    }
  }

  // Add more keyword catches
  searchParts.push('"SAMPLE ORDER"');
  searchParts.push('"Sample Allowance"');
  searchParts.push('"ARRIVAL SAMPLE REQUEST"');
  searchParts.push('"sampling instructions"');
  searchParts.push('"send samples"');
  searchParts.push('"LETTER OF ENTRY"');
  searchParts.push('"Delivery Order"');
  searchParts.push('"Pier to Whse"');

  var query = '';
  if (searchParts.length > 0) query = '(' + searchParts.join(' OR ') + ')';
  query += ' -label:PDF_Processed';
  query += ' after:2026/02/15';
  return query;
}

function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);
  return label;
}

/**
 * One-time utility: strips PDF_Processed label from threads after a given date
 * so they get reprocessed on the next trigger run.
 * Run manually from the script editor, then delete when done.
 */
function reprocessEmailsFrom() {
  var label = GmailApp.getUserLabelByName('PDF_Processed');
  if (!label) { Logger.log('No PDF_Processed label found'); return; }

  var query = 'label:PDF_Processed after:2026/02/15';
  var threads = GmailApp.search(query, 0, 100);
  Logger.log('Found ' + threads.length + ' threads to reprocess');

  for (var i = 0; i < threads.length; i++) {
    threads[i].removeLabel(label);
    Logger.log('Unlabeled: ' + threads[i].getFirstMessageSubject());
  }

  Logger.log('Done. ' + threads.length + ' threads will be reprocessed on next run.');
}

// logError() and logWarning() — moved to Errorlog.js to avoid duplicates
