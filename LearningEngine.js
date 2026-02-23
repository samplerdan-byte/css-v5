// ============================================================
// LearningEngine.gs — Extraction Learning & Correction System
// Commodity Sampler Services
//
// How it works:
//   1. CAPTURE: When emails are processed, logExtraction() saves
//      the raw extracted values to the "Extraction Log" sheet.
//   2. LEARN: When you manually correct data in All Orders,
//      "Learn from Corrections" compares current vs. original,
//      detects what changed, and writes rules to "Correction Rules".
//   3. APPLY: applyCorrections() runs after every extraction,
//      checking rules for matching sender/field and fixing
//      values before they hit the sheet.
//
// Rule types:
//   - VALUE_MAP:  "When sender matches X and field is Y,
//                  replace value A with value B"
//   - DEFAULT:    "When sender matches X and field Y is empty,
//                  fill in value B"
//   - REGEX_MAP:  "When field Y matches regex pattern A,
//                  replace with value B"
// ============================================================

var LEARN_CONFIG = {
  logSheetName: 'Extraction Log',
  rulesSheetName: 'Correction Rules',
  minConfidence: 5,
  trackedFields: [
    'Sender', 'Receiver', 'Warehouse', 'Description',
    'Mark #', 'Container #', 'Cargo #', 'Bag Count',
    'Sample Weight', 'Sample Type', 'Shipping Process',
    'Reference', 'Sample Order #', 'Shipping Line'
  ]
};


// ============================================================
// SETUP — Create the Log and Rules sheets
// ============================================================

function setupLearningSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var logSheet = ss.getSheetByName(LEARN_CONFIG.logSheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(LEARN_CONFIG.logSheetName);
    var logHeaders = [
      'Timestamp', 'CS Sample #', 'CS Order #', 'Sender Pattern',
      'Email Subject', 'Email From'
    ].concat(LEARN_CONFIG.trackedFields.map(function(f) { return 'Raw: ' + f; }));

    logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders])
      .setFontWeight('bold').setBackground('#1a3a25').setFontColor('#fff');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 140);
  }

  var rulesSheet = ss.getSheetByName(LEARN_CONFIG.rulesSheetName);
  if (!rulesSheet) {
    rulesSheet = ss.insertSheet(LEARN_CONFIG.rulesSheetName);
    var rulesHeaders = [
      'Rule ID', 'Sender Pattern', 'Field', 'Rule Type',
      'Match Value', 'Corrected Value',
      'Times Seen', 'Auto Apply', 'Last Updated',
      'Created By', 'Notes'
    ];

    rulesSheet.getRange(1, 1, 1, rulesHeaders.length).setValues([rulesHeaders])
      .setFontWeight('bold').setBackground('#1a3a25').setFontColor('#fff');
    rulesSheet.setFrozenRows(1);
    rulesSheet.setColumnWidth(1, 80);
    rulesSheet.setColumnWidth(2, 180);
    rulesSheet.setColumnWidth(3, 120);
    rulesSheet.setColumnWidth(5, 200);
    rulesSheet.setColumnWidth(6, 200);
  }

  SpreadsheetApp.getUi().alert(
    'Learning System Ready!\n\n' +
    '• "Extraction Log" sheet created\n' +
    '• "Correction Rules" sheet created\n\n' +
    'Workflow:\n' +
    '1. Process emails as normal\n' +
    '2. Fix any wrong data in All Orders\n' +
    '3. Run "Learn from Corrections" to teach the system\n' +
    '4. Future emails auto-apply your corrections'
  );
}


// ============================================================
// CAPTURE — Log raw extracted values after processing an email
// ============================================================

function logExtraction(sampleData, emailMeta) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(LEARN_CONFIG.logSheetName);
    if (!logSheet) return;

    var senderPattern = _buildSenderPattern(
      emailMeta ? emailMeta.from : '',
      sampleData.sender || sampleData.sourceEmail || ''
    );

    var fieldMap = {
      'Sender': sampleData.sender || '',
      'Receiver': sampleData.receiver || '',
      'Warehouse': sampleData.warehouse || '',
      'Description': sampleData.description || '',
      'Mark #': sampleData.mark || '',
      'Container #': sampleData.container || '',
      'Cargo #': sampleData.cargo || '',
      'Bag Count': sampleData.bagCount || '',
      'Sample Weight': sampleData.sampleWeight || '',
      'Sample Type': sampleData.sampleType || '',
      'Shipping Process': sampleData.shippingProcess || '',
      'Reference': sampleData.reference || '',
      'Sample Order #': sampleData.sampleOrderNum || '',
      'Shipping Line': sampleData.shippingLine || ''
    };

    var row = [
      new Date(),
      sampleData.csSampleNum || '',
      sampleData.sampleOrderNum || '',
      senderPattern,
      emailMeta ? emailMeta.subject : '',
      emailMeta ? emailMeta.from : ''
    ];

    for (var f = 0; f < LEARN_CONFIG.trackedFields.length; f++) {
      row.push(fieldMap[LEARN_CONFIG.trackedFields[f]] || '');
    }

    logSheet.appendRow(row);

  } catch (e) {
    Logger.log('logExtraction error (non-fatal): ' + e);
  }
}


// ============================================================
// LEARN — Compare All Orders current values vs. Extraction Log
// ============================================================

function learnFromCorrections() {
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var logSheet = ss.getSheetByName(LEARN_CONFIG.logSheetName);
  if (!logSheet || logSheet.getLastRow() < 2) {
    ui.alert('No extraction log found. Process some emails first, then correct any errors and run this again.');
    return;
  }

  var mainName = (typeof CONFIG !== 'undefined' && CONFIG.mainSheetName) ? CONFIG.mainSheetName : 'All Orders';
  var mainSheet = ss.getSheetByName(mainName);
  if (!mainSheet || mainSheet.getLastRow() < 2) {
    ui.alert('All Orders sheet is empty.');
    return;
  }

  // Load extraction log
  var logData = logSheet.getDataRange().getValues();
  var logBySample = {};

  for (var i = 1; i < logData.length; i++) {
    var sampleId = String(logData[i][1]).trim();
    if (!sampleId) continue;
    logBySample[sampleId] = {
      senderPattern: String(logData[i][3]).trim(),
      rawValues: {},
      rowIdx: i
    };
    for (var f = 0; f < LEARN_CONFIG.trackedFields.length; f++) {
      var colIdx = 6 + f;
      logBySample[sampleId].rawValues[LEARN_CONFIG.trackedFields[f]] =
        String(logData[i][colIdx] || '').trim();
    }
  }

  // Load current sheet values
  var col = _getColumnMap(mainSheet);
  var sheetData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();

  // Compare and find corrections
  var corrections = [];

  for (var r = 0; r < sheetData.length; r++) {
    var sid = col['CS Sample #'] !== undefined ? String(sheetData[r][col['CS Sample #']] || '').trim() : '';
    if (!sid || !logBySample[sid]) continue;

    var logEntry = logBySample[sid];

    for (var fi = 0; fi < LEARN_CONFIG.trackedFields.length; fi++) {
      var fieldName = LEARN_CONFIG.trackedFields[fi];
      if (col[fieldName] === undefined) continue;

      var currentVal = String(sheetData[r][col[fieldName]] || '').trim();
      var rawVal = logEntry.rawValues[fieldName] || '';

      if (currentVal === rawVal) continue;
      if (!currentVal && !rawVal) continue;

      corrections.push({
        senderPattern: logEntry.senderPattern,
        field: fieldName,
        rawValue: rawVal,
        correctedValue: currentVal,
        sampleId: sid
      });
    }
  }

  if (corrections.length === 0) {
    ui.alert('No corrections detected.\n\nEither:\n• No changes have been made since extraction\n• Changes match what was originally extracted');
    return;
  }

  // Group corrections into rules
  var ruleMap = {};
  for (var c = 0; c < corrections.length; c++) {
    var corr = corrections[c];
    var ruleType = corr.rawValue ? 'VALUE_MAP' : 'DEFAULT';
    var key = corr.senderPattern + '|' + corr.field + '|' + ruleType + '|' + corr.rawValue;

    if (!ruleMap[key]) {
      ruleMap[key] = {
        senderPattern: corr.senderPattern,
        field: corr.field,
        ruleType: ruleType,
        matchValue: corr.rawValue,
        correctedValue: corr.correctedValue,
        count: 0,
        samples: []
      };
    }

    if (ruleMap[key].correctedValue === corr.correctedValue) {
      ruleMap[key].count++;
      ruleMap[key].samples.push(corr.sampleId);
    }
  }

  // Write/update rules
  var rulesSheet = ss.getSheetByName(LEARN_CONFIG.rulesSheetName);
  if (!rulesSheet) {
    setupLearningSystem();
    rulesSheet = ss.getSheetByName(LEARN_CONFIG.rulesSheetName);
  }

  var existingRules = _loadRules(rulesSheet);
  var newRules = 0;
  var updatedRules = 0;
  var now = new Date();
  var user = '';
  try { user = Session.getActiveUser().getEmail(); } catch (e) { /* ignore */ }

  var ruleKeys = Object.keys(ruleMap);
  for (var k = 0; k < ruleKeys.length; k++) {
    var rule = ruleMap[ruleKeys[k]];
    var existingKey = rule.senderPattern + '|' + rule.field + '|' + rule.matchValue;
    var existing = existingRules[existingKey];

    if (existing) {
      var newCount = existing.timesSeen + rule.count;
      var shouldAutoApply = newCount >= LEARN_CONFIG.minConfidence ? 'YES' : 'NO';

      // Batch the 4 adjacent column updates into one setValues call (cols 6-9)
      rulesSheet.getRange(existing.row, 6, 1, 4).setValues([[
        rule.correctedValue, newCount, shouldAutoApply, now
      ]]);
      updatedRules++;
    } else {
      var isAutoApply = rule.count >= LEARN_CONFIG.minConfidence ? 'YES' : 'NO';
      var ruleId = 'R' + String(Date.now()).slice(-6) + String(k);

      rulesSheet.appendRow([
        ruleId,
        rule.senderPattern,
        rule.field,
        rule.ruleType,
        rule.matchValue,
        rule.correctedValue,
        rule.count,
        isAutoApply,
        now,
        user,
        'Samples: ' + rule.samples.join(', ')
      ]);
      newRules++;
    }
  }

  // Count active rules
  var activeRules = 0;
  if (rulesSheet.getLastRow() > 1) {
    var allAuto = rulesSheet.getRange(2, 8, rulesSheet.getLastRow() - 1, 1).getValues();
    for (var a = 0; a < allAuto.length; a++) {
      if (String(allAuto[a][0]).trim() === 'YES') activeRules++;
    }
  }

  Logger.log('learnFromCorrections: ' + corrections.length + ' corrections, ' +
    newRules + ' new rules, ' + updatedRules + ' updated, ' + activeRules + ' active');

  ui.alert(
    'Learning Complete!\n\n' +
    corrections.length + ' correction(s) detected\n' +
    newRules + ' new rule(s) created\n' +
    updatedRules + ' existing rule(s) updated\n' +
    activeRules + ' rule(s) now auto-applying\n\n' +
    'Rules need ' + LEARN_CONFIG.minConfidence + '+ occurrences to auto-apply.\n' +
    'You can also set any rule to YES manually in the Correction Rules sheet.'
  );

  } catch (e) {
    Logger.log('learnFromCorrections error: ' + e);
    try { SpreadsheetApp.getUi().alert('Error in Learn from Corrections: ' + e); } catch(e2) {}
  }
}


// ============================================================
// APPLY — Run correction rules on freshly extracted data
//
// Usage in EmailExtraction_v3.gs:
//   logExtraction(sampleData, emailMeta);         // log RAW first
//   sampleData = applyCorrections(sampleData, senderEmail);
//   addDataToSheet(sheet, sampleData, ...);
// ============================================================

function applyCorrections(sampleData, senderEmail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rulesSheet = ss.getSheetByName(LEARN_CONFIG.rulesSheetName);
    if (!rulesSheet || rulesSheet.getLastRow() < 2) return sampleData;

    var senderPattern = _buildSenderPattern(senderEmail, sampleData.sender || '');

    var rulesData = rulesSheet.getRange(2, 1, rulesSheet.getLastRow() - 1, 8).getValues();
    var applicableRules = [];

    for (var i = 0; i < rulesData.length; i++) {
      var autoApply = String(rulesData[i][7]).trim();
      if (autoApply !== 'YES') continue;

      var rulePattern = String(rulesData[i][1]).trim().toLowerCase();
      if (!rulePattern) continue;

      var senderLower = senderPattern.toLowerCase();
      if (senderLower !== rulePattern) continue;  // exact match — prevents cross-client leakage

      applicableRules.push({
        field: String(rulesData[i][2]).trim(),
        ruleType: String(rulesData[i][3]).trim(),
        matchValue: String(rulesData[i][4]).trim(),
        correctedValue: String(rulesData[i][5]).trim()
      });
    }

    if (applicableRules.length === 0) return sampleData;

    var fieldToKey = {
      'Sender': 'sender',
      'Receiver': 'receiver',
      'Warehouse': 'warehouse',
      'Description': 'description',
      'Mark #': 'mark',
      'Container #': 'container',
      'Cargo #': 'cargo',
      'Bag Count': 'bagCount',
      'Sample Weight': 'sampleWeight',
      'Sample Type': 'sampleType',
      'Shipping Process': 'shippingProcess',
      'Reference': 'reference',
      'Sample Order #': 'sampleOrderNum',
      'Shipping Line': 'shippingLine'
    };

    var appliedCount = 0;

    for (var r = 0; r < applicableRules.length; r++) {
      var rule = applicableRules[r];
      var dataKey = fieldToKey[rule.field];
      if (!dataKey) continue;

      var currentVal = String(sampleData[dataKey] || '').trim();

      if (rule.ruleType === 'VALUE_MAP') {
        if (currentVal.toLowerCase() === rule.matchValue.toLowerCase()) {
          sampleData[dataKey] = rule.correctedValue;
          appliedCount++;
          Logger.log('Learning: ' + rule.field + ' "' + currentVal + '" -> "' + rule.correctedValue + '"');
        }
      } else if (rule.ruleType === 'DEFAULT') {
        if (!currentVal) {
          sampleData[dataKey] = rule.correctedValue;
          appliedCount++;
          Logger.log('Learning: ' + rule.field + ' (empty) -> "' + rule.correctedValue + '"');
        }
      } else if (rule.ruleType === 'REGEX_MAP') {
        try {
          var re = new RegExp(rule.matchValue, 'i');
          if (re.test(currentVal)) {
            sampleData[dataKey] = rule.correctedValue;
            appliedCount++;
            Logger.log('Learning: ' + rule.field + ' regex matched -> "' + rule.correctedValue + '"');
          }
        } catch (regexErr) {
          Logger.log('Learning: bad regex in rule: ' + rule.matchValue);
        }
      }
    }

    if (appliedCount > 0) {
      Logger.log('Learning engine applied ' + appliedCount + ' correction(s) for sender: ' + senderPattern);
    }

    return sampleData;

  } catch (e) {
    Logger.log('applyCorrections error (non-fatal): ' + e);
    return sampleData;
  }
}


// ============================================================
// VIEW — Show current rules summary in a sidebar
// ============================================================

function showLearningDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rulesSheet = ss.getSheetByName(LEARN_CONFIG.rulesSheetName);

  var totalRules = 0;
  var activeRules = 0;
  var rulesBySender = {};

  if (rulesSheet && rulesSheet.getLastRow() > 1) {
    var data = rulesSheet.getRange(2, 1, rulesSheet.getLastRow() - 1, 8).getValues();
    totalRules = data.length;

    for (var i = 0; i < data.length; i++) {
      var sender = String(data[i][1]).trim() || 'Unknown';
      var auto = String(data[i][7]).trim();
      if (auto === 'YES') activeRules++;
      if (!rulesBySender[sender]) rulesBySender[sender] = { total: 0, active: 0 };
      rulesBySender[sender].total++;
      if (auto === 'YES') rulesBySender[sender].active++;
    }
  }

  var logSheet = ss.getSheetByName(LEARN_CONFIG.logSheetName);
  var logCount = logSheet ? Math.max(0, logSheet.getLastRow() - 1) : 0;

  var senderRows = '';
  var senders = Object.keys(rulesBySender).sort();
  for (var s = 0; s < senders.length; s++) {
    var info = rulesBySender[senders[s]];
    senderRows += '<tr><td>' + senders[s] + '</td>' +
      '<td style="text-align:center">' + info.total + '</td>' +
      '<td style="text-align:center">' + info.active + '</td></tr>';
  }

  var htmlStr = '<style>' +
    'body{font-family:Arial,sans-serif;padding:15px;background:#f5f5f5}' +
    '.card{background:#fff;border-radius:8px;padding:15px;margin-bottom:12px;box-shadow:0 1px 4px rgba(0,0,0,0.1)}' +
    'h2{color:#2E5339;font-size:16px;margin:0 0 10px}' +
    '.stats{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:15px}' +
    '.stat{background:#e8f5e9;padding:12px;border-radius:8px;text-align:center}' +
    '.stat .num{font-size:24px;font-weight:bold;color:#2E5339}' +
    '.stat .label{font-size:10px;color:#666;text-transform:uppercase}' +
    'table{width:100%;border-collapse:collapse;font-size:11px}' +
    'th{background:#2E5339;color:#fff;padding:6px 8px;text-align:left}' +
    'td{padding:5px 8px;border-bottom:1px solid #eee}' +
    '.btn{display:block;width:100%;padding:12px;font-size:13px;font-weight:bold;border:none;border-radius:8px;cursor:pointer;margin-bottom:6px}' +
    '.btn-learn{background:#2E5339;color:#fff}' +
    '.btn-learn:hover{background:#1e4620}' +
    '.btn-view{background:#fff;color:#2E5339;border:2px solid #2E5339}' +
    '.btn-view:hover{background:#e8f5e9}' +
    '.note{font-size:10px;color:#999;text-align:center;margin-top:8px}' +
    '</style>';

  htmlStr += '<div class="card">' +
    '<h2>Learning Dashboard</h2>' +
    '<div class="stats">' +
    '<div class="stat"><div class="num">' + logCount + '</div><div class="label">Extractions Logged</div></div>' +
    '<div class="stat"><div class="num">' + totalRules + '</div><div class="label">Total Rules</div></div>' +
    '<div class="stat"><div class="num">' + activeRules + '</div><div class="label">Auto-Applying</div></div>' +
    '</div></div>';

  htmlStr += '<div class="card"><h2>Rules by Sender</h2>';
  if (senderRows) {
    htmlStr += '<table><thead><tr><th>Sender</th><th>Rules</th><th>Active</th></tr></thead>' +
      '<tbody>' + senderRows + '</tbody></table>';
  } else {
    htmlStr += '<p style="color:#999;text-align:center;font-size:12px;">No rules yet. Process emails, correct errors, then run Learn.</p>';
  }
  htmlStr += '</div>';

  htmlStr += '<div class="card">' +
    '<button class="btn btn-learn" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close()}).learnFromCorrections()">Learn from Corrections</button>' +
    '<button class="btn btn-view" onclick="google.script.run.showRulesSheet()">View Rules Sheet</button>' +
    '<div class="note">Rules auto-apply after ' + LEARN_CONFIG.minConfidence + '+ matching corrections.<br>Override manually by setting Auto Apply to YES in the rules sheet.</div>' +
    '</div>';

  var html = HtmlService.createHtmlOutput(htmlStr)
    .setWidth(380).setHeight(550).setTitle('Learning System');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showRulesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LEARN_CONFIG.rulesSheetName);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Correction Rules sheet not found. Run Setup first.');
  }
}


// ============================================================
// RESET — Clear learned corrections for a specific sender
// ============================================================

function resetSenderRules() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Reset Sender Rules',
    'Enter sender pattern to clear rules for (e.g. "louis dreyfus"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var pattern = response.getResponseText().trim().toLowerCase();
  if (!pattern) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rulesSheet = ss.getSheetByName(LEARN_CONFIG.rulesSheetName);
  if (!rulesSheet || rulesSheet.getLastRow() < 2) {
    ui.alert('No rules to clear.');
    return;
  }

  var data = rulesSheet.getRange(2, 2, rulesSheet.getLastRow() - 1, 1).getValues();
  var deleted = 0;

  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim().toLowerCase().indexOf(pattern) !== -1) {
      rulesSheet.deleteRow(i + 2);
      deleted++;
    }
  }

  ui.alert('Cleared ' + deleted + ' rule(s) matching "' + pattern + '".');
}


// ============================================================
// HELPERS
// ============================================================

function _buildSenderPattern(emailFrom, senderName) {
  var name = String(senderName || '').trim();
  if (name) {
    name = name.replace(/\b(company|corp|corporation|llc|ltd|inc|group|trading|commodities)\b/gi, '')
               .replace(/[.]/g, '')
               .replace(/\s+/g, ' ').trim();
    if (name.length >= 3) return name.toLowerCase();
  }

  var email = String(emailFrom || '').trim().toLowerCase();
  var match = email.match(/@([^.]+)\./);
  if (match) return match[1];

  return email || 'unknown';
}

function _loadRules(rulesSheet) {
  var rules = {};
  if (!rulesSheet || rulesSheet.getLastRow() < 2) return rules;

  var data = rulesSheet.getRange(2, 1, rulesSheet.getLastRow() - 1, 11).getValues();
  for (var i = 0; i < data.length; i++) {
    var sender = String(data[i][1]).trim().toLowerCase();
    var field = String(data[i][2]).trim();
    var matchVal = String(data[i][4]).trim();
    var key = sender + '|' + field + '|' + matchVal;

    rules[key] = {
      row: i + 2,
      ruleId: String(data[i][0]).trim(),
      correctedValue: String(data[i][5]).trim(),
      timesSeen: parseInt(data[i][6], 10) || 0,
      autoApply: String(data[i][7]).trim()
    };
  }
  return rules;
}
