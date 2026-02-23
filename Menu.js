// ============================================================
// Menu.gs — onOpen menu, triggers
// ============================================================

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    var role = _getUserRole();

    // ── EDITOR: Full access (dan@commoditysampler.com) ──
    if (role === 'editor') {
      _buildEditorMenus(ui);
      _applySheetVisibility(role);
    }
    // ── BILLING: samplerdan@gmail.com ──
    else if (role === 'billing') {
      _buildBillingMenus(ui);
      _applySheetVisibility(role);
    }
    // ── FIELD: danboy1217@gmail.com ──
    else if (role === 'field') {
      _buildFieldMenus(ui);
      _applySheetVisibility(role);
    }
    // ── UNKNOWN: read-only, no menus ──
    else {
      ui.createMenu('☕ CSS System')
        .addItem('🔒 No access — contact Dan', '_noAccess')
        .addToUi();
      _applySheetVisibility('none');
    }

    Logger.log('CSS Menu created for role: ' + role);

  } catch (e) {
    Logger.log('onOpen MENU ERROR: ' + e.message + ' | Stack: ' + e.stack);
    try {
      SpreadsheetApp.getUi().createMenu('☕ CSS (ERROR)')
        .addItem('⚠️ Menu failed: ' + e.message, 'listTriggers')
        .addToUi();
    } catch (e2) { /* give up */ }
  }
}

// ============================================================
// TRIGGERS
// ============================================================

function createTimeTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var toDelete = [];

    // Collect existing triggers to delete
    if (triggers && Array.isArray(triggers)) {
      triggers.forEach(function(trigger) {
        if (!trigger) return;
        try {
          var handlerFunc = trigger.getHandlerFunction();
          if (handlerFunc === 'processPDFsFromGmail' ||
              handlerFunc === 'scheduledBackup' ||
              handlerFunc === 'processPendingMoves') {
            toDelete.push(trigger);
          }
        } catch (e) {
          Logger.log('Could not inspect trigger: ' + e.message);
        }
      });
    }

    // Delete collected triggers
    for (var i = 0; i < toDelete.length; i++) {
      try {
        ScriptApp.deleteTrigger(toDelete[i]);
      } catch (e) {
        Logger.log('Could not delete trigger: ' + e.message);
      }
    }

    // Create new triggers with error handling
    try {
      ScriptApp.newTrigger('processPDFsFromGmail')
        .timeBased().everyMinutes(15).create();
    } catch (e) {
      Logger.log('Failed to create processPDFsFromGmail trigger: ' + e.message);
      throw e;
    }

    try {
      ScriptApp.newTrigger('scheduledBackup')
        .timeBased().atHour(23).everyDays(1).create();
    } catch (e) {
      Logger.log('Failed to create scheduledBackup trigger: ' + e.message);
      throw e;
    }

    try {
      ScriptApp.newTrigger('processPendingMoves')
        .timeBased().everyHours(1).create();
    } catch (e) {
      Logger.log('Failed to create processPendingMoves trigger: ' + e.message);
      throw e;
    }

    Logger.log('Time triggers created');
    SpreadsheetApp.getUi().alert('Triggers Created!\n\n• Email processing: every 15 min\n• Backup: daily at 11 PM\n• Move shipped orders: hourly');
  } catch (e) {
    Logger.log('createTimeTriggers error: ' + e.message);
    SpreadsheetApp.getUi().alert('Error creating triggers: ' + e.message);
  }
}

function deleteAllTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var deleted = 0;
    var failed = 0;

    if (triggers && Array.isArray(triggers)) {
      triggers.forEach(function(trigger) {
        if (!trigger) return;
        try {
          ScriptApp.deleteTrigger(trigger);
          deleted++;
        } catch (e) {
          Logger.log('Could not delete trigger: ' + e.message);
          failed++;
        }
      });
    }

    Logger.log('All triggers deletion: deleted=' + deleted + ', failed=' + failed);
    SpreadsheetApp.getUi().alert('Trigger deletion complete.\n\nDeleted: ' + deleted + '\nFailed: ' + failed);
  } catch (e) {
    Logger.log('deleteAllTriggers error: ' + e.message);
    SpreadsheetApp.getUi().alert('Error deleting triggers: ' + e.message);
  }
}

function listTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var info = 'Current Triggers:\n\n';

    if (!triggers || triggers.length === 0) {
      info += 'No triggers set up.';
    } else {
      triggers.forEach(function(trigger) {
        if (trigger) info += '• ' + trigger.getHandlerFunction() + '\n';
      });
    }

    SpreadsheetApp.getUi().alert(info);
  } catch (e) {
    Logger.log('listTriggers error: ' + e.message);
    SpreadsheetApp.getUi().alert('Error listing triggers: ' + e.message);
  }
}


// ============================================================
// ROLE-BASED ACCESS
// ============================================================

var USER_ROLES = {
  'dan@commoditysampler.com': 'editor',
  'samplerdan@gmail.com':     'billing',
  'danboy1217@gmail.com':     'field'
};

function _getUserRole() {
  try {
    var email = '';
    try {
      var activeUser = Session.getActiveUser();
      if (activeUser) {
        email = String(activeUser.getEmail()).toLowerCase().trim();
      }
    } catch (e) {
      Logger.log('getActiveUser failed: ' + e.message);
    }

    // getActiveUser() returns empty in simple triggers — fall back to effective user
    if (!email) {
      try {
        var effectiveUser = Session.getEffectiveUser();
        if (effectiveUser) {
          email = String(effectiveUser.getEmail()).toLowerCase().trim();
        }
      } catch (e) {
        Logger.log('getEffectiveUser failed: ' + e.message);
      }
    }

    // If still empty, fail-closed to 'none' instead of defaulting to editor
    if (!email || email.length === 0) {
      Logger.log('Could not determine user email — denying access');
      return 'none';
    }

    Logger.log('User detected: ' + email);
    var roles = _loadUserRoles();
    if (!roles || typeof roles !== 'object') {
      Logger.log('No user roles config found');
      return 'none';
    }
    return roles[email] || 'none';
  } catch (e) {
    Logger.log('Could not get user email: ' + e.message);
    return 'none'; // fail-closed
  }
}

function _noAccess() {
  SpreadsheetApp.getUi().alert(
    'Access Denied\n\n' +
    'Your account does not have access to this system.\n' +
    'Contact Dan at dan@commoditysampler.com.'
  );
}


// ============================================================
// EDITOR MENUS (dan@commoditysampler.com — full access)
// ============================================================

function _buildEditorMenus(ui) {

  ui.createMenu('☕ CSS System')

    // === TOP-LEVEL: Core daily actions ===
    .addItem('📧 Process New Emails', 'processPDFsFromGmail')
    .addItem('🔄 Full Refresh', 'fullRefresh')
    .addItem('✍️ Manual Order Entry', 'showEnhancedManualEntry')
    .addItem('📋 View Order Details', 'viewOrderDetails')
    .addItem('🤖 AI Assistant', 'openAiAssistant')
    .addSeparator()

    // === SCAN IN ===
    .addSubMenu(ui.createMenu('📷 Scan In')
      .addItem('📷 Barcode Scanner', 'openBarcodeScanner')
      .addItem('📱 Web Scanner', 'openWebScanner')
      .addItem('📸 Sample Photos', 'showPhotoUpload'))
    .addSeparator()

    // === PRINT ===
    .addSubMenu(ui.createMenu('🖨️ Print')
      .addItem('☑️ Print Checked Samples', 'printCheckedSamples')
      .addItem('🔄 Reprint Label', 'reprintLabel')
      .addItem('📋 Today\'s Cover Sheets', 'printTodaysCoverSheets')
      .addItem('🏷️ Generate Labels', 'generateLabels'))
    .addSeparator()

    // === REPORTS ===
    .addSubMenu(ui.createMenu('📊 Reports')
      .addItem('📊 Dashboard', 'refreshDashboard')
      .addItem('📈 Trend Data', 'generateTrendData')
      .addSeparator()
      .addItem('📋 Open Orders', 'generateOpenOrdersReport')
      .addItem('🏭 Field Report', 'generateFieldReport')
      .addItem('📄 Printable Order Forms', 'generatePrintableOrderForm')
      .addItem('📈 End of Day', 'generateEndOfDayReport')
      .addItem('📥 Import Offline Edits', 'showImportOfflineEdits')
      .addSeparator()
      .addItem('📈 Market Intel', 'buildMarketIntelSheet'))
    .addSeparator()

    // === ORDER MANAGEMENT ===
    .addSubMenu(ui.createMenu('📦 Orders')
      .addItem('🔄 Refresh Live Orders', 'updateLiveOrdersView')
      .addItem('📦 Move Shipped → Completed', 'processPendingMovesWithAlert')
      .addItem('🗄️ Archive Old (90+ days)', 'archiveOldOrders')
      .addItem('🔧 Repair Receivers', 'repairReceivers')
      .addSeparator()
      .addItem('📧 Send Customer Reports', 'sendDailyCustomerReports')
      .addItem('🗑️ Delete Selected Rows', 'deleteSelectedRows'))
    .addSeparator()

    // === EXPORT ===
    .addSubMenu(ui.createMenu('📤 Export')
      .addItem('📅 Today\'s Orders', 'exportTodaysOrders')
      .addItem('📊 All Orders', 'exportAllOrders')
      .addItem('🗄️ Export for SQL', 'exportForSQL'))

    // === MY VIEW ===
    .addSubMenu(ui.createMenu('👤 My View')
      .addItem('➕ Create My View', 'createMyView')
      .addItem('🔄 Refresh My View', 'refreshMyView'))
    .addSeparator()

    // === ADMIN ===
    .addSubMenu(ui.createMenu('⚙️ Admin')
      .addItem('🚀 Setup Everything', 'setupEverything')
      .addSeparator()
      .addItem('🛠️ Initial Setup', 'setupSheet')
      .addItem('📊 Setup Tracking', 'setupTrackingSystem')
      .addItem('🏭 Setup Field Columns', 'setupFieldColumns')
      .addItem('📇 Setup Contacts', 'setupContacts')
      .addItem('📧 Setup Customer Emails', 'setupCustomerEmailSheet')
      .addItem('🏭 Setup Warehouse Emails', 'setupWarehouseEmails')
      .addItem('✏️ Setup Field Report Editors', 'setupFieldReportEditors')
      .addSeparator()
      .addItem('🔧 Fix Print Checkboxes', 'fixAllPrintCheckboxes')
      .addItem('🔄 Sync Print Column', 'syncPrintColumnToAllSheets')
      .addItem('🎨 Conditional Formatting', 'setupConditionalFormatting')
      .addSeparator()
      .addItem('🔒 Lock Headers', 'lockAllHeadersWithAlert')
      .addItem('🔓 Unlock Headers', 'unlockHeaders')
      .addSeparator()
      .addItem('🏷️ Bulk Label Old Emails', 'bulkLabelProcessed')
      .addItem('🛑 Cancel Historical Import', 'cancelHistoricalImport')
      .addItem('💾 Create Backup', 'createBackup')
      .addItem('🗑️ Clear Cover Sheet Flags', 'clearCoverSheetPrintedFlags'))
    .addSeparator()
    .addItem('🔐 Manage Sheet Access', 'showAccessManager')
    .addItem('👥 Manage Menu Roles', 'showRoleManager')
    .addSeparator()

    // === CLIENT PORTAL ===
    .addSubMenu(ui.createMenu('🌐 Client Portal')
      .addItem('⚙️ Setup Portal', 'setupPortalAccess')
      .addItem('🔑 Generate Token', 'generatePortalToken')
      .addItem('🔗 List Portal Links', 'listPortalLinks')
      .addItem('🚫 Revoke Access', 'revokePortalToken'))

    // === BRIEFING ===
    .addSubMenu(ui.createMenu('☀️ Briefing')
      .addItem('📧 Send Now', 'sendMorningBriefing')
      .addItem('⚙️ Setup Daily (7 AM)', 'setupMorningBriefing')
      .addItem('🚫 Remove Trigger', 'removeMorningBriefing'))

    // === LEARNING ===
    .addSubMenu(ui.createMenu('🧠 Learning')
      .addItem('⚙️ Setup', 'setupLearningSystem')
      .addItem('🧠 Learn from Corrections', 'learnFromCorrections')
      .addItem('📊 Dashboard', 'showLearningDashboard')
      .addItem('🗑️ Reset Rules', 'resetSenderRules'))

    .addToUi();

  // === SCAN OUT — top-level menu ===
  ui.createMenu('📦 Scan Out')
    .addItem('📦 Ship Samples', 'showScanOut')
    .addItem('📦 Move Shipped → Completed', 'processPendingMovesWithAlert')
    .addSeparator()
    .addItem('🏷️ Generate Label (EasyPost)', 'showLabelGenerator')
    .addItem('🔍 Track Package', 'showTrackPackage')
    .addSeparator()
    .addItem('⚙️ Setup EasyPost', 'setupEasyPost')
    .addItem('✅ Test EasyPost Connection', 'testEasyPostConnection')
    .addToUi();

  // === BILLING — top-level menu ===
  ui.createMenu('💰 Billing')
    .addItem('Generate Invoice', 'generateInvoice')
    .addItem('Preview Invoice', 'previewInvoice')
    .addSeparator()
    .addItem('Edit Line Item', 'editLineItemService')
    .addItem('Mark Sent', 'markInvoiceSent')
    .addItem('Mark Paid', 'markInvoicePaid')
    .addItem('Void Invoice', 'markInvoiceVoid')
    .addSeparator()
    .addItem('Export to QB Online (CSV)', 'exportInvoicesQBO')
    .addItem('Export to QB Desktop (IIF)', 'exportInvoicesIIF')
    .addSeparator()
    .addItem('Dashboard', 'showBillingDashboard')
    .addItem('Setup Billing', 'setupBilling')
    .addToUi();

  // === HISTORICAL IMPORT — top-level menu ===
  ui.createMenu('📜 Historical Import')
    .addItem('▶️ Start Import (4+ Years)', 'historicalImport')
    .addItem('🔄 Restart Import (Fresh)', 'restartHistoricalImport')
    .addItem('🛑 Cancel Import', 'cancelHistoricalImport')
    .addToUi();
}


// ============================================================
// BILLING MENUS (samplerdan@gmail.com — billing + user view)
// ============================================================

function _buildBillingMenus(ui) {

  // === BILLING — full billing access ===
  ui.createMenu('💰 Billing')
    .addItem('Generate Invoice', 'generateInvoice')
    .addItem('Preview Invoice', 'previewInvoice')
    .addSeparator()
    .addItem('Edit Line Item', 'editLineItemService')
    .addItem('Mark Sent', 'markInvoiceSent')
    .addItem('Mark Paid', 'markInvoicePaid')
    .addItem('Void Invoice', 'markInvoiceVoid')
    .addSeparator()
    .addItem('Export to QB Online (CSV)', 'exportInvoicesQBO')
    .addItem('Export to QB Desktop (IIF)', 'exportInvoicesIIF')
    .addSeparator()
    .addItem('Dashboard', 'showBillingDashboard')
    .addToUi();

  // === MY VIEW — read-only view ===
  ui.createMenu('👤 My View')
    .addItem('➕ Create My View', 'createMyView')
    .addItem('🔄 Refresh My View', 'refreshMyView')
    .addSeparator()
    .addItem('📋 View Order Details', 'viewOrderDetails')
    .addToUi();
}


// ============================================================
// FIELD MENUS (danboy1217@gmail.com — field report + user view)
// ============================================================

function _buildFieldMenus(ui) {

  // === FIELD REPORT — edit + import ===
  ui.createMenu('🏭 Field Report')
    .addItem('🏭 Edit Field Report', 'generateFieldReport')
    .addItem('📥 Import Offline Edits', 'showImportOfflineEdits')
    .addToUi();

  // === MY VIEW — read-only view ===
  ui.createMenu('👤 My View')
    .addItem('➕ Create My View', 'createMyView')
    .addItem('🔄 Refresh My View', 'refreshMyView')
    .addSeparator()
    .addItem('📋 View Order Details', 'viewOrderDetails')
    .addToUi();
}


// ============================================================
// SHEET ACCESS MANAGER (editor-only admin UI)
// ============================================================

/**
 * Opens the access manager modal — checkboxes to set View/Edit per user per sheet.
 * Only callable by editor role.
 */
function showAccessManager() {
  var role = _getUserRole();
  if (role !== 'editor') {
    SpreadsheetApp.getUi().alert('Access Denied — Editor only.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().map(function(s) { return s.getName(); });
  var config = _loadAccessConfig();

  // Build user list from USER_ROLES (exclude editor)
  var users = [];
  for (var email in USER_ROLES) {
    if (USER_ROLES[email] !== 'editor') {
      users.push({ email: email, role: USER_ROLES[email] });
    }
  }

  var html = '<html><head><style>' +
    'body { font-family: Arial, sans-serif; margin: 0; padding: 16px; background: #f8f9fa; }' +
    'h2 { margin: 0 0 4px 0; font-size: 18px; }' +
    '.subtitle { color: #666; font-size: 12px; margin-bottom: 16px; }' +
    'table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }' +
    'th { background: #2d5016; color: #fff; padding: 8px 6px; font-size: 11px; text-align: center; position: sticky; top: 0; }' +
    'th:first-child { text-align: left; min-width: 160px; }' +
    'td { padding: 6px; border-bottom: 1px solid #eee; font-size: 12px; text-align: center; }' +
    'td:first-child { text-align: left; font-weight: 500; }' +
    'tr:hover { background: #f0f7e8; }' +
    '.user-col { min-width: 100px; }' +
    '.cb-group { display: flex; gap: 8px; justify-content: center; align-items: center; }' +
    '.cb-group label { font-size: 10px; color: #555; cursor: pointer; display: flex; align-items: center; gap: 2px; }' +
    '.cb-group input { margin: 0; cursor: pointer; }' +
    '.btn-row { margin-top: 16px; text-align: right; }' +
    '.btn { padding: 8px 20px; border: none; border-radius: 6px; font-size: 13px; cursor: pointer; font-weight: 600; }' +
    '.btn-save { background: #2d5016; color: #fff; }' +
    '.btn-save:hover { background: #3d6b1e; }' +
    '.btn-cancel { background: #e0e0e0; color: #333; margin-right: 8px; }' +
    '.btn-cancel:hover { background: #ccc; }' +
    '.saving { color: #2d5016; font-weight: 600; display: none; margin-right: 12px; }' +
    '.sheet-hidden { color: #999; font-style: italic; }' +
    '</style></head><body>' +
    '<h2>Sheet Access Manager</h2>' +
    '<div class="subtitle">Editor (dan@commoditysampler.com) always has full access to all sheets.</div>' +
    '<div style="overflow-x:auto; max-height: 480px; overflow-y: auto;">' +
    '<table><thead><tr><th>Sheet Name</th>';

  // Column headers per user
  for (var u = 0; u < users.length; u++) {
    var label = users[u].role.charAt(0).toUpperCase() + users[u].role.slice(1);
    html += '<th class="user-col">' + label + '<br><span style="font-weight:normal;font-size:10px;">' + users[u].email.split('@')[0] + '</span></th>';
  }
  html += '</tr></thead><tbody>';

  // Rows per sheet
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i];
    html += '<tr><td>' + sheetName + '</td>';

    for (var j = 0; j < users.length; j++) {
      var email = users[j].email;
      var key = email + '|' + sheetName;
      var viewChecked = (config[key] && config[key].view) ? ' checked' : '';
      var editChecked = (config[key] && config[key].edit) ? ' checked' : '';

      html += '<td><div class="cb-group">' +
        '<label><input type="checkbox" data-email="' + email + '" data-sheet="' + _escapeHtmlAttr(sheetName) + '" data-perm="view"' + viewChecked + '> View</label>' +
        '<label><input type="checkbox" data-email="' + email + '" data-sheet="' + _escapeHtmlAttr(sheetName) + '" data-perm="edit"' + editChecked + '> Edit</label>' +
        '</div></td>';
    }
    html += '</tr>';
  }

  html += '</tbody></table></div>' +
    '<div class="btn-row">' +
    '<span id="saving" class="saving">Saving...</span>' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-save" onclick="saveConfig()">Save & Apply</button>' +
    '</div>' +
    '<script>' +
    'function saveConfig() {' +
    '  var config = {};' +
    '  var boxes = document.querySelectorAll("input[type=checkbox]");' +
    '  for (var i = 0; i < boxes.length; i++) {' +
    '    var email = boxes[i].getAttribute("data-email");' +
    '    var sheet = boxes[i].getAttribute("data-sheet");' +
    '    var perm = boxes[i].getAttribute("data-perm");' +
    '    var key = email + "|" + sheet;' +
    '    if (!config[key]) config[key] = { view: false, edit: false };' +
    '    config[key][perm] = boxes[i].checked;' +
    '  }' +
    '  document.getElementById("saving").style.display = "inline";' +
    '  google.script.run.withSuccessHandler(function() {' +
    '    document.getElementById("saving").textContent = "Saved!";' +
    '    setTimeout(function() { google.script.host.close(); }, 1000);' +
    '  }).withFailureHandler(function(e) {' +
    '    document.getElementById("saving").style.display = "none";' +
    '    alert("Error saving: " + e.message);' +
    '  }).saveAccessConfig(JSON.stringify(config));' +
    '}' +
    '</script></body></html>';

  var dialog = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(580)
    .setTitle('Sheet Access Manager');
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Sheet Access Manager');
}

// ============================================================
// ROLE MANAGER — editor-only modal to assign user roles
// ============================================================

/**
 * Opens modal to manage user→role assignments.
 * Editor can add/remove users and change their role.
 */
function showRoleManager() {
  var role = _getUserRole();
  if (role !== 'editor') {
    SpreadsheetApp.getUi().alert('Access Denied — Editor only.');
    return;
  }

  var roles = _loadUserRoles();
  var availableRoles = ['editor', 'billing', 'field', 'none'];

  var html = '<html><head><style>' +
    'body { font-family: Arial, sans-serif; margin: 0; padding: 16px; background: #f8f9fa; }' +
    'h2 { margin: 0 0 4px 0; font-size: 18px; }' +
    '.subtitle { color: #666; font-size: 12px; margin-bottom: 16px; }' +
    'table { border-collapse: collapse; width: 100%; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }' +
    'th { background: #2d5016; color: #fff; padding: 8px 10px; font-size: 12px; text-align: left; }' +
    'td { padding: 8px 10px; border-bottom: 1px solid #eee; font-size: 13px; }' +
    'tr:hover { background: #f0f7e8; }' +
    'input[type=email] { width: 100%; padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; box-sizing: border-box; }' +
    'select { padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; background: #fff; }' +
    '.btn-remove { background: #dc3545; color: #fff; border: none; border-radius: 4px; padding: 4px 10px; cursor: pointer; font-size: 12px; }' +
    '.btn-remove:hover { background: #c82333; }' +
    '.btn-add { background: #2d5016; color: #fff; border: none; border-radius: 6px; padding: 6px 16px; cursor: pointer; font-size: 13px; font-weight: 600; }' +
    '.btn-add:hover { background: #3d6b1e; }' +
    '.btn-row { margin-top: 16px; text-align: right; }' +
    '.btn { padding: 8px 20px; border: none; border-radius: 6px; font-size: 13px; cursor: pointer; font-weight: 600; }' +
    '.btn-save { background: #2d5016; color: #fff; }' +
    '.btn-save:hover { background: #3d6b1e; }' +
    '.btn-cancel { background: #e0e0e0; color: #333; margin-right: 8px; }' +
    '.btn-cancel:hover { background: #ccc; }' +
    '.saving { color: #2d5016; font-weight: 600; display: none; margin-right: 12px; }' +
    '.role-desc { font-size: 11px; color: #888; margin-top: 12px; }' +
    '.add-row { margin-top: 12px; display: flex; gap: 8px; align-items: center; }' +
    '</style></head><body>' +
    '<h2>Manage Menu Roles</h2>' +
    '<div class="subtitle">Assign which menus each user sees. Changes take effect on their next spreadsheet open.</div>' +
    '<table><thead><tr><th>Email</th><th>Role</th><th></th></tr></thead>' +
    '<tbody id="roleRows">';

  // Existing users
  var emails = Object.keys(roles);
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    var userRole = roles[email];
    html += _buildRoleRow(email, userRole, availableRoles);
  }

  html += '</tbody></table>' +
    '<div class="add-row">' +
    '<input type="email" id="newEmail" placeholder="user@example.com" style="flex:1;">' +
    '<select id="newRole">';
  for (var r = 0; r < availableRoles.length; r++) {
    html += '<option value="' + availableRoles[r] + '">' + availableRoles[r] + '</option>';
  }
  html += '</select>' +
    '<button class="btn-add" onclick="addUser()">+ Add</button>' +
    '</div>' +
    '<div class="role-desc">' +
    '<strong>Roles:</strong> ' +
    '<b>editor</b> = full access &nbsp;|&nbsp; ' +
    '<b>billing</b> = billing + my view &nbsp;|&nbsp; ' +
    '<b>field</b> = field report + my view &nbsp;|&nbsp; ' +
    '<b>none</b> = no access' +
    '</div>' +
    '<div class="btn-row">' +
    '<span id="saving" class="saving">Saving...</span>' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-save" onclick="saveRoles()">Save</button>' +
    '</div>';

  // Client-side JS
  html += '<script>' +
    'var availableRoles = ' + JSON.stringify(availableRoles) + ';' +

    'function addUser() {' +
    '  var email = document.getElementById("newEmail").value.trim().toLowerCase();' +
    '  if (!email || email.indexOf("@") === -1) { alert("Enter a valid email"); return; }' +
    '  var role = document.getElementById("newRole").value;' +
    '  var tbody = document.getElementById("roleRows");' +
    '  var existing = tbody.querySelectorAll("tr[data-email]");' +
    '  for (var i = 0; i < existing.length; i++) {' +
    '    if (existing[i].getAttribute("data-email") === email) { alert("User already exists"); return; }' +
    '  }' +
    '  var tr = document.createElement("tr");' +
    '  tr.setAttribute("data-email", email);' +
    '  var opts = "";' +
    '  for (var r = 0; r < availableRoles.length; r++) {' +
    '    opts += "<option value=\\"" + availableRoles[r] + "\\"" + (availableRoles[r] === role ? " selected" : "") + ">" + availableRoles[r] + "</option>";' +
    '  }' +
    '  tr.innerHTML = "<td>" + email + "</td><td><select class=\\"role-select\\">" + opts + "</select></td><td><button class=\\"btn-remove\\" onclick=\\"removeUser(this)\\">Remove</button></td>";' +
    '  tbody.appendChild(tr);' +
    '  document.getElementById("newEmail").value = "";' +
    '}' +

    'function removeUser(btn) {' +
    '  var row = btn.closest("tr");' +
    '  row.parentNode.removeChild(row);' +
    '}' +

    'function saveRoles() {' +
    '  var config = {};' +
    '  var rows = document.querySelectorAll("#roleRows tr[data-email]");' +
    '  for (var i = 0; i < rows.length; i++) {' +
    '    var email = rows[i].getAttribute("data-email");' +
    '    var select = rows[i].querySelector(".role-select");' +
    '    config[email] = select.value;' +
    '  }' +
    '  document.getElementById("saving").style.display = "inline";' +
    '  google.script.run.withSuccessHandler(function() {' +
    '    document.getElementById("saving").textContent = "Saved!";' +
    '    setTimeout(function() { google.script.host.close(); }, 1000);' +
    '  }).withFailureHandler(function(e) {' +
    '    document.getElementById("saving").style.display = "none";' +
    '    alert("Error saving: " + e.message);' +
    '  }).saveUserRoles(JSON.stringify(config));' +
    '}' +
    '</script></body></html>';

  var dialog = HtmlService.createHtmlOutput(html)
    .setWidth(560)
    .setHeight(480)
    .setTitle('Manage Menu Roles');
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Manage Menu Roles');
}

/** Builds a single table row for the role manager */
function _buildRoleRow(email, currentRole, availableRoles) {
  var opts = '';
  for (var r = 0; r < availableRoles.length; r++) {
    var sel = (availableRoles[r] === currentRole) ? ' selected' : '';
    opts += '<option value="' + availableRoles[r] + '"' + sel + '>' + availableRoles[r] + '</option>';
  }
  return '<tr data-email="' + _escapeHtmlAttr(email) + '">' +
    '<td>' + _escapeHtmlAttr(email) + '</td>' +
    '<td><select class="role-select">' + opts + '</select></td>' +
    '<td><button class="btn-remove" onclick="removeUser(this)">Remove</button></td>' +
    '</tr>';
}


function _escapeHtmlAttr(str) {
  return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}


// ============================================================
// ACCESS CONFIG — stored in ScriptProperties
// ============================================================

var ACCESS_CONFIG_KEY = 'SHEET_ACCESS_CONFIG';
var USER_ROLES_KEY = 'USER_ROLES_CONFIG';

function _loadAccessConfig() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (!props) return {};
    var raw = props.getProperty(ACCESS_CONFIG_KEY);
    if (!raw || typeof raw !== 'string') return {};

    var config = {};
    try {
      config = JSON.parse(raw);
    } catch (parseErr) {
      Logger.log('Error parsing access config JSON: ' + parseErr.message);
      return {};
    }

    return (config && typeof config === 'object') ? config : {};
  } catch (e) {
    Logger.log('Error loading access config: ' + e.message);
    return {};
  }
}

/**
 * Called from client-side JS in showAccessManager.
 * Saves config and immediately applies visibility.
 */
function saveAccessConfig(jsonStr) {
  try {
    var role = _getUserRole();
    if (role !== 'editor') throw new Error('Editor only');

    if (!jsonStr || typeof jsonStr !== 'string') {
      throw new Error('Invalid config JSON');
    }

    // Validate JSON before saving
    var config = JSON.parse(jsonStr);
    if (!config || typeof config !== 'object') {
      throw new Error('Config must be an object');
    }

    var props = PropertiesService.getScriptProperties();
    if (!props) throw new Error('Could not access script properties');

    props.setProperty(ACCESS_CONFIG_KEY, jsonStr);
    Logger.log('Access config saved');

    // Apply immediately for all roles (will take effect on their next onOpen)
    // But also apply now for any non-editor currently viewing
    // Nothing to do here — _applySheetVisibility runs on each user's onOpen
  } catch (e) {
    Logger.log('saveAccessConfig error: ' + e.message);
    throw e;
  }
}


// ============================================================
// USER ROLES CONFIG — stored in ScriptProperties, falls back to hardcoded
// ============================================================

/**
 * Loads user→role mapping. ScriptProperties override hardcoded USER_ROLES.
 * Returns object like { 'email@example.com': 'editor', ... }
 */
function _loadUserRoles() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (!props) return USER_ROLES || {};
    var raw = props.getProperty(USER_ROLES_KEY);
    if (raw && typeof raw === 'string') {
      var config = {};
      try {
        config = JSON.parse(raw);
      } catch (parseErr) {
        Logger.log('Error parsing user roles JSON: ' + parseErr.message);
        return USER_ROLES || {};
      }
      if (config && typeof config === 'object') return config;
    }
  } catch (e) {
    Logger.log('Error loading user roles config: ' + e.message);
  }
  // Fall back to hardcoded
  return USER_ROLES || {};
}

/**
 * Saves user→role mapping to ScriptProperties. Called from showRoleManager modal.
 * @param {string} jsonStr — JSON string of { email: role, ... }
 */
function saveUserRoles(jsonStr) {
  try {
    var role = _getUserRole();
    if (role !== 'editor') throw new Error('Editor only');

    if (!jsonStr || typeof jsonStr !== 'string') {
      throw new Error('Invalid roles JSON');
    }

    // Validate JSON before saving
    var config = JSON.parse(jsonStr);
    if (!config || typeof config !== 'object') {
      throw new Error('Config must be an object');
    }

    var props = PropertiesService.getScriptProperties();
    if (!props) throw new Error('Could not access script properties');

    props.setProperty(USER_ROLES_KEY, jsonStr);
    Logger.log('User roles config saved: ' + jsonStr);
  } catch (e) {
    Logger.log('saveUserRoles error: ' + e.message);
    throw e;
  }
}


// ============================================================
// SHEET VISIBILITY — enforced on every onOpen
// ============================================================

/**
 * Shows/hides sheets based on saved access config for the given role.
 * Editor always sees everything. Unknown users see nothing.
 */
function _applySheetVisibility(role) {
  try {
    if (role === 'editor') return; // editor sees all, no changes

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var sheets = ss.getSheets();
    if (!sheets || sheets.length === 0) return;

    var config = _loadAccessConfig();
    var email = '';

    // Get the email for this role
    try {
      email = Session.getActiveUser().getEmail().toLowerCase().trim();
    } catch (e) {
      Logger.log('Cannot get email for visibility: ' + e.message);
      return;
    }

    if (!email) return;

    // If no config saved yet, do nothing (editor hasn't configured yet)
    if (Object.keys(config).length === 0) return;

    // We need at least one sheet visible at all times
    var visibleCount = 0;
    var sheetsToHide = [];
    var sheetsToShow = [];

    for (var i = 0; i < sheets.length; i++) {
      if (!sheets[i]) continue;
      var name = sheets[i].getName();
      var key = email + '|' + name;
      var hasView = config[key] && config[key].view;
      var hasEdit = config[key] && config[key].edit;

      if (role === 'none') {
        // Unknown users — hide everything
        sheetsToHide.push(sheets[i]);
      } else if (hasView || hasEdit) {
        sheetsToShow.push(sheets[i]);
        visibleCount++;
      } else {
        sheetsToHide.push(sheets[i]);
      }
    }

    // Safety: must keep at least 1 sheet visible
    if (visibleCount === 0 && sheetsToShow.length === 0) {
      // Show first sheet as fallback
      if (sheets.length > 0 && sheets[0]) {
        sheetsToShow.push(sheets[0]);
        // Remove from hide list
        sheetsToHide = sheetsToHide.filter(function(s) { return s && s.getName() !== sheets[0].getName(); });
      }
    }

    // Show first, then hide (can't hide all sheets)
    for (var s = 0; s < sheetsToShow.length; s++) {
      try {
        if (sheetsToShow[s]) sheetsToShow[s].showSheet();
      } catch (e) {
        Logger.log('Could not show sheet: ' + e.message);
      }
    }
    for (var h = 0; h < sheetsToHide.length; h++) {
      try {
        if (sheetsToHide[h]) sheetsToHide[h].hideSheet();
      } catch (e) {
        Logger.log('Could not hide sheet: ' + e.message);
      }
    }
  } catch (e) {
    Logger.log('_applySheetVisibility error: ' + e.message);
  }
}
