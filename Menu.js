// ============================================================
// Menu.gs — onOpen menu, triggers
// ============================================================

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();

    ui.createMenu('☕ CSS System')
      .addItem('📧 Process New Emails', 'processPDFsFromGmail')
      .addItem('📋 View Order Details', 'viewOrderDetails')
      .addItem('✍️ Manual Order Entry', 'showEnhancedManualEntry')
      .addSeparator()
      .addSubMenu(ui.createMenu('📱 Scanner')
        .addItem('📷 Receiving Scanner (Scan In)', 'openBarcodeScanner')
        .addItem('📦 Shipping Scanner (Scan Out)', 'showScanOut')
        .addItem('📱 Web Scanner', 'openWebScanner')
        .addItem('📸 Sample Photos', 'showPhotoUpload'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📊 Reports')
        .addItem('📊 Operations Dashboard', 'refreshDashboard')
        .addSeparator()
        .addItem('📋 Open Orders Report', 'generateOpenOrdersReport')
        .addItem('🏭 Field Report', 'generateFieldReport')
        .addItem('📥 Import Offline Edits', 'showImportOfflineEdits')
        .addItem('📄 Printable Order Forms', 'generatePrintableOrderForm')
        .addItem('📈 End of Day Report', 'generateEndOfDayReport'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🖨️ Print')
        .addItem('☑️ Print Checked Samples', 'printCheckedSamples')
        .addItem('🔄 Reprint Selected Label', 'reprintLabel')
        .addItem('📋 Print Today\'s Cover Sheets', 'printTodaysCoverSheets')
        .addItem('🏷️ Generate Labels (Selected)', 'generateLabels'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📤 Export')
        .addItem('📅 Export Today\'s Orders', 'exportTodaysOrders')
        .addItem('📊 Export All Orders', 'exportAllOrders')
        .addItem('🗄️ Export for SQL', 'exportForSQL'))
      .addSeparator()
      .addSubMenu(ui.createMenu('👤 My View')
        .addItem('➕ Create My View', 'createMyView')
        .addItem('🔄 Refresh My View', 'refreshMyView'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📦 Orders')
        .addItem('🔄 Update Live Orders', 'updateLiveOrdersView')
        .addItem('📦 Move Shipped → Completed', 'processPendingMovesWithAlert')
        .addItem('🗄️ Archive Old Orders (90+ days)', 'archiveOldOrders')
        .addItem('📧 Send Daily Customer Reports', 'sendDailyCustomerReports')
        .addItem('🔍 Check Sheet Data', 'checkSheetData')
        .addItem('🗑️ Delete Selected Rows (Synced)', 'deleteSelectedRows'))
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ Setup')
        .addItem('🛠️ Initial Setup', 'setupSheet')
        .addItem('📊 Setup Tracking System', 'setupTrackingSystem')
        .addItem('📧 Setup Customer Emails', 'setupCustomerEmailSheet')
        .addItem('🏭 Setup Warehouse Emails', 'setupWarehouseEmails')
        .addItem('🏭 Setup Field Report Columns', 'setupFieldColumns')
        .addItem('✏️ Setup Field Report Editors', 'setupFieldReportEditors')
        .addItem('📇 Setup Contacts (Ship From/To)', 'setupContacts')
        .addItem('☑️ Add Print Checkboxes', 'addPrintCheckboxColumn')
        .addItem('🔧 Fix All Print Checkboxes', 'fixAllPrintCheckboxes')
        .addItem('🔄 Sync Print Column', 'syncPrintColumnToAllSheets')
        .addItem('🎨 Setup Conditional Formatting', 'setupConditionalFormatting')
        .addItem('🔒 Lock Headers', 'lockAllHeadersWithAlert')
        .addItem('🔓 Unlock Headers', 'unlockHeaders')
        .addItem('🔧 Fix Completed Orders Headers', 'fixCompletedOrdersHeaders')
        .addItem('🔍 Diagnose Shipping Issue', 'diagnoseShippingIssue'))
      .addSeparator()
      .addSubMenu(ui.createMenu('💾 Backup')
        .addItem('💾 Create Backup Now', 'createBackup')
        .addItem('🗑️ Clear Cover Sheet Flags', 'clearCoverSheetPrintedFlags'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🌐 Client Portal')
        .addItem('⚙️ Setup Portal Access', 'setupPortalAccess')
        .addItem('🔑 Generate Token', 'generatePortalToken')
        .addItem('🔗 List All Portal Links', 'listPortalLinks')
        .addItem('🚫 Revoke Access (selected row)', 'revokePortalToken'))
      .addSeparator()
      .addSubMenu(ui.createMenu('☀️ Briefing')
        .addItem('📧 Send Briefing Now', 'sendMorningBriefing')
        .addItem('⚙️ Setup Daily Trigger (7 AM)', 'setupMorningBriefing')
        .addItem('🚫 Remove Daily Trigger', 'removeMorningBriefing'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🧠 Learning')
        .addItem('⚙️ Setup Learning System', 'setupLearningSystem')
        .addItem('🧠 Learn from Corrections', 'learnFromCorrections')
        .addItem('📊 Learning Dashboard', 'showLearningDashboard')
        .addItem('🗑️ Reset Sender Rules', 'resetSenderRules'))
      .addToUi();

    ui.createMenu('💰 Billing')
      .addItem('Setup Billing Sheets', 'setupBilling')
      .addSeparator()
      .addItem('Generate Invoices...', 'generateInvoice')
      .addItem('Preview Invoice...', 'previewInvoice')
      .addSeparator()
      .addItem('Edit Line Item...', 'editLineItemService')
      .addItem('Mark Invoice Sent', 'markInvoiceSent')
      .addItem('Mark Invoice Paid', 'markInvoicePaid')
      .addItem('Void Invoice', 'markInvoiceVoid')
      .addSeparator()
      .addItem('Export to QB Online (CSV)', 'exportInvoicesQBO')
      .addItem('Export to QB Desktop (IIF)', 'exportInvoicesIIF')
      .addSeparator()
      .addItem('Billing Dashboard', 'showBillingDashboard')
      .addToUi();

    ui.createMenu('📈 Market Intel')
      .addItem('Refresh Dashboard', 'buildMarketIntelSheet')
      .addToUi();

    Logger.log('CSS Menu created');

  } catch (e) {
    // If menu fails, at least log WHY so it shows in Executions
    Logger.log('onOpen MENU ERROR: ' + e.message + ' | Stack: ' + e.stack);
    // Try to show a minimal menu so you're not locked out
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
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processPDFsFromGmail' ||
        trigger.getHandlerFunction() === 'scheduledBackup' ||
        trigger.getHandlerFunction() === 'processPendingMoves') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('processPDFsFromGmail')
    .timeBased().everyMinutes(15).create();

  ScriptApp.newTrigger('scheduledBackup')
    .timeBased().atHour(23).everyDays(1).create();

  ScriptApp.newTrigger('processPendingMoves')
    .timeBased().everyHours(1).create();

  Logger.log('Time triggers created');
  SpreadsheetApp.getUi().alert('Triggers Created!\n\n• Email processing: every 15 min\n• Backup: daily at 11 PM\n• Move shipped orders: hourly');
}

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) { ScriptApp.deleteTrigger(trigger); });
  Logger.log('All triggers deleted');
  SpreadsheetApp.getUi().alert('All triggers deleted.');
}

function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var info = 'Current Triggers:\n\n';

  if (triggers.length === 0) {
    info += 'No triggers set up.';
  } else {
    triggers.forEach(function(trigger) {
      info += '• ' + trigger.getHandlerFunction() + '\n';
    });
  }

  SpreadsheetApp.getUi().alert(info);
}