// ============================================================
// Views.gs — Personal view sheets
// ============================================================

function createMyView() {
  Logger.log('createMyView started');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!mainSheet) {
    ui.alert(CONFIG.mainSheetName + ' sheet not found!');
    return;
  }

  var email = Session.getActiveUser().getEmail();
  if (!email) {
    ui.alert('Could not determine your email. Please make sure you are signed in.');
    return;
  }
  var userName = email.split('@')[0];
  userName = userName.charAt(0).toUpperCase() + userName.slice(1);
  var viewName = 'View - ' + userName;

  var existing = ss.getSheetByName(viewName);
  if (existing) {
    var response = ui.alert('View Exists',
      '"' + viewName + '" already exists.\n\nReset it? (This will clear your column customizations)',
      ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  try {
    var viewSheet = existing || ss.insertSheet(viewName);
    viewSheet.clear();

    var lastCol = mainSheet.getLastColumn();
    var headers = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    viewSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    viewSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('white');
    viewSheet.setFrozenRows(1);

    var count = _refreshViewSheet(mainSheet, viewSheet);

    ss.setActiveSheet(viewSheet);
    Logger.log('createMyView: created "' + viewName + '" with ' + count + ' rows');
    ui.alert('✅ Personal View Created!\n\n"' + viewName + '" — ' + count + ' rows\n\n' +
      'You can now:\n' +
      '• Drag columns to any order you prefer\n' +
      '• Right-click a column → Hide columns you don\'t need\n' +
      '• Delete columns you never want to see\n\n' +
      'Run "🔄 Refresh My View" anytime to sync latest data.');
  } catch (e) {
    Logger.log('createMyView error: ' + e.message);
    ui.alert('Error creating view: ' + e.message);
  }
}

function refreshMyView() {
  Logger.log('refreshMyView started');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!mainSheet) {
    ui.alert(CONFIG.mainSheetName + ' sheet not found!');
    return;
  }

  var email = Session.getActiveUser().getEmail();
  if (!email) {
    ui.alert('Could not determine your email. Please make sure you are signed in.');
    return;
  }
  var userName = email.split('@')[0];
  userName = userName.charAt(0).toUpperCase() + userName.slice(1);
  var viewName = 'View - ' + userName;

  var viewSheet = ss.getSheetByName(viewName);
  if (!viewSheet) {
    ui.alert('No personal view found.\nRun "Create My View" first.');
    return;
  }

  try {
    var count = _refreshViewSheet(mainSheet, viewSheet);
    Logger.log('refreshMyView: refreshed "' + viewName + '" with ' + count + ' rows');
    ui.alert('✅ View refreshed: ' + count + ' rows synced.');
  } catch (e) {
    Logger.log('refreshMyView error: ' + e.message);
    ui.alert('Error refreshing view: ' + e.message);
  }
}

function _refreshViewSheet(mainSheet, viewSheet) {
  if (!mainSheet || !viewSheet) {
    Logger.log('_refreshViewSheet: missing mainSheet or viewSheet');
    return 0;
  }

  var viewLastCol = viewSheet.getLastColumn();
  if (viewLastCol < 1) return 0;

  var viewHeaders = viewSheet.getRange(1, 1, 1, viewLastCol).getValues()[0];

  var mainLastRow = mainSheet.getLastRow();
  var mainLastCol = mainSheet.getLastColumn();

  if (viewSheet.getLastRow() > 1) {
    viewSheet.getRange(2, 1, viewSheet.getLastRow() - 1, viewLastCol).clear();
  }

  if (mainLastRow < 2) return 0;

  var mainCol = _getColumnMap(mainSheet);
  var mainData = mainSheet.getRange(2, 1, mainLastRow - 1, mainLastCol).getValues();

  var colMapping = [];
  for (var c = 0; c < viewHeaders.length; c++) {
    colMapping.push(mainCol[viewHeaders[c]] !== undefined ? mainCol[viewHeaders[c]] : -1);
  }

  var viewData = [];
  for (var r = 0; r < mainData.length; r++) {
    var row = [];
    for (var c = 0; c < colMapping.length; c++) {
      row.push(colMapping[c] >= 0 ? mainData[r][colMapping[c]] : '');
    }
    viewData.push(row);
  }

  if (viewData.length > 0) {
    viewSheet.getRange(2, 1, viewData.length, viewHeaders.length).setValues(viewData);
  }

  Logger.log('_refreshViewSheet: ' + viewData.length + ' rows, ' + viewHeaders.length + ' columns');
  return viewData.length;
}
