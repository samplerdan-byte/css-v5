// ============================================================
// Export.gs — CSV export, SQL export, backup functions
// ============================================================

function exportTodaysOrders() {
  Logger.log('exportTodaysOrders started');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName || '');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data to export!'); return; }

  // Validate CONFIG.mainSheetName is a string
  if (!CONFIG.mainSheetName || typeof CONFIG.mainSheetName !== 'string') {
    ui.alert('Invalid sheet configuration');
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  var col = _getColumnMap(sheet);
  if (col['Timestamp'] === undefined) { ui.alert('Timestamp column not found!'); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const todaysData = data.filter(row => {
    const rowDate = new Date(row[col['Timestamp']]);
    rowDate.setHours(0, 0, 0, 0);
    return rowDate.getTime() === today.getTime();
  });

  if (todaysData.length === 0) {
    ui.alert('No orders found for today!');
    return;
  }

  // V5: perf — Array.join instead of string concat in loop (O(n²) alloc → O(n))
  var csvLines = [headers.map(escapeCSVField).join(',')];
  todaysData.forEach(function(row) {
    var safeRow = row.map(function(v) { return String(v || '').trim(); });
    csvLines.push(safeRow.map(escapeCSVField).join(','));
  });
  var csvContent = csvLines.join('\n') + '\n';

  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fileName = 'CSS_Orders_' + dateStr + '.csv';

  try {
    const folder = DriveApp.getRootFolder();
    const file = folder.createFile(fileName, csvContent, MimeType.CSV);
    Logger.log('exportTodaysOrders: exported ' + todaysData.length + ' orders to ' + fileName);

    ui.alert('Export Complete!\n\n' +
      'Exported ' + todaysData.length + ' orders to:\n' +
      file.getName() + '\n\n' +
      'File URL:\n' + file.getUrl());
  } catch (e) {
    Logger.log('exportTodaysOrders error creating file: ' + e.message);
    ui.alert('Export failed: ' + e.message);
  }
}

function exportAllOrders() {
  Logger.log('exportAllOrders started');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName || '');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) { ui.alert('No data to export!'); return; }
  if (lastCol < 1) { ui.alert('Sheet has no columns!'); return; }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // V5: perf — Array.join instead of string concat in loop (O(n²) alloc → O(n))
  var csvLines = [headers.map(escapeCSVField).join(',')];
  data.forEach(function(row) {
    var safeRow = row.map(function(v) { return String(v || '').trim(); });
    csvLines.push(safeRow.map(escapeCSVField).join(','));
  });
  var csvContent = csvLines.join('\n') + '\n';

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  const fileName = 'CSS_All_Orders_' + dateStr + '.csv';

  try {
    const folder = DriveApp.getRootFolder();
    const file = folder.createFile(fileName, csvContent, MimeType.CSV);
    Logger.log('exportAllOrders: exported ' + data.length + ' orders to ' + fileName);

    ui.alert('Export Complete!\n\n' +
      'Exported ' + data.length + ' orders to:\n' +
      file.getName() + '\n\n' +
      'File URL:\n' + file.getUrl());
  } catch (e) {
    Logger.log('exportAllOrders error creating file: ' + e.message);
    ui.alert('Export failed: ' + e.message);
  }
}

function exportForSQL() {
  Logger.log('exportForSQL started');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName || '');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data to export!'); return; }

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) { ui.alert('Sheet has no columns!'); return; }

  var col = _getColumnMap(sheet);
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Safe column lookup helper — returns undefined index as -1 to avoid undefined access
  function safeCol(name) { return col[name] !== undefined ? col[name] : -1; }
  function safeVal(row, name) { var idx = safeCol(name); return idx >= 0 ? row[idx] : ''; }

  // V5: perf — Array.join instead of string concat in loop (O(n²) alloc → O(n))
  var sqlParts = [
    '-- CSS Orders Export\n-- Generated: ' + new Date().toISOString() + '\n',
    'CREATE TABLE IF NOT EXISTS css_orders (\n' +
    '  id INT AUTO_INCREMENT PRIMARY KEY,\n' +
    '  timestamp DATETIME,\n' +
    '  cs_order_num VARCHAR(20),\n' +
    '  cs_sample_num VARCHAR(25),\n' +
    '  sender VARCHAR(100),\n' +
    '  receiver VARCHAR(200),\n' +
    '  warehouse VARCHAR(100),\n' +
    '  description VARCHAR(200),\n' +
    '  sample_order_num VARCHAR(50),\n' +
    '  cargo VARCHAR(20),\n' +
    '  mark VARCHAR(30),\n' +
    '  container VARCHAR(20),\n' +
    '  reference VARCHAR(50),\n' +
    '  bag_count VARCHAR(20),\n' +
    '  weight VARCHAR(30),\n' +
    '  sample_weight VARCHAR(20),\n' +
    '  status VARCHAR(20),\n' +
    '  tracking_number VARCHAR(50),\n' +
    '  shipped_date DATETIME\n' +
    ');\n'
  ];

  data.forEach(function(row) {
    var values = [
      typeof formatDateForSQL === 'function' ? formatDateForSQL(safeVal(row, 'Timestamp')) : String(safeVal(row, 'Timestamp') || '').trim(),
      String(safeVal(row, 'CS Order #') || '').trim(),
      String(safeVal(row, 'CS Sample #') || '').trim(),
      String(safeVal(row, 'Sender') || '').trim(),
      String(safeVal(row, 'Receiver') || '').trim(),
      String(safeVal(row, 'Warehouse') || '').trim(),
      String(safeVal(row, 'Description') || '').trim(),
      String(safeVal(row, 'Sample Order #') || '').trim(),
      String(safeVal(row, 'Cargo #') || '').trim(),
      String(safeVal(row, 'Mark #') || '').trim(),
      String(safeVal(row, 'Container #') || '').trim(),
      String(safeVal(row, 'Reference') || '').trim(),
      String(safeVal(row, 'Bag Count') || '').trim(),
      String(safeVal(row, 'Weight') || '').trim(),
      String(safeVal(row, 'Sample Weight') || '').trim(),
      String(safeVal(row, 'Status') || '').trim(),
      String(safeVal(row, 'Tracking Number') || '').trim(),
      typeof formatDateForSQL === 'function' ? formatDateForSQL(safeVal(row, 'Shipped Date')) : String(safeVal(row, 'Shipped Date') || '').trim()
    ];

    var escapedValues = values.map(function(v) {
      if (v === null || v === undefined || v === '') return 'NULL';
      return "'" + String(v || '').replace(/'/g, "''") + "'";
    });

    sqlParts.push('INSERT INTO css_orders (timestamp, cs_order_num, cs_sample_num, sender, receiver, warehouse, description, sample_order_num, cargo, mark, container, reference, bag_count, weight, sample_weight, status, tracking_number, shipped_date) VALUES (' + escapedValues.join(', ') + ');');
  });

  var sqlStatements = sqlParts.join('\n') + '\n';
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fileName = 'CSS_Orders_SQL_' + dateStr + '.sql';

  try {
    const folder = DriveApp.getRootFolder();
    const file = folder.createFile(fileName, sqlStatements, MimeType.PLAIN_TEXT);
    Logger.log('exportForSQL: exported ' + data.length + ' records to ' + fileName);

    ui.alert('SQL Export Complete!\n\n' +
      'Exported ' + data.length + ' records to:\n' +
      file.getName() + '\n\n' +
      'File URL:\n' + file.getUrl());
  } catch (e) {
    Logger.log('exportForSQL error creating file: ' + e.message);
    ui.alert('SQL Export failed: ' + e.message);
  }
}

// ============================================================
// BACKUP FUNCTIONS
// ============================================================

function createBackup() {
  Logger.log('createBackup started');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  const backupName = ss.getName() + ' - Backup ' + dateStr;

  try {
    const backup = ss.copy(backupName);
    Logger.log('createBackup: created ' + backupName);

    ui.alert('Backup Created!\n\n' +
      'Name: ' + backupName + '\n\n' +
      'URL: ' + backup.getUrl());
  } catch (e) {
    Logger.log('createBackup error: ' + e.message);
    ui.alert('Backup failed: ' + e.message);
  }
}

function scheduledBackup() {
  Logger.log('scheduledBackup started');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssName = String(ss.getName() || '').trim();
    if (!ssName) {
      Logger.log('scheduledBackup: spreadsheet name is empty');
      return;
    }

    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const backupName = ssName + ' - Auto Backup ' + dateStr;

    const backup = ss.copy(backupName);
    Logger.log('Scheduled backup created: ' + backupName);

    // Clean up old backups (keep last 7 days)
    // Use search query to find files whose name starts with the backup prefix
    var backupPrefix = ssName + ' - Auto Backup';
    var files = DriveApp.searchFiles('title contains "' + backupPrefix.replace(/"/g, '\\"') + '"');
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 7);

    while (files.hasNext()) {
      var file = files.next();
      var fileName = String(file.getName() || '').trim();
      if (fileName.indexOf(backupPrefix) === 0 && file.getDateCreated() < cutoffDate) {
        file.setTrashed(true);
        Logger.log('Trashed old backup: ' + fileName);
      }
    }
  } catch (e) {
    Logger.log('scheduledBackup error: ' + e.message);
  }
}
