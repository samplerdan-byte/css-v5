// ============================================================
// Export.gs — CSV export, SQL export, backup functions
// ============================================================

function exportTodaysOrders() {
  Logger.log('exportTodaysOrders started');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data to export!'); return; }

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

  let csvContent = headers.map(escapeCSVField).join(',') + '\n';
  todaysData.forEach(row => {
    csvContent += row.map(escapeCSVField).join(',') + '\n';
  });

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
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) { ui.alert('No data to export!'); return; }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let csvContent = headers.map(escapeCSVField).join(',') + '\n';
  data.forEach(row => {
    csvContent += row.map(escapeCSVField).join(',') + '\n';
  });

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
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data to export!'); return; }

  var col = _getColumnMap(sheet);
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Safe column lookup helper — returns undefined index as -1 to avoid undefined access
  function safeCol(name) { return col[name] !== undefined ? col[name] : -1; }
  function safeVal(row, name) { var idx = safeCol(name); return idx >= 0 ? row[idx] : ''; }

  let sqlStatements = '-- CSS Orders Export\n-- Generated: ' + new Date().toISOString() + '\n\n';
  sqlStatements += 'CREATE TABLE IF NOT EXISTS css_orders (\n';
  sqlStatements += '  id INT AUTO_INCREMENT PRIMARY KEY,\n';
  sqlStatements += '  timestamp DATETIME,\n';
  sqlStatements += '  cs_order_num VARCHAR(20),\n';
  sqlStatements += '  cs_sample_num VARCHAR(25),\n';
  sqlStatements += '  sender VARCHAR(100),\n';
  sqlStatements += '  receiver VARCHAR(200),\n';
  sqlStatements += '  warehouse VARCHAR(100),\n';
  sqlStatements += '  description VARCHAR(200),\n';
  sqlStatements += '  sample_order_num VARCHAR(50),\n';
  sqlStatements += '  cargo VARCHAR(20),\n';
  sqlStatements += '  mark VARCHAR(30),\n';
  sqlStatements += '  container VARCHAR(20),\n';
  sqlStatements += '  reference VARCHAR(50),\n';
  sqlStatements += '  bag_count VARCHAR(20),\n';
  sqlStatements += '  weight VARCHAR(30),\n';
  sqlStatements += '  sample_weight VARCHAR(20),\n';
  sqlStatements += '  status VARCHAR(20),\n';
  sqlStatements += '  tracking_number VARCHAR(50),\n';
  sqlStatements += '  shipped_date DATETIME\n';
  sqlStatements += ');\n\n';

  data.forEach(row => {
    const values = [
      typeof formatDateForSQL === 'function' ? formatDateForSQL(safeVal(row, 'Timestamp')) : safeVal(row, 'Timestamp'),
      safeVal(row, 'CS Order #'),
      safeVal(row, 'CS Sample #'),
      safeVal(row, 'Sender'),
      safeVal(row, 'Receiver'),
      safeVal(row, 'Warehouse'),
      safeVal(row, 'Description'),
      safeVal(row, 'Sample Order #'),
      safeVal(row, 'Cargo #'),
      safeVal(row, 'Mark #'),
      safeVal(row, 'Container #'),
      safeVal(row, 'Reference'),
      safeVal(row, 'Bag Count'),
      safeVal(row, 'Weight'),
      safeVal(row, 'Sample Weight'),
      safeVal(row, 'Status'),
      safeVal(row, 'Tracking Number'),
      typeof formatDateForSQL === 'function' ? formatDateForSQL(safeVal(row, 'Shipped Date')) : safeVal(row, 'Shipped Date')
    ];

    const escapedValues = values.map(v => {
      if (v === null || v === undefined || v === '') return 'NULL';
      return "'" + String(v).replace(/'/g, "''") + "'";
    });

    sqlStatements += 'INSERT INTO css_orders (timestamp, cs_order_num, cs_sample_num, sender, receiver, warehouse, description, sample_order_num, cargo, mark, container, reference, bag_count, weight, sample_weight, status, tracking_number, shipped_date) VALUES (' + escapedValues.join(', ') + ');\n';
  });

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
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const backupName = ss.getName() + ' - Auto Backup ' + dateStr;

    const backup = ss.copy(backupName);
    Logger.log('Scheduled backup created: ' + backupName);

    // Clean up old backups (keep last 7 days)
    // Use search query to find files whose name starts with the backup prefix
    var backupPrefix = ss.getName() + ' - Auto Backup';
    var files = DriveApp.searchFiles('title contains "' + backupPrefix.replace(/"/g, '\\"') + '"');
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 7);

    while (files.hasNext()) {
      var file = files.next();
      if (file.getName().indexOf(backupPrefix) === 0 && file.getDateCreated() < cutoffDate) {
        file.setTrashed(true);
        Logger.log('Trashed old backup: ' + file.getName());
      }
    }
  } catch (e) {
    Logger.log('scheduledBackup error: ' + e.message);
  }
}
