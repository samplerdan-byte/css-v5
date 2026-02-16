// ============================================================
// Export.gs — CSV export, SQL export, backup functions
// ============================================================

function exportTodaysOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  var col = _getColumnMap(sheet);
  const lastRow = sheet.getLastRow();
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

  const folder = DriveApp.getRootFolder();
  const file = folder.createFile(fileName, csvContent, MimeType.CSV);

  ui.alert('Export Complete!\n\n' +
    'Exported ' + todaysData.length + ' orders to:\n' +
    file.getName() + '\n\n' +
    'File URL:\n' + file.getUrl());
}

function exportAllOrders() {
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

  const folder = DriveApp.getRootFolder();
  const file = folder.createFile(fileName, csvContent, MimeType.CSV);

  ui.alert('Export Complete!\n\n' +
    'Exported ' + data.length + ' orders to:\n' +
    file.getName() + '\n\n' +
    'File URL:\n' + file.getUrl());
}

function exportForSQL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  var col = _getColumnMap(sheet);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

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
      formatDateForSQL(row[col['Timestamp']]),
      row[col['CS Order #']],
      row[col['CS Sample #']],
      row[col['Sender']],
      row[col['Receiver']],
      row[col['Warehouse']],
      row[col['Description']],
      row[col['Sample Order #']],
      row[col['Cargo #']],
      row[col['Mark #']],
      row[col['Container #']],
      row[col['Reference']],
      row[col['Bag Count']],
      row[col['Weight']],
      row[col['Sample Weight']],
      row[col['Status']],
      row[col['Tracking Number']],
      formatDateForSQL(row[col['Shipped Date']])
    ];

    const escapedValues = values.map(v => {
      if (v === null || v === undefined || v === '') return 'NULL';
      return "'" + String(v).replace(/'/g, "''") + "'";
    });

    sqlStatements += 'INSERT INTO css_orders (timestamp, cs_order_num, cs_sample_num, sender, receiver, warehouse, description, sample_order_num, cargo, mark, container, reference, bag_count, weight, sample_weight, status, tracking_number, shipped_date) VALUES (' + escapedValues.join(', ') + ');\n';
  });

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fileName = 'CSS_Orders_SQL_' + dateStr + '.sql';

  const folder = DriveApp.getRootFolder();
  const file = folder.createFile(fileName, sqlStatements, MimeType.PLAIN_TEXT);

  ui.alert('SQL Export Complete!\n\n' +
    'Exported ' + data.length + ' records to:\n' +
    file.getName() + '\n\n' +
    'File URL:\n' + file.getUrl());
}

// ============================================================
// BACKUP FUNCTIONS
// ============================================================

function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  const backupName = ss.getName() + ' - Backup ' + dateStr;

  const backup = ss.copy(backupName);

  ui.alert('Backup Created!\n\n' +
    'Name: ' + backupName + '\n\n' +
    'URL: ' + backup.getUrl());
}

function scheduledBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const backupName = ss.getName() + ' - Auto Backup ' + dateStr;

  const backup = ss.copy(backupName);
  Logger.log('Scheduled backup created: ' + backupName);

  // Clean up old backups (keep last 7 days)
  const files = DriveApp.getFilesByName(ss.getName() + ' - Auto Backup');
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 7);

  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoffDate) {
      file.setTrashed(true);
      Logger.log('Trashed old backup: ' + file.getName());
    }
  }
}
