// ============================================================
// Setup.gs — Sheet setup, header protection, conditional formatting
// ============================================================

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!mainSheet) mainSheet = ss.insertSheet(CONFIG.mainSheetName);

  const headers = [
    'Timestamp', 'CS Order #', 'CS Sample #', 'Sender', 'Receiver', 'Warehouse',
    'Description', 'Sample Order #', 'Cargo #', 'Print', 'Mark #', 'Container #', 'Reference',
    'Bag Count', 'Weight', 'Sample Weight', 'P #', 'S #', 'Shipping Process',
    'Comments', 'Source Email', 'QR Data',
    'Status', 'Scanned Date', 'Scanned By', 'Shipped Date', 'Tracking Number', 'Report Date', 'Report Sent To',
    'Email Link', 'Attachments', 'Photos', 'Container Status', 'Container ETA',
    'Sample Type', 'Shipping Line', 'Shipping Notes'
  ];
  mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  mainSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  mainSheet.setFrozenRows(1);

  // Find Status column and set defaults
  const statusColIdx = headers.indexOf('Status');
  const lastRow = mainSheet.getLastRow();
  if (lastRow > 1 && statusColIdx !== -1) {
    const statusCol = statusColIdx + 1;
    const statusRange = mainSheet.getRange(2, statusCol, lastRow - 1, 1);
    const statusValues = statusRange.getValues();
    const updatedValues = statusValues.map(row =>
      row[0] === '' ? [CONFIG.statusValues.RECEIVED] : row
    );
    statusRange.setValues(updatedValues);
  }

  // Find Print column and add checkboxes
  const printColIdx = headers.indexOf('Print');
  if (lastRow > 1 && printColIdx !== -1) {
    mainSheet.getRange(2, printColIdx + 1, lastRow - 1, 1).insertCheckboxes();
  }
  if (printColIdx !== -1) {
    mainSheet.setColumnWidth(printColIdx + 1, 50);
  }

  // Setup Scan Log sheet
  let scanSheet = ss.getSheetByName(CONFIG.scanSheetName);
  if (!scanSheet) scanSheet = ss.insertSheet(CONFIG.scanSheetName);

  const scanHeaders = ['Scan Timestamp', 'CS Order #', 'CS Sample #', 'Sample Order #', 'Cargo #', 'Mark #', 'Container #', 'Reference', 'Description', 'Bag Count', 'Weight', 'Sample Weight', 'P #', 'S #', 'Warehouse', 'Shipping Process', 'Comments', 'Scanned By', 'Notes'];
  scanSheet.getRange(1, 1, 1, scanHeaders.length).setValues([scanHeaders]);
  scanSheet.getRange(1, 1, 1, scanHeaders.length).setFontWeight('bold');
  scanSheet.setFrozenRows(1);

  try { lockAllHeaders(); } catch(e) { Logger.log('Header lock skipped: ' + e); }

  SpreadsheetApp.getUi().alert('Setup Complete!\n\n✓ Tracking columns added\n✓ Print column with checkboxes added\n✓ Attachment columns added\n✓ Container tracking columns added\n✓ Header row locked\n✓ Existing rows set to "Received" status');
}

function setupTrackingSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'Setup Barcode Tracking System',
    'This will:\n' +
    '1. Add status tracking columns to main sheet\n' +
    '2. Create "Live Orders" view sheet\n' +
    '3. Create "Completed Orders" archive sheet\n' +
    '4. Create "Daily Report Log" sheet\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  try {
    const mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
    if (!mainSheet) {
      ui.alert('Main sheet not found! Please run "Setup Sheets" first.');
      return;
    }

    const lastCol = mainSheet.getLastColumn();
    const headers = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    if (!headers.includes('Status')) {
      const newHeaders = ['Status', 'Scanned Date', 'Scanned By', 'Shipped Date', 'Tracking Number', 'Report Date', 'Report Sent To'];
      const startCol = lastCol + 1;

      mainSheet.getRange(1, startCol, 1, newHeaders.length).setValues([newHeaders]);
      mainSheet.getRange(1, startCol, 1, newHeaders.length).setFontWeight('bold').setBackground('#d9ead3');

      const lastRow = mainSheet.getLastRow();
      if (lastRow > 1) {
        const statusData = [];
        for (let i = 2; i <= lastRow; i++) {
          statusData.push(['Received', '', '', '', '', '', '']);
        }
        mainSheet.getRange(2, startCol, statusData.length, 7).setValues(statusData);
      }
      Logger.log('✓ Added status tracking columns');
    }

    // Create Live Orders sheet
    let liveSheet = ss.getSheetByName('Live Orders');
    if (!liveSheet) {
      liveSheet = ss.insertSheet('Live Orders');
      const allHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues();
      liveSheet.getRange(1, 1, 1, allHeaders[0].length).setValues(allHeaders);
      liveSheet.getRange(1, 1, 1, allHeaders[0].length).setFontWeight('bold').setBackground('#d9ead3');
      liveSheet.getRange(1, 1, 1, allHeaders[0].length).createFilter();
      liveSheet.setFrozenRows(1);
      Logger.log('✓ Created Live Orders sheet');
    }

    // Create Completed Orders sheet
    let completedSheet = ss.getSheetByName('Completed Orders');
    if (!completedSheet) {
      completedSheet = ss.insertSheet('Completed Orders');
      const allHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues();
      completedSheet.getRange(1, 1, 1, allHeaders[0].length).setValues(allHeaders);
      completedSheet.getRange(1, 1, 1, allHeaders[0].length).setFontWeight('bold').setBackground('#f4cccc');
      completedSheet.getRange(1, 1, 1, allHeaders[0].length).createFilter();
      completedSheet.setFrozenRows(1);
      Logger.log('✓ Created Completed Orders sheet');
    }

    // Create Daily Report Log sheet
    let reportLogSheet = ss.getSheetByName('Daily Report Log');
    if (!reportLogSheet) {
      reportLogSheet = ss.insertSheet('Daily Report Log');
      const reportHeaders = ['Date', 'Customer', 'Samples Count', 'Email Sent To', 'Status', 'Report Generated By'];
      reportLogSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]);
      reportLogSheet.getRange(1, 1, 1, reportHeaders.length).setFontWeight('bold').setBackground('#cfe2f3');
      reportLogSheet.setFrozenRows(1);
      Logger.log('✓ Created Daily Report Log sheet');
    }

    updateLiveOrdersView();
    lockAllHeaders();

    ui.alert(
      'Tracking System Setup Complete!',
      '✓ Status columns added\n' +
      '✓ Live Orders sheet created\n' +
      '✓ Completed Orders sheet created\n' +
      '✓ Daily Report Log created\n' +
      '✓ Headers locked on all sheets\n\n' +
      'Next: Use "Barcode Scanner" to scan samples!',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('Error setting up tracking system: ' + e.toString());
    Logger.log('ERROR: ' + e.toString());
  }
}

function setupCustomerEmailSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  var sheet = ss.getSheetByName('Customer Emails');
  if (sheet) {
    ui.alert('Customer Emails sheet already exists');
    return;
  }

  sheet = ss.insertSheet('Customer Emails');

  var headers = ['Customer Name', 'Match Pattern', 'Emails', 'Active', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4A7C59').setFontColor('white');
  sheet.setFrozenRows(1);

  var sampleData = [
    ['Serengeti Trading Company', 'Serengeti Trading Company', 'tyler@serengeti.com, sales@serengeti.com', true, 'Comma-separated emails'],
    ['Paragon Coffee', 'Paragon Coffee Trading', 'orders@paragoncoffee.com', true, ''],
    ['Louis Dreyfus', 'Louis Dreyfus', '', true, 'Add emails'],
    ['Olam', 'Olam', '', true, ''],
    ['Volcafe', 'Volcafe', '', true, '']
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  sheet.getRange(2, 4, sampleData.length, 1).insertCheckboxes();

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 250);
  sheet.getRange(2, 3, 100, 1).setWrap(true);

  ui.alert('✅ Customer Emails sheet created!');
}

function setupConditionalFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(CONFIG.mainSheetName + ' sheet not found!');
    return;
  }

  var col = _getColumnMap(sheet);
  var reviewCol = col['Needs Review'];

  if (reviewCol === undefined) {
    SpreadsheetApp.getUi().alert('Add "Needs Review" column to your sheet first!');
    return;
  }

  var lastRow = Math.max(sheet.getLastRow(), 1000);
  var lastCol = sheet.getLastColumn();
  var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);

  var colLetter = String.fromCharCode(65 + reviewCol);

  sheet.clearConditionalFormatRules();

  var rules = [];

  var yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + colLetter + '2<>""')
    .setBackground('#fff3cd')
    .setRanges([dataRange])
    .build();
  rules.push(yellowRule);

  var statusCol = col['Status'];
  if (statusCol !== undefined) {
    var statusLetter = String.fromCharCode(65 + statusCol);
    var greenRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusLetter + '2="Completed"')
      .setBackground('#d4edda')
      .setRanges([dataRange])
      .build();
    rules.push(greenRule);
  }

  sheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert('✅ Conditional formatting applied!\n\n• Yellow = Needs Review\n• Green = Completed');
}

function setupFieldColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  var ui = SpreadsheetApp.getUi();
  if (!sheet) { ui.alert('Main sheet not found!'); return; }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastCol = headers.length;

  var newCols = ['Sample Type', 'Shipping Line', 'Shipping Notes'];
  var added = [];
  newCols.forEach(function(colName) {
    if (headers.indexOf(colName) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(colName).setFontWeight('bold');
      added.push(colName);
    }
  });

  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastRow = Math.max(sheet.getLastRow(), 2);

  var stIdx = headers.indexOf('Sample Type');
  if (stIdx !== -1 && lastRow > 1) {
    var stRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Warehouse Sample', 'Photos With Sample', 'Container Supervision', 'FCC Grading-Cocoa', 'Exchange Samples'], true)
      .setAllowInvalid(true).build();
    sheet.getRange(2, stIdx + 1, lastRow - 1, 1).setDataValidation(stRule);
  }

  var slIdx = headers.indexOf('Shipping Line');
  if (slIdx !== -1 && lastRow > 1) {
    var slRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['MSC', 'ZIM', 'HAPAG-LLOYD', 'MAERSK', 'SEABOARD MARINE', 'ONE', 'CMA-CGM', 'EVERGREEN', 'COSCO'], true)
      .setAllowInvalid(true).build();
    sheet.getRange(2, slIdx + 1, lastRow - 1, 1).setDataValidation(slRule);
  }

  var csIdx = headers.indexOf('Container Status');
  if (csIdx !== -1 && lastRow > 1) {
    var csRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Afloat', 'Discharged At Pier', 'Picked Up By Warehouse', 'Stripped In', 'Not Selected', 'Customs Hold', 'Cancelled', 'Warehouse Investigating'], true)
      .setAllowInvalid(true).build();
    sheet.getRange(2, csIdx + 1, lastRow - 1, 1).setDataValidation(csRule);
  }

  ui.alert('✅ Field Report columns ready!\n\n' +
    (added.length > 0 ? 'Added: ' + added.join(', ') + '\n' : 'All columns already exist.\n') +
    'Dropdowns set for: Sample Type, Shipping Line, Container Status\n' +
    'Tip: Dropdowns allow custom values — just type to add new options.');
}

function setupWarehouseEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Warehouse Emails');

  if (!sheet) {
    sheet = ss.insertSheet('Warehouse Emails');
    var headers = ['Warehouse', 'Emails', 'Active', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 350);
    sheet.setColumnWidth(3, 60);
    sheet.setColumnWidth(4, 250);

    var sampleData = [
      ['Continental Terminal', '', 'Y', ''],
      ['RPM Avenel', '', 'Y', ''],
      ['Keurig Green Mountain', '', 'Y', ''],
      ['GreenStar', '', 'Y', ''],
      ['Dupuy Storage', '', 'Y', '']
    ];
    sheet.getRange(2, 1, sampleData.length, 4).setValues(sampleData);
    SpreadsheetApp.getUi().alert('✅ Warehouse Emails sheet created!\n\nAdd email addresses (comma-separated) for each warehouse.');
  } else {
    SpreadsheetApp.getUi().alert('Warehouse Emails sheet already exists. Opening it now.');
    ss.setActiveSheet(sheet);
  }
}

function setupFieldReportEditors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Field Report Editors');

  // Upgrade existing sheet if it has old format
  if (sheet) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('Allowed Columns') === -1) {
      var notesIdx = headers.indexOf('Notes');
      if (notesIdx === -1) notesIdx = headers.length;
      sheet.insertColumnAfter(notesIdx);
      sheet.getRange(1, notesIdx + 1).setValue('Allowed Columns').setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
      sheet.setColumnWidth(notesIdx + 1, 400);
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        for (var r = 2; r <= lastRow; r++) {
          sheet.getRange(r, notesIdx + 1).setValue('ALL');
        }
      }
      SpreadsheetApp.getUi().alert(
        '✅ Field Report Editors sheet upgraded!\n\n' +
        '"Allowed Columns" column added. All existing editors set to ALL access.\n\n' +
        'To restrict a user, change their Allowed Columns to a comma-separated list:\n' +
        'e.g. Container Status, Shipping Notes, Container ETA'
      );
      return;
    }
    SpreadsheetApp.getUi().alert('Field Report Editors sheet already exists.\nEdit it directly to manage users and permissions.');
    return;
  }

  sheet = ss.insertSheet('Field Report Editors');
  sheet.getRange(1, 1).setValue('Email');
  sheet.getRange(1, 2).setValue('Name');
  sheet.getRange(1, 3).setValue('Allowed Columns');
  sheet.getRange(1, 4).setValue('Notes');

  sheet.getRange(2, 1).setValue(Session.getActiveUser().getEmail());
  sheet.getRange(2, 2).setValue('Owner');
  sheet.getRange(2, 3).setValue('ALL');
  sheet.getRange(2, 4).setValue('Full access — auto-added on setup');

  sheet.getRange(3, 1).setValue('warehouse@example.com');
  sheet.getRange(3, 2).setValue('Example — Warehouse Staff');
  sheet.getRange(3, 3).setValue('Container Status, Shipping Notes, Container ETA, Cargo #');
  sheet.getRange(3, 4).setValue('DELETE THIS ROW — example only');

  sheet.getRange(4, 1).setValue('field@example.com');
  sheet.getRange(4, 2).setValue('Example — Field Worker');
  sheet.getRange(4, 3).setValue('Container Status, Shipping Notes, Container ETA');
  sheet.getRange(4, 4).setValue('DELETE THIS ROW — example only');

  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 200);

  sheet.getRange(6, 1).setValue('📋 Valid column names for "Allowed Columns":');
  sheet.getRange(7, 1).setValue('ALL = everything  |  Or pick from:');
  sheet.getRange(8, 1).setValue('Reference, Description, Mark #, Container #, Cargo #, Receiver, Bag Count, Sample Weight, Shipping Line, Shipping Notes, Container Status, Container ETA');
  sheet.getRange(6, 1, 3, 1).setFontColor('#666').setFontStyle('italic');

  SpreadsheetApp.getUi().alert(
    '✅ Field Report Editors sheet created!\n\n' +
    'Your email was added with ALL access.\n\n' +
    'To add editors:\n' +
    '• Enter their email in column A\n' +
    '• Set Allowed Columns to ALL or a comma-separated list\n' +
    '• Delete the example rows\n\n' +
    'Column names: Reference, Description, Mark #, Container #, Cargo #,\n' +
    'Receiver, Bag Count, Sample Weight, Shipping Line, Shipping Notes,\n' +
    'Container Status, Container ETA'
  );
}

// ============================================================
// HEADER PROTECTION
// ============================================================

function lockAllHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ['All Orders', 'Live Orders', 'Completed Orders', 'Scan Log'];
  var locked = [];

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(function(p) {
      if (p.getDescription() === 'Header Row - Locked') {
        p.remove();
      }
    });

    var protection = sheet.getRange(1, 1, 1, lastCol).protect()
      .setDescription('Header Row - Locked');

    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    locked.push(name);
  });

  return locked;
}

function lockAllHeadersWithAlert() {
  var locked = lockAllHeaders();
  if (locked && locked.length > 0) {
    SpreadsheetApp.getUi().alert('🔒 Headers Locked!\n\nProtected sheets:\n• ' + locked.join('\n• '));
  } else {
    SpreadsheetApp.getUi().alert('No sheets found to lock.');
  }
}

function unlockHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ['ALL Orders', 'Live Orders', 'Completed Orders', 'Scan Log'];
  var unlocked = [];

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;

    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(function(p) {
      if (p.getDescription() === 'Header Row - Locked') {
        p.remove();
        unlocked.push(name);
      }
    });
  });

  SpreadsheetApp.getUi().alert('🔓 Headers unlocked' + (unlocked.length > 0 ? ' on: ' + unlocked.join(', ') : ' (none were locked)'));
}

// ============================================================
// SYNC PRINT COLUMN TO ALL SHEETS
// ============================================================

function syncPrintColumnToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!mainSheet) {
    ui.alert('Main sheet not found!');
    return;
  }

  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const printIdx = mainHeaders.indexOf('Print');

  if (printIdx === -1) {
    ui.alert('Print column not found in main sheet. Run Setup first.');
    return;
  }

  const printColNum = printIdx + 1;
  Logger.log('Print column found at position ' + printColNum);

  const sheetsToSync = [
    CONFIG.liveOrdersSheetName,
    CONFIG.completedOrdersSheetName || 'Completed Orders',
    'View - Danboy1217'
  ];

  const synced = [];

  sheetsToSync.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(sheetName + ': not found, skipping');
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const existingPrintIdx = headers.indexOf('Print');

    if (existingPrintIdx === -1) {
      if (sheet.getLastColumn() >= printColNum) {
        sheet.insertColumnBefore(printColNum);
      }
      sheet.getRange(1, printColNum).setValue('Print');
      sheet.getRange(1, printColNum).setFontWeight('bold');
      sheet.setColumnWidth(printColNum, 50);

      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, printColNum, lastRow - 1, 1).insertCheckboxes();
      }
      synced.push(sheetName + ' (column inserted at ' + printColNum + ')');
    } else if (existingPrintIdx !== printIdx) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, existingPrintIdx + 1, lastRow - 1, 1).insertCheckboxes();
      }
      synced.push(sheetName + ' (checkboxes refreshed)');
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, printColNum, lastRow - 1, 1).insertCheckboxes();
      }
      synced.push(sheetName + ' (verified)');
    }
  });

  ui.alert('Print Column Sync Complete!\n\n• ' + synced.join('\n• '));
}
