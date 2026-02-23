// ============================================================
// Setup.gs — Sheet setup, header protection, conditional formatting
// ============================================================

function setupSheet() {
  Logger.log('setupSheet: Starting sheet setup');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
    if (!mainSheet) {
      Logger.log('setupSheet: Creating main sheet ' + CONFIG.mainSheetName);
      mainSheet = ss.insertSheet(CONFIG.mainSheetName);
    }

  const headers = [
    'Timestamp', 'CS Order #', 'CS Sample #', 'Sender', 'Receiver', 'Warehouse',
    'Description', 'Sample Order #', 'Cargo #', 'Print', 'Mark #', 'Container #', 'Reference',
    'Bag Count', 'Weight', 'Sample Weight', 'P #', 'S #', 'Shipping Process',
    'Comments', 'Source Email', 'QR Data',
    'Status', 'Scanned Date', 'Scanned By', 'Shipped Date', 'Tracking Number', 'Report Date', 'Report Sent To',
    'Email Link', 'Attachments', 'Photos', 'Container Status', 'Container ETA',
    'Sample Type', 'Shipping Line', 'Shipping Notes', 'B/L #', 'Ship Status'
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
    const updatedValues = statusValues.map(row => {
      var cellValue = String(row[0] || '').trim();
      return cellValue === '' ? [CONFIG.statusValues.RECEIVED || 'Received'] : row;
    });
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
    if (!scanSheet) {
      Logger.log('setupSheet: Creating scan log sheet ' + CONFIG.scanSheetName);
      scanSheet = ss.insertSheet(CONFIG.scanSheetName);
    }

    const scanHeaders = ['Scan Timestamp', 'CS Order #', 'CS Sample #', 'Sample Order #', 'Cargo #', 'Mark #', 'Container #', 'Reference', 'Description', 'Bag Count', 'Weight', 'Sample Weight', 'P #', 'S #', 'Warehouse', 'Shipping Process', 'Comments', 'Scanned By', 'Notes'];
    if (scanSheet && scanHeaders.length > 0) {
      scanSheet.getRange(1, 1, 1, scanHeaders.length).setValues([scanHeaders]);
      scanSheet.getRange(1, 1, 1, scanHeaders.length).setFontWeight('bold');
      scanSheet.setFrozenRows(1);
    }

    try {
      Logger.log('setupSheet: Locking headers');
      lockAllHeaders();
    } catch(e) {
      Logger.log('setupSheet: Header lock skipped: ' + e.toString());
    }

    Logger.log('setupSheet: Complete');
    SpreadsheetApp.getUi().alert('Setup Complete!\n\n✓ Tracking columns added\n✓ Print column with checkboxes added\n✓ Attachment columns added\n✓ Container tracking columns added\n✓ Header row locked\n✓ Existing rows set to "Received" status');
  } catch (e) {
    Logger.log('setupSheet error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Setup error: ' + e.toString());
  }
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
      const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues();
      if (mainHeaders && mainHeaders.length > 0 && mainHeaders[0].length > 0) {
        liveSheet.getRange(1, 1, 1, mainHeaders[0].length).setValues(mainHeaders);
        liveSheet.getRange(1, 1, 1, mainHeaders[0].length).setFontWeight('bold').setBackground('#d9ead3');
        liveSheet.getRange(1, 1, 1, mainHeaders[0].length).createFilter();
        liveSheet.setFrozenRows(1);
        Logger.log('✓ Created Live Orders sheet');
      }
    }

    // Create Completed Orders sheet
    let completedSheet = ss.getSheetByName('Completed Orders');
    if (!completedSheet) {
      completedSheet = ss.insertSheet('Completed Orders');
      const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues();
      if (mainHeaders && mainHeaders.length > 0 && mainHeaders[0].length > 0) {
        completedSheet.getRange(1, 1, 1, mainHeaders[0].length).setValues(mainHeaders);
        completedSheet.getRange(1, 1, 1, mainHeaders[0].length).setFontWeight('bold').setBackground('#f4cccc');
        completedSheet.getRange(1, 1, 1, mainHeaders[0].length).createFilter();
        completedSheet.setFrozenRows(1);
        Logger.log('✓ Created Completed Orders sheet');
      }
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
  Logger.log('setupConditionalFormatting: Starting');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.mainSheetName);

    if (!sheet) {
      Logger.log('setupConditionalFormatting: Sheet not found - ' + CONFIG.mainSheetName);
      SpreadsheetApp.getUi().alert(CONFIG.mainSheetName + ' sheet not found!');
      return;
    }

    var col = _getColumnMap(sheet);
    var reviewCol = col['Needs Review'];

    if (reviewCol === undefined) {
      Logger.log('setupConditionalFormatting: Needs Review column not found');
      SpreadsheetApp.getUi().alert('Add "Needs Review" column to your sheet first!');
      return;
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      Logger.log('setupConditionalFormatting: sheet is empty');
      SpreadsheetApp.getUi().alert('Sheet must have data and columns!');
      return;
    }
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);

    var colLetter = _colToLetter(reviewCol);

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
      var statusLetter = _colToLetter(statusCol);
      var greenRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + statusLetter + '2="Completed"')
        .setBackground('#d4edda')
        .setRanges([dataRange])
        .build();
      rules.push(greenRule);
    }

    sheet.setConditionalFormatRules(rules);

    Logger.log('setupConditionalFormatting: Complete - ' + rules.length + ' rules applied');
    SpreadsheetApp.getUi().alert('✅ Conditional formatting applied!\n\n• Yellow = Needs Review\n• Green = Completed');
  } catch (e) {
    Logger.log('setupConditionalFormatting error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error: ' + e.toString());
  }
}

function setupFieldColumns() {
  Logger.log('setupFieldColumns: Starting');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.mainSheetName);
    var ui = SpreadsheetApp.getUi();
    if (!sheet) {
      Logger.log('setupFieldColumns: Main sheet not found');
      ui.alert('Main sheet not found!');
      return;
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastCol = headers.length;

    // Guard against null/undefined headers
    if (!headers || headers.length === 0) {
      Logger.log('setupFieldColumns: no headers found');
      ui.alert('Sheet must have headers!');
      return;
    }

    var newCols = ['Sample Type', 'Shipping Line', 'Shipping Notes', 'B/L #', 'Ship Status'];
    var added = [];
    newCols.forEach(function(colName) {
      var colNameStr = String(colName || '').trim();
      if (colNameStr && headers.indexOf(colNameStr) === -1) {
        lastCol++;
        sheet.getRange(1, lastCol).setValue(colNameStr).setFontWeight('bold');
        added.push(colNameStr);
      }
    });

    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) lastRow = 2;

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

    var ssIdx = headers.indexOf('Ship Status');
    if (ssIdx !== -1 && lastRow > 1) {
      var ssRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Afloat', 'Landed', 'Discharged', 'At Warehouse', 'In Transit', 'Customs Hold', 'Released'], true)
        .setAllowInvalid(true).build();
      sheet.getRange(2, ssIdx + 1, lastRow - 1, 1).setDataValidation(ssRule);
    }

    Logger.log('setupFieldColumns: Complete - added ' + added.length + ' columns');
    ui.alert('✅ Field Report columns ready!\n\n' +
      (added.length > 0 ? 'Added: ' + added.join(', ') + '\n' : 'All columns already exist.\n') +
      'Dropdowns set for: Sample Type, Shipping Line, Container Status, Ship Status\n' +
      'Tip: Dropdowns allow custom values — just type to add new options.');
  } catch (e) {
    Logger.log('setupFieldColumns error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error: ' + e.toString());
  }
}

function setupWarehouseEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Warehouse Emails');

  if (!sheet) {
    sheet = ss.insertSheet('Warehouse Emails');
    var headers = ['Warehouse', 'Emails', 'Active', 'Notes'];
    if (headers.length > 0) {
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
      if (sampleData.length > 0) {
        sheet.getRange(2, 1, sampleData.length, 4).setValues(sampleData);
      }
      SpreadsheetApp.getUi().alert('✅ Warehouse Emails sheet created!\n\nAdd email addresses (comma-separated) for each warehouse.');
    }
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
    if (!headers || headers.length === 0) {
      SpreadsheetApp.getUi().alert('Field Report Editors sheet has no headers. Please create it manually.');
      return;
    }
    if (headers.indexOf('Allowed Columns') === -1) {
      var notesIdx = headers.indexOf('Notes');
      if (notesIdx === -1) notesIdx = headers.length;
      sheet.insertColumnAfter(notesIdx);
      if (typeof _clearColumnMapCache === 'function') _clearColumnMapCache();
      sheet.getRange(1, notesIdx + 1).setValue('Allowed Columns').setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
      sheet.setColumnWidth(notesIdx + 1, 400);
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2 && lastRow <= 10000) { // Guard against absurdly large sheets
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
  Logger.log('lockAllHeaders: Starting header protection');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = ['All Orders', 'Live Orders', 'Completed Orders', 'Scan Log'];
    var locked = [];

    sheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) {
        Logger.log('lockAllHeaders: Sheet not found - ' + name);
        return;
      }

      var lastCol = sheet.getLastColumn();
      if (lastCol < 1) {
        Logger.log('lockAllHeaders: Sheet has no columns - ' + name);
        return;
      }

      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      if (protections && protections.length > 0) {
        protections.forEach(function(p) {
          if (String(p.getDescription() || '') === 'Header Row - Locked') {
            p.remove();
          }
        });
      }

      var protection = sheet.getRange(1, 1, 1, lastCol).protect()
        .setDescription('Header Row - Locked');

      var editors = protection.getEditors();
      if (editors && editors.length > 0) {
        protection.removeEditors(editors);
      }
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }

      locked.push(name);
      Logger.log('lockAllHeaders: Locked ' + name);
    });

    Logger.log('lockAllHeaders: Complete - locked ' + locked.length + ' sheets');
    return locked;
  } catch (e) {
    Logger.log('lockAllHeaders error: ' + e.toString());
    return [];
  }
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
  Logger.log('unlockHeaders: Starting header unlock');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = ['All Orders', 'Live Orders', 'Completed Orders', 'Scan Log'];
    var unlocked = [];

    sheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) {
        Logger.log('unlockHeaders: Sheet not found - ' + name);
        return;
      }

      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      if (protections && protections.length > 0) {
        protections.forEach(function(p) {
          if (String(p.getDescription() || '') === 'Header Row - Locked') {
            p.remove();
            unlocked.push(name);
            Logger.log('unlockHeaders: Unlocked ' + name);
          }
        });
      }
    });

    Logger.log('unlockHeaders: Complete - unlocked ' + unlocked.length + ' sheets');
    SpreadsheetApp.getUi().alert('🔓 Headers unlocked' + (unlocked.length > 0 ? ' on: ' + unlocked.join(', ') : ' (none were locked)'));
  } catch (e) {
    Logger.log('unlockHeaders error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error: ' + e.toString());
  }
}

// ============================================================
// SYNC PRINT COLUMN TO ALL SHEETS
// ============================================================

function syncPrintColumnToAllSheets() {
  Logger.log('syncPrintColumnToAllSheets: Starting');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
    if (!mainSheet) {
      Logger.log('syncPrintColumnToAllSheets: Main sheet not found');
      ui.alert('Main sheet not found!');
      return;
    }

    const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    if (!mainHeaders || mainHeaders.length === 0) {
      Logger.log('syncPrintColumnToAllSheets: Main sheet has no headers');
      ui.alert('Main sheet has no headers. Run Setup first.');
      return;
    }

    const printIdx = mainHeaders.indexOf('Print');

    if (printIdx === -1) {
      Logger.log('syncPrintColumnToAllSheets: Print column not found');
      ui.alert('Print column not found in main sheet. Run Setup first.');
      return;
    }

    const printColNum = printIdx + 1;
    Logger.log('syncPrintColumnToAllSheets: Print column found at position ' + printColNum);

    const sheetsToSync = [
      CONFIG.liveOrdersSheetName,
      CONFIG.completedOrdersSheetName || 'Completed Orders',
      'View - Danboy1217'
    ];

    const synced = [];

    sheetsToSync.forEach(function(sheetName) {
      var sheetNameStr = String(sheetName || '').trim();
      const sheet = ss.getSheetByName(sheetNameStr);
      if (!sheet) {
        Logger.log('syncPrintColumnToAllSheets: ' + sheetNameStr + ' not found, skipping');
        return;
      }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!headers || headers.length === 0) {
        Logger.log('syncPrintColumnToAllSheets: ' + sheetNameStr + ' has no headers, skipping');
        return;
      }

      const existingPrintIdx = headers.indexOf('Print');

      if (existingPrintIdx === -1) {
        if (sheet.getLastColumn() >= printColNum) {
          sheet.insertColumnBefore(printColNum);
          if (typeof _clearColumnMapCache === 'function') _clearColumnMapCache();
        }
        sheet.getRange(1, printColNum).setValue('Print');
        sheet.getRange(1, printColNum).setFontWeight('bold');
        sheet.setColumnWidth(printColNum, 50);

        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.getRange(2, printColNum, lastRow - 1, 1).insertCheckboxes();
        }
        synced.push(sheetNameStr + ' (column inserted at ' + printColNum + ')');
      } else if (existingPrintIdx !== printIdx) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.getRange(2, existingPrintIdx + 1, lastRow - 1, 1).insertCheckboxes();
        }
        synced.push(sheetNameStr + ' (checkboxes refreshed)');
      } else {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.getRange(2, printColNum, lastRow - 1, 1).insertCheckboxes();
        }
        synced.push(sheetNameStr + ' (verified)');
      }
    });

    Logger.log('syncPrintColumnToAllSheets: Complete - synced ' + synced.length + ' sheets');
    ui.alert('Print Column Sync Complete!\n\n• ' + synced.join('\n• '));
  } catch (e) {
    Logger.log('syncPrintColumnToAllSheets error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error: ' + e.toString());
  }
}

// ============================================================
// SETUP EVERYTHING — One-click full system setup
// ============================================================

function setupEverything() {
  Logger.log('setupEverything: Starting full system setup');
  var ui = SpreadsheetApp.getUi();
  var log = [];

  ui.alert('🚀 Full System Setup',
    'This will set up the entire CSS system:\n\n' +
    '1. Main sheet + all columns\n' +
    '2. Tracking system (Scan Log, Live Orders, Completed)\n' +
    '3. Field columns + dropdowns\n' +
    '4. Conditional formatting\n' +
    '5. Print checkboxes\n' +
    '6. Header protection\n' +
    '7. Customer email sheet\n' +
    '8. Warehouse email sheet\n' +
    '9. Field report editors\n' +
    '10. Contacts sheet\n' +
    '11. Time triggers\n\n' +
    'Click OK to begin.', ui.ButtonSet.OK_CANCEL);

  try {
    // 1. Main sheet + headers
    try {
      setupSheet();
      log.push('✅ Main sheet + columns');
      Logger.log('setupEverything: Step 1 complete');
    } catch(e) {
      log.push('❌ Main sheet: ' + e.message);
      Logger.log('setupEverything: Step 1 failed: ' + e.toString());
    }

    // 2. Tracking system
    try {
      setupTrackingSystem();
      log.push('✅ Tracking system (Scan Log, Live Orders, Completed)');
      Logger.log('setupEverything: Step 2 complete');
    } catch(e) {
      log.push('❌ Tracking system: ' + e.message);
      Logger.log('setupEverything: Step 2 failed: ' + e.toString());
    }

    // 3. Field columns + dropdowns
    try {
      setupFieldColumns();
      log.push('✅ Field columns + dropdowns');
      Logger.log('setupEverything: Step 3 complete');
    } catch(e) {
      log.push('❌ Field columns: ' + e.message);
      Logger.log('setupEverything: Step 3 failed: ' + e.toString());
    }

    // 4. Conditional formatting
    try {
      setupConditionalFormatting();
      log.push('✅ Conditional formatting');
      Logger.log('setupEverything: Step 4 complete');
    } catch(e) {
      log.push('❌ Conditional formatting: ' + e.message);
      Logger.log('setupEverything: Step 4 failed: ' + e.toString());
    }

    // 5. Print checkboxes
    try {
      if (typeof addPrintCheckboxColumn === 'function') {
        addPrintCheckboxColumn();
        log.push('✅ Print checkboxes');
      }
      Logger.log('setupEverything: Step 5 complete');
    } catch(e) {
      log.push('❌ Print checkboxes: ' + e.message);
      Logger.log('setupEverything: Step 5 failed: ' + e.toString());
    }

    // 6. Header protection
    try {
      var locked = lockAllHeaders();
      log.push('✅ Headers locked (' + (locked ? locked.length : 0) + ' sheets)');
      Logger.log('setupEverything: Step 6 complete');
    } catch(e) {
      log.push('❌ Header lock: ' + e.message);
      Logger.log('setupEverything: Step 6 failed: ' + e.toString());
    }

    // 7. Customer email sheet
    try {
      setupCustomerEmailSheet();
      log.push('✅ Customer email sheet');
      Logger.log('setupEverything: Step 7 complete');
    } catch(e) {
      log.push('❌ Customer emails: ' + e.message);
      Logger.log('setupEverything: Step 7 failed: ' + e.toString());
    }

    // 8. Warehouse email sheet
    try {
      setupWarehouseEmails();
      log.push('✅ Warehouse email sheet');
      Logger.log('setupEverything: Step 8 complete');
    } catch(e) {
      log.push('❌ Warehouse emails: ' + e.message);
      Logger.log('setupEverything: Step 8 failed: ' + e.toString());
    }

    // 9. Field report editors
    try {
      setupFieldReportEditors();
      log.push('✅ Field report editors');
      Logger.log('setupEverything: Step 9 complete');
    } catch(e) {
      log.push('❌ Field report editors: ' + e.message);
      Logger.log('setupEverything: Step 9 failed: ' + e.toString());
    }

    // 10. Contacts
    try {
      if (typeof setupContacts === 'function') {
        setupContacts();
        log.push('✅ Contacts sheet');
      }
      Logger.log('setupEverything: Step 10 complete');
    } catch(e) {
      log.push('❌ Contacts: ' + e.message);
      Logger.log('setupEverything: Step 10 failed: ' + e.toString());
    }

    // 11. Time triggers
    try {
      createTimeTriggers();
      log.push('✅ Time triggers');
      Logger.log('setupEverything: Step 11 complete');
    } catch(e) {
      log.push('❌ Triggers: ' + e.message);
      Logger.log('setupEverything: Step 11 failed: ' + e.toString());
    }

    // Final report
    Logger.log('setupEverything: Complete');
    ui.alert('🎉 Setup Complete!\n\n' + log.join('\n'));
  } catch (e) {
    Logger.log('setupEverything error: ' + e.toString());
    ui.alert('Unexpected error: ' + e.toString());
  }
}
