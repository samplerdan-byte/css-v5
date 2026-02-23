// ============================================================
// 💰 BILLING MODULE — CSS Billing & QuickBooks Integration
// ============================================================
// Rate sheet based on CSS Rate Sheet effective June 4, 2022
// Supports: Invoice generation, QuickBooks Online CSV export,
//           QuickBooks Desktop IIF export, invoice tracking
//
// ⚠️  CLEANED: Removed duplicate functions that already exist in Main.gs:
//     - escapeCSVField, formatDateForSQL, getTrackingUrl, cleanReceiverName
//     These are defined once in Main.gs and shared across all .gs files.
//
// Add to existing CSS project as Billing.gs (ONE copy only)
// Then run setupBilling() from the menu to initialize
// ============================================================

// ============================================================
// RATE SHEET CONFIGURATION
// ============================================================

var BILLING_CONFIG = {
  rateSheetName: 'Rate Sheet',
  invoicesSheetName: 'Invoices',
  lineItemsSheetName: 'Invoice Line Items',
  invoicePrefix: 'CSS-',
  defaultTerms: 'Net 30',
  companyName: 'Commodity Sampler Services',
  companyTagline: 'Integrity Through Independence',
  // QuickBooks Online settings
  qboIncomeAccount: 'Sampling Services',
  qboTaxCode: 'NON'  // Non-taxable
};

// Master rate table — matches the 6/4/2022 rate sheet
// These seed the Rate Sheet tab; Dan edits rates there going forward
var MASTER_RATES = [
  // [Service Code, Service Description, Unit, Base Rate, Notes]
  ['REG',        'Regular Sample (up to 5 lbs)',              'per sample',    28.00,  ''],
  ['INT-10',     'Intensive Sample — 10%',                    'per sample',    43.00,  'Add $15 to regular'],
  ['INT-20',     'Intensive Sample — 20%',                    'per sample',    63.00,  'Add $35 to regular'],
  ['INT-30',     'Intensive Sample — 30%',                    'per sample',    98.00,  'Add $70 to regular'],
  ['INT-50',     'Intensive Sample — 50%+',                   'per sample',     0.00,  'Call for quote'],
  ['LRG-6',      'Large Sample (6–9 lbs)',                    'per sample',    38.00,  'Add $10 to regular'],
  ['LRG-10',     'Large Sample (10–25 lbs)',                  'per sample',    48.00,  'Add $20 to regular'],
  ['LRG-26',     'Large Sample (26+ lbs)',                    'per sample',    63.00,  'Add $35 to regular'],
  ['KEURIG',     'Keurig Sample 30% Sampling',                'per sample',    98.00,  ''],
  ['BARREL',     'Barrel Sampling',                           'per sample',   100.00,  ''],
  ['MICRO',      'Microlots / Small Bag Cts (<19 bags)',      'per sample',    38.00,  'Add $10 to regular'],
  ['MOIST',      'Moisture Reading',                          'per lot',       55.00,  ''],
  ['SCREEN',     'Screening Coffee Samples',                  'per lot',       55.00,  ''],
  ['INV-COUNT',  'Inventory Coffee: Physical Count',          'per lot',       30.00,  ''],
  ['BAG-MOVE',   'Bag Moves (up to 3 bags)',                  'per move',      95.00,  ''],
  ['BAG-MOVE-A', 'Bag Moves — Additional bags',              'per bag',       95.00,  'Same rate per bag'],
  ['SPLIT-BAG',  'Splitting Bags',                            'per bag',       45.00,  ''],
  ['INSPECT',    'Inspection (Infestation/Wet/Quality)',       'per inspection',110.00, ''],
  ['SIFT',       'Sifting Coffee Bag Inspection',             'per bag',      110.00,  ''],
  ['CONT-IN',    'Container Supervision — Inbound',           'per container',180.00,  ''],
  ['CONT-OUT',   'Container Supervision — Outbound',          'per container',120.00,  ''],
  ['WEIGH-IN',   'Weight Inspection — Inbound',               'per container',180.00,  ''],
  ['REWEIGH',    'Reweights',                                 'per container',120.00,  ''],
  ['PHOTO-W',    'Photographs — With Sample',                 'per sample',    10.00,  ''],
  ['PHOTO-WO',   'Photographs — Without Sample',              'per photo set', 25.00,  ''],
  ['RUSH',       'Same Day Rush Delivery NJ/NYC',             'per delivery', 150.00,  ''],
  ['INV-COCOA',  'Inventory Cocoa',                           'per lot',       55.00,  ''],
  ['INV-COCBUT', 'Inventory Cocoa Butter',                    'per lot',       79.50,  ''],
  ['VESSEL',     'Vessel Supervision',                        'per MT',         4.49,  ''],
  ['NSF',        'NSF Checks',                                'per check',     35.00,  ''],
  ['SHIP-FEE',   'Shipping (CSS FedEx/UPS Account)',          'percentage',    20.00,  'Added as 20% of subtotal']
];


// ============================================================
// SETUP — Creates Rate Sheet, Invoices, Line Items sheets
// ============================================================

function setupBilling() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var resp = ui.alert(
    'Setup Billing System',
    'This will create:\n' +
    '1. Rate Sheet — editable pricing\n' +
    '2. Invoices — invoice tracking\n' +
    '3. Invoice Line Items — detail rows\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  
  // ── Rate Sheet ──
  var rateSheet = ss.getSheetByName(BILLING_CONFIG.rateSheetName);
  if (!rateSheet) {
    rateSheet = ss.insertSheet(BILLING_CONFIG.rateSheetName);
    var rateHeaders = ['Service Code', 'Service Description', 'Unit', 'Rate', 'Notes', 'Active'];
    rateSheet.getRange(1, 1, 1, rateHeaders.length).setValues([rateHeaders]);
    rateSheet.getRange(1, 1, 1, rateHeaders.length).setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
    rateSheet.setFrozenRows(1);
    
    var rateData = MASTER_RATES.map(function(r) {
      return [r[0], r[1], r[2], r[3], r[4], true];
    });
    rateSheet.getRange(2, 1, rateData.length, 6).setValues(rateData);
    rateSheet.getRange(2, 6, rateData.length, 1).insertCheckboxes();
    
    // Format rate column as currency
    rateSheet.getRange(2, 4, rateData.length, 1).setNumberFormat('$#,##0.00');
    
    rateSheet.setColumnWidth(1, 100);
    rateSheet.setColumnWidth(2, 300);
    rateSheet.setColumnWidth(3, 120);
    rateSheet.setColumnWidth(4, 80);
    rateSheet.setColumnWidth(5, 200);
    rateSheet.setColumnWidth(6, 60);
    
    Logger.log('✓ Rate Sheet created with ' + rateData.length + ' services');
  }
  
  // ── Invoices ──
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  if (!invSheet) {
    invSheet = ss.insertSheet(BILLING_CONFIG.invoicesSheetName);
    var invHeaders = [
      'Invoice #', 'Invoice Date', 'Due Date', 'Customer',
      'Subtotal', 'Shipping Fee', 'Total', 'Status',
      'QB Exported', 'QB Export Date', 'Payment Date', 'Payment Method',
      'Notes', 'Created By'
    ];
    invSheet.getRange(1, 1, 1, invHeaders.length).setValues([invHeaders]);
    invSheet.getRange(1, 1, 1, invHeaders.length).setFontWeight('bold').setBackground('#1a3a25').setFontColor('#fff');
    invSheet.setFrozenRows(1);
    
    invSheet.getRange(2, 5, 500, 3).setNumberFormat('$#,##0.00');
    
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Draft', 'Sent', 'Paid', 'Overdue', 'Void'], true)
      .setAllowInvalid(false).build();
    invSheet.getRange(2, 8, 500, 1).setDataValidation(statusRule);
    
    invSheet.setColumnWidth(1, 100);
    invSheet.setColumnWidth(2, 100);
    invSheet.setColumnWidth(3, 100);
    invSheet.setColumnWidth(4, 200);
    invSheet.setColumnWidth(7, 100);
    invSheet.setColumnWidth(8, 80);
    
    Logger.log('✓ Invoices sheet created');
  }
  
  // ── Line Items ──
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  if (!liSheet) {
    liSheet = ss.insertSheet(BILLING_CONFIG.lineItemsSheetName);
    var liHeaders = [
      'Invoice #', 'Line #', 'Service Code', 'Service Description',
      'CS Sample #', 'CS Order #', 'Container #', 'Reference',
      'Quantity', 'Unit', 'Unit Price', 'Line Total',
      'Customer', 'Warehouse', 'Notes'
    ];
    liSheet.getRange(1, 1, 1, liHeaders.length).setValues([liHeaders]);
    liSheet.getRange(1, 1, 1, liHeaders.length).setFontWeight('bold').setBackground('#1a3a25').setFontColor('#fff');
    liSheet.setFrozenRows(1);
    
    liSheet.getRange(2, 11, 1000, 2).setNumberFormat('$#,##0.00');
    
    liSheet.setColumnWidth(1, 100);
    liSheet.setColumnWidth(4, 280);
    
    Logger.log('✓ Invoice Line Items sheet created');
  }
  
  try { lockAllHeaders(); } catch(e) {}
  
  ui.alert(
    '✅ Billing System Ready!\n\n' +
    '✓ Rate Sheet — edit prices anytime\n' +
    '✓ Invoices — tracks all invoices\n' +
    '✓ Invoice Line Items — detailed billing\n\n' +
    'Use ☕ CSS System → 💰 Billing to:\n' +
    '• Generate invoices from shipped orders\n' +
    '• Preview & print invoices\n' +
    '• Export to QuickBooks'
  );
}


// ============================================================
// AUTO-BILL: Called when samples move to Completed
// Creates line items + draft invoice automatically
// ============================================================

function _autoBillSamples(movedRows, ss) {
  try {
    if (!movedRows || movedRows.length === 0) return;

    var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
    var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
    if (!invSheet || !liSheet) return; // billing not set up yet — skip silently

    var rates = _loadRates();
    if (Object.keys(rates).length === 0) return; // no rate sheet

    // Get column map from Completed Orders (same headers as All Orders)
    var completedSheet = ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders');
    if (!completedSheet) return;
    var col = _getColumnMap(completedSheet);

    // Build set of already-billed samples
    var billedSamples = {};
    if (liSheet.getLastRow() >= 2) {
      var billedData = liSheet.getRange(2, 5, liSheet.getLastRow() - 1, 1).getValues();
      for (var b = 0; b < billedData.length; b++) {
        var bs = String(billedData[b][0]).trim();
        if (bs) billedSamples[bs] = true;
      }
    }

    // Build sample objects from row data
    var unbilledSamples = [];
    for (var i = 0; i < movedRows.length; i++) {
      var row = movedRows[i];
      var csSample = col['CS Sample #'] !== undefined ? String(row[col['CS Sample #']] || '').trim() : '';
      if (!csSample || billedSamples[csSample]) continue;

      unbilledSamples.push({
        csSample: csSample,
        csOrder: col['CS Order #'] !== undefined ? String(row[col['CS Order #']] || '').trim() : '',
        sender: col['Sender'] !== undefined ? String(row[col['Sender']] || '').trim() : '',
        receiver: col['Receiver'] !== undefined ? String(row[col['Receiver']] || '').trim() : '',
        warehouse: col['Warehouse'] !== undefined ? String(row[col['Warehouse']] || '').trim() : '',
        container: col['Container #'] !== undefined ? String(row[col['Container #']] || '').trim() : '',
        reference: col['Reference'] !== undefined ? String(row[col['Reference']] || '').trim() : '',
        sampleOrderNum: col['Sample Order #'] !== undefined ? String(row[col['Sample Order #']] || '').trim() : '',
        bagCount: col['Bag Count'] !== undefined ? String(row[col['Bag Count']] || '').trim() : '',
        sampleWeight: col['Sample Weight'] !== undefined ? String(row[col['Sample Weight']] || '').trim() : '',
        sampleType: col['Sample Type'] !== undefined ? String(row[col['Sample Type']] || '').trim() : '',
        shippingProcess: col['Shipping Process'] !== undefined ? String(row[col['Shipping Process']] || '').trim() : '',
        trackingNumber: col['Tracking Number'] !== undefined ? String(row[col['Tracking Number']] || '').trim() : '',
        description: col['Description'] !== undefined ? String(row[col['Description']] || '').trim() : ''
      });
    }

    if (unbilledSamples.length === 0) return;

    // Group by customer (sender)
    var byCustomer = {};
    for (var j = 0; j < unbilledSamples.length; j++) {
      var cust = unbilledSamples[j].sender || 'Unknown';
      if (!byCustomer[cust]) byCustomer[cust] = [];
      byCustomer[cust].push(unbilledSamples[j]);
    }

    var today = new Date();
    var dueDate = new Date(today);
    dueDate.setDate(dueDate.getDate() + 30);
    var totalBilled = 0;

    var customerNames = Object.keys(byCustomer).sort();
    for (var ci = 0; ci < customerNames.length; ci++) {
      var customerName = customerNames[ci];
      var samples = byCustomer[customerName];

      var invoiceNum = _getNextInvoiceNumber();
      var subtotal = 0;
      var lineNum = 0;
      var lineItemRows = [];

      for (var si = 0; si < samples.length; si++) {
        var sample = samples[si];
        lineNum++;
        var serviceCode = _detectServiceCode(sample, rates);
        var rateInfo = rates[serviceCode] || rates['REG'] || { description: 'Regular Sample', rate: 28.00, unit: 'per sample' };
        var qty = 1;
        var unitPrice = rateInfo.rate;
        var lineTotal = qty * unitPrice;
        subtotal += lineTotal;

        lineItemRows.push([
          invoiceNum, lineNum, serviceCode, rateInfo.description,
          sample.csSample, sample.csOrder, sample.container,
          sample.reference || sample.sampleOrderNum,
          qty, rateInfo.unit, unitPrice, lineTotal,
          customerName, sample.warehouse, sample.description
        ]);
      }

      // Write line items
      if (lineItemRows.length > 0) {
        var liLastRow = liSheet.getLastRow() + 1;
        liSheet.getRange(liLastRow, 1, lineItemRows.length, lineItemRows[0].length).setValues(lineItemRows);
        liSheet.getRange(liLastRow, 11, lineItemRows.length, 2).setNumberFormat('$#,##0.00');
        totalBilled += lineItemRows.length;
      }

      // Write invoice header
      var total = subtotal;
      invSheet.appendRow([
        invoiceNum, today, dueDate, customerName,
        subtotal, 0, total, 'Draft',
        false, '', '', '', samples.length + ' samples',
        ''
      ]);
      var invLastRow = invSheet.getLastRow();
      invSheet.getRange(invLastRow, 5, 1, 3).setNumberFormat('$#,##0.00');
      invSheet.getRange(invLastRow, 2, 1, 1).setNumberFormat('MM/dd/yyyy');
      invSheet.getRange(invLastRow, 3, 1, 1).setNumberFormat('MM/dd/yyyy');
    }

    Logger.log('Auto-billed ' + totalBilled + ' samples across ' + customerNames.length + ' invoice(s)');
  } catch (e) {
    Logger.log('Auto-bill error (non-fatal): ' + e);
  }
}


// ============================================================
// LOAD RATES FROM SHEET (so edits in the sheet are respected)
// ============================================================

function _loadRates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BILLING_CONFIG.rateSheetName);
  if (!sheet || sheet.getLastRow() < 2) return {};
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var rates = {};
  for (var i = 0; i < data.length; i++) {
    var code = String(data[i][0]).trim();
    var active = data[i][5];
    if (code && active) {
      rates[code] = {
        code: code,
        description: String(data[i][1]),
        unit: String(data[i][2]),
        rate: parseFloat(data[i][3]) || 0,
        notes: String(data[i][4])
      };
    }
  }
  return rates;
}


// ============================================================
// AUTO-DETECT SERVICE TYPE FROM ORDER DATA
// ============================================================

function _detectServiceCode(sample, rates) {
  var sType = String(sample.sampleType || sample.shippingProcess || '').toLowerCase();
  var weight = parseFloat(sample.sampleWeight) || 0;
  var bagCount = parseInt(sample.bagCount) || 0;
  
  if (sType.indexOf('container sup') !== -1 || sType.indexOf('supervision') !== -1) {
    if (sType.indexOf('out') !== -1) return 'CONT-OUT';
    return 'CONT-IN';
  }
  if (sType.indexOf('photo') !== -1) {
    if (sType.indexOf('with sample') !== -1 || sType.indexOf('with samp') !== -1) return 'PHOTO-W';
    if (sType.indexOf('without') !== -1) return 'PHOTO-WO';
    return 'PHOTO-W';
  }
  if (sType.indexOf('fcc') !== -1 || sType.indexOf('cocoa') !== -1) {
    if (sType.indexOf('butter') !== -1) return 'INV-COCBUT';
    return 'INV-COCOA';
  }
  if (sType.indexOf('exchange') !== -1) return 'REG';
  if (sType.indexOf('keurig') !== -1) return 'KEURIG';
  if (sType.indexOf('barrel') !== -1) return 'BARREL';
  
  var pctMatch = sType.match(/(\d+)\s*%/);
  if (pctMatch) {
    var pct = parseInt(pctMatch[1]);
    if (pct >= 50) return 'INT-50';
    if (pct >= 30) return 'INT-30';
    if (pct >= 20) return 'INT-20';
    if (pct >= 10) return 'INT-10';
  }
  if (sType.indexOf('intensive') !== -1) return 'INT-10';
  if (sType.indexOf('moisture') !== -1) return 'MOIST';
  if (sType.indexOf('screen') !== -1) return 'SCREEN';
  if (sType.indexOf('sift') !== -1) return 'SIFT';
  if (sType.indexOf('inspect') !== -1) return 'INSPECT';
  if (sType.indexOf('vessel') !== -1) return 'VESSEL';
  if (bagCount > 0 && bagCount < 19) return 'MICRO';
  
  if (weight > 0) {
    if (weight >= 26) return 'LRG-26';
    if (weight >= 10) return 'LRG-10';
    if (weight >= 6)  return 'LRG-6';
  }
  
  return 'REG';
}


// ============================================================
// GENERATE NEXT INVOICE NUMBER
// ============================================================

function _getNextInvoiceNumber() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  if (!sheet || sheet.getLastRow() < 2) return BILLING_CONFIG.invoicePrefix + '1001';
  
  var invNums = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var highest = 1000;
  for (var i = 0; i < invNums.length; i++) {
    var num = String(invNums[i][0]).replace(BILLING_CONFIG.invoicePrefix, '');
    var parsed = parseInt(num);
    if (!isNaN(parsed) && parsed > highest) highest = parsed;
  }
  return BILLING_CONFIG.invoicePrefix + (highest + 1);
}


// ============================================================
// GENERATE INVOICE FROM COMPLETED ORDERS
// ============================================================

function generateInvoice() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  var completedSheet = ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders');
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  
  if (!mainSheet) { ui.alert('Main sheet not found. Run Setup first.'); return; }
  
  var rates = _loadRates();
  if (Object.keys(rates).length === 0) {
    ui.alert('No rates found. Run Setup → Setup Billing first.');
    return;
  }
  
  var billedSamples = {};
  if (liSheet && liSheet.getLastRow() >= 2) {
    var billedData = liSheet.getRange(2, 5, liSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < billedData.length; i++) {
      var s = String(billedData[i][0]).trim();
      if (s) billedSamples[s] = true;
    }
  }
  
  var allSamples = [];
  
  function gatherFromSheet(sheet) {
    if (!sheet || sheet.getLastRow() < 2) return;
    var col = _getColumnMap(sheet);
    if (col['Status'] === undefined) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][col['Status']] || '').trim().toLowerCase();
      if (status !== 'shipped' && status !== 'completed') continue;
      
      var csSample = String(data[i][col['CS Sample #']] || '').trim();
      if (!csSample || billedSamples[csSample]) continue;
      
      allSamples.push({
        csSample: csSample,
        csOrder: String(data[i][col['CS Order #']] || '').trim(),
        sender: String(data[i][col['Sender']] || '').trim(),
        receiver: String(data[i][col['Receiver']] || '').trim(),
        warehouse: String(data[i][col['Warehouse']] || '').trim(),
        container: String(data[i][col['Container #']] || '').trim(),
        cargo: String(data[i][col['Cargo #']] || '').trim(),
        mark: String(data[i][col['Mark #']] || '').trim(),
        reference: String(data[i][col['Reference']] || '').trim(),
        sampleOrderNum: String(data[i][col['Sample Order #']] || '').trim(),
        bagCount: String(data[i][col['Bag Count']] || '').trim(),
        sampleWeight: String(data[i][col['Sample Weight']] || '').trim(),
        sampleType: col['Sample Type'] !== undefined ? String(data[i][col['Sample Type']] || '').trim() : '',
        shippingProcess: String(data[i][col['Shipping Process']] || '').trim(),
        trackingNumber: col['Tracking Number'] !== undefined ? String(data[i][col['Tracking Number']] || '').trim() : '',
        shippedDate: col['Shipped Date'] !== undefined ? data[i][col['Shipped Date']] : '',
        description: String(data[i][col['Description']] || '').trim()
      });
    }
  }
  
  gatherFromSheet(mainSheet);
  gatherFromSheet(completedSheet);
  
  if (allSamples.length === 0) {
    ui.alert('No unbilled shipped/completed orders found.\n\nAll shipped samples have already been invoiced.');
    return;
  }
  
  var byCustomer = {};
  allSamples.forEach(function(s) {
    var cust = s.sender || 'Unknown';
    if (!byCustomer[cust]) byCustomer[cust] = [];
    byCustomer[cust].push(s);
  });
  
  var customerNames = Object.keys(byCustomer).sort();
  
  var selectHtml = '<html><head><style>' +
    'body { font-family: Arial; padding: 15px; }' +
    '.cust { padding: 8px; margin: 5px 0; background: #f0f7f2; border-radius: 5px; display: flex; align-items: center; }' +
    '.cust label { flex: 1; cursor: pointer; font-size: 13px; }' +
    '.count { color: #666; font-size: 11px; }' +
    '.btn { padding: 10px 20px; background: #4A7C59; color: white; border: none; border-radius: 5px; font-size: 13px; cursor: pointer; margin: 5px; }' +
    '.btn:hover { background: #3a6249; }' +
    '.btn.sec { background: #6c757d; }' +
    '.btns { text-align: center; margin-top: 15px; }' +
    '.summary { background: #e8f5e9; padding: 10px; border-radius: 5px; margin-bottom: 10px; text-align: center; font-size: 13px; }' +
    '</style></head><body>' +
    '<div class="summary"><strong>' + allSamples.length + ' unbilled samples</strong> across ' + customerNames.length + ' customer(s)</div>';
  
  customerNames.forEach(function(cust, idx) {
    var samples = byCustomer[cust];
    selectHtml += '<div class="cust">' +
      '<input type="checkbox" id="c' + idx + '" checked value="' + cust.replace(/"/g, '&quot;') + '">' +
      '<label for="c' + idx + '">&nbsp;' + cust + ' <span class="count">(' + samples.length + ' samples)</span></label>' +
      '</div>';
  });
  
  selectHtml += '<div class="btns">' +
    '<button class="btn" onclick="generate()">💰 Generate Invoice(s)</button>' +
    '<button class="btn sec" onclick="google.script.host.close()">Cancel</button>' +
    '</div>' +
    '<script>' +
    'function generate() {' +
    '  var checked = [];' +
    '  document.querySelectorAll("input[type=checkbox]:checked").forEach(function(cb) { checked.push(cb.value); });' +
    '  if (checked.length === 0) { alert("Select at least one customer"); return; }' +
    '  google.script.run.withSuccessHandler(function(r) { alert(r); google.script.host.close(); })' +
    '    .withFailureHandler(function(e) { alert("Error: " + e.message); })' +
    '    .generateInvoicesForCustomers(checked);' +
    '}' +
    '</script></body></html>';
  
  var output = HtmlService.createHtmlOutput(selectHtml).setWidth(450).setHeight(400);
  ui.showModalDialog(output, '💰 Generate Invoices');
}


// ============================================================
// GENERATE INVOICES FOR SELECTED CUSTOMERS
// ============================================================

function generateInvoicesForCustomers(customerList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  
  if (!invSheet || !liSheet) return 'Error: Billing sheets not found. Run Setup Billing first.';
  
  var rates = _loadRates();
  
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  var completedSheet = ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders');
  
  var billedSamples = {};
  if (liSheet.getLastRow() >= 2) {
    var billedData = liSheet.getRange(2, 5, liSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < billedData.length; i++) {
      var s = String(billedData[i][0]).trim();
      if (s) billedSamples[s] = true;
    }
  }
  
  var allSamples = [];
  function gatherFromSheet(sheet) {
    if (!sheet || sheet.getLastRow() < 2) return;
    var col = _getColumnMap(sheet);
    if (col['Status'] === undefined) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][col['Status']] || '').trim().toLowerCase();
      if (status !== 'shipped' && status !== 'completed') continue;
      var csSample = String(data[i][col['CS Sample #']] || '').trim();
      if (!csSample || billedSamples[csSample]) continue;
      allSamples.push({
        csSample: csSample,
        csOrder: String(data[i][col['CS Order #']] || '').trim(),
        sender: String(data[i][col['Sender']] || '').trim(),
        receiver: String(data[i][col['Receiver']] || '').trim(),
        warehouse: String(data[i][col['Warehouse']] || '').trim(),
        container: String(data[i][col['Container #']] || '').trim(),
        reference: String(data[i][col['Reference']] || '').trim(),
        sampleOrderNum: String(data[i][col['Sample Order #']] || '').trim(),
        bagCount: String(data[i][col['Bag Count']] || '').trim(),
        sampleWeight: String(data[i][col['Sample Weight']] || '').trim(),
        sampleType: col['Sample Type'] !== undefined ? String(data[i][col['Sample Type']] || '').trim() : '',
        shippingProcess: String(data[i][col['Shipping Process']] || '').trim(),
        trackingNumber: col['Tracking Number'] !== undefined ? String(data[i][col['Tracking Number']] || '').trim() : '',
        description: String(data[i][col['Description']] || '').trim()
      });
    }
  }
  gatherFromSheet(mainSheet);
  gatherFromSheet(completedSheet);
  
  var byCustomer = {};
  allSamples.forEach(function(s) {
    var cust = s.sender || 'Unknown';
    if (!byCustomer[cust]) byCustomer[cust] = [];
    byCustomer[cust].push(s);
  });
  
  var invoicesCreated = 0;
  var totalLineItems = 0;
  var today = new Date();
  var dueDate = new Date(today);
  dueDate.setDate(dueDate.getDate() + 30);
  
  customerList.forEach(function(customerName) {
    var samples = byCustomer[customerName];
    if (!samples || samples.length === 0) return;
    
    var invoiceNum = _getNextInvoiceNumber();
    var subtotal = 0;
    var lineNum = 0;
    var hasShipping = false;
    var lineItemRows = [];
    
    samples.forEach(function(sample) {
      lineNum++;
      var serviceCode = _detectServiceCode(sample, rates);
      var rateInfo = rates[serviceCode] || rates['REG'] || { description: 'Regular Sample', rate: 28.00, unit: 'per sample' };
      var qty = 1;
      var unitPrice = rateInfo.rate;
      var lineTotal = qty * unitPrice;
      subtotal += lineTotal;
      
      if (sample.trackingNumber) hasShipping = true;
      
      lineItemRows.push([
        invoiceNum, lineNum, serviceCode, rateInfo.description,
        sample.csSample, sample.csOrder, sample.container,
        sample.reference || sample.sampleOrderNum,
        qty, rateInfo.unit, unitPrice, lineTotal,
        customerName, sample.warehouse, sample.description
      ]);
    });
    
    // Shipping fee NOT auto-added — most customers ship on their own account.
    // To add manually: insert a SHIP-FEE line item on the Invoice Line Items sheet.
    var shippingFee = 0;
    var total = subtotal + shippingFee;
    
    if (lineItemRows.length > 0) {
      var liLastRow = liSheet.getLastRow() + 1;
      liSheet.getRange(liLastRow, 1, lineItemRows.length, lineItemRows[0].length).setValues(lineItemRows);
      liSheet.getRange(liLastRow, 11, lineItemRows.length, 2).setNumberFormat('$#,##0.00');
      totalLineItems += lineItemRows.length;
    }
    
    invSheet.appendRow([
      invoiceNum, today, dueDate, customerName,
      subtotal, shippingFee, total, 'Draft',
      false, '', '', '', samples.length + ' samples',
      Session.getActiveUser().getEmail()
    ]);
    
    var invLastRow = invSheet.getLastRow();
    invSheet.getRange(invLastRow, 5, 1, 3).setNumberFormat('$#,##0.00');
    invSheet.getRange(invLastRow, 2, 1, 1).setNumberFormat('MM/dd/yyyy');
    invSheet.getRange(invLastRow, 3, 1, 1).setNumberFormat('MM/dd/yyyy');
    
    invoicesCreated++;
    Logger.log('✓ Invoice ' + invoiceNum + ' created for ' + customerName + ': $' + total.toFixed(2));
  });
  
  return '✅ Created ' + invoicesCreated + ' invoice(s) with ' + totalLineItems + ' line items.\n\n' +
    'View them in the "Invoices" and "Invoice Line Items" sheets.\n' +
    'Use Billing → Preview Invoice to review before sending.';
}


// ============================================================
// PREVIEW / PRINT INVOICE
// ============================================================

function previewInvoice() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  
  if (!invSheet || invSheet.getLastRow() < 2) {
    ui.alert('No invoices found. Generate an invoice first.');
    return;
  }
  
  var activeSheet = ss.getActiveSheet();
  var invNum = '';
  
  if (activeSheet.getName() === BILLING_CONFIG.invoicesSheetName) {
    var row = activeSheet.getActiveCell().getRow();
    if (row >= 2) {
      invNum = String(activeSheet.getRange(row, 1).getValue()).trim();
    }
  }
  
  if (!invNum) {
    var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 8).getValues();
    var options = invData.map(function(r) {
      return String(r[0]) + ' — ' + String(r[3]) + ' — $' + parseFloat(r[6] || 0).toFixed(2) + ' [' + String(r[7]) + ']';
    }).reverse().slice(0, 20);
    
    var resp = ui.prompt('Select Invoice', 'Enter invoice number (e.g. CSS-1001):\n\nRecent invoices:\n' + options.join('\n'), ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    invNum = resp.getResponseText().trim().toUpperCase();
  }
  
  if (!invNum) { ui.alert('No invoice selected.'); return; }
  
  var html = _generateInvoiceHtml(invNum);
  if (!html) { ui.alert('Invoice ' + invNum + ' not found.'); return; }
  
  var output = HtmlService.createHtmlOutput(html).setWidth(850).setHeight(900);
  ui.showModalDialog(output, '📄 Invoice ' + invNum);
}

function _generateInvoiceHtml(invoiceNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  
  if (!invSheet || !liSheet) return null;
  
  var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 14).getValues();
  var invoice = null;
  for (var i = 0; i < invData.length; i++) {
    if (String(invData[i][0]).trim().toUpperCase() === invoiceNum) {
      invoice = {
        number: String(invData[i][0]),
        date: invData[i][1],
        dueDate: invData[i][2],
        customer: String(invData[i][3]),
        subtotal: parseFloat(invData[i][4]) || 0,
        shippingFee: parseFloat(invData[i][5]) || 0,
        total: parseFloat(invData[i][6]) || 0,
        status: String(invData[i][7]),
        notes: String(invData[i][12])
      };
      break;
    }
  }
  if (!invoice) return null;
  
  var liData = liSheet.getRange(2, 1, liSheet.getLastRow() - 1, 15).getValues();
  var lineItems = [];
  for (var i = 0; i < liData.length; i++) {
    if (String(liData[i][0]).trim().toUpperCase() === invoiceNum) {
      lineItems.push({
        lineNum: liData[i][1],
        serviceCode: String(liData[i][2]),
        description: String(liData[i][3]),
        csSample: String(liData[i][4]),
        csOrder: String(liData[i][5]),
        container: String(liData[i][6]),
        reference: String(liData[i][7]),
        qty: parseFloat(liData[i][8]) || 1,
        unit: String(liData[i][9]),
        unitPrice: parseFloat(liData[i][10]) || 0,
        lineTotal: parseFloat(liData[i][11]) || 0,
        warehouse: String(liData[i][13]),
        notes: String(liData[i][14])
      });
    }
  }
  
  var custAddress = _getCustomerAddress(invoice.customer);
  
  function fmtDate(d) {
    if (!d) return '';
    try {
      var dt = new Date(d);
      return String(dt.getMonth() + 1).padStart(2, '0') + '/' +
        String(dt.getDate()).padStart(2, '0') + '/' + dt.getFullYear();
    } catch(e) { return ''; }
  }
  
  function fmtMoney(n) {
    return '$' + parseFloat(n || 0).toFixed(2);
  }
  
  var rowsHtml = lineItems.map(function(li) {
    return '<tr>' +
      '<td>' + li.description + '</td>' +
      '<td>' + (li.csSample || '') + '</td>' +
      '<td>' + (li.container || '') + '</td>' +
      '<td>' + (li.reference || '') + '</td>' +
      '<td class="r">' + li.qty + '</td>' +
      '<td class="r">' + fmtMoney(li.unitPrice) + '</td>' +
      '<td class="r">' + fmtMoney(li.lineTotal) + '</td>' +
      '</tr>';
  }).join('');
  
  var statusColor = '#666';
  if (invoice.status === 'Paid') statusColor = '#28a745';
  if (invoice.status === 'Overdue') statusColor = '#dc3545';
  if (invoice.status === 'Sent') statusColor = '#007bff';
  
  var html = '<!DOCTYPE html><html><head><style>' +
    '* { box-sizing: border-box; }' +
    'body { font-family: Arial, sans-serif; margin: 0; padding: 15px; background: #f5f5f5; }' +
    '.invoice { background: #fff; max-width: 800px; margin: 0 auto; padding: 30px; border: 1px solid #ddd; }' +
    '.inv-header { display: flex; justify-content: space-between; border-bottom: 3px solid #2E5339; padding-bottom: 15px; margin-bottom: 20px; }' +
    '.company { }' +
    '.company h1 { font-size: 18px; color: #2E5339; margin: 0; letter-spacing: 1px; }' +
    '.company .tag { font-size: 9px; color: #888; font-style: italic; }' +
    '.inv-title { text-align: right; }' +
    '.inv-title h2 { font-size: 28px; color: #2E5339; margin: 0; }' +
    '.inv-title .num { font-size: 16px; color: #4A7C59; }' +
    '.inv-title .status { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: bold; color: #fff; background: ' + statusColor + '; }' +
    '.meta { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 25px; }' +
    '.meta-box { }' +
    '.meta-label { font-size: 9px; text-transform: uppercase; color: #999; margin-bottom: 3px; }' +
    '.meta-value { font-size: 13px; font-weight: bold; color: #333; }' +
    '.meta-addr { font-size: 11px; color: #555; line-height: 1.5; }' +
    'table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }' +
    'th { background: #2E5339; color: #fff; padding: 8px 6px; text-align: left; font-size: 10px; }' +
    'td { padding: 6px; border-bottom: 1px solid #eee; font-size: 10px; }' +
    'tr:nth-child(even) { background: #fafafa; }' +
    '.r { text-align: right; }' +
    '.totals { display: flex; justify-content: flex-end; margin-top: 10px; }' +
    '.totals-box { width: 250px; }' +
    '.total-row { display: flex; justify-content: space-between; padding: 5px 0; font-size: 12px; }' +
    '.total-row.grand { border-top: 2px solid #2E5339; margin-top: 5px; padding-top: 8px; font-size: 16px; font-weight: bold; color: #2E5339; }' +
    '.terms { margin-top: 25px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 10px; color: #888; }' +
    '.no-print { text-align: center; margin-bottom: 15px; }' +
    '.btn { padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-size: 12px; margin: 5px; }' +
    '.btn-print { background: #4A7C59; color: white; }' +
    '.btn-close { background: #6c757d; color: white; }' +
    '@media print { .no-print { display: none; } body { background: white; padding: 0; } .invoice { border: none; box-shadow: none; } }' +
    '</style></head><body>' +
    '<div class="no-print">' +
    '<button class="btn btn-print" onclick="window.print()">🖨️ Print Invoice</button>' +
    '<button class="btn btn-close" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<div class="invoice">' +
    '<div class="inv-header">' +
    '<div class="company"><h1>COMMODITY SAMPLER SERVICES</h1><div class="tag">"Integrity Through Independence"</div></div>' +
    '<div class="inv-title"><h2>INVOICE</h2><div class="num">' + invoice.number + '</div><div class="status">' + invoice.status + '</div></div>' +
    '</div>' +
    '<div class="meta">' +
    '<div class="meta-box">' +
    '<div class="meta-label">Bill To</div>' +
    '<div class="meta-value">' + invoice.customer + '</div>' +
    (custAddress ? '<div class="meta-addr">' + custAddress + '</div>' : '') +
    '</div>' +
    '<div class="meta-box" style="text-align:right;">' +
    '<div><span class="meta-label">Invoice Date: </span><span class="meta-value">' + fmtDate(invoice.date) + '</span></div>' +
    '<div><span class="meta-label">Due Date: </span><span class="meta-value">' + fmtDate(invoice.dueDate) + '</span></div>' +
    '<div><span class="meta-label">Terms: </span><span class="meta-value">' + BILLING_CONFIG.defaultTerms + '</span></div>' +
    '</div>' +
    '</div>' +
    '<table>' +
    '<thead><tr><th>Service</th><th>Sample #</th><th>Container</th><th>Reference</th><th class="r">Qty</th><th class="r">Rate</th><th class="r">Amount</th></tr></thead>' +
    '<tbody>' + rowsHtml + '</tbody>' +
    '</table>' +
    '<div class="totals"><div class="totals-box">' +
    '<div class="total-row"><span>Subtotal</span><span>' + fmtMoney(invoice.subtotal) + '</span></div>' +
    (invoice.shippingFee > 0 ? '<div class="total-row"><span>Shipping Fee (20%)</span><span>' + fmtMoney(invoice.shippingFee) + '</span></div>' : '') +
    '<div class="total-row grand"><span>Total Due</span><span>' + fmtMoney(invoice.total) + '</span></div>' +
    '</div></div>' +
    '<div class="terms">' +
    '<strong>Terms:</strong> ' + BILLING_CONFIG.defaultTerms + '<br>' +
    'Thank you for your business.' +
    '</div>' +
    '</div></body></html>';
  
  return html;
}

function _getCustomerAddress(customerName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contacts');
    if (!sheet || sheet.getLastRow() < 2) return '';
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      var company = String(data[i][1]).trim();
      if (company.toLowerCase() === customerName.toLowerCase() ||
          customerName.toLowerCase().indexOf(company.toLowerCase()) !== -1 ||
          company.toLowerCase().indexOf(customerName.toLowerCase()) !== -1) {
        var parts = [];
        if (data[i][2]) parts.push(String(data[i][2]));
        if (data[i][3]) parts.push(String(data[i][3]));
        var cityLine = '';
        if (data[i][4]) cityLine += String(data[i][4]);
        if (data[i][5]) cityLine += ', ' + String(data[i][5]);
        if (data[i][6]) cityLine += ' ' + String(data[i][6]);
        if (cityLine) parts.push(cityLine);
        return parts.join('<br>');
      }
    }
  } catch(e) {}
  return '';
}


// ============================================================
// QUICKBOOKS ONLINE — CSV EXPORT
// ============================================================

function exportInvoicesQBO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  
  if (!invSheet || !liSheet || invSheet.getLastRow() < 2) {
    ui.alert('No invoices to export.');
    return;
  }
  
  var invCol = _getColumnMap(invSheet);
  var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).getValues();
  
  var unexported = [];
  for (var i = 0; i < invData.length; i++) {
    var exported = invData[i][invCol['QB Exported']];
    var status = String(invData[i][invCol['Status']]).trim();
    if (!exported && status !== 'Void') {
      unexported.push({
        rowIdx: i,
        sheetRow: i + 2,
        number: String(invData[i][invCol['Invoice #']]),
        date: invData[i][invCol['Invoice Date']],
        dueDate: invData[i][invCol['Due Date']],
        customer: String(invData[i][invCol['Customer']]),
        total: parseFloat(invData[i][invCol['Total']]) || 0
      });
    }
  }
  
  if (unexported.length === 0) {
    ui.alert('All invoices have been exported to QuickBooks already.');
    return;
  }
  
  var resp = ui.alert(
    'Export to QuickBooks',
    'Found ' + unexported.length + ' unexported invoice(s).\n\n' +
    'This will create a CSV file that you can import into\n' +
    'QuickBooks Online via Settings → Import Data → Invoices.\n\n' +
    'Export now?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  
  var liData = liSheet.getRange(2, 1, liSheet.getLastRow() - 1, 15).getValues();
  var linesByInvoice = {};
  for (var i = 0; i < liData.length; i++) {
    var invNum = String(liData[i][0]).trim();
    if (!linesByInvoice[invNum]) linesByInvoice[invNum] = [];
    linesByInvoice[invNum].push({
      serviceCode: String(liData[i][2]),
      description: String(liData[i][3]),
      csSample: String(liData[i][4]),
      qty: parseFloat(liData[i][8]) || 1,
      unitPrice: parseFloat(liData[i][10]) || 0,
      lineTotal: parseFloat(liData[i][11]) || 0
    });
  }
  
  var csvLines = [];
  csvLines.push([
    'InvoiceNo', 'Customer', 'InvoiceDate', 'DueDate', 'Terms',
    'ItemDescription', 'ItemQuantity', 'ItemRate', 'ItemAmount',
    'Memo'
  ].map(escapeCSVField).join(','));
  
  unexported.forEach(function(inv) {
    var lines = linesByInvoice[inv.number] || [];
    var dateStr = _fmtDateCSV(inv.date);
    var dueStr = _fmtDateCSV(inv.dueDate);
    
    if (lines.length === 0) {
      csvLines.push([
        inv.number, inv.customer, dateStr, dueStr, BILLING_CONFIG.defaultTerms,
        'Sampling Services', 1, inv.total, inv.total, ''
      ].map(escapeCSVField).join(','));
    } else {
      lines.forEach(function(li, idx) {
        var desc = li.description;
        if (li.csSample) desc += ' [' + li.csSample + ']';
        csvLines.push([
          inv.number, inv.customer, dateStr, dueStr,
          idx === 0 ? BILLING_CONFIG.defaultTerms : '',
          desc, li.qty, li.unitPrice, li.lineTotal, ''
        ].map(escapeCSVField).join(','));
      });
    }
  });
  
  var csvContent = csvLines.join('\n');
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var fileName = 'CSS_QBO_Invoices_' + dateStr + '.csv';
  
  var folder = DriveApp.getRootFolder();
  var file = folder.createFile(fileName, csvContent, MimeType.CSV);
  
  var qbExportedCol = invCol['QB Exported'];
  var qbExportDateCol = invCol['QB Export Date'];
  unexported.forEach(function(inv) {
    invSheet.getRange(inv.sheetRow, qbExportedCol + 1).setValue(true);
    invSheet.getRange(inv.sheetRow, qbExportDateCol + 1).setValue(new Date());
  });
  
  ui.alert(
    '✅ QuickBooks Export Complete!\n\n' +
    'Exported ' + unexported.length + ' invoice(s) to:\n' +
    fileName + '\n\n' +
    'File: ' + file.getUrl() + '\n\n' +
    'To import into QuickBooks Online:\n' +
    '1. Go to Settings (⚙️) → Import Data\n' +
    '2. Select "Invoices"\n' +
    '3. Upload this CSV file\n' +
    '4. Map columns and import'
  );
}


// ============================================================
// QUICKBOOKS DESKTOP — IIF EXPORT
// ============================================================

function exportInvoicesIIF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  
  if (!invSheet || !liSheet || invSheet.getLastRow() < 2) {
    ui.alert('No invoices to export.');
    return;
  }
  
  var invCol = _getColumnMap(invSheet);
  var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).getValues();
  
  var unexported = [];
  for (var i = 0; i < invData.length; i++) {
    var exported = invData[i][invCol['QB Exported']];
    var status = String(invData[i][invCol['Status']]).trim();
    if (!exported && status !== 'Void') {
      unexported.push({
        rowIdx: i,
        sheetRow: i + 2,
        number: String(invData[i][invCol['Invoice #']]),
        date: invData[i][invCol['Invoice Date']],
        dueDate: invData[i][invCol['Due Date']],
        customer: String(invData[i][invCol['Customer']]),
        total: parseFloat(invData[i][invCol['Total']]) || 0
      });
    }
  }
  
  if (unexported.length === 0) {
    ui.alert('All invoices have been exported.');
    return;
  }
  
  var liData = liSheet.getRange(2, 1, liSheet.getLastRow() - 1, 15).getValues();
  var linesByInvoice = {};
  for (var i = 0; i < liData.length; i++) {
    var invNum = String(liData[i][0]).trim();
    if (!linesByInvoice[invNum]) linesByInvoice[invNum] = [];
    linesByInvoice[invNum].push({
      description: String(liData[i][3]),
      csSample: String(liData[i][4]),
      qty: parseFloat(liData[i][8]) || 1,
      unitPrice: parseFloat(liData[i][10]) || 0,
      lineTotal: parseFloat(liData[i][11]) || 0
    });
  }
  
  var iif = '';
  iif += '!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tTERMS\tDUEDATE\n';
  iif += '!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tQNTY\tPRICE\n';
  iif += '!ENDTRNS\n';
  
  unexported.forEach(function(inv) {
    var lines = linesByInvoice[inv.number] || [];
    var dateStr = _fmtDateIIF(inv.date);
    var dueStr = _fmtDateIIF(inv.dueDate);
    
    iif += 'TRNS\tINVOICE\t' + dateStr + '\tAccounts Receivable\t' + inv.customer + '\t\t' +
      inv.total.toFixed(2) + '\t' + inv.number + '\t\t' + BILLING_CONFIG.defaultTerms + '\t' + dueStr + '\n';
    
    lines.forEach(function(li) {
      var desc = li.description;
      if (li.csSample) desc += ' [' + li.csSample + ']';
      iif += 'SPL\tINVOICE\t' + dateStr + '\t' + BILLING_CONFIG.qboIncomeAccount + '\t' + inv.customer +
        '\t\t-' + li.lineTotal.toFixed(2) + '\t' + inv.number + '\t' + desc + '\t' + li.qty + '\t' + li.unitPrice.toFixed(2) + '\n';
    });
    
    iif += 'ENDTRNS\n';
  });
  
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var fileName = 'CSS_QB_Desktop_' + dateStr + '.iif';
  
  var folder = DriveApp.getRootFolder();
  var file = folder.createFile(fileName, iif, MimeType.PLAIN_TEXT);
  
  var qbExportedCol = invCol['QB Exported'];
  var qbExportDateCol = invCol['QB Export Date'];
  unexported.forEach(function(inv) {
    invSheet.getRange(inv.sheetRow, qbExportedCol + 1).setValue(true);
    invSheet.getRange(inv.sheetRow, qbExportDateCol + 1).setValue(new Date());
  });
  
  ui.alert(
    '✅ QuickBooks Desktop Export Complete!\n\n' +
    'Exported ' + unexported.length + ' invoice(s) to:\n' +
    fileName + '\n\n' +
    'File: ' + file.getUrl() + '\n\n' +
    'To import into QuickBooks Desktop:\n' +
    '1. File → Utilities → Import → IIF Files\n' +
    '2. Select the downloaded .iif file\n' +
    '3. Review imported invoices'
  );
}


// ============================================================
// MARK INVOICE AS SENT / PAID / VOID
// ============================================================

function markInvoiceSent() {
  _updateInvoiceStatus('Sent');
}

function markInvoicePaid() {
  _updateInvoiceStatus('Paid');
}

function markInvoiceVoid() {
  _updateInvoiceStatus('Void');
}

function _updateInvoiceStatus(newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== BILLING_CONFIG.invoicesSheetName) {
    ui.alert('Please select a row in the Invoices sheet first.');
    return;
  }
  
  var row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert('Select an invoice row.'); return; }
  
  var col = _getColumnMap(sheet);
  var invNum = String(sheet.getRange(row, col['Invoice #'] + 1).getValue());
  var currentStatus = String(sheet.getRange(row, col['Status'] + 1).getValue());
  
  var resp = ui.alert(
    'Update Invoice Status',
    'Change ' + invNum + ' from "' + currentStatus + '" to "' + newStatus + '"?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  
  sheet.getRange(row, col['Status'] + 1).setValue(newStatus);
  
  if (newStatus === 'Paid') {
    var pmtCol = col['Payment Date'];
    if (pmtCol !== undefined) {
      sheet.getRange(row, pmtCol + 1).setValue(new Date());
    }
  }
  
  ui.alert('✅ ' + invNum + ' marked as ' + newStatus);
}


// ============================================================
// BILLING DASHBOARD
// ============================================================

function showBillingDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  
  if (!invSheet || invSheet.getLastRow() < 2) {
    ui.alert('No invoices yet. Generate invoices first.');
    return;
  }
  
  var col = _getColumnMap(invSheet);
  var data = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).getValues();
  
  var stats = { draft: 0, sent: 0, paid: 0, overdue: 0, void_: 0 };
  var totals = { draft: 0, sent: 0, paid: 0, overdue: 0 };
  var byCustomer = {};
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][col['Status']]).trim();
    var total = parseFloat(data[i][col['Total']]) || 0;
    var customer = String(data[i][col['Customer']]);
    var dueDate = data[i][col['Due Date']];
    
    if (status === 'Sent' && dueDate) {
      var due = new Date(dueDate);
      due.setHours(0, 0, 0, 0);
      if (due < today) status = 'Overdue';
    }
    
    if (status === 'Draft') { stats.draft++; totals.draft += total; }
    else if (status === 'Sent') { stats.sent++; totals.sent += total; }
    else if (status === 'Paid') { stats.paid++; totals.paid += total; }
    else if (status === 'Overdue') { stats.overdue++; totals.overdue += total; }
    else if (status === 'Void') { stats.void_++; }
    
    if (status !== 'Void') {
      if (!byCustomer[customer]) byCustomer[customer] = { invoiced: 0, paid: 0, outstanding: 0 };
      byCustomer[customer].invoiced += total;
      if (status === 'Paid') byCustomer[customer].paid += total;
      else byCustomer[customer].outstanding += total;
    }
  }
  
  var customerNames = Object.keys(byCustomer).sort();

  var custRows = customerNames.map(function(c) {
    var d = byCustomer[c];
    var escapedName = c.replace(/'/g, "\\'").replace(/"/g, '&quot;');
    return '<tr class="cust-row" onclick="showCustomer(\'' + escapedName + '\')">' +
      '<td class="cust-link">' + c + '</td>' +
      '<td class="r">$' + d.invoiced.toFixed(2) + '</td>' +
      '<td class="r">$' + d.paid.toFixed(2) + '</td>' +
      '<td class="r' + (d.outstanding > 0 ? ' ow' : '') + '">$' + d.outstanding.toFixed(2) + '</td></tr>';
  }).join('');

  var html = '<!DOCTYPE html><html><head><style>' +
    '* { box-sizing: border-box; }' +
    'body { font-family: Arial, sans-serif; padding: 15px; background: #f5f5f5; margin: 0; }' +
    '.dash { max-width: 800px; margin: 0 auto; }' +
    'h1 { color: #2E5339; font-size: 18px; border-bottom: 3px solid #4A7C59; padding-bottom: 8px; margin-top: 0; }' +

    /* Cards */
    '.cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 20px; }' +
    '.card { background: #fff; padding: 15px; border-radius: 8px; text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }' +
    '.card .num { font-size: 24px; font-weight: bold; }' +
    '.card .lbl { font-size: 10px; color: #888; text-transform: uppercase; }' +
    '.card .amt { font-size: 13px; color: #2E5339; font-weight: bold; margin-top: 3px; }' +
    '.card.draft .num { color: #6c757d; }' +
    '.card.sent .num { color: #007bff; }' +
    '.card.paid .num { color: #28a745; }' +
    '.card.overdue .num { color: #dc3545; }' +

    /* Customer table */
    'table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }' +
    'th { background: #2E5339; color: #fff; padding: 8px; text-align: left; font-size: 11px; }' +
    'td { padding: 6px 8px; border-bottom: 1px solid #eee; font-size: 11px; }' +
    '.r { text-align: right; }' +
    '.ow { color: #dc3545; font-weight: bold; }' +
    '.cust-row { cursor: pointer; transition: background 0.15s; }' +
    '.cust-row:hover { background: #e8f5e9; }' +
    '.cust-link { color: #1565c0; font-weight: bold; text-decoration: underline; }' +

    /* Drilldown panel */
    '#drilldown { display: none; margin-top: 15px; }' +
    '.dd-back { background: none; border: none; color: #1565c0; cursor: pointer; font-size: 12px; padding: 0; margin-bottom: 10px; }' +
    '.dd-back:hover { text-decoration: underline; }' +
    '.dd-header { background: #2E5339; color: #fff; padding: 12px 15px; border-radius: 8px 8px 0 0; display: flex; justify-content: space-between; align-items: center; }' +
    '.dd-header h3 { margin: 0; font-size: 15px; }' +
    '.dd-total { font-size: 13px; opacity: 0.9; }' +
    '.dd-body { background: #fff; border-radius: 0 0 8px 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }' +
    '.dd-loading { text-align: center; padding: 30px; color: #888; }' +

    /* Invoice cards */
    '.inv-card { border-bottom: 1px solid #eee; }' +
    '.inv-card:last-child { border-bottom: none; }' +
    '.inv-top { padding: 10px 15px; display: flex; justify-content: space-between; align-items: center; cursor: pointer; transition: background 0.15s; }' +
    '.inv-top:hover { background: #f8f8f8; }' +
    '.inv-num { font-weight: bold; color: #1e3c72; font-size: 13px; }' +
    '.inv-date { color: #888; font-size: 10px; margin-left: 10px; }' +
    '.inv-right { display: flex; align-items: center; gap: 10px; }' +
    '.inv-total { font-weight: bold; font-size: 13px; color: #2E5339; }' +
    '.badge { padding: 2px 8px; border-radius: 10px; font-size: 9px; font-weight: bold; color: #fff; }' +
    '.badge-draft { background: #6c757d; }' +
    '.badge-sent { background: #007bff; }' +
    '.badge-paid { background: #28a745; }' +
    '.badge-overdue { background: #dc3545; }' +
    '.inv-arrow { color: #ccc; font-size: 14px; transition: transform 0.2s; }' +
    '.inv-arrow.open { transform: rotate(90deg); }' +

    /* Line items */
    '.li-table { display: none; padding: 0 15px 12px; }' +
    '.li-table.open { display: block; }' +
    '.li-table table { box-shadow: none; margin-bottom: 5px; }' +
    '.li-table th { background: #4A7C59; font-size: 9px; padding: 5px 6px; }' +
    '.li-table td { font-size: 10px; padding: 4px 6px; }' +
    '.li-table .svc-code { font-family: monospace; font-size: 9px; color: #4A7C59; font-weight: bold; background: #e8f5e9; padding: 1px 5px; border-radius: 3px; }' +
    '.li-table .email-link { color: #1565c0; text-decoration: none; font-size: 13px; }' +
    '.li-table .email-link:hover { color: #0d47a1; }' +
    '.li-subtotal { text-align: right; font-size: 11px; color: #666; padding: 4px 6px; }' +
    '.li-subtotal strong { color: #2E5339; }' +

    /* Buttons */
    '.btns { text-align: center; margin-top: 15px; }' +
    '.btn { padding: 8px 16px; border: none; border-radius: 5px; cursor: pointer; font-size: 11px; margin: 3px; }' +
    '.btn-c { background: #6c757d; color: #fff; }' +
    '.btn-c:hover { background: #5a6268; }' +
    '</style></head><body><div class="dash">' +

    '<h1>&#x1F4B0; Billing Dashboard</h1>' +

    /* Status cards */
    '<div class="cards">' +
    '<div class="card draft"><div class="num">' + stats.draft + '</div><div class="lbl">Draft</div><div class="amt">$' + totals.draft.toFixed(2) + '</div></div>' +
    '<div class="card sent"><div class="num">' + stats.sent + '</div><div class="lbl">Sent</div><div class="amt">$' + totals.sent.toFixed(2) + '</div></div>' +
    '<div class="card paid"><div class="num">' + stats.paid + '</div><div class="lbl">Paid</div><div class="amt">$' + totals.paid.toFixed(2) + '</div></div>' +
    '<div class="card overdue"><div class="num">' + stats.overdue + '</div><div class="lbl">Overdue</div><div class="amt">$' + totals.overdue.toFixed(2) + '</div></div>' +
    '</div>' +

    /* Customer table */
    '<div id="custSection">' +
    '<h3 style="color:#2E5339;margin-bottom:8px;">By Customer <span style="font-size:10px;color:#888;font-weight:normal;">(click to view invoices)</span></h3>' +
    '<table><thead><tr><th>Customer</th><th class="r">Invoiced</th><th class="r">Paid</th><th class="r">Outstanding</th></tr></thead>' +
    '<tbody>' + custRows + '</tbody></table>' +
    '</div>' +

    /* Drilldown panel */
    '<div id="drilldown">' +
    '<button class="dd-back" onclick="backToList()">&#x2190; Back to all customers</button>' +
    '<div class="dd-header"><h3 id="ddCustName"></h3><div class="dd-total" id="ddTotal"></div></div>' +
    '<div class="dd-body" id="ddBody"><div class="dd-loading">Loading invoices...</div></div>' +
    '</div>' +

    '<div class="btns">' +
    '<button class="btn btn-c" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '</div>' +

    '<script>' +
    'function showCustomer(name) {' +
    '  document.getElementById("custSection").style.display = "none";' +
    '  document.getElementById("drilldown").style.display = "block";' +
    '  document.getElementById("ddCustName").textContent = name;' +
    '  document.getElementById("ddTotal").textContent = "Loading...";' +
    '  document.getElementById("ddBody").innerHTML = \'<div class="dd-loading">Loading invoices...</div>\';' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data) { renderDrilldown(data); })' +
    '    .withFailureHandler(function(e) { document.getElementById("ddBody").innerHTML = \'<div class="dd-loading" style="color:#dc3545;">Error: \' + e.message + \'</div>\'; })' +
    '    .getCustomerInvoiceDetails(name);' +
    '}' +

    'function backToList() {' +
    '  document.getElementById("drilldown").style.display = "none";' +
    '  document.getElementById("custSection").style.display = "block";' +
    '}' +

    'function renderDrilldown(data) {' +
    '  if (data.error) { document.getElementById("ddBody").innerHTML = \'<div class="dd-loading" style="color:#dc3545;">\' + data.error + \'</div>\'; return; }' +
    '  var invs = data.invoices;' +
    '  var totalOut = 0;' +
    '  for (var i = 0; i < invs.length; i++) { if (invs[i].status !== "Paid") totalOut += invs[i].total; }' +
    '  document.getElementById("ddTotal").textContent = invs.length + " invoice(s) | Outstanding: $" + totalOut.toFixed(2);' +
    '  if (invs.length === 0) { document.getElementById("ddBody").innerHTML = \'<div class="dd-loading">No invoices found</div>\'; return; }' +
    '  var h = "";' +
    '  for (var i = 0; i < invs.length; i++) {' +
    '    var inv = invs[i];' +
    '    var bc = inv.status === "Paid" ? "badge-paid" : inv.status === "Sent" ? "badge-sent" : inv.status === "Overdue" ? "badge-overdue" : "badge-draft";' +
    '    h += \'<div class="inv-card">\';' +
    '    h += \'<div class="inv-top" onclick="toggleInv(\' + i + \')">\';' +
    '    h += \'<div><span class="inv-num">\' + inv.number + \'</span><span class="inv-date">\' + inv.date + \'</span></div>\';' +
    '    h += \'<div class="inv-right"><span class="badge \' + bc + \'">\' + inv.status + \'</span><span class="inv-total">$\' + inv.total.toFixed(2) + \'</span><span class="inv-arrow" id="arrow\' + i + \'">&#x25B6;</span></div>\';' +
    '    h += \'</div>\';' +
    '    h += \'<div class="li-table" id="li\' + i + \'">\';' +
    '    if (inv.lineItems.length > 0) {' +
    '      h += \'<table><thead><tr><th>Service</th><th>Description</th><th>Sample #</th><th>Order #</th><th>Container</th><th>Reference</th><th class="r">Rate</th><th class="r">Amount</th><th>Email</th></tr></thead><tbody>\';' +
    '      for (var j = 0; j < inv.lineItems.length; j++) {' +
    '        var li = inv.lineItems[j];' +
    '        var emailCell = li.emailLink ? \'<a href="\' + li.emailLink + \'" target="_blank" class="email-link" title="Open original email">&#x1F4E7;</a>\' : "";' +
    '        h += "<tr>";' +
    '        h += \'<td><span class="svc-code">\' + li.serviceCode + \'</span></td>\';' +
    '        h += "<td>" + li.description + (li.productDesc ? \'<br><span style="color:#888;font-size:9px;">\' + li.productDesc + "</span>" : "") + "</td>";' +
    '        h += "<td>" + li.csSample + "</td>";' +
    '        h += "<td>" + li.csOrder + "</td>";' +
    '        h += "<td>" + li.container + "</td>";' +
    '        h += "<td>" + li.reference + "</td>";' +
    '        h += \'<td class="r">$\' + li.unitPrice.toFixed(2) + "</td>";' +
    '        h += \'<td class="r"><strong>$\' + li.lineTotal.toFixed(2) + "</strong></td>";' +
    '        h += \'<td style="text-align:center">\' + emailCell + "</td>";' +
    '        h += "</tr>";' +
    '      }' +
    '      h += "</tbody></table>";' +
    '      h += \'<div class="li-subtotal">\' + inv.lineItems.length + " line items | <strong>$" + inv.total.toFixed(2) + "</strong></div>";' +
    '    } else { h += \'<div class="dd-loading">No line items</div>\'; }' +
    '    h += "</div></div>";' +
    '  }' +
    '  document.getElementById("ddBody").innerHTML = h;' +
    '  if (invs.length === 1) toggleInv(0);' +
    '}' +

    'function toggleInv(idx) {' +
    '  var el = document.getElementById("li" + idx);' +
    '  var arrow = document.getElementById("arrow" + idx);' +
    '  if (el.classList.contains("open")) {' +
    '    el.classList.remove("open"); arrow.classList.remove("open");' +
    '  } else {' +
    '    el.classList.add("open"); arrow.classList.add("open");' +
    '  }' +
    '}' +
    '</script></body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(850).setHeight(700);
  ui.showModalDialog(output, '💰 Billing Dashboard');
}


// ============================================================
// CUSTOMER INVOICE DRILLDOWN (called from dashboard)
// ============================================================

function getCustomerInvoiceDetails(customerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  if (!invSheet || !liSheet) return { error: 'Billing sheets not found' };

  // Build email link lookup from All Orders + Completed Orders
  var emailLookup = {};
  var orderSheets = [
    ss.getSheetByName(CONFIG.mainSheetName),
    ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders')
  ];
  for (var si = 0; si < orderSheets.length; si++) {
    var oSheet = orderSheets[si];
    if (!oSheet || oSheet.getLastRow() < 2) continue;
    var oCol = _getColumnMap(oSheet);
    var sampleIdx = oCol['CS Sample #'];
    var emailIdx = oCol['Email Link'];
    var orderIdx = oCol['CS Order #'];
    var senderIdx = oCol['Sender'];
    if (sampleIdx === undefined) continue;
    var oData = oSheet.getRange(2, 1, oSheet.getLastRow() - 1, oSheet.getLastColumn()).getValues();
    for (var oi = 0; oi < oData.length; oi++) {
      var sid = String(oData[oi][sampleIdx] || '').trim();
      if (sid) {
        if (emailIdx !== undefined) emailLookup[sid] = String(oData[oi][emailIdx] || '');
      }
    }
  }

  // Get invoices for this customer
  var invCol = _getColumnMap(invSheet);
  var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).getValues();
  var today = new Date(); today.setHours(0,0,0,0);

  var invoices = [];
  for (var i = 0; i < invData.length; i++) {
    var cust = String(invData[i][invCol['Customer']] || '').trim();
    if (cust !== customerName) continue;
    var status = String(invData[i][invCol['Status']] || '').trim();
    if (status === 'Void') continue;

    // Check overdue
    if (status === 'Sent') {
      var due = invData[i][invCol['Due Date']];
      if (due) {
        var d = new Date(due); d.setHours(0,0,0,0);
        if (d < today) status = 'Overdue';
      }
    }

    var invDate = invData[i][invCol['Invoice Date']];
    var dateStr = '';
    try { dateStr = Utilities.formatDate(new Date(invDate), Session.getScriptTimeZone(), 'MM/dd/yyyy'); } catch(e) {}

    invoices.push({
      number: String(invData[i][invCol['Invoice #']] || ''),
      date: dateStr,
      status: status,
      subtotal: parseFloat(invData[i][invCol['Subtotal']]) || 0,
      shippingFee: parseFloat(invData[i][invCol['Shipping Fee']]) || 0,
      total: parseFloat(invData[i][invCol['Total']]) || 0,
      lineItems: []
    });
  }

  // Get line items grouped by invoice
  var invNums = {};
  for (var j = 0; j < invoices.length; j++) invNums[invoices[j].number.toUpperCase()] = j;

  if (liSheet.getLastRow() >= 2) {
    var liData = liSheet.getRange(2, 1, liSheet.getLastRow() - 1, 15).getValues();
    for (var k = 0; k < liData.length; k++) {
      var liInvNum = String(liData[k][0] || '').trim().toUpperCase();
      var idx = invNums[liInvNum];
      if (idx === undefined) continue;

      var csSample = String(liData[k][4] || '').trim();
      var emailLink = emailLookup[csSample] || '';

      invoices[idx].lineItems.push({
        serviceCode: String(liData[k][2] || ''),
        description: String(liData[k][3] || ''),
        csSample: csSample,
        csOrder: String(liData[k][5] || ''),
        container: String(liData[k][6] || ''),
        reference: String(liData[k][7] || ''),
        qty: parseFloat(liData[k][8]) || 1,
        unit: String(liData[k][9] || ''),
        unitPrice: parseFloat(liData[k][10]) || 0,
        lineTotal: parseFloat(liData[k][11]) || 0,
        warehouse: String(liData[k][13] || ''),
        productDesc: String(liData[k][14] || ''),
        emailLink: emailLink
      });
    }
  }

  return { customer: customerName, invoices: invoices };
}


// ============================================================
// EDIT LINE ITEM SERVICE CODE (override auto-detect)
// ============================================================

function editLineItemService() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== BILLING_CONFIG.lineItemsSheetName) {
    ui.alert('Please select a row in the Invoice Line Items sheet.');
    return;
  }
  
  var row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert('Select a line item row.'); return; }
  
  var rates = _loadRates();
  var codes = Object.keys(rates).sort();
  
  var currentCode = String(sheet.getRange(row, 3).getValue()).trim();
  var optionList = codes.map(function(c) {
    return c + ' — ' + rates[c].description + ' ($' + rates[c].rate.toFixed(2) + ')' + (c === currentCode ? ' ← current' : '');
  }).join('\n');
  
  var resp = ui.prompt(
    'Change Service Code',
    'Current: ' + currentCode + '\n\nAvailable codes:\n' + optionList + '\n\nEnter new service code:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  var newCode = resp.getResponseText().trim().toUpperCase();
  if (!rates[newCode]) { ui.alert('Invalid service code: ' + newCode); return; }
  
  var rate = rates[newCode];
  var qty = parseFloat(sheet.getRange(row, 9).getValue()) || 1;
  
  sheet.getRange(row, 3).setValue(newCode);
  sheet.getRange(row, 4).setValue(rate.description);
  sheet.getRange(row, 10).setValue(rate.unit);
  sheet.getRange(row, 11).setValue(rate.rate);
  sheet.getRange(row, 12).setValue(qty * rate.rate);
  
  var invNum = String(sheet.getRange(row, 1).getValue()).trim();
  _recalcInvoiceTotal(invNum);
  
  ui.alert('✅ Updated to ' + newCode + ' — ' + rate.description + ' ($' + rate.rate.toFixed(2) + ')');
}

function _recalcInvoiceTotal(invoiceNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  if (!liSheet || !invSheet) return;
  
  var liData = liSheet.getRange(2, 1, liSheet.getLastRow() - 1, 12).getValues();
  var subtotal = 0;
  var shippingFee = 0;
  
  for (var i = 0; i < liData.length; i++) {
    if (String(liData[i][0]).trim().toUpperCase() === invoiceNum) {
      var code = String(liData[i][2]).trim();
      var lineTotal = parseFloat(liData[i][11]) || 0;
      if (code === 'SHIP-FEE') {
        shippingFee = lineTotal;
      } else {
        subtotal += lineTotal;
      }
    }
  }
  
  var rates = _loadRates();
  if (rates['SHIP-FEE'] && shippingFee > 0) {
    shippingFee = subtotal * (rates['SHIP-FEE'].rate / 100);
    for (var i = 0; i < liData.length; i++) {
      if (String(liData[i][0]).trim().toUpperCase() === invoiceNum && String(liData[i][2]).trim() === 'SHIP-FEE') {
        liSheet.getRange(i + 2, 11).setValue(shippingFee);
        liSheet.getRange(i + 2, 12).setValue(shippingFee);
        break;
      }
    }
  }
  
  var total = subtotal + shippingFee;
  
  var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).getValues();
  var invCol = _getColumnMap(invSheet);
  for (var i = 0; i < invData.length; i++) {
    if (String(invData[i][invCol['Invoice #']]).trim().toUpperCase() === invoiceNum) {
      invSheet.getRange(i + 2, invCol['Subtotal'] + 1).setValue(subtotal);
      invSheet.getRange(i + 2, invCol['Shipping Fee'] + 1).setValue(shippingFee);
      invSheet.getRange(i + 2, invCol['Total'] + 1).setValue(total);
      break;
    }
  }
}


// ============================================================
// BILLING-SPECIFIC DATE HELPERS
// (escapeCSVField, formatDateForSQL, getTrackingUrl, cleanReceiverName
//  are all defined in Main.gs — do NOT redefine them here)
// ============================================================

function _fmtDateCSV(d) {
  if (!d) return '';
  try {
    var dt = new Date(d);
    return String(dt.getMonth() + 1).padStart(2, '0') + '/' +
      String(dt.getDate()).padStart(2, '0') + '/' + dt.getFullYear();
  } catch(e) { return ''; }
}

function _fmtDateIIF(d) {
  return _fmtDateCSV(d);
}
