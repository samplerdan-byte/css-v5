// ============================================================
// RunTestOrder.gs — Generate test orders for all service types
//
// Run from menu or script editor:
//   runAllTestOrders()     — inserts one order per service type (29 orders)
//   runTestOrdersMinimal() — inserts 5 core orders (sample, photo, inspection, container, cocoa)
//   shipAllTestOrders()    — marks all TEST- orders as Shipped + moves to Completed (triggers billing)
//   clearTestOrders()      — removes all rows with "TEST-" in Comments
// ============================================================

/**
 * Insert test orders covering every billable service on the rate sheet.
 * Each order triggers a different _detectServiceCode() path so you can
 * verify scanning, billing, cover sheets, photos, and reports end-to-end.
 */
function runAllTestOrders() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Insert Test Orders',
    'This will add 29 test orders (one per rate-sheet service) to All Orders.\n\n' +
    'All test rows are tagged "TEST-xxx" in Comments so you can clear them later.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!sheet) { ui.alert('All Orders sheet not found. Run Setup first.'); return; }

  var orders = _buildTestOrders();
  var added = 0;

  for (var i = 0; i < orders.length; i++) {
    try {
      var data = {};
      for (var k in orders[i]) {
        if (k !== '_testCode') data[k] = orders[i][k];
      }
      addDataToSheet(sheet, data, { skipLiveUpdate: true, skipAutoPrint: true });
      added++;
    } catch (e) {
      Logger.log('Test order ' + i + ' (' + orders[i]._testCode + ') failed: ' + e);
    }
  }

  try { updateLiveOrdersView(); } catch (e) {}

  ui.alert('Done! Added ' + added + ' of ' + orders.length + ' test orders.\n\n' +
    'Look for "TEST-" in the Comments column.\n' +
    'Run clearTestOrders() when finished testing.');
}

/**
 * Minimal set: 5 orders covering the main workflows.
 */
function runTestOrdersMinimal() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Insert 5 Core Test Orders',
    'Adds:\n' +
    '1. Regular warehouse sample (REG)\n' +
    '2. Photos with sample (PHOTO-W)\n' +
    '3. Inspection (INSPECT)\n' +
    '4. Container supervision inbound (CONT-IN)\n' +
    '5. FCC cocoa grading (INV-COCOA)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!sheet) { ui.alert('All Orders sheet not found.'); return; }

  var all = _buildTestOrders();
  var keys = ['REG', 'PHOTO-W', 'INSPECT', 'CONT-IN', 'INV-COCOA'];
  var added = 0;

  for (var i = 0; i < all.length; i++) {
    if (keys.indexOf(all[i]._testCode) !== -1) {
      try {
        var data = {};
        for (var k in all[i]) {
          if (k !== '_testCode') data[k] = all[i][k];
        }
        addDataToSheet(sheet, data, { skipLiveUpdate: true, skipAutoPrint: true });
        added++;
      } catch (e) { Logger.log('Mini test ' + i + ' failed: ' + e); }
    }
  }

  try { updateLiveOrdersView(); } catch (e) {}
  ui.alert('Done! Added ' + added + ' test orders.');
}

/**
 * Remove all test orders (rows where Comments contains "TEST-").
 */
function clearTestOrders() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Clear Test Orders',
    'This will DELETE all rows in All Orders where Comments contains "TEST-".\n\n' +
    'This cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!sheet) { ui.alert('Sheet not found.'); return; }

  var col = _getColumnMap(sheet);
  var commentsIdx = col['Comments'];
  if (commentsIdx === undefined) { ui.alert('Comments column not found.'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows.'); return; }

  var data = sheet.getRange(2, commentsIdx + 1, lastRow - 1, 1).getValues();
  var deleted = 0;

  // Delete bottom-up to avoid row-shift issues
  for (var r = data.length - 1; r >= 0; r--) {
    if (String(data[r][0]).indexOf('TEST-') !== -1) {
      sheet.deleteRow(r + 2);
      deleted++;
    }
  }

  // Also clear from Live Orders and Completed Orders
  var viewSheets = ['Live Orders', 'Completed Orders'];
  for (var s = 0; s < viewSheets.length; s++) {
    var vSheet = ss.getSheetByName(viewSheets[s]);
    if (!vSheet || vSheet.getLastRow() < 2) continue;
    var vCol = _getColumnMap(vSheet);
    var vCommentsIdx = vCol['Comments'];
    if (vCommentsIdx === undefined) continue;
    var vData = vSheet.getRange(2, vCommentsIdx + 1, vSheet.getLastRow() - 1, 1).getValues();
    for (var vr = vData.length - 1; vr >= 0; vr--) {
      if (String(vData[vr][0]).indexOf('TEST-') !== -1) {
        vSheet.deleteRow(vr + 2);
      }
    }
  }

  ui.alert('Deleted ' + deleted + ' test order(s) from All Orders.\nAlso cleaned Live Orders and Completed Orders.');
}


/**
 * Ship all test orders: set Status=Shipped, Tracking=TEST123456789,
 * Shipped Date=now. Does NOT move or delete rows — everything stays
 * in All Orders so you can review. Run clearTestOrders() when done.
 */
function shipAllTestOrders() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Ship All Test Orders',
    'This will:\n' +
    '1. Mark all TEST- orders as "Shipped"\n' +
    '2. Set tracking number to TEST123456789\n' +
    '3. Set shipped date to now\n\n' +
    'Orders stay in All Orders for review.\n' +
    'Run clearTestOrders() when you\'re done.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!sheet) { ui.alert('All Orders sheet not found.'); return; }

  var col = _getColumnMap(sheet);
  var commentsIdx = col['Comments'];
  var statusIdx = col['Status'];
  var trackingIdx = col['Tracking Number'];
  var shipDateIdx = col['Shipped Date'];

  if (commentsIdx === undefined || statusIdx === undefined) {
    ui.alert('Required columns not found (Comments, Status).'); return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows.'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var now = new Date();
  var shipped = 0;

  for (var r = 0; r < data.length; r++) {
    var comments = String(data[r][commentsIdx] || '');
    if (comments.indexOf('TEST-') === -1) continue;

    var rowNum = r + 2;

    sheet.getRange(rowNum, statusIdx + 1).setValue('Shipped');

    if (trackingIdx !== undefined) {
      sheet.getRange(rowNum, trackingIdx + 1).setValue('TEST123456789');
    }

    if (shipDateIdx !== undefined) {
      sheet.getRange(rowNum, shipDateIdx + 1).setValue(now);
    }

    shipped++;
  }

  SpreadsheetApp.flush();

  ui.alert(
    'Done!\n\n' +
    '✓ ' + shipped + ' test orders marked as Shipped\n' +
    '✓ Tracking: TEST123456789\n\n' +
    'All rows remain in All Orders for review.\n' +
    'Run clearTestOrders() to remove them when done.'
  );
}


// ============================================================
// TEST ORDER DATA — one per service code
// ============================================================

function _buildTestOrders() {
  var senders = [
    'Louis Dreyfus Company', 'Olam International', 'Volcafe Ltd',
    'Sucafina SA', 'Serengeti Trading Co', 'Atlantic USA Inc',
    'Armenia Coffee Corp', 'Paragon Coffee Trading'
  ];
  var warehouses = [
    'Continental Terminal', 'RPM Avenel', 'Keurig Green Mountain',
    'GreenStar', 'Dupuy Storage'
  ];
  var receivers = [
    'Starbucks Coffee, Attn: QC Lab', 'Nestle USA, Attn: Green Coffee',
    'JDE Peets, Attn: Imports', 'Massimo Zanetti USA, Attn: Quality',
    'S&D Coffee, Attn: Receiving'
  ];
  var origins = [
    'Colombia Supremo', 'Brazil Santos 2/3', 'Ethiopia Yirgacheffe Gr.2',
    'Guatemala Antigua SHB', 'Honduras HG EP', 'Vietnam Robusta Gr.1',
    'Peru HB MCM', 'Costa Rica Tarrazu SHB', 'Kenya AA',
    'Uganda Robusta Screen 15', 'Mexico Altura HG', 'Papua New Guinea A/X'
  ];
  var shippingLines = ['MSC', 'MAERSK', 'HAPAG-LLOYD', 'ZIM', 'CMA-CGM', 'ONE', 'COSCO'];

  function pick(arr) { return arr[Math.floor(Math.random() * arr.length)]; }
  function cont() {
    var p = ['MSCU','MSKU','MRKU','TCLU','CMAU','OOLU'];
    return pick(p) + (1000000 + Math.floor(Math.random() * 8999999));
  }
  function mark() {
    return '0' + (10 + Math.floor(Math.random() * 89)) + '/' +
      (1000 + Math.floor(Math.random() * 8999)) + '/' +
      (100 + Math.floor(Math.random() * 899));
  }
  function cargo() { return 'C-' + (20000 + Math.floor(Math.random() * 79999)); }
  function bol() { return 'MEDU' + (1000000 + Math.floor(Math.random() * 8999999)); }
  function so() { return 'SO-2026-' + (1000 + Math.floor(Math.random() * 8999)); }

  var orders = [];

  // ── 1. REG — Regular Sample (up to 5 lbs) — $28 ──
  orders.push({
    _testCode: 'REG',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: 'PO-' + (40000 + Math.floor(Math.random() * 9999)),
    description: pick(origins), bagCount: '250', weight: '', sampleWeight: '2 LB',
    pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-REG Regular warehouse sample',
    sourceEmail: 'TEST-REG',
    sampleType: 'Warehouse Sample',
    containerETA: '2026-03-10', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'Afloat'
  });

  // ── 2. INT-10 — Intensive 10% — $43 ──
  orders.push({
    _testCode: 'INT-10',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '300', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-INT10 Intensive 10% sample',
    sourceEmail: 'TEST-INT10',
    sampleType: 'Warehouse Sample 10%',
    containerETA: '2026-03-12', shippingLine: pick(shippingLines),
    shippingNotes: 'Intensive sampling requested', bol: bol(), shipStatus: 'Afloat'
  });

  // ── 3. INT-20 — Intensive 20% — $63 ──
  orders.push({
    _testCode: 'INT-20',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '200', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-INT20 Intensive 20% sample',
    sourceEmail: 'TEST-INT20',
    sampleType: 'Warehouse Sample 20%',
    containerETA: '2026-03-14', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'Landed'
  });

  // ── 4. INT-30 — Intensive 30% — $98 ──
  orders.push({
    _testCode: 'INT-30',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '275', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'UPS Next Day Air',
    comments: 'TEST-INT30 Intensive 30% sample',
    sourceEmail: 'TEST-INT30',
    sampleType: 'Warehouse Sample 30%',
    containerETA: '2026-03-08', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 5. INT-50 — Intensive 50%+ — call for quote ──
  orders.push({
    _testCode: 'INT-50',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '150', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-INT50 Intensive 50%+ call for quote',
    sourceEmail: 'TEST-INT50',
    sampleType: 'Warehouse Sample 50%',
    containerETA: '2026-03-20', shippingLine: pick(shippingLines),
    shippingNotes: 'Quote pending', bol: bol(), shipStatus: 'Afloat'
  });

  // ── 6. LRG-6 — Large Sample 6-9 lbs — $38 ──
  orders.push({
    _testCode: 'LRG-6',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '8 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-LRG6 Large sample 8 lbs',
    sourceEmail: 'TEST-LRG6',
    sampleType: 'Warehouse Sample',
    containerETA: '2026-03-15', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'Landed'
  });

  // ── 7. LRG-10 — Large Sample 10-25 lbs — $48 ──
  orders.push({
    _testCode: 'LRG-10',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '300', weight: '',
    sampleWeight: '15 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-LRG10 Large sample 15 lbs',
    sourceEmail: 'TEST-LRG10',
    sampleType: 'Warehouse Sample',
    containerETA: '2026-03-18', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 8. LRG-26 — Large Sample 26+ lbs — $63 ──
  orders.push({
    _testCode: 'LRG-26',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '400', weight: '',
    sampleWeight: '30 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Freight',
    comments: 'TEST-LRG26 Large sample 30 lbs',
    sourceEmail: 'TEST-LRG26',
    sampleType: 'Warehouse Sample',
    containerETA: '2026-03-22', shippingLine: pick(shippingLines),
    shippingNotes: 'Heavy sample — use freight', bol: bol(), shipStatus: 'Discharged'
  });

  // ── 9. KEURIG — Keurig 30% Sampling — $98 ──
  orders.push({
    _testCode: 'KEURIG',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: 'Keurig Green Mountain', sender: pick(senders),
    receiver: 'Keurig Dr Pepper, Attn: QC',
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-KEURIG Keurig 30% sampling',
    sourceEmail: 'TEST-KEURIG',
    sampleType: 'Keurig Sample',
    containerETA: '2026-03-25', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'Afloat'
  });

  // ── 10. BARREL — Barrel Sampling — $100 ──
  orders.push({
    _testCode: 'BARREL',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: 'Jamaica Blue Mountain Barrel', bagCount: '12',
    weight: '', sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-BARREL Barrel sampling',
    sourceEmail: 'TEST-BARREL',
    sampleType: 'Barrel Sampling',
    containerETA: '2026-04-01', shippingLine: pick(shippingLines),
    shippingNotes: '12 barrels', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 11. MICRO — Microlots (<19 bags) — $38 ──
  orders.push({
    _testCode: 'MICRO',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: 'Ethiopia Guji Natural Microlot', bagCount: '10',
    weight: '', sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-MICRO Microlot 10 bags',
    sourceEmail: 'TEST-MICRO',
    sampleType: 'Warehouse Sample',
    containerETA: '2026-03-28', shippingLine: pick(shippingLines),
    shippingNotes: 'Small lot', bol: bol(), shipStatus: 'Landed'
  });

  // ── 12. MOIST — Moisture Reading — $55 ──
  orders.push({
    _testCode: 'MOIST',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-MOIST Moisture reading',
    sourceEmail: 'TEST-MOIST',
    sampleType: 'Moisture Reading',
    containerETA: '2026-03-11', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 13. SCREEN — Screening — $55 ──
  orders.push({
    _testCode: 'SCREEN',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '300', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-SCREEN Screening coffee samples',
    sourceEmail: 'TEST-SCREEN',
    sampleType: 'Screen Analysis',
    containerETA: '2026-03-16', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'Landed'
  });

  // ── 14. INV-COUNT — Inventory Physical Count — $30 ──
  orders.push({
    _testCode: 'INV-COUNT',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: '',
    reference: '', description: pick(origins), bagCount: '500', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-INVCOUNT Inventory physical count',
    sourceEmail: 'TEST-INVCOUNT',
    sampleType: 'Inventory Count',
    containerETA: '', shippingLine: '',
    shippingNotes: 'Count only — no sample shipped', bol: '', shipStatus: ''
  });

  // ── 15. BAG-MOVE — Bag Moves (up to 3) — $95 ──
  orders.push({
    _testCode: 'BAG-MOVE',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '3', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-BAGMOVE Bag move 3 bags',
    sourceEmail: 'TEST-BAGMOVE',
    sampleType: 'Bag Move',
    containerETA: '', shippingLine: '',
    shippingNotes: 'Move 3 bags to new location', bol: '', shipStatus: 'At Warehouse'
  });

  // ── 16. SPLIT-BAG — Splitting Bags — $45 ──
  orders.push({
    _testCode: 'SPLIT-BAG',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '5', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-SPLITBAG Splitting 5 bags',
    sourceEmail: 'TEST-SPLITBAG',
    sampleType: 'Split Bag',
    containerETA: '', shippingLine: '',
    shippingNotes: 'Split into 2 lots', bol: '', shipStatus: 'At Warehouse'
  });

  // ── 17. INSPECT — Inspection — $110 ──
  orders.push({
    _testCode: 'INSPECT',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-INSPECT Quality inspection',
    sourceEmail: 'TEST-INSPECT',
    sampleType: 'Inspection — Quality',
    containerETA: '2026-03-09', shippingLine: pick(shippingLines),
    shippingNotes: 'Buyer reported quality concern', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 18. SIFT — Sifting Inspection — $110 ──
  orders.push({
    _testCode: 'SIFT',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-SIFT Sifting inspection',
    sourceEmail: 'TEST-SIFT',
    sampleType: 'Sift Inspection',
    containerETA: '2026-03-13', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 19. CONT-IN — Container Supervision Inbound — $180 ──
  orders.push({
    _testCode: 'CONT-IN',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '275', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-CONTIN Container supervision inbound',
    sourceEmail: 'TEST-CONTIN',
    sampleType: 'Container Supervision Inbound',
    containerETA: '2026-03-17', shippingLine: pick(shippingLines),
    shippingNotes: 'Supervise stripping', bol: bol(), shipStatus: 'Discharged'
  });

  // ── 20. CONT-OUT — Container Supervision Outbound — $120 ──
  orders.push({
    _testCode: 'CONT-OUT',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-CONTOUT Container supervision outbound',
    sourceEmail: 'TEST-CONTOUT',
    sampleType: 'Container Supervision Outbound',
    containerETA: '', shippingLine: pick(shippingLines),
    shippingNotes: 'Supervise stuffing for export', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 21. WEIGH-IN — Weight Inspection Inbound — $180 ──
  orders.push({
    _testCode: 'WEIGH-IN',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '275',
    weight: '41,250 LB', sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-WEIGHIN Weight inspection inbound',
    sourceEmail: 'TEST-WEIGHIN',
    sampleType: 'Weight Inspection',
    containerETA: '2026-03-19', shippingLine: pick(shippingLines),
    shippingNotes: 'Verify weights at stripping', bol: bol(), shipStatus: 'Discharged'
  });

  // ── 22. REWEIGH — Reweights — $120 ──
  orders.push({
    _testCode: 'REWEIGH',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250',
    weight: '37,500 LB', sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-REWEIGH Reweigh lot',
    sourceEmail: 'TEST-REWEIGH',
    sampleType: 'Reweigh',
    containerETA: '', shippingLine: '',
    shippingNotes: 'Buyer dispute — reweigh entire lot', bol: '', shipStatus: 'At Warehouse'
  });

  // ── 23. PHOTO-W — Photos With Sample — $10 (add-on) ──
  orders.push({
    _testCode: 'PHOTO-W',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-PHOTOW Photos with sample',
    sourceEmail: 'TEST-PHOTOW',
    sampleType: 'Photos With Sample',
    containerETA: '2026-03-21', shippingLine: pick(shippingLines),
    shippingNotes: 'Photo bags + marks + container', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 24. PHOTO-WO — Photos Without Sample — $25 ──
  orders.push({
    _testCode: 'PHOTO-WO',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-PHOTOWO Photos without sample',
    sourceEmail: 'TEST-PHOTOWO',
    sampleType: 'Photos Without Sample',
    containerETA: '2026-03-23', shippingLine: pick(shippingLines),
    shippingNotes: 'Photo documentation only — no sample', bol: bol(), shipStatus: 'At Warehouse'
  });

  // ── 25. RUSH — Same Day Rush Delivery — $150 ──
  orders.push({
    _testCode: 'RUSH',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'Rush Delivery',
    comments: 'TEST-RUSH Same day rush NYC',
    sourceEmail: 'TEST-RUSH',
    sampleType: 'Warehouse Sample',
    containerETA: '', shippingLine: '',
    shippingNotes: 'URGENT — hand deliver to NYC office today', bol: '', shipStatus: 'At Warehouse'
  });

  // ── 26. INV-COCOA — Inventory Cocoa — $55 ──
  orders.push({
    _testCode: 'INV-COCOA',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: 'Ghana Cocoa Beans Grade 1', bagCount: '200',
    weight: '', sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-COCOA Cocoa inventory',
    sourceEmail: 'TEST-COCOA',
    sampleType: 'FCC Grading-Cocoa',
    containerETA: '2026-03-26', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'Afloat'
  });

  // ── 27. INV-COCBUT — Inventory Cocoa Butter — $79.50 ──
  orders.push({
    _testCode: 'INV-COCBUT',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: 'Ivory Coast Cocoa Butter Deodorized',
    bagCount: '100', weight: '', sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders), receiver: pick(receivers),
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-COCBUT Cocoa butter inventory',
    sourceEmail: 'TEST-COCBUT',
    sampleType: 'FCC Grading-Cocoa Butter',
    containerETA: '2026-03-30', shippingLine: pick(shippingLines),
    shippingNotes: '', bol: bol(), shipStatus: 'Afloat'
  });

  // ── 28. VESSEL — Vessel Supervision — $4.49/MT ──
  orders.push({
    _testCode: 'VESSEL',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: '',
    reference: '', description: 'Brazil Santos 2/3 Bulk Vessel', bagCount: '',
    weight: '500 MT', sampleWeight: '', pNumber: '', sNumber: '',
    warehouse: 'Port Newark', sender: pick(senders), receiver: '',
    shippingProcess: '',
    comments: 'TEST-VESSEL Vessel supervision 500 MT',
    sourceEmail: 'TEST-VESSEL',
    sampleType: 'Vessel Supervision',
    containerETA: '2026-04-05', shippingLine: pick(shippingLines),
    shippingNotes: 'Bulk discharge — supervise at pier', bol: bol(), shipStatus: 'Afloat'
  });

  // ── 29. EXCHANGE — Exchange Samples (bills as REG) — $28 ──
  orders.push({
    _testCode: 'EXCHANGE',
    sampleOrderNum: so(), cargo: cargo(), mark: mark(), container: cont(),
    reference: '', description: pick(origins), bagCount: '250', weight: '',
    sampleWeight: '2 LB', pNumber: '', sNumber: '',
    warehouse: pick(warehouses), sender: pick(senders),
    receiver: 'ICE Futures US, Attn: Grading',
    shippingProcess: 'FedEx Overnight',
    comments: 'TEST-EXCHANGE Exchange sample for grading',
    sourceEmail: 'TEST-EXCHANGE',
    sampleType: 'Exchange Samples',
    containerETA: '', shippingLine: '',
    shippingNotes: 'Exchange grading sample', bol: '', shipStatus: 'At Warehouse'
  });

  return orders;
}
