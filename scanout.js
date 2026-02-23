// ============================================================================
// SCANOUT.gs — Shipping Scanner, Barcode Scanning, QR, Labels
// Commodity Sampler Services
// v19: Auto-alternating mode — scan order → tracking → order → tracking
//      Order-grouped tracking (one scan ships all samples in order)
//      Consistent Live Orders refresh after every ship operation
// ============================================================================

var SHEET_NAMES = {
  allOrders: (typeof CONFIG !== 'undefined' && CONFIG.mainSheetName) || 'All Orders',
  liveOrders: (typeof CONFIG !== 'undefined' && CONFIG.liveOrdersSheetName) || 'Live Orders',
  completedOrders: (typeof CONFIG !== 'undefined' && CONFIG.completedOrdersSheetName) || 'Completed Orders'
};

var SAMPLE_COLUMN_NAME = 'CS Sample #';

// ============================================================
// SIDEBAR LAUNCHER
// ============================================================

function showScanOut() {
  var html = HtmlService.createHtmlOutput(getScanOutHTML())
    .setWidth(400)
    .setHeight(700)
    .setTitle('Scan Out - Shipping');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// PREFETCH INDEX
// ============================================================

function getClientIndex() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheet = ss.getSheetByName(SHEET_NAMES.allOrders);
  if (!allSheet) return { error: 'All Orders sheet not found' };

  var lastRow = allSheet.getLastRow();
  var lastCol = allSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { error: 'Sheet empty' };

  // Load only headers + needed columns instead of entire sheet
  var headers = allSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var allData = allSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var sampleCol = findSampleColumn(headers);
  if (sampleCol === -1) return { error: 'CS Sample # column not found in All Orders sheet' };

  var cols = {
    sample: sampleCol,
    order: headers.indexOf('CS Order #'),
    status: headers.indexOf('Status'),
    tracking: headers.indexOf('Tracking Number'),
    shipDate: headers.indexOf('Shipped Date'),
    receiver: headers.indexOf('Receiver'),
    desc: headers.indexOf('Description')
  };

  var samples = {};
  var orders = {};
  var receiverSet = {};
  var today = new Date(); today.setHours(0,0,0,0);
  var shippedToday = 0;

  for (var i = 1; i < allData.length; i++) {
    var id = String(allData[i][sampleCol]).trim();
    if (!id) continue;

    var status = cols.status >= 0 ? String(allData[i][cols.status]).trim() : '';
    var receiver = cols.receiver >= 0 ? String(allData[i][cols.receiver]).trim() : '';
    var desc = cols.desc >= 0 ? String(allData[i][cols.desc]).trim() : '';
    var details = receiver;
    if (desc) details += (details ? ' | ' : '') + desc;

    if (status === 'Shipped' && cols.shipDate >= 0 && allData[i][cols.shipDate]) {
      try {
        var d = new Date(allData[i][cols.shipDate]); d.setHours(0,0,0,0);
        if (d.getTime() === today.getTime()) shippedToday++;
      } catch(e) {}
    }

    var trackInfo = '';
    if ((status === 'Shipped' || status === 'Completed') && cols.tracking >= 0 && allData[i][cols.tracking]) {
      trackInfo = String(allData[i][cols.tracking]);
    }
    var shipDateStr = '';
    if ((status === 'Shipped' || status === 'Completed') && cols.shipDate >= 0 && allData[i][cols.shipDate]) {
      try { shipDateStr = Utilities.formatDate(new Date(allData[i][cols.shipDate]), Session.getScriptTimeZone(), 'MM/dd'); } catch(e) {}
    }

    var orderId = cols.order >= 0 ? String(allData[i][cols.order]).trim() : '';

    samples[id] = {
      r: i + 1, d: details, s: status, recv: receiver, t: trackInfo, sd: shipDateStr, o: orderId
    };

    if (orderId) {
      if (!orders[orderId]) orders[orderId] = [];
      orders[orderId].push(id);
    }

    if (receiver) receiverSet[receiver] = true;
  }

  var carriers = {};
  try {
    var receiverNames = Object.keys(receiverSet);
    for (var j = 0; j < receiverNames.length; j++) {
      try {
        var c = lookupCarrierByCompany(receiverNames[j]);
        if (c && (c.fedex || c.ups)) carriers[receiverNames[j]] = c;
      } catch(e) {}
    }
  } catch(e) {}

  var liveSet = {};
  var liveSheet = ss.getSheetByName(SHEET_NAMES.liveOrders);
  if (liveSheet) {
    var liveData = liveSheet.getDataRange().getValues();
    var liveCol = findSampleColumn(liveData[0]);
    if (liveCol >= 0) {
      for (var k = 1; k < liveData.length; k++) {
        var lid = String(liveData[k][liveCol]).trim();
        if (lid) liveSet[lid] = true;
      }
    }
  }

  return {
    samples: samples, orders: orders, cols: cols, colCount: headers.length,
    carriers: carriers, liveSet: liveSet, shippedToday: shippedToday
  };
}

// ============================================================
// FIND ROW BY SAMPLE # (safe — never uses stale row numbers)
// ============================================================

function _findRowBySample(sheet, sampleId) {
  var col = _getColumnMap(sheet);
  var sampleIdx = col['CS Sample #'];
  if (sampleIdx === undefined) return -1;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  
  var samples = sheet.getRange(2, sampleIdx + 1, lastRow - 1, 1).getValues();
  
  for (var i = 0; i < samples.length; i++) {
    if (String(samples[i][0]).trim() === String(sampleId).trim()) {
      return i + 2;
    }
  }
  return -1;
}

// ============================================================
// SINGLE SHIP (by sample ID, not row)
// ============================================================

function scanOutShip(row, cols, trackingNum) {
  // row param kept for backward compat but we don't trust it
  // We need the sample ID — get it from the sheet at that row as fallback
  try {
    trackingNum = String(trackingNum).trim();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.allOrders);
    if (!sheet) return { success: false, message: 'Sheet not found' };
    
    row = parseInt(row, 10);
    if (!row || row < 2) return { success: false, message: 'Invalid row' };
    
    // Read the sample ID at that row to verify
    var col = _getColumnMap(sheet);
    var sampleIdx = col['CS Sample #'];
    if (sampleIdx === undefined) return { success: false, message: 'CS Sample # column not found' };
    var sampleId = String(sheet.getRange(row, sampleIdx + 1).getValue()).trim();

    if (!sampleId) return { success: false, message: 'No sample at row ' + row };
    
    // Now find the ACTUAL current row for this sample
    var actualRow = _findRowBySample(sheet, sampleId);
    if (actualRow < 2) return { success: false, message: sampleId + ' not found in sheet' };
    
    if (cols.status >= 0)   sheet.getRange(actualRow, cols.status + 1).setValue('Shipped');
    if (cols.shipDate >= 0) sheet.getRange(actualRow, cols.shipDate + 1).setValue(new Date());
    
    if (cols.tracking >= 0) {
      var link = getTrackingUrl(trackingNum);
      if (link) {
        sheet.getRange(actualRow, cols.tracking + 1).setFormula('=HYPERLINK("' + link + '","' + trackingNum + '")');
      } else {
        sheet.getRange(actualRow, cols.tracking + 1).setValue(trackingNum);
      }
    }
    
    SpreadsheetApp.flush();
    try { updateLiveOrdersView(); } catch(ignore) {}
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// BATCH SHIP (by sample ID, not row)
// ============================================================

function scanOutShipBatch(items, cols) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.allOrders);
    if (!sheet) return { success: false, message: 'Sheet not found' };
    
    var col = _getColumnMap(sheet);
    var sampleIdx = col['CS Sample #'];
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2 || sampleIdx === undefined) {
      return { success: false, message: 'No data or sample column missing' };
    }
    
    // Build sample-to-row lookup from CURRENT sheet state
    var sampleData = sheet.getRange(2, sampleIdx + 1, lastRow - 1, 1).getValues();
    var sampleToRow = {};
    for (var s = 0; s < sampleData.length; s++) {
      var sid = String(sampleData[s][0]).trim();
      if (sid) sampleToRow[sid] = s + 2;
    }
    
    var now = new Date();
    var results = [];

    // Collect all updates, then batch write once (instead of per-item setValue calls)
    var statusUpdates = [];
    var shipDateUpdates = [];
    var trackingUpdates = [];

    for (var i = 0; i < items.length; i++) {
      try {
        // Guard against null/undefined id or tracking coming from client
        if (!items[i] || items[i].id == null || items[i].id === '') {
          results.push({ id: '(unknown)', success: false, message: 'Missing sample ID at index ' + i });
          continue;
        }
        if (items[i].tracking == null || items[i].tracking === '') {
          results.push({ id: items[i].id, success: false, message: 'Missing tracking number for ' + items[i].id });
          continue;
        }
        var sampleId = String(items[i].id).trim();
        var trackingNum = String(items[i].tracking).trim();
        if (!sampleId) {
          results.push({ id: items[i].id, success: false, message: 'Blank sample ID at index ' + i });
          continue;
        }
        if (!trackingNum) {
          results.push({ id: items[i].id, success: false, message: 'Blank tracking number for ' + items[i].id });
          continue;
        }

        var actualRow = sampleToRow[sampleId];
        if (!actualRow) {
          results.push({ id: items[i].id, success: false, message: sampleId + ' not found' });
          continue;
        }

        if (cols.status >= 0)   statusUpdates.push({ row: actualRow, col: cols.status + 1, value: 'Shipped' });
        if (cols.shipDate >= 0) shipDateUpdates.push({ row: actualRow, col: cols.shipDate + 1, value: now });

        if (cols.tracking >= 0) {
          var link = getTrackingUrl(trackingNum);
          if (link) {
            trackingUpdates.push({ row: actualRow, col: cols.tracking + 1, formula: '=HYPERLINK("' + link + '","' + trackingNum + '")' });
          } else {
            trackingUpdates.push({ row: actualRow, col: cols.tracking + 1, value: trackingNum });
          }
        }

        results.push({ id: items[i].id, success: true });
      } catch(e) {
        results.push({ id: items[i].id, success: false, message: e.message });
      }
    }

    // Batch write — use RangeList for same-value columns (1 API call instead of N)
    if (statusUpdates.length > 0) {
      var statusRanges = statusUpdates.map(function(u) { return sheet.getRange(u.row, u.col).getA1Notation(); });
      sheet.getRangeList(statusRanges).setValue('Shipped');
    }
    if (shipDateUpdates.length > 0) {
      var dateRanges = shipDateUpdates.map(function(u) { return sheet.getRange(u.row, u.col).getA1Notation(); });
      sheet.getRangeList(dateRanges).setValue(now);
    }
    // Tracking needs individual writes (formulas differ per row)
    trackingUpdates.forEach(function(u) {
      if (u.formula) sheet.getRange(u.row, u.col).setFormula(u.formula);
      else sheet.getRange(u.row, u.col).setValue(u.value);
    });

    SpreadsheetApp.flush();
    try { updateLiveOrdersView(); } catch(ignore) {}
    return { success: true, results: results, shipped: results.filter(function(r) { return r.success; }).length };
  } catch(e) {
    return { success: false, message: 'Batch error: ' + e.message };
  }
}

// ============================================================
// FORCE MOVE
// ============================================================

function forceMoveSample(sampleId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: 'System busy — try again' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var liveSheet = ss.getSheetByName(SHEET_NAMES.liveOrders);
    var completedSheet = ss.getSheetByName(SHEET_NAMES.completedOrders);
    if (!liveSheet || !completedSheet) return { success: false, message: 'Sheets not found' };

    var liveData = liveSheet.getDataRange().getValues();
    var sampleCol = findSampleColumn(liveData[0]);
    if (sampleCol === -1) return { success: false, message: 'Sample column not found' };

    sampleId = String(sampleId).trim();

    for (var i = liveData.length - 1; i >= 1; i--) {
      if (String(liveData[i][sampleCol]).trim() === sampleId) {
        var rowNum = i + 1;
        var rowData = liveSheet.getRange(rowNum, 1, 1, liveSheet.getLastColumn()).getValues()[0];
        completedSheet.appendRow(rowData);
        liveSheet.deleteRow(rowNum);
        return { success: true, message: sampleId + ' moved to Completed Orders' };
      }
    }
    return { success: false, message: sampleId + ' not found in Live Orders' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// scanSampleBarcode() — moved to WebApp.js to avoid duplicates

function openBarcodeScanner() {
  const html = HtmlService.createHtmlOutputFromFile('BarcodeScanner').setWidth(750).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '☕ Barcode Scanner');
}

function openWebScanner() {
  const html = HtmlService.createHtmlOutputFromFile('WebAppScanner').setWidth(700).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📱 QR/Barcode Scanner');
}

// ============================================================
// HELPERS
// ============================================================

function findSampleColumn(headers) {
  if (!headers) return -1;
  var idx = headers.indexOf(SAMPLE_COLUMN_NAME);
  if (idx >= 0) return idx;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).indexOf('CS Sample') >= 0) return i;
  }
  return -1;
}

function getTrackingUrl(tracking) {
  if (!tracking) return null;
  tracking = String(tracking).trim().toUpperCase();
  if (/^1Z[A-Z0-9]{16,}$/i.test(tracking)) return 'https://www.ups.com/track?tracknum=' + tracking;
  if (/^(94|93|92|91|90)\d{18,22}$/.test(tracking)) return 'https://tools.usps.com/go/TrackConfirmAction?tLabels=' + tracking;
  if (/^\d{12,22}$/.test(tracking)) return 'https://www.fedex.com/fedextrack/?trknbr=' + tracking;
  if (/^\d{10,}$/.test(tracking)) return 'https://www.fedex.com/fedextrack/?trknbr=' + tracking;
  return null;
}

// buildQRDataString() — moved to QRcode.js to avoid duplicates

function autoPrintLabel(rowData, col) {
  Logger.log('Auto-print requested for sample: ' + rowData[col['CS Sample #']]);
}

// ============================================================
// SCAN OUT SIDEBAR HTML (v19 — Auto-alternating mode)
// ============================================================

function getScanOutHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; padding: 10px; background: #f5f5f5; position: relative; }
    h2 { text-align: center; color: #2c5f2d; margin-bottom: 6px; font-size: 18px; }
    
    /* MODE TABS */
    .mode-tabs { display: flex; margin-bottom: 8px; border-radius: 8px; overflow: hidden; border: 2px solid #2c5f2d; }
    .mode-tab {
      flex: 1; padding: 10px 0; text-align: center; font-size: 13px; font-weight: bold;
      cursor: pointer; border: none; background: #fff; color: #2c5f2d; transition: all 0.15s;
    }
    .mode-tab.active { background: #2c5f2d; color: #fff; }
    .mode-tab:not(.active):hover { background: #e8f5e9; }
    
    /* INPUT */
    .input-row { display: flex; gap: 6px; margin-bottom: 8px; }
    .input-row input {
      flex: 1; padding: 14px; font-size: 18px; text-align: center;
      border: 3px solid #2c5f2d; border-radius: 10px; background: #fff;
    }
    .input-row input:focus { outline: none; border-color: #28a745; box-shadow: 0 0 0 3px rgba(40,167,69,0.3); }
    .input-row input.tracking-mode { border-color: #1976d2 !important; background: #e3f2fd !important; }
    .go-btn {
      padding: 14px 18px; font-size: 18px; font-weight: bold;
      background: #28a745; color: #fff; border: none; border-radius: 10px; cursor: pointer; min-width: 60px;
    }
    
    /* STATUS BAR */
    #status {
      padding: 10px; border-radius: 8px; text-align: center;
      font-size: 14px; font-weight: bold; margin-bottom: 8px; min-height: 38px;
    }
    .ready { background: #e8f5e9; color: #2c5f2d; }
    .loading { background: #e3f2fd; color: #1565c0; }
    .scanning { background: #fff3cd; color: #856404; }
    .success { background: #d4edda; color: #155724; }
    .error { background: #f8d7da; color: #721c24; }
    .tracking-ready { background: #e3f2fd; color: #1565c0; }
    .wrong-mode { background: #fff3cd; color: #856404; border: 2px solid #ffc107; }
    
    /* QUEUE */
    #queueSection { background: #fff; border-radius: 8px; border: 1px solid #ddd; margin-bottom: 8px; max-height: 320px; overflow-y: auto; }
    #queueHeader {
      display: flex; justify-content: space-between; align-items: center;
      padding: 8px 10px; border-bottom: 1px solid #eee; position: sticky; top: 0; background: #fff; z-index: 1;
    }
    #queueHeader h3 { font-size: 12px; color: #333; margin: 0; }
    #queueCount { font-size: 11px; color: #666; }
    .q-empty { color: #999; font-style: italic; text-align: center; padding: 20px; font-size: 12px; }
    .q-item {
      display: flex; align-items: center; padding: 6px 10px; border-bottom: 1px solid #f5f5f5;
      font-size: 11px; transition: background 0.15s;
    }
    .q-item.q-active { background: #e3f2fd; border-left: 3px solid #1976d2; }
    .q-item.q-done { background: #f0faf0; opacity: 0.7; }
    .q-item.q-sending { background: #fff8e1; }
    .q-item.q-error { background: #ffeaea; }
    .q-num { width: 20px; font-weight: bold; color: #999; font-size: 10px; }
    .q-sample { font-weight: bold; color: #1e3c72; min-width: 90px; }
    .q-details { flex: 1; color: #666; font-size: 9px; margin-left: 6px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .q-tracking { font-family: monospace; font-size: 9px; color: #1976d2; margin-left: 4px; }
    .q-status {
      margin-left: 4px; padding: 1px 6px; border-radius: 8px; font-size: 9px; font-weight: bold; white-space: nowrap;
    }
    .q-status.pending { background: #e0e0e0; color: #666; }
    .q-status.ready { background: #bbdefb; color: #1565c0; }
    .q-status.sending { background: #fff3cd; color: #856404; }
    .q-status.done { background: #c8e6c9; color: #2e7d32; }
    .q-status.fail { background: #ffcdd2; color: #c62828; }
    .q-remove { background: none; border: none; color: #ccc; cursor: pointer; font-size: 14px; padding: 0 2px; margin-left: 4px; }
    .q-remove:hover { color: #f44336; }
    
    /* ACTION BUTTONS */
    .action-row { display: flex; gap: 6px; margin-bottom: 8px; }
    .action-btn {
      flex: 1; padding: 10px; font-size: 12px; font-weight: bold;
      border: none; border-radius: 8px; cursor: pointer;
    }
    .action-btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .btn-clear { background: #f5f5f5; color: #666; border: 1px solid #ddd; }
    .btn-clear:hover:not(:disabled) { background: #eee; }
    .btn-ship { background: #2c5f2d; color: #fff; }
    .btn-ship:hover:not(:disabled) { background: #1e4620; }
    
    /* COUNTER */
    #counter {
      text-align: center; padding: 6px; background: #2c5f2d;
      color: #fff; border-radius: 8px; font-size: 12px; margin-bottom: 8px;
    }
    #counter span { font-size: 16px; font-weight: bold; }
    
    /* CARRIER INFO */
    #carrierInfo {
      display: none; background: #fff; border-radius: 8px; padding: 8px;
      margin-bottom: 8px; border: 2px solid #1976d2;
    }
    #carrierInfo h4 { font-size: 10px; color: #1565c0; margin-bottom: 4px; text-align: center; }
    .carrier-row { display: flex; justify-content: space-between; align-items: center; padding: 2px 0; }
    .carrier-label { font-size: 9px; color: #666; text-transform: uppercase; font-weight: bold; }
    .carrier-acct { font-size: 14px; font-weight: bold; font-family: monospace; color: #1e3c72; letter-spacing: 1px; }
    .carrier-preferred { background: #1976d2; color: #fff; font-size: 8px; padding: 1px 5px; border-radius: 8px; font-weight: bold; }
    
    /* POPUP */
    #popupOverlay {
      display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
      background: rgba(0,0,0,0.7); z-index: 1000; justify-content: center; align-items: center;
    }
    #popupBox {
      background: #fff; border-radius: 12px; padding: 20px; text-align: center;
      max-width: 300px; box-shadow: 0 8px 30px rgba(0,0,0,0.3);
    }
    #popupIcon { font-size: 50px; margin-bottom: 10px; }
    #popupTitle { font-size: 18px; font-weight: bold; margin-bottom: 8px; }
    #popupMessage { font-size: 12px; color: #666; margin-bottom: 15px; line-height: 1.4; white-space: pre-line; }
    .popup-buttons { display: flex; gap: 8px; justify-content: center; flex-wrap: wrap; }
    .popup-btn { padding: 10px 20px; font-size: 13px; font-weight: bold; border: none; border-radius: 20px; cursor: pointer; }
    .popup-btn-ok { background: #dc3545; color: #fff; }
    .popup-btn-secondary { background: #6c757d; color: #fff; }
    .popup-btn-move { background: #17a2b8; color: #fff; }
    .popup-btn-warn { background: #ffc107; color: #333; }
    
    /* MISC */
    #refreshBtn {
      position: absolute; top: 8px; right: 10px; background: none; border: none;
      font-size: 16px; cursor: pointer; color: #999; padding: 2px 6px;
    }
    #refreshBtn:hover { color: #333; }
    #moveBtn {
      width: 100%; padding: 8px; background: #17a2b8; color: #fff;
      border: none; border-radius: 5px; font-size: 12px; font-weight: bold;
      cursor: pointer; margin-bottom: 4px;
    }
    #moveBtn:hover { background: #138496; }
    #moveStatus { font-size: 10px; color: #888; text-align: center; min-height: 14px; margin-bottom: 6px; }
    #debugLog { font-size: 9px; color: #aaa; text-align: center; margin-top: 4px; min-height: 12px; }
  </style>
</head>
<body>

  <button id="refreshBtn" onclick="loadIndex()" title="Refresh data">🔄</button>

  <div id="popupOverlay" onclick="closePopup()">
    <div id="popupBox" onclick="event.stopPropagation()">
      <div id="popupIcon">❌</div>
      <div id="popupTitle">Error</div>
      <div id="popupMessage">Something went wrong</div>
      <div class="popup-buttons" id="popupButtons">
        <button class="popup-btn popup-btn-ok" onclick="closePopup()">OK</button>
      </div>
    </div>
  </div>

  <h2>📤 SCAN OUT</h2>

  <!-- MODE TABS -->
  <div class="mode-tabs">
    <button class="mode-tab active" id="tabSamples" onclick="switchMode(0)">1. Scan Samples</button>
    <button class="mode-tab" id="tabTracking" onclick="switchMode(1)">2. Scan Tracking</button>
  </div>

  <form onsubmit="submitScan(); return false;" autocomplete="off">
    <div class="input-row">
      <input type="text" id="scanBox" placeholder="Loading..." disabled
             autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
      <button type="submit" class="go-btn" id="goBtn" disabled>GO</button>
    </div>
  </form>

  <div id="status" class="loading">Loading order data...</div>

  <!-- CARRIER (shows for active item in tracking mode) -->
  <div id="carrierInfo">
    <h4>📦 Shipping Account</h4>
    <div id="carrierContent"></div>
  </div>

  <!-- QUEUE -->
  <div id="queueSection">
    <div id="queueHeader">
      <h3>📋 Queue</h3>
      <span id="queueCount">0 samples</span>
    </div>
    <div id="queueList">
      <div class="q-empty">Scan sample or order barcodes to build your queue</div>
    </div>
  </div>

  <!-- ACTIONS -->
  <div class="action-row">
    <button class="action-btn btn-clear" id="clearBtn" onclick="clearQueue()" disabled>✕ Clear</button>
    <button class="action-btn btn-ship" id="shipAllBtn" onclick="shipAll()" disabled>🚀 Ship All</button>
  </div>

  <div id="counter">Shipped Today: <span id="count">0</span></div>

  <button id="moveBtn" onclick="runBatchMove()">📦 Move Shipped → Completed</button>
  <div id="moveStatus"></div>

  <div id="debugLog"></div>

<script>
/* ============================================================
   STATE
   ============================================================ */
var IDX = null;
var mode = 0;           /* 0 = scanning samples, 1 = scanning tracking */
var queue = [];         /* [{ id, entry, tracking, status }] */
var trackingIdx = 0;    /* next item in queue awaiting tracking */
var shippedCount = 0;
var inflight = 0;       /* server calls in progress */

function dbg(msg) { document.getElementById("debugLog").textContent = msg; }

/* ============================================================
   CROSS-MODE VALIDATION
   Prevents tracking numbers from being scanned as samples
   and sample/order barcodes from being scanned as tracking.
   ============================================================ */
function looksLikeTracking(value) {
  if (!value) return false;
  var v = value.trim().toUpperCase();
  /* UPS: starts with 1Z */
  if (/^1Z[A-Z0-9]{16,}$/.test(v)) return true;
  /* USPS: starts with 90-94 + 18-22 digits */
  if (/^(94|93|92|91|90)\\d{18,22}$/.test(v)) return true;
  /* FedEx: 12-22 pure digits */
  if (/^\\d{12,22}$/.test(v)) return true;
  return false;
}

function looksLikeSampleOrOrder(value) {
  if (!IDX) return false;
  var v = value.replace(/[\\[\\]]/g, "").trim();
  if (IDX.samples[v]) return true;
  if (IDX.orders[v]) return true;
  return false;
}

/* ============================================================
   LOAD INDEX
   ============================================================ */
function loadIndex() {
  setStatus("loading", "Loading order data...");
  document.getElementById("scanBox").disabled = true;
  document.getElementById("goBtn").disabled = true;
  var t0 = performance.now();
  google.script.run
    .withSuccessHandler(function(data) {
      if (data.error) { setStatus("error", data.error); return; }
      IDX = data;
      shippedCount = data.shippedToday || 0;
      document.getElementById("count").textContent = shippedCount;
      var n = Object.keys(data.samples).length;
      var nOrders = Object.keys(data.orders || {}).length;
      setStatus("ready", "Ready — scan sample or order barcode");
      document.getElementById("scanBox").disabled = false;
      document.getElementById("scanBox").placeholder = "Scan sample or order barcode...";
      document.getElementById("goBtn").disabled = false;
      document.getElementById("scanBox").focus();
      dbg("v19 | " + n + " samples, " + nOrders + " orders | " + Math.round(performance.now()-t0) + "ms");
    })
    .withFailureHandler(function(e) { setStatus("error", "Load failed: " + e.message); })
    .getClientIndex();
}

/* ============================================================
   MODE SWITCHING
   ============================================================ */
function switchMode(m) {
  mode = m;
  document.getElementById("tabSamples").className = "mode-tab" + (m===0 ? " active" : "");
  document.getElementById("tabTracking").className = "mode-tab" + (m===1 ? " active" : "");
  var sb = document.getElementById("scanBox");
  if (m === 0) {
    sb.placeholder = "Scan sample or order barcode...";
    sb.classList.remove("tracking-mode");
    if (queue.length === 0) {
      setStatus("ready", "Scan samples or orders into queue");
    }
    document.getElementById("carrierInfo").style.display = "none";
  } else {
    sb.placeholder = "Scan TRACKING barcode...";
    sb.classList.add("tracking-mode");
    if (queue.length === 0) {
      setStatus("error", "No samples in queue — scan samples first");
    } else {
      advanceTracking();
    }
  }
  sb.value = "";
  sb.focus();
}

function advanceTracking() {
  /* Find next item that needs tracking */
  while (trackingIdx < queue.length && queue[trackingIdx].status !== "pending") trackingIdx++;
  if (trackingIdx >= queue.length) {
    setStatus("success", "✓ All " + queue.length + " samples have tracking!");
    document.getElementById("carrierInfo").style.display = "none";
    renderQueue();
    return;
  }
  var item = queue[trackingIdx];
  setStatus("tracking-ready", "Scan tracking for: " + item.id);
  showCarrierInfo(item.entry.recv ? IDX.carriers[item.entry.recv] : null);
  renderQueue();
}

/* ============================================================
   SUBMIT SCAN
   ============================================================ */
function submitScan() {
  var sb = document.getElementById("scanBox");
  var value = sb.value.trim();
  sb.value = "";
  if (!value || value.length < 2) { sb.focus(); return; }
  // Reject pathologically long inputs (scanner glitch / paste of garbage)
  if (value.length > 256) {
    setStatus("error", "⚠ Scan too long (" + value.length + " chars) — try again");
    sb.focus();
    return;
  }

  if (mode === 0) {
    /* GUARD: reject tracking numbers in sample mode */
    if (looksLikeTracking(value)) {
      setStatus("wrong-mode", "⚠ That looks like a tracking number! Switch to Tracking tab.");
      sb.focus();
      return;
    }
    addSampleToQueue(value);
  } else {
    /* GUARD: reject sample/order barcodes in tracking mode */
    if (looksLikeSampleOrOrder(value)) {
      setStatus("wrong-mode", "⚠ That\\'s a sample/order barcode! Switch to Samples tab.");
      sb.focus();
      return;
    }
    assignTracking(value);
  }
  sb.focus();
}

/* ============================================================
   MODE 0: ADD SAMPLE TO QUEUE (instant)
   ============================================================ */
function addSampleToQueue(id) {
  if (!IDX) { showPopup("error", "Not Ready", "Still loading..."); return; }
  id = id.replace(/[\\[\\]]/g, "").trim();

  /* Check if this is an order number */
  if (IDX.orders[id]) {
    addOrderToQueue(id);
    return;
  }

  /* Single sample */
  addOneSample(id);
}

function addOrderToQueue(orderId) {
  var sampleIds = IDX.orders[orderId];
  var added = 0;
  var skipped = 0;
  var alreadyShipped = 0;

  for (var i = 0; i < sampleIds.length; i++) {
    var sid = sampleIds[i];
    var entry = IDX.samples[sid];
    if (!entry) continue;

    /* Skip already shipped */
    if (entry.s === "Shipped" || entry.s === "Completed") { alreadyShipped++; continue; }

    /* Skip duplicates already in queue */
    var dup = false;
    for (var j = 0; j < queue.length; j++) {
      if (queue[j].id === sid) { dup = true; skipped++; break; }
    }
    if (dup) continue;

    queue.push({ id: sid, entry: entry, tracking: "", status: "pending" });
    added++;
  }

  var msg = "Order " + orderId + ": " + added + " sample" + (added !== 1 ? "s" : "") + " added";
  if (skipped > 0) msg += ", " + skipped + " already queued";
  if (alreadyShipped > 0) msg += ", " + alreadyShipped + " already shipped";

  if (added > 0) {
    setStatus("success", "✓ " + msg);
    updateButtons();
    renderQueue();
    switchMode(1);
  } else {
    setStatus("error", "⚠ " + msg);
    updateButtons();
    renderQueue();
  }
}

function addOneSample(id) {

  /* Duplicate check */
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].id === id) {
      setStatus("error", "⚠ " + id + " already in queue (#" + (i+1) + ")");
      return;
    }
  }

  var entry = IDX.samples[id];
  if (!entry) {
    showPopup("error", "Not Found", id + " not in system.");
    return;
  }

  if (entry.s === "Shipped" || entry.s === "Completed") {
    var inLive = !!IDX.liveSet[id];
    if (inLive) {
      showPopup("stuck", "Stuck Sample!",
        id + " is marked shipped but still in Live Orders.\\n\\n" + (entry.t ? "Has tracking: " + entry.t : "No tracking on file"),
        '<button class="popup-btn popup-btn-move" onclick="forceMove(\\'' + id + '\\')">Move to Completed</button>' +
        '<button class="popup-btn popup-btn-secondary" onclick="closePopup()">Cancel</button>'
      );
    } else {
      var info = "";
      if (entry.sd) info = entry.sd;
      if (entry.t) info += (info ? " | " : "") + entry.t;
      showPopup("warning", "Already Shipped!", id + "\\n\\n" + info);
    }
    return;
  }

  queue.push({ id: id, entry: entry, tracking: "", status: "pending" });
  setStatus("success", "✓ " + id + " added (#" + queue.length + ")");
  updateButtons();
  renderQueue();
  switchMode(1);
}

/* ============================================================
   MODE 1: ASSIGN TRACKING (order-grouped, then back to samples)
   ============================================================ */
function assignTracking(tracking) {
  if (trackingIdx >= queue.length) {
    setStatus("error", "All samples already have tracking!");
    switchMode(0);
    return;
  }

  var item = queue[trackingIdx];
  var orderId = item.entry.o;

  /* Group all pending items with same order for single-scan tracking */
  var group = [];
  if (orderId) {
    for (var gi = 0; gi < queue.length; gi++) {
      if (queue[gi].status === "pending" && queue[gi].entry.o === orderId) {
        group.push(gi);
      }
    }
  }
  if (group.length <= 1) group = [trackingIdx];

  /* Apply tracking to all in group */
  var batchItems = [];
  for (var bi = 0; bi < group.length; bi++) {
    var gIdx = group[bi];
    queue[gIdx].tracking = tracking;
    queue[gIdx].status = "sending";
    batchItems.push({ id: queue[gIdx].id, row: queue[gIdx].entry.r, tracking: tracking });
  }
  renderQueue();

  var groupCopy = group.slice();
  var groupLen = group.length;
  inflight++;
  google.script.run
    .withSuccessHandler(function(r) {
      inflight--;
      if (r.success) {
        for (var rj = 0; rj < r.results.length; rj++) {
          var res = r.results[rj];
          for (var rk = 0; rk < queue.length; rk++) {
            if (queue[rk].id === res.id) {
              queue[rk].status = res.success ? "done" : "fail";
              if (res.success) { queue[rk].entry.s = "Shipped"; shippedCount++; }
              else { queue[rk].error = res.message; }
            }
          }
        }
        document.getElementById("count").textContent = shippedCount;
      } else {
        for (var fi = 0; fi < groupCopy.length; fi++) {
          queue[groupCopy[fi]].status = "fail";
          queue[groupCopy[fi]].error = r.message;
        }
      }
      renderQueue();
      checkAllDone();
    })
    .withFailureHandler(function(e) {
      inflight--;
      for (var ei = 0; ei < groupCopy.length; ei++) {
        queue[groupCopy[ei]].status = "fail";
        queue[groupCopy[ei]].error = e.message;
      }
      renderQueue();
    })
    .scanOutShipBatch(batchItems, IDX.cols);

  /* Advance past all grouped items */
  trackingIdx = Math.max.apply(null, group) + 1;

  /* Switch back to sample mode for next order */
  switchMode(0);
  setStatus("success", "✓ Tracking → " + groupLen + " sample" + (groupLen !== 1 ? "s" : "") + " — scan next order");
}

function checkAllDone() {
  if (inflight > 0) return;
  var allDone = queue.every(function(q) { return q.status === "done" || q.status === "fail"; });
  if (allDone && queue.length > 0) {
    var ok = queue.filter(function(q) { return q.status === "done"; }).length;
    var fail = queue.filter(function(q) { return q.status === "fail"; }).length;
    var msg = "✓ All done! " + ok + " shipped";
    if (fail > 0) msg += ", " + fail + " failed";
    setStatus(fail > 0 ? "error" : "success", msg);
  }
}

/* ============================================================
   SHIP ALL (batch — alternative to one-by-one tracking)
   ============================================================ */
function shipAll() {
  /* Only ship items that have tracking assigned */
  var items = [];
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].tracking && queue[i].status !== "done") {
      items.push({ id: queue[i].id, row: queue[i].entry.r, tracking: queue[i].tracking });
      queue[i].status = "sending";
    }
  }
  if (items.length === 0) {
    showPopup("warning", "Nothing to Ship", "Scan tracking numbers first, then Ship All.\\n\\nOr switch to Tracking tab and scan them one by one.");
    return;
  }
  renderQueue();
  setStatus("scanning", "Shipping " + items.length + " samples...");
  document.getElementById("shipAllBtn").disabled = true;

  google.script.run
    .withSuccessHandler(function(r) {
      document.getElementById("shipAllBtn").disabled = false;
      if (r.success) {
        for (var j = 0; j < r.results.length; j++) {
          var res = r.results[j];
          for (var k = 0; k < queue.length; k++) {
            if (queue[k].id === res.id) {
              queue[k].status = res.success ? "done" : "fail";
              if (res.success) {
                queue[k].entry.s = "Shipped";
                shippedCount++;
              } else {
                queue[k].error = res.message;
              }
            }
          }
        }
        document.getElementById("count").textContent = shippedCount;
        var failCount = queue.filter(function(q) { return q.status === "fail"; }).length;
        var okCount = r.shipped || 0;
        setStatus(failCount > 0 ? "error" : "success", "✓ " + okCount + " shipped" + (failCount > 0 ? ", " + failCount + " failed" : "!"));
      } else {
        // Top-level failure — mark all "sending" items as failed so they don't stay stuck
        for (var si = 0; si < queue.length; si++) {
          if (queue[si].status === "sending") {
            queue[si].status = "fail";
            queue[si].error = r.message || "Batch error";
          }
        }
        setStatus("error", r.message || "Batch error");
      }
      renderQueue();
    })
    .withFailureHandler(function(e) {
      document.getElementById("shipAllBtn").disabled = false;
      setStatus("error", "Batch error: " + e.message);
      for (var i = 0; i < queue.length; i++) {
        if (queue[i].status === "sending") queue[i].status = "fail";
      }
      renderQueue();
    })
    .scanOutShipBatch(items, IDX.cols);
}

/* ============================================================
   QUEUE RENDERING
   ============================================================ */
function renderQueue() {
  var el = document.getElementById("queueList");
  document.getElementById("queueCount").textContent = queue.length + " sample" + (queue.length !== 1 ? "s" : "");

  if (queue.length === 0) {
    el.innerHTML = '<div class="q-empty">Scan sample or order barcodes to build your queue</div>';
    return;
  }

  var h = "";
  for (var i = 0; i < queue.length; i++) {
    var q = queue[i];
    var cls = "q-item";
    if (mode === 1 && i === trackingIdx && q.status === "pending") cls += " q-active";
    if (q.status === "done") cls += " q-done";
    if (q.status === "sending") cls += " q-sending";
    if (q.status === "fail") cls += " q-error";

    var statusLabel = "";
    if (q.status === "pending" && !q.tracking) statusLabel = '<span class="q-status pending">waiting</span>';
    if (q.status === "pending" && q.tracking) statusLabel = '<span class="q-status ready">ready</span>';
    if (q.status === "sending") statusLabel = '<span class="q-status sending">saving...</span>';
    if (q.status === "done") statusLabel = '<span class="q-status done">✓ shipped</span>';
    if (q.status === "fail") statusLabel = '<span class="q-status fail">✗ failed</span>';

    h += '<div class="' + cls + '">' +
      '<span class="q-num">' + (i+1) + '.</span>' +
      '<span class="q-sample">' + q.id + '</span>' +
      '<span class="q-details">' + (q.entry.d || "") + '</span>';
    if (q.tracking) h += '<span class="q-tracking">' + q.tracking.slice(-6) + '</span>';
    h += statusLabel;
    if (q.status === "pending" && mode === 0) {
      h += '<button class="q-remove" onclick="removeFromQueue(' + i + ')" title="Remove">✕</button>';
    }
    h += '</div>';
  }
  el.innerHTML = h;

  /* Auto-scroll to active item */
  if (mode === 1) {
    var active = el.querySelector(".q-active");
    if (active) active.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }
}

function removeFromQueue(idx) {
  queue.splice(idx, 1);
  if (trackingIdx > idx) trackingIdx--;
  updateButtons();
  renderQueue();
  setStatus("ready", "Removed. " + queue.length + " in queue.");
}

function clearQueue() {
  if (queue.some(function(q) { return q.status === "sending"; })) {
    showPopup("warning", "Wait", "Some items are still being saved. Wait a moment.");
    return;
  }
  queue = [];
  trackingIdx = 0;
  updateButtons();
  renderQueue();
  switchMode(0);
}

function updateButtons() {
  document.getElementById("clearBtn").disabled = queue.length === 0;
  var hasTracking = queue.some(function(q) { return q.tracking && q.status !== "done"; });
  document.getElementById("shipAllBtn").disabled = !hasTracking;
}

/* ============================================================
   CARRIER DISPLAY
   ============================================================ */
function showCarrierInfo(carrier) {
  var el = document.getElementById("carrierInfo");
  var content = document.getElementById("carrierContent");
  if (!carrier || (!carrier.fedex && !carrier.ups)) { el.style.display = "none"; return; }
  var h = "";
  if (carrier.fedex) {
    h += '<div class="carrier-row"><span class="carrier-label">FedEx</span>';
    if (carrier.preferred === "FedEx") h += '<span class="carrier-preferred">★ PREFERRED</span>';
    h += '</div><div class="carrier-row"><span class="carrier-acct">' + carrier.fedex + '</span></div>';
  }
  if (carrier.ups) {
    if (carrier.fedex) h += '<hr style="border:none;border-top:1px solid #eee;margin:3px 0">';
    h += '<div class="carrier-row"><span class="carrier-label">UPS</span>';
    if (carrier.preferred === "UPS") h += '<span class="carrier-preferred">★ PREFERRED</span>';
    h += '</div><div class="carrier-row"><span class="carrier-acct">' + carrier.ups + '</span></div>';
  }
  content.innerHTML = h;
  el.style.display = "block";
}

/* ============================================================
   POPUP, FORCE MOVE, BATCH MOVE
   ============================================================ */
function showPopup(type, title, msg, buttons) {
  document.getElementById("popupBox").className = "popup-" + type;
  document.getElementById("popupIcon").textContent = type === "error" ? "❌" : type === "stuck" ? "🔄" : "⚠️";
  document.getElementById("popupTitle").textContent = title;
  document.getElementById("popupMessage").textContent = msg;
  document.getElementById("popupButtons").innerHTML = buttons || '<button class="popup-btn ' + (type === "warning" ? "popup-btn-warn" : "popup-btn-ok") + '" onclick="closePopup()">OK</button>';
  document.getElementById("popupOverlay").style.display = "flex";
}

function closePopup() {
  document.getElementById("popupOverlay").style.display = "none";
  document.getElementById("scanBox").focus();
}

function setStatus(cls, msg) {
  var el = document.getElementById("status");
  el.className = cls;
  el.textContent = msg;
}

function forceMove(sampleId) {
  closePopup();
  setStatus("scanning", "Moving " + sampleId + "...");
  google.script.run
    .withSuccessHandler(function(r) {
      if (r.success) {
        if (IDX && IDX.liveSet) delete IDX.liveSet[sampleId];
        setStatus("success", "✓ " + sampleId + " moved!");
      } else { showPopup("error", "Move Failed", r.message); }
    })
    .withFailureHandler(function(e) { showPopup("error", "Error", e.message); })
    .forceMoveSample(sampleId);
}

function runBatchMove() {
  var btn = document.getElementById("moveBtn");
  btn.textContent = "Moving...";
  btn.disabled = true;
  google.script.run
    .withSuccessHandler(function(r) {
      btn.textContent = "📦 Move Shipped → Completed";
      btn.disabled = false;
      var ts = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
      document.getElementById("moveStatus").textContent = (r.message || "Done") + " (" + ts + ")";
    })
    .withFailureHandler(function(e) {
      btn.textContent = "📦 Move Shipped → Completed";
      btn.disabled = false;
      document.getElementById("moveStatus").textContent = "Error: " + e.message;
    })
    .processPendingMoves();
}

/* ============================================================
   INIT
   ============================================================ */
loadIndex();
</script>
</body>
</html>`;
}
