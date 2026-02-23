// ============================================================
// QRCode.gs — QR Code Generation, Parsing & Label Styles
// Extracted from Main.gs — February 4, 2026
//
// Functions in this file:
//   buildQRDataString(data, csSampleNumber)  — builds pipe-delimited QR payload
//   getQRCodeUrl(qrData, size)               — returns quickchart.io image URL
//   parseQRData(code)                        — parses QR/barcode string → fields object
//   generateArrivalLabels()                  — compact horizontal: QR + date + warehouse + customer
//   generateArrivalLabelsVertical()          — vertical: QR on top, text below
//   generateKeurigLabels()                   — 4×4 full-detail: two-column + barcode footer
//   testQRCodeGeneration()                   — dialog to verify QR API works
//   _extractCertification(description)       — helper: RFA → RAINFOREST ALLIANCE, etc.
//   _formatMarkDashes(mark)                  — helper: 003/0311/00417 → 003-00311-00417
//
// NOTE: All functions are globally available to Main.gs, ScanOut.gs, etc.
//       After deploying this file, update Main.gs to call buildQRDataString()
//       instead of inline qrData construction (see MIGRATION NOTES at bottom).
// ============================================================


// ============================================================
// 🔧 QR DATA STRING BUILDER
// Centralizes the pipe-delimited format used by all QR codes.
// Called by addDataToSheet, addMultipleSamplesAsNewOrder, etc.
// ============================================================
function buildQRDataString(data, csSampleNumber) {
  return 'ORDER:' + (data.sampleOrderNum || '') +
    '|CARGO:' + (data.cargo || '') +
    '|MARK:' + (data.mark || '') +
    '|CONTAINER:' + (data.container || '') +
    '|REF:' + (data.reference || '') +
    '|DESC:' + (data.description || '') +
    '|BAGS:' + (data.bagCount || '') +
    '|WEIGHT:' + (data.weight || '') +
    '|SAMPLE:' + (data.sampleWeight || '') +
    '|P:' + (data.pNumber || '') +
    '|S:' + (data.sNumber || '') +
    '|WAREHOUSE:' + (data.warehouse || '') +
    '|CS_SAMPLE:' + (csSampleNumber || '');
}


// ============================================================
// 🔗 QR CODE URL GENERATOR
// Returns a quickchart.io URL for a given data string.
// size = pixel width/height (default 350)
// ============================================================
function getQRCodeUrl(qrData, size) {
  size = size || 350;
  return 'https://quickchart.io/qr?text=' + encodeURIComponent(qrData) +
    '&size=' + size + '&ecLevel=H&margin=2';
}


// ============================================================
// 🔍 QR DATA PARSER
// Takes a raw scanned string and extracts fields.
// Returns { success, fields, barcodeOnly? }
//
// Handles three formats:
//   1. Full labeled: ORDER:...|CARGO:...|...
//   2. Pipe-delimited positional: value|value|...
//   3. Bare barcode: 211657-01
// ============================================================
function parseQRData(code) {
  Logger.log('parseQRData: Processing code length=' + (code ? code.length : 0));
  try {
    var fields = {
      sampleOrderNum: '', cargo: '', mark: '', container: '', reference: '',
      description: '', bagCount: '', weight: '', sampleWeight: '',
      pNumber: '', sNumber: '', warehouse: '', originalSampleNum: ''
    };

    var cleanCode = code;

    // Try to extract the full labeled QR pattern if embedded in noise
    if (code.indexOf('ORDER:') !== -1 && code.indexOf('CARGO:') !== -1) {
      var qrMatch = code.match(
        /ORDER:[^|]+\|CARGO:[^|]+\|MARK:[^|]+\|CONTAINER:[^|]+\|REF:[^|]+\|DESC:[^|]+\|BAGS:[^|]+\|WEIGHT:[^|]*\|SAMPLE:[^|]+\|P:[^|]*\|S:[^|]*\|WAREHOUSE:[^|]+\|CS_SAMPLE:\d{6}-\d{2}/
      );
      if (qrMatch) cleanCode = qrMatch[0];
    }

    // ── Format 1: Full labeled ORDER:...|CARGO:... ──
    if (cleanCode.indexOf('ORDER:') !== -1 && cleanCode.indexOf('CARGO:') !== -1) {
      var pos, endPos;

      pos = cleanCode.indexOf('ORDER:'); endPos = cleanCode.indexOf('|CARGO:');
      if (pos !== -1 && endPos !== -1) fields.sampleOrderNum = cleanCode.substring(pos + 6, endPos).trim();

      pos = cleanCode.indexOf('CARGO:'); endPos = cleanCode.indexOf('|MARK:');
      if (pos !== -1 && endPos !== -1) fields.cargo = cleanCode.substring(pos + 6, endPos).trim();

      pos = cleanCode.indexOf('MARK:'); endPos = cleanCode.indexOf('|CONTAINER:');
      if (pos !== -1 && endPos !== -1) fields.mark = cleanCode.substring(pos + 5, endPos).trim();

      pos = cleanCode.indexOf('CONTAINER:'); endPos = cleanCode.indexOf('|REF:');
      if (pos !== -1 && endPos !== -1) fields.container = cleanCode.substring(pos + 10, endPos).trim();

      pos = cleanCode.indexOf('REF:'); endPos = cleanCode.indexOf('|DESC:');
      if (pos !== -1 && endPos !== -1) {
        fields.reference = cleanCode.substring(pos + 4, endPos).trim();
        fields.pNumber = fields.reference;
      }

      pos = cleanCode.indexOf('DESC:'); endPos = cleanCode.indexOf('|BAGS:');
      if (pos !== -1 && endPos !== -1) fields.description = cleanCode.substring(pos + 5, endPos).trim();

      pos = cleanCode.indexOf('BAGS:'); endPos = cleanCode.indexOf('|WEIGHT:');
      if (pos !== -1 && endPos !== -1) fields.bagCount = cleanCode.substring(pos + 5, endPos).trim();

      pos = cleanCode.indexOf('WEIGHT:'); endPos = cleanCode.indexOf('|SAMPLE:');
      if (pos !== -1 && endPos !== -1) fields.weight = cleanCode.substring(pos + 7, endPos).trim();

      pos = cleanCode.indexOf('SAMPLE:'); endPos = cleanCode.indexOf('|P:');
      if (pos !== -1 && endPos !== -1) fields.sampleWeight = cleanCode.substring(pos + 7, endPos).trim();

      pos = cleanCode.indexOf('|P:'); endPos = cleanCode.indexOf('|S:');
      if (pos !== -1 && endPos !== -1) {
        var pVal = cleanCode.substring(pos + 3, endPos).trim();
        if (pVal) fields.pNumber = pVal;
      }

      pos = cleanCode.indexOf('|S:'); endPos = cleanCode.indexOf('|WAREHOUSE:');
      if (pos !== -1 && endPos !== -1) fields.sNumber = cleanCode.substring(pos + 3, endPos).trim();

      pos = cleanCode.indexOf('WAREHOUSE:'); endPos = cleanCode.indexOf('|CS_SAMPLE:');
      if (pos !== -1 && endPos !== -1) fields.warehouse = cleanCode.substring(pos + 10, endPos).trim();

      pos = cleanCode.indexOf('CS_SAMPLE:');
      if (pos !== -1) {
        endPos = cleanCode.indexOf('|', pos + 10);
        if (endPos === -1) endPos = cleanCode.length;
        fields.originalSampleNum = cleanCode.substring(pos + 10, endPos).trim();
      }

      return { success: true, fields: fields };
    }

    // ── Format 2: Pipe-delimited positional ──
    if (cleanCode.indexOf('|') !== -1) {
      var parts = cleanCode.split('|');
      parts.forEach(function(part) {
        var t = String(part).trim();
        if (!t) return;
        if (/^\d{6}-\d{2}$/.test(t)) fields.originalSampleNum = t;
        else if (/^[A-Z]{4}\d{7}$/.test(t)) fields.container = t;
        else if (/^\d{3}\/\d{4}\/\d{4,5}$/.test(t)) fields.mark = t;
        else if (/^C\d{6}$/.test(t)) fields.cargo = t;
        else if (/^P\d+[-]\d+/.test(t)) { fields.reference = t; fields.pNumber = t; }
        else if (/warehouse/i.test(t)) fields.warehouse = t;
        else if (/^\d{1,4}$/.test(t) && parseInt(t) > 10) fields.bagCount = t;
        else if (/\d+\s*(LB|lb|kg|KG)/i.test(t)) {
          if (!fields.sampleWeight) fields.sampleWeight = t; else fields.weight = t;
        }
        else if (t.length > 10 && /[a-zA-Z]/.test(t) && !fields.description) fields.description = t;
        else if (!fields.sampleOrderNum && t.length > 3) fields.sampleOrderNum = t;
      });
      return { success: true, fields: fields };
    }

    // ── Format 3: Bare barcode (e.g. 211657-01) ──
    if (/^\d{6}-\d{2}$/.test(cleanCode)) {
      fields.originalSampleNum = cleanCode;
      return { success: true, fields: fields, barcodeOnly: true };
    }

    return { success: false, message: 'Could not parse QR code format' };
  } catch (e) {
    return { success: false, message: 'Parse error: ' + e.toString() };
  }
}

// Backward-compatible alias — existing code calls parseQRForQueue
function parseQRForQueue(code) {
  return parseQRData(code);
}


// ============================================================
// 🏷️ ARRIVAL LABELS — Compact format
//
// Matches the Keurig-style label:
//   ┌──────────────────┐
//   │                  │
//   │    [QR CODE]     │
//   │                  │
//   │ Arrived 1/5/2026 │
//   │ CONTINENTAL NJ   │
//   │ VOLCAFE NORTH AME│
//   └──────────────────┘
//
// Select rows in All Orders → run from menu.
// Designed for small label stock (2.25" x 1.25" or similar).
// ============================================================
function generateArrivalLabels() {
  Logger.log('generateArrivalLabels: Starting label generation');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  var selection = sheet.getActiveRange();
  var startRow = selection.getRow();
  var numRows = selection.getNumRows();
  if (startRow === 1) { ui.alert('Please select data rows (not header).'); return; }

  var answer = ui.alert(
    'Generate ' + numRows + ' arrival label(s)?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  var col = _getColumnMap(sheet);
  var labels = [];
  var allRows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  for (var i = 0; i < numRows; i++) {
    var row = allRows[i];

    // Build QR data payload
    var qrDataStr = row[col['QR Data']] || '';
    if (!qrDataStr || qrDataStr === 'undefined' || qrDataStr.trim() === '') {
      qrDataStr = buildQRDataString({
        sampleOrderNum: row[col['Sample Order #']] || '',
        cargo:          row[col['Cargo #']] || '',
        mark:           row[col['Mark #']] || '',
        container:      row[col['Container #']] || '',
        reference:      row[col['Reference']] || '',
        description:    row[col['Description']] || '',
        bagCount:       row[col['Bag Count']] || '',
        weight:         row[col['Weight']] || '',
        sampleWeight:   row[col['Sample Weight']] || '',
        pNumber:        row[col['P #']] || '',
        sNumber:        row[col['S #']] || '',
        warehouse:      row[col['Warehouse']] || ''
      }, row[col['CS Sample #']] || '');
    }

    // Arrival date — use Timestamp from the row, or today
    var ts = row[col['Timestamp']];
    var arrivalDate;
    if (ts instanceof Date && !isNaN(ts)) {
      arrivalDate = (ts.getMonth() + 1) + '/' + ts.getDate() + '/' + ts.getFullYear();
    } else {
      var now = new Date();
      arrivalDate = (now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear();
    }

    // Warehouse — short name, uppercased
    var warehouse = String(row[col['Warehouse']] || '').toUpperCase();
    // Trim to something concise like "CONTINENTAL NJ"
    warehouse = warehouse
      .replace(/\bWAREHOUSE\b/gi, '')
      .replace(/\bDISTRIBUTION\b/gi, '')
      .replace(/\bSERVICES\b/gi, '')
      .replace(/\bINC\.?\b/gi, '')
      .replace(/\bLLC\b/gi, '')
      .replace(/,/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();

    // Customer — sender name, uppercased, truncated to ~20 chars
    var customer = String(row[col['Sender']] || '').toUpperCase();
    if (customer.length > 22) customer = customer.substring(0, 22);

    labels.push({
      qrData: qrDataStr,
      arrivalDate: arrivalDate,
      warehouse: warehouse,
      customer: customer,
      csSample: row[col['CS Sample #']] || '',
      container: row[col['Container #']] || ''
    });
  }

  // ── Build HTML ──
  var html = '<!DOCTYPE html>\n<html>\n<head>\n<style>\n';
  html += '* { margin: 0; padding: 0; box-sizing: border-box; }\n';
  html += '@media print {\n';
  html += '  @page { size: 2.25in 1.25in; margin: 0; }\n';
  html += '  body { margin: 0; }\n';
  html += '  .label { page-break-after: always; border: none !important; box-shadow: none !important; }\n';
  html += '  .label:last-child { page-break-after: auto; }\n';
  html += '  .no-print { display: none !important; }\n';
  html += '}\n';
  html += '@media screen {\n';
  html += '  body { background: #e0e0e0; padding: 20px; font-family: Arial, sans-serif; }\n';
  html += '  .label { margin: 15px auto; box-shadow: 0 2px 8px rgba(0,0,0,0.2); }\n';
  html += '}\n';
  html += '.label {\n';
  html += '  width: 2.25in; height: 1.25in;\n';
  html += '  background: white;\n';
  html += '  padding: 0.06in 0.08in;\n';
  html += '  display: flex;\n';
  html += '  flex-direction: row;\n';
  html += '  align-items: center;\n';
  html += '  gap: 0.08in;\n';
  html += '  overflow: hidden;\n';
  html += '}\n';
  html += '.qr-box {\n';
  html += '  width: 1in; min-width: 1in; height: 1in;\n';
  html += '  display: flex; align-items: center; justify-content: center;\n';
  html += '}\n';
  html += '.qr-box img { width: 100%; height: 100%; object-fit: contain; }\n';
  html += '.info-box {\n';
  html += '  flex: 1;\n';
  html += '  display: flex; flex-direction: column;\n';
  html += '  justify-content: center;\n';
  html += '  overflow: hidden;\n';
  html += '}\n';
  html += '.arrived {\n';
  html += '  font-size: 9pt; font-weight: bold; color: #000;\n';
  html += '  margin-bottom: 2px; white-space: nowrap;\n';
  html += '}\n';
  html += '.warehouse {\n';
  html += '  font-size: 8pt; font-weight: bold; color: #000;\n';
  html += '  margin-bottom: 1px; white-space: nowrap;\n';
  html += '  overflow: hidden; text-overflow: ellipsis;\n';
  html += '}\n';
  html += '.customer {\n';
  html += '  font-size: 8pt; font-weight: bold; color: #000;\n';
  html += '  white-space: nowrap;\n';
  html += '  overflow: hidden; text-overflow: ellipsis;\n';
  html += '}\n';
  html += '</style>\n</head>\n<body>\n';

  // Controls
  html += '<div class="no-print" style="text-align:center; margin-bottom:15px; font-family:Arial;">\n';
  html += '  <h3>Arrival Labels — ' + labels.length + ' label(s)</h3>\n';
  html += '  <p style="color:#666; font-size:13px;">Format: QR + Arrived date + Warehouse + Customer</p>\n';
  html += '  <div style="margin:10px 0;">\n';
  html += '    <label style="font-size:13px;">Label size: </label>\n';
  html += '    <select id="sizeSelect" onchange="changeSize(this.value)" style="padding:4px;">\n';
  html += '      <option value="2.25in|1.25in" selected>2.25" × 1.25" (standard)</option>\n';
  html += '      <option value="2in|2in">2" × 2" (square)</option>\n';
  html += '      <option value="3in|2in">3" × 2" (wide)</option>\n';
  html += '      <option value="4in|2in">4" × 2" (shipping)</option>\n';
  html += '    </select>\n';
  html += '  </div>\n';
  html += '  <button onclick="window.print()" style="padding:10px 24px; font-size:15px; background:#4CAF50; color:white; border:none; border-radius:4px; cursor:pointer; margin:5px;">🖨️ Print Labels</button>\n';
  html += '  <button onclick="google.script.host.close()" style="padding:10px 24px; font-size:15px; background:#666; color:white; border:none; border-radius:4px; cursor:pointer; margin:5px;">Close</button>\n';
  html += '</div>\n';

  // Labels
  for (var j = 0; j < labels.length; j++) {
    var lb = labels[j];
    var qrUrl = getQRCodeUrl(lb.qrData, 300);

    html += '<div class="label">\n';
    html += '  <div class="qr-box"><img src="' + qrUrl + '" alt="QR" /></div>\n';
    html += '  <div class="info-box">\n';
    html += '    <div class="arrived">Arrived&nbsp;&nbsp;&nbsp;' + lb.arrivalDate + '</div>\n';
    html += '    <div class="warehouse">' + lb.warehouse + '</div>\n';
    html += '    <div class="customer">' + lb.customer + '</div>\n';
    html += '  </div>\n';
    html += '</div>\n';
  }

  // Size-change script
  html += '<script>\n';
  html += 'function changeSize(val) {\n';
  html += '  var parts = val.split("|");\n';
  html += '  var labels = document.querySelectorAll(".label");\n';
  html += '  for (var i = 0; i < labels.length; i++) {\n';
  html += '    labels[i].style.width = parts[0];\n';
  html += '    labels[i].style.height = parts[1];\n';
  html += '  }\n';
  html += '  // Update print page size\n';
  html += '  var style = document.getElementById("dynamicPageSize");\n';
  html += '  if (!style) { style = document.createElement("style"); style.id = "dynamicPageSize"; document.head.appendChild(style); }\n';
  html += '  style.textContent = "@media print { @page { size: " + parts[0] + " " + parts[1] + "; } }";\n';
  html += '}\n';
  html += '</script>\n';

  html += '</body>\n</html>';

  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(600);
  ui.showModalDialog(htmlOutput, 'Arrival Labels — ' + labels.length + ' label(s)');
}


// ============================================================
// 🏷️ ARRIVAL LABELS — VERTICAL (large QR on top)
//
// Closer to the exact photo layout:
//   ┌──────────────┐
//   │  [QR CODE]   │
//   │              │
//   │ Arrived date │
//   │ WAREHOUSE    │
//   │ CUSTOMER     │
//   └──────────────┘
//
// For square or tall label stock.
// ============================================================
function generateArrivalLabelsVertical() {
  Logger.log('generateArrivalLabelsVertical: Starting vertical label generation');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  var selection = sheet.getActiveRange();
  var startRow = selection.getRow();
  var numRows = selection.getNumRows();
  if (startRow === 1) { ui.alert('Please select data rows (not header).'); return; }

  var answer = ui.alert(
    'Generate ' + numRows + ' vertical arrival label(s)?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  var col = _getColumnMap(sheet);
  var labels = [];
  var allRows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  for (var i = 0; i < numRows; i++) {
    var row = allRows[i];

    var qrDataStr = row[col['QR Data']] || '';
    if (!qrDataStr || qrDataStr === 'undefined' || qrDataStr.trim() === '') {
      qrDataStr = buildQRDataString({
        sampleOrderNum: row[col['Sample Order #']] || '',
        cargo:          row[col['Cargo #']] || '',
        mark:           row[col['Mark #']] || '',
        container:      row[col['Container #']] || '',
        reference:      row[col['Reference']] || '',
        description:    row[col['Description']] || '',
        bagCount:       row[col['Bag Count']] || '',
        weight:         row[col['Weight']] || '',
        sampleWeight:   row[col['Sample Weight']] || '',
        pNumber:        row[col['P #']] || '',
        sNumber:        row[col['S #']] || '',
        warehouse:      row[col['Warehouse']] || ''
      }, row[col['CS Sample #']] || '');
    }

    var ts = row[col['Timestamp']];
    var arrivalDate;
    if (ts instanceof Date && !isNaN(ts)) {
      arrivalDate = (ts.getMonth() + 1) + '/' + ts.getDate() + '/' + ts.getFullYear();
    } else {
      var now = new Date();
      arrivalDate = (now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear();
    }

    var warehouse = String(row[col['Warehouse']] || '').toUpperCase()
      .replace(/\bWAREHOUSE\b/gi, '').replace(/\bDISTRIBUTION\b/gi, '')
      .replace(/\bSERVICES\b/gi, '').replace(/\bINC\.?\b/gi, '')
      .replace(/\bLLC\b/gi, '').replace(/,/g, '').replace(/\s{2,}/g, ' ').trim();

    var customer = String(row[col['Sender']] || '').toUpperCase();
    if (customer.length > 25) customer = customer.substring(0, 25);

    labels.push({
      qrData: qrDataStr,
      arrivalDate: arrivalDate,
      warehouse: warehouse,
      customer: customer
    });
  }

  // ── HTML — vertical layout matching the photo exactly ──
  var html = '<!DOCTYPE html>\n<html>\n<head>\n<style>\n';
  html += '* { margin: 0; padding: 0; box-sizing: border-box; }\n';
  html += '@media print {\n';
  html += '  @page { size: 2in 2in; margin: 0; }\n';
  html += '  body { margin: 0; }\n';
  html += '  .label { page-break-after: always; border: none !important; box-shadow: none !important; }\n';
  html += '  .label:last-child { page-break-after: auto; }\n';
  html += '  .no-print { display: none !important; }\n';
  html += '}\n';
  html += '@media screen {\n';
  html += '  body { background: #e0e0e0; padding: 20px; font-family: Arial, sans-serif; }\n';
  html += '  .label { margin: 15px auto; box-shadow: 0 2px 8px rgba(0,0,0,0.2); }\n';
  html += '}\n';
  html += '.label {\n';
  html += '  width: 2in; height: 2in;\n';
  html += '  background: white;\n';
  html += '  padding: 0.06in;\n';
  html += '  display: flex;\n';
  html += '  flex-direction: column;\n';
  html += '  align-items: center;\n';
  html += '  justify-content: center;\n';
  html += '  overflow: hidden;\n';
  html += '}\n';
  html += '.qr-box {\n';
  html += '  width: 1.4in; height: 1.4in;\n';
  html += '  display: flex; align-items: center; justify-content: center;\n';
  html += '  margin-bottom: 0.04in;\n';
  html += '}\n';
  html += '.qr-box img { width: 100%; height: 100%; object-fit: contain; }\n';
  html += '.text-area {\n';
  html += '  width: 100%;\n';
  html += '  text-align: left;\n';
  html += '  padding-left: 0.08in;\n';
  html += '}\n';
  html += '.arrived {\n';
  html += '  font-size: 8pt; font-weight: bold; color: #000;\n';
  html += '  margin-bottom: 1px; white-space: nowrap;\n';
  html += '}\n';
  html += '.warehouse {\n';
  html += '  font-size: 7.5pt; font-weight: bold; color: #000;\n';
  html += '  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\n';
  html += '}\n';
  html += '.customer {\n';
  html += '  font-size: 7.5pt; font-weight: bold; color: #000;\n';
  html += '  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\n';
  html += '}\n';
  html += '</style>\n</head>\n<body>\n';

  html += '<div class="no-print" style="text-align:center; margin-bottom:15px;">\n';
  html += '  <h3>Arrival Labels (Vertical) — ' + labels.length + ' label(s)</h3>\n';
  html += '  <p style="color:#666; font-size:13px;">QR on top · Arrived + Warehouse + Customer below</p>\n';
  html += '  <button onclick="window.print()" style="padding:10px 24px; font-size:15px; background:#4CAF50; color:white; border:none; border-radius:4px; cursor:pointer; margin:5px;">🖨️ Print</button>\n';
  html += '  <button onclick="google.script.host.close()" style="padding:10px 24px; font-size:15px; background:#666; color:white; border:none; border-radius:4px; cursor:pointer; margin:5px;">Close</button>\n';
  html += '</div>\n';

  for (var j = 0; j < labels.length; j++) {
    var lb = labels[j];
    var qrUrl = getQRCodeUrl(lb.qrData, 400);

    html += '<div class="label">\n';
    html += '  <div class="qr-box"><img src="' + qrUrl + '" alt="QR" /></div>\n';
    html += '  <div class="text-area">\n';
    html += '    <div class="arrived">Arrived&nbsp;&nbsp;&nbsp;&nbsp;' + lb.arrivalDate + '</div>\n';
    html += '    <div class="warehouse">' + lb.warehouse + '</div>\n';
    html += '    <div class="customer">' + lb.customer + '</div>\n';
    html += '  </div>\n';
    html += '</div>\n';
  }

  html += '</body>\n</html>';

  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(600);
  ui.showModalDialog(htmlOutput, 'Arrival Labels (Vertical) — ' + labels.length);
}


// ============================================================
// 🧪 TEST — QR Code API Verification
// ============================================================
function testQRCodeGeneration() {
  Logger.log('testQRCodeGeneration: Testing QR code API');
  var ui = SpreadsheetApp.getUi();
  var testData = 'ORDER:TEST|CARGO:123|CONTAINER:MSMU1234567|CS_SAMPLE:211700-01';
  var qrUrl = getQRCodeUrl(testData, 200);

  var html = '<!DOCTYPE html>\n<html><head><style>\n';
  html += 'body { font-family: Arial; padding: 20px; text-align: center; }\n';
  html += 'img { border: 2px solid #333; padding: 10px; margin: 20px; }\n';
  html += '.url { background: #f5f5f5; padding: 10px; margin: 20px; word-wrap: break-word; font-size: 12px; }\n';
  html += '</style></head><body>\n';
  html += '<h2>QR Code Test</h2>\n';
  html += '<p>Test Data: ' + testData + '</p>\n';
  html += '<p>API: QuickChart.io</p>\n';
  html += '<h3>QR Code:</h3>\n';
  html += '<img src="' + qrUrl + '" alt="QR Code" onerror="this.alt=\'FAILED TO LOAD\';" />\n';
  html += '<div class="url"><strong>URL:</strong><br>' + qrUrl + '</div>\n';
  html += '<p>If you see a QR code above, the API works!</p>\n';
  html += '<button onclick="window.open(\'' + qrUrl + '\')">Open QR URL in New Tab</button>\n';
  html += '</body></html>';

  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(600).setHeight(500);
  ui.showModalDialog(htmlOutput, 'QR Code Test — QuickChart.io');
}


// ============================================================
// 🏷️ KEURIG LABELS — 4" × 4" Full-Detail Label
//
// Matches the full warehouse label layout:
//   ┌─────────────────────────────────────────────┐
//   │ [QR CODE]         │  4000107013             │
//   │                   │  COLOMBIA               │
//   │                   │  RAINFOREST ALLIANCE     │
//   │                   │  0.000      lbs          │
//   │ Arrived 1/5/2026  │  Bags 20       2026      │
//   │ CONTINENTAL NJ    │  Sampled  2/4/2026       │
//   │ VOLCAFE NORTH AME │  C375519                 │
//   │                   │  S18377.004              │
//   │ CO RFA EXCELSO EP │                          │
//   │ TCKU1274162       │  003-00311-00417         │
//   │ Commodity Sampler │  |||||||||||||||||||||||  │
//   │ Services     [☕]  │  211563-01               │
//   └─────────────────────────────────────────────┘
// ============================================================

/**
 * Extract certification name from description abbreviation.
 * E.g. "CO RFA EXCELSO EP" → "RAINFOREST ALLIANCE"
 */
function _extractCertification(description) {
  if (!description) return '';
  var desc = String(description).toUpperCase();

  var certs = [
    { abbr: /\bRFA\b/,  full: 'RAINFOREST ALLIANCE' },
    { abbr: /\bRA\b/,   full: 'RAINFOREST ALLIANCE' },
    { abbr: /\bFT\b/,   full: 'FAIR TRADE' },
    { abbr: /\bORG\b/,  full: 'ORGANIC' },
    { abbr: /\bUTZ\b/,  full: 'UTZ CERTIFIED' },
    { abbr: /\b4C\b/,   full: '4C CERTIFIED' },
    { abbr: /\bAAA\b/,  full: 'AAA SUSTAINABLE' },
    { abbr: /\bFLO\b/,  full: 'FAIRTRADE (FLO)' }
  ];

  var found = [];
  for (var i = 0; i < certs.length; i++) {
    if (certs[i].abbr.test(desc)) {
      found.push(certs[i].full);
    }
  }
  return found.join(' / ');
}

/**
 * Reformat ICO mark from slashes to dashes with zero-padding.
 * "003/0311/00417" → "003-00311-00417"
 * "003/0279/60336" → "003-00279-60336"
 */
function _formatMarkDashes(mark) {
  if (!mark) return '';
  var parts = String(mark).split(/[\/\-]/);
  if (parts.length === 3) {
    return parts[0].padStart(3, '0') + '-' +
           parts[1].padStart(5, '0') + '-' +
           parts[2].padStart(5, '0');
  }
  return String(mark).replace(/\//g, '-');
}


function generateKeurigLabels() {
  Logger.log('generateKeurigLabels: Starting Keurig label generation');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  var selection = sheet.getActiveRange();
  var startRow = selection.getRow();
  var numRows = selection.getNumRows();
  if (startRow === 1) { ui.alert('Please select data rows (not header).'); return; }

  var answer = ui.alert(
    'Generate ' + numRows + ' Keurig-style label(s)?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  var col = _getColumnMap(sheet);
  var labels = [];
  var allRows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  for (var i = 0; i < numRows; i++) {
    var row = allRows[i];

    // ── QR data ──
    var qrDataStr = row[col['QR Data']] || '';
    if (!qrDataStr || qrDataStr === 'undefined' || qrDataStr.trim() === '') {
      qrDataStr = buildQRDataString({
        sampleOrderNum: row[col['Sample Order #']] || '',
        cargo:          row[col['Cargo #']] || '',
        mark:           row[col['Mark #']] || '',
        container:      row[col['Container #']] || '',
        reference:      row[col['Reference']] || '',
        description:    row[col['Description']] || '',
        bagCount:       row[col['Bag Count']] || '',
        weight:         row[col['Weight']] || '',
        sampleWeight:   row[col['Sample Weight']] || '',
        pNumber:        row[col['P #']] || '',
        sNumber:        row[col['S #']] || '',
        warehouse:      row[col['Warehouse']] || ''
      }, row[col['CS Sample #']] || '');
    }

    // ── Arrived date (from Timestamp) ──
    var ts = row[col['Timestamp']];
    var arrivedDate = '';
    if (ts instanceof Date && !isNaN(ts)) {
      arrivedDate = (ts.getMonth() + 1) + '/' + ts.getDate() + '/' + ts.getFullYear();
    } else {
      var now = new Date();
      arrivedDate = (now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear();
    }

    // ── Year from timestamp ──
    var tsYear = '';
    if (ts instanceof Date && !isNaN(ts)) {
      tsYear = String(ts.getFullYear());
    } else {
      tsYear = String(new Date().getFullYear());
    }

    // ── Sampled date (Scanned Date if available, else today) ──
    var sampledDate = '';
    if (col['Scanned Date'] !== undefined && row[col['Scanned Date']]) {
      var sd = row[col['Scanned Date']];
      if (sd instanceof Date && !isNaN(sd)) {
        sampledDate = (sd.getMonth() + 1) + '/' + sd.getDate() + '/' + sd.getFullYear();
      }
    }
    if (!sampledDate) {
      var today = new Date();
      sampledDate = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();
    }

    // ── Warehouse short name ──
    var warehouse = String(row[col['Warehouse']] || '').toUpperCase()
      .replace(/\bWAREHOUSE\b/gi, '').replace(/\bDISTRIBUTION\b/gi, '')
      .replace(/\bSERVICES\b/gi, '').replace(/\bINC\.?\b/gi, '')
      .replace(/\bLLC\b/gi, '').replace(/,/g, '').replace(/\s{2,}/g, ' ').trim();

    // ── Customer (sender) — strip company suffixes ──
    var customer = String(row[col['Sender']] || '').toUpperCase()
      .replace(/\bCOMPANY\b/gi, '')
      .replace(/\bCORPORATION\b/gi, '')
      .replace(/\bCORP\.?\b/gi, '')
      .replace(/\bINC\.?\b/gi, '')
      .replace(/\bLLC\b/gi, '')
      .replace(/\bLTD\.?\b/gi, '')
      .replace(/\bNA\b/gi, '')
      .replace(/,/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
    if (customer.length > 25) customer = customer.substring(0, 25);

    // ── Description ──
    var description = String(row[col['Description']] || '').toUpperCase();

    // ── Country from ICO mark ──
    var mark = row[col['Mark #']] || '';
    var country = '';
    if (typeof getCountryFromICOMark === 'function') {
      country = getCountryFromICOMark(mark);
    }
    country = (country || '').toUpperCase();

    // ── Certification from description ──
    var certification = _extractCertification(description);

    // ── Mark reformatted with dashes ──
    var markFormatted = _formatMarkDashes(mark);

    // ── Weight ──
    var weight = row[col['Weight']] || '0.000';

    // ── Other fields ──
    var pNumber      = String(row[col['P #']] || '');
    var bagCount     = String(row[col['Bag Count']] || '');
    var cargo        = String(row[col['Cargo #']] || '');
    var sampleOrder  = String(row[col['Sample Order #']] || '');
    var container    = String(row[col['Container #']] || '');
    var csSample     = String(row[col['CS Sample #']] || '');

    labels.push({
      qrData: qrDataStr,
      arrivedDate: arrivedDate,
      sampledDate: sampledDate,
      tsYear: tsYear,
      warehouse: warehouse,
      customer: customer,
      description: description,
      country: country,
      certification: certification,
      weight: weight,
      pNumber: pNumber,
      bagCount: bagCount,
      cargo: cargo,
      sampleOrder: sampleOrder,
      container: container,
      csSample: csSample,
      markFormatted: markFormatted
    });
  }

  // ── Build HTML ──
  var html = '<!DOCTYPE html>\n<html>\n<head>\n<style>\n';

  // Reset
  html += '* { margin: 0; padding: 0; box-sizing: border-box; }\n';

  // Print
  html += '@media print {\n';
  html += '  @page { size: 4in 4in; margin: 0; }\n';
  html += '  body { margin: 0; }\n';
  html += '  .label { page-break-after: always; border: none !important; box-shadow: none !important; }\n';
  html += '  .label:last-child { page-break-after: auto; }\n';
  html += '  .no-print { display: none !important; }\n';
  html += '}\n';

  // Screen
  html += '@media screen {\n';
  html += '  body { background: #d0d0d0; padding: 20px; font-family: Arial, sans-serif; }\n';
  html += '  .label { margin: 20px auto; box-shadow: 0 3px 12px rgba(0,0,0,0.25); }\n';
  html += '}\n';

  // Label container
  html += '.label {\n';
  html += '  width: 4in; height: 4in;\n';
  html += '  background: white;\n';
  html += '  padding: 0.12in;\n';
  html += '  display: grid;\n';
  html += '  grid-template-columns: 1.85in 1fr;\n';
  html += '  grid-template-rows: 1fr auto;\n';
  html += '  font-family: Arial, Helvetica, sans-serif;\n';
  html += '  overflow: hidden;\n';
  html += '}\n';

  // Left column
  html += '.left-col {\n';
  html += '  display: flex;\n';
  html += '  flex-direction: column;\n';
  html += '  justify-content: space-between;\n';
  html += '  padding-right: 0.1in;\n';
  html += '}\n';

  // QR
  html += '.qr-box {\n';
  html += '  width: 1.65in; height: 1.65in;\n';
  html += '  display: flex; align-items: center; justify-content: center;\n';
  html += '  margin-bottom: 0.04in;\n';
  html += '}\n';
  html += '.qr-box img { width: 100%; height: 100%; object-fit: contain; }\n';

  // Left text lines
  html += '.left-line {\n';
  html += '  font-size: 10pt; font-weight: bold; color: #000;\n';
  html += '  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\n';
  html += '  line-height: 1.35;\n';
  html += '}\n';
  html += '.left-line.arrived { font-size: 10pt; }\n';
  html += '.left-line.warehouse { font-size: 10pt; }\n';
  html += '.left-line.customer { font-size: 10pt; }\n';
  html += '.left-line.desc { font-size: 9pt; margin-top: 0.06in; }\n';
  html += '.left-line.container { font-size: 10pt; }\n';

  // CSS branding
  html += '.css-branding {\n';
  html += '  display: flex; align-items: center; gap: 0.06in;\n';
  html += '  margin-top: 0.04in;\n';
  html += '}\n';
  html += '.css-branding .css-name {\n';
  html += '  font-size: 10pt; font-weight: bold; color: #000; line-height: 1.2;\n';
  html += '}\n';
  html += '.css-branding .css-logo {\n';
  html += '  width: 0.45in; height: 0.45in;\n';
  html += '  border: 1px solid #999; border-radius: 4px;\n';
  html += '  display: flex; align-items: center; justify-content: center;\n';
  html += '  font-size: 22pt;\n';
  html += '}\n';

  // Right column
  html += '.right-col {\n';
  html += '  display: flex;\n';
  html += '  flex-direction: column;\n';
  html += '  justify-content: space-between;\n';
  html += '  padding-left: 0.08in;\n';
  html += '}\n';
  html += '.right-line {\n';
  html += '  font-size: 10.5pt; font-weight: bold; color: #000;\n';
  html += '  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\n';
  html += '  line-height: 1.4;\n';
  html += '}\n';
  html += '.right-line.reference { font-size: 11pt; }\n';
  html += '.right-line.country { font-size: 11pt; }\n';
  html += '.right-line.cert { font-size: 10pt; }\n';
  html += '.right-line.weight { font-size: 10.5pt; }\n';
  html += '.right-line.bags { font-size: 10.5pt; }\n';
  html += '.right-line.sampled { font-size: 10pt; }\n';
  html += '.right-line.cargo { font-size: 11pt; }\n';
  html += '.right-line.mark { font-size: 10pt; }\n';
  html += '.right-line.sample-order { font-size: 11pt; }\n';

  // Bottom section (spans both columns) — centered CS Sample + wide barcode
  html += '.bottom-section {\n';
  html += '  grid-column: 1 / -1;\n';
  html += '  display: flex;\n';
  html += '  flex-direction: column;\n';
  html += '  align-items: center;\n';
  html += '  border-top: 1px solid #ccc;\n';
  html += '  padding-top: 0.06in;\n';
  html += '  margin-top: 0.04in;\n';
  html += '}\n';
  html += '.cs-sample-text {\n';
  html += '  font-size: 14pt; font-weight: bold; color: #000;\n';
  html += '  text-align: center;\n';
  html += '  margin-bottom: 0.04in;\n';
  html += '}\n';
  html += '.barcode-box { width: 100%; height: 0.45in; }\n';
  html += '.barcode-box img { width: 100%; height: 100%; object-fit: fill; }\n';

  // Weight layout helpers
  html += '.weight-row { display: flex; justify-content: space-between; align-items: baseline; }\n';
  html += '.weight-val { }\n';
  html += '.weight-unit { margin-left: 0.3in; }\n';
  html += '.bags-row { display: flex; justify-content: space-between; align-items: baseline; }\n';

  html += '</style>\n</head>\n<body>\n';

  // Controls
  html += '<div class="no-print" style="text-align:center; margin-bottom:20px; font-family:Arial;">\n';
  html += '  <h3>☕ Keurig Labels — ' + labels.length + ' label(s)</h3>\n';
  html += '  <p style="color:#666; font-size:13px;">4" × 4" full-detail warehouse label</p>\n';
  html += '  <button onclick="window.print()" style="padding:10px 24px; font-size:15px; background:#4CAF50; color:white; border:none; border-radius:4px; cursor:pointer; margin:5px;">🖨️ Print Labels</button>\n';
  html += '  <button onclick="google.script.host.close()" style="padding:10px 24px; font-size:15px; background:#666; color:white; border:none; border-radius:4px; cursor:pointer; margin:5px;">Close</button>\n';
  html += '</div>\n';

  // ── Render each label ──
  for (var j = 0; j < labels.length; j++) {
    var lb = labels[j];
    var qrUrl = getQRCodeUrl(lb.qrData, 450);
    var barcodeUrl = 'https://barcode.tec-it.com/barcode.ashx?data=' + encodeURIComponent(lb.csSample) + '&code=Code128&translate-esc=on&hidehrt=True&dpi=200';

    html += '<div class="label">\n';

    // ── Left column ──
    html += '  <div class="left-col">\n';
    html += '    <div class="qr-box"><img src="' + qrUrl + '" alt="QR" /></div>\n';
    html += '    <div class="left-line arrived">Arrived&nbsp;&nbsp;&nbsp;&nbsp;' + lb.arrivedDate + '</div>\n';
    html += '    <div class="left-line warehouse">' + lb.warehouse + '</div>\n';
    html += '    <div class="left-line customer">' + lb.customer + '</div>\n';
    html += '    <div class="left-line desc">' + lb.description + '</div>\n';
    html += '    <div class="left-line container">' + lb.container + '</div>\n';
    html += '    <div class="css-branding">\n';
    html += '      <div class="css-name">Commodity<br>Sampler<br>Services</div>\n';
    html += '      <div class="css-logo">☕</div>\n';
    html += '    </div>\n';
    html += '  </div>\n';

    // ── Right column ──
    html += '  <div class="right-col">\n';
    html += '    <div class="right-line reference">' + lb.pNumber + '</div>\n';
    html += '    <div class="right-line country">' + lb.country + '</div>\n';
    html += '    <div class="right-line cert">' + lb.certification + '</div>\n';
    html += '    <div class="right-line weight">\n';
    html += '      <span class="weight-val">' + lb.weight + '</span>\n';
    html += '      <span class="weight-unit">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;lbs</span>\n';
    html += '    </div>\n';
    html += '    <div class="right-line bags">\n';
    html += '      <span>Bags&nbsp;&nbsp;' + lb.bagCount + '</span>\n';
    html += '      <span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + lb.tsYear + '</span>\n';
    html += '    </div>\n';
    html += '    <div class="right-line sampled">Sampled&nbsp;&nbsp;' + lb.sampledDate + '</div>\n';
    html += '    <div class="right-line cargo">' + lb.cargo + '</div>\n';
    html += '    <div class="right-line mark">' + lb.markFormatted + '</div>\n';
    html += '    <div class="right-line sample-order">' + lb.sampleOrder + '</div>\n';
    html += '  </div>\n';

    // ── Bottom section (full width) — CS Sample + wide barcode ──
    html += '  <div class="bottom-section">\n';
    html += '    <div class="cs-sample-text">' + lb.csSample + '</div>\n';
    html += '    <div class="barcode-box"><img src="' + barcodeUrl + '" alt="Barcode" /></div>\n';
    html += '  </div>\n';

    html += '</div>\n';
  }

  html += '</body>\n</html>';

  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(700);
  ui.showModalDialog(htmlOutput, 'Keurig Labels — ' + labels.length + ' label(s)');
}
function testQRUrl() {
  Logger.log('testQRUrl: Testing QR URL fetch');
  try {
    var url = getQRCodeUrl('TEST123', 350);
    Logger.log('URL: ' + url);
    var response = UrlFetchApp.fetch(url);
    Logger.log('Response code: ' + response.getResponseCode());
    Logger.log('Response body: ' + response.getContentText().substring(0, 200));
  } catch (e) {
    Logger.log('testQRUrl error: ' + e.toString());
  }
}