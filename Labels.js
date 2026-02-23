// ============================================================
// Labels.gs — Label Generation (4x6, 4x4, 3x2, 2x1), Print Checkboxes
// Commodity Sampler Services
//
// NOTE: getQRCodeUrl lives in QRcode.gs (quickchart.io version)
//       getBarcodeUrl is unique to this file
// ============================================================


// ============================================================
// BARCODE HELPER (unique to Labels.gs)
// ============================================================

function getBarcodeUrl(data) {
  var encoded = encodeURIComponent(data);
  return 'https://barcodeapi.org/api/128/' + encoded;
}

// ============================================================
// GENERATE LABELS WITH SIZE (Router)
// ============================================================

function generateLabelsWithSize(size) {
  var samples = _getSamplesFromCache(false);
  if (!samples) return _expiredHtml();
  switch(size) {
    case '3x2': return generateLabelsHtml3x2(samples);
    case '4x4': return generateLabelsHtml4x4(samples);
    case '2x1': return generateLabelsHtml2x1(samples);
    case '4x6': default: return generateLabelsHtml4x6(samples);
  }
}

function generateLabelsWithSizeNoCheckbox(size) {
  var samples = _getSamplesFromCache(true);
  if (!samples) return _expiredHtml();
  switch(size) {
    case '3x2': return generateLabelsHtml3x2NoCheckbox(samples);
    case '4x4': return generateLabelsHtml4x4NoCheckbox(samples);
    case '2x1': return generateLabelsHtml2x1NoCheckbox(samples);
    case '4x6': default: return generateLabelsHtml4x6NoCheckbox(samples);
  }
}

function _getSamplesFromCache(removeAfter) {
  var cache = CacheService.getUserCache();
  var json = cache.get('samplesToPrint');
  if (!json) return null;
  var samples = JSON.parse(json);
  if (removeAfter) cache.remove('samplesToPrint');
  return samples;
}

function _expiredHtml() {
  return '<p style="padding:20px;font-family:Arial;">Session expired. Please try again.</p>';
}

// ============================================================
// SHARED: clearAndClose script block
// ============================================================

var CLEAR_SCRIPT = `
  <script>
    function clearAndClose() {
      google.script.run.withSuccessHandler(function() {
        google.script.host.close();
      }).clearPrintCheckboxesFromCache();
    }
  </script>`;

// ============================================================
// 4x6 LABEL (Standard)
// ============================================================

function generateLabelsHtml4x6(samples) { return _build4x6(samples, true); }
function generateLabelsHtml4x6NoCheckbox(samples) { return _build4x6(samples, false); }

function _build4x6(samples, hasCheckbox) {
  var html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    @page { size: 4in 6in; margin: 0; }
    body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
    .label {
      width: 4in; height: 6in; padding: 0.15in; box-sizing: border-box;
      page-break-after: always; border: 1px solid #ccc;
      display: flex; flex-direction: column;
    }
    .label:last-child { page-break-after: auto; }
    .header { text-align: center; border-bottom: 2px solid #2E5339; padding-bottom: 6px; margin-bottom: 8px; }
    .company-name { font-size: 14pt; font-weight: bold; color: #2E5339; margin: 0; }
    .tagline { font-size: 9pt; font-style: italic; color: #666; margin: 2px 0; }
    .contact { font-size: 8pt; color: #333; }
    .content { flex: 1; }
    .address-row { display: flex; gap: 10px; margin-bottom: 8px; }
    .address-box { flex: 1; padding: 8px; border-radius: 4px; }
    .ship-from { background: #fff3cd; border: 2px solid #ffc107; }
    .ship-to { background: #e8f5e9; border: 2px solid #4A7C59; }
    .address-label { font-size: 9pt; font-weight: bold; color: #666; margin-bottom: 3px; }
    .address-value { font-size: 12pt; font-weight: bold; color: #000; }
    .warehouse-box { background: #e3f2fd; border: 2px solid #1976d2; padding: 8px; border-radius: 4px; margin-bottom: 8px; text-align: center; }
    .warehouse-label { font-size: 9pt; font-weight: bold; color: #666; }
    .warehouse-value { font-size: 13pt; font-weight: bold; color: #1976d2; }
    .shipping-method { text-align: center; font-size: 12pt; font-weight: bold; margin-bottom: 8px; padding: 6px; background: #f5f5f5; border-radius: 4px; }
    .field { font-size: 11pt; margin: 4px 0; }
    .field-label { font-weight: bold; display: inline-block; width: 90px; }
    .field-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 4px 12px; }
    .codes-row { display: flex; justify-content: space-between; align-items: center; margin-top: auto; padding-top: 8px; border-top: 1px solid #ddd; }
    .qr-code { width: 1in; height: 1in; }
    .barcode-container { flex: 1; text-align: center; padding-left: 10px; display: flex; flex-direction: column; align-items: center; justify-content: center; }
    .barcode { height: 0.7in; width: 100%; max-width: 2.6in; }
    .barcode-text { font-size: 9pt; font-weight: bold; margin-top: 2px; }
    .print-btn { display: block; margin: 20px auto; padding: 15px 40px; font-size: 16pt; background: #4A7C59; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .print-btn:hover { background: #3a6249; }
    .clear-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #dc3545; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .close-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #6c757d; color: white; border: none; cursor: pointer; border-radius: 5px; }
    @media print { .no-print { display: none; } body { margin: 0; } .label { border: none; } }
  </style>
</head>
<body>
  <div class="no-print" style="text-align: center; padding: 15px; background: #f0f0f0;">
    <button class="print-btn" onclick="window.print()">🖨️ Print ${samples.length} Labels (4×6)</button>
    ${hasCheckbox
      ? '<button class="clear-btn" onclick="clearAndClose()">Clear Checkboxes & Close</button>'
      : '<button class="close-btn" onclick="google.script.host.close()">Close</button>'}
  </div>
`;

  samples.forEach(function(s) {
    var qrData = 'CS:' + s.csSample + '|' + s.container + '|' + s.mark + '|' + s.cargo;
    var qrUrl = getQRCodeUrl(qrData, 120);
    var barcodeUrl = getBarcodeUrl(s.csSample);
    
    html += `
    <div class="label">
      <div class="header">
        <div class="company-name">COMMODITY SAMPLER SERVICES</div>
        <div class="tagline">Integrity Through Independence</div>
        <div class="contact">info@commoditysampler.com | 732-583-0666</div>
      </div>
      <div class="content">
        <div class="address-row">
          <div class="address-box ship-from">
            <div class="address-label">📤 SHIP FROM</div>
            <div class="address-value">${s.sender || 'N/A'}</div>
          </div>
          <div class="address-box ship-to">
            <div class="address-label">📦 SHIP TO</div>
            <div class="address-value">${s.receiver || 'N/A'}</div>
          </div>
        </div>
        <div class="warehouse-box">
          <div class="warehouse-label">🏭 WAREHOUSE</div>
          <div class="warehouse-value">${s.warehouse || 'N/A'}</div>
        </div>
        <div class="shipping-method">🚚 ${s.shippingProcess || 'Standard Shipping'}</div>
        <div class="field-grid">
          <div class="field"><span class="field-label">Container:</span> ${s.container || '-'}</div>
          <div class="field"><span class="field-label">Mark:</span> ${s.mark || '-'}</div>
          <div class="field"><span class="field-label">Cargo:</span> ${s.cargo || '-'}</div>
          <div class="field"><span class="field-label">Reference:</span> ${s.reference || '-'}</div>
          <div class="field"><span class="field-label">Description:</span> ${s.description || '-'}</div>
          <div class="field"><span class="field-label">Bags:</span> ${s.bagCount || '-'}</div>
          <div class="field"><span class="field-label">Sample Wt:</span> ${s.sampleWeight || '-'}</div>
          <div class="field"><span class="field-label">Order:</span> ${s.sampleOrderNum || '-'}</div>
        </div>
      </div>
      <div class="codes-row">
        <img src="${qrUrl}" class="qr-code" alt="QR">
        <div class="barcode-container">
          <img src="${barcodeUrl}" class="barcode" alt="Barcode">
          <div class="barcode-text">${s.csSample}</div>
        </div>
      </div>
    </div>
`;
  });

  html += (hasCheckbox ? CLEAR_SCRIPT : '') + `</body></html>`;
  return html;
}

// ============================================================
// 4x4 LABEL (Compact)
// ============================================================

function generateLabelsHtml4x4(samples) { return _build4x4(samples, true); }
function generateLabelsHtml4x4NoCheckbox(samples) { return _build4x4(samples, false); }

function _build4x4(samples, hasCheckbox) {
  var html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    @page { size: 4in 4in; margin: 0; }
    body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
    .label {
      width: 4in; height: 4in; padding: 0.1in; box-sizing: border-box;
      page-break-after: always; border: 1px solid #ccc;
      display: flex; flex-direction: column;
    }
    .label:last-child { page-break-after: auto; }
    .header { text-align: center; border-bottom: 2px solid #2E5339; padding-bottom: 4px; margin-bottom: 6px; }
    .company-name { font-size: 11pt; font-weight: bold; color: #2E5339; margin: 0; }
    .tagline { font-size: 7pt; font-style: italic; color: #666; margin: 1px 0; }
    .contact { font-size: 7pt; color: #333; }
    .content { flex: 1; }
    .address-row { display: flex; gap: 6px; margin-bottom: 5px; }
    .address-box { flex: 1; padding: 5px; border-radius: 3px; }
    .ship-from { background: #fff3cd; border: 2px solid #ffc107; }
    .ship-to { background: #e8f5e9; border: 2px solid #4A7C59; }
    .address-label { font-size: 7pt; font-weight: bold; color: #666; margin-bottom: 2px; }
    .address-value { font-size: 10pt; font-weight: bold; color: #000; }
    .warehouse-box { background: #e3f2fd; border: 2px solid #1976d2; padding: 5px; border-radius: 3px; margin-bottom: 5px; text-align: center; }
    .warehouse-label { font-size: 7pt; font-weight: bold; color: #666; }
    .warehouse-value { font-size: 10pt; font-weight: bold; color: #1976d2; }
    .shipping-method { text-align: center; font-size: 9pt; font-weight: bold; margin-bottom: 5px; padding: 4px; background: #f5f5f5; border-radius: 3px; }
    .field { font-size: 9pt; margin: 2px 0; }
    .field-label { font-weight: bold; display: inline-block; width: 75px; }
    .field-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 2px 8px; }
    .codes-row { display: flex; justify-content: space-between; align-items: center; margin-top: auto; padding-top: 5px; border-top: 1px solid #ddd; }
    .qr-code { width: 0.85in; height: 0.85in; }
    .barcode-container { flex: 1; text-align: center; padding-left: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; }
    .barcode { height: 0.55in; width: 100%; max-width: 2.2in; }
    .barcode-text { font-size: 9pt; font-weight: bold; margin-top: 2px; }
    .print-btn { display: block; margin: 20px auto; padding: 15px 40px; font-size: 16pt; background: #4A7C59; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .clear-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #dc3545; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .close-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #6c757d; color: white; border: none; cursor: pointer; border-radius: 5px; }
    @media print { .no-print { display: none; } .label { border: none; } }
  </style>
</head>
<body>
  <div class="no-print" style="text-align: center; padding: 15px; background: #f0f0f0;">
    <button class="print-btn" onclick="window.print()">🖨️ Print ${samples.length} Labels (4×4)</button>
    ${hasCheckbox
      ? '<button class="clear-btn" onclick="clearAndClose()">Clear Checkboxes & Close</button>'
      : '<button class="close-btn" onclick="google.script.host.close()">Close</button>'}
  </div>
`;

  samples.forEach(function(s) {
    var qrData = 'CS:' + s.csSample + '|' + s.container + '|' + s.mark + '|' + s.cargo;
    var qrUrl = getQRCodeUrl(qrData, 100);
    var barcodeUrl = getBarcodeUrl(s.csSample);
    
    html += `
    <div class="label">
      <div class="header">
        <div class="company-name">COMMODITY SAMPLER SERVICES</div>
        <div class="tagline">Integrity Through Independence</div>
        <div class="contact">info@commoditysampler.com | 732-583-0666</div>
      </div>
      <div class="content">
        <div class="address-row">
          <div class="address-box ship-from">
            <div class="address-label">📤 FROM</div>
            <div class="address-value">${s.sender || 'N/A'}</div>
          </div>
          <div class="address-box ship-to">
            <div class="address-label">📦 TO</div>
            <div class="address-value">${s.receiver || 'N/A'}</div>
          </div>
        </div>
        <div class="warehouse-box">
          <div class="warehouse-label">🏭 WAREHOUSE</div>
          <div class="warehouse-value">${s.warehouse || 'N/A'}</div>
        </div>
        <div class="shipping-method">🚚 ${s.shippingProcess || 'Standard'}</div>
        <div class="field-grid">
          <div class="field"><span class="field-label">Container:</span> ${s.container || '-'}</div>
          <div class="field"><span class="field-label">Mark:</span> ${s.mark || '-'}</div>
          <div class="field"><span class="field-label">Cargo:</span> ${s.cargo || '-'}</div>
          <div class="field"><span class="field-label">Bags:</span> ${s.bagCount || '-'}</div>
        </div>
      </div>
      <div class="codes-row">
        <img src="${qrUrl}" class="qr-code" alt="QR">
        <div class="barcode-container">
          <img src="${barcodeUrl}" class="barcode" alt="Barcode">
          <div class="barcode-text">${s.csSample}</div>
        </div>
      </div>
    </div>
`;
  });

  html += (hasCheckbox ? CLEAR_SCRIPT : '') + `</body></html>`;
  return html;
}

// ============================================================
// 3x2 LABEL (Mid)
// ============================================================

function generateLabelsHtml3x2(samples) { return _build3x2(samples, true); }
function generateLabelsHtml3x2NoCheckbox(samples) { return _build3x2(samples, false); }

function _build3x2(samples, hasCheckbox) {
  var html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    @page { size: 3in 2in; margin: 0; }
    body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
    .label {
      width: 3in; height: 2in; padding: 0.08in; box-sizing: border-box;
      page-break-after: always; border: 1px solid #ccc;
      display: flex; flex-direction: column;
    }
    .label:last-child { page-break-after: auto; }
    .header { text-align: center; font-size: 8pt; font-weight: bold; color: #2E5339; border-bottom: 1px solid #2E5339; padding-bottom: 3px; margin-bottom: 4px; }
    .tagline { font-size: 6pt; font-style: italic; color: #666; }
    .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 2px 8px; font-size: 9pt; margin-bottom: 4px; }
    .info-item { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .info-label { font-weight: bold; color: #333; }
    .address-row { display: flex; gap: 6px; margin-bottom: 4px; }
    .address-box { flex: 1; padding: 3px 5px; border-radius: 3px; font-size: 8pt; }
    .ship-from { background: #fff3cd; border-left: 2px solid #ffc107; }
    .ship-to { background: #e8f5e9; border-left: 2px solid #4A7C59; }
    .address-label { font-size: 6pt; font-weight: bold; color: #666; }
    .address-value { font-size: 9pt; font-weight: bold; }
    .warehouse { text-align: center; font-size: 8pt; color: #1976d2; font-weight: bold; margin-bottom: 3px; }
    .barcode-container { text-align: center; flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; }
    .barcode { height: 0.45in; max-width: 2.8in; }
    .barcode-text { font-size: 7pt; font-weight: bold; margin-top: 1px; }
    .print-btn { display: block; margin: 20px auto; padding: 15px 40px; font-size: 16pt; background: #4A7C59; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .clear-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #dc3545; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .close-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #6c757d; color: white; border: none; cursor: pointer; border-radius: 5px; }
    @media print { .no-print { display: none; } .label { border: none; } }
  </style>
</head>
<body>
  <div class="no-print" style="text-align: center; padding: 15px; background: #f0f0f0;">
    <button class="print-btn" onclick="window.print()">🖨️ Print ${samples.length} Labels (3×2)</button>
    ${hasCheckbox
      ? '<button class="clear-btn" onclick="clearAndClose()">Clear Checkboxes & Close</button>'
      : '<button class="close-btn" onclick="google.script.host.close()">Close</button>'}
  </div>
`;

  samples.forEach(function(s) {
    var barcodeUrl = getBarcodeUrl(s.csSample);
    var containerShort = (s.container || '-').substring(0, 18);
    var markShort = (s.mark || '-').substring(0, 18);
    var cargoShort = (s.cargo || '-').substring(0, 18);
    var senderShort = (s.sender || '-').substring(0, 18);
    var receiverShort = (s.receiver || '-').substring(0, 18);
    var warehouseShort = (s.warehouse || '-').substring(0, 20);
    
    html += `
    <div class="label">
      <div class="header">
        COMMODITY SAMPLER SERVICES
        <div class="tagline">Integrity Through Independence</div>
      </div>
      <div class="address-row">
        <div class="address-box ship-from">
          <div class="address-label">📤 FROM</div>
          <div class="address-value">${senderShort}</div>
        </div>
        <div class="address-box ship-to">
          <div class="address-label">📦 TO</div>
          <div class="address-value">${receiverShort}</div>
        </div>
      </div>
      <div class="warehouse">🏭 ${warehouseShort}</div>
      <div class="info-grid">
        <div class="info-item"><span class="info-label">C:</span> ${containerShort}</div>
        <div class="info-item"><span class="info-label">M:</span> ${markShort}</div>
        <div class="info-item"><span class="info-label">Cargo:</span> ${cargoShort}</div>
        <div class="info-item"><span class="info-label">Bags:</span> ${s.bagCount || '-'}</div>
      </div>
      <div class="barcode-container">
        <img src="${barcodeUrl}" class="barcode" alt="Barcode">
        <div class="barcode-text">${s.csSample}</div>
      </div>
    </div>
`;
  });

  html += (hasCheckbox ? CLEAR_SCRIPT : '') + `</body></html>`;
  return html;
}

// ============================================================
// 2x1 LABEL (Mini)
// ============================================================

function generateLabelsHtml2x1(samples) { return _build2x1(samples, true); }
function generateLabelsHtml2x1NoCheckbox(samples) { return _build2x1(samples, false); }

function _build2x1(samples, hasCheckbox) {
  var html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    @page { size: 2in 1in; margin: 0; }
    body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
    .label {
      width: 2in; height: 1in; padding: 0.03in; box-sizing: border-box;
      page-break-after: always; border: 1px solid #ccc;
      display: flex; flex-direction: column;
    }
    .label:last-child { page-break-after: auto; }
    .header { text-align: center; font-size: 5pt; font-weight: bold; color: #2E5339; margin-bottom: 1px; }
    .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0px 4px; font-size: 6pt; margin-bottom: 2px; }
    .info-item { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .info-label { font-weight: bold; }
    .barcode-container { text-align: center; flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; }
    .barcode { height: 0.3in; max-width: 1.9in; }
    .barcode-text { font-size: 5pt; font-weight: bold; }
    .print-btn { display: block; margin: 20px auto; padding: 15px 40px; font-size: 16pt; background: #4A7C59; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .clear-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #dc3545; color: white; border: none; cursor: pointer; border-radius: 5px; }
    .close-btn { display: block; margin: 10px auto; padding: 10px 30px; font-size: 12pt; background: #6c757d; color: white; border: none; cursor: pointer; border-radius: 5px; }
    @media print { .no-print { display: none; } .label { border: none; } }
  </style>
</head>
<body>
  <div class="no-print" style="text-align: center; padding: 15px; background: #f0f0f0;">
    <button class="print-btn" onclick="window.print()">🖨️ Print ${samples.length} Labels (2×1)</button>
    ${hasCheckbox
      ? '<button class="clear-btn" onclick="clearAndClose()">Clear Checkboxes & Close</button>'
      : '<button class="close-btn" onclick="google.script.host.close()">Close</button>'}
  </div>
`;

  samples.forEach(function(s) {
    var barcodeUrl = getBarcodeUrl(s.csSample);
    var containerShort = (s.container || '-').substring(0, 14);
    var markShort = (s.mark || '-').substring(0, 14);
    var cargoShort = (s.cargo || '-').substring(0, 14);
    var warehouseShort = (s.warehouse || '-').substring(0, 10);
    
    html += `
    <div class="label">
      <div class="header">CSS | ${warehouseShort}</div>
      <div class="info-grid">
        <div class="info-item"><span class="info-label">C:</span> ${containerShort}</div>
        <div class="info-item"><span class="info-label">M:</span> ${markShort}</div>
        <div class="info-item"><span class="info-label">Cargo:</span> ${cargoShort}</div>
        <div class="info-item"><span class="info-label">📦</span> ${(s.receiver || '-').substring(0, 12)}</div>
      </div>
      <div class="barcode-container">
        <img src="${barcodeUrl}" class="barcode" alt="Barcode">
        <div class="barcode-text">${s.csSample}</div>
      </div>
    </div>
`;
  });

  html += (hasCheckbox ? CLEAR_SCRIPT : '') + `</body></html>`;
  return html;
}

// ============================================================
// CLEAR PRINT CHECKBOXES
// ============================================================

function clearPrintCheckboxesFromCache() {
  var cache = CacheService.getUserCache();
  var samplesJson = cache.get('samplesToPrint');
  if (!samplesJson) return;
  
  var samples = JSON.parse(samplesJson);
  clearPrintCheckboxes(samples);
  cache.remove('samplesToPrint');
}

function clearPrintCheckboxes(samples) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGroups = {};
  samples.forEach(function(s) {
    if (!sheetGroups[s.sheet]) sheetGroups[s.sheet] = [];
    sheetGroups[s.sheet].push(s.row);
  });
  
  for (var sheetName in sheetGroups) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    var col = _getColumnMap(sheet);
    var printIdx = col['Print'];
    if (printIdx === undefined) continue;
    sheetGroups[sheetName].forEach(function(row) {
      sheet.getRange(row, printIdx + 1).setValue(false);
    });
  }
}

// ============================================================
// PRINT CHECKED SAMPLES - Entry Point
// ============================================================

function printCheckedSamples() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var col = _getColumnMap(sheet);
  
  var printIdx = col['Print'];
  if (printIdx === undefined) {
    SpreadsheetApp.getUi().alert('No "Print" column found on this sheet.');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var samples = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][printIdx] === true) {
      samples.push({
        row: i + 1,
        sheet: sheetName,
        csOrder: data[i][col['CS Order #']] || '',
        csSample: data[i][col['CS Sample #']] || '',
        sender: data[i][col['Sender']] || '',
        receiver: data[i][col['Receiver']] || '',
        warehouse: data[i][col['Warehouse']] || '',
        description: data[i][col['Description']] || '',
        sampleOrderNum: data[i][col['Sample Order #']] || '',
        cargo: data[i][col['Cargo #']] || '',
        mark: data[i][col['Mark #']] || '',
        container: data[i][col['Container #']] || '',
        reference: data[i][col['Reference']] || '',
        bagCount: data[i][col['Bag Count']] || '',
        weight: data[i][col['Weight']] || '',
        sampleWeight: data[i][col['Sample Weight']] || '',
        shippingProcess: data[i][col['Shipping Process']] || '',
        comments: data[i][col['Comments']] || ''
      });
    }
  }
  
  if (samples.length === 0) {
    SpreadsheetApp.getUi().alert('No samples checked for printing. Check the Print column for rows you want to print.');
    return;
  }
  
  var cache = CacheService.getUserCache();
  cache.put('samplesToPrint', JSON.stringify(samples), 300);
  
  var html = HtmlService.createHtmlOutput(generateLabelSizeChooser(samples.length))
    .setWidth(400)
    .setHeight(300)
    .setTitle('Print Labels');
  SpreadsheetApp.getUi().showModalDialog(html, 'Print ' + samples.length + ' Labels');
}

function generateLabelSizeChooser(count) {
  return `
    <div style="font-family:Arial;text-align:center;padding:20px;">
      <h2>Print ${count} Label${count > 1 ? 's' : ''}</h2>
      <p>Choose label size:</p>
      <button onclick="pickSize('4x6')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">4×6 Standard</button><br>
      <button onclick="pickSize('4x4')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">4×4 Compact</button><br>
      <button onclick="pickSize('3x2')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">3×2 Mid</button><br>
      <button onclick="pickSize('2x1')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">2×1 Mini</button>
    </div>
    <script>
      function pickSize(size) {
        google.script.run.withSuccessHandler(function(html) {
          document.open();
          document.write(html);
          document.close();
        }).generateLabelsWithSize(size);
      }
    </script>`;
}

// ============================================================
// GENERATE LABELS (from selected rows)
// ============================================================

function generateLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow === 1) { 
    ui.alert('Please select data rows (not header).'); 
    return; 
  }
  
  var col = _getColumnMap(sheet);
  const samples = [];
  var allRows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  for (let i = 0; i < numRows; i++) {
    const rowData = allRows[i];

    samples.push({
      row: startRow + i,
      sheet: sheet.getName(),
      csOrder: rowData[col['CS Order #']] || '',
      csSample: rowData[col['CS Sample #']] || '',
      sender: rowData[col['Sender']] || '',
      receiver: rowData[col['Receiver']] || '',
      warehouse: rowData[col['Warehouse']] || '',
      description: rowData[col['Description']] || '',
      sampleOrderNum: rowData[col['Sample Order #']] || '',
      cargo: rowData[col['Cargo #']] || '',
      mark: rowData[col['Mark #']] || '',
      container: rowData[col['Container #']] || '',
      reference: rowData[col['Reference']] || '',
      bagCount: rowData[col['Bag Count']] || '',
      weight: rowData[col['Weight']] || '',
      sampleWeight: rowData[col['Sample Weight']] || '',
      shippingProcess: rowData[col['Shipping Process']] || '',
      comments: rowData[col['Comments']] || ''
    });
  }
  
  if (samples.length === 0) {
    ui.alert('No rows selected');
    return;
  }
  
  var cache = CacheService.getUserCache();
  cache.put('samplesToPrint', JSON.stringify(samples), 300);
  
  var pickerHtml = generateLabelSizePickerNoCheckbox(samples.length);
  var htmlOutput = HtmlService.createHtmlOutput(pickerHtml)
    .setWidth(400)
    .setHeight(320);
  ui.showModalDialog(htmlOutput, 'Select Label Size - ' + samples.length + ' Labels');
}

function generateLabelSizePickerNoCheckbox(count) {
  return `
    <div style="font-family:Arial;text-align:center;padding:20px;">
      <h2>Print ${count} Label${count > 1 ? 's' : ''}</h2>
      <p>Choose label size:</p>
      <button onclick="pickSize('4x6')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">4×6 Standard</button><br>
      <button onclick="pickSize('4x4')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">4×4 Compact</button><br>
      <button onclick="pickSize('3x2')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">3×2 Mid</button><br>
      <button onclick="pickSize('2x1')" style="margin:8px;padding:12px 24px;font-size:14pt;background:#4A7C59;color:white;border:none;border-radius:5px;cursor:pointer;">2×1 Mini</button>
    </div>
    <script>
      function pickSize(size) {
        google.script.run.withSuccessHandler(function(html) {
          document.open();
          document.write(html);
          document.close();
        }).generateLabelsWithSizeNoCheckbox(size);
      }
    </script>`;
}

// ============================================================
// REPRINT LABEL
// ============================================================

function reprintLabel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const row = sheet.getActiveCell().getRow();
  
  if (row < 2) {
    ui.alert('Please select a data row');
    return;
  }
  
  var col = _getColumnMap(sheet);
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var sample = {
    row: row,
    sheet: sheet.getName(),
    csOrder: data[col['CS Order #']] || '',
    csSample: data[col['CS Sample #']] || '',
    sender: data[col['Sender']] || '',
    receiver: data[col['Receiver']] || '',
    warehouse: data[col['Warehouse']] || '',
    description: data[col['Description']] || '',
    sampleOrderNum: data[col['Sample Order #']] || '',
    cargo: data[col['Cargo #']] || '',
    mark: data[col['Mark #']] || '',
    container: data[col['Container #']] || '',
    reference: data[col['Reference']] || '',
    bagCount: data[col['Bag Count']] || '',
    sampleWeight: data[col['Sample Weight']] || '',
    shippingProcess: data[col['Shipping Process']] || ''
  };
  
  if (!sample.csSample) {
    ui.alert('No sample number found in this row');
    return;
  }
  
  var cache = CacheService.getUserCache();
  cache.put('samplesToPrint', JSON.stringify([sample]), 300);
  
  var pickerHtml = generateLabelSizePickerNoCheckbox(1);
  var htmlOutput = HtmlService.createHtmlOutput(pickerHtml)
    .setWidth(400)
    .setHeight(350);
  ui.showModalDialog(htmlOutput, 'Reprint Label: ' + sample.csSample);
}

// ============================================================
// FIX ALL PRINT CHECKBOXES
// ============================================================

function fixAllPrintCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheets = [
    CONFIG.mainSheetName,
    CONFIG.liveOrdersSheetName,
    CONFIG.completedOrdersSheetName || 'Completed Orders',
    'View - Danboy1217'
  ];
  
  sheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var printIdx = headers.indexOf('Print');
    
    if (printIdx === -1) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, printIdx + 1, lastRow - 1, 1).insertCheckboxes();
    }
  });
  
  SpreadsheetApp.getUi().alert('✅ Print checkboxes fixed on all sheets');
}
