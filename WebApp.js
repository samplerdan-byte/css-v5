// ============================================================
// 🌐 WEB APP — doGet + Server-Side Handlers
// ============================================================
// Deployed web app for mobile/tablet use in the field
// Features: QR scan to input, drag-and-drop PDF, manual entry
// 
// Deploy: Deploy → New Deployment → Web App → Anyone with link
// ============================================================

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : 'main';
  var action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : '';

  // Validate page parameter against allowlist
  var validPages = ['main', 'scanner', 'portal'];
  if (validPages.indexOf(page) === -1) page = 'main';

  // Validate action parameter against allowlist
  var validActions = ['', 'import', 'refresh', 'process'];
  if (validActions.indexOf(action) === -1) action = '';

  // Remote action triggers (called via URL) — require token auth
  if (action === 'import' || action === 'refresh' || action === 'process') {
    var actionToken = (e && e.parameter && e.parameter.token) ? e.parameter.token : '';
    var expectedToken = PropertiesService.getScriptProperties().getProperty('WEBAPP_ACTION_TOKEN');
    if (!expectedToken || actionToken !== expectedToken) {
      return ContentService.createTextOutput('Unauthorized — invalid or missing token')
        .setMimeType(ContentService.MimeType.TEXT);
    }
  }
  if (action === 'import') {
    // Clear old state and create trigger for fresh historical import
    _removeHistoricalTrigger();
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('hist_running');
    props.deleteProperty('hist_queryIndex');
    props.deleteProperty('hist_offset');
    props.deleteProperty('hist_totalProcessed');
    props.deleteProperty('hist_totalSkipped');
    props.deleteProperty('hist_totalThreads');
    ScriptApp.newTrigger('historicalImport').timeBased().after(5000).create();
    return ContentService.createTextOutput('Historical import trigger created — starts in ~5 seconds')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  if (action === 'refresh') {
    ScriptApp.newTrigger('fullRefresh').timeBased().after(5000).create();
    return ContentService.createTextOutput('Full refresh trigger created — starts in ~5 seconds')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  if (action === 'process') {
    ScriptApp.newTrigger('processPDFsFromGmail').timeBased().after(5000).create();
    return ContentService.createTextOutput('Email processing trigger created — starts in ~5 seconds')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // Client Status Portal — token-based access
  if (page === 'portal') {
    try {
      return HtmlService.createTemplateFromFile('ClientPortal')
        .evaluate()
        .setTitle('CSS — Order Status Portal')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (templateErr) {
      Logger.log('doGet: ClientPortal template missing or failed to load');
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family:Arial;padding:40px;text-align:center;">' +
        '<h2 style="color:#c0392b;">Portal Unavailable</h2>' +
        '<p>The portal is temporarily unavailable. Please try again later.</p>' +
        '</body></html>'
      ).setTitle('CSS — Portal Unavailable');
    }
  }

  // Legacy scanner URL
  if (page === 'scanner') {
    try {
      return HtmlService.createTemplateFromFile('WebAppScanner')
        .evaluate()
        .setTitle('CSS Scanner')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (templateErr) {
      Logger.log('doGet: WebAppScanner template missing or failed to load');
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family:Arial;padding:40px;text-align:center;">' +
        '<h2 style="color:#c0392b;">Scanner Unavailable</h2>' +
        '<p>The scanner is temporarily unavailable. Please try again later.</p>' +
        '</body></html>'
      ).setTitle('CSS — Scanner Unavailable');
    }
  }

  // Default: full web app
  try {
    return HtmlService.createTemplateFromFile('WebAppScanner')
      .evaluate()
      .setTitle('CSS V5 — Commodity Sampler Services')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (templateErr) {
    Logger.log('doGet: WebAppScanner (default) template missing or failed to load');
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:Arial;padding:40px;text-align:center;">' +
      '<h2 style="color:#2E5339;">CSS — Commodity Sampler Services</h2>' +
      '<p>The application is temporarily unavailable. Please try again later.</p>' +
      '</body></html>'
    ).setTitle('CSS — Commodity Sampler Services');
  }
}
// ============================================================
// SERVER-SIDE: Process scanned QR data (order entry / lookup)
// ============================================================

function webAppProcessQR(qrText) {
  try {
    Logger.log('WebApp QR scan: ' + String(qrText).substring(0, 200).replace(/[\r\n]/g, ' '));
    
    // Parse QR data — CSS labels use pipe-delimited format
    // ORDER:xxx|CARGO:xxx|MARK:xxx|CONTAINER:xxx|REF:xxx|DESC:xxx|BAGS:xxx|WEIGHT:xxx|SAMPLE:xxx|P:xxx|S:xxx|WAREHOUSE:xxx|CS_SAMPLE:xxx
    var data = {};
    if (qrText.indexOf('|') !== -1) {
      qrText.split('|').forEach(function(part) {
        var colonIdx = part.indexOf(':');
        if (colonIdx !== -1) {
          var key = part.substring(0, colonIdx).trim();
          var val = part.substring(colonIdx + 1).trim();
          data[key] = val;
        }
      });
    }
    
    // If it's a CSS QR code with CS_SAMPLE, look up the order
    if (data['CS_SAMPLE']) {
      return _webLookupSample(data['CS_SAMPLE']);
    }
    
    // If it's a plain sample number (e.g., "211700-03")
    if (/^\d{5,}-\d{1,3}$/.test(qrText.trim())) {
      return _webLookupSample(qrText.trim());
    }
    
    // If it has ORDER data, it could be a new sample to process
    if (data['ORDER'] || data['CARGO'] || data['CONTAINER']) {
      return {
        success: true,
        type: 'parsed_qr',
        data: {
          sampleOrderNum: data['ORDER'] || '',
          cargo: data['CARGO'] || '',
          mark: data['MARK'] || '',
          container: data['CONTAINER'] || '',
          reference: data['REF'] || '',
          description: data['DESC'] || '',
          bagCount: data['BAGS'] || '',
          weight: data['WEIGHT'] || '',
          sampleWeight: data['SAMPLE'] || '',
          pNumber: data['P'] || '',
          sNumber: data['S'] || '',
          warehouse: data['WAREHOUSE'] || '',
          csSample: data['CS_SAMPLE'] || ''
        },
        message: '✅ QR parsed — review and submit'
      };
    }
    
    // Generic barcode — try as sample number
    return _webLookupSample(qrText.trim());
    
  } catch (e) {
    Logger.log('WebApp QR error: ' + String(e).substring(0, 500).replace(/[\r\n]/g, ' '));
    return { success: false, message: '❌ Error processing QR code. Please try again.' };
  }
}

function _webLookupSample(sampleId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [
    ss.getSheetByName(CONFIG.mainSheetName),
    ss.getSheetByName(CONFIG.liveOrdersSheetName),
    ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders')
  ];
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (!sheet || sheet.getLastRow() < 2) continue;
    
    var col = _getColumnMap(sheet);
    var csSampleIdx = col['CS Sample #'];
    if (csSampleIdx === undefined) continue;
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][csSampleIdx]).trim() === sampleId) {
        return {
          success: true,
          type: 'found',
          data: {
            csSample: sampleId,
            csOrder: String(data[i][col['CS Order #']] || ''),
            sender: String(data[i][col['Sender']] || ''),
            receiver: String(data[i][col['Receiver']] || ''),
            warehouse: String(data[i][col['Warehouse']] || ''),
            description: String(data[i][col['Description']] || ''),
            sampleOrderNum: String(data[i][col['Sample Order #']] || ''),
            cargo: String(data[i][col['Cargo #']] || ''),
            mark: String(data[i][col['Mark #']] || ''),
            container: String(data[i][col['Container #']] || ''),
            reference: String(data[i][col['Reference']] || ''),
            bagCount: String(data[i][col['Bag Count']] || ''),
            sampleWeight: String(data[i][col['Sample Weight']] || ''),
            status: String(data[i][col['Status']] || ''),
            trackingNumber: col['Tracking Number'] !== undefined ? String(data[i][col['Tracking Number']] || '') : ''
          },
          message: '✅ Found: ' + sampleId + ' — ' + String(data[i][col['Sender']] || '') + ' / ' + String(data[i][col['Description']] || '')
        };
      }
    }
  }
  
  return { success: false, type: 'not_found', message: '❌ Sample ' + sampleId + ' not found' };
}

// ============================================================
// SERVER-SIDE: Scan to mark as Received (barcode receiving)
// ============================================================

function scanSampleBarcode(barcode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
    if (!mainSheet) return { success: false, message: 'Sheet not found' };
    
    var col = _getColumnMap(mainSheet);
    var lastRow = mainSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No data' };
    
    var sampleIdx = col['CS Sample #'];
    var statusIdx = col['Status'];
    var scanDateIdx = col['Scanned Date'];
    var scanByIdx = col['Scanned By'];
    
    if (sampleIdx === undefined || statusIdx === undefined) {
      return { success: false, message: 'Required columns not found' };
    }
    
    // Parse QR data if pipe-delimited
    var searchId = barcode;
    if (barcode.indexOf('|') !== -1) {
      var parts = {};
      barcode.split('|').forEach(function(p) {
        var ci = p.indexOf(':');
        if (ci !== -1) parts[p.substring(0, ci).trim()] = p.substring(ci + 1).trim();
      });
      searchId = parts['CS_SAMPLE'] || barcode;
    }
    searchId = String(searchId).trim();
    
    // Search
    var data = mainSheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][sampleIdx]).trim() === searchId) {
        var currentStatus = String(data[i][statusIdx]).trim();
        
        if (currentStatus === 'Scanned') {
          return { success: false, message: '⚠️ ' + searchId + ' already scanned' };
        }
        if (currentStatus === 'Shipped' || currentStatus === 'Completed') {
          return { success: false, message: '⚠️ ' + searchId + ' already ' + currentStatus };
        }
        
        // Mark as Scanned
        var rowNum = i + 2;
        mainSheet.getRange(rowNum, statusIdx + 1).setValue(CONFIG.statusValues.SCANNED);
        if (scanDateIdx !== undefined) mainSheet.getRange(rowNum, scanDateIdx + 1).setValue(new Date());
        if (scanByIdx !== undefined) {
          try { mainSheet.getRange(rowNum, scanByIdx + 1).setValue(Session.getActiveUser().getEmail()); } catch(e) {}
        }
        
        try { updateLiveOrdersView(); } catch(e) {}
        
        return {
          success: true,
          message: '✅ ' + searchId + ' scanned in — ' + String(data[i][col['Description']] || '') + ' / ' + String(data[i][col['Sender']] || '')
        };
      }
    }
    
    return { success: false, message: '❌ ' + searchId + ' not found in system' };
  } catch (e) {
    Logger.log('scanSampleBarcode error: ' + String(e).substring(0, 500).replace(/[\r\n]/g, ' '));
    return { success: false, message: '❌ Error processing barcode. Please try again.' };
  }
}

// ============================================================
// SERVER-SIDE: Submit manual order from web app
// ============================================================

function webAppSubmitOrder(orderData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.mainSheetName);
    if (!sheet) return { success: false, message: 'Main sheet not found' };
    
    // Validate minimum data
    if (!orderData.sender && !orderData.sampleOrderNum && !orderData.container) {
      return { success: false, message: 'Need at least a sender, order #, or container #' };
    }
    
    addDataToSheet(sheet, {
      sender: orderData.sender || '',
      receiver: orderData.receiver || '',
      warehouse: orderData.warehouse || '',
      description: orderData.description || '',
      sampleOrderNum: orderData.sampleOrderNum || '',
      cargo: orderData.cargo || '',
      mark: orderData.mark || '',
      container: orderData.container || '',
      reference: orderData.reference || '',
      bagCount: orderData.bagCount || '',
      weight: orderData.weight || '',
      sampleWeight: orderData.sampleWeight || '',
      pNumber: orderData.pNumber || '',
      sNumber: orderData.sNumber || '',
      shippingProcess: orderData.shippingProcess || '',
      comments: orderData.comments || 'Entered via Web App',
      sourceEmail: 'Web App',
      sampleType: orderData.sampleType || ''
    });
    
    return { success: true, message: '✅ Order submitted successfully' };
  } catch (e) {
    Logger.log('webAppSubmitOrder error: ' + String(e).substring(0, 500).replace(/[\r\n]/g, ' '));
    return { success: false, message: '❌ Error submitting order. Please try again.' };
  }
}

// ============================================================
// SERVER-SIDE: Process dropped PDF (base64)
// ============================================================

function webAppProcessPDF(base64Data, fileName) {
  try {
    // Sanitize filename — strip path separators, control characters, limit length
    // Use the sanitized version everywhere (blob, Drive title, logging)
    var safeFileName = String(fileName || 'upload.pdf')
      .replace(/[\r\n\t]/g, '')          // no line breaks
      .replace(/[\/\\<>:"|?*]/g, '_')    // no path/shell special chars
      .replace(/\.\./g, '_')             // no directory traversal
      .substring(0, 100)                 // length cap
      .trim() || 'upload.pdf';
    Logger.log('WebApp PDF drop: ' + safeFileName + ' (' + (base64Data ? base64Data.length : 0) + ' chars)');

    // Basic size guard — Apps Script base64: ~4MB PDF ≈ ~5.5M chars
    if (!base64Data || typeof base64Data !== 'string' || base64Data.length > 8000000) {
      return { success: false, message: '⚠️ PDF too large or invalid. Try a smaller file.' };
    }

    // Decode base64 to blob
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'application/pdf', safeFileName);

    // Extract text from PDF
    // Google Apps Script can create a temp file and use OCR via Drive
    var tempFile = DriveApp.createFile(blob);
    var docFile = null;
    var text = '';

    // Convert PDF to Google Doc for text extraction
    try {
      docFile = Drive.Files.copy(
        { title: 'CSS_Temp_Extract_' + safeFileName, mimeType: MimeType.GOOGLE_DOCS },
        tempFile.getId(),
        { ocr: true, ocrLanguage: 'en' }
      );

      var doc = DocumentApp.openById(docFile.id);
      text = doc.getBody().getText();
    } finally {
      // Clean up temp files even if extraction fails
      try { if (docFile) DriveApp.getFileById(docFile.id).setTrashed(true); } catch(ignore) {}
      try { tempFile.setTrashed(true); } catch(ignore) {}
    }
    
    if (!text || text.trim().length === 0) {
      return { success: false, message: '⚠️ Could not extract text from PDF. Try manual entry.' };
    }
    
    // Parse the extracted text using existing format detection
    var parsed = _parseExtractedPDFText(text, safeFileName);

    return {
      success: true,
      type: 'pdf_parsed',
      extracted: parsed,
      rawText: text.substring(0, 2000), // First 2000 chars for review
      message: '✅ Extracted ' + parsed.length + ' sample(s) from ' + safeFileName
    };
    
  } catch (e) {
    Logger.log('PDF processing error: ' + String(e).substring(0, 500).replace(/[\r\n]/g, ' '));
    return { success: false, message: '❌ PDF processing failed. Try manual entry instead.' };
  }
}

function _parseExtractedPDFText(text, fileName) {
  var samples = [];
  var lines = text.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l; });
  
  // Try to detect format and extract fields
  // This is a simplified parser — the main processPDFsFromGmail has more sophisticated parsing
  
  var currentSample = {};
  var senderFound = '';
  var warehouseFound = '';
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var lower = line.toLowerCase();
    
    // Detect sender
    if (lower.indexOf('serengeti') !== -1) senderFound = 'Serengeti Trading Company';
    else if (lower.indexOf('paragon') !== -1) senderFound = 'Paragon Coffee Trading';
    else if (lower.indexOf('louis dreyfus') !== -1 || lower.indexOf('ldcom') !== -1) senderFound = 'Louis Dreyfus';
    else if (lower.indexOf('olam') !== -1) senderFound = 'Olam';
    else if (lower.indexOf('volcafe') !== -1) senderFound = 'Volcafe';
    else if (lower.indexOf('sucafina') !== -1) senderFound = 'Sucafina';
    else if (lower.indexOf('ecom') !== -1) senderFound = 'ECOM Trading';
    
    // Detect warehouse
    if (lower.indexOf('continental') !== -1) warehouseFound = 'Continental Terminals';
    else if (lower.indexOf('seaboard') !== -1) warehouseFound = 'Seaboard Marine';
    else if (lower.indexOf('gramercy') !== -1) warehouseFound = 'Gramercy';
    else if (lower.indexOf('greenport') !== -1) warehouseFound = 'Greenport';
    
    // Container number pattern: 4 letters + 7 digits
    var containerMatch = line.match(/[A-Z]{4}\d{7}/);
    if (containerMatch) currentSample.container = containerMatch[0];
    
    // ICO mark pattern: digits/digits/digits
    var markMatch = line.match(/\b(\d{1,3}\/\d+\/\d+)\b/);
    if (markMatch) currentSample.mark = markMatch[1];
    
    // Cargo number
    var cargoMatch = line.match(/\b[A-Z]?\d{4,8}\b/);
    if (lower.indexOf('cargo') !== -1 && cargoMatch) currentSample.cargo = cargoMatch[0];
    
    // Bag count
    var bagMatch = line.match(/(\d+)\s*bags?/i);
    if (bagMatch) currentSample.bagCount = bagMatch[1];
    
    // Sample weight
    var weightMatch = line.match(/(\d+(?:\.\d+)?)\s*(?:lbs?|pounds?|oz|grams?|g|kg)/i);
    if (weightMatch) currentSample.sampleWeight = weightMatch[0];
    
    // Order/reference number
    if (lower.indexOf('order') !== -1 || lower.indexOf('ref') !== -1) {
      var refMatch = line.match(/[#:\s]+([A-Z0-9-]+)/i);
      if (refMatch) currentSample.sampleOrderNum = refMatch[1];
    }
  }
  
  // Build sample entry
  currentSample.sender = senderFound || '';
  currentSample.warehouse = warehouseFound || '';
  
  // Try to get country from mark
  if (currentSample.mark && typeof getOriginFromMark === 'function') {
    currentSample.description = getOriginFromMark(currentSample.mark) || '';
  }
  
  if (Object.keys(currentSample).length > 2) { // More than just sender/warehouse
    samples.push(currentSample);
  }
  
  return samples;
}

// ============================================================
// SERVER-SIDE: Get dropdown data for web app forms
// ============================================================

function webAppGetFormData() {
  var customers = typeof getCustomerList === 'function' ? getCustomerList() : [];
  var warehouses = typeof getWarehouseList === 'function' ? getWarehouseList() : [];
  var contacts = typeof getContactsList === 'function' ? getContactsList() : { shippers: [], receivers: [] };
  
  return {
    customers: customers,
    warehouses: warehouses,
    shippers: contacts.shippers || [],
    receivers: contacts.receivers || []
  };
}

// ============================================================
// SERVER-SIDE: Ship sample from web app
// ============================================================

function webAppShipSample(sampleId, trackingNumber) {
  try {
    var result = fastLookup(sampleId);
    if (!result.success) {
      if (result.alreadyShipped) return { success: false, message: '⚠️ ' + sampleId + ' already shipped' };
      return { success: false, message: '❌ ' + sampleId + ' not found' };
    }
    
    var shipResult = fastShip(result.row, trackingNumber, result.cols);
    if (shipResult.success) {
      return { success: true, message: '✅ ' + sampleId + ' shipped — Tracking: ' + trackingNumber };
    }
    return { success: false, message: '❌ Failed to update shipping status' };
  } catch (e) {
    Logger.log('webAppShipSample error: ' + String(e).substring(0, 500).replace(/[\r\n]/g, ' '));
    return { success: false, message: '❌ Error processing ship request. Please try again.' };
  }
}
