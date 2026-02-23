// ============================================================
// ORDERS.gs — Order viewing, cover sheets, reports, lifecycle
// ============================================================
// Contains: order details, cover sheets, open orders report,
// printable forms, live orders view, pending moves, archiving
// ============================================================

// ============================================================
// HELPER: Get all samples for an order across ALL sheets
// ============================================================

// In-memory cache for getAllSamplesForOrder — cleared on each new script execution.
var _getAllSamplesCache = {};

function getAllSamplesForOrder(csOrderNum) {
  if (_getAllSamplesCache[csOrderNum]) return _getAllSamplesCache[csOrderNum];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetsToSearch = [
    CONFIG.mainSheetName,
    CONFIG.liveOrdersSheetName,
    CONFIG.completedOrdersSheetName || 'Completed Orders'
  ];
  
  var samples = [];
  var seenSamples = {};
  var orderInfo = {
    csOrder: csOrderNum,
    emailLink: '',
    sender: '',
    receiver: '',
    shippingProcess: '',
    warehouse: '',
    timestamp: '',
    sampleOrderNum: '',
    attachments: '',
    sourceEmail: '',
    bol: '',
    shipStatus: '',
    shippingLine: '',
    containerETA: '',
    shippingNotes: ''
  };
  
  sheetsToSearch.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    var col = _getColumnMap(sheet);
    if (col['CS Order #'] === undefined) return;
    
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    // QUOTA GUARD: Track execution time
    var startTime = Date.now();

    for (var i = 0; i < data.length; i++) {
      // QUOTA GUARD: Check every 100 rows for timeout
      if (i % 100 === 0) {
        var timeCheck = (typeof _quotaCheckExecutionTime === 'function') ? _quotaCheckExecutionTime(startTime) : { shouldStop: false };
        if (timeCheck.shouldStop) {
          Logger.log('archiveCompletedOrders: Execution timeout approaching after ' + i + ' rows. Stopping to avoid 6-min limit.');
          break;
        }
      }

      var rowOrderNum = String(data[i][col['CS Order #']] || '').trim();
      
      if (rowOrderNum === csOrderNum) {
        var csSampleNum = String(data[i][col['CS Sample #']] || '').trim();
        
        if (seenSamples[csSampleNum]) continue;
        seenSamples[csSampleNum] = true;
        
        var sample = {
          csSample: csSampleNum,
          container: data[i][col['Container #']] || '',
          cargo: data[i][col['Cargo #']] || '',
          mark: data[i][col['Mark #']] || '',
          reference: data[i][col['Reference']] || '',
          description: data[i][col['Description']] || '',
          bagCount: data[i][col['Bag Count']] || '',
          sampleWeight: data[i][col['Sample Weight']] || '',
          status: data[i][col['Status']] || '',
          trackingNumber: data[i][col['Tracking Number']] || '',
          comments: data[i][col['Comments']] || '',
          receiver: data[i][col['Receiver']] || '',
          shippingProcess: data[i][col['Shipping Process']] || '',
          source: sheetName
        };
        samples.push(sample);
        
        if (!orderInfo.emailLink && col['Email Link'] !== undefined) {
          orderInfo.emailLink = data[i][col['Email Link']] || '';
        }
        if (!orderInfo.sender) orderInfo.sender = data[i][col['Sender']] || '';
        if (!orderInfo.receiver) orderInfo.receiver = data[i][col['Receiver']] || '';
        if (!orderInfo.shippingProcess) orderInfo.shippingProcess = data[i][col['Shipping Process']] || '';
        if (!orderInfo.warehouse) orderInfo.warehouse = data[i][col['Warehouse']] || '';
        if (!orderInfo.timestamp) orderInfo.timestamp = data[i][col['Timestamp']] || '';
        if (!orderInfo.sampleOrderNum) orderInfo.sampleOrderNum = data[i][col['Sample Order #']] || '';
        if (!orderInfo.attachments && col['Attachments'] !== undefined) {
          orderInfo.attachments = data[i][col['Attachments']] || '';
        }
        if (!orderInfo.sourceEmail && col['Source Email'] !== undefined) {
          orderInfo.sourceEmail = data[i][col['Source Email']] || '';
        }
        if (!orderInfo.bol && col['B/L #'] !== undefined) {
          orderInfo.bol = data[i][col['B/L #']] || '';
        }
        if (!orderInfo.shipStatus && col['Ship Status'] !== undefined) {
          orderInfo.shipStatus = data[i][col['Ship Status']] || '';
        }
        if (!orderInfo.shippingLine && col['Shipping Line'] !== undefined) {
          orderInfo.shippingLine = data[i][col['Shipping Line']] || '';
        }
        if (!orderInfo.containerETA && col['Container ETA'] !== undefined) {
          orderInfo.containerETA = data[i][col['Container ETA']] || '';
        }
        if (!orderInfo.shippingNotes && col['Shipping Notes'] !== undefined) {
          orderInfo.shippingNotes = data[i][col['Shipping Notes']] || '';
        }
      }
    }
  });
  
  samples.sort(function(a, b) {
    return a.csSample.localeCompare(b.csSample);
  });

  var result = { orderInfo: orderInfo, samples: samples };
  _getAllSamplesCache[csOrderNum] = result;
  return result;
}

// ============================================================
// VIEW ORDER DETAILS
// ============================================================

function viewOrderDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  var row = sheet.getActiveCell().getRow();
  if (row < 2) {
    ui.alert('Please select a data row (not the header)');
    return;
  }
  
  var col = _getColumnMap(sheet);
  var selectedRowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var csOrderNum = String(selectedRowData[col['CS Order #']] || '').trim();
  
  if (!csOrderNum) {
    ui.alert('No CS Order # found in selected row');
    return;
  }
  
  var result = getAllSamplesForOrder(csOrderNum);
  
  if (result.samples.length === 0) {
    ui.alert('No samples found for CS Order # ' + csOrderNum);
    return;
  }
  
  var html = generateOrderDetailsHtml(result.orderInfo, result.samples);
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(850);
  ui.showModalDialog(htmlOutput, '📋 Order ' + csOrderNum + ' - ' + result.samples.length + ' Sample(s)');
}

function viewOrderCoverSheet() {
  viewOrderDetails();
}

// ============================================================
// CODE128B BARCODE SCRIPT (shared by cover sheets & details)
// ============================================================

function _getBarcodeScript() {
  return 'function drawCode128B(canvas,text){' +
    'var P=[[2,1,2,2,2,2],[2,2,2,1,2,2],[2,2,2,2,2,1],[1,2,1,2,2,3],[1,2,1,3,2,2],' +
    '[1,3,1,2,2,2],[1,2,2,2,1,3],[1,2,2,3,1,2],[1,3,2,2,1,2],[2,2,1,2,1,3],' +
    '[2,2,1,3,1,2],[2,3,1,2,1,2],[1,1,2,2,3,2],[1,2,2,1,3,2],[1,2,2,2,3,1],' +
    '[1,1,3,2,2,2],[1,2,3,1,2,2],[1,2,3,2,2,1],[2,2,3,2,1,1],[2,2,1,1,3,2],' +
    '[2,2,1,2,3,1],[2,1,3,2,1,2],[2,2,3,1,1,2],[3,1,2,1,3,1],[3,1,1,2,2,2],' +
    '[3,2,1,1,2,2],[3,2,1,2,2,1],[3,1,2,2,1,2],[3,2,2,1,1,2],[3,2,2,2,1,1],' +
    '[2,1,2,1,2,3],[2,1,2,3,2,1],[2,3,2,1,2,1],[1,1,1,3,2,3],[1,3,1,1,2,3],' +
    '[1,3,1,3,2,1],[1,1,2,3,1,3],[1,3,2,1,1,3],[1,3,2,3,1,1],[2,1,1,3,1,3],' +
    '[2,3,1,1,1,3],[2,3,1,3,1,1],[1,1,2,1,3,3],[1,1,2,3,3,1],[1,3,2,1,3,1],' +
    '[1,1,3,1,2,3],[1,1,3,3,2,1],[1,3,3,1,2,1],[3,1,3,1,2,1],[2,1,1,3,3,1],' +
    '[2,3,1,1,3,1],[2,1,3,1,1,3],[2,1,3,3,1,1],[2,1,3,1,3,1],[3,1,1,1,2,3],' +
    '[3,1,1,3,2,1],[3,3,1,1,2,1],[3,1,2,1,1,3],[3,1,2,3,1,1],[3,3,2,1,1,1],' +
    '[3,1,4,1,1,1],[2,2,1,4,1,1],[4,3,1,1,1,1],[1,1,1,2,2,4],[1,1,1,4,2,2],' +
    '[1,2,1,1,2,4],[1,2,1,4,2,1],[1,4,1,1,2,2],[1,4,1,2,2,1],[1,1,2,2,1,4],' +
    '[1,1,2,4,1,2],[1,2,2,1,1,4],[1,2,2,4,1,1],[1,4,2,1,1,2],[1,4,2,2,1,1],' +
    '[2,4,1,2,1,1],[2,2,1,1,1,4],[4,1,3,1,1,1],[2,4,1,1,1,2],[1,3,4,1,1,1],' +
    '[1,1,1,2,4,2],[1,2,1,1,4,2],[1,2,1,2,4,1],[1,1,4,2,1,2],[1,2,4,1,1,2],' +
    '[1,2,4,2,1,1],[4,1,1,2,1,2],[4,2,1,1,1,2],[4,2,1,2,1,1],[2,1,2,1,4,1],' +
    '[2,1,4,1,2,1],[4,1,2,1,2,1],[1,1,1,1,4,3],[1,1,1,3,4,1],[1,3,1,1,4,1],' +
    '[1,1,4,1,1,3],[1,1,4,3,1,1],[4,1,1,1,1,3],[4,1,1,3,1,1],[1,1,3,1,4,1],' +
    '[1,1,4,1,3,1],[3,1,1,1,4,1],[4,1,1,1,3,1],[2,1,1,4,1,2],[2,1,1,2,1,4],' +
    '[2,1,1,2,3,2]];' +
    'var STOP=[2,3,3,1,1,1,2];var codes=[104];var cs=104;' +
    'for(var i=0;i<text.length;i++){var c=text.charCodeAt(i)-32;codes.push(c);cs+=(i+1)*c;}' +
    'codes.push(cs%103);codes.push(106);var bits=[];' +
    'for(var i=0;i<codes.length;i++){var p=codes[i]===106?STOP:P[codes[i]];' +
    'for(var j=0;j<p.length;j++){var on=j%2===0;for(var k=0;k<p[j];k++)bits.push(on?1:0);}}' +
    'var ctx=canvas.getContext("2d");var w=canvas.width,h=canvas.height;' +
    'ctx.fillStyle="#fff";ctx.fillRect(0,0,w,h);var bw=w/bits.length;ctx.fillStyle="#000";' +
    'for(var i=0;i<bits.length;i++){if(bits[i])ctx.fillRect(i*bw,0,Math.ceil(bw),h-12);}' +
    'ctx.font="bold 10px Arial";ctx.textAlign="center";ctx.fillText(text,w/2,h-1);}' +
    'var _bc=document.querySelectorAll(".order-barcode");' +
    'for(var _i=0;_i<_bc.length;_i++){var _c=_bc[_i].getAttribute("data-code");if(_c)drawCode128B(_bc[_i],_c);}';
}

// ============================================================
// PRINT TODAY'S COVER SHEETS (Searches ALL sheets)
// ============================================================

function printTodaysCoverSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();
  
  if (!mainSheet) { ui.alert('Main sheet not found'); return; }
  
  var col = _getColumnMap(mainSheet);
  var lastRow = mainSheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data found'); return; }
  
  var allData = mainSheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn()).getValues();
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var coverSheetPrintedIdx = col['Cover Sheet Printed'];
  if (coverSheetPrintedIdx === undefined) {
    var lastCol = mainSheet.getLastColumn();
    mainSheet.getRange(1, lastCol + 1).setValue('Cover Sheet Printed');
    mainSheet.getRange(1, lastCol + 1).setFontWeight('bold');
    coverSheetPrintedIdx = lastCol;
    col['Cover Sheet Printed'] = coverSheetPrintedIdx;
  }
  
  var ordersToday = {};
  var timestampIdx = col['Timestamp'];
  
  for (var i = 0; i < allData.length; i++) {
    var timestamp = allData[i][timestampIdx];
    var rowDate = new Date(timestamp);
    rowDate.setHours(0, 0, 0, 0);
    
    if (rowDate.getTime() === today.getTime()) {
      var alreadyPrinted = allData[i][coverSheetPrintedIdx];
      if (!alreadyPrinted) {
        var csOrder = String(allData[i][col['CS Order #']]).trim();
        if (csOrder && !ordersToday[csOrder]) {
          ordersToday[csOrder] = { rows: [] };
        }
        if (csOrder) {
          ordersToday[csOrder].rows.push(i + 2);
        }
      }
    }
  }
  
  var orderKeys = Object.keys(ordersToday);
  if (orderKeys.length === 0) {
    ui.alert('No unprinted cover sheets from today.\n\nAll orders from today have already been printed.');
    return;
  }
  
  var fullOrders = {};
  var totalSamples = 0;
  var emailCount = 0;

  orderKeys.forEach(function(csOrderNum) {
    var result = getAllSamplesForOrder(csOrderNum);
    var emailHtml = null;
    if (result.orderInfo.emailLink) {
      var emailData = getEmailHtmlFromLink(result.orderInfo.emailLink);
      if (emailData) {
        emailHtml = emailData;
        emailCount++;
      }
    }
    fullOrders[csOrderNum] = {
      orderInfo: result.orderInfo,
      samples: result.samples,
      rows: ordersToday[csOrderNum].rows,
      emailHtml: emailHtml
    };
    totalSamples += result.samples.length;
  });
  
  var response = ui.alert(
    'Print Today\'s Cover Sheets',
    'Found ' + orderKeys.length + ' orders with ' + totalSamples + ' total samples.\n' +
    emailCount + ' original emails found.\n\n' +
    '(Includes both live and completed samples)\n\nPrint all cover sheets?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  var html = generateAllCoverSheetsHtmlFull(fullOrders);
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(850)
    .setHeight(900);
  ui.showModalDialog(htmlOutput, 'Print ' + orderKeys.length + ' Cover Sheets');
}

// ============================================================
// GENERATE ALL COVER SHEETS HTML (Full order data)
// ============================================================

function generateAllCoverSheetsHtmlFull(orders) {
  var orderKeys = Object.keys(orders);
  
  var pages = orderKeys.map(function(orderNum) {
    var order = orders[orderNum];
    var coverHtml = generateSingleCoverSheetHtmlFull(order.orderInfo, order.samples);
    
    var emailHtml = '';
    if (order.emailHtml) {
      emailHtml = '<div class="page-break"></div>' +
        '<div class="email-page">' +
        '<div class="email-header-bar">' +
        '<div class="email-label">📧 Original Order Email — CS Order #' + orderNum + '</div>' +
        '</div>' +
        '<div class="email-meta">' +
        '<div><strong>From:</strong> ' + (order.emailHtml.from || '') + '</div>' +
        '<div><strong>Subject:</strong> ' + (order.emailHtml.subject || '') + '</div>' +
        '<div><strong>Date:</strong> ' + (order.emailHtml.date || '') + '</div>' +
        '</div>' +
        '<div class="email-body">' + order.emailHtml.body + '</div>' +
        '</div>';
    } else {
      emailHtml = '<div class="no-email-note">' +
        '📧 No original email available for this order' +
        (order.orderInfo.sourceEmail ? ' (Source: ' + order.orderInfo.sourceEmail + ')' : '') +
        '</div>';
    }
    
    return coverHtml + emailHtml;
  }).join('<div class="page-break"></div>');
  
  var html = '<!DOCTYPE html><html><head><style>' +
    '* { box-sizing: border-box; } ' +
    'body { font-family: Arial, sans-serif; margin: 0; padding: 10px; background: #f5f5f5; } ' +
    '.controls { text-align: center; padding: 15px; background: white; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); } ' +
    '.btn { display: inline-block; padding: 12px 25px; margin: 5px; border: none; border-radius: 5px; cursor: pointer; font-size: 18pt; text-decoration: none; } ' +
    '.print-btn { background: #4A7C59; color: white; } .print-btn:hover { background: #3a6249; } ' +
    '.mark-btn { background: #0066cc; color: white; } .mark-btn:hover { background: #0052a3; } ' +
    '.summary { font-size: 21pt; margin-bottom: 10px; } ' +
    '.page-break { page-break-after: always; height: 20px; border-bottom: 2px dashed #ccc; margin: 20px 0; } ' +
    '.cover-sheet { background: white; border: 2px solid #2E5339; border-radius: 8px; padding: 20px; margin: 0 auto 20px auto; max-width: 750px; } ' +
    '.header { text-align: center; border-bottom: 3px solid #4A7C59; padding-bottom: 12px; margin-bottom: 15px; } ' +
    '.company-name { font-size: 27pt; font-weight: bold; color: #2E5339; margin: 0; } ' +
    '.tagline { font-size: 13.5pt; color: #666; font-style: italic; margin: 3px 0 0 0; } ' +
    '.order-header { display: flex; justify-content: space-between; align-items: center; background: #f0f7f2; padding: 12px; border-radius: 5px; margin-bottom: 12px; } ' +
    '.order-id-block { display: flex; flex-direction: column; align-items: flex-start; } ' +
    '.order-number { font-size: 30pt; font-weight: bold; color: #4A7C59; } ' +
    '.order-barcode { margin-top: 4px; } ' +
    '.sample-counts { display: flex; gap: 8px; } ' +
    '.sample-count { font-size: 15pt; padding: 5px 10px; border-radius: 12px; border: 2px solid; } ' +
    '.count-live { background: #fff3cd; border-color: #ffc107; color: #856404; } ' +
    '.count-completed { background: #d4edda; border-color: #28a745; color: #155724; } ' +
    '.count-total { background: white; border-color: #4A7C59; color: #4A7C59; } ' +
    '.info-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; margin-bottom: 12px; } ' +
    '.info-box { background: #fafafa; padding: 10px; border-radius: 5px; border-left: 4px solid #4A7C59; } ' +
    '.info-label { font-size: 12pt; color: #666; text-transform: uppercase; margin-bottom: 2px; } ' +
    '.info-value { font-size: 16.5pt; font-weight: bold; color: #2E5339; } ' +
    '.receiver-group { margin-bottom: 15px; } ' +
    '.receiver-header { background: #2E5339; color: #fff; padding: 10px 15px; font-size: 16.5pt; font-weight: bold; border-radius: 6px 6px 0 0; } ' +
    'table { width: 100%; border-collapse: collapse; font-size: 12pt; } ' +
    'th { background: #4A7C59; color: white; padding: 6px 4px; text-align: left; font-weight: bold; } ' +
    'td { padding: 5px 4px; border-bottom: 1px solid #ddd; } ' +
    'tr:nth-child(even) { background: #f9f9f9; } ' +
    '.status-badge { display: inline-block; padding: 2px 6px; border-radius: 8px; font-weight: bold; font-size: 10.5pt; } ' +
    '.status-received { background: #fff3cd; color: #856404; } ' +
    '.status-scanned { background: #cce5ff; color: #004085; } ' +
    '.status-shipped { background: #d4edda; color: #155724; } ' +
    '.source-badge { display: inline-block; padding: 1px 5px; border-radius: 6px; font-size: 9pt; font-weight: bold; margin-left: 3px; } ' +
    '.source-live { background: #fff3cd; color: #856404; } ' +
    '.source-completed { background: #d4edda; color: #155724; } ' +
    '.timestamp { text-align: center; font-size: 12pt; color: #999; margin-top: 8px; } ' +
    '.email-page { background: white; border: 2px solid #1976d2; border-radius: 8px; padding: 0; margin: 0 auto 20px auto; max-width: 750px; overflow: hidden; } ' +
    '.email-header-bar { background: #1976d2; color: white; padding: 10px 16px; font-size: 18pt; font-weight: bold; } ' +
    '.email-label { margin: 0; } ' +
    '.email-meta { padding: 12px 16px; background: #e3f2fd; font-size: 15pt; border-bottom: 1px solid #bbdefb; } ' +
    '.email-meta div { margin-bottom: 3px; } .email-meta div:last-child { margin-bottom: 0; } ' +
    '.email-body { padding: 16px; font-size: 15pt; line-height: 1.5; overflow: hidden; word-wrap: break-word; } ' +
    '.email-body img { max-width: 100%; height: auto; } .email-body table { font-size: inherit; } ' +
    '.no-email-note { text-align: center; padding: 10px 16px; margin: 10px auto 20px; max-width: 750px; background: #fff3cd; color: #856404; border: 1px solid #ffc107; border-radius: 5px; font-size: 15pt; font-style: italic; } ' +
    '.shipping-detail-row { display: grid; grid-template-columns: 1fr 2fr 1fr 1fr 1fr; gap: 8px; background: #e3f2fd; border: 2px solid #1976d2; border-radius: 5px; padding: 8px 12px; margin-bottom: 12px; align-items: start; } ' +
    '.ship-field .info-label { font-size: 9pt; } ' +
    '.ship-field .info-value { font-size: 13pt; } ' +
    '.ship-field-note { min-width: 0; } ' +
    '.ship-note-input { width: 100%; border: 1px solid #90caf9; border-radius: 3px; padding: 3px 6px; font-size: 11pt; font-family: Arial, sans-serif; background: #fff; } ' +
    '.ship-select { width: 100%; border: 1px solid #90caf9; border-radius: 3px; padding: 3px 4px; font-size: 11pt; font-family: Arial, sans-serif; background: #fff; } ' +
    '.carrier-select { padding: 2px 4px; border: 1px solid #ccc; border-radius: 3px; font-size: 8pt; } ' +
    '@media print { .controls { display: none; } body { background: white; padding: 0; } .cover-sheet { border: 1px solid #000; margin: 0; page-break-inside: avoid; } .email-page { border: 1px solid #000; margin: 0; page-break-inside: avoid; } .page-break { height: 0; border: none; margin: 0; } .no-email-note { page-break-inside: avoid; } .carrier-select, .ship-select { border: none; background: transparent; -webkit-appearance: none; -moz-appearance: none; font-weight: bold; padding: 0; } .ship-note-input { border: none; background: transparent; padding: 0; font-weight: bold; } .shipping-detail-row { border: 1px solid #000; background: #f0f0f0; } }' +
    '</style></head><body>' +
    '<div class="controls">' +
    '<div class="summary">📋 ' + orderKeys.length + ' Cover Sheets Ready (With Original Emails)</div>' +
    '<button class="btn print-btn" onclick="window.print()">🖨️ Print All Cover Sheets + Emails</button>' +
    '<button class="btn mark-btn" onclick="google.script.run.withSuccessHandler(onMarked).markCoverSheetsAsPrinted()">✅ Mark as Printed & Close</button>' +
    '</div>' +
    pages +
    '<script>function onMarked(result) { alert(result); google.script.host.close(); }\n' +
    'function _shipSave(el){var row=el.closest(".shipping-detail-row");var ord=row.getAttribute("data-order");var field=el.getAttribute("data-field");var val=el.value||"";el.style.outline="2px solid #ffc107";google.script.run.withSuccessHandler(function(){el.style.outline="2px solid #28a745";setTimeout(function(){el.style.outline="";},1200);}).withFailureHandler(function(e){el.style.outline="2px solid #dc3545";alert("Save failed: "+e);}).updateCoverSheetField(ord,field,val);}\n' +
    'document.addEventListener("change",function(e){var t=e.target;if(t.classList.contains("ship-select")||t.classList.contains("ship-note-input")){_shipSave(t);}});\n' +
    'document.addEventListener("blur",function(e){var t=e.target;if(t.classList.contains("ship-note-input")){_shipSave(t);}},true);\n' +
    _getBarcodeScript() +
    '</script>' +
    '</body></html>';
  
  return html;
}

// ============================================================
// HELPER: Build shipping info bar HTML
// ============================================================

function _buildShippingInfoBar(order) {
  // Shipping Line dropdown
  var shipLines = ['', 'MSC', 'ZIM', 'HAPAG-LLOYD', 'MAERSK', 'SEABOARD MARINE', 'ONE', 'CMA-CGM', 'EVERGREEN', 'COSCO'];
  var shipLineOptions = shipLines.map(function(sl) {
    var sel = (order.shippingLine && order.shippingLine.toUpperCase() === sl.toUpperCase()) ? ' selected' : '';
    return '<option value="' + sl + '"' + sel + '>' + (sl || '--') + '</option>';
  }).join('');

  // Vessel Status dropdown
  var statuses = ['', 'Afloat', 'Landed', 'Discharged', 'At Warehouse', 'In Transit', 'Customs Hold', 'Released'];
  var statusOptions = statuses.map(function(st) {
    var sel = (order.shipStatus && order.shipStatus.toLowerCase() === st.toLowerCase()) ? ' selected' : '';
    return '<option value="' + st + '"' + sel + '>' + (st || '--') + '</option>';
  }).join('');

  var oid = (order.csOrder || '').replace(/"/g, '&quot;');

  return '<div class="shipping-detail-row" data-order="' + oid + '">' +
    '<div class="ship-field"><div class="info-label">B/L</div><div class="info-value">' + (order.bol || '-') + '</div></div>' +
    '<div class="ship-field ship-field-note"><div class="info-label">Note</div>' +
      '<input type="text" class="ship-note-input" data-field="Shipping Notes" value="' + (order.shippingNotes || '').replace(/"/g, '&quot;') + '" placeholder="Add note...">' +
    '</div>' +
    '<div class="ship-field"><div class="info-label">Ship Line</div>' +
      '<select class="ship-select" data-field="Shipping Line">' + shipLineOptions + '</select>' +
    '</div>' +
    '<div class="ship-field"><div class="info-label">Status</div>' +
      '<select class="ship-select" data-field="Ship Status">' + statusOptions + '</select>' +
    '</div>' +
    '<div class="ship-field"><div class="info-label">ETA</div><div class="info-value">' + (order.containerETA || '-') + '</div></div>' +
    '</div>';
}

// ============================================================
// UPDATE SHIPPING FIELD FROM COVER SHEET
// Called via google.script.run from cover sheet dropdowns/inputs
// ============================================================

function updateCoverSheetField(csOrderNum, fieldName, value) {
  var allowed = ['Shipping Notes', 'Shipping Line', 'Ship Status'];
  if (allowed.indexOf(fieldName) === -1) {
    throw new Error('Field not allowed: ' + fieldName);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToUpdate = [
    (typeof CONFIG !== 'undefined' && CONFIG.mainSheetName) ? CONFIG.mainSheetName : 'All Orders',
    'Completed Orders'
  ];

  var updated = 0;

  sheetsToUpdate.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    var col = _getColumnMap(sheet);
    if (col['CS Order #'] === undefined || col[fieldName] === undefined) return;

    var orderColIdx = col['CS Order #'];
    var fieldColIdx = col[fieldName];
    var data = sheet.getRange(2, orderColIdx + 1, sheet.getLastRow() - 1, 1).getValues();

    // V5: perf — collect matching rows then batch-write via RangeList (N setValue → 1 setValue)
    var matchingA1 = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(csOrderNum).trim()) {
        matchingA1.push(sheet.getRange(i + 2, fieldColIdx + 1).getA1Notation());
        updated++;
      }
    }
    if (matchingA1.length > 0) {
      sheet.getRangeList(matchingA1).setValue(value);
    }
  });

  if (updated === 0) {
    throw new Error('Order ' + csOrderNum + ' not found');
  }
  return updated + ' row(s) updated';
}

// ============================================================
// GENERATE SINGLE COVER SHEET HTML (Full order data)
// ============================================================

function generateSingleCoverSheetHtmlFull(order, samples) {
  var statusCounts = { live: 0, completed: 0 };
  samples.forEach(function(s) {
    if (s.status && s.status.toLowerCase() === 'shipped') {
      statusCounts.completed++;
    } else {
      statusCounts.live++;
    }
  });

  var receiverGroups = {};
  var receiverOrder = [];
  samples.forEach(function(s) {
    var recv = s.receiver || order.receiver || 'N/A';
    if (!receiverGroups[recv]) {
      receiverGroups[recv] = { samples: [], shippingProcess: s.shippingProcess || order.shippingProcess || 'Standard' };
      receiverOrder.push(recv);
    }
    receiverGroups[recv].samples.push(s);
  });

  var receiverGroupsHtml = receiverOrder.map(function(recv) {
    var group = receiverGroups[recv];

    var carrier = null;
    try { carrier = lookupCarrierByCompany(recv); } catch(e) { logError('_buildOrdersHtml', 'Error looking up carrier for receiver', { receiver: recv, error: e.message }); }
    var carrierLine = '';
    if (carrier && (carrier.fedex || carrier.ups)) {
      var parts = [];
      if (carrier.fedex) parts.push('FedEx: ' + carrier.fedex + (carrier.preferred === 'FedEx' ? ' ★' : ''));
      if (carrier.ups) parts.push('UPS: ' + carrier.ups + (carrier.preferred === 'UPS' ? ' ★' : ''));
      carrierLine = '<div style="font-size:13.5pt;color:#cde;margin-top:4px;">📋 ' + parts.join(' &nbsp;|&nbsp; ') + '</div>';
    }

    // Build carrier dropdown options
    var carrierOptions = '<option value="">--</option>';
    if (carrier) {
      if (carrier.fedex) carrierOptions += '<option value="FedEx"' + (carrier.preferred === 'FedEx' ? ' selected' : '') + '>FedEx</option>';
      if (carrier.ups) carrierOptions += '<option value="UPS"' + (carrier.preferred === 'UPS' ? ' selected' : '') + '>UPS</option>';
    } else {
      carrierOptions += '<option value="FedEx">FedEx</option><option value="UPS">UPS</option>';
    }
    carrierOptions += '<option value="USPS">USPS</option><option value="Pickup">Pickup</option>';

    var rows = group.samples.map(function(s) {
      var sourceBadge = '';
      if (s.source === 'Completed Orders') {
        sourceBadge = '<span class="source-badge source-completed">✓</span>';
      }
      return '<tr>' +
        '<td><strong>' + s.csSample + '</strong>' + sourceBadge + '</td>' +
        '<td>' + (s.container || '-') + '</td>' +
        '<td>' + (s.cargo || '-') + '</td>' +
        '<td>' + (s.mark || '-') + '</td>' +
        '<td>' + (s.description || '-') + '</td>' +
        '<td>' + (s.bagCount || '-') + '</td>' +
        '<td>' + (s.sampleWeight || '-') + '</td>' +
        '<td><select class="carrier-select">' + carrierOptions + '</select></td>' +
        '</tr>';
    }).join('');

    return '<div class="receiver-group">' +
      '<div class="receiver-header">' +
      '<span class="ship-icon">📦</span> Ship To: ' + recv +
      '<span style="float:right; font-size:18px; opacity:0.8;">🚚 ' + group.shippingProcess + '</span>' +
      carrierLine +
      '</div>' +
      '<table><thead><tr>' +
      '<th>CS Sample #</th><th>Container</th><th>Cargo</th><th>Mark</th><th>Description</th><th>Bags</th><th>Weight</th><th>Carrier</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table></div>';
  }).join('');

  var shippingBar = _buildShippingInfoBar(order);

  return '<div class="cover-sheet">' +
    '<div class="header"><p class="company-name">COMMODITY SAMPLER SERVICES</p><p class="tagline">"Integrity Through Independence"</p></div>' +
    '<div class="order-header">' +
    '<div class="order-id-block">' +
    '<div class="order-number">CS Order #' + order.csOrder + '</div>' +
    '<canvas class="order-barcode" data-code="' + order.csOrder + '" width="200" height="50"></canvas>' +
    '</div>' +
    '<div class="sample-counts">' +
    '<span class="sample-count count-live">📋 ' + statusCounts.live + ' Live</span>' +
    '<span class="sample-count count-completed">✅ ' + statusCounts.completed + ' Done</span>' +
    '<span class="sample-count count-total">📦 ' + samples.length + ' Total</span>' +
    '</div></div>' +
    '<div class="info-grid">' +
    '<div class="info-box"><div class="info-label">Client</div><div class="info-value">' + (order.sender || '-') + '</div></div>' +
    '<div class="info-box"><div class="info-label">Warehouse</div><div class="info-value">' + (order.warehouse || '-') + '</div></div>' +
    '<div class="info-box"><div class="info-label">Order #</div><div class="info-value">' + (order.sampleOrderNum || '-') + '</div></div>' +
    '</div>' +
    shippingBar +
    receiverGroupsHtml +
    '<div class="timestamp">Created: ' + order.timestamp + '</div></div>';
}

// ============================================================
// GENERATE ORDER DETAILS HTML (Shows source sheet)
// ============================================================

function generateOrderDetailsHtml(order, samples) {
  var emailSection = '';
  if (order.emailLink) {
    emailSection = '<div class="email-section">' +
      '<a href="' + order.emailLink + '" target="_blank" class="btn email-btn">📧 Open Original Email in Gmail</a>' +
      '<span class="email-note">Click to view the original order email</span></div>';
  } else {
    emailSection = '<div class="email-section no-email"><span class="email-missing">📧 No email link available for this order</span></div>';
  }
  
  var attachmentSection = '';
  if (order.attachments) {
    var urls = order.attachments.split('\n').filter(function(u) { return u.trim(); });
    if (urls.length > 0) {
      var attachLinks = urls.map(function(url, i) {
        return '<a href="' + url + '" target="_blank" class="btn attach-btn">📎 PDF ' + (i + 1) + '</a>';
      }).join(' ');
      attachmentSection = '<div class="attachments-section"><strong>Attachments:</strong> ' + attachLinks + '</div>';
    }
  }
  
  var statusCounts = { live: 0, completed: 0 };
  samples.forEach(function(s) {
    if (s.status && s.status.toLowerCase() === 'shipped') { statusCounts.completed++; }
    else { statusCounts.live++; }
  });
  
  var receiverGroups = {};
  var receiverOrder = [];
  samples.forEach(function(s) {
    var recv = s.receiver || order.receiver || 'N/A';
    if (!receiverGroups[recv]) {
      receiverGroups[recv] = { samples: [], shippingProcess: s.shippingProcess || order.shippingProcess || 'Standard' };
      receiverOrder.push(recv);
    }
    receiverGroups[recv].samples.push(s);
  });
  
  var receiverGroupsHtml = receiverOrder.map(function(recv) {
    var group = receiverGroups[recv];
    
    var carrier = null;
    try { carrier = lookupCarrierByCompany(recv); } catch(e) { logError('_buildReceiverStatusHtml', 'Error looking up carrier for receiver', { receiver: recv, error: e.message }); }
    var carrierLine = '';
    if (carrier && (carrier.fedex || carrier.ups)) {
      var parts = [];
      if (carrier.fedex) parts.push('FedEx: ' + carrier.fedex + (carrier.preferred === 'FedEx' ? ' ★' : ''));
      if (carrier.ups) parts.push('UPS: ' + carrier.ups + (carrier.preferred === 'UPS' ? ' ★' : ''));
      carrierLine = '<div style="font-size:13.5pt;color:#cde;margin-top:4px;">📋 ' + parts.join(' &nbsp;|&nbsp; ') + '</div>';
    }
    
    var rows = group.samples.map(function(s) {
      var statusClass = 'status-' + (s.status || 'received').toLowerCase().replace(' ', '-');
      var trackingCell = s.trackingNumber || '-';
      if (s.trackingNumber) {
        var trackUrl = getTrackingUrl(s.trackingNumber);
        if (trackUrl) {
          trackingCell = '<a href="' + trackUrl + '" target="_blank">' + s.trackingNumber + '</a>';
        }
      }
      var sourceBadge = '';
      if (s.source === 'Completed Orders') {
        sourceBadge = '<span class="source-badge source-completed">Completed</span>';
      } else if (s.source === 'Live Orders' || s.source === 'All Orders') {
        sourceBadge = '<span class="source-badge source-live">Live</span>';
      }
      return '<tr>' +
        '<td><strong>' + s.csSample + '</strong></td>' +
        '<td>' + (s.container || '-') + '</td>' +
        '<td>' + (s.cargo || '-') + '</td>' +
        '<td>' + (s.mark || '-') + '</td>' +
        '<td>' + (s.description || '-') + '</td>' +
        '<td>' + (s.bagCount || '-') + '</td>' +
        '<td><span class="status-badge ' + statusClass + '">' + (s.status || 'Received') + '</span></td>' +
        '<td>' + trackingCell + '</td>' +
        '<td>' + sourceBadge + '</td>' +
        '</tr>';
    }).join('');
    
    return '<div class="receiver-group" style="margin-bottom:15px;">' +
      '<div style="background:#2E5339;color:#fff;padding:10px 15px;font-size:16.5pt;font-weight:bold;border-radius:6px 6px 0 0;">' +
      '📦 Ship To: ' + recv +
      '<span style="float:right;font-size:15pt;opacity:0.8;">🚚 ' + group.shippingProcess + '</span>' +
      carrierLine + '</div>' +
      '<table><thead><tr>' +
      '<th>CS Sample #</th><th>Container</th><th>Cargo</th><th>Mark</th><th>Description</th><th>Bags</th><th>Status</th><th>Tracking</th><th>Sheet</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table></div>';
  }).join('');
  
  var html = '<!DOCTYPE html><html><head><style>' +
    '* { box-sizing: border-box; } body { font-family: Arial, sans-serif; margin: 0; padding: 15px; background: #f5f5f5; } ' +
    '.cover-sheet { background: white; border: 2px solid #2E5339; border-radius: 8px; padding: 20px; margin: 0 auto; } ' +
    '.header { text-align: center; border-bottom: 3px solid #4A7C59; padding-bottom: 12px; margin-bottom: 15px; } ' +
    '.company-name { font-size: 27pt; font-weight: bold; color: #2E5339; margin: 0; } ' +
    '.tagline { font-size: 13.5pt; color: #666; font-style: italic; margin: 3px 0 0 0; } ' +
    '.order-header { display: flex; justify-content: space-between; align-items: center; background: #f0f7f2; padding: 15px; border-radius: 5px; margin-bottom: 15px; } ' +
    '.order-id-block { display: flex; flex-direction: column; align-items: flex-start; } ' +
    '.order-number { font-size: 33pt; font-weight: bold; color: #4A7C59; } ' +
    '.order-barcode { margin-top: 4px; } ' +
    '.sample-counts { display: flex; gap: 10px; } .sample-count { font-size: 18pt; padding: 8px 12px; border-radius: 15px; border: 2px solid; } ' +
    '.count-live { background: #fff3cd; border-color: #ffc107; color: #856404; } ' +
    '.count-completed { background: #d4edda; border-color: #28a745; color: #155724; } ' +
    '.count-total { background: white; border-color: #4A7C59; color: #4A7C59; } ' +
    '.email-section { background: #e3f2fd; padding: 15px; border-radius: 8px; margin-bottom: 15px; text-align: center; border: 2px solid #1976d2; } ' +
    '.email-section.no-email { background: #fff3cd; border-color: #ffc107; } ' +
    '.email-btn { display: inline-block; padding: 12px 25px; background: #1976d2; color: white; text-decoration: none; border-radius: 5px; font-size: 21pt; font-weight: bold; } ' +
    '.email-btn:hover { background: #1565c0; } .email-note { display: block; margin-top: 8px; font-size: 15pt; color: #666; } ' +
    '.email-missing { color: #856404; font-style: italic; } ' +
    '.attachments-section { background: #f5f5f5; padding: 12px; border-radius: 5px; margin-bottom: 15px; } ' +
    '.attach-btn { display: inline-block; padding: 8px 15px; background: #6c757d; color: white; text-decoration: none; border-radius: 4px; margin: 3px; font-size: 15pt; } ' +
    '.info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 15px; } ' +
    '.info-box { background: #fafafa; padding: 12px; border-radius: 5px; border-left: 4px solid #4A7C59; } ' +
    '.info-label { font-size: 13.5pt; color: #666; text-transform: uppercase; margin-bottom: 3px; } ' +
    '.info-value { font-size: 16.5pt; font-weight: bold; color: #2E5339; } ' +
    'table { width: 100%; border-collapse: collapse; font-size: 13.5pt; } ' +
    'th { background: #4A7C59; color: white; padding: 8px 5px; text-align: left; font-weight: bold; } ' +
    'td { padding: 6px 5px; border-bottom: 1px solid #ddd; } tr:nth-child(even) { background: #f9f9f9; } tr:hover { background: #e8f5e9; } td a { color: #1976d2; } ' +
    '.status-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-weight: bold; font-size: 12pt; } ' +
    '.status-received { background: #fff3cd; color: #856404; } .status-scanned { background: #cce5ff; color: #004085; } ' +
    '.status-shipped { background: #d4edda; color: #155724; } .status-reported { background: #d1ecf1; color: #0c5460; } ' +
    '.source-badge { display: inline-block; padding: 2px 6px; border-radius: 8px; font-size: 10.5pt; font-weight: bold; } ' +
    '.source-live { background: #fff3cd; color: #856404; } .source-completed { background: #d4edda; color: #155724; } ' +
    '.btn { display: inline-block; padding: 10px 18px; margin: 5px; border: none; border-radius: 5px; cursor: pointer; font-size: 16.5pt; text-decoration: none; text-align: center; } ' +
    '.print-btn { background: #4A7C59; color: white; } .close-btn { background: #6c757d; color: white; } ' +
    '.btn-row { text-align: center; margin-top: 15px; padding-top: 12px; border-top: 1px solid #ddd; } ' +
    '.timestamp { text-align: center; font-size: 12pt; color: #999; margin-top: 10px; } ' +
    '.shipping-detail-row { display: grid; grid-template-columns: 1fr 2fr 1fr 1fr 1fr; gap: 8px; background: #e3f2fd; border: 2px solid #1976d2; border-radius: 5px; padding: 8px 12px; margin-bottom: 15px; align-items: start; } ' +
    '.ship-field .info-label { font-size: 9pt; } .ship-field .info-value { font-size: 13pt; } ' +
    '.ship-field-note { min-width: 0; } ' +
    '.ship-note-input { width: 100%; border: 1px solid #90caf9; border-radius: 3px; padding: 3px 6px; font-size: 11pt; font-family: Arial, sans-serif; background: #fff; } ' +
    '.ship-select { width: 100%; border: 1px solid #90caf9; border-radius: 3px; padding: 3px 4px; font-size: 11pt; font-family: Arial, sans-serif; background: #fff; } ' +
    '@media print { body { background: white; padding: 0; } .btn-row, .email-section, .attachments-section { display: none; } .cover-sheet { border: 1px solid #000; } table { font-size: 12pt; } .ship-select { border: none; background: transparent; -webkit-appearance: none; -moz-appearance: none; font-weight: bold; padding: 0; } .ship-note-input { border: none; background: transparent; padding: 0; font-weight: bold; } .shipping-detail-row { border: 1px solid #000; background: #f0f0f0; } }' +
    '</style></head><body>' +
    '<div class="cover-sheet">' +
    '<div class="header"><p class="company-name">COMMODITY SAMPLER SERVICES</p><p class="tagline">"Integrity Through Independence"</p></div>' +
    '<div class="order-header">' +
    '<div class="order-id-block">' +
    '<div class="order-number">CS Order #' + order.csOrder + '</div>' +
    '<canvas class="order-barcode" data-code="' + order.csOrder + '" width="200" height="50"></canvas>' +
    '</div>' +
    '<div class="sample-counts">' +
    '<span class="sample-count count-live">📋 ' + statusCounts.live + ' Live</span>' +
    '<span class="sample-count count-completed">✅ ' + statusCounts.completed + ' Completed</span>' +
    '<span class="sample-count count-total">📦 ' + samples.length + ' Total</span>' +
    '</div></div>' +
    emailSection + attachmentSection +
    '<div class="info-grid" style="grid-template-columns: 1fr 1fr 1fr;">' +
    '<div class="info-box"><div class="info-label">Client</div><div class="info-value">' + (order.sender || '-') + '</div></div>' +
    '<div class="info-box"><div class="info-label">Warehouse</div><div class="info-value">' + (order.warehouse || '-') + '</div></div>' +
    '<div class="info-box"><div class="info-label">Order # / Reference</div><div class="info-value">' + (order.sampleOrderNum || '-') + '</div></div>' +
    '</div>' +
    _buildShippingInfoBar(order) +
    receiverGroupsHtml +
    '<div class="timestamp">Order Date: ' + order.timestamp + ' | Source: ' + (order.sourceEmail || 'N/A') + '</div>' +
    '<div class="btn-row">' +
    '<button class="btn print-btn" onclick="window.print()">🖨️ Print Cover Sheet</button>' +
    '<button class="btn close-btn" onclick="google.script.host.close()">Close</button>' +
    '</div></div>' +
    '<script>' +
    'function _shipSave(el){var row=el.closest(".shipping-detail-row");var ord=row.getAttribute("data-order");var field=el.getAttribute("data-field");var val=el.value||"";el.style.outline="2px solid #ffc107";google.script.run.withSuccessHandler(function(){el.style.outline="2px solid #28a745";setTimeout(function(){el.style.outline="";},1200);}).withFailureHandler(function(e){el.style.outline="2px solid #dc3545";alert("Save failed: "+e);}).updateCoverSheetField(ord,field,val);}\n' +
    'document.addEventListener("change",function(e){var t=e.target;if(t.classList.contains("ship-select")||t.classList.contains("ship-note-input")){_shipSave(t);}});\n' +
    'document.addEventListener("blur",function(e){var t=e.target;if(t.classList.contains("ship-note-input")){_shipSave(t);}},true);\n' +
    _getBarcodeScript() + '</script>' +
    '</body></html>';

  return html;
}

// ============================================================
// COVER SHEET PRINT FLAGS
// ============================================================

function markCoverSheetsAsPrinted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);

  if (!sheet) return 'Error: Main sheet not found';

  var col = _getColumnMap(sheet);
  var coverSheetPrintedIdx = col['Cover Sheet Printed'];
  
  if (coverSheetPrintedIdx === undefined) {
    return 'Error: Cover Sheet Printed column not found';
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '✅ Marked 0 samples as printed';

  var allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var timestampIdx = col['Timestamp'];
  var printedCol = coverSheetPrintedIdx + 1;

  if (timestampIdx === undefined) return 'Error: Timestamp column not found';

  var now = new Date();
  var rowsToMark = [];
  for (var i = 0; i < allData.length; i++) {
    var timestamp = allData[i][timestampIdx];
    if (!timestamp) continue;
    var rowDate = new Date(timestamp);
    if (isNaN(rowDate.getTime())) continue;
    rowDate.setHours(0, 0, 0, 0);

    if (rowDate.getTime() === today.getTime() && !allData[i][coverSheetPrintedIdx]) {
      rowsToMark.push(i + 2);
    }
  }

  // Batch write using RangeList (1 API call instead of N)
  if (rowsToMark.length > 0) {
    var a1List = rowsToMark.map(function(r) {
      return sheet.getRange(r, printedCol).getA1Notation();
    });
    sheet.getRangeList(a1List).setValue(now);
  }

  return '✅ Marked ' + rowsToMark.length + ' samples as printed';
}

function clearCoverSheetPrintedFlags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Main sheet not found');
    return;
  }

  var col = _getColumnMap(sheet);
  var coverSheetPrintedIdx = col['Cover Sheet Printed'];
  
  if (coverSheetPrintedIdx === undefined) {
    ui.alert('Cover Sheet Printed column not found');
    return;
  }
  
  var response = ui.alert(
    'Clear Print Flags',
    'This will clear all "Cover Sheet Printed" timestamps, allowing all orders to be printed again.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, coverSheetPrintedIdx + 1, lastRow - 1, 1).clearContent();
  }
  
  ui.alert('✅ All cover sheet print flags cleared');
}

// ============================================================
// UPDATE LIVE ORDERS VIEW (excludes Shipped)
// ============================================================

function updateLiveOrdersView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!mainSheet) return;

  var col = _getColumnMap(mainSheet);
  var statusIdx = col['Status'];
  var totalCols = mainSheet.getLastColumn();
  var headers = mainSheet.getRange(1, 1, 1, totalCols).getValues();

  // Auto-create Live Orders if missing
  var liveSheet = ss.getSheetByName(CONFIG.liveOrdersSheetName || 'Live Orders');
  if (!liveSheet) {
    liveSheet = ss.insertSheet(CONFIG.liveOrdersSheetName || 'Live Orders');
    liveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    liveSheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold').setBackground('#d9ead3');
    liveSheet.setFrozenRows(1);
    try { liveSheet.getRange(1, 1, 1, headers[0].length).createFilter(); } catch(e) { logError('setupOrders', 'Failed to create filter on Live Orders sheet', { error: e.message }); }
    Logger.log('Created Live Orders sheet');
  }

  // Auto-create Completed Orders if missing
  var completedSheet = ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders');
  if (!completedSheet) {
    completedSheet = ss.insertSheet(CONFIG.completedOrdersSheetName || 'Completed Orders');
    completedSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    completedSheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold').setBackground('#f4cccc');
    completedSheet.setFrozenRows(1);
    try { completedSheet.getRange(1, 1, 1, headers[0].length).createFilter(); } catch(e) { logError('setupOrders', 'Failed to create filter on Completed Orders sheet', { error: e.message }); }
    Logger.log('Created Completed Orders sheet');
  }

  var lastRow = mainSheet.getLastRow();
  var liveRows = [];
  var completedRows = [];

  if (lastRow >= 2) {
    var data = mainSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
    var formulas = mainSheet.getRange(2, 1, lastRow - 1, totalCols).getFormulas();

    var liveStatuses = [CONFIG.statusValues.RECEIVED, CONFIG.statusValues.SCANNED];
    var completedStatuses = [CONFIG.statusValues.SHIPPED, 'Completed', CONFIG.statusValues.REPORTED, CONFIG.statusValues.ARCHIVED];

    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][statusIdx] || '').trim();
      var rowData = data[i].slice();
      for (var j = 0; j < formulas[i].length; j++) {
        if (formulas[i][j]) rowData[j] = formulas[i][j];
      }

      if (liveStatuses.indexOf(status) !== -1) {
        liveRows.push(rowData);
      } else if (completedStatuses.indexOf(status) !== -1) {
        completedRows.push(rowData);
      } else if (status === '') {
        // No status = treat as Received (live)
        liveRows.push(rowData);
      }
    }
  }

  // Populate Live Orders
  if (liveSheet.getLastRow() > 1) {
    liveSheet.getRange(2, 1, liveSheet.getLastRow() - 1, liveSheet.getLastColumn() || totalCols).clear();
  }
  if (liveRows.length > 0) {
    liveSheet.getRange(2, 1, liveRows.length, totalCols).setValues(liveRows);
    var printIdx = col['Print'];
    if (printIdx !== undefined) {
      liveSheet.getRange(2, printIdx + 1, liveRows.length, 1).insertCheckboxes();
    }
  }
  Logger.log('Live Orders updated: ' + liveRows.length + ' active samples');

  // Populate Completed Orders — preserve rows permanently moved by processPendingMoves
  // (those rows exist only in Completed, not in All Orders anymore)
  var sampleIdx = col['CS Sample #'];
  var newCompletedSamples = {};
  for (var ci = 0; ci < completedRows.length; ci++) {
    var csId = (sampleIdx !== undefined) ? String(completedRows[ci][sampleIdx] || '').trim() : '';
    if (csId) newCompletedSamples[csId] = true;
  }

  // Read existing Completed rows before clearing — keep any not in All Orders
  var preservedRows = [];
  if (completedSheet.getLastRow() > 1 && sampleIdx !== undefined) {
    var existingCompleted = completedSheet.getRange(2, 1, completedSheet.getLastRow() - 1, completedSheet.getLastColumn()).getValues();
    var existingFormulas = completedSheet.getRange(2, 1, completedSheet.getLastRow() - 1, completedSheet.getLastColumn()).getFormulas();
    for (var ec = 0; ec < existingCompleted.length; ec++) {
      var ecSample = String(existingCompleted[ec][sampleIdx] || '').trim();
      if (ecSample && !newCompletedSamples[ecSample]) {
        // This row was permanently moved — not in All Orders, preserve it
        var ecRow = existingCompleted[ec].slice();
        for (var ef = 0; ef < existingFormulas[ec].length; ef++) {
          if (existingFormulas[ec][ef]) ecRow[ef] = existingFormulas[ec][ef];
        }
        preservedRows.push(ecRow);
      }
    }
  }

  if (completedSheet.getLastRow() > 1) {
    completedSheet.getRange(2, 1, completedSheet.getLastRow() - 1, completedSheet.getLastColumn() || totalCols).clear();
  }
  var allCompletedRows = completedRows.concat(preservedRows);
  if (allCompletedRows.length > 0) {
    completedSheet.getRange(2, 1, allCompletedRows.length, totalCols).setValues(allCompletedRows);
  }
  Logger.log('Completed Orders updated: ' + completedRows.length + ' from All Orders + ' + preservedRows.length + ' preserved (permanently moved)');

  // Update user view sheets
  var userSheets = ['View - Danboy1217'];
  userSheets.forEach(function(sheetName) {
    var userSheet = ss.getSheetByName(sheetName);
    if (userSheet) {
      // Ensure headers exist
      if (userSheet.getLastColumn() < totalCols || String(userSheet.getRange(1, 1).getValue()).trim() === '') {
        userSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        userSheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold').setFontColor('#000000').setBackground('#d9ead3');
        userSheet.setFrozenRows(1);
        try { userSheet.getRange(1, 1, 1, headers[0].length).createFilter(); } catch(e) { logError('setupOrders', 'Failed to create filter on user sheet', { user: userEmail, error: e.message }); }
      }
      if (userSheet.getLastRow() > 1) {
        userSheet.getRange(2, 1, userSheet.getLastRow() - 1, userSheet.getLastColumn() || totalCols).clear();
      }
      if (liveRows.length > 0) {
        userSheet.getRange(2, 1, liveRows.length, totalCols).setValues(liveRows);
        var printIdx2 = col['Print'];
        if (printIdx2 !== undefined) {
          userSheet.getRange(2, printIdx2 + 1, liveRows.length, 1).insertCheckboxes();
        }
      }
      Logger.log('User sheet "' + sheetName + '" updated: ' + liveRows.length + ' rows');
    }
  });
}

// ============================================================
// OPEN ORDERS REPORT
// ============================================================

function generateOpenOrdersReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();
  
  if (!mainSheet) { ui.alert('Main sheet not found!'); return; }
  
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data in system!'); return; }

  // V5: perf — read headers + data in one getValues() call (2 API calls → 1)
  const allRows = mainSheet.getRange(1, 1, lastRow, mainSheet.getLastColumn()).getValues();
  const headers = allRows[0];
  const data = allRows.slice(1);
  
  const statusColIndex = headers.indexOf('Status');
  const csSampleColIndex = headers.indexOf('CS Sample #');
  const csOrderColIndex = headers.indexOf('CS Order #');
  const sampleOrderColIndex = headers.indexOf('Sample Order #');
  const senderColIndex = headers.indexOf('Sender');
  const descColIndex = headers.indexOf('Description');
  const containerColIndex = headers.indexOf('Container #');
  const cargoColIndex = headers.indexOf('Cargo #');
  const markColIndex = headers.indexOf('Mark #');
  const warehouseColIndex = headers.indexOf('Warehouse');
  const timestampColIndex = headers.indexOf('Timestamp');
  
  if (statusColIndex === -1) { ui.alert('Status column not found!'); return; }
  
  const openOrders = data.filter(row => {
    const status = row[statusColIndex];
    return status === 'Received' || status === 'Scanned' || status === '';
  });
  
  if (openOrders.length === 0) { ui.alert('No open orders!'); return; }
  
  const warehouseGroups = {};
  
  openOrders.forEach(row => {
    const warehouse = row[warehouseColIndex] || 'Unknown Warehouse';
    const sampleOrderNum = row[sampleOrderColIndex] || 'Unknown';
    const sender = row[senderColIndex] || 'Unknown Sender';
    
    if (!warehouseGroups[warehouse]) warehouseGroups[warehouse] = {};
    
    const orderKey = sampleOrderNum + '|' + sender;
    if (!warehouseGroups[warehouse][orderKey]) {
      warehouseGroups[warehouse][orderKey] = { sampleOrderNum: sampleOrderNum, sender: sender, samples: [] };
    }
    
    warehouseGroups[warehouse][orderKey].samples.push({
      csSample: row[csSampleColIndex], csOrder: row[csOrderColIndex],
      sampleOrder: row[sampleOrderColIndex], sender: sender,
      description: row[descColIndex], container: row[containerColIndex],
      cargo: row[cargoColIndex], mark: row[markColIndex],
      warehouse: warehouse, status: row[statusColIndex] || 'Received',
      timestamp: row[timestampColIndex]
    });
  });
  
  const sortedWarehouses = Object.keys(warehouseGroups).sort();
  
  var html = '<!DOCTYPE html><html><head><style>' +
    '@media print { @page { size: letter; margin: 0.5in; } body { margin: 0; } .no-print { display: none; } .warehouse-section { page-break-after: always; } .warehouse-section:last-child { page-break-after: auto; } } ' +
    'body { font-family: Arial; padding: 20px; max-width: 900px; margin: 0 auto; } ' +
    'h1 { color: #2c5f2d; border-bottom: 3px solid #4CAF50; padding-bottom: 10px; } ' +
    '.summary { background: #e8f5e9; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 4px solid #4CAF50; } .summary h2 { margin-top: 0; color: #2c5f2d; } ' +
    '.warehouse-section { margin-bottom: 30px; border: 2px solid #4CAF50; border-radius: 8px; overflow: hidden; } ' +
    '.warehouse-header { background: #4CAF50; color: white; padding: 15px; font-weight: bold; font-size: 27px; } ' +
    '.warehouse-summary { background: #f1f8f4; padding: 10px 15px; border-bottom: 1px solid #4CAF50; font-size: 21px; } ' +
    '.order-section { background: white; padding: 15px; margin: 0; border-bottom: 1px solid #ddd; } ' +
    '.order-header { background: #f5f5f5; padding: 10px; border-left: 4px solid #66bb6a; margin-bottom: 15px; font-weight: bold; } ' +
    'table { width: 100%; border-collapse: collapse; margin-top: 10px; } ' +
    'th { background: #66bb6a; color: white; padding: 10px; text-align: left; font-size: 18px; } ' +
    'td { padding: 8px; border-bottom: 1px solid #ddd; font-size: 16px; } tr:hover { background: #f9f9f9; } ' +
    '.status-received { background: #fff3cd; color: #856404; padding: 3px 8px; border-radius: 3px; font-size: 15px; font-weight: bold; } ' +
    '.status-scanned { background: #d4edda; color: #155724; padding: 3px 8px; border-radius: 3px; font-size: 15px; font-weight: bold; } ' +
    'button { padding: 12px 24px; font-size: 24px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; margin: 10px 5px; } ' +
    '.no-print { text-align: center; margin-bottom: 20px; }' +
    '</style></head><body>' +
    '<div class="no-print"><button onclick="window.print()">🖨️ Print Report</button><button onclick="window.close()">Close</button></div>' +
    '<h1>☕ Open Orders Report</h1>' +
    '<div class="summary"><h2>Summary</h2>' +
    '<p><strong>Total Open Samples:</strong> ' + openOrders.length + '</p>' +
    '<p><strong>Warehouses:</strong> ' + sortedWarehouses.length + '</p>' +
    '<p><strong>Report Date:</strong> ' + new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString() + '</p></div>';

  sortedWarehouses.forEach(function(warehouse) {
    var orders = warehouseGroups[warehouse];
    var sortedOrderKeys = Object.keys(orders).sort(function(a, b) { return orders[a].sender.toUpperCase().localeCompare(orders[b].sender.toUpperCase()); });
    var warehouseSampleCount = 0;
    sortedOrderKeys.forEach(function(key) { warehouseSampleCount += orders[key].samples.length; });
    
    html += '<div class="warehouse-section"><div class="warehouse-header">📍 ' + warehouse + '</div>' +
      '<div class="warehouse-summary"><strong>Orders:</strong> ' + sortedOrderKeys.length + ' | <strong>Samples:</strong> ' + warehouseSampleCount + '</div>';
    
    sortedOrderKeys.forEach(function(orderKey) {
      var orderData = orders[orderKey];
      html += '<div class="order-section"><div class="order-header">📦 ' + orderData.sampleOrderNum + ' - ' + orderData.sender + '</div>' +
        '<p><strong>Samples:</strong> ' + orderData.samples.length + '</p>' +
        '<table><thead><tr><th>CS Sample #</th><th>Status</th><th>Container #</th><th>Mark #</th><th>Cargo #</th><th>Description</th></tr></thead><tbody>';
      orderData.samples.forEach(function(sample) {
        var statusClass = sample.status === 'Scanned' ? 'status-scanned' : 'status-received';
        html += '<tr><td><strong>' + sample.csSample + '</strong></td><td><span class="' + statusClass + '">' + sample.status + '</span></td>' +
          '<td>' + (sample.container || 'N/A') + '</td><td>' + (sample.mark || 'N/A') + '</td>' +
          '<td>' + (sample.cargo || 'N/A') + '</td><td>' + (sample.description || 'N/A') + '</td></tr>';
      });
      html += '</tbody></table></div>';
    });
    html += '</div>';
  });
  
  html += '</body></html>';
  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(950).setHeight(700);
  ui.showModalDialog(htmlOutput, 'Open Orders Report - ' + openOrders.length + ' samples');
}

// ============================================================
// PRINTABLE ORDER FORM
// ============================================================

function generatePrintableOrderForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet) { ui.alert('Sheet not found!'); return; }
  
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  if (startRow === 1) { ui.alert('Please select data rows (not header).'); return; }
  
  var col = _getColumnMap(sheet);
  const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  
  const orderGroups = {};
  data.forEach(function(row) {
    const sampleOrderNum = row[col['Sample Order #']] || 'Unknown';
    if (!orderGroups[sampleOrderNum]) orderGroups[sampleOrderNum] = [];
    orderGroups[sampleOrderNum].push({
      timestamp: row[col['Timestamp']], csOrderNum: row[col['CS Order #']],
      csSampleNum: row[col['CS Sample #']], sender: row[col['Sender']],
      receiver: row[col['Receiver']], warehouse: row[col['Warehouse']],
      description: row[col['Description']], sampleOrderNum: row[col['Sample Order #']],
      cargo: row[col['Cargo #']], mark: row[col['Mark #']],
      container: row[col['Container #']], reference: row[col['Reference']],
      bagCount: row[col['Bag Count']], weight: row[col['Weight']],
      sampleWeight: row[col['Sample Weight']], pNumber: row[col['P #']],
      sNumber: row[col['S #']]
    });
  });
  
  var html = '<!DOCTYPE html><html><head><style>' +
    '@media print { @page { size: letter; margin: 0.5in; } body { margin: 0; } .order-form { page-break-after: always; } .order-form:last-child { page-break-after: auto; } .no-print { display: none; } } ' +
    'body { font-family: Arial; font-size: 16px; margin: 20px; background: #f5f5f5; } .no-print { text-align: center; margin-bottom: 20px; } ' +
    'button { padding: 12px 24px; font-size: 24px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; } ' +
    '.order-form { background: white; padding: 20px; margin-bottom: 20px; border: 1px solid #ddd; min-height: 10in; } ' +
    '.order-header { font-size: 21px; font-weight: bold; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #333; } ' +
    '.sample-row { margin: 15px 0; padding: 15px 0; border-bottom: 1px solid #ddd; } .sample-header { display: flex; justify-content: space-between; margin-bottom: 10px; } ' +
    '.sample-id { font-weight: bold; font-size: 18px; } .sample-details { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 15px; margin-left: 20px; } ' +
    '.detail-row { display: flex; } .detail-label { font-weight: bold; min-width: 140px; }' +
    '</style></head><body>' +
    '<div class="no-print"><h2>Printable Order Forms</h2><p><strong>' + Object.keys(orderGroups).length + ' form(s)</strong></p><button onclick="window.print()">Print All Order Forms</button></div>';

  Object.keys(orderGroups).sort().forEach(function(sampleOrderNum) {
    var samples = orderGroups[sampleOrderNum];
    var firstSample = samples[0];
    var date = new Date(firstSample.timestamp);
    var formattedDate = String(date.getMonth() + 1).padStart(2, '0') + '/' + String(date.getDate()).padStart(2, '0') + '/' + date.getFullYear();
    
    html += '<div class="order-form"><div class="order-header">' + sampleOrderNum + ' - ' + (firstSample.sender || 'Unknown Sender') + '</div>' +
      '<div style="font-style:italic;color:#666;margin-bottom:15px;">' + formattedDate + '</div>';

    samples.forEach(function(sample) {
      var eta = new Date(sample.timestamp); eta.setDate(eta.getDate() + 30);
      var etaFormatted = String(eta.getMonth() + 1).padStart(2, '0') + '/' + String(eta.getDate()).padStart(2, '0') + '/' + eta.getFullYear();
      
      html += '<div class="sample-row"><div class="sample-header"><div class="sample-id">' + (sample.csSampleNum || sample.sampleOrderNum) + '</div>' +
        '<div style="font-weight:bold;text-align:right">ETA: ' + etaFormatted + '</div></div><div class="sample-details">' +
        '<div class="detail-row"><span class="detail-label">Description:</span><span>' + (sample.description || 'N/A') + '</span></div>' +
        '<div class="detail-row"><span class="detail-label">Warehouse:</span><span>' + (sample.warehouse || 'N/A') + '</span></div>' +
        '<div class="detail-row"><span class="detail-label">Mark #:</span><span>' + (sample.mark || 'N/A') + '</span></div>' +
        '<div class="detail-row"><span class="detail-label">Container #:</span><span>' + (sample.container || 'N/A') + '</span></div>' +
        '<div class="detail-row"><span class="detail-label">Cargo #:</span><span>' + (sample.cargo || 'N/A') + '</span></div>' +
        '<div class="detail-row"><span class="detail-label">Bags:</span><span>' + (sample.bagCount || 'N/A') + '</span></div>' +
        '<div class="detail-row"><span class="detail-label">Reference:</span><span>' + (sample.reference || 'N/A') + '</span></div>' +
        '<div class="detail-row"><span class="detail-label">Sample Weight:</span><span>' + (sample.sampleWeight || 'N/A') + '</span></div>' +
        '</div></div>';
    });
    html += '</div>';
  });
  
  html += '</body></html>';
  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(900).setHeight(700);
  ui.showModalDialog(htmlOutput, 'Printable Order Forms - ' + Object.keys(orderGroups).length + ' form(s)');
}

// ============================================================
// PROCESS PENDING MOVES (Shipped → Completed Orders)
// ============================================================

function processPendingMoves() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  var completedSheet = ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders');
  
  if (!mainSheet) return { moved: 0, message: 'Main sheet not found' };
  
  SpreadsheetApp.flush();
  
  var mainLastCol = mainSheet.getLastColumn();
  var mainHeaders = mainSheet.getRange(1, 1, 1, mainLastCol).getValues()[0];
  
  if (!completedSheet) {
    completedSheet = ss.insertSheet(CONFIG.completedOrdersSheetName || 'Completed Orders');
    completedSheet.getRange(1, 1, 1, mainHeaders.length).setValues([mainHeaders]);
    completedSheet.getRange(1, 1, 1, mainHeaders.length).setFontWeight('bold').setBackground('#f4cccc');
    completedSheet.setFrozenRows(1);
  } else {
    if (completedSheet.getLastColumn() < mainLastCol) {
      completedSheet.getRange(1, 1, 1, mainHeaders.length).setValues([mainHeaders]);
      completedSheet.getRange(1, 1, 1, mainHeaders.length).setFontWeight('bold').setBackground('#f4cccc');
    }
  }
  
  var col = _getColumnMap(mainSheet);
  var statusIdx = col['Status'];
  var sampleIdx = col['CS Sample #'];
  var printIdx = col['Print'];
  
  if (statusIdx === undefined) return { moved: 0, message: 'Status column not found' };
  if (sampleIdx === undefined) return { moved: 0, message: 'CS Sample # column not found' };
  
  var lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return { moved: 0, message: 'No data in main sheet' };
  
  var data = mainSheet.getRange(2, 1, lastRow - 1, mainLastCol).getValues();
  
  // Collect shipped samples by CS Sample #
  var shippedSamples = [];
  
  for (var i = 0; i < data.length; i++) {
    var rowStatus = String(data[i][statusIdx] || '').trim().toLowerCase();
    var sampleNum = String(data[i][sampleIdx] || '').trim();
    
    if (rowStatus === 'shipped' && sampleNum) {
      shippedSamples.push({ sample: sampleNum, rowData: data[i].slice() });
    }
  }
  
  if (shippedSamples.length === 0) return { moved: 0, message: 'No shipped samples to move' };
  
  // Batch copy to Completed — update status from Shipped → Completed
  var rowsToAppend = shippedSamples.map(function(s) {
    var row = s.rowData.slice();
    if (statusIdx !== undefined) row[statusIdx] = 'Completed';
    return row;
  });
  var newStartRow = completedSheet.getLastRow() + 1;
  completedSheet.getRange(newStartRow, 1, rowsToAppend.length, mainLastCol).setValues(rowsToAppend);

  if (printIdx !== undefined) {
    completedSheet.getRange(newStartRow, printIdx + 1, rowsToAppend.length, 1).insertCheckboxes();
  }

  SpreadsheetApp.flush();

  // Auto-bill the completed samples
  var autoBillFailed = false;
  try {
    if (typeof _autoBillSamples === 'function') {
      _autoBillSamples(rowsToAppend, ss);
    }
  } catch (e) {
    autoBillFailed = true;
    Logger.log('Auto-billing FAILED: ' + e);
    if (typeof logError === 'function') {
      logError('processPendingMoves', 'Auto-billing failed for ' + shippedSamples.length + ' samples — these need manual billing', {
        samples: shippedSamples.map(function(s) { return s.sample; }).join(', '),
        error: e.message
      });
    }
  }

  // Build set of sample #s to delete
  var toDelete = {};
  for (var j = 0; j < shippedSamples.length; j++) {
    toDelete[shippedSamples[j].sample] = true;
  }
  
  // Re-read main sheet (rows may have shifted if anything else changed)
  // Find and delete by CS Sample # — bottom-up
  var freshLastRow = mainSheet.getLastRow();
  if (freshLastRow >= 2) {
    var freshSamples = mainSheet.getRange(2, sampleIdx + 1, freshLastRow - 1, 1).getValues();
    var deleteCount = 0;
    
    for (var k = freshSamples.length - 1; k >= 0; k--) {
      var id = String(freshSamples[k][0] || '').trim();
      if (toDelete[id]) {
        mainSheet.deleteRow(k + 2);
        deleteCount++;
        delete toDelete[id]; // only delete first match per sample
      }
    }
  }
  
  SpreadsheetApp.flush();
  try { updateLiveOrdersView(); } catch(e) { logError('markSamplesShipped', 'Error updating Live Orders view', { samplesCount: shippedSamples.length, error: e.message }); }

  var remaining = Object.keys(toDelete).length;
  var msg = 'Moved ' + shippedSamples.length + ' samples to Completed Orders';
  if (remaining > 0) msg += ' (' + remaining + ' could not be deleted from main sheet)';
  if (autoBillFailed) msg += '\n⚠️ Auto-billing failed — run Billing → Generate Invoice manually for these samples';

  return { moved: shippedSamples.length, message: msg, billingFailed: autoBillFailed };
}

function processPendingMovesWithAlert() {
  var result = processPendingMoves();
  SpreadsheetApp.getUi().alert(result.message);
}

// ============================================================
// ARCHIVE OLD ORDERS
// ============================================================

function archiveOldOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const completedSheet = ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders');
  const ui = SpreadsheetApp.getUi();
  
  if (!completedSheet) { ui.alert('Completed Orders sheet not found!'); return; }
  
  var cutoffDays = 90;
  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - cutoffDays);
  
  var col = _getColumnMap(completedSheet);
  var lastRow = completedSheet.getLastRow();
  
  if (lastRow < 2) { ui.alert('No completed orders to archive.'); return; }
  
  var data = completedSheet.getRange(2, 1, lastRow - 1, completedSheet.getLastColumn()).getValues();
  
  var oldRows = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var timestamp = new Date(data[i][col['Timestamp']]);
    if (timestamp < cutoffDate) {
      oldRows.push(i + 2);
    }
  }
  
  if (oldRows.length === 0) { ui.alert('No orders older than ' + cutoffDays + ' days found.'); return; }
  
  var response = ui.alert(
    'Archive Old Orders',
    'Found ' + oldRows.length + ' orders older than ' + cutoffDays + ' days.\n\nThis will:\n1. Export them to a CSV file\n2. Delete them from Completed Orders\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;

  // ── Billing integrity check: warn if any sample about to be archived has no Invoice Line Item ──
  // Archiving a sample without billing means the invoice for that work is permanently lost.
  var billedSamplesSet = {};
  try {
    var liSheet = ss.getSheetByName('Invoice Line Items');
    if (liSheet && liSheet.getLastRow() >= 2) {
      var liData = liSheet.getRange(2, 5, liSheet.getLastRow() - 1, 1).getValues();
      for (var bi = 0; bi < liData.length; bi++) {
        var bs = String(liData[bi][0]).trim();
        if (bs) billedSamplesSet[bs] = true;
      }
    }
  } catch (billingCheckErr) {
    Logger.log('archiveOldOrders: billing check error (continuing): ' + billingCheckErr);
  }

  var unbilledInArchive = [];
  var csSampleColIdx = col['CS Sample #'];
  if (csSampleColIdx !== undefined) {
    for (var ub = 0; ub < oldRows.length; ub++) {
      var dataIdxUb = oldRows[ub] - 2;
      var sampleNumUb = String(data[dataIdxUb][csSampleColIdx] || '').trim();
      if (sampleNumUb && !billedSamplesSet[sampleNumUb]) {
        unbilledInArchive.push(sampleNumUb);
      }
    }
  }

  if (unbilledInArchive.length > 0) {
    logWarning('archiveOldOrders',
      'Archiving ' + unbilledInArchive.length + ' sample(s) with no Invoice Line Item',
      { count: unbilledInArchive.length, samples: unbilledInArchive.slice(0, 20).join(', ') });
    Logger.log('archiveOldOrders BILLING WARNING: ' + unbilledInArchive.length + ' unbilled samples: ' +
      unbilledInArchive.slice(0, 10).join(', '));
    var billingResp = ui.alert(
      '\u26a0\ufe0f Unbilled Samples Detected',
      unbilledInArchive.length + ' sample(s) being archived have no billing record in Invoice Line Items.\n\n' +
      'First ' + Math.min(10, unbilledInArchive.length) + ': ' + unbilledInArchive.slice(0, 10).join(', ') + '\n\n' +
      'Consider running Billing \u2192 Generate Invoice first.\n\nContinue archiving anyway?',
      ui.ButtonSet.YES_NO
    );
    if (billingResp !== ui.Button.YES) return;
  }

  var lastColForCsv = completedSheet.getLastColumn();
  var headers = completedSheet.getRange(1, 1, 1, lastColForCsv).getValues()[0];
  var csvContent = headers.map(escapeCSVField).join(',') + '\n';

  // Batch-read all archived rows in one API call instead of one call per row.
  // data[] is already loaded above; oldRows contains 1-based sheet row numbers (descending).
  // Map them back to 0-based data indices (rowNum - 2) and collect in forward order for CSV.
  var oldRowIndices = oldRows.slice().reverse(); // ascending order for CSV
  oldRowIndices.forEach(function(rowNum) {
    var dataIdx = rowNum - 2; // row 2 → index 0
    csvContent += data[dataIdx].map(escapeCSVField).join(',') + '\n';
  });
  
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var fileName = 'CSS_Archived_Orders_' + dateStr + '.csv';
  var file = DriveApp.getRootFolder().createFile(fileName, csvContent, MimeType.CSV);
  
  // CRITICAL: Delete rows in REVERSE order to avoid index shift corruption
  oldRows.sort(function(a, b) { return b - a; });
  oldRows.forEach(function(rowNum) {
    completedSheet.deleteRow(rowNum);
  });
  
  ui.alert('Archive Complete!\n\n' +
    'Archived ' + oldRows.length + ' orders to:\n' +
    fileName + '\n\nFile URL: ' + file.getUrl());
}

// ============================================================
// ORDER/SAMPLE NUMBER GENERATION & DATA ENTRY
// ============================================================

function _buildOrderCache(sheet) {
  var cache = {
    highestOrderNumber: 211656,
    orders: {}
  };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return cache;

  var col = _getColumnMap(sheet);
  var csOrderCol = col['CS Order #'];
  var csSampleCol = col['CS Sample #'];
  var sampleOrdCol = col['Sample Order #'];

  // If essential columns are missing, return default cache (safe: lock in addDataToSheet will validate)
  if (csOrderCol === undefined) return cache;

  var numRows = Math.min(500, lastRow - 1);
  var startRow = lastRow - numRows + 1;
  var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var rowCSOrder = data[i][csOrderCol];
    var rowCSSample = (csSampleCol !== undefined) ? data[i][csSampleCol] : null;
    var rowSampleOrder = (sampleOrdCol !== undefined) ? data[i][sampleOrdCol] : null;

    if (rowCSOrder) {
      var orderNum = parseInt(String(rowCSOrder).replace(/\D/g, ''));
      if (!isNaN(orderNum) && orderNum > cache.highestOrderNumber) {
        cache.highestOrderNumber = orderNum;
      }

      var soKey = String(rowSampleOrder || '').trim();
      if (soKey) {
        if (!cache.orders[soKey]) {
          cache.orders[soKey] = {
            csOrderNumber: String(rowCSOrder),
            highestSample: 0
          };
        }
        if (rowCSSample) {
          var m = String(rowCSSample).match(/-(\d+)$/);
          if (m) {
            var sn = parseInt(m[1]);
            if (sn > cache.orders[soKey].highestSample) {
              cache.orders[soKey].highestSample = sn;
            }
          }
        }
      }
    }
  }

  return cache;
}

function generateOrderAndSampleNumber(sheet, data, cache) {
  // Fast path: use in-memory cache
  if (cache) {
    var soKey = String(data.sampleOrderNum || '').trim();

    if (soKey && cache.orders[soKey]) {
      var entry = cache.orders[soKey];
      entry.highestSample++;
      var sampleNumber = entry.csOrderNumber + '-' + String(entry.highestSample).padStart(2, '0');
      return { orderNumber: entry.csOrderNumber, sampleNumber: sampleNumber };
    } else {
      cache.highestOrderNumber++;
      var newOrderNum = String(cache.highestOrderNumber);
      if (soKey) {
        cache.orders[soKey] = { csOrderNumber: newOrderNum, highestSample: 1 };
      }
      return { orderNumber: newOrderNum, sampleNumber: newOrderNum + '-01' };
    }
  }

  // Original path: read from sheet
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { orderNumber: '211657', sampleNumber: '211657-01' };
  }

  var col = _getColumnMap(sheet);
  var csOrderCol = col['CS Order #'];
  var csSampleCol = col['CS Sample #'];
  var sampleOrdCol = col['Sample Order #'];

  const numRowsToCheck = Math.min(500, lastRow - 1);
  const startRow = lastRow - numRowsToCheck + 1;
  const dataRange = sheet.getRange(startRow, 1, numRowsToCheck, sheet.getLastColumn());
  const values = dataRange.getValues();

  const orderMap = {};
  let highestOrderNumber = 211656;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowSampleOrderNum = row[sampleOrdCol];
    const rowCSOrderNumber = row[csOrderCol];
    const rowCSSampleNumber = row[csSampleCol];

    if (rowCSOrderNumber) {
      const orderNum = parseInt(String(rowCSOrderNumber).replace(/\D/g, ''));
      if (!isNaN(orderNum) && orderNum > highestOrderNumber) {
        highestOrderNumber = orderNum;
      }
      if (rowSampleOrderNum && !orderMap[rowSampleOrderNum]) {
        orderMap[rowSampleOrderNum] = String(rowCSOrderNumber);
      }
    }
  }

  const currentSampleOrderNum = String(data.sampleOrderNum || '').trim();

  if (currentSampleOrderNum && orderMap[currentSampleOrderNum]) {
    let highestSampleInOrder = 0;
    const existingCSOrder = orderMap[currentSampleOrderNum];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowSampleOrderNum = String(row[sampleOrdCol] || '').trim();
      const rowCSOrderNumber = String(row[csOrderCol] || '').trim();
      const rowCSSampleNumber = row[csSampleCol];

      if (rowSampleOrderNum === currentSampleOrderNum && rowCSOrderNumber === existingCSOrder) {
        if (rowCSSampleNumber) {
          const sampleMatch = String(rowCSSampleNumber).match(/-(\d+)$/);
          if (sampleMatch) {
            const sampleNum = parseInt(sampleMatch[1]);
            if (sampleNum > highestSampleInOrder) {
              highestSampleInOrder = sampleNum;
            }
          }
        }
      }
    }

    const orderNumber = existingCSOrder;
    const nextSampleNum = highestSampleInOrder + 1;
    const sampleNumber = `${orderNumber}-${String(nextSampleNum).padStart(2, '0')}`;
    return { orderNumber: orderNumber, sampleNumber: sampleNumber };
  } else {
    const nextOrderNum = highestOrderNumber + 1;
    const orderNumber = String(nextOrderNum);
    const sampleNumber = `${orderNumber}-01`;
    return { orderNumber: orderNumber, sampleNumber: sampleNumber };
  }
}

function getNextOrderNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);

  if (!sheet) return { success: false, message: 'Sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, orderNumber: '211657' };

  var col = _getColumnMap(sheet);
  var csOrderIdx = col['CS Order #'];
  if (csOrderIdx === undefined) {
    return { success: false, message: 'CS Order # column not found' };
  }
  const numRowsToCheck = Math.min(500, lastRow - 1);
  const startRow = lastRow - numRowsToCheck + 1;
  const values = sheet.getRange(startRow, csOrderIdx + 1, numRowsToCheck, 1).getValues();

  let highestOrderNumber = 211656;
  for (let i = 0; i < values.length; i++) {
    const orderNum = parseInt(String(values[i][0]).replace(/\D/g, ''));
    if (!isNaN(orderNum) && orderNum > highestOrderNumber) highestOrderNumber = orderNum;
  }

  return { success: true, orderNumber: String(highestOrderNumber + 1) };
}

function getNextSampleInOrder(orderNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);

  if (!sheet) return 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sampleCol = headers.indexOf('CS Sample #');
  const samples = sheet.getRange(2, sampleCol + 1, lastRow - 1, 1).getValues();

  let maxSampleNum = 0;
  const orderPrefix = String(orderNum) + '-';

  samples.forEach(row => {
    const sample = String(row[0]);
    if (sample.startsWith(orderPrefix)) {
      const sampleNum = parseInt(sample.split('-')[1]);
      if (!isNaN(sampleNum) && sampleNum > maxSampleNum) maxSampleNum = sampleNum;
    }
  });

  return maxSampleNum + 1;
}

// ============================================================
// ADD DATA TO SHEET
// ============================================================

function addDataToSheet(sheet, data, options) {
  options = options || {};
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    Logger.log('=== addDataToSheet called ===');

    if (data.receiver) {
      data.receiver = cleanReceiverName(data.receiver);
    }

    if ((!data.description || data.description.trim() === '') && data.mark) {
      const country = getCountryFromICOMark(data.mark);
      if (country) {
        data.description = country;
      }
    }

    // Use _orderGroupKey for grouping if present
    var groupingData = Object.assign({}, data);
    if (data._orderGroupKey) {
      groupingData.sampleOrderNum = data._orderGroupKey;
      delete data._orderGroupKey;
    }

    // --- CRITICAL SECTION: lock to prevent duplicate order/sample numbers ---
    if (!lock.tryLock(15000)) {
      throw new Error('System busy — could not acquire write lock. Try again.');
    }
    lockAcquired = true;

    // If using cache, verify it hasn't gone stale from a concurrent writer
    if (options.orderCache) {
      var _col = _getColumnMap(sheet);
      var _lastRow = sheet.getLastRow();
      if (_lastRow > 1) {
        var _checkRows = Math.min(10, _lastRow - 1);
        var _recent = sheet.getRange(_lastRow - _checkRows + 1, _col['CS Order #'] + 1, _checkRows, 1).getValues();
        for (var _r = 0; _r < _recent.length; _r++) {
          var _n = parseInt(String(_recent[_r][0]).replace(/\D/g, ''));
          if (!isNaN(_n) && _n > options.orderCache.highestOrderNumber) {
            options.orderCache.highestOrderNumber = _n;
            Logger.log('⚠️ Lock: cache was stale — updated highestOrderNumber to ' + _n);
          }
        }
      }
    }

    // Generate order and sample numbers
    const orderInfo = generateOrderAndSampleNumber(sheet, groupingData, options.orderCache);
    const csOrderNumber = orderInfo.orderNumber;
    const csSampleNumber = orderInfo.sampleNumber;

    const fullComments = data.comments || '';

    // Build QR data string
    const qrData = typeof buildQRDataString === 'function'
      ? buildQRDataString(data, csSampleNumber)
      : (data.sampleOrderNum + '|' + data.cargo + '|' + data.mark + '|' + data.container);

    // Get column map FIRST before using it
    var col = _getColumnMap(sheet);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = new Array(headers.length).fill('');

    // Verify all required columns exist before building row
    var requiredCols = ['Timestamp', 'CS Order #', 'CS Sample #', 'Sender', 'Receiver',
                        'Warehouse', 'Description', 'Sample Order #', 'Cargo #', 'Mark #',
                        'Container #', 'Reference', 'Bag Count', 'Weight', 'Sample Weight',
                        'P #', 'S #', 'Shipping Process', 'Comments', 'Source Email', 'QR Data', 'Status'];
    for (var rc = 0; rc < requiredCols.length; rc++) {
      if (col[requiredCols[rc]] === undefined) {
        throw new Error('Required column missing: ' + requiredCols[rc]);
      }
    }

    // Formula injection guard: prepend a single quote to any string
    // starting with =, +, -, or @ so Sheets treats it as plain text.
    function _safeStr(val) {
      if (val === null || val === undefined) return '';
      var s = String(val);
      if (s.length > 0 && '=+-@'.indexOf(s.charAt(0)) !== -1) return "'" + s;
      return s;
    }

    // Build row by header name
    row[col['Timestamp']] = new Date();
    row[col['CS Order #']] = csOrderNumber;
    row[col['CS Sample #']] = csSampleNumber;
    row[col['Sender']] = _safeStr(data.sender);
    row[col['Receiver']] = _safeStr(data.receiver);
    row[col['Warehouse']] = _safeStr(data.warehouse);
    row[col['Description']] = _safeStr(data.description);
    row[col['Sample Order #']] = _safeStr(data.sampleOrderNum);
    row[col['Cargo #']] = _safeStr(data.cargo);
    row[col['Mark #']] = _safeStr(data.mark);
    row[col['Container #']] = _safeStr(data.container);
    row[col['Reference']] = _safeStr(data.reference);
    row[col['Bag Count']] = _safeStr(data.bagCount);
    row[col['Weight']] = _safeStr(data.weight);
    row[col['Sample Weight']] = _safeStr(data.sampleWeight);
    row[col['P #']] = _safeStr(data.pNumber);
    row[col['S #']] = _safeStr(data.sNumber);
    row[col['Shipping Process']] = _safeStr(data.shippingProcess);
    row[col['Comments']] = _safeStr(fullComments);
    row[col['Source Email']] = _safeStr(data.sourceEmail);
    row[col['QR Data']] = _safeStr(qrData);
    row[col['Status']] = 'Received';

    // Print column
    if (col['Print'] !== undefined) {
      row[col['Print']] = false;
    }

    // Flag orders that need manual review
    var needsReview = '';
    if (!data.container) needsReview += 'No container; ';
    if (!data.mark) needsReview += 'No mark; ';
    if (!data.sampleOrderNum) needsReview += 'No order #; ';
    if (data.comments && data.comments.indexOf('REVIEW') >= 0) needsReview += 'Flagged; ';
    if (col['Needs Review'] !== undefined) {
      row[col['Needs Review']] = needsReview ? '⚠️ ' + needsReview.slice(0, -2) : '';
    }

    if (col['Email Link'] !== undefined) row[col['Email Link']] = _safeStr(data.emailLink);
    if (col['Attachments'] !== undefined) row[col['Attachments']] = _safeStr(data.attachments);
    if (col['Photos'] !== undefined) row[col['Photos']] = _safeStr(data.photos);
    if (col['Container Status'] !== undefined) row[col['Container Status']] = _safeStr(data.containerStatus);
    if (col['Container ETA'] !== undefined) row[col['Container ETA']] = _safeStr(data.containerETA);
    if (col['Sample Type'] !== undefined) row[col['Sample Type']] = _safeStr(data.sampleType);
    if (col['Shipping Line'] !== undefined) row[col['Shipping Line']] = _safeStr(data.shippingLine);
    if (col['Shipping Notes'] !== undefined) row[col['Shipping Notes']] = _safeStr(data.shippingNotes);
    if (col['B/L #'] !== undefined) row[col['B/L #']] = _safeStr(data.bol);
    if (col['Ship Status'] !== undefined) row[col['Ship Status']] = _safeStr(data.shipStatus);

    // Append the row
    sheet.appendRow(row);

    // Capture the row number while still under lock — prevents a concurrent writer
    // from appending another row before we read lastRow for checkbox/hyperlink inserts.
    var lastRow = sheet.getLastRow();

    // Add checkbox to Print column for new row (inside lock window)
    var printColIdx = col['Print'];
    if (printColIdx !== undefined) {
      sheet.getRange(lastRow, printColIdx + 1).insertCheckboxes();
    }

    // Format tracking link if present (inside lock window)
    if (data.trackingNumber) {
      var trackingCol = col['Tracking Number'];
      if (trackingCol !== undefined) {
        var url = getTrackingUrl(data.trackingNumber);
        if (url) {
          sheet.getRange(lastRow, trackingCol + 1).setFormula('=HYPERLINK("' + url + '","' + data.trackingNumber + '")');
        }
      }
    }

    Logger.log('✓ Row added: ' + data.container + ' | CS Order: ' + csOrderNumber + ' | Sample: ' + csSampleNumber + ' | Status: Received');

    // --- END CRITICAL SECTION: release lock ---
    lock.releaseLock();
    lockAcquired = false;

    // Skip live update in batch mode
    if (!options.skipLiveUpdate) {
      try {
        updateLiveOrdersView();
      } catch (e) {
        Logger.log('Live Orders update failed (continuing): ' + e);
      }
    }

    // Auto-print if not skipped
    if (!options.skipAutoPrint) {
      try {
        if (typeof autoPrintLabel === 'function') {
          autoPrintLabel(row, col);
        }
      } catch (e) {
        Logger.log('Label print failed (continuing): ' + e);
      }
    }

  } catch (e) {
    Logger.log('ERROR in addDataToSheet: ' + e);
    Logger.log('Data: ' + JSON.stringify(data));
    throw e;
  } finally {
    if (lockAcquired) {
      try { lock.releaseLock(); } catch (releaseErr) {
        logError('addDataToSheet', 'Error releasing lock', { error: releaseErr.message });
      }
    }
  }
}

// ============================================================
// DELETE SELECTED ROWS (syncs across All Orders, Live Orders, User tabs)
// ============================================================

function deleteSelectedRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var sheetName = sheet.getName();

  // Only allow from synced sheets
  var allowedSheets = [CONFIG.mainSheetName, CONFIG.liveOrdersSheetName, 'View - Danboy1217'];
  if (allowedSheets.indexOf(sheetName) === -1) {
    ui.alert('Delete sync only works from All Orders, Live Orders, or User view tabs.');
    return;
  }

  var selection = sheet.getActiveRange();
  if (!selection || selection.getRow() < 2) {
    ui.alert('Select one or more data rows first (not the header).');
    return;
  }

  var col = _getColumnMap(sheet);
  var sampleIdx = col['CS Sample #'];
  if (sampleIdx === undefined) {
    ui.alert('CS Sample # column not found on this sheet.');
    return;
  }

  // Collect CS Sample #s from selected rows
  var startRow = selection.getRow();
  var numRows = selection.getNumRows();
  var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  var samplesToDelete = [];
  for (var i = 0; i < data.length; i++) {
    var sampleNum = String(data[i][sampleIdx] || '').trim();
    if (sampleNum) samplesToDelete.push(sampleNum);
  }

  if (samplesToDelete.length === 0) {
    ui.alert('No CS Sample #s found in selected rows.');
    return;
  }

  var answer = ui.alert(
    '⚠️ Delete ' + samplesToDelete.length + ' Sample(s)',
    'This will delete from All Orders, Live Orders, and User tabs:\n\n' +
    samplesToDelete.join('\n') +
    '\n\nThis cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  // Always delete from main sheet (source of truth)
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!mainSheet) { ui.alert('Main sheet not found!'); return; }

  var mainCol = _getColumnMap(mainSheet);
  var mainSampleIdx = mainCol['CS Sample #'];
  var mainLastRow = mainSheet.getLastRow();
  var deleteCount = 0;

  if (mainLastRow >= 2) {
    var mainData = mainSheet.getRange(2, mainSampleIdx + 1, mainLastRow - 1, 1).getValues();

    // Build list of row numbers to delete (bottom-up to preserve indices)
    var rowsToDelete = [];
    for (var i = 0; i < mainData.length; i++) {
      var val = String(mainData[i][0] || '').trim();
      if (samplesToDelete.indexOf(val) !== -1) {
        rowsToDelete.push(i + 2);
      }
    }

    // Delete bottom-up
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var j = 0; j < rowsToDelete.length; j++) {
      mainSheet.deleteRow(rowsToDelete[j]);
      deleteCount++;
    }
  }

  // Rebuild Live Orders + User tabs from main sheet
  try {
    updateLiveOrdersView();
  } catch (e) {
    Logger.log('Live Orders rebuild failed: ' + e);
  }

  ui.alert('✅ Deleted ' + deleteCount + ' row(s) from All Orders.\nLive Orders and User tabs refreshed.');
}

// ============================================================
// REPAIR RECEIVERS — Re-extract receiver from original emails
// using the corrected parsers for Serengeti, Sucafina, Covoya
// ============================================================

function repairReceivers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!sheet) { SpreadsheetApp.getUi().alert('No "All Orders" sheet found.'); return; }

  var col = _getColumnMap(sheet);
  var senderCol = col['Sender'];
  var receiverCol = col['Receiver'];
  var emailLinkCol = col['Email Link'];
  var cargoCol = col['Cargo #'];
  var containerCol = col['Container #'];

  if (senderCol === undefined || receiverCol === undefined) {
    SpreadsheetApp.getUi().alert('Missing Sender or Receiver column.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data rows.'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Target senders that had parser bugs
  var targetSenders = /serengeti|sucafina|covoya/i;

  // Group rows by email thread to avoid re-fetching the same email
  var threadGroups = {};
  var targetRows = [];

  for (var i = 0; i < data.length; i++) {
    var sender = (data[i][senderCol] || '').toString();
    if (!targetSenders.test(sender)) continue;

    var emailLink = emailLinkCol !== undefined ? (data[i][emailLinkCol] || '').toString() : '';
    var threadId = '';
    if (emailLink) {
      var tidMatch = emailLink.match(/\/([a-f0-9]+)\s*$/i);
      if (tidMatch) threadId = tidMatch[1];
    }

    targetRows.push({
      rowIndex: i,
      sheetRow: i + 2,
      sender: sender,
      currentReceiver: (data[i][receiverCol] || '').toString(),
      cargo: cargoCol !== undefined ? (data[i][cargoCol] || '').toString() : '',
      container: containerCol !== undefined ? (data[i][containerCol] || '').toString() : '',
      threadId: threadId
    });

    if (threadId && !threadGroups[threadId]) {
      threadGroups[threadId] = null; // placeholder, will fetch
    }
  }

  if (targetRows.length === 0) {
    SpreadsheetApp.getUi().alert('No Serengeti/Sucafina/Covoya orders found to repair.');
    return;
  }

  // Fetch unique threads and re-extract orders
  var threadIds = Object.keys(threadGroups);
  Logger.log('repairReceivers: Found ' + targetRows.length + ' rows across ' + threadIds.length + ' threads');

  for (var t = 0; t < threadIds.length; t++) {
    try {
      var thread = GmailApp.getThreadById(threadIds[t]);
      if (!thread) continue;
      var messages = thread.getMessages();
      var message = messages[0]; // first message in thread

      var emailBody = message.getPlainBody();
      var subject = message.getSubject();
      var senderEmail = message.getFrom();

      // Extract PDF text
      var pdfText = '';
      var attachments = message.getAttachments();
      for (var a = 0; a < attachments.length; a++) {
        if (attachments[a].getContentType() === 'application/pdf') {
          try {
            var text = extractTextFromPDF(attachments[a]);
            if (text) pdfText += text + '\n\n';
          } catch (e) { /* skip bad PDF */ }
        }
      }

      // Re-extract using fixed parsers
      var result = extractOrderFromEmail(emailBody, pdfText, senderEmail, subject);
      if (result.success && result.orders.length > 0) {
        // Build a map: cargo/container → receiver
        var receiverMap = [];
        for (var o = 0; o < result.orders.length; o++) {
          var order = result.orders[o];
          var recv = (order.receiver && typeof order.receiver === 'object')
            ? (order.receiver.name || order.receiver.address || '')
            : (order.receiver || '');
          receiverMap.push({
            cargo: (order.cargo || '').toUpperCase(),
            container: (order.container || '').toUpperCase(),
            receiver: recv
          });
        }
        threadGroups[threadIds[t]] = receiverMap;
      }
    } catch (e) {
      Logger.log('repairReceivers: Error fetching thread ' + threadIds[t] + ': ' + e);
    }
  }

  // Now match rows to re-extracted receivers and update
  var updates = [];
  for (var r = 0; r < targetRows.length; r++) {
    var row = targetRows[r];
    if (!row.threadId || !threadGroups[row.threadId]) continue;

    var receiverMap = threadGroups[row.threadId];
    var newReceiver = '';

    // Try matching by cargo # first, then container #
    var rowCargo = row.cargo.toUpperCase();
    var rowContainer = row.container.toUpperCase();

    for (var m = 0; m < receiverMap.length; m++) {
      if (rowCargo && receiverMap[m].cargo && rowCargo === receiverMap[m].cargo) {
        newReceiver = receiverMap[m].receiver;
        break;
      }
    }
    if (!newReceiver) {
      for (var m2 = 0; m2 < receiverMap.length; m2++) {
        if (rowContainer && receiverMap[m2].container && rowContainer === receiverMap[m2].container) {
          newReceiver = receiverMap[m2].receiver;
          break;
        }
      }
    }
    // Fallback: if only one receiver extracted, use it for all rows from this thread
    if (!newReceiver && receiverMap.length > 0) {
      var uniqueReceivers = [];
      for (var u = 0; u < receiverMap.length; u++) {
        if (receiverMap[u].receiver && uniqueReceivers.indexOf(receiverMap[u].receiver) === -1) {
          uniqueReceivers.push(receiverMap[u].receiver);
        }
      }
      if (uniqueReceivers.length === 1) newReceiver = uniqueReceivers[0];
    }

    if (newReceiver && newReceiver !== row.currentReceiver) {
      updates.push({
        sheetRow: row.sheetRow,
        oldReceiver: row.currentReceiver,
        newReceiver: newReceiver,
        sender: row.sender,
        cargo: row.cargo
      });
      // Write the fix
      sheet.getRange(row.sheetRow, receiverCol + 1).setValue(newReceiver);
    }
  }

  // Report
  var msg = '🔧 Receiver Repair Complete\n\n';
  msg += 'Scanned: ' + targetRows.length + ' rows\n';
  msg += 'Updated: ' + updates.length + ' rows\n\n';
  for (var u2 = 0; u2 < updates.length; u2++) {
    msg += 'Row ' + updates[u2].sheetRow + ': "' + updates[u2].oldReceiver + '" → "' + updates[u2].newReceiver + '"';
    if (updates[u2].cargo) msg += ' (Cargo: ' + updates[u2].cargo + ')';
    msg += '\n';
  }

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);

  // Refresh Live Orders if anything changed
  if (updates.length > 0) {
    try { updateLiveOrdersView(); } catch (e) {
      logError('repairReceiverInfo', 'Error refreshing Live Orders after repair', { updatesCount: updates.length, error: e.message });
    }
  }
}