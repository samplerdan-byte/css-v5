// ============================================================
// Reports.gs — End of day report, field report system
//
// FIXED Feb 15, 2026:
//   - Added generateOfflineFieldReport() — was called by the
//     Field Report sidebar "Save Offline" button but never defined
// ============================================================

// ============================================================
// END OF DAY REPORT
// ============================================================

function generateEndOfDayReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found!'); return; }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  var col = _getColumnMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows found.'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let received = 0, scanned = 0, shipped = 0;
  const customerCounts = {};
  const warehouseCounts = {};

  data.forEach(row => {
    const rowDate = new Date(row[col['Timestamp']]);
    rowDate.setHours(0, 0, 0, 0);

    if (rowDate.getTime() === today.getTime()) {
      received++;

      const sender = row[col['Sender']] || 'Unknown';
      customerCounts[sender] = (customerCounts[sender] || 0) + 1;

      const warehouse = row[col['Warehouse']] || 'Unknown';
      warehouseCounts[warehouse] = (warehouseCounts[warehouse] || 0) + 1;
    }

    const scannedDate = row[col['Scanned Date']];
    if (scannedDate) {
      const scanDate = new Date(scannedDate);
      scanDate.setHours(0, 0, 0, 0);
      if (scanDate.getTime() === today.getTime()) {
        scanned++;
      }
    }

    const shippedDate = row[col['Shipped Date']];
    if (shippedDate) {
      const shipDate = new Date(shippedDate);
      shipDate.setHours(0, 0, 0, 0);
      if (shipDate.getTime() === today.getTime()) {
        shipped++;
      }
    }
  });

  let customerBreakdown = '';
  Object.keys(customerCounts).sort((a, b) => customerCounts[b] - customerCounts[a]).forEach(customer => {
    customerBreakdown += `<tr><td>${customer}</td><td>${customerCounts[customer]}</td></tr>`;
  });

  let warehouseBreakdown = '';
  Object.keys(warehouseCounts).sort((a, b) => warehouseCounts[b] - warehouseCounts[a]).forEach(warehouse => {
    warehouseBreakdown += `<tr><td>${warehouse}</td><td>${warehouseCounts[warehouse]}</td></tr>`;
  });

  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
    .report { background: white; padding: 25px; border-radius: 8px; max-width: 600px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
    h1 { color: #2E5339; border-bottom: 3px solid #4A7C59; padding-bottom: 10px; }
    .date { color: #666; font-size: 14pt; margin-bottom: 20px; }
    .stats { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-bottom: 25px; }
    .stat-box { background: #f0f7f2; padding: 20px; border-radius: 8px; text-align: center; border-left: 4px solid #4A7C59; }
    .stat-number { font-size: 32pt; font-weight: bold; color: #2E5339; }
    .stat-label { font-size: 10pt; color: #666; text-transform: uppercase; }
    h3 { color: #4A7C59; margin-top: 25px; border-bottom: 2px solid #4A7C59; padding-bottom: 5px; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th, td { padding: 8px 12px; text-align: left; border-bottom: 1px solid #ddd; }
    th { background: #4A7C59; color: white; }
    tr:hover { background: #f9f9f9; }
    .no-print { text-align: center; margin-top: 20px; }
    button { padding: 12px 24px; font-size: 14pt; background: #4A7C59; color: white; border: none; border-radius: 5px; cursor: pointer; margin: 5px; }
    button:hover { background: #3a6249; }
    @media print { .no-print { display: none; } }
  </style>
</head>
<body>
  <div class="report">
    <h1>☕ End of Day Report</h1>
    <div class="date">${dateStr}</div>
    <div class="stats">
      <div class="stat-box"><div class="stat-number">${received}</div><div class="stat-label">Received</div></div>
      <div class="stat-box"><div class="stat-number">${scanned}</div><div class="stat-label">Scanned</div></div>
      <div class="stat-box"><div class="stat-number">${shipped}</div><div class="stat-label">Shipped</div></div>
    </div>
    <h3>By Customer</h3>
    <table><thead><tr><th>Customer</th><th>Samples</th></tr></thead>
    <tbody>${customerBreakdown || '<tr><td colspan="2">No data</td></tr>'}</tbody></table>
    <h3>By Warehouse</h3>
    <table><thead><tr><th>Warehouse</th><th>Samples</th></tr></thead>
    <tbody>${warehouseBreakdown || '<tr><td colspan="2">No data</td></tr>'}</tbody></table>
    <div class="no-print">
      <button onclick="window.print()">🖨️ Print Report</button>
      <button onclick="google.script.host.close()">Close</button>
    </div>
  </div>
</body>
</html>
`;

  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(700);
  ui.showModalDialog(htmlOutput, 'End of Day Report');
}

// ============================================================
// FIELD REPORT — DATA
// ============================================================

function getWarehouseEmailConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Warehouse Emails');
  if (!sheet || sheet.getLastRow() < 2) return {};

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  var config = {};
  data.forEach(function(row) {
    var wh = String(row[0] || '').trim();
    var emails = String(row[1] || '').trim();
    var active = String(row[2] || '').trim().toUpperCase();
    if (wh && emails && active === 'Y') {
      config[wh] = emails;
    }
  });
  return config;
}

function getFieldReportData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);

  if (!mainSheet || mainSheet.getLastRow() < 2) return { warehouses: {} };

  var col = _getColumnMap(mainSheet);
  var data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();

  var liveStatuses = [CONFIG.statusValues.RECEIVED, CONFIG.statusValues.SCANNED];
  var warehouses = {};

  function safeRead(row, colName) {
    return col[colName] !== undefined ? String(data[row][col[colName]] || '').trim() : '';
  }

  function formatDate(val) {
    if (!val) return '';
    try {
      var d = new Date(val);
      if (isNaN(d.getTime())) return String(val);
      return String(d.getMonth() + 1).padStart(2, '0') + '/' +
        String(d.getDate()).padStart(2, '0') + '/' + d.getFullYear();
    } catch(e) { return String(val); }
  }

  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][col['Status']] || '').trim();
    if (liveStatuses.indexOf(status) === -1) continue;

    var warehouse = safeRead(i, 'Warehouse') || 'Unknown';
    var sender = safeRead(i, 'Sender') || 'Unknown';
    var csOrder = safeRead(i, 'CS Order #') || 'Unknown';
    var receiver = safeRead(i, 'Receiver') || '';

    if (!warehouses[warehouse]) warehouses[warehouse] = {};
    if (!warehouses[warehouse][sender]) warehouses[warehouse][sender] = {};

    if (!warehouses[warehouse][sender][csOrder]) {
      warehouses[warehouse][sender][csOrder] = {
        date: formatDate(data[i][col['Timestamp']]),
        sortDate: data[i][col['Timestamp']] ? new Date(data[i][col['Timestamp']]).getTime() : 0,
        sampleOrderNum: safeRead(i, 'Sample Order #'),
        shippingLine: safeRead(i, 'Shipping Line'),
        shippingNotes: safeRead(i, 'Shipping Notes'),
        vesselStatus: safeRead(i, 'Container Status'),
        eta: formatDate(data[i][col['Container ETA']]),
        emailLink: '',
        firstRow: i + 2,
        receivers: {}
      };
    }

    var order = warehouses[warehouse][sender][csOrder];
    var rowEmailLink = safeRead(i, 'Email Link');
    if (rowEmailLink && !order.emailLink) order.emailLink = rowEmailLink;

    var recvKey = receiver || 'Unassigned';
    if (!order.receivers[recvKey]) order.receivers[recvKey] = [];

    order.receivers[recvKey].push({
      reference: safeRead(i, 'Reference') || safeRead(i, 'Sample Order #'),
      description: safeRead(i, 'Description'),
      mark: safeRead(i, 'Mark #'),
      container: safeRead(i, 'Container #'),
      cargo: safeRead(i, 'Cargo #'),
      sampleType: safeRead(i, 'Sample Type') || safeRead(i, 'Shipping Process'),
      bagCount: safeRead(i, 'Bag Count'),
      sampleWeight: safeRead(i, 'Sample Weight'),
      csSample: safeRead(i, 'CS Sample #'),
      status: status,
      rowNumber: i + 2
    });
  }

  return { warehouses: warehouses };
}

function _typeIcon(type) {
  if (!type) return '';
  var t = type.toLowerCase();
  if (t.indexOf('photo') !== -1) return '📷';
  if (t.indexOf('warehouse') !== -1) return '📦';
  if (t.indexOf('container sup') !== -1) return '🔍';
  if (t.indexOf('exchange') !== -1) return '☕';
  if (t.indexOf('fcc') !== -1) return '🏷️';
  return '';
}

function _typeLabel(type) {
  if (!type) return '';
  var icon = _typeIcon(type);
  return icon ? icon + ' ' + type : type;
}

// ============================================================
// FIELD REPORT — EDITOR PERMISSIONS
// ============================================================

function getEditorPermissions() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return { canEdit: false, columns: [] };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Field Report Editors');
    if (!sheet || sheet.getLastRow() < 2) return { canEdit: false, columns: [] };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var emailIdx = headers.indexOf('Email');
    var colIdx = headers.indexOf('Allowed Columns');
    if (emailIdx === -1) return { canEdit: false, columns: [] };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === email.toLowerCase()) {
        if (colIdx === -1) return { canEdit: true, columns: 'ALL' };
        var colStr = String(data[i][colIdx] || '').trim();
        if (!colStr) return { canEdit: false, columns: [] };
        if (colStr.toUpperCase() === 'ALL') return { canEdit: true, columns: 'ALL' };
        var cols = colStr.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c; });
        return { canEdit: cols.length > 0, columns: cols };
      }
    }
    return { canEdit: false, columns: [] };
  } catch (e) {
    return { canEdit: false, columns: [] };
  }
}

function isFieldReportEditor() {
  return getEditorPermissions().canEdit;
}

function updateFieldReportCell(rowNumber, columnName, newValue) {
  var perms = getEditorPermissions();
  if (!perms.canEdit) return { success: false, message: 'You do not have edit access.' };

  if (perms.columns !== 'ALL') {
    if (perms.columns.indexOf(columnName) === -1) {
      return { success: false, message: 'You don\'t have permission to edit "' + columnName + '".' };
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!sheet) return { success: false, message: 'Sheet not found.' };

  var col = _getColumnMap(sheet);
  var colIdx = col[columnName];
  if (colIdx === undefined) return { success: false, message: 'Column "' + columnName + '" not found in sheet.' };

  try {
    sheet.getRange(rowNumber, colIdx + 1).setValue(newValue);
    return { success: true, message: '✅ Updated ' + columnName };
  } catch (e) {
    return { success: false, message: '❌ Error: ' + e.message };
  }
}

// generateOfflineFieldReport() — moved to fieldreportoffline.js to avoid duplicates

// ============================================================
// FIELD REPORT — DISPLAY (generateFieldReport)
// ============================================================

function generateFieldReport() {
  var reportData = getFieldReportData();
  var emailConfig = getWarehouseEmailConfig();
  var perms = getEditorPermissions();
  var warehouseNames = Object.keys(reportData.warehouses).sort();

  if (warehouseNames.length === 0) {
    SpreadsheetApp.getUi().alert('No active orders found.');
    return;
  }

  var grandTotal = 0;
  var warehouseStats = {};
  warehouseNames.forEach(function(wh) {
    var customers = reportData.warehouses[wh];
    var count = 0;
    var custCount = Object.keys(customers).length;
    for (var c in customers) {
      for (var o in customers[c]) {
        var order = customers[c][o];
        for (var r in order.receivers) count += order.receivers[r].length;
      }
    }
    warehouseStats[wh] = { sampleCount: count, customerCount: custCount };
    grandTotal += count;
  });

  // -- SIDEBAR --
  var sidebarItems = warehouseNames.map(function(wh) {
    var stats = warehouseStats[wh];
    var hasEmail = emailConfig[wh] ? true : false;
    var esc = wh.replace(/'/g, "\\'");
    return '<div class="si" onclick="scrollTo(\'' + esc + '\')">' +
      '<div class="sn">' + wh + '</div>' +
      '<div class="ss">' + stats.sampleCount + ' samples &middot; ' + stats.customerCount + ' clients</div>' +
      '<div class="sa">' +
      '<button class="sb" onclick="event.stopPropagation();printWh(\'' + esc + '\')">🖨️</button>' +
      '<button class="sb' + (hasEmail ? '' : ' dim') + '" onclick="event.stopPropagation();emailWh(\'' + esc + '\',' + hasEmail + ')" title="' + (hasEmail ? emailConfig[wh] : 'No emails') + '">📧</button>' +
      '</div></div>';
  }).join('');

  // -- REPORT BODY --
  var reportSections = warehouseNames.map(function(wh) {
    var customers = reportData.warehouses[wh];
    var customerNames = Object.keys(customers).sort();
    var whId = wh.replace(/[^a-zA-Z0-9]/g, '_');

    var customerHtml = customerNames.map(function(cust) {
      var orders = customers[cust];
      var orderKeys = Object.keys(orders).sort(function(a, b) {
        return (orders[a].sortDate || 0) - (orders[b].sortDate || 0);
      });

      var ordersHtml = orderKeys.map(function(orderNum) {
        var order = orders[orderNum];
        var receiverNames = Object.keys(order.receivers);
        var multiRecv = receiverNames.length > 1;
        var oRow = order.firstRow;

        function edt(row, colName, val, placeholder) {
          var canCol = perms.canEdit && (perms.columns === 'ALL' || perms.columns.indexOf(colName) !== -1);
          if (!canCol) return val || '';
          var esc = colName.replace(/'/g, "\\'");
          if (val) return '<span class="ec" onclick="editCell(this,' + row + ',\'' + esc + '\')" title="Click to edit ' + colName + '">' + val + '</span>';
          return '<span class="ec" onclick="editCell(this,' + row + ',\'' + esc + '\')" title="Click to edit ' + colName + '"><em class="ep">+ ' + (placeholder || colName.toLowerCase()) + '</em></span>';
        }

        var orderNumHtml = orderNum;
        if (order.emailLink) {
          orderNumHtml = '<a href="' + order.emailLink + '" target="_blank" class="on olink" title="Open email">' + orderNum + '</a>';
        } else {
          orderNumHtml = '<span class="on">' + orderNum + '</span>';
        }

        var shipLineHtml = edt(oRow, 'Shipping Line', order.shippingLine, 'line');
        var shipNotesHtml = edt(oRow, 'Shipping Notes', order.shippingNotes, 'notes');
        var vesselStatusHtml = edt(oRow, 'Container Status', order.vesselStatus, 'status');
        var etaHtml = edt(oRow, 'Container ETA', order.eta, 'eta');

        var shipParts = [];
        if (shipLineHtml) shipParts.push(shipLineHtml);
        if (shipNotesHtml) shipParts.push(shipNotesHtml);
        if (vesselStatusHtml) shipParts.push(vesselStatusHtml);
        var shipStr = shipParts.join(' ');

        var receiverSections = receiverNames.map(function(recv) {
          var samples = order.receivers[recv];
          var recvHdr = '';
          if (multiRecv) {
            recvHdr = '<div class="rh">→ ' + recv + '</div>';
          }

          var rows = samples.map(function(s) {
            var r = s.rowNumber;
            var recvCell = '';
            if (!multiRecv) {
              var recvVal = recv !== 'Unassigned' ? recv : '';
              recvCell = '<td>' + edt(r, 'Receiver', recvVal, 'receiver') + '</td>';
            }
            return '<tr>' +
              '<td>' + edt(r, 'Reference', s.reference, 'ref') + '</td>' +
              '<td class="dc">' + edt(r, 'Description', s.description, 'desc') + '</td>' +
              '<td>' + edt(r, 'Mark #', s.mark, 'mark') + '</td>' +
              '<td>' + edt(r, 'Container #', s.container, 'container') + '</td>' +
              '<td>' + edt(r, 'Cargo #', s.cargo, 'cargo') + '</td>' +
              recvCell +
              '<td>' + edt(r, 'Bag Count', s.bagCount, 'bags') + '</td>' +
              '<td>' + edt(r, 'Sample Weight', s.sampleWeight, 'wt') + '</td>' +
              '</tr>';
          }).join('');

          var recvTh = !multiRecv ? '<th>Receiver</th>' : '';
          return recvHdr +
            '<table><thead><tr>' +
            '<th>Ref</th><th class="dh">Description</th><th>Mark</th><th>Container</th><th>Cargo</th>' + recvTh + '<th>Bags</th><th>Wt</th>' +
            '</tr></thead><tbody>' + rows + '</tbody></table>';
        }).join('');

        return '<div class="ob">' +
          '<div class="ol">' +
            orderNumHtml +
            (shipStr ? ' <span class="si2">' + shipStr + '</span>' : '') +
            '<span class="ei">ETA: ' + etaHtml + '</span>' +
          '</div>' +
          receiverSections +
          '</div>';
      }).join('');

      return '<div class="cb">' +
        '<div class="ch">' + cust + '</div>' +
        ordersHtml +
        '</div>';
    }).join('');

    return '<div class="ws" id="wh-' + whId + '">' +
      '<div class="wh">' + wh + '</div>' +
      customerHtml +
      '</div>';
  }).join('');

  var today = new Date();
  var dateStr = today.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

  var html = '<!DOCTYPE html><html><head><style>\n' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }\n' +
    'body { font-family: Arial, sans-serif; display: flex; height: 100vh; background: #f5f5f5; }\n' +
    '.sidebar { width: 230px; min-width: 230px; background: #1a3a25; color: #fff; overflow-y: auto; padding: 12px 0; display: flex; flex-direction: column; }\n' +
    '.shdr { padding: 0 12px 12px; border-bottom: 1px solid rgba(255,255,255,0.1); margin-bottom: 8px; }\n' +
    '.stitle { font-size: 14px; font-weight: bold; color: #8cc99e; }\n' +
    '.ssub { font-size: 10px; color: #6b9e7d; margin-top: 2px; }\n' +
    '.slbl { font-size: 9px; color: #6b9e7d; text-transform: uppercase; padding: 6px 12px 3px; font-weight: bold; }\n' +
    '.si { padding: 8px 12px; cursor: pointer; border-left: 3px solid transparent; }\n' +
    '.si:hover { background: rgba(255,255,255,0.08); border-left-color: #4caf50; }\n' +
    '.sn { font-size: 12px; font-weight: bold; }\n' +
    '.ss { font-size: 9px; color: #8cc99e; margin-top: 1px; }\n' +
    '.sa { margin-top: 4px; display: flex; gap: 4px; }\n' +
    '.sb { padding: 3px 8px; border: none; border-radius: 3px; cursor: pointer; font-size: 11px; background: rgba(255,255,255,0.12); color: #fff; }\n' +
    '.sb:hover { background: rgba(255,255,255,0.25); }\n' +
    '.sb.dim { opacity: 0.35; }\n' +
    '.sfoot { margin-top: auto; padding: 12px; border-top: 1px solid rgba(255,255,255,0.1); }\n' +
    '.bpa { width: 100%; padding: 8px; background: #4A7C59; color: #fff; border: none; border-radius: 5px; font-size: 12px; font-weight: bold; cursor: pointer; }\n' +
    '.bcl { width: 100%; padding: 6px; background: rgba(255,255,255,0.1); color: #ccc; border: none; border-radius: 5px; font-size: 11px; cursor: pointer; margin-top: 5px; }\n' +
    '.main { flex: 1; overflow-y: auto; padding: 15px; }\n' +
    '.rpth { text-align: center; margin-bottom: 15px; padding: 15px; background: #fff; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }\n' +
    '.rptt { font-size: 18px; font-weight: bold; color: #2E5339; letter-spacing: 1px; }\n' +
    '.rpttag { font-size: 9px; color: #888; font-style: italic; }\n' +
    '.rpts { font-size: 14px; color: #2E5339; margin-top: 6px; }\n' +
    '.rptm { font-size: 11px; color: #666; margin-top: 3px; }\n' +
    '.ws { background: #fff; border-radius: 6px; margin-bottom: 15px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }\n' +
    '.wh { background: #2E5339; color: #fff; padding: 12px 16px; font-size: 16px; font-weight: bold; letter-spacing: 0.5px; }\n' +
    '.cb { border-bottom: 2px solid #ddd; }\n' +
    '.cb:last-child { border-bottom: none; }\n' +
    '.ch { padding: 6px 16px; font-weight: bold; font-size: 13px; color: #222; background: #f0f0f0; border: 1px solid #ccc; display: inline-block; margin: 10px 0 0 16px; border-radius: 3px; }\n' +
    '.ob { padding: 4px 16px 8px; border-bottom: 1px solid #eee; }\n' +
    '.ob:last-child { border-bottom: none; }\n' +
    '.ol { display: flex; justify-content: space-between; align-items: baseline; margin: 1px 0 5px; font-size: 11px; }\n' +
    '.on { font-weight: bold; color: #333; margin-right: 6px; }\n' +
    '.olink { color: #1a73e8; text-decoration: none; cursor: pointer; }\n' +
    '.olink:hover { text-decoration: underline; }\n' +
    '.ec { cursor: pointer; border-bottom: 1px dashed #2E5339; padding: 0 2px; }\n' +
    '.ec:hover { background: #e8f5e9; border-radius: 2px; }\n' +
    '.ep { color: #aaa; font-size: 9px; font-style: italic; }\n' +
    '.eip { border: 1px solid #2E5339; border-radius: 3px; padding: 2px 4px; font-size: 10px; font-family: Arial, sans-serif; outline: none; min-width: 60px; }\n' +
    '.eip:focus { box-shadow: 0 0 0 2px rgba(46,83,57,0.3); }\n' +
    '.si2 { color: #555; flex: 1; }\n' +
    '.ei { color: #c0392b; font-weight: bold; white-space: nowrap; }\n' +
    '.rh { padding: 4px 0 2px; font-size: 11px; font-weight: bold; color: #2E5339; border-top: 1px dashed #bbb; margin-top: 4px; }\n' +
    'table { width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 3px; }\n' +
    'th { background: #e8f5e9; color: #2E5339; padding: 3px 4px; text-align: left; font-size: 9px; border-bottom: 1px solid #c8e6c9; white-space: nowrap; }\n' +
    'td { padding: 3px 4px; border-bottom: 1px solid #f0f0f0; vertical-align: top; font-size: 10px; }\n' +
    '.dc { max-width: 220px; word-wrap: break-word; }\n' +
    '.dh { min-width: 120px; }\n' +
    '#toast { position: fixed; bottom: 15px; right: 15px; padding: 10px 18px; border-radius: 6px; font-weight: bold; font-size: 12px; z-index: 1000; display: none; box-shadow: 0 3px 10px rgba(0,0,0,0.2); }\n' +
    '.tok { background: #d4edda; color: #155724; display: block !important; }\n' +
    '.terr { background: #f8d7da; color: #721c24; display: block !important; }\n' +
    '.twait { background: #d1ecf1; color: #0c5460; display: block !important; }\n' +
    '@media print { .sidebar { display: none !important; } body { display: block; } .main { padding: 0; overflow: visible; } .ws { break-inside: avoid; box-shadow: none; border: 1px solid #999; margin-bottom: 10px; } .rpth { box-shadow: none; border: 1px solid #999; } #toast { display: none !important; } .ec { border-bottom: none; cursor: default; } .ep { display: none; } }\n' +
    'body.ps .ws { display: none; } body.ps .ws.pt { display: block; }\n' +
    '</style></head><body>\n' +
    '<div class="sidebar">\n' +
    '<div class="shdr"><div class="stitle">🏭 Field Report</div><div class="ssub">' + dateStr + '</div><div class="ssub">' + grandTotal + ' active samples</div>' + (perms.canEdit ? '<div class="ssub" style="color:#4caf50;margin-top:4px;">✏️ Edit mode' + (perms.columns === 'ALL' ? ' (full)' : '') + '</div>' : '') + '</div>\n' +
    '<div class="slbl">WAREHOUSES</div>\n' +
    sidebarItems +
    '<div class="sfoot">\n' +
    '<button class="bpa" onclick="window.print()">🖨️ Print Full Report</button>\n' +
    '<button class="bcl" style="background:#e67e22;color:#fff;" onclick="downloadOffline()">📥 Save Offline</button>\n' +
    '<button class="bcl" onclick="google.script.host.close()">Close</button>\n' +
    '</div></div>\n' +
    '<div class="main">\n' +
    '<div class="rpth">\n' +
    '<div class="rptt">COMMODITY SAMPLER SERVICES</div>\n' +
    '<div class="rpttag">"Integrity Through Independence"</div>\n' +
    '<div class="rpts">🏭 Field Report</div>\n' +
    '<div class="rptm">' + dateStr + ' &middot; ' + warehouseNames.length + ' warehouses &middot; ' + grandTotal + ' active samples</div>\n' +
    '</div>\n' +
    reportSections +
    '</div>\n' +
    '<div id="toast"></div>\n' +
    '<script>\n' +
    'function scrollTo(n){var e=document.getElementById("wh-"+n.replace(/[^a-zA-Z0-9]/g,"_"));if(e)e.scrollIntoView({behavior:"smooth",block:"start"});}\n' +
    'function printWh(n){var id="wh-"+n.replace(/[^a-zA-Z0-9]/g,"_"),e=document.getElementById(id);if(!e)return;document.body.classList.add("ps");e.classList.add("pt");window.print();setTimeout(function(){document.body.classList.remove("ps");e.classList.remove("pt");},500);}\n' +
    'function emailWh(n,has){if(!has){toast("terr","No emails for "+n+". Setup \\u2192 Setup Warehouse Emails.");return;}if(!confirm("Send Field Report for "+n+"?"))return;toast("twait","Sending...");google.script.run.withSuccessHandler(function(r){toast(r.success?"tok":"terr",r.message);}).withFailureHandler(function(e){toast("terr","Error: "+e.message);}).sendWarehouseFieldReport(n);}\n' +
    'function toast(c,m){var e=document.getElementById("toast");e.className=c;e.textContent=m;if(c!=="twait")setTimeout(function(){e.className="";e.style.display="none";},5000);}\n' +
    'function downloadOffline(){toast("twait","Generating offline report...");google.script.run.withSuccessHandler(function(r){if(!r.success){toast("terr",r.message);return;}var a=document.createElement("a");a.href="data:text/html;charset=utf-8,"+encodeURIComponent(r.html);a.download="CSS_Field_Report_Offline_"+new Date().toISOString().slice(0,10)+".html";document.body.appendChild(a);a.click();a.remove();toast("tok","\\u2705 Offline report downloaded!");}).withFailureHandler(function(e){toast("terr","Error: "+e.message);}).generateOfflineFieldReport();}\n' +
    'function editCell(el,row,col){' +
      'if(el.querySelector("input"))return;' +
      'var cur=el.querySelector("em")?"":(el.textContent||"").trim();' +
      'var inp=document.createElement("input");' +
      'inp.type="text";inp.className="eip";inp.value=cur;' +
      'inp.style.width=Math.max(60,el.offsetWidth)+"px";' +
      'el.textContent="";el.appendChild(inp);inp.focus();inp.select();' +
      'function save(){' +
        'var nv=inp.value.trim();' +
        'el.textContent=nv||"";' +
        'if(!nv)el.innerHTML=\'<em class="ep">+ \'+col.toLowerCase().replace(/[^a-z]/g," ").trim()+\'</em>\';' +
        'if(nv!==cur){' +
          'toast("twait","Saving...");' +
          'google.script.run.withSuccessHandler(function(r){toast(r.success?"tok":"terr",r.message);})' +
          '.withFailureHandler(function(e){toast("terr","Error: "+e.message);el.textContent=cur;})' +
          '.updateFieldReportCell(row,col,nv);' +
        '}' +
      '}' +
      'inp.addEventListener("blur",save);' +
      'inp.addEventListener("keydown",function(e){if(e.key==="Enter"){e.preventDefault();inp.blur();}if(e.key==="Escape"){inp.value=cur;inp.blur();}});' +
    '}\n' +
    '</script></body></html>';

  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🏭 Field Report — ' + dateStr);
}

// ============================================================
// FIELD REPORT — EMAIL TO WAREHOUSE
// ============================================================

function sendWarehouseFieldReport(warehouseName) {
  var reportData = getFieldReportData();
  var warehouseData = reportData.warehouses[warehouseName];

  if (!warehouseData) return { success: false, message: 'No active orders for ' + warehouseName };

  var emailConfig = getWarehouseEmailConfig();
  var emails = emailConfig[warehouseName] || '';
  if (!emails) return { success: false, message: 'No email addresses for ' + warehouseName + '. Setup → Setup Warehouse Emails.' };

  var recipients = emails.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e; });
  if (recipients.length === 0) return { success: false, message: 'No valid emails for ' + warehouseName };

  var today = new Date().toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  var customerNames = Object.keys(warehouseData).sort();
  var totalSamples = 0;

  var bodyHtml = '';
  customerNames.forEach(function(cust) {
    var orders = warehouseData[cust];
    var orderKeys = Object.keys(orders).sort(function(a, b) { return (orders[a].sortDate || 0) - (orders[b].sortDate || 0); });

    bodyHtml += '<tr style="background:#2E5339;color:#fff;"><td colspan="8" style="padding:10px;font-weight:bold;font-size:13px;">' + cust + '</td></tr>';

    orderKeys.forEach(function(orderNum) {
      var order = orders[orderNum];
      var shipParts = [];
      if (order.shippingLine) shipParts.push(order.shippingLine);
      if (order.shippingNotes) shipParts.push(order.shippingNotes);
      if (order.vesselStatus) shipParts.push(order.vesselStatus);
      var shipStr = shipParts.join(' ');

      bodyHtml += '<tr style="background:#e8f5e9;"><td colspan="8" style="padding:6px 10px;font-size:11px;">' +
        '<strong>' + orderNum + '</strong> &nbsp; ' +
        '<span style="color:#555;">' + shipStr + '</span>' +
        '<span style="float:right;color:#c0392b;font-weight:bold;">ETA: ' + (order.eta || '') + '</span>' +
        '</td></tr>';

      var receiverNames = Object.keys(order.receivers);
      var multiRecv = receiverNames.length > 1;

      receiverNames.forEach(function(recv) {
        var samples = order.receivers[recv];
        totalSamples += samples.length;

        if (multiRecv) {
          bodyHtml += '<tr><td colspan="8" style="padding:4px 10px;font-size:11px;font-weight:bold;color:#2E5339;border-top:1px dashed #ccc;">→ ' + recv + '</td></tr>';
        }

        samples.forEach(function(s, idx) {
          var bg = idx % 2 === 0 ? '#fff' : '#f9f9f9';
          var recvCell = !multiRecv ? '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;">' + recv + '</td>' : '';
          bodyHtml += '<tr style="background:' + bg + ';">' +
            '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;">' + (s.reference || '') + '</td>' +
            '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;max-width:200px;">' + (s.description || '') + '</td>' +
            '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;">' + (s.mark || '') + '</td>' +
            '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;">' + (s.container || '') + '</td>' +
            '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;">' + (s.cargo || '') + '</td>' +
            recvCell +
            '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;">' + (s.bagCount || '') + '</td>' +
            '<td style="padding:4px;border-bottom:1px solid #eee;font-size:10px;">' + (s.sampleWeight || '') + '</td>' +
            '</tr>';
        });
      });
    });
  });

  var emailHtml = '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;margin:0;padding:20px;background:#f5f5f5;">' +
    '<div style="max-width:950px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">' +
    '<div style="background:#2E5339;color:#fff;padding:20px;text-align:center;">' +
    '<div style="font-size:18px;font-weight:bold;letter-spacing:2px;">COMMODITY SAMPLER SERVICES</div>' +
    '<div style="font-size:10px;opacity:0.7;font-style:italic;">"Integrity Through Independence"</div>' +
    '</div>' +
    '<div style="padding:15px;">' +
    '<div style="background:#f0f7f2;padding:12px;border-radius:6px;margin-bottom:12px;text-align:center;">' +
    '<div style="font-size:16px;font-weight:bold;color:#2E5339;">🏭 ' + warehouseName + ' — Field Report</div>' +
    '<div style="color:#666;margin-top:4px;font-size:12px;">' + today + ' | ' + customerNames.length + ' client(s) | ' + totalSamples + ' sample(s)</div>' +
    '</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:10px;">' +
    '<thead><tr style="background:#4A7C59;color:#fff;">' +
    '<th style="padding:5px;text-align:left;">Ref</th>' +
    '<th style="padding:5px;text-align:left;">Description</th>' +
    '<th style="padding:5px;text-align:left;">Mark</th>' +
    '<th style="padding:5px;text-align:left;">Container</th>' +
    '<th style="padding:5px;text-align:left;">Cargo</th>' +
    '<th style="padding:5px;text-align:left;">Receiver</th>' +
    '<th style="padding:5px;text-align:left;">Bags</th>' +
    '<th style="padding:5px;text-align:left;">Wt</th>' +
    '</tr></thead><tbody>' + bodyHtml + '</tbody></table>' +
    '</div>' +
    '<div style="text-align:center;padding:12px;color:#999;font-size:9px;">Generated by CSS System | ' + today + '</div>' +
    '</div></body></html>';

  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: 'CSS Field Report — ' + warehouseName + ' — ' + new Date().toLocaleDateString(),
      htmlBody: emailHtml
    });
    return { success: true, message: '✅ Report sent to: ' + recipients.join(', ') };
  } catch (e) {
    return { success: false, message: '❌ Email error: ' + e.message };
  }
}
