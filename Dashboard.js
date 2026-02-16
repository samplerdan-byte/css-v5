// ============================================================
// 📊 DASHBOARD MODULE — CSS Operations Dashboard
// ============================================================
// Creates and refreshes a "Dashboard" sheet tab with live
// operational metrics pulled from All Orders, Completed Orders,
// and Invoices sheets.
//
// Add to existing CSS project as Dashboard.gs
// Menu item: ☕ CSS System → Reports → 📊 Operations Dashboard
// Can also be triggered on a time-based trigger for auto-refresh
// ============================================================

var DASH_CONFIG = {
  sheetName: 'Dashboard',
  headerBg: '#2E5339',
  headerFont: '#FFFFFF',
  cardBg: '#F0F7F2',
  alertBg: '#FFF3CD',
  alertCriticalBg: '#F8D7DA',
  sectionBg: '#E8F5E9',
  numberFont: '#2E5339',
  alertFont: '#856404',
  alertCriticalFont: '#721C24'
};


// ============================================================
// MAIN DASHBOARD BUILDER
// ============================================================

function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  var completedSheet = ss.getSheetByName(CONFIG.completedOrdersSheetName || 'Completed Orders');
  var invSheet = ss.getSheetByName(BILLING_CONFIG.invoicesSheetName);
  
  if (!mainSheet) {
    try { SpreadsheetApp.getUi().alert('Main sheet not found. Run Setup first.'); } catch(e) {}
    return;
  }

  // ── Gather all order data ──
  var allOrders = [];
  _gatherDashData(mainSheet, allOrders);
  _gatherDashData(completedSheet, allOrders);

  var now = new Date();
  var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var yesterday = new Date(todayStart.getTime() - 86400000);
  var weekAgo = new Date(todayStart.getTime() - 7 * 86400000);
  var monthAgo = new Date(todayStart.getTime() - 30 * 86400000);

  // ── Compute metrics ──
  var metrics = {
    total: allOrders.length,
    receivedToday: 0,
    scannedToday: 0,
    shippedToday: 0,
    completedToday: 0,
    pendingNotScanned: 0,
    inProgress: 0,
    shipped: 0,
    completed: 0,
    over24h: 0,
    over48h: 0,
    over72h: 0,
    thisWeekReceived: 0,
    thisWeekShipped: 0,
    thisMonthReceived: 0,
    thisMonthShipped: 0,
    byWarehouse: {},
    byCustomer: {},
    byStatus: {},
    recentOrders: [],
    unbilledCount: 0,
    avgTurnaround: []
  };

  // Collect billed samples for unbilled count
  var billedSamples = {};
  try {
    var liSheet = ss.getSheetByName(BILLING_CONFIG.lineItemsSheetName);
    if (liSheet && liSheet.getLastRow() >= 2) {
      var billedData = liSheet.getRange(2, 5, liSheet.getLastRow() - 1, 1).getValues();
      for (var b = 0; b < billedData.length; b++) {
        var bs = String(billedData[b][0]).trim();
        if (bs) billedSamples[bs] = true;
      }
    }
  } catch(e) {}

  for (var i = 0; i < allOrders.length; i++) {
    var o = allOrders[i];
    var status = (o.status || '').toLowerCase();
    var statusClean = o.status || 'Unknown';

    // Status counts
    if (!metrics.byStatus[statusClean]) metrics.byStatus[statusClean] = 0;
    metrics.byStatus[statusClean]++;

    // Warehouse counts
    var wh = o.warehouse || 'Unknown';
    if (!metrics.byWarehouse[wh]) metrics.byWarehouse[wh] = { total: 0, pending: 0, shipped: 0, completed: 0 };
    metrics.byWarehouse[wh].total++;

    // Customer counts
    var cust = o.sender || 'Unknown';
    if (!metrics.byCustomer[cust]) metrics.byCustomer[cust] = { total: 0, thisWeek: 0, thisMonth: 0 };
    metrics.byCustomer[cust].total++;

    // Date-based metrics
    var recvDate = o.receivedDate ? new Date(o.receivedDate) : null;
    var shipDate = o.shippedDate ? new Date(o.shippedDate) : null;

    if (recvDate && recvDate >= todayStart) metrics.receivedToday++;
    if (recvDate && recvDate >= weekAgo) {
      metrics.thisWeekReceived++;
      metrics.byCustomer[cust].thisWeek++;
    }
    if (recvDate && recvDate >= monthAgo) {
      metrics.thisMonthReceived++;
      metrics.byCustomer[cust].thisMonth++;
    }

    if (o.scannedDate && new Date(o.scannedDate) >= todayStart) metrics.scannedToday++;

    if (shipDate && shipDate >= todayStart) metrics.shippedToday++;
    if (shipDate && shipDate >= weekAgo) metrics.thisWeekShipped++;
    if (shipDate && shipDate >= monthAgo) metrics.thisMonthShipped++;

    if (status === 'shipped' || status === 'completed') {
      if (shipDate && shipDate >= todayStart) metrics.completedToday++;
      metrics.byWarehouse[wh].shipped++;
      if (status === 'completed') metrics.byWarehouse[wh].completed++;
    }

    // Pending (not yet scanned in / still "New" or "Pending")
    if (status === 'new' || status === 'pending' || status === 'received') {
      metrics.pendingNotScanned++;
      metrics.byWarehouse[wh].pending++;

      // Age alerts
      if (recvDate) {
        var ageHrs = (now - recvDate) / 3600000;
        if (ageHrs > 72) metrics.over72h++;
        else if (ageHrs > 48) metrics.over48h++;
        else if (ageHrs > 24) metrics.over24h++;
      }
    }

    if (status === 'in progress' || status === 'scanned' || status === 'scanning') {
      metrics.inProgress++;
    }

    if (status === 'shipped') metrics.shipped++;
    if (status === 'completed') metrics.completed++;

    // Unbilled shipped/completed
    if ((status === 'shipped' || status === 'completed') && o.csSample && !billedSamples[o.csSample]) {
      metrics.unbilledCount++;
    }

    // Turnaround time (received → shipped)
    if (recvDate && shipDate && shipDate > recvDate) {
      var turnHrs = (shipDate - recvDate) / 3600000;
      metrics.avgTurnaround.push(turnHrs);
    }

    // Recent orders (last 10 by received date)
    if (recvDate && recvDate >= weekAgo) {
      metrics.recentOrders.push(o);
    }
  }

  // Sort recent orders by date desc, take top 15
  metrics.recentOrders.sort(function(a, b) {
    return (new Date(b.receivedDate || 0)) - (new Date(a.receivedDate || 0));
  });
  metrics.recentOrders = metrics.recentOrders.slice(0, 15);

  // Avg turnaround
  var avgTurn = 0;
  if (metrics.avgTurnaround.length > 0) {
    var sum = 0;
    for (var t = 0; t < metrics.avgTurnaround.length; t++) sum += metrics.avgTurnaround[t];
    avgTurn = sum / metrics.avgTurnaround.length;
  }

  // Billing summary
  var billingStats = { draft: 0, sent: 0, paid: 0, overdue: 0, draftAmt: 0, sentAmt: 0, paidAmt: 0, overdueAmt: 0 };
  if (invSheet && invSheet.getLastRow() >= 2) {
    var invCol = _getColumnMap(invSheet);
    var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).getValues();
    for (var j = 0; j < invData.length; j++) {
      var invStatus = String(invData[j][invCol['Status']]).trim();
      var invTotal = parseFloat(invData[j][invCol['Total']]) || 0;
      var invDue = invData[j][invCol['Due Date']];
      if (invStatus === 'Sent' && invDue) {
        var dueD = new Date(invDue);
        dueD.setHours(0,0,0,0);
        if (dueD < todayStart) invStatus = 'Overdue';
      }
      if (invStatus === 'Draft') { billingStats.draft++; billingStats.draftAmt += invTotal; }
      else if (invStatus === 'Sent') { billingStats.sent++; billingStats.sentAmt += invTotal; }
      else if (invStatus === 'Paid') { billingStats.paid++; billingStats.paidAmt += invTotal; }
      else if (invStatus === 'Overdue') { billingStats.overdue++; billingStats.overdueAmt += invTotal; }
    }
  }

  // ── Build Dashboard Sheet ──
  var dashSheet = ss.getSheetByName(DASH_CONFIG.sheetName);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(DASH_CONFIG.sheetName);
  }
  dashSheet.clear();
  dashSheet.setColumnWidth(1, 30);   // spacer
  dashSheet.setColumnWidth(2, 200);
  dashSheet.setColumnWidth(3, 100);
  dashSheet.setColumnWidth(4, 30);   // spacer
  dashSheet.setColumnWidth(5, 200);
  dashSheet.setColumnWidth(6, 100);
  dashSheet.setColumnWidth(7, 30);   // spacer
  dashSheet.setColumnWidth(8, 200);
  dashSheet.setColumnWidth(9, 100);

  var row = 1;
  var fmtDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');

  // ── HEADER ──
  dashSheet.getRange(row, 1, 1, 9).merge()
    .setValue('📊 COFFEE SAMPLING DASHBOARD')
    .setFontSize(18).setFontWeight('bold').setFontColor(DASH_CONFIG.headerFont)
    .setBackground(DASH_CONFIG.headerBg).setHorizontalAlignment('center');
  row++;
  dashSheet.getRange(row, 1, 1, 9).merge()
    .setValue('Last Updated: ' + fmtDate)
    .setFontSize(9).setFontColor('#888').setHorizontalAlignment('center').setBackground('#f5f5f5');
  row += 2;

  // ── TODAY'S SUMMARY ──
  row = _dashSection(dashSheet, row, "TODAY'S SUMMARY");
  row = _dashCard(dashSheet, row, 2, 'Received Today', metrics.receivedToday);
  _dashCard(dashSheet, row - 1, 5, 'Scanned Today', metrics.scannedToday);
  _dashCard(dashSheet, row - 1, 8, 'Shipped Today', metrics.shippedToday);
  row++;
  row = _dashCard(dashSheet, row, 2, 'Pending (Not Scanned)', metrics.pendingNotScanned);
  _dashCard(dashSheet, row - 1, 5, 'In Progress', metrics.inProgress);
  _dashCard(dashSheet, row - 1, 8, 'Unbilled Samples', metrics.unbilledCount);
  row++;

  // ── ALERTS ──
  row = _dashSection(dashSheet, row, '⚠️ AGING ALERTS');
  var alertColor = metrics.over24h > 0 ? DASH_CONFIG.alertBg : '#fff';
  var alertFontC = metrics.over24h > 0 ? DASH_CONFIG.alertFont : '#333';
  row = _dashCard(dashSheet, row, 2, 'Over 24 Hours', metrics.over24h, alertColor, alertFontC);
  var alert48Color = metrics.over48h > 0 ? DASH_CONFIG.alertCriticalBg : '#fff';
  var alert48Font = metrics.over48h > 0 ? DASH_CONFIG.alertCriticalFont : '#333';
  _dashCard(dashSheet, row - 1, 5, 'Over 48 Hours', metrics.over48h, alert48Color, alert48Font);
  var alert72Color = metrics.over72h > 0 ? DASH_CONFIG.alertCriticalBg : '#fff';
  var alert72Font = metrics.over72h > 0 ? DASH_CONFIG.alertCriticalFont : '#333';
  _dashCard(dashSheet, row - 1, 8, 'Over 72 Hours', metrics.over72h, alert72Color, alert72Font);
  row++;

  // ── VOLUME TRENDS ──
  row = _dashSection(dashSheet, row, '📈 VOLUME');
  row = _dashCard(dashSheet, row, 2, 'This Week Received', metrics.thisWeekReceived);
  _dashCard(dashSheet, row - 1, 5, 'This Week Shipped', metrics.thisWeekShipped);
  var turnText = avgTurn > 0 ? (avgTurn < 24 ? Math.round(avgTurn) + ' hrs' : (avgTurn / 24).toFixed(1) + ' days') : 'N/A';
  _dashCard(dashSheet, row - 1, 8, 'Avg Turnaround', turnText);
  row++;
  row = _dashCard(dashSheet, row, 2, 'This Month Received', metrics.thisMonthReceived);
  _dashCard(dashSheet, row - 1, 5, 'This Month Shipped', metrics.thisMonthShipped);
  _dashCard(dashSheet, row - 1, 8, 'Total All Time', metrics.total);
  row++;

  // ── BY WAREHOUSE ──
  row = _dashSection(dashSheet, row, '🏭 BY WAREHOUSE');
  dashSheet.getRange(row, 2).setValue('Warehouse').setFontWeight('bold').setFontSize(9);
  dashSheet.getRange(row, 3).setValue('Pending').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  dashSheet.getRange(row, 5).setValue('Shipped').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  dashSheet.getRange(row, 6).setValue('Completed').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  dashSheet.getRange(row, 8).setValue('Total').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  dashSheet.getRange(row, 2, 1, 8).setBackground('#E8F5E9');
  row++;

  var whKeys = Object.keys(metrics.byWarehouse).sort();
  for (var w = 0; w < whKeys.length; w++) {
    var wd = metrics.byWarehouse[whKeys[w]];
    dashSheet.getRange(row, 2).setValue(whKeys[w]).setFontSize(10);
    dashSheet.getRange(row, 3).setValue(wd.pending).setHorizontalAlignment('right').setFontSize(10);
    dashSheet.getRange(row, 5).setValue(wd.shipped).setHorizontalAlignment('right').setFontSize(10);
    dashSheet.getRange(row, 6).setValue(wd.completed).setHorizontalAlignment('right').setFontSize(10);
    dashSheet.getRange(row, 8).setValue(wd.total).setHorizontalAlignment('right').setFontSize(10);
    if (wd.pending > 0) dashSheet.getRange(row, 3).setFontColor('#dc3545').setFontWeight('bold');
    if (w % 2 === 0) dashSheet.getRange(row, 2, 1, 8).setBackground('#fafafa');
    row++;
  }
  row++;

  // ── TOP CUSTOMERS (this month) ──
  row = _dashSection(dashSheet, row, '👤 TOP CUSTOMERS (30 days)');
  dashSheet.getRange(row, 2).setValue('Customer').setFontWeight('bold').setFontSize(9);
  dashSheet.getRange(row, 3).setValue('This Week').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  dashSheet.getRange(row, 5).setValue('This Month').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  dashSheet.getRange(row, 6).setValue('All Time').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('right');
  dashSheet.getRange(row, 2, 1, 8).setBackground('#E8F5E9');
  row++;

  var custKeys = Object.keys(metrics.byCustomer).sort(function(a, b) {
    return metrics.byCustomer[b].thisMonth - metrics.byCustomer[a].thisMonth;
  }).slice(0, 15);

  for (var c = 0; c < custKeys.length; c++) {
    var cd = metrics.byCustomer[custKeys[c]];
    if (cd.thisMonth === 0 && cd.thisWeek === 0) continue;
    dashSheet.getRange(row, 2).setValue(custKeys[c]).setFontSize(10);
    dashSheet.getRange(row, 3).setValue(cd.thisWeek).setHorizontalAlignment('right').setFontSize(10);
    dashSheet.getRange(row, 5).setValue(cd.thisMonth).setHorizontalAlignment('right').setFontSize(10);
    dashSheet.getRange(row, 6).setValue(cd.total).setHorizontalAlignment('right').setFontSize(10);
    if (c % 2 === 0) dashSheet.getRange(row, 2, 1, 8).setBackground('#fafafa');
    row++;
  }
  row++;

  // ── BILLING SNAPSHOT ──
  row = _dashSection(dashSheet, row, '💰 BILLING SNAPSHOT');
  row = _dashCard(dashSheet, row, 2, 'Draft Invoices', billingStats.draft + ' ($' + billingStats.draftAmt.toFixed(0) + ')');
  _dashCard(dashSheet, row - 1, 5, 'Sent / Awaiting', billingStats.sent + ' ($' + billingStats.sentAmt.toFixed(0) + ')');
  var odColor = billingStats.overdue > 0 ? DASH_CONFIG.alertCriticalBg : '#fff';
  var odFont = billingStats.overdue > 0 ? DASH_CONFIG.alertCriticalFont : '#333';
  _dashCard(dashSheet, row - 1, 8, 'Overdue', billingStats.overdue + ' ($' + billingStats.overdueAmt.toFixed(0) + ')', odColor, odFont);
  row++;
  row = _dashCard(dashSheet, row, 2, 'Paid (all time)', billingStats.paid + ' ($' + billingStats.paidAmt.toFixed(0) + ')');
  _dashCard(dashSheet, row - 1, 5, 'Unbilled Samples', metrics.unbilledCount);
  row++;

  // ── STATUS BREAKDOWN ──
  row = _dashSection(dashSheet, row, '📋 STATUS BREAKDOWN');
  var statKeys = Object.keys(metrics.byStatus).sort();
  for (var s = 0; s < statKeys.length; s++) {
    dashSheet.getRange(row, 2).setValue(statKeys[s]).setFontSize(10);
    dashSheet.getRange(row, 3).setValue(metrics.byStatus[statKeys[s]]).setHorizontalAlignment('right').setFontSize(10).setFontWeight('bold');
    if (s % 2 === 0) dashSheet.getRange(row, 2, 1, 2).setBackground('#fafafa');
    row++;
  }
  row++;

  // ── RECENT ORDERS ──
  if (metrics.recentOrders.length > 0) {
    row = _dashSection(dashSheet, row, '🕐 RECENT ORDERS (last 7 days)');
    dashSheet.getRange(row, 2).setValue('CS Sample #').setFontWeight('bold').setFontSize(9);
    dashSheet.getRange(row, 3).setValue('Status').setFontWeight('bold').setFontSize(9);
    dashSheet.getRange(row, 5).setValue('Customer').setFontWeight('bold').setFontSize(9);
    dashSheet.getRange(row, 6).setValue('Warehouse').setFontWeight('bold').setFontSize(9);
    dashSheet.getRange(row, 8).setValue('Received').setFontWeight('bold').setFontSize(9);
    dashSheet.getRange(row, 2, 1, 8).setBackground('#E8F5E9');
    row++;

    for (var r = 0; r < metrics.recentOrders.length; r++) {
      var ro = metrics.recentOrders[r];
      dashSheet.getRange(row, 2).setValue(ro.csSample || ro.csOrder || '—').setFontSize(10);
      dashSheet.getRange(row, 3).setValue(ro.status || '—').setFontSize(10);
      dashSheet.getRange(row, 5).setValue(ro.sender || '—').setFontSize(10);
      dashSheet.getRange(row, 6).setValue(ro.warehouse || '—').setFontSize(10);
      var recvStr = '';
      try { recvStr = Utilities.formatDate(new Date(ro.receivedDate), Session.getScriptTimeZone(), 'MM/dd hh:mm a'); } catch(e) {}
      dashSheet.getRange(row, 8).setValue(recvStr).setFontSize(10);
      if (r % 2 === 0) dashSheet.getRange(row, 2, 1, 8).setBackground('#fafafa');
      row++;
    }
  }

  // Protect the dashboard
  try {
    var protection = dashSheet.protect().setDescription('Dashboard — auto-generated');
    protection.setWarningOnly(true);
  } catch(e) {}

  // Move dashboard to first position
  try { ss.setActiveSheet(dashSheet); ss.moveActiveSheet(1); } catch(e) {}

  Logger.log('Dashboard refreshed at ' + fmtDate);
}


// ============================================================
// HELPER: Gather order data from a sheet
// ============================================================

function _gatherDashData(sheet, results) {
  if (!sheet || sheet.getLastRow() < 2) return;
  var col = _getColumnMap(sheet);
  if (col['Status'] === undefined) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var csSample = String(data[i][col['CS Sample #']] || '').trim();
    if (!csSample && !String(data[i][col['CS Order #']] || '').trim()) continue; // skip empty rows

    results.push({
      csSample: csSample,
      csOrder: String(data[i][col['CS Order #']] || '').trim(),
      status: String(data[i][col['Status']] || '').trim(),
      sender: String(data[i][col['Sender']] || '').trim(),
      receiver: String(data[i][col['Receiver']] || '').trim(),
      warehouse: String(data[i][col['Warehouse']] || '').trim(),
      container: String(data[i][col['Container #']] || '').trim(),
      receivedDate: col['Received Date'] !== undefined ? data[i][col['Received Date']] : null,
      scannedDate: col['Scanned Date'] !== undefined ? data[i][col['Scanned Date']] : (col['Scan In Date'] !== undefined ? data[i][col['Scan In Date']] : null),
      shippedDate: col['Shipped Date'] !== undefined ? data[i][col['Shipped Date']] : (col['Ship Date'] !== undefined ? data[i][col['Ship Date']] : null)
    });
  }
}


// ============================================================
// HELPER: Section header
// ============================================================

function _dashSection(sheet, row, title) {
  sheet.getRange(row, 2, 1, 8).merge()
    .setValue(title)
    .setFontSize(12).setFontWeight('bold').setFontColor(DASH_CONFIG.headerBg)
    .setBackground(DASH_CONFIG.sectionBg);
  return row + 1;
}


// ============================================================
// HELPER: Metric card (label + value in two cells)
// ============================================================

function _dashCard(sheet, row, col, label, value, bgColor, fontColor) {
  bgColor = bgColor || DASH_CONFIG.cardBg;
  fontColor = fontColor || DASH_CONFIG.numberFont;
  
  sheet.getRange(row, col).setValue(label)
    .setFontSize(9).setFontColor('#666').setBackground(bgColor)
    .setVerticalAlignment('middle');
  sheet.getRange(row, col + 1).setValue(value)
    .setFontSize(16).setFontWeight('bold').setFontColor(fontColor)
    .setBackground(bgColor).setHorizontalAlignment('right')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 36);
  return row + 1;
}
