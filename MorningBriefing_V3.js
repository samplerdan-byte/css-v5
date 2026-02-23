// ============================================================
// MorningBriefing_v3.gs — Daily Status Email (All-in-One)
// ============================================================
// Merged: MorningBriefing.gs + DailyDigest.gs + HealthCheck.gs
// Replaces all 3 files. Delete originals after deploying.
//
// Covers in one compact email:
//   • Quick stats bar (yesterday in/out, pending, ready)
//   • Week/month volume + by-client breakdown
//   • Needs Review + recent errors (from HealthCheck)
//   • Received yesterday
//   • Pending unscanned (grouped by warehouse)
//   • Ready to ship (grouped by receiver)
//   • Shipped yesterday
//   • Container ETAs today / overdue
//   • Overdue invoices
//
// Menu items:
//   .addSubMenu(ui.createMenu('☀️ Briefing')
//     .addItem('📊 Quick Brief', 'showBriefingNow')
//     .addItem('📧 Send Briefing Email', 'sendMorningBriefing')
//     .addItem('⚙️ Setup Daily Trigger (7 AM)', 'setupMorningBriefing')
//     .addItem('🚫 Remove Daily Trigger', 'removeMorningBriefing'))
// ============================================================

var BRIEFING_CONFIG = {
  recipientEmail: Session.getActiveUser().getEmail(),
  sendHour: 7,
  timezone: 'America/New_York',
  subjectPrefix: '☕ CSS Daily Briefing',
  includeInvoices: true,
  skipWeekends: false    // Set true to skip Sat/Sun
};

// ============================================================
// TRIGGER SETUP
// ============================================================

function setupMorningBriefing() {
  _removeBriefingTriggers();
  ScriptApp.newTrigger('sendMorningBriefing')
    .timeBased()
    .atHour(BRIEFING_CONFIG.sendHour)
    .everyDays(1)
    .inTimezone(BRIEFING_CONFIG.timezone)
    .create();
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Morning Briefing Scheduled!\n\n' +
      'Daily at ' + BRIEFING_CONFIG.sendHour + ':00 AM Eastern.\n' +
      'To: ' + BRIEFING_CONFIG.recipientEmail + '\n\n' +
      'Run "Send Briefing Now" to test.'
    );
  } catch(e) {
    Logger.log('Briefing trigger created. To: ' + BRIEFING_CONFIG.recipientEmail);
  }
}

function removeMorningBriefing() {
  _removeBriefingTriggers();
  try { SpreadsheetApp.getUi().alert('✅ Daily briefing trigger removed.'); }
  catch(e) { Logger.log('Briefing trigger removed.'); }
}

function _removeBriefingTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendMorningBriefing' ||
        t.getHandlerFunction() === 'sendDailyDigest' ||
        t.getHandlerFunction() === 'sendDailyHealthCheck') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ============================================================
// MAIN — Gather all data and send briefing
// ============================================================

function sendMorningBriefing() {
  try {
    Logger.log('sendMorningBriefing() called');
    // Weekend skip
    if (BRIEFING_CONFIG.skipWeekends) {
      var dow = new Date().getDay();
      if (dow === 0 || dow === 6) { Logger.log('sendMorningBriefing: skipping weekend'); return; }
    }

    if (!BRIEFING_CONFIG.recipientEmail) {
      Logger.log('sendMorningBriefing: no recipient email configured');
      return;
    }

    Logger.log('sendMorningBriefing: gathering briefing data');
    var b = _gatherBriefingData();
    var html = _buildBriefingHtml(b);
    var subject = BRIEFING_CONFIG.subjectPrefix + ' — ' + b.date;
    var plain = 'CSS Briefing — ' + b.date +
      '\nYesterday: ' + b.receivedYesterday.length + ' in, ' + b.shippedYesterday.length + ' out' +
      '\nPending: ' + b.pendingUnscanned.length + ' unscanned, ' + b.scannedReady.length + ' ready';

    // QUOTA GUARD: Check email quota before sending
    var emailQuota = (typeof _quotaCheckEmailQuota === 'function') ? _quotaCheckEmailQuota() : { canSend: true, remaining: 100, countToday: 0, limit: 1500 };
    if (!emailQuota.canSend) {
      var quotaErr = 'sendMorningBriefing: Email quota exhausted (' + emailQuota.countToday + '/' + emailQuota.limit + ' today)';
      Logger.log(quotaErr);
      if (typeof logError === 'function') logError('sendMorningBriefing', 'Email quota exhausted', { count: emailQuota.countToday, limit: emailQuota.limit });
      throw new Error(quotaErr);
    }

    MailApp.sendEmail({ to: BRIEFING_CONFIG.recipientEmail, subject: subject, body: plain, htmlBody: html });
    // QUOTA GUARD: Log the send
    if (typeof _quotaLogEmailSend === 'function') _quotaLogEmailSend();
    Logger.log('sendMorningBriefing: briefing sent to ' + BRIEFING_CONFIG.recipientEmail + ' (' + emailQuota.remaining + ' sends remaining)');

    try { SpreadsheetApp.getUi().alert('✅ Briefing sent to ' + BRIEFING_CONFIG.recipientEmail); }
    catch(e) { logError('sendMorningBriefing', 'Error showing alert (non-critical)', { error: e.message }); }
  } catch(e) {
    Logger.log('sendMorningBriefing: error: ' + e.message);
    if (typeof logError === 'function') logError('sendMorningBriefing', 'Error sending briefing email', { recipient: BRIEFING_CONFIG.recipientEmail, error: e.message });
    try { SpreadsheetApp.getUi().alert('Briefing failed: ' + e); } catch(e2) {}
  }
}

// ============================================================
// QUICK BRIEF — Show live stats in sidebar (no email)
// ============================================================

function showBriefingNow() {
  try {
    Logger.log('showBriefingNow() called');
    var b = _gatherBriefingData();
    var sidebarHtml = _buildSidebarHtml(b);
    var htmlOutput = HtmlService.createHtmlOutput(sidebarHtml)
      .setTitle('☀️ Quick Brief')
      .setWidth(320);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch(e) {
    Logger.log('showBriefingNow: error: ' + e.message);
    if (typeof logError === 'function') logError('showBriefingNow', 'Error displaying briefing sidebar', { error: e.message });
  }
}

// ============================================================
// DATA GATHERING — Shared by email + sidebar
// ============================================================

function _gatherBriefingData() {
  try {
    Logger.log('_gatherBriefingData: gathering briefing data');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var today = _startOfDay(now);
    var yesterday = new Date(today.getTime() - 86400000);
    var weekAgo = new Date(today.getTime() - 7 * 86400000);
    var monthAgo = new Date(today.getTime() - 30 * 86400000);
  
  var b = {
    date: Utilities.formatDate(now, BRIEFING_CONFIG.timezone, 'EEEE, MMMM d, yyyy'),
    time: Utilities.formatDate(now, BRIEFING_CONFIG.timezone, 'h:mm a'),
    receivedYesterday: [],
    pendingUnscanned: [],
    scannedReady: [],
    shippedYesterday: [],
    containersToday: [],
    containersOverdue: [],
    overdueInvoices: [],
    needsReview: 0,
    recentErrors: [],
    byClient: {},
    statsWeek: { received: 0, shipped: 0 },
    statsMonth: { received: 0, shipped: 0 },
    statsTotal: { received: 0, scanned: 0, shipped: 0 }
  };
  
  // ── Main Sheet ──
  var mainSheetName = (typeof CONFIG !== 'undefined' && CONFIG.mainSheetName) ? CONFIG.mainSheetName : 'All Orders';
  Logger.log('_gatherBriefingData: loading ' + mainSheetName);
  var mainSheet = ss.getSheetByName(mainSheetName);
  if (mainSheet && mainSheet.getLastRow() >= 2) {
    var col = _getColumnMap(mainSheet);
    var data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    var reviewIdx = col['Needs Review'];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var status = String(row[col['Status']] || '').trim();
      var sender = String(row[col['Sender']] || '');
      var csSample = String(row[col['CS Sample #']] || '');
      var desc = String(row[col['Description']] || '');
      var container = String(row[col['Container #']] || '');
      var receiver = String(row[col['Receiver']] || '');
      var warehouse = String(row[col['Warehouse']] || '');
      var timestamp = row[col['Timestamp']];
      var shippedDate = col['Shipped Date'] !== undefined ? row[col['Shipped Date']] : null;
      var tracking = col['Tracking Number'] !== undefined ? String(row[col['Tracking Number']] || '') : '';
      var containerETA = col['Container ETA'] !== undefined ? String(row[col['Container ETA']] || '') : '';
      
      var o = { csSample: csSample, sender: sender, desc: desc, container: container,
                receiver: receiver, warehouse: warehouse, tracking: tracking, status: status,
                containerETA: containerETA };
      
      var tsDate = timestamp ? new Date(timestamp) : null;
      var shipDate = shippedDate ? new Date(shippedDate) : null;
      
      if (reviewIdx !== undefined && row[reviewIdx]) b.needsReview++;
      
      // Normalize tsDate to start of day for accurate comparisons
      var tsDateNorm = tsDate ? _startOfDay(tsDate) : null;
      if (tsDateNorm && tsDateNorm >= yesterday && tsDateNorm < today) {
        b.receivedYesterday.push(o);
        var clientKey = sender || 'Unknown';
        b.byClient[clientKey] = (b.byClient[clientKey] || 0) + 1;
      }

      // Only count 'Received' status (or truly empty) as pending, not all non-scanned statuses
      if ((status === '' || status === 'Received' || status === 'New') && tsDateNorm && tsDateNorm < today) b.pendingUnscanned.push(o);
      if (status === 'Scanned') b.scannedReady.push(o);
      // Normalize shipDate to start of day for accurate comparisons
      var shipDateNorm = shipDate ? _startOfDay(shipDate) : null;
      if ((status === 'Shipped' || status === 'Completed') && shipDateNorm && shipDateNorm >= yesterday && shipDateNorm < today) b.shippedYesterday.push(o);
      
      if (containerETA) {
        try {
          var etaDate = new Date(containerETA);
          if (!isNaN(etaDate.getTime())) {
            var etaDay = _startOfDay(etaDate);
            if (etaDay.getTime() === today.getTime()) b.containersToday.push(o);
            else if (etaDay < today && status !== 'Shipped' && status !== 'Completed') b.containersOverdue.push(o);
          }
        } catch(e) {}
      }
      
      if (tsDateNorm && tsDateNorm >= weekAgo) b.statsWeek.received++;
      if (tsDateNorm && tsDateNorm >= monthAgo) b.statsMonth.received++;
      if (shipDateNorm && shipDateNorm >= weekAgo) b.statsWeek.shipped++;
      if (shipDateNorm && shipDateNorm >= monthAgo) b.statsMonth.shipped++;

      if (status === '' || status === 'Received' || status === 'New') b.statsTotal.received++;
      else if (status === 'Scanned') b.statsTotal.scanned++;
      else if (status === 'Shipped' || status === 'Completed') b.statsTotal.shipped++;
    }
  }
  
  // ── Completed Orders ──
  var compSheetName = (typeof CONFIG !== 'undefined' && CONFIG.completedOrdersSheetName) ? CONFIG.completedOrdersSheetName : 'Completed Orders';
  var compSheet = ss.getSheetByName(compSheetName);
  if (compSheet && compSheet.getLastRow() >= 2) {
    var compCol = _getColumnMap(compSheet);
    var compData = compSheet.getRange(2, 1, compSheet.getLastRow() - 1, compSheet.getLastColumn()).getValues();
    
    for (var j = 0; j < compData.length; j++) {
      var crow = compData[j];
      var cShipDate = compCol['Shipped Date'] !== undefined ? crow[compCol['Shipped Date']] : null;
      var cTimestamp = crow[compCol['Timestamp']];
      
      if (cShipDate) {
        var sd = new Date(cShipDate);
        if (sd >= yesterday && sd < today) {
          b.shippedYesterday.push({
            csSample: String(crow[compCol['CS Sample #']] || ''),
            sender: String(crow[compCol['Sender']] || ''),
            desc: String(crow[compCol['Description']] || ''),
            receiver: String(crow[compCol['Receiver']] || ''),
            tracking: compCol['Tracking Number'] !== undefined ? String(crow[compCol['Tracking Number']] || '') : ''
          });
        }
        if (sd >= weekAgo) b.statsWeek.shipped++;
        if (sd >= monthAgo) b.statsMonth.shipped++;
      }
      if (cTimestamp) {
        var ct = new Date(cTimestamp);
        if (ct >= weekAgo) b.statsWeek.received++;
        if (ct >= monthAgo) b.statsMonth.received++;
      }
      b.statsTotal.shipped++;
    }
  }
  
  // ── Recent Errors ──
  if (typeof getRecentErrors === 'function') {
    try { b.recentErrors = getRecentErrors(24) || []; } catch(e) {}
  }
  
  // ── Overdue Invoices ──
  if (BRIEFING_CONFIG.includeInvoices) {
    var invSheet = ss.getSheetByName('Invoices');
    if (invSheet && invSheet.getLastRow() >= 2) {
      var invCol = _getColumnMap(invSheet);
      var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).getValues();
      
      var invStatusCol = invCol['Status'];
      var invDueDateCol = invCol['Due Date'];
      var invNumCol = invCol['Invoice #'];
      var invCustCol = invCol['Customer'];
      var invTotalCol = invCol['Total'];

      if (invStatusCol === undefined || invDueDateCol === undefined) {
        Logger.log('_gatherBriefingData: Invoices sheet missing expected columns');
      } else {
        for (var k = 0; k < invData.length; k++) {
          var inv = invData[k];
          var invStatus = String(inv[invStatusCol] || '').trim();
          var dueDate = inv[invDueDateCol];

          if (invStatus === 'Sent' && dueDate) {
            var dd = new Date(dueDate);
            var ddNorm = _startOfDay(dd);
            if (ddNorm < today) {
              var daysOverdue = Math.floor((today.getTime() - ddNorm.getTime()) / 86400000);
              b.overdueInvoices.push({
                invoiceNum: invNumCol !== undefined ? String(inv[invNumCol] || '') : '',
                customer: invCustCol !== undefined ? String(inv[invCustCol] || '') : '',
                total: invTotalCol !== undefined ? (parseFloat(inv[invTotalCol]) || 0) : 0,
                dueDate: Utilities.formatDate(dd, BRIEFING_CONFIG.timezone, 'MM/dd'),
                daysOverdue: Math.max(1, daysOverdue)
              });
            }
          }
        }
      }
    }
  }

    return b;
  } catch (e) {
    Logger.log('_gatherBriefingData: error: ' + e.message);
    if (typeof logError === 'function') logError('_gatherBriefingData', 'Error gathering briefing data', { error: e.message });
    return {
      date: Utilities.formatDate(new Date(), BRIEFING_CONFIG.timezone, 'EEEE, MMMM d, yyyy'),
      time: Utilities.formatDate(new Date(), BRIEFING_CONFIG.timezone, 'h:mm a'),
      receivedYesterday: [],
      pendingUnscanned: [],
      scannedReady: [],
      shippedYesterday: [],
      containersToday: [],
      containersOverdue: [],
      overdueInvoices: [],
      needsReview: 0,
      recentErrors: [],
      byClient: {},
      statsWeek: { received: 0, shipped: 0 },
      statsMonth: { received: 0, shipped: 0 },
      statsTotal: { received: 0, scanned: 0, shipped: 0 }
    };
  }
}

// ============================================================
// COMPACT HTML — designed to fit one screen
// ============================================================

function _buildBriefingHtml(b) {
  var G = '#2E5339', GL = '#e8f5e9', Y = '#fff3cd', R = '#f8d7da';
  var h = '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
    '<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:0;">' +
    '<div style="max-width:620px;margin:0 auto;background:#fff;">';
  
  // ── Header (compact) ──
  h += '<div style="background:' + G + ';padding:14px 20px;text-align:center;">' +
    '<span style="color:#fff;font-size:16px;font-weight:bold;letter-spacing:1px;">☀️ MORNING BRIEFING</span>' +
    '<span style="color:#8cc99e;font-size:11px;margin-left:10px;">' + b.date + '</span></div>';
  
  // ── Stats Bar (3 columns) ──
  h += '<table style="width:100%;background:#f9f9f9;border-bottom:1px solid #eee;font-size:12px;" cellpadding="0" cellspacing="0"><tr>' +
    _statTd('📥 Yesterday', b.receivedYesterday.length + ' in / ' + b.shippedYesterday.length + ' out') +
    _statTd('⏳ Pending', b.pendingUnscanned.length + ' unscanned') +
    _statTd('📦 Ready', b.scannedReady.length + ' to ship') +
    '</tr></table>';
  
  // ── Volume + Alerts row (side by side) ──
  h += '<table style="width:100%;font-size:11px;border-bottom:1px solid #eee;" cellpadding="0" cellspacing="0"><tr>';
  
  // Left: volume stats
  h += '<td style="padding:10px 16px;vertical-align:top;width:50%;border-right:1px solid #eee;">';
  h += '<table style="width:100%;font-size:11px;border-collapse:collapse;">';
  h += '<tr style="color:#888;"><td></td><td style="text-align:center;">In</td><td style="text-align:center;">Out</td></tr>';
  h += '<tr><td style="color:#888;">7d</td><td style="text-align:center;font-weight:bold;">' + b.statsWeek.received + '</td><td style="text-align:center;font-weight:bold;">' + b.statsWeek.shipped + '</td></tr>';
  h += '<tr><td style="color:#888;">30d</td><td style="text-align:center;font-weight:bold;">' + b.statsMonth.received + '</td><td style="text-align:center;font-weight:bold;">' + b.statsMonth.shipped + '</td></tr>';
  h += '<tr style="border-top:1px solid #eee;"><td style="color:#888;padding-top:4px;">Active</td>';
  h += '<td style="text-align:center;font-weight:bold;padding-top:4px;">' + b.statsTotal.received + ' pend</td>';
  h += '<td style="text-align:center;font-weight:bold;padding-top:4px;">' + b.statsTotal.scanned + ' scan</td></tr>';
  h += '</table></td>';
  
  // Right: alerts + by-client
  h += '<td style="padding:10px 16px;vertical-align:top;width:50%;">';
  
  // Alerts
  var alerts = [];
  if (b.needsReview > 0) alerts.push('⚠️ ' + b.needsReview + ' need review');
  if (b.containersOverdue.length > 0) alerts.push('🚨 ' + b.containersOverdue.length + ' containers overdue');
  if (b.overdueInvoices.length > 0) alerts.push('💰 ' + b.overdueInvoices.length + ' invoices overdue');
  if (b.recentErrors.length > 0) alerts.push('🔴 ' + b.recentErrors.length + ' errors (24h)');
  
  if (alerts.length > 0) {
    h += '<div style="font-size:11px;margin-bottom:6px;">' + alerts.join('<br>') + '</div>';
  } else {
    h += '<div style="font-size:11px;color:#888;margin-bottom:6px;">✅ No alerts</div>';
  }
  
  // By-client yesterday
  var clientKeys = Object.keys(b.byClient).sort(function(a,c) { return b.byClient[c] - b.byClient[a]; });
  if (clientKeys.length > 0) {
    h += '<div style="font-size:10px;color:#888;border-top:1px solid #eee;padding-top:4px;">Yesterday by client:</div>';
    var maxShow = Math.min(clientKeys.length, 4);
    for (var ci = 0; ci < maxShow; ci++) {
      var ck = clientKeys[ci];
      h += '<div style="font-size:10px;color:#555;">' + _escHtml(_truncate(ck, 22)) + ': ' + b.byClient[ck] + '</div>';
    }
    if (clientKeys.length > maxShow) h += '<div style="font-size:10px;color:#aaa;">+' + (clientKeys.length - maxShow) + ' more</div>';
  }
  h += '</td></tr></table>';
  
  // ── Sections (compact rows) ──
  
  // Received Yesterday
  if (b.receivedYesterday.length > 0) {
    h += _secHdr('📥 Received Yesterday', b.receivedYesterday.length);
    h += _compactTable(b.receivedYesterday, 'warehouse', GL);
  }
  
  // Pending Unscanned
  if (b.pendingUnscanned.length > 0) {
    h += _secHdr('⚠️ Pending Unscanned', b.pendingUnscanned.length, Y);
    // Group by warehouse
    var byWh = _groupBy(b.pendingUnscanned, 'warehouse');
    Object.keys(byWh).sort().forEach(function(wh) {
      h += '<div style="padding:2px 16px;font-size:10px;color:#666;font-weight:bold;background:#fafafa;">📍 ' + (wh || 'Unknown') + ' (' + byWh[wh].length + ')</div>';
      h += _compactTable(byWh[wh], '', '#fffde7');
    });
  }
  
  // Ready to Ship
  if (b.scannedReady.length > 0) {
    h += _secHdr('📦 Ready to Ship', b.scannedReady.length);
    var byRecv = _groupBy(b.scannedReady, 'receiver');
    Object.keys(byRecv).sort().forEach(function(recv) {
      h += '<div style="padding:2px 16px;font-size:10px;color:#666;font-weight:bold;background:#fafafa;">📬 ' + (recv || 'No receiver') + ' (' + byRecv[recv].length + ')</div>';
      h += _compactTable(byRecv[recv], '', '#e3f2fd');
    });
  }
  
  // Shipped Yesterday
  if (b.shippedYesterday.length > 0) {
    h += _secHdr('✅ Shipped Yesterday', b.shippedYesterday.length);
    b.shippedYesterday.forEach(function(o) {
      var extra = o.tracking ? ' 📦' + o.tracking : '';
      h += _row(o.csSample, o.sender, (o.desc || '') + extra, GL);
    });
  }
  
  // Container ETAs
  if (b.containersToday.length > 0) {
    h += _secHdr('🚢 Arriving Today', b.containersToday.length);
    b.containersToday.forEach(function(o) {
      h += _row(o.container, o.sender, o.warehouse, '#e8eaf6');
    });
  }
  if (b.containersOverdue.length > 0) {
    h += _secHdr('🚨 Containers Overdue', b.containersOverdue.length, R);
    b.containersOverdue.forEach(function(o) {
      h += _row(o.container, o.sender, 'ETA: ' + o.containerETA, '#ffebee');
    });
  }
  
  // Overdue Invoices
  if (b.overdueInvoices.length > 0) {
    h += _secHdr('💰 Overdue Invoices', b.overdueInvoices.length, R);
    b.overdueInvoices.forEach(function(inv) {
      h += '<div style="padding:4px 16px;border-bottom:1px solid #f0f0f0;font-size:11px;background:#fff8f8;">' +
        '<b>' + inv.invoiceNum + '</b> ' + _escHtml(inv.customer) +
        '<span style="float:right;color:#d32f2f;font-weight:bold;">$' + Number(inv.total).toFixed(2) + '</span>' +
        '<br><span style="color:#888;">Due ' + inv.dueDate + ' (' + inv.daysOverdue + 'd ago)</span></div>';
    });
  }
  
  // Recent Errors (from HealthCheck)
  if (b.recentErrors.length > 0) {
    h += _secHdr('🔴 Errors (24h)', b.recentErrors.length, R);
    var maxErr = Math.min(b.recentErrors.length, 5);
    for (var ei = 0; ei < maxErr; ei++) {
      var err = b.recentErrors[ei];
      h += '<div style="padding:3px 16px;font-size:10px;color:#721c24;background:#fff5f5;border-bottom:1px solid #f0f0f0;">' +
        '<b>' + _escHtml(err.functionName || '') + '</b>: ' + _escHtml(err.error || '') + '</div>';
    }
    if (b.recentErrors.length > 5) {
      h += '<div style="padding:3px 16px;font-size:10px;color:#888;">+' + (b.recentErrors.length - 5) + ' more</div>';
    }
  }
  
  // All clear
  if (b.pendingUnscanned.length === 0 && b.scannedReady.length === 0 &&
      b.containersOverdue.length === 0 && b.overdueInvoices.length === 0 &&
      b.needsReview === 0 && b.recentErrors.length === 0) {
    h += '<div style="padding:20px;text-align:center;color:#888;font-size:13px;">✨ All clear — nothing needs attention.</div>';
  }
  
  // Footer
  h += '<div style="padding:10px 16px;background:#f9f9f9;text-align:center;font-size:9px;color:#aaa;">' +
    'CSS — "Integrity Through Independence" • ' +
    Utilities.formatDate(new Date(), BRIEFING_CONFIG.timezone, 'h:mm a') + ' • ' +
    '<a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '" style="color:' + G + ';">Open Sheet</a></div>';
  
  h += '</div></body></html>';
  return h;
}

// ── HTML Helpers (compact) ──

// ============================================================
// SIDEBAR HTML — Interactive in-app brief
// ============================================================

function _buildSidebarHtml(b) {
  var G = '#2E5339';
  var s = '<html><head><style>';
  s += 'body{font-family:Arial,sans-serif;margin:0;padding:0;font-size:12px;color:#333;background:#fafafa;}';
  s += '.hdr{background:' + G + ';color:#fff;padding:12px 14px;text-align:center;}';
  s += '.hdr h2{margin:0;font-size:15px;letter-spacing:1px;} .hdr .sub{color:#8cc99e;font-size:10px;margin-top:3px;}';
  s += '.stats{display:flex;background:#fff;border-bottom:1px solid #eee;}';
  s += '.stat{flex:1;text-align:center;padding:10px 4px;} .stat .n{font-size:18px;font-weight:bold;color:' + G + ';} .stat .l{font-size:9px;color:#888;}';
  s += '.alert{margin:6px 10px;padding:6px 10px;border-radius:4px;font-size:11px;font-weight:bold;}';
  s += '.alert-r{background:#f8d7da;color:#721c24;} .alert-y{background:#fff3cd;color:#856404;} .alert-g{background:#d4edda;color:#155724;}';
  s += '.sec{margin-top:8px;} .sec-h{background:#f0f0f0;padding:6px 12px;font-weight:bold;font-size:11px;border-top:1px solid #e0e0e0;cursor:pointer;} .sec-h:hover{background:#e8e8e8;}';
  s += '.sec-h .cnt{float:right;color:#888;font-weight:normal;}';
  s += '.row{padding:4px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:#fff;}';
  s += '.row:hover{background:#f5f9ff;} .row b{font-family:monospace;font-size:10px;} .row .meta{color:#888;}';
  s += '.grp{padding:3px 12px;font-size:10px;color:#666;font-weight:bold;background:#f7f7f7;}';
  s += '.vol{padding:8px 12px;background:#fff;}';
  s += '.vol table{width:100%;font-size:11px;border-collapse:collapse;} .vol td{padding:3px 4px;} .vol .lbl{color:#888;} .vol .num{text-align:center;font-weight:bold;}';
  s += '.inv-row{padding:4px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:#fff8f8;}';
  s += '.inv-row .amt{float:right;color:#d32f2f;font-weight:bold;}';
  s += '.err-row{padding:3px 12px;font-size:10px;color:#721c24;background:#fff5f5;border-bottom:1px solid #f0f0f0;}';
  s += '.foot{padding:10px;text-align:center;font-size:9px;color:#aaa;} .foot a{color:' + G + ';}';
  s += '.hidden{display:none;}';
  s += '.btn-row{padding:8px 12px;text-align:center;background:#fff;border-bottom:1px solid #eee;}';
  s += '.btn{padding:6px 16px;font-size:11px;border:none;border-radius:4px;cursor:pointer;margin:0 4px;}';
  s += '.btn-refresh{background:' + G + ';color:#fff;} .btn-email{background:#eee;color:#333;}';
  s += '</style></head><body>';

  // Header
  s += '<div class="hdr"><h2>☀️ QUICK BRIEF</h2><div class="sub">' + b.date + ' • ' + b.time + '</div></div>';

  // Big stat boxes
  s += '<div class="stats">';
  s += '<div class="stat"><div class="n">' + b.pendingUnscanned.length + '</div><div class="l">⏳ Pending</div></div>';
  s += '<div class="stat"><div class="n">' + b.scannedReady.length + '</div><div class="l">📦 Ready</div></div>';
  s += '<div class="stat"><div class="n">' + b.statsTotal.shipped + '</div><div class="l">✅ Shipped</div></div>';
  s += '</div>';

  // Buttons
  s += '<div class="btn-row">';
  s += '<button class="btn btn-refresh" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close();google.script.run.showBriefingNow();}).withFailureHandler(function(){}).showBriefingNow()">↻ Refresh</button>';
  s += '<button class="btn btn-email" onclick="google.script.run.sendMorningBriefing()">📧 Email Brief</button>';
  s += '</div>';

  // Alerts
  var hasAlerts = false;
  if (b.needsReview > 0) { s += '<div class="alert alert-y">⚠️ ' + b.needsReview + ' orders need review</div>'; hasAlerts = true; }
  if (b.containersOverdue.length > 0) { s += '<div class="alert alert-r">🚨 ' + b.containersOverdue.length + ' containers overdue</div>'; hasAlerts = true; }
  if (b.overdueInvoices.length > 0) { s += '<div class="alert alert-r">💰 ' + b.overdueInvoices.length + ' invoices overdue</div>'; hasAlerts = true; }
  if (b.recentErrors.length > 0) { s += '<div class="alert alert-r">🔴 ' + b.recentErrors.length + ' errors in last 24h</div>'; hasAlerts = true; }
  if (!hasAlerts) { s += '<div class="alert alert-g">✅ No alerts</div>'; }

  // Volume stats
  s += '<div class="vol"><table>';
  s += '<tr><td class="lbl"></td><td class="num">In</td><td class="num">Out</td></tr>';
  s += '<tr><td class="lbl">Yesterday</td><td class="num">' + b.receivedYesterday.length + '</td><td class="num">' + b.shippedYesterday.length + '</td></tr>';
  s += '<tr><td class="lbl">7 days</td><td class="num">' + b.statsWeek.received + '</td><td class="num">' + b.statsWeek.shipped + '</td></tr>';
  s += '<tr><td class="lbl">30 days</td><td class="num">' + b.statsMonth.received + '</td><td class="num">' + b.statsMonth.shipped + '</td></tr>';
  s += '</table></div>';

  // By-client yesterday
  var clientKeys = Object.keys(b.byClient).sort(function(a,c) { return b.byClient[c] - b.byClient[a]; });
  if (clientKeys.length > 0) {
    s += _sidebarSection('📥 Yesterday by Client', clientKeys.length + ' clients');
    for (var ci = 0; ci < clientKeys.length; ci++) {
      var ck = clientKeys[ci];
      s += '<div class="row"><span class="meta">' + b.byClient[ck] + '×</span> ' + _escHtml(ck) + '</div>';
    }
  }

  // Pending Unscanned
  if (b.pendingUnscanned.length > 0) {
    s += _sidebarSection('⚠️ Pending Unscanned', b.pendingUnscanned.length);
    var byWh = _groupBy(b.pendingUnscanned, 'warehouse');
    Object.keys(byWh).sort().forEach(function(wh) {
      s += '<div class="grp">📍 ' + (wh || 'Unknown') + ' (' + byWh[wh].length + ')</div>';
      byWh[wh].forEach(function(o) {
        s += '<div class="row"><b>' + _escHtml(o.csSample) + '</b> <span class="meta">' + _escHtml(o.sender) + '</span></div>';
      });
    });
  }

  // Ready to Ship
  if (b.scannedReady.length > 0) {
    s += _sidebarSection('📦 Ready to Ship', b.scannedReady.length);
    var byRecv = _groupBy(b.scannedReady, 'receiver');
    Object.keys(byRecv).sort().forEach(function(recv) {
      s += '<div class="grp">📬 ' + (recv || 'No receiver') + ' (' + byRecv[recv].length + ')</div>';
      byRecv[recv].forEach(function(o) {
        s += '<div class="row"><b>' + _escHtml(o.csSample) + '</b> <span class="meta">' + _escHtml(o.sender) + '</span></div>';
      });
    });
  }

  // Shipped Yesterday
  if (b.shippedYesterday.length > 0) {
    s += _sidebarSection('✅ Shipped Yesterday', b.shippedYesterday.length);
    b.shippedYesterday.forEach(function(o) {
      s += '<div class="row"><b>' + _escHtml(o.csSample) + '</b> <span class="meta">' + _escHtml(o.sender) + '</span>' +
        (o.tracking ? ' <span class="meta">📦' + _escHtml(o.tracking) + '</span>' : '') + '</div>';
    });
  }

  // Container ETAs
  if (b.containersToday.length > 0) {
    s += _sidebarSection('🚢 Arriving Today', b.containersToday.length);
    b.containersToday.forEach(function(o) {
      s += '<div class="row"><b>' + _escHtml(o.container) + '</b> <span class="meta">' + _escHtml(o.sender) + ' · ' + _escHtml(o.warehouse) + '</span></div>';
    });
  }
  if (b.containersOverdue.length > 0) {
    s += _sidebarSection('🚨 Containers Overdue', b.containersOverdue.length);
    b.containersOverdue.forEach(function(o) {
      s += '<div class="row" style="background:#fff5f5;"><b>' + _escHtml(o.container) + '</b> <span class="meta">ETA: ' + _escHtml(o.containerETA) + '</span></div>';
    });
  }

  // Overdue Invoices
  if (b.overdueInvoices.length > 0) {
    s += _sidebarSection('💰 Overdue Invoices', b.overdueInvoices.length);
    b.overdueInvoices.forEach(function(inv) {
      s += '<div class="inv-row"><b>' + _escHtml(inv.invoiceNum) + '</b> ' + _escHtml(inv.customer) +
        '<span class="amt">$' + Number(inv.total).toFixed(2) + '</span>' +
        '<br><span class="meta">Due ' + inv.dueDate + ' (' + inv.daysOverdue + 'd)</span></div>';
    });
  }

  // Recent Errors
  if (b.recentErrors.length > 0) {
    s += _sidebarSection('🔴 Errors (24h)', b.recentErrors.length);
    var maxErr = Math.min(b.recentErrors.length, 5);
    for (var ei = 0; ei < maxErr; ei++) {
      var err = b.recentErrors[ei];
      s += '<div class="err-row"><b>' + _escHtml(err.functionName || '') + '</b>: ' + _escHtml(err.error || '') + '</div>';
    }
    if (b.recentErrors.length > 5) s += '<div class="row" style="color:#888;">+' + (b.recentErrors.length - 5) + ' more</div>';
  }

  // Footer
  s += '<div class="foot">Updated ' + b.time + '</div>';
  s += '</body></html>';
  return s;
}

function _sidebarSection(title, count) {
  return '<div class="sec"><div class="sec-h">' + title + '<span class="cnt">' + count + '</span></div></div>';
}

function _statTd(label, value) {
  return '<td style="text-align:center;padding:10px 6px;width:33%;">' +
    '<div style="font-size:9px;color:#888;">' + label + '</div>' +
    '<div style="font-size:12px;font-weight:bold;color:#333;">' + value + '</div></td>';
}

function _secHdr(title, count, bg) {
  return '<div style="padding:6px 16px;background:' + (bg || '#f5f5f5') + ';border-top:1px solid #eee;">' +
    '<b style="font-size:12px;">' + title + '</b>' +
    '<span style="float:right;font-size:10px;color:#888;">' + count + '</span></div>';
}

function _row(id, sender, detail, bg) {
  return '<div style="padding:3px 16px;border-bottom:1px solid #f5f5f5;font-size:11px;background:' + (bg || '#fff') + ';">' +
    '<b style="font-family:monospace;font-size:10px;">' + _escHtml(id) + '</b>' +
    '<span style="color:#888;"> ' + _escHtml(_truncate(sender, 20)) + '</span>' +
    (detail ? ' <span style="color:#555;">' + _escHtml(_truncate(detail, 40)) + '</span>' : '') +
    '</div>';
}

function _compactTable(items, extraField, bg) {
  var out = '';
  items.forEach(function(o) {
    var extra = extraField && o[extraField] ? o[extraField] : '';
    out += _row(o.csSample, o.sender, (o.desc || '') + (extra ? ' · ' + extra : ''), bg);
  });
  return out;
}

function _groupBy(arr, field) {
  var groups = {};
  arr.forEach(function(o) {
    var key = o[field] || 'Unknown';
    if (!groups[key]) groups[key] = [];
    groups[key].push(o);
  });
  return groups;
}

function _truncate(s, max) {
  if (!s) return '';
  s = String(s);
  return s.length > max ? s.substring(0, max - 1) + '…' : s;
}

function _escHtml(s) {
  if (!s) return '';
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function _startOfDay(d) {
  var s = new Date(d);
  s.setHours(0, 0, 0, 0);
  return s;
}