// ============================================================
// 🌐 CLIENT PORTAL — Read-Only Status Portal for Trading Houses
// ============================================================
// Gives customers a branded link to check their own order status.
// Token-based access — no Google login required.
//
// URL format: your-webapp-url?page=portal&token=XXXXXXXX
//
// Setup: Run setupPortalAccess() to create the access sheet,
//        then use generatePortalToken('Customer Name') to issue tokens.
// ============================================================

var PORTAL_CONFIG = {
  accessSheetName: 'Portal Access',
  tokenLength: 12
};

// ============================================================
// SETUP — Creates Portal Access sheet
// ============================================================

function setupPortalAccess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var sheet = ss.getSheetByName(PORTAL_CONFIG.accessSheetName);
  if (sheet) {
    ui.alert('Portal Access sheet already exists.\nEdit it directly to manage client tokens.');
    ss.setActiveSheet(sheet);
    return;
  }
  
  sheet = ss.insertSheet(PORTAL_CONFIG.accessSheetName);
  var headers = ['Customer Name', 'Match Pattern', 'Access Token', 'Active', 'Created', 'Last Login', 'Login Count', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 200);
  
  // Checkboxes for Active column
  sheet.getRange(2, 4, 100, 1).insertCheckboxes();
  
  ui.alert(
    '✅ Portal Access sheet created!\n\n' +
    'To add a client:\n' +
    '1. Run "Generate Portal Token" from the Billing/Portal menu\n' +
    '2. Or manually add rows with a unique token\n\n' +
    'Share the portal link with clients:\n' +
    'your-webapp-url?page=portal&token=THEIR_TOKEN'
  );
}

// ============================================================
// GENERATE TOKEN — Creates a unique access token for a customer
// ============================================================

function generatePortalToken(customerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PORTAL_CONFIG.accessSheetName);
  
  if (!sheet) {
    setupPortalAccess();
    sheet = ss.getSheetByName(PORTAL_CONFIG.accessSheetName);
  }
  
  // Generate random token
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  var token = '';
  for (var i = 0; i < PORTAL_CONFIG.tokenLength; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  
  // If no customerName provided, prompt
  if (!customerName) {
    var ui = SpreadsheetApp.getUi();
    var resp = ui.prompt('Generate Portal Token', 'Enter customer name:', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    customerName = resp.getResponseText().trim();
    if (!customerName) return;
  }
  
  // Default match pattern = customer name (can be edited)
  var matchPattern = customerName;
  
  // Check for existing token for this customer
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === customerName.toLowerCase()) {
        var ui2 = SpreadsheetApp.getUi();
        var resp2 = ui2.alert(
          'Token Exists',
          customerName + ' already has a token: ' + data[i][2] + '\n\nGenerate a NEW token (old one will stop working)?',
          ui2.ButtonSet.YES_NO
        );
        if (resp2 !== ui2.Button.YES) return;
        // Update existing row
        sheet.getRange(i + 2, 3).setValue(token);
        sheet.getRange(i + 2, 5).setValue(new Date());
        
        _showPortalLink(customerName, token);
        return;
      }
    }
  }
  
  // Add new row
  sheet.appendRow([customerName, matchPattern, token, true, new Date(), '', 0, '']);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 4).insertCheckboxes();
  sheet.getRange(lastRow, 4).setValue(true);
  
  _showPortalLink(customerName, token);
}

function _showPortalLink(customerName, token) {
  var webAppUrl = ScriptApp.getService().getUrl();
  var portalUrl = webAppUrl + '?page=portal&token=' + token;
  
  var ui = SpreadsheetApp.getUi();
  
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:15px;">' +
    '<h3 style="color:#2E5339;">✅ Portal Token Generated</h3>' +
    '<p><strong>Customer:</strong> ' + customerName + '</p>' +
    '<p><strong>Token:</strong> <code style="background:#f0f0f0;padding:3px 8px;border-radius:4px;font-size:16px;">' + token + '</code></p>' +
    '<p style="margin-top:15px;"><strong>Portal Link:</strong></p>' +
    '<input type="text" value="' + portalUrl + '" readonly ' +
    'style="width:100%;padding:10px;font-size:12px;border:2px solid #2E5339;border-radius:6px;margin-bottom:10px;" ' +
    'onclick="this.select();">' +
    '<p style="font-size:11px;color:#666;">Share this link with ' + customerName + '. They can bookmark it — no login needed.</p>' +
    '<button onclick="navigator.clipboard.writeText(\'' + portalUrl + '\');this.textContent=\'✅ Copied!\';setTimeout(function(){},2000);" ' +
    'style="padding:10px 20px;background:#2E5339;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;">📋 Copy Link</button>' +
    '</div>'
  ).setWidth(500).setHeight(320);
  
  ui.showModalDialog(html, 'Portal Link for ' + customerName);
}

// ============================================================
// MANAGE TOKENS — Bulk generate, revoke, list
// ============================================================

function revokePortalToken() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PORTAL_CONFIG.accessSheetName);
  var ui = SpreadsheetApp.getUi();
  
  if (!sheet) { ui.alert('No Portal Access sheet. Run Setup first.'); return; }
  
  var row = sheet.getActiveCell().getRow();
  if (row < 2 || sheet !== ss.getActiveSheet()) {
    ui.alert('Select a row in the Portal Access sheet to revoke.');
    return;
  }
  
  var customer = sheet.getRange(row, 1).getValue();
  var resp = ui.alert('Revoke Access', 'Disable portal access for ' + customer + '?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  
  sheet.getRange(row, 4).setValue(false); // Active = false
  ui.alert('✅ Portal access revoked for ' + customer);
}

function listPortalLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PORTAL_CONFIG.accessSheetName);
  var ui = SpreadsheetApp.getUi();
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('No portal tokens found. Generate tokens first.');
    return;
  }
  
  var webAppUrl = ScriptApp.getService().getUrl();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  
  var rows = '';
  data.forEach(function(row) {
    var active = row[3];
    var token = row[2];
    var url = webAppUrl + '?page=portal&token=' + token;
    var logins = row[6] || 0;
    var lastLogin = row[5] ? new Date(row[5]).toLocaleDateString() : 'Never';
    
    rows += '<tr style="' + (active ? '' : 'opacity:0.4;') + '">' +
      '<td>' + row[0] + '</td>' +
      '<td><code>' + token + '</code></td>' +
      '<td>' + (active ? '✅' : '❌') + '</td>' +
      '<td>' + logins + '</td>' +
      '<td>' + lastLogin + '</td>' +
      '<td><a href="' + url + '" target="_blank" style="color:#1976d2;">Open</a></td>' +
      '</tr>';
  });
  
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:15px;}table{width:100%;border-collapse:collapse;}th{background:#2E5339;color:#fff;padding:8px;text-align:left;font-size:11px;}td{padding:6px 8px;border-bottom:1px solid #eee;font-size:11px;}code{background:#f0f0f0;padding:2px 6px;border-radius:3px;}</style>' +
    '<h3 style="color:#2E5339;">🔗 Portal Links</h3>' +
    '<table><thead><tr><th>Customer</th><th>Token</th><th>Active</th><th>Logins</th><th>Last Login</th><th>Link</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>'
  ).setWidth(700).setHeight(400);
  
  ui.showModalDialog(html, 'Portal Links');
}

// ============================================================
// VALIDATE TOKEN — Called by portal doGet
// ============================================================

function _validatePortalToken(token) {
  if (!token) return null;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PORTAL_CONFIG.accessSheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][2]).trim() === token && data[i][3] === true) {
      // Update last login + count
      var row = i + 2;
      sheet.getRange(row, 6).setValue(new Date());
      var count = parseInt(data[i][6]) || 0;
      sheet.getRange(row, 7).setValue(count + 1);
      
      return {
        customerName: String(data[i][0]).trim(),
        matchPattern: String(data[i][1]).trim()
      };
    }
  }
  
  return null;
}

// ============================================================
// GET PORTAL DATA — Filtered to one customer
// ============================================================

function getPortalData(token) {
  var auth = _validatePortalToken(token);
  if (!auth) return { success: false, message: 'Invalid or expired access token.' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pattern = auth.matchPattern.toLowerCase();
  
  var result = {
    success: true,
    customerName: auth.customerName,
    orders: [],
    stats: { total: 0, received: 0, scanned: 0, shipped: 0 }
  };
  
  // Search main sheet and completed orders
  var sheetNames = [CONFIG.mainSheetName, CONFIG.completedOrdersSheetName || 'Completed Orders'];
  
  sheetNames.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    
    var col = _getColumnMap(sheet);
    if (col['Sender'] === undefined) return;
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var sender = String(data[i][col['Sender']] || '').toLowerCase();
      
      if (sender.indexOf(pattern) !== -1) {
        var status = String(data[i][col['Status']] || 'Received').trim();
        var trackingNum = col['Tracking Number'] !== undefined ? String(data[i][col['Tracking Number']] || '') : '';
        var trackingUrl = '';
        if (trackingNum && typeof getTrackingUrl === 'function') {
          trackingUrl = getTrackingUrl(trackingNum);
        }
        
        var order = {
          csSample: String(data[i][col['CS Sample #']] || ''),
          csOrder: String(data[i][col['CS Order #']] || ''),
          sampleOrderNum: String(data[i][col['Sample Order #']] || ''),
          description: String(data[i][col['Description']] || ''),
          container: String(data[i][col['Container #']] || ''),
          cargo: String(data[i][col['Cargo #']] || ''),
          mark: String(data[i][col['Mark #']] || ''),
          receiver: String(data[i][col['Receiver']] || ''),
          warehouse: String(data[i][col['Warehouse']] || ''),
          bagCount: String(data[i][col['Bag Count']] || ''),
          sampleWeight: String(data[i][col['Sample Weight']] || ''),
          status: status,
          trackingNumber: trackingNum,
          trackingUrl: trackingUrl,
          containerStatus: col['Container Status'] !== undefined ? String(data[i][col['Container Status']] || '') : '',
          containerETA: col['Container ETA'] !== undefined ? String(data[i][col['Container ETA']] || '') : '',
          shippedDate: col['Shipped Date'] !== undefined ? _portalFormatDate(data[i][col['Shipped Date']]) : '',
          receivedDate: _portalFormatDate(data[i][col['Timestamp']]),
          source: sheetName
        };
        
        result.orders.push(order);
        result.stats.total++;
        
        var sl = status.toLowerCase();
        if (sl === 'received') result.stats.received++;
        else if (sl === 'scanned') result.stats.scanned++;
        else if (sl === 'shipped' || sl === 'completed') result.stats.shipped++;
      }
    }
  });
  
  // Sort: active first (Received, Scanned), then shipped, by date desc
  var statusOrder = { 'Received': 0, 'Scanned': 1, 'Shipped': 2, 'Completed': 3 };
  result.orders.sort(function(a, b) {
    var sa = statusOrder[a.status] !== undefined ? statusOrder[a.status] : 9;
    var sb = statusOrder[b.status] !== undefined ? statusOrder[b.status] : 9;
    if (sa !== sb) return sa - sb;
    // Within same status, newest first
    return (b.receivedDate || '').localeCompare(a.receivedDate || '');
  });
  
  return result;
}

function _portalFormatDate(d) {
  if (!d) return '';
  try {
    var dt = new Date(d);
    if (isNaN(dt.getTime())) return String(d);
    return String(dt.getMonth() + 1).padStart(2, '0') + '/' +
      String(dt.getDate()).padStart(2, '0') + '/' + dt.getFullYear();
  } catch(e) { return ''; }
}

// ============================================================
// SERVE PORTAL PAGE — Called from doGet when page=portal
// ============================================================

function servePortalPage(token) {
  return HtmlService.createTemplateFromFile('ClientPortal')
    .evaluate()
    .setTitle('CSS — Order Status Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// UPDATE doGet to route portal requests
// ============================================================
// 
// In your WebApp.gs doGet(), add this BEFORE the default return:
//
//   if (page === 'portal') {
//     return servePortalPage(e.parameter.token || '');
//   }
//
// ============================================================

// ============================================================
// MENU ITEMS — Add to onOpen
// ============================================================
//
// Add to your ☕ CSS System menu:
//
//   .addSubMenu(ui.createMenu('🌐 Client Portal')
//     .addItem('⚙️ Setup Portal Access', 'setupPortalAccess')
//     .addItem('🔑 Generate Token', 'generatePortalToken')
//     .addItem('🔗 List All Portal Links', 'listPortalLinks')
//     .addItem('🚫 Revoke Access (selected row)', 'revokePortalToken'))
//
// ============================================================
