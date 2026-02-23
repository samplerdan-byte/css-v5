// ============================================================
// Email.gs — Customer email reports, Gmail helpers, attachments
// ============================================================

// ============================================================
// CUSTOMER EMAIL FUNCTIONS
// ============================================================

function sendDailyCustomerReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.mainSheetName);
  const emailSheet = ss.getSheetByName('Customer Emails');
  const ui = SpreadsheetApp.getUi();

  if (!mainSheet) { ui.alert('Main sheet not found!'); return; }
  if (!emailSheet) { ui.alert('Customer Emails sheet not found! Run Setup → Customer Email Sheet first.'); return; }

  var col = _getColumnMap(mainSheet);
  const lastRow = mainSheet.getLastRow();
  const data = mainSheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn()).getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Filter shipped today
  const shippedToday = data.filter(row => {
    const shippedDate = row[col['Shipped Date']];
    if (!shippedDate) return false;
    const shipDate = new Date(shippedDate);
    shipDate.setHours(0, 0, 0, 0);
    return shipDate.getTime() === today.getTime();
  });

  if (shippedToday.length === 0) {
    ui.alert('No samples shipped today.');
    return;
  }

  // Get customer email mappings
  const emailLastRow = emailSheet.getLastRow();
  const emailData = emailSheet.getRange(2, 1, emailLastRow - 1, 5).getValues();

  const customerEmails = {};
  emailData.forEach(row => {
    const customerName = row[0];
    const matchPattern = row[1];
    const emails = row[2];
    const active = row[3];

    if (active && emails) {
      customerEmails[matchPattern] = {
        name: customerName,
        emails: emails.split(',').map(e => e.trim()).filter(e => e)
      };
    }
  });

  // Group shipped samples by customer
  const customerSamples = {};
  shippedToday.forEach(row => {
    const sender = row[col['Sender']] || 'Unknown';

    let matchedCustomer = null;
    for (const pattern in customerEmails) {
      if (sender.toLowerCase().includes(pattern.toLowerCase())) {
        matchedCustomer = pattern;
        break;
      }
    }

    if (matchedCustomer) {
      if (!customerSamples[matchedCustomer]) {
        customerSamples[matchedCustomer] = [];
      }
      customerSamples[matchedCustomer].push({
        csSample: row[col['CS Sample #']],
        container: row[col['Container #']],
        cargo: row[col['Cargo #']],
        mark: row[col['Mark #']],
        tracking: row[col['Tracking Number']],
        description: row[col['Description']]
      });
    }
  });

  if (Object.keys(customerSamples).length === 0) {
    ui.alert('No matching customers found for shipped samples.\n\nCheck your Customer Emails sheet patterns.');
    return;
  }

  // Preview before sending
  let preview = 'Ready to send reports:\n\n';
  for (const pattern in customerSamples) {
    const config = customerEmails[pattern];
    preview += `• ${config.name}: ${customerSamples[pattern].length} samples\n`;
    preview += `  To: ${config.emails.join(', ')}\n\n`;
  }

  const response = ui.alert('Send Daily Reports?', preview, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // Send emails
  let sent = 0;
  var sendResults = {};  // track per-pattern success/failure
  var failed = [];
  for (const pattern in customerSamples) {
    const config = customerEmails[pattern];
    const samples = customerSamples[pattern];

    const subject = 'CSS Daily Shipping Report - ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy');

    let body = '<h2>Commodity Sampler Services</h2>';
    body += '<h3>Daily Shipping Report for ' + config.name + '</h3>';
    body += '<p>Date: ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy') + '</p>';
    body += '<p>Samples Shipped: ' + samples.length + '</p>';
    body += '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">';
    body += '<tr style="background: #4A7C59; color: white;"><th>CS Sample #</th><th>Container</th><th>Cargo</th><th>Mark</th><th>Tracking</th></tr>';

    samples.forEach(s => {
      let trackingLink = s.tracking;
      if (s.tracking) {
        const url = getTrackingUrl(s.tracking);
        if (url) trackingLink = '<a href="' + url + '">' + s.tracking + '</a>';
      }
      body += '<tr>';
      body += '<td>' + s.csSample + '</td>';
      body += '<td>' + (s.container || 'N/A') + '</td>';
      body += '<td>' + (s.cargo || 'N/A') + '</td>';
      body += '<td>' + (s.mark || 'N/A') + '</td>';
      body += '<td>' + (trackingLink || 'N/A') + '</td>';
      body += '</tr>';
    });

    body += '</table>';
    body += '<p style="color: #666; margin-top: 20px;">This is an automated report from Commodity Sampler Services.</p>';

    try {
      GmailApp.sendEmail(config.emails.join(','), subject, '', { htmlBody: body });
      sent++;
      sendResults[pattern] = 'Sent';
      Logger.log('Sent report to ' + config.name + ': ' + config.emails.join(', '));
    } catch (e) {
      sendResults[pattern] = 'Failed: ' + e.message;
      failed.push(config.name + ': ' + e.message);
      Logger.log('Failed to send to ' + config.name + ': ' + e);
    }
  }

  // Log the reports — only log actual status per customer
  const reportLogSheet = ss.getSheetByName('Daily Report Log');
  if (reportLogSheet) {
    for (const pattern in customerSamples) {
      const config = customerEmails[pattern];
      reportLogSheet.appendRow([
        new Date(),
        config.name,
        customerSamples[pattern].length,
        config.emails.join(', '),
        sendResults[pattern] || 'Unknown',
        Session.getActiveUser().getEmail()
      ]);
    }
  }

  var msg = 'Sent ' + sent + ' customer reports.';
  if (failed.length > 0) {
    msg += '\n\nFAILED (' + failed.length + '):\n' + failed.join('\n');
  }
  ui.alert('Reports Status', msg, ui.ButtonSet.OK);
}

function getCustomerEmailList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName('Customer Emails');

  if (!emailSheet) return [];

  const lastRow = emailSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = emailSheet.getRange(2, 1, lastRow - 1, 5).getValues();

  return data.map(row => ({
    name: row[0],
    pattern: row[1],
    emails: row[2],
    active: row[3],
    notes: row[4]
  })).filter(c => c.name);
}

function lookupCustomerEmails(senderName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName('Customer Emails');

  if (!emailSheet) return null;

  const lastRow = emailSheet.getLastRow();
  if (lastRow < 2) return null;

  const data = emailSheet.getRange(2, 1, lastRow - 1, 5).getValues();

  for (var i = 0; i < data.length; i++) {
    var pattern = data[i][1];
    var emails = data[i][2];
    var active = data[i][3];

    if (active && pattern && senderName.toLowerCase().includes(pattern.toLowerCase())) {
      return emails;
    }
  }

  return null;
}

// ============================================================
// GMAIL MESSAGE LINK
// ============================================================

// getGmailMessageLink() — moved to Emailextraction_v3.js to avoid duplicates

function getEmailHtmlFromLink(emailLink) {
  if (!emailLink) return null;

  try {
    var match = String(emailLink).match(/\/([a-f0-9]+)$/i);
    if (!match) {
      Logger.log('Could not extract thread ID from: ' + emailLink);
      return null;
    }

    var threadId = match[1];
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      Logger.log('Thread not found: ' + threadId);
      return null;
    }

    var messages = thread.getMessages();
    if (!messages || messages.length === 0) return null;

    var msg = messages[0];
    var body = msg.getBody();
    var from = msg.getFrom();
    var subject = msg.getSubject();
    var date = msg.getDate();
    var dateStr = '';
    try {
      dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a');
    } catch(e) {
      dateStr = String(date);
    }

    return {
      from: from,
      subject: subject,
      date: dateStr,
      body: body
    };

  } catch (e) {
    Logger.log('Error fetching email: ' + e);
    return null;
  }
}

// ============================================================
// SAVE EMAIL ATTACHMENTS TO DRIVE
// ============================================================

// saveEmailPDFs() and saveEmailImages() — moved to Emailextraction_v3.js to avoid duplicates
