// ============================================================
// ManualEntry.gs — Manual order entry sidebar + web app
//
// CLEANED Feb 15, 2026:
//   - Removed dead processManualEntry() (superseded by processManualOrderEntry)
//   - Removed dead processPastedText() (superseded by manual entry sidebar)
// ============================================================

// ============================================================
// MANUAL ORDER ENTRY (primary function — used by sidebar)
// ============================================================

function processManualOrderEntry(orderData) {
  if (!orderData || typeof orderData !== 'object') {
    return { success: false, message: 'No order data provided.' };
  }
  if (!Array.isArray(orderData.samples) || orderData.samples.length === 0) {
    return { success: false, message: 'No samples provided. Add at least one sample.' };
  }

  Logger.log('processManualOrderEntry: ' + orderData.samples.length + ' sample(s), customer=' + String(orderData.customer || 'none').substring(0, 100).replace(/[\r\n]/g, ' '));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mainSheetName);

  if (!sheet) return { success: false, message: 'Sheet "' + CONFIG.mainSheetName + '" not found.' };

  if (orderData.samples && orderData.samples.length > 0) {
    var validation = validateOrderData(orderData.samples[0]);
    if (!validation.valid) {
      return { success: false, message: '⚠️ Validation error:\n' + validation.errors.join('\n') };
    }
  }

  try {
    const results = [];
    const customer = (typeof CUSTOMERS !== 'undefined' && CUSTOMERS[orderData.customer]) 
      ? CUSTOMERS[orderData.customer] 
      : { name: orderData.customer || 'Manual Entry' };
    
    // Build sender string from sender object or fall back to customer name
    var senderStr = '';
    if (orderData.sender && orderData.sender.company) {
      senderStr = orderData.sender.company;
    } else {
      senderStr = customer.name;
    }
    
    // Build receiver string from receiver object
    let receiverStr = '';
    if (orderData.receiver && orderData.receiver.company) {
      receiverStr = orderData.receiver.company;
      if (orderData.receiver.attention) receiverStr += ', Attn: ' + orderData.receiver.attention;
    }
    
    let addressStr = '';
    if (orderData.receiver) {
      if (orderData.receiver.address) addressStr += orderData.receiver.address;
      if (orderData.receiver.city) addressStr += (addressStr ? ', ' : '') + orderData.receiver.city;
      if (orderData.receiver.state) addressStr += (addressStr ? ', ' : '') + orderData.receiver.state;
      if (orderData.receiver.zip) addressStr += ' ' + orderData.receiver.zip;
    }
    
    for (const sample of orderData.samples) {
      const sampleData = {
        sampleOrderNum: orderData.orderNumber || '',
        cargo: sample.cargo || '',
        mark: sample.marks || '',
        container: sample.container || '',
        reference: sample.reference || orderData.orderNumber || '',
        description: sample.description || '',
        bagCount: sample.bags || '',
        weight: '',
        sampleWeight: orderData.sampleWeight || customer.defaultSampleSize || '2 LB',
        pNumber: sample.pNumber || '',
        sNumber: sample.sNumber || '',
        warehouse: (typeof WAREHOUSES !== 'undefined' && WAREHOUSES[orderData.warehouse]) 
          ? WAREHOUSES[orderData.warehouse].fullName 
          : (orderData.warehouse || ''),
        sender: senderStr,
        receiver: receiverStr,
        shippingProcess: orderData.shipping || customer.defaultShipping || 'FedEx Overnight',
        comments: 'Manual Entry' + (orderData.fedexAccount ? ' | FedEx: ' + orderData.fedexAccount : ''),
        sourceEmail: orderData.orderNumber || senderStr || 'Manual Entry',
        sampleType: orderData.sampleType || 'Warehouse Sample',
        containerETA: orderData.eta || '',
        shippingLine: orderData.shippingLine || '',
        shippingNotes: orderData.shippingNotes || '',
        bol: orderData.bolNumber || '',
        shipStatus: orderData.shipStatus || ''
      };
      
      addDataToSheet(sheet, sampleData, { skipLiveUpdate: true, skipAutoPrint: true });
      results.push(sample.cargo || sample.container || 'Sample');
    }
    
    // Single live update after all samples added (avoids re-triggering sidebar)
    try { updateLiveOrdersView(); } catch (e) { Logger.log('Live update after manual entry: ' + e); }

    Logger.log('processManualOrderEntry success: added ' + results.length + ' sample(s)');
    return { success: true, message: '✅ Added ' + results.length + ' sample(s): ' + results.join(', ') };
  } catch (e) {
    Logger.log('processManualOrderEntry error: ' + String(e).substring(0, 500).replace(/[\r\n]/g, ' '));
    return { success: false, message: 'Error processing manual order entry. Please try again.' };
  }
}

// ============================================================
// SHOW MANUAL ENTRY SIDEBAR
// ============================================================

function showEnhancedManualEntry() {
  var customers = typeof getCustomerList === 'function' ? getCustomerList() : [];
  var warehouses = typeof getWarehouseList === 'function' ? getWarehouseList() : [];
  var contacts = typeof getContactsList === 'function' ? getContactsList() : { shippers: [], receivers: [] };

  // HTML-escape helper to prevent XSS via sheet data in option tags
  function escHtml(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  let customerOptions = '<option value="">-- Select or type new --</option>';
  for (const c of customers) {
    customerOptions += '<option value="' + escHtml(c.key) + '">' + escHtml(c.name) + '</option>';
  }

  let warehouseOptions = '<option value="">-- Select --</option>';
  for (const w of warehouses) {
    warehouseOptions += '<option value="' + escHtml(w.key) + '">' + escHtml(w.name) + '</option>';
  }

  const template = HtmlService.createTemplateFromFile('ManualOrderEntry');
  template.customerOptions = customerOptions;
  template.warehouseOptions = warehouseOptions;
  template.contactsJson = JSON.stringify(contacts);
  
  const html = template.evaluate()
    .setWidth(450).setHeight(750).setTitle('New Sample Order');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// WEB APP ENTRY POINT — moved to WebApp.js to avoid duplicates
// ============================================================
