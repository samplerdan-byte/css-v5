// ============================================================================
// CustomerConfig.gs — Customer & Warehouse Configuration
// Commodity Sampler Services
// 
// This file provides the missing functions that are called from Code.gs
// and TestExamples.gs. It bridges to the KNOWN_CLIENTS and WAREHOUSES
// objects defined in Emailextraction v2.gs.
// ============================================================================


// ============================================================================
// CUSTOMERS OBJECT — Alias for backward compatibility
// Code.gs references CUSTOMERS but Emailextraction uses KNOWN_CLIENTS
// ============================================================================

var CUSTOMERS = KNOWN_CLIENTS;  // Alias to the master list in Emailextraction v2.gs


// ============================================================================
// getCustomerList() — Returns array of customers for dropdowns
// Called by: Code.gs (showEnhancedManualEntry)
// ============================================================================

function getCustomerList() {
  try {
    var list = [];
    for (var key in KNOWN_CLIENTS) {
      if (KNOWN_CLIENTS.hasOwnProperty(key)) {
        var client = KNOWN_CLIENTS[key];
        list.push({
          key: key,
          name: client.name,
          defaultShipping: client.defaultShipping || 'FedEx',
          defaultSampleSize: client.defaultSampleSize || '2 lb',
          warehouse: client.warehouse || '',
          fedexAccount: client.fedexAccount || ''
        });
      }
    }
    // Sort alphabetically by name
    list.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
    Logger.log('getCustomerList: returned ' + list.length + ' customers');
    return list;
  } catch (e) {
    Logger.log('getCustomerList ERROR: ' + e.message);
    return [];
  }
}


// ============================================================================
// getWarehouseList() — Returns array of warehouses for dropdowns
// Called by: Code.gs (showEnhancedManualEntry)
// ============================================================================

function getWarehouseList() {
  try {
    var list = [];
    for (var key in WAREHOUSES) {
      if (WAREHOUSES.hasOwnProperty(key)) {
        var wh = WAREHOUSES[key];
        list.push({
          key: key,
          name: wh.name,
          fullName: wh.name,
          fullAddress: wh.fullAddress || '',
          addresses: wh.addresses || []
        });
      }
    }
    Logger.log('getWarehouseList: returned ' + list.length + ' warehouses');
    return list;
  } catch (e) {
    Logger.log('getWarehouseList ERROR: ' + e.message);
    return [];
  }
}


// ============================================================================
// getCustomerDefaults() — Get default settings for a customer by key
// Called by: TestExamples.gs
// ============================================================================

function getCustomerDefaults(customerKey) {
  try {
    if (!customerKey) return null;

    // Normalize key to lowercase with null/undefined guard
    var key = String(customerKey || '').toLowerCase().trim();
    if (!key) return null;

    // Direct lookup
    if (KNOWN_CLIENTS[key]) {
      var client = KNOWN_CLIENTS[key];
      Logger.log('getCustomerDefaults: exact match for "' + key + '"');
      return {
        name: String(client.name || ''),
        shipping: String(client.defaultShipping || 'FedEx').trim(),
        sampleWeight: String(client.defaultSampleSize || '2 lb').trim(),
        warehouse: String(client.warehouse || '').trim(),
        fedexAccount: String(client.fedexAccount || '').trim(),
        defaultReceiver: client.defaultReceiver || null
      };
    }

    // Try partial match on name
    for (var k in KNOWN_CLIENTS) {
      if (KNOWN_CLIENTS.hasOwnProperty(k)) {
        var clientName = String(KNOWN_CLIENTS[k].name || '').toLowerCase();
        if (clientName.indexOf(key) >= 0) {
          var client2 = KNOWN_CLIENTS[k];
          Logger.log('getCustomerDefaults: partial match "' + key + '" -> "' + client2.name + '"');
          return {
            name: String(client2.name || ''),
            shipping: String(client2.defaultShipping || 'FedEx').trim(),
            sampleWeight: String(client2.defaultSampleSize || '2 lb').trim(),
            warehouse: String(client2.warehouse || '').trim(),
            fedexAccount: String(client2.fedexAccount || '').trim(),
            defaultReceiver: client2.defaultReceiver || null
          };
        }
      }
    }

    Logger.log('getCustomerDefaults: no match for "' + key + '"');
    return null;
  } catch (e) {
    Logger.log('getCustomerDefaults ERROR: ' + e.message);
    return null;
  }
}


// ============================================================================
// detectCustomer() — Identify customer from email content and sender
// Called by: TestExamples.gs (testCustomerConfig)
// Wrapper for identifyClient() in Emailextraction v2.gs
// ============================================================================

function detectCustomer(content, senderEmail, subject) {
  try {
    if (!senderEmail && !content && !subject) {
      Logger.log('detectCustomer: all inputs empty, returning null');
      return null;
    }
    // identifyClient is defined in Emailextraction v2.gs
    if (typeof identifyClient === 'function') {
      var contentStr = String(content || '').trim();
      var senderStr = String(senderEmail || '').trim();
      var subjectStr = String(subject || '').trim();
      var allText = contentStr + '\n' + subjectStr;
      var result = identifyClient(senderStr, allText);
      Logger.log('detectCustomer: sender="' + senderStr + '" -> ' + (result ? result.name || JSON.stringify(result) : 'null'));
      return result;
    }
    Logger.log('detectCustomer: identifyClient not available');
    return null;
  } catch (e) {
    Logger.log('detectCustomer ERROR: ' + e.message);
    return null;
  }
}


// ============================================================================
// parseEmailOrder() — Parse email content into order object
// Called by: TestExamples.gs (exampleParseEmail)
// Wrapper for extractOrderFromEmail() in Emailextraction v2.gs
// ============================================================================

function parseEmailOrder(emailBody, fromEmail, subject) {
  try {
    if (!emailBody && !fromEmail && !subject) {
      Logger.log('parseEmailOrder: all inputs empty');
      return null;
    }
    // extractOrderFromEmail is defined in Emailextraction v2.gs
    if (typeof extractOrderFromEmail === 'function') {
      var result = extractOrderFromEmail(String(emailBody || ''), '', String(fromEmail || ''), String(subject || ''));
      if (result && result.success && result.orders && result.orders.length > 0) {
        var order = result.orders[0];
        Logger.log('parseEmailOrder: parsed order for client "' + (result.client ? result.client.name : 'Unknown') + '"');
        return {
          client: String(result.client && result.client.name ? result.client.name : 'Unknown').trim(),
          sampleOrderNum: String(order.sampleOrderNum || '').trim(),
          container: String(order.container || '').trim(),
          mark: String(order.mark || '').trim(),
          cargo: String(order.cargo || '').trim(),
          reference: String(order.reference || '').trim(),
          warehouse: String(order.warehouse || '').trim(),
          receiver: String(order.receiver || '').trim(),
          sampleSize: String(order.sampleSize || '2 lb').trim(),
          shipping: String(order.shipping || 'FedEx').trim(),
          description: String(order.description || '').trim(),
          bags: String(order.bags || '').trim(),
          flags: result.flags || []
        };
      }
      Logger.log('parseEmailOrder: extraction returned no orders');
    } else {
      Logger.log('parseEmailOrder: extractOrderFromEmail not available');
    }
    return null;
  } catch (e) {
    Logger.log('parseEmailOrder ERROR: ' + e.message);
    return null;
  }
}


// ============================================================================
// extractCommonFields() — Extract container, mark, cargo from text
// Called by: TestExamples.gs (testCustomerConfig)
// Uses extraction functions from Emailextraction v2.gs
// ============================================================================

function extractCommonFields(text) {
  try {
    var textStr = String(text || '').trim();
    if (!textStr) {
      Logger.log('extractCommonFields: empty text input');
      return { containers: [], marks: [], cargos: [], bags: [] };
    }

    var fields = {
      containers: [],
      marks: [],
      cargos: [],
      bags: []
    };

    // Use the extraction functions from Emailextraction v2.gs
    if (typeof extractContainers === 'function') {
      fields.containers = extractContainers(textStr) || [];
    }
    if (typeof extractMarks === 'function') {
      fields.marks = extractMarks(textStr) || [];
    }
    if (typeof extractCargos === 'function') {
      fields.cargos = extractCargos(textStr) || [];
    }
    if (typeof extractBagCounts === 'function') {
      fields.bags = extractBagCounts(textStr) || [];
    }

    Logger.log('extractCommonFields: containers=' + fields.containers.length +
               ' marks=' + fields.marks.length +
               ' cargos=' + fields.cargos.length +
               ' bags=' + fields.bags.length);
    return fields;
  } catch (e) {
    Logger.log('extractCommonFields ERROR: ' + e.message);
    return { containers: [], marks: [], cargos: [], bags: [] };
  }
}



