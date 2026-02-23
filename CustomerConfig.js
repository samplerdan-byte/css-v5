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

    // Normalize key to lowercase
    var key = String(customerKey).toLowerCase().trim();
    if (!key) return null;

    // Direct lookup
    if (KNOWN_CLIENTS[key]) {
      var client = KNOWN_CLIENTS[key];
      Logger.log('getCustomerDefaults: exact match for "' + key + '"');
      return {
        name: client.name,
        shipping: client.defaultShipping || 'FedEx',
        sampleWeight: client.defaultSampleSize || '2 lb',
        warehouse: client.warehouse || '',
        fedexAccount: client.fedexAccount || '',
        defaultReceiver: client.defaultReceiver || null
      };
    }

    // Try partial match on name
    for (var k in KNOWN_CLIENTS) {
      if (KNOWN_CLIENTS[k].name.toLowerCase().indexOf(key) >= 0) {
        var client2 = KNOWN_CLIENTS[k];
        Logger.log('getCustomerDefaults: partial match "' + key + '" -> "' + client2.name + '"');
        return {
          name: client2.name,
          shipping: client2.defaultShipping || 'FedEx',
          sampleWeight: client2.defaultSampleSize || '2 lb',
          warehouse: client2.warehouse || '',
          fedexAccount: client2.fedexAccount || '',
          defaultReceiver: client2.defaultReceiver || null
        };
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
      var allText = (content || '') + '\n' + (subject || '');
      var result = identifyClient(senderEmail, allText);
      Logger.log('detectCustomer: sender="' + (senderEmail || '') + '" -> ' + (result ? result.name || JSON.stringify(result) : 'null'));
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
      var result = extractOrderFromEmail(emailBody, '', fromEmail, subject);
      if (result && result.success && result.orders && result.orders.length > 0) {
        var order = result.orders[0];
        Logger.log('parseEmailOrder: parsed order for client "' + (result.client ? result.client.name : 'Unknown') + '"');
        return {
          client: result.client ? result.client.name : 'Unknown',
          sampleOrderNum: order.sampleOrderNum || '',
          container: order.container || '',
          mark: order.mark || '',
          cargo: order.cargo || '',
          reference: order.reference || '',
          warehouse: order.warehouse || '',
          receiver: order.receiver || '',
          sampleSize: order.sampleSize || '2 lb',
          shipping: order.shipping || 'FedEx',
          description: order.description || '',
          bags: order.bags || '',
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
    if (!text) {
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
      fields.containers = extractContainers(text) || [];
    }
    if (typeof extractMarks === 'function') {
      fields.marks = extractMarks(text) || [];
    }
    if (typeof extractCargos === 'function') {
      fields.cargos = extractCargos(text) || [];
    }
    if (typeof extractBagCounts === 'function') {
      fields.bags = extractBagCounts(text) || [];
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



