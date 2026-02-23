// ============================================================================
// CustomerConfig.gs — Customer & Warehouse Configuration
// Commodity Sampler Services
// 
// This file provides functions called from Code.gs, InputManual.gs, and
// WebApp.gs. It bridges to the KNOWN_CLIENTS and WAREHOUSES objects
// defined in Emailextraction v2.gs.
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
    if (!KNOWN_CLIENTS || typeof KNOWN_CLIENTS !== 'object') {
      Logger.log('getCustomerList: KNOWN_CLIENTS not available');
      return [];
    }
    for (var key in KNOWN_CLIENTS) {
      if (KNOWN_CLIENTS.hasOwnProperty(key)) {
        var client = KNOWN_CLIENTS[key];
        if (!client) continue; // Guard against null entries
        list.push({
          key: key,
          name: String(client.name || ''),
          defaultShipping: String(client.defaultShipping || 'FedEx'),
          defaultSampleSize: String(client.defaultSampleSize || '2 lb'),
          warehouse: String(client.warehouse || ''),
          fedexAccount: String(client.fedexAccount || '')
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
    if (!WAREHOUSES || typeof WAREHOUSES !== 'object') {
      Logger.log('getWarehouseList: WAREHOUSES not available');
      return [];
    }
    for (var key in WAREHOUSES) {
      if (WAREHOUSES.hasOwnProperty(key)) {
        var wh = WAREHOUSES[key];
        if (!wh) continue; // Guard against null entries
        list.push({
          key: key,
          name: String(wh.name || ''),
          fullName: String(wh.name || ''),
          fullAddress: String(wh.fullAddress || ''),
          addresses: Array.isArray(wh.addresses) ? wh.addresses : []
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



