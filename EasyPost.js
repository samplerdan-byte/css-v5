// ============================================================================
// EasyPost.gs — FedEx/UPS Label Generation via EasyPost API
// Commodity Sampler Services
// Supports: own-account shipping + third-party billing to receiver accounts
// ============================================================================

// ============================================================
// CONFIG
// ============================================================

var EASYPOST_CONFIG = {
  apiUrl: 'https://api.easypost.com/v2',
  testUrl: 'https://api.easypost.com/v2',  // same base, test vs prod is determined by API key
  propKey: 'EASYPOST_API_KEY',
  modePropKey: 'EASYPOST_MODE',  // 'test' or 'production'
  // CSS shipper address — update if office moves
  shipperAddress: {
    company: 'Commodity Sampler Services',
    street1: '',   // ← Dan fills in via setupEasyPost()
    street2: '',
    city: '',
    state: '',
    zip: '',
    country: 'US',
    phone: '',
    email: 'dan@commoditysampler.com'
  }
};


// ============================================================
// SETUP — store API key + shipper address in ScriptProperties
// ============================================================

function setupEasyPost() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  // API Key
  var currentKey = props.getProperty(EASYPOST_CONFIG.propKey) || '';
  var masked = currentKey ? currentKey.slice(0, 8) + '...' + currentKey.slice(-4) : '(not set)';

  var resp = ui.prompt(
    'EasyPost Setup (1/3)',
    'Enter your EasyPost API key.\n\nCurrent: ' + masked + '\n\nLeave blank to keep current.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var newKey = resp.getResponseText().trim();
  if (newKey) props.setProperty(EASYPOST_CONFIG.propKey, newKey);

  // Mode
  var currentMode = props.getProperty(EASYPOST_CONFIG.modePropKey) || 'test';
  var modeResp = ui.prompt(
    'EasyPost Setup (2/3)',
    'Mode: "test" or "production"\n\nCurrent: ' + currentMode + '\n\nTest mode creates fake labels for testing.\nLeave blank to keep current.',
    ui.ButtonSet.OK_CANCEL
  );
  if (modeResp.getSelectedButton() !== ui.Button.OK) return;
  var newMode = modeResp.getResponseText().trim().toLowerCase();
  if (newMode === 'test' || newMode === 'production') {
    props.setProperty(EASYPOST_CONFIG.modePropKey, newMode);
  }

  // Shipper address
  var currentAddr = props.getProperty('EASYPOST_SHIPPER_ADDRESS');
  var addrResp = ui.prompt(
    'EasyPost Setup (3/3)',
    'Enter CSS shipper address as:\nStreet, City, State, Zip, Phone\n\nExample:\n123 Main St, Jersey City, NJ, 07305, 201-555-1234\n\n' +
    (currentAddr ? 'Current: ' + currentAddr : '(not set)') +
    '\n\nLeave blank to keep current.',
    ui.ButtonSet.OK_CANCEL
  );
  if (addrResp.getSelectedButton() !== ui.Button.OK) return;
  var addrText = addrResp.getResponseText().trim();
  if (addrText) {
    props.setProperty('EASYPOST_SHIPPER_ADDRESS', addrText);
  }

  ui.alert('✅ EasyPost Setup Complete!\n\n' +
    'API Key: ' + (newKey ? 'Updated' : 'Unchanged') + '\n' +
    'Mode: ' + (props.getProperty(EASYPOST_CONFIG.modePropKey) || 'test') + '\n' +
    'Address: ' + (addrText || 'Unchanged') + '\n\n' +
    'Run "Test EasyPost Connection" to verify.');
}


// ============================================================
// GET API KEY + SHIPPER ADDRESS
// ============================================================

function _getEasyPostKey() {
  var key = PropertiesService.getScriptProperties().getProperty(EASYPOST_CONFIG.propKey);
  if (!key) throw new Error('EasyPost API key not set. Run setupEasyPost() first.');
  return key;
}

function _getShipperAddress() {
  var addrStr = PropertiesService.getScriptProperties().getProperty('EASYPOST_SHIPPER_ADDRESS');
  if (!addrStr) throw new Error('Shipper address not set. Run setupEasyPost() first.');

  var parts = addrStr.split(',').map(function(p) { return p.trim(); });
  if (parts.length < 4) throw new Error('Shipper address must be: Street, City, State, Zip[, Phone]');

  return {
    company: 'Commodity Sampler Services',
    street1: parts[0],
    city: parts[1],
    state: parts[2],
    zip: parts[3],
    country: 'US',
    phone: parts[4] || '',
    email: 'dan@commoditysampler.com'
  };
}


// ============================================================
// API CALL HELPER
// ============================================================

function _easyPostFetch(endpoint, payload) {
  var apiKey = _getEasyPostKey();
  var url = EASYPOST_CONFIG.apiUrl + endpoint;

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // QUOTA GUARD: Check UrlFetch quota before making API call
  var fetchQuota = (typeof _quotaCheckUrlFetchQuota === 'function') ? _quotaCheckUrlFetchQuota() : { remaining: 20000, countToday: 0, canFetch: true };
  if (!fetchQuota.canFetch) {
    Logger.log('EasyPost: UrlFetch quota exhausted (' + fetchQuota.countToday + '/' + fetchQuota.limit + ' today)');
    throw new Error('EasyPost: Daily UrlFetch quota exhausted. Max ' + fetchQuota.limit + ' calls/day.');
  }

  Logger.log('EasyPost POST ' + endpoint + ' (quota: ' + fetchQuota.remaining + ' remaining)');
  var response;
  try {
    response = fetchWithRetry(url, options, 3, 'EasyPost POST ' + endpoint);
    // QUOTA GUARD: Log the fetch
    if (typeof _quotaLogUrlFetch === 'function') _quotaLogUrlFetch();
  } catch (fetchErr) {
    Logger.log('EasyPost fetch error on POST ' + endpoint + ': ' + fetchErr.message);
    throw new Error('EasyPost network error: ' + fetchErr.message);
  }

  var code = response.getResponseCode();
  var body;
  try {
    body = JSON.parse(response.getContentText());
  } catch (parseErr) {
    Logger.log('EasyPost parse error on POST ' + endpoint + ' (HTTP ' + code + ')');
    throw new Error('EasyPost API error (' + code + '): Invalid response body');
  }

  if (code >= 400) {
    var errMsg = 'EasyPost API error (' + code + ')';
    if (body.error && body.error.message) {
      // Scrub API key from error messages before logging or returning
      var safeErrMsg = String(body.error.message);
      if (apiKey && apiKey.length > 8) {
        safeErrMsg = safeErrMsg.replace(new RegExp(apiKey.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), '[REDACTED]');
      }
      errMsg += ': ' + safeErrMsg;
    }
    Logger.log(errMsg + ' on POST ' + endpoint);
    throw new Error(errMsg);
  }

  return body;
}

function _easyPostGet(endpoint) {
  var apiKey = _getEasyPostKey();
  var url = EASYPOST_CONFIG.apiUrl + endpoint;

  var options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    muteHttpExceptions: true
  };

  // QUOTA GUARD: Check UrlFetch quota before making API call
  var fetchQuota = (typeof _quotaCheckUrlFetchQuota === 'function') ? _quotaCheckUrlFetchQuota() : { remaining: 20000, countToday: 0, canFetch: true };
  if (!fetchQuota.canFetch) {
    Logger.log('EasyPost: UrlFetch quota exhausted (' + fetchQuota.countToday + '/' + fetchQuota.limit + ' today)');
    throw new Error('EasyPost: Daily UrlFetch quota exhausted. Max ' + fetchQuota.limit + ' calls/day.');
  }

  Logger.log('EasyPost GET ' + endpoint + ' (quota: ' + fetchQuota.remaining + ' remaining)');
  var response;
  try {
    response = fetchWithRetry(url, options, 3, 'EasyPost GET ' + endpoint);
    // QUOTA GUARD: Log the fetch
    if (typeof _quotaLogUrlFetch === 'function') _quotaLogUrlFetch();
  } catch (fetchErr) {
    Logger.log('EasyPost fetch error on GET ' + endpoint + ': ' + fetchErr.message);
    throw new Error('EasyPost network error: ' + fetchErr.message);
  }

  var code = response.getResponseCode();
  var body;
  try {
    body = JSON.parse(response.getContentText());
  } catch (parseErr) {
    Logger.log('EasyPost parse error on GET ' + endpoint + ' (HTTP ' + code + ')');
    throw new Error('EasyPost API error (' + code + '): Invalid response body');
  }

  if (code >= 400) {
    var errMsg = 'EasyPost API error (' + code + ')';
    if (body.error && body.error.message) {
      var safeErrMsg = String(body.error.message);
      if (apiKey && apiKey.length > 8) {
        safeErrMsg = safeErrMsg.replace(new RegExp(apiKey.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), '[REDACTED]');
      }
      errMsg += ': ' + safeErrMsg;
    }
    Logger.log(errMsg + ' on GET ' + endpoint);
    throw new Error(errMsg);
  }

  return body;
}


// ============================================================
// TEST CONNECTION
// ============================================================

function testEasyPostConnection() {
  try {
    var shipper = _getShipperAddress();

    // Create a test address to verify API key works
    var result = _easyPostFetch('/addresses', {
      address: {
        company: shipper.company,
        street1: shipper.street1,
        city: shipper.city,
        state: shipper.state,
        zip: shipper.zip,
        country: shipper.country,
        phone: shipper.phone,
        verify: ['delivery']
      }
    });

    var mode = PropertiesService.getScriptProperties().getProperty(EASYPOST_CONFIG.modePropKey) || 'test';
    var verified = result.verifications && result.verifications.delivery && result.verifications.delivery.success;

    SpreadsheetApp.getUi().alert(
      '✅ EasyPost Connection Successful!\n\n' +
      'Mode: ' + mode.toUpperCase() + '\n' +
      'Address ID: ' + result.id + '\n' +
      'Address verified: ' + (verified ? 'Yes ✓' : 'No — check address') + '\n\n' +
      'Shipper: ' + shipper.company + '\n' +
      shipper.street1 + '\n' +
      shipper.city + ', ' + shipper.state + ' ' + shipper.zip
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ EasyPost Connection Failed\n\n' + e.message);
  }
}


// ============================================================
// CREATE SHIPMENT + BUY LABEL
// ============================================================

/**
 * Creates a shipment and buys the cheapest label.
 *
 * @param {Object} toAddress - { company, name, street1, street2, city, state, zip, country, phone }
 * @param {Object} parcel - { length, width, height, weight } (inches/oz)
 * @param {Object} opts - { carrier, thirdPartyAccount, thirdPartyZip, service }
 * @returns {Object} { trackingNumber, labelUrl, carrier, service, rate, shipmentId }
 */
function createShipmentAndBuyLabel(toAddress, parcel, opts) {
  if (!toAddress || typeof toAddress !== 'object') throw new Error('toAddress is required');
  if (!parcel || typeof parcel !== 'object') throw new Error('parcel is required');
  if (!(toAddress.street1 || toAddress.address)) throw new Error('toAddress must include street1 or address');
  if (!(toAddress.zip)) throw new Error('toAddress must include zip');

  opts = opts || {};
  var shipper = _getShipperAddress();
  Logger.log('createShipmentAndBuyLabel: to=' + (toAddress.company || toAddress.name || 'unknown') + ', carrier=' + (opts.carrier || 'any'));

  // Build shipment payload
  var shipmentPayload = {
    shipment: {
      from_address: {
        company: shipper.company,
        street1: shipper.street1,
        street2: shipper.street2 || '',
        city: shipper.city,
        state: shipper.state,
        zip: shipper.zip,
        country: shipper.country,
        phone: shipper.phone,
        email: shipper.email
      },
      to_address: {
        company: toAddress.company || '',
        name: toAddress.name || toAddress.attention || toAddress.company || '',
        street1: toAddress.street1 || toAddress.address || '',
        street2: toAddress.street2 || '',
        city: toAddress.city || '',
        state: toAddress.state || '',
        zip: toAddress.zip || '',
        country: toAddress.country || 'US',
        phone: toAddress.phone || '',
        email: toAddress.email || ''
      },
      parcel: {
        length: parcel.length || 12,
        width: parcel.width || 10,
        height: parcel.height || 6,
        weight: parcel.weight || 32  // 2 lbs default for coffee samples
      },
      options: {}
    }
  };

  // Third-party billing
  if (opts.thirdPartyAccount) {
    shipmentPayload.shipment.options.bill_third_party_account = opts.thirdPartyAccount;
    shipmentPayload.shipment.options.bill_third_party_country = 'US';
    if (opts.thirdPartyZip) {
      shipmentPayload.shipment.options.bill_third_party_postal_code = opts.thirdPartyZip;
    }
  }

  // Create shipment (gets back rates)
  var shipment = _easyPostFetch('/shipments', shipmentPayload);

  if (!shipment.rates || shipment.rates.length === 0) {
    throw new Error('No rates returned. Check addresses and parcel dimensions.');
  }

  // Pick the best rate
  var selectedRate = _selectBestRate(shipment.rates, opts.carrier, opts.service);

  // Buy the label
  var buyResult = _easyPostFetch('/shipments/' + shipment.id + '/buy', {
    rate: { id: selectedRate.id }
  });

  var trackingCode = buyResult.tracking_code || null;
  var rawLabelUrl = (buyResult.postage_label && buyResult.postage_label.label_url) ? buyResult.postage_label.label_url : null;
  // Validate label URL is a real HTTPS URL before returning it
  var labelUrl = (rawLabelUrl && /^https:\/\/.{5,}/.test(rawLabelUrl)) ? rawLabelUrl : null;

  if (!trackingCode) {
    Logger.log('WARNING: EasyPost buy succeeded but tracking_code is missing. Shipment ID: ' + buyResult.id);
  }
  if (!labelUrl && rawLabelUrl) {
    Logger.log('WARNING: EasyPost label URL failed validation: ' + rawLabelUrl);
  }

  Logger.log('Label purchased: tracking=' + (trackingCode || 'MISSING') + ', carrier=' + selectedRate.carrier + ' ' + selectedRate.service + ', rate=$' + selectedRate.rate);

  return {
    trackingNumber: trackingCode,
    labelUrl: labelUrl,
    carrier: selectedRate.carrier,
    service: selectedRate.service,
    rate: selectedRate.rate,
    shipmentId: buyResult.id,
    estimatedDays: selectedRate.est_delivery_days || null
  };
}


/**
 * Select best rate: prefer specified carrier, then cheapest.
 */
function _selectBestRate(rates, preferredCarrier, preferredService) {
  // Filter by carrier if specified
  var filtered = rates;
  if (preferredCarrier) {
    var carrierUpper = preferredCarrier.toUpperCase().replace(/\s+/g, '');
    var carrierFiltered = rates.filter(function(r) {
      var rc = r.carrier.toUpperCase().replace(/\s+/g, '');
      return rc.indexOf(carrierUpper) >= 0 || carrierUpper.indexOf(rc) >= 0;
    });
    if (carrierFiltered.length > 0) filtered = carrierFiltered;
  }

  // Filter by service if specified
  if (preferredService) {
    var svcUpper = preferredService.toUpperCase();
    var svcFiltered = filtered.filter(function(r) {
      return r.service.toUpperCase().indexOf(svcUpper) >= 0;
    });
    if (svcFiltered.length > 0) filtered = svcFiltered;
  }

  // Sort by price, pick cheapest
  // Safety: if filtering left nothing (unexpected carrier/service combo), fall back to full list
  if (filtered.length === 0) {
    Logger.log('_selectBestRate: no rates matched filters (carrier=' + preferredCarrier + ', service=' + preferredService + ') — falling back to cheapest overall');
    filtered = rates;
  }
  filtered.sort(function(a, b) { return parseFloat(a.rate) - parseFloat(b.rate); });
  if (!filtered[0]) throw new Error('No rates available to select from');
  return filtered[0];
}


// ============================================================
// GET RATES (without buying — for rate shopping UI)
// ============================================================

function getShippingRates(toAddress, parcel, opts) {
  if (!toAddress || typeof toAddress !== 'object') throw new Error('toAddress is required');
  if (!parcel || typeof parcel !== 'object') throw new Error('parcel is required');

  opts = opts || {};
  var shipper = _getShipperAddress();
  Logger.log('getShippingRates: to=' + (toAddress.company || toAddress.name || 'unknown'));

  var shipmentPayload = {
    shipment: {
      from_address: {
        company: shipper.company,
        street1: shipper.street1,
        city: shipper.city,
        state: shipper.state,
        zip: shipper.zip,
        country: shipper.country,
        phone: shipper.phone
      },
      to_address: {
        company: toAddress.company || '',
        name: toAddress.name || toAddress.attention || '',
        street1: toAddress.street1 || toAddress.address || '',
        city: toAddress.city || '',
        state: toAddress.state || '',
        zip: toAddress.zip || '',
        country: toAddress.country || 'US',
        phone: toAddress.phone || ''
      },
      parcel: {
        length: parcel.length || 12,
        width: parcel.width || 10,
        height: parcel.height || 6,
        weight: parcel.weight || 32
      },
      options: {}
    }
  };

  if (opts.thirdPartyAccount) {
    shipmentPayload.shipment.options.bill_third_party_account = opts.thirdPartyAccount;
    shipmentPayload.shipment.options.bill_third_party_country = 'US';
  }

  var shipment = _easyPostFetch('/shipments', shipmentPayload);

  if (!shipment.rates || shipment.rates.length === 0) {
    return { rates: [], shipmentId: shipment.id };
  }

  // Sort by price
  var sorted = shipment.rates.sort(function(a, b) {
    return parseFloat(a.rate) - parseFloat(b.rate);
  });

  return {
    shipmentId: shipment.id,
    rates: sorted.map(function(r) {
      return {
        id: r.id,
        carrier: r.carrier,
        service: r.service,
        rate: r.rate,
        currency: r.currency,
        days: r.est_delivery_days,
        deliveryDate: r.delivery_date
      };
    })
  };
}


// ============================================================
// BUY LABEL FROM EXISTING SHIPMENT (after rate shopping)
// ============================================================

function buyLabelForShipment(shipmentId, rateId) {
  if (!shipmentId) throw new Error('shipmentId is required');
  if (!rateId) throw new Error('rateId is required');

  // Validate IDs contain only safe characters (EasyPost IDs are alphanumeric with underscores)
  shipmentId = String(shipmentId);
  rateId = String(rateId);
  if (!/^[a-zA-Z0-9_\-]+$/.test(shipmentId)) throw new Error('Invalid shipment ID format');
  if (!/^[a-zA-Z0-9_\-]+$/.test(rateId)) throw new Error('Invalid rate ID format');

  Logger.log('buyLabelForShipment: shipment=' + shipmentId + ', rate=' + rateId);
  var result = _easyPostFetch('/shipments/' + shipmentId + '/buy', {
    rate: { id: rateId }
  });

  var rawLabelUrl2 = (result.postage_label && result.postage_label.label_url) ? result.postage_label.label_url : null;
  var labelUrl2 = (rawLabelUrl2 && /^https:\/\/.{5,}/.test(rawLabelUrl2)) ? rawLabelUrl2 : null;
  if (!labelUrl2 && rawLabelUrl2) {
    Logger.log('WARNING: buyLabelForShipment label URL failed validation: ' + rawLabelUrl2);
  }
  Logger.log('Label bought: tracking=' + (result.tracking_code || 'none'));
  return {
    trackingNumber: result.tracking_code || null,
    labelUrl: labelUrl2,
    carrier: result.selected_rate ? result.selected_rate.carrier : '',
    service: result.selected_rate ? result.selected_rate.service : '',
    rate: result.selected_rate ? result.selected_rate.rate : '',
    shipmentId: result.id
  };
}


// ============================================================
// TRACK PACKAGE
// ============================================================

function trackPackage(trackingNumber, carrier) {
  if (!trackingNumber || !String(trackingNumber).trim()) return { error: 'Tracking number is required' };
  trackingNumber = String(trackingNumber).trim();

  try {
    Logger.log('trackPackage: ' + trackingNumber + (carrier ? ' (' + carrier + ')' : ''));
    var payload = { tracker: { tracking_code: trackingNumber } };
    if (carrier) payload.tracker.carrier = carrier;

    var result = _easyPostFetch('/trackers', payload);

    return {
      status: result.status,
      statusDetail: result.status_detail,
      carrier: result.carrier,
      estDelivery: result.est_delivery_date,
      trackingDetails: (result.tracking_details || []).map(function(td) {
        return {
          status: td.status,
          message: td.message,
          datetime: td.datetime,
          city: td.tracking_location ? td.tracking_location.city : '',
          state: td.tracking_location ? td.tracking_location.state : ''
        };
      })
    };
  } catch (e) {
    return { error: e.message };
  }
}


// ============================================================
// SHIP FROM SCAN OUT — integrates with existing scan-out flow
// Called from the sidebar when user wants to generate a label
// ============================================================

/**
 * Generate a shipping label for a sample being scanned out.
 * Looks up receiver address + carrier account from Contacts.
 *
 * @param {string} sampleId - CS Sample # to ship
 * @param {Object} parcelOverrides - optional { length, width, height, weight }
 * @returns {Object} { success, trackingNumber, labelUrl, carrier, service, rate, error }
 */
function generateLabelForSample(sampleId, parcelOverrides) {
  if (!sampleId || !String(sampleId).trim()) return { success: false, error: 'Sample ID is required' };
  sampleId = String(sampleId).trim();

  try {
    Logger.log('generateLabelForSample: ' + sampleId);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.allOrders);
    if (!sheet) return { success: false, error: 'All Orders sheet not found' };

    // Find the sample row
    var col = _getColumnMap(sheet);
    var row = _findRowBySample(sheet, sampleId);
    if (row < 2) return { success: false, error: sampleId + ' not found' };

    // Get receiver name
    var receiverIdx = col['Receiver'];
    if (receiverIdx === undefined) return { success: false, error: 'Receiver column not found' };
    var receiverName = String(sheet.getRange(row, receiverIdx + 1).getValue()).trim();
    if (!receiverName) return { success: false, error: 'No receiver for ' + sampleId };

    // Look up receiver address + carrier accounts from Contacts
    var contactInfo = _getReceiverContact(receiverName);
    if (!contactInfo) return { success: false, error: 'Receiver "' + receiverName + '" not found in Contacts' };
    if (!contactInfo.address || !contactInfo.city || !contactInfo.state || !contactInfo.zip) {
      return { success: false, error: 'Incomplete address for "' + receiverName + '". Update Contacts sheet.' };
    }

    // Determine carrier + third-party billing
    var carrier = lookupCarrierByCompany(receiverName);
    var preferredCarrier = (carrier && carrier.preferred) || 'FedEx';
    var thirdPartyAcct = '';
    if (carrier) {
      if (preferredCarrier.toUpperCase().indexOf('UPS') >= 0 && carrier.ups) {
        thirdPartyAcct = carrier.ups;
      } else if (carrier.fedex) {
        thirdPartyAcct = carrier.fedex;
      }
    }

    // Build destination address
    var toAddress = {
      company: contactInfo.company,
      name: contactInfo.attention || contactInfo.company,
      street1: contactInfo.address,
      city: contactInfo.city,
      state: contactInfo.state,
      zip: contactInfo.zip,
      country: 'US',
      phone: contactInfo.phone || '',
      email: contactInfo.email || ''
    };

    // Default parcel (coffee sample box)
    var parcel = {
      length: 12, width: 10, height: 6, weight: 32  // ~2 lbs
    };
    if (parcelOverrides) {
      if (parcelOverrides.length) parcel.length = parcelOverrides.length;
      if (parcelOverrides.width) parcel.width = parcelOverrides.width;
      if (parcelOverrides.height) parcel.height = parcelOverrides.height;
      if (parcelOverrides.weight) parcel.weight = parcelOverrides.weight;
    }

    // Create shipment and buy label
    var result = createShipmentAndBuyLabel(toAddress, parcel, {
      carrier: preferredCarrier,
      thirdPartyAccount: thirdPartyAcct,
      thirdPartyZip: contactInfo.zip
    });

    Logger.log('generateLabelForSample success: ' + sampleId + ' → ' + result.trackingNumber);
    return {
      success: true,
      trackingNumber: result.trackingNumber,
      labelUrl: result.labelUrl,
      carrier: result.carrier,
      service: result.service,
      rate: result.rate,
      estimatedDays: result.estimatedDays,
      billedTo: thirdPartyAcct ? receiverName + ' (' + thirdPartyAcct + ')' : 'CSS Account'
    };

  } catch (e) {
    Logger.log('generateLabelForSample error: ' + sampleId + ' — ' + e.message);
    return { success: false, error: e.message };
  }
}


/**
 * Look up receiver's full contact info from Contacts sheet.
 */
function _getReceiverContact(receiverName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet || sheet.getLastRow() < 2) return null;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  var companyCol = 1;
  var attentionCol = 2;
  var addressCol = 3;
  var cityCol = 4;
  var stateCol = 5;
  var zipCol = 6;
  var phoneCol = 7;
  var emailCol = 8;

  var searchName = String(receiverName).trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var company = String(data[i][companyCol]).trim().toLowerCase();
    if (company === searchName ||
        company.indexOf(searchName) >= 0 ||
        searchName.indexOf(company) >= 0) {
      return {
        company: String(data[i][companyCol]).trim(),
        attention: String(data[i][attentionCol] || '').trim(),
        address: String(data[i][addressCol] || '').trim(),
        city: String(data[i][cityCol] || '').trim(),
        state: String(data[i][stateCol] || '').trim(),
        zip: String(data[i][zipCol] || '').trim(),
        phone: String(data[i][phoneCol] || '').trim(),
        email: String(data[i][emailCol] || '').trim()
      };
    }
  }

  // Try word-level match
  var words = searchName.split(/[\s\/&,\-]+/).filter(function(w) { return w.length > 3; });
  for (var w = 0; w < words.length; w++) {
    for (var j = 0; j < data.length; j++) {
      var comp = String(data[j][companyCol]).trim().toLowerCase();
      if (comp.indexOf(words[w]) >= 0) {
        return {
          company: String(data[j][companyCol]).trim(),
          attention: String(data[j][attentionCol] || '').trim(),
          address: String(data[j][addressCol] || '').trim(),
          city: String(data[j][cityCol] || '').trim(),
          state: String(data[j][stateCol] || '').trim(),
          zip: String(data[j][zipCol] || '').trim(),
          phone: String(data[j][phoneCol] || '').trim(),
          email: String(data[j][emailCol] || '').trim()
        };
      }
    }
  }

  return null;
}


// ============================================================
// BATCH LABEL GENERATION
// Generate labels for all samples in a list
// ============================================================

function generateLabelsForBatch(sampleIds, parcelOverrides) {
  if (!Array.isArray(sampleIds) || sampleIds.length === 0) return [];
  Logger.log('generateLabelsForBatch: ' + sampleIds.length + ' sample(s)');

  var results = [];
  for (var i = 0; i < sampleIds.length; i++) {
    var result = generateLabelForSample(sampleIds[i], parcelOverrides);
    results.push({
      sampleId: sampleIds[i],
      success: result.success,
      trackingNumber: result.trackingNumber || '',
      labelUrl: result.labelUrl || '',
      carrier: result.carrier || '',
      service: result.service || '',
      rate: result.rate || '',
      billedTo: result.billedTo || '',
      error: result.error || ''
    });
    // Small delay to avoid rate limiting
    if (i < sampleIds.length - 1) Utilities.sleep(500);
  }
  return results;
}


// ============================================================
// VOID / REFUND SHIPMENT
// ============================================================

function voidShipment(shipmentId) {
  if (!shipmentId) return { success: false, error: 'shipmentId is required' };

  // Validate shipmentId format to prevent path injection
  shipmentId = String(shipmentId);
  if (!/^[a-zA-Z0-9_\-]+$/.test(shipmentId)) return { success: false, error: 'Invalid shipment ID format' };

  try {
    Logger.log('voidShipment: ' + shipmentId);
    var result = _easyPostFetch('/shipments/' + shipmentId + '/refund', {});
    Logger.log('voidShipment success: ' + shipmentId + ' → ' + result.refund_status);
    return { success: true, status: result.refund_status };
  } catch (e) {
    Logger.log('voidShipment error: ' + shipmentId + ' — ' + e.message);
    return { success: false, error: e.message };
  }
}


// ============================================================
// LABEL GENERATOR UI — modal to generate a label for selected sample
// ============================================================

function showLabelGenerator() {
  var html = '<html><head><style>' +
    'body { font-family: Arial, sans-serif; padding: 16px; background: #f8f9fa; }' +
    'h2 { margin: 0 0 12px; font-size: 18px; color: #2d5016; }' +
    'label { display: block; font-size: 11px; color: #555; font-weight: bold; margin: 8px 0 3px; }' +
    'input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; box-sizing: border-box; }' +
    'input:focus { border-color: #2d5016; outline: none; }' +
    '.parcel-row { display: flex; gap: 6px; }' +
    '.parcel-row > div { flex: 1; }' +
    '.btn { padding: 10px 20px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; width: 100%; margin-top: 12px; }' +
    '.btn-gen { background: #2d5016; color: #fff; }' +
    '.btn-gen:hover { background: #3d6b1e; }' +
    '.btn-gen:disabled { opacity: 0.5; cursor: not-allowed; }' +
    '#result { margin-top: 12px; padding: 12px; border-radius: 8px; display: none; font-size: 13px; }' +
    '.result-ok { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }' +
    '.result-err { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }' +
    '.result-row { display: flex; justify-content: space-between; padding: 2px 0; }' +
    '.result-label { font-weight: bold; }' +
    'a { color: #2d5016; font-weight: bold; }' +
    '#status { font-size: 12px; color: #666; text-align: center; margin-top: 8px; min-height: 16px; }' +
    '</style></head><body>' +
    '<h2>🏷️ Generate Shipping Label</h2>' +
    '<label>CS Sample #</label>' +
    '<input type="text" id="sampleId" placeholder="e.g. CS-2025-001" autofocus>' +
    '<label>Parcel Dimensions (optional — defaults: 12x10x6 in, 2 lbs)</label>' +
    '<div class="parcel-row">' +
    '<div><label>L (in)</label><input type="number" id="pLen" placeholder="12"></div>' +
    '<div><label>W (in)</label><input type="number" id="pWid" placeholder="10"></div>' +
    '<div><label>H (in)</label><input type="number" id="pHt" placeholder="6"></div>' +
    '<div><label>Wt (oz)</label><input type="number" id="pWt" placeholder="32"></div>' +
    '</div>' +
    '<button class="btn btn-gen" id="genBtn" onclick="generate()">Generate Label</button>' +
    '<div id="status"></div>' +
    '<div id="result"></div>' +
    '<script>' +
    'function esc(s){var d=document.createElement("div");d.textContent=s||"";return d.innerHTML;}' +
    'function generate() {' +
    '  var sid = document.getElementById("sampleId").value.trim();' +
    '  if (!sid) { alert("Enter a sample ID"); return; }' +
    '  document.getElementById("genBtn").disabled = true;' +
    '  document.getElementById("status").textContent = "Generating label...";' +
    '  document.getElementById("result").style.display = "none";' +
    '  var parcel = {};' +
    '  var l = document.getElementById("pLen").value; if (l) parcel.length = parseFloat(l);' +
    '  var w = document.getElementById("pWid").value; if (w) parcel.width = parseFloat(w);' +
    '  var h = document.getElementById("pHt").value; if (h) parcel.height = parseFloat(h);' +
    '  var wt = document.getElementById("pWt").value; if (wt) parcel.weight = parseFloat(wt);' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r) {' +
    '      document.getElementById("genBtn").disabled = false;' +
    '      document.getElementById("status").textContent = "";' +
    '      var el = document.getElementById("result");' +
    '      if (r.success) {' +
    '        el.className = "result-ok";' +
    '        var h = \'<div class="result-row"><span class="result-label">Tracking:</span><span>\' + esc(r.trackingNumber) + \'</span></div>\';' +
    '        h += \'<div class="result-row"><span class="result-label">Carrier:</span><span>\' + esc(r.carrier) + \' \' + esc(r.service) + \'</span></div>\';' +
    '        h += \'<div class="result-row"><span class="result-label">Rate:</span><span>$\' + esc(r.rate) + \'</span></div>\';' +
    '        h += \'<div class="result-row"><span class="result-label">Billed To:</span><span>\' + esc(r.billedTo) + \'</span></div>\';' +
    '        if (r.labelUrl) {' +
    '          var a = document.createElement("a"); a.href = r.labelUrl; a.target = "_blank"; a.textContent = "📄 Open Label PDF";' +
    '          h += \'<div style="margin-top:8px;text-align:center" id="labelLink"></div>\';' +
    '          el.innerHTML = h;' +
    '          document.getElementById("labelLink").appendChild(a);' +
    '        } else { el.innerHTML = h; }' +
    '      } else {' +
    '        el.className = "result-err";' +
    '        el.textContent = r.error;' +
    '      }' +
    '      el.style.display = "block";' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      document.getElementById("genBtn").disabled = false;' +
    '      document.getElementById("status").textContent = "";' +
    '      var el = document.getElementById("result");' +
    '      el.className = "result-err";' +
    '      el.textContent = e.message;' +
    '      el.style.display = "block";' +
    '    })' +
    '    .generateLabelForSample(sid, Object.keys(parcel).length > 0 ? parcel : null);' +
    '}' +
    '</script></body></html>';

  var dialog = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(420)
    .setTitle('Generate Shipping Label');
  SpreadsheetApp.getUi().showModalDialog(dialog, '🏷️ Generate Shipping Label');
}


// ============================================================
// TRACK PACKAGE UI
// ============================================================

function showTrackPackage() {
  var html = '<html><head><style>' +
    'body { font-family: Arial, sans-serif; padding: 16px; background: #f8f9fa; }' +
    'h2 { margin: 0 0 12px; font-size: 18px; color: #2d5016; }' +
    'label { display: block; font-size: 11px; color: #555; font-weight: bold; margin: 8px 0 3px; }' +
    'input { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px; box-sizing: border-box; font-family: monospace; }' +
    'input:focus { border-color: #2d5016; outline: none; }' +
    '.btn { padding: 10px 20px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; width: 100%; margin-top: 10px; background: #2d5016; color: #fff; }' +
    '.btn:hover { background: #3d6b1e; }' +
    '.btn:disabled { opacity: 0.5; }' +
    '#result { margin-top: 12px; display: none; }' +
    '.status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 13px; font-weight: bold; margin-bottom: 8px; }' +
    '.status-delivered { background: #d4edda; color: #155724; }' +
    '.status-transit { background: #cce5ff; color: #004085; }' +
    '.status-other { background: #fff3cd; color: #856404; }' +
    '.timeline { list-style: none; padding: 0; margin: 8px 0 0; }' +
    '.timeline li { padding: 6px 0 6px 16px; border-left: 2px solid #ddd; font-size: 11px; position: relative; }' +
    '.timeline li:before { content: ""; width: 8px; height: 8px; border-radius: 50%; background: #2d5016; position: absolute; left: -5px; top: 10px; }' +
    '.tl-date { color: #999; font-size: 10px; }' +
    '.tl-loc { color: #666; }' +
    '#status { font-size: 12px; color: #666; text-align: center; margin-top: 8px; }' +
    '</style></head><body>' +
    '<h2>🔍 Track Package</h2>' +
    '<label>Tracking Number</label>' +
    '<input type="text" id="trackNum" placeholder="Enter tracking number" autofocus>' +
    '<button class="btn" id="trackBtn" onclick="doTrack()">Track</button>' +
    '<div id="status"></div>' +
    '<div id="result"></div>' +
    '<script>' +
    'function esc(s){var d=document.createElement("div");d.textContent=s||"";return d.innerHTML;}' +
    'function doTrack() {' +
    '  var num = document.getElementById("trackNum").value.trim();' +
    '  if (!num) { alert("Enter a tracking number"); return; }' +
    '  document.getElementById("trackBtn").disabled = true;' +
    '  document.getElementById("status").textContent = "Tracking...";' +
    '  document.getElementById("result").style.display = "none";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r) {' +
    '      document.getElementById("trackBtn").disabled = false;' +
    '      document.getElementById("status").textContent = "";' +
    '      var el = document.getElementById("result");' +
    '      if (r.error) { el.textContent = r.error; el.style.color = "#c62828"; el.style.display = "block"; return; }' +
    '      var statusCls = r.status === "delivered" ? "status-delivered" : (r.status === "in_transit" ? "status-transit" : "status-other");' +
    '      var h = \'<span class="status-badge \' + esc(statusCls) + \'">\' + esc((r.status || "unknown").toUpperCase()) + \'</span>\';' +
    '      h += \'<div style="font-size:12px;color:#666;margin-bottom:6px">Carrier: \' + esc(r.carrier || "?") + \'</div>\';' +
    '      if (r.estDelivery) h += \'<div style="font-size:12px;color:#666">Est. Delivery: \' + esc(r.estDelivery) + \'</div>\';' +
    '      if (r.trackingDetails && r.trackingDetails.length > 0) {' +
    '        h += \'<ul class="timeline">\';' +
    '        for (var i = 0; i < Math.min(r.trackingDetails.length, 15); i++) {' +
    '          var td = r.trackingDetails[i];' +
    '          h += \'<li>\' + esc(td.message);' +
    '          if (td.city || td.state) h += \' <span class="tl-loc">(\' + esc([td.city, td.state].filter(Boolean).join(", ")) + \')</span>\';' +
    '          if (td.datetime) h += \'<br><span class="tl-date">\' + esc(new Date(td.datetime).toLocaleString()) + \'</span>\';' +
    '          h += \'</li>\';' +
    '        }' +
    '        h += \'</ul>\';' +
    '      }' +
    '      el.innerHTML = h;' +
    '      el.style.display = "block";' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      document.getElementById("trackBtn").disabled = false;' +
    '      document.getElementById("status").textContent = "";' +
    '      var el = document.getElementById("result");' +
    '      el.textContent = e.message;' +
    '      el.style.color = "#c62828";' +
    '      el.style.display = "block";' +
    '    })' +
    '    .trackPackage(num);' +
    '}' +
    '</script></body></html>';

  var dialog = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(500)
    .setTitle('Track Package');
  SpreadsheetApp.getUi().showModalDialog(dialog, '🔍 Track Package');
}
