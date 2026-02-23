// ============================================================================
// EmailExtraction_v3.gs
// Commodity Sampler Services — Email/PDF Extraction + Gmail Processing
// ============================================================================


// ============================================================================
// SECTION 1: CONFIGURATION — EXCLUDED SENDERS
// ============================================================================

var EXCLUDED_DOMAINS = [
  '@commoditysampler.com',
  '@ddscocoa.com'
];

function isExcludedSender(senderEmail) {
  senderEmail = (senderEmail || '').toLowerCase();
  for (var i = 0; i < EXCLUDED_DOMAINS.length; i++) {
    if (senderEmail.indexOf(EXCLUDED_DOMAINS[i]) >= 0) return true;
  }
  return false;
}


// ============================================================================
// SECTION 2: CONFIGURATION — KNOWN CLIENTS
// ============================================================================

var KNOWN_CLIENTS = {
  'olam': {
    name: 'Olam Americas',
    domains: ['@ofi.com', '@olamamericas.com'],
    keywords: ['Olam Americas', 'ofi.com'],
    defaultShipping: 'FedEx Priority',
    defaultSampleSize: '2 lb',
    warehouse: 'RPM_AVENEL'
  },
  'coffeeamerica': {
    name: 'Coffee America USA Corp',
    domains: ['@coffeeamericausa.com'],
    keywords: ['Coffee America', 'Pier to Whse', 'Delivery Order'],
    defaultShipping: 'FedEx Priority',
    defaultSampleSize: '2 lb',
    fedexAccount: '6701-4695-2',
    warehouse: 'RPM_AVENEL',
    useCustomParser: true
  },
  'ally': {
    name: 'Ally Coffee Trading S.A.',
    domains: ['@allycoffee.com'],
    keywords: ['Ally Coffee', 'ALLY SAMPLING ORDER', 'Ally Brazilian'],
    defaultShipping: 'FedEx Standard Overnight',
    defaultSampleSize: '2 lb',
    warehouse: 'RPM_AVENEL',
    useCustomParser: true
  },
  'rothfos': {
    name: 'Rothfos (Neumann Kaffee Gruppe)',
    domains: ['@nkg.coffee', '@rothfos.com'],
    keywords: ['Rothfos', 'Neumann Gruppe USA Inc. dba Rothfos'],
    defaultShipping: 'FedEx 2 Day',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200',
    useCustomParser: true
  },
  'interamerican': {
    name: 'InterAmerican Coffee (NKG)',
    domains: ['@nkg.coffee'],
    keywords: ['InterAmerican Coffee', 'Neumann Gruppe USA Inc. dba InterAmerican'],
    defaultShipping: 'FedEx 2 Day',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200',
    useCustomParser: true
  },
  'atlantic_usa': {
    name: 'Atlantic (USA), LLC',
    domains: ['@ecomtrading.com'],
    keywords: ['Atlantic (USA), LLC', 'LETTER OF ENTRY', 'LOE'],
    defaultShipping: 'UPS Ground',
    defaultSampleSize: '2 lb',
    warehouse: 'RPM_AVENEL',
    requiresSamplingInstructions: true,
    useCustomParser: true
  },
  'atlantic_specialty': {
    name: 'Atlantic Specialty Coffee, Inc.',
    domains: ['@ecomtrading.com'],
    keywords: ['Atlantic Specialty Coffee', 'ASCILAB', 'ascitraffic'],
    defaultShipping: 'UPS 2 Day',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200'
  },
  'excelco': {
    name: 'Excelco Trading LP',
    domains: ['@exceltrade.com', '@excelco.onmicrosoft.com'],
    keywords: ['Excelco', 'Sample Allowance', 'Sopex'],
    defaultShipping: 'UPS 2nd Day',
    defaultSampleSize: '2 lb',
    warehouse: 'RPM_AVENEL',
    isBroker: true,
    brokeredClients: ['equalexchange', 'linglebrothers']
  },
  'equalexchange': {
    name: 'Equal Exchange',
    domains: [],
    keywords: ['Equal Exchange', 'EQUAL EXCHANGE'],
    broker: 'excelco',
    defaultShipping: 'UPS 2nd Day',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200'
  },
  'linglebrothers': {
    name: 'Lingle Brothers Coffee',
    domains: [],
    keywords: ['Lingle Brothers', 'LINGLE BROTHERS'],
    defaultShipping: 'FedEx',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200'
  },
  'covoya': {
    name: 'Covoya Specialty Coffee',
    domains: ['@covoya.com'],
    keywords: ['Covoya', 'ARRIVAL SAMPLE REQUEST'],
    defaultShipping: 'FedEx',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200'
  },
  'osito': {
    name: 'Osito Coffee',
    domains: ['@ositocoffee.com'],
    keywords: ['Osito Coffee'],
    defaultShipping: 'FedEx Ground',
    defaultSampleSize: '500g',
    fedexAccount: '625178762',
    warehouse: 'CONTINENTAL_200',
    useCustomParser: true
  },
  'paragon': {
    name: 'Paragon Coffee Trading Company L.P.',
    domains: ['@paragoncoffee.com'],
    keywords: ['Paragon Coffee', 'PAR_KS_SAMPLING_ORDER', 'PARAGON COFFEE'],
    defaultShipping: 'FedEx',
    defaultSampleSize: '2 lb',
    sampleSizeToSelf: '5 lb',
    warehouse: null
  },
  'coffeesource': {
    name: 'The Coffee Source LLC',
    domains: ['@coffee.cr', '@coffeecr.onmicrosoft.com'],
    keywords: ['The Coffee Source', 'coffee.cr'],
    defaultShipping: 'FedEx',
    defaultSampleSize: '2 lb',
    fedexAccount: '9462-6555-1',
    warehouse: 'CONTINENTAL_200',
    useCustomParser: true
  },
  'listbeisler': {
    name: 'List + Beisler Corp',
    domains: ['@list-beisler.us'],
    keywords: ['List + Beisler', 'L+B Corp'],
    defaultShipping: 'FedEx 2 Day',
    defaultSampleSize: '2 lb',
    fedexAccount: '763817172',
    warehouse: 'CONTINENTAL_200'
  },
  'sucafina': {
    name: 'Sucafina NA',
    domains: ['@sucafina.com', '@tastify.com'],
    keywords: ['Sucafina', 'SUCAFINA', 'Tastify'],
    defaultShipping: 'FedEx',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200',
    useCustomParser: true
  },
  'serengeti': {
    name: 'Serengeti Trading Company',
    domains: ['@serengetitrading.com'],
    keywords: ['Serengeti Trading', 'SERENGETI', 'Serengeti -'],
    defaultShipping: 'FedEx 2 Day',
    defaultSampleSize: '4 lb',
    warehouse: 'CONTINENTAL_200',
    defaultReceiver: {
      name: 'Serengeti Trading Company',
      address: '19100 Hamilton Pool Rd',
      city: 'Dripping Springs',
      state: 'TX',
      zip: '78620'
    },
    useCustomParser: true
  },
  'icc': {
    name: 'International Coffee Corporation',
    domains: ['@iccnola.com'],
    keywords: ['International Coffee Corporation', 'ICC'],
    broker: 'ict',
    defaultShipping: 'FedEx Standard Overnight',
    defaultSampleSize: '2 lb',
    fedexAccount: '461992226',
    warehouse: 'RPM_AVENEL',
    defaultReceiver: {
      name: 'International Coffee Corporation',
      address: '734 Martin Behrman Ave',
      city: 'Metairie',
      state: 'LA',
      zip: '70005'
    },
    useCustomParser: true
  },
  'ict': {
    name: 'Intercontinental Coffee Trading',
    domains: ['@ictcoffee.com'],
    keywords: ['Intercontinental Coffee Trading', 'ICT', 'SAMPLE ORDER', 'Sample Allowance'],
    defaultShipping: 'FedEx 2 Day',
    defaultSampleSize: '2 lb',
    fedexAccount: '461992226',
    warehouse: 'RPM_AVENEL',
    isBroker: true,
    brokeredClients: ['icc']
  },
  'amcof': {
    name: 'American Coffee Corporation',
    domains: ['@amcof.com'],
    keywords: ['American Coffee Corp', 'American Coffee Corporation', 'AMCOF'],
    defaultShipping: 'FedEx',
    defaultSampleSize: '2 lb',
    warehouse: 'RPM_AVENEL',
    useCustomParser: true
  },
  'armenia': {
    name: 'Armenia Coffee Corp',
    domains: ['@armeniacoffee.com'],
    keywords: ['Armenia Coffee', 'Armenia reference'],
    defaultShipping: 'UPS Ground',
    defaultSampleSize: '2 lb',
    warehouse: 'CONTINENTAL_200',
    useCustomParser: true
  },
  'ldc': {
    name: 'Louis Dreyfus Company',
    domains: ['@ldc.com'],
    keywords: ['Louis Dreyfus', 'LDC'],
    defaultShipping: 'FedEx Standard Overnight',
    defaultSampleSize: '2 lb',
    warehouse: 'RPM_AVENEL',
    useCustomParser: true
  }
};


// ============================================================================
// SECTION 3: CONFIGURATION — ICO COUNTRY CODES
// ============================================================================

var ICO_COUNTRIES = {
  '001': 'Bolivia',   '01': 'Bolivia',   '1': 'Bolivia',
  '002': 'Brazil',    '02': 'Brazil',    '2': 'Brazil',
  '003': 'Colombia',  '03': 'Colombia',  '3': 'Colombia',
  '005': 'Costa Rica','05': 'Costa Rica', '5': 'Costa Rica',
  '006': 'Cuba',      '06': 'Cuba',      '6': 'Cuba',
  '007': 'Dominican Republic', '07': 'Dominican Republic', '7': 'Dominican Republic',
  '008': 'Ecuador',   '08': 'Ecuador',   '8': 'Ecuador',
  '009': 'El Salvador','09': 'El Salvador','9': 'El Salvador',
  '010': 'Ethiopia',  '10': 'Ethiopia',
  '011': 'Guatemala', '11': 'Guatemala',
  '012': 'Haiti',     '12': 'Haiti',
  '013': 'Honduras',  '13': 'Honduras',
  '014': 'India',     '14': 'India',
  '015': 'Indonesia', '15': 'Indonesia',
  '016': 'Mexico',    '16': 'Mexico',
  '017': 'Nicaragua', '17': 'Nicaragua',
  '018': 'Panama',    '18': 'Panama',
  '019': 'Papua New Guinea', '19': 'Papua New Guinea',
  '020': 'Paraguay',  '20': 'Paraguay',
  '021': 'Peru',      '21': 'Peru',
  '023': 'Mexico',    '23': 'Mexico',
  '027': 'Rwanda',    '27': 'Rwanda',
  '030': 'Peru',      '30': 'Peru',
  '033': 'Tanzania',  '33': 'Tanzania',
  '035': 'Uganda',    '35': 'Uganda',
  '036': 'Venezuela', '36': 'Venezuela',
  '0145': 'Vietnam',  '145': 'Vietnam'
};

function getOriginFromMark(mark) {
  if (!mark) return '';
  var parts = String(mark).split('/');
  if (parts.length < 2) return '';
  return ICO_COUNTRIES[parts[0]] || '';
}

function getCountryFromICOMark(mark) {
  return getOriginFromMark(mark);
}


// ============================================================================
// SECTION 4: CONFIGURATION — WAREHOUSES
// ============================================================================

var WAREHOUSES = {
  'RPM_AVENEL': {
    name: 'RPM AVENEL',
    region: 'CSS NJ',
    addresses: ['1500 Rahway', 'Avenel', 'RPM Avenel'],
    fullAddress: '1500 Rahway Avenue, Avenel, NJ 07001'
  },
  'RPM_WOODBRIDGE': {
    name: 'RPM WOODBRIDGE',
    region: 'CSS NJ',
    addresses: ['2900 Woodbridge', 'Woodbridge'],
    fullAddress: '2900 Woodbridge Avenue, Edison, NJ 08837'
  },
  'RPM_TALMAGE': {
    name: 'RPM TALMAGE',
    region: 'CSS NJ',
    addresses: ['Talmage'],
    fullAddress: '44 Talmadge Road, Edison, NJ 08817'
  },
  'RPM_UNASSIGNED': {
    name: 'RPM UNASSIGNED',
    region: 'CSS NJ',
    addresses: [],
    fullAddress: ''
  },
  'CONTINENTAL_200': {
    name: 'CONTINENTAL 200',
    region: 'CSS NJ',
    addresses: ['200 Middlesex', 'Carteret', 'Continental', 'CTI'],
    fullAddress: '200 Middlesex Avenue, Carteret, NJ 07008'
  },
  'CONTINENTAL_300': {
    name: 'CONTINENTAL 300',
    region: 'CSS NJ',
    addresses: ['300 Middlesex', '300 Mac Lane', 'Keasby'],
    fullAddress: '300 Mac Lane, Keasby, NJ 08832'
  },
  'Florence': {
    name: 'Florence Warehouse Inc',
    region: 'CSS NJ',
    addresses: ['Florence Warehouse', 'Florence', 'Mecca', '580 Marin'],
    fullAddress: '580 Marin Blvd, Jersey City, NJ 07310'
  },
  'Cadeco': {
    name: 'Cadeco',
    region: 'CSS HOUSTON',
    addresses: ['Cadeco', '5610 Clinton'],
    fullAddress: '5610 Clinton Dr, Houston, TX 77020'
  },
  'Dupuy': {
    name: 'Dupuy Storage',
    region: 'CSS HOUSTON',
    addresses: ['Dupuy Storage', 'Dupuy', '7703 Cannon'],
    fullAddress: '7703 Cannon St, Houston, TX 77021'
  },
  'RPM_NORTHDALE': {
    name: 'RPM NORTHDALE',
    region: 'CSS HOUSTON',
    addresses: ['RPM Northdale', 'Northdale', '6462 Northdale'],
    fullAddress: '6462 Northdale St, Houston, TX 77087'
  }
};


// ============================================================================
// SECTION 5: CONFIGURATION — SAMPLING INDICATORS
// ============================================================================

var SAMPLING_INDICATORS = [
  'please send sample',
  'send sample',
  'pull sample',
  'please sample',
  'sampling order',
  'sample order',
  'sample request',
  'sample allowance',
  'sampler to sample',
  'commodity sampler',
  'sampling procedure',
  'allow sampling',
  'offering sample',
  'sample size',
  'sample qty',
  'draw and send',
  '\\d+\\s*(lb|lbs|g|gram)\\s+sample',
  'standard\\s+\\d+#'
];


// ============================================================================
// SECTION 6: GMAIL PROCESSING PIPELINE
// ============================================================================

function processPDFsFromGmail() {
  var t0 = new Date();
  Logger.log('=== STARTING PROCESS (OPTIMIZED) ===');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.mainSheetName);
  Logger.log('Sheet found: ' + (sheet ? 'YES' : 'NO'));

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Please run setupSheet() first!');
    return;
  }

  var col = _getColumnMap(sheet);
  var attachColIdx = (col['Attachments'] !== undefined) ? col['Attachments'] : -1;
  var photoColIdx = (col['Photos'] !== undefined) ? col['Photos'] : -1;
  var sampleColIdx = (col['CS Sample #'] !== undefined) ? col['CS Sample #'] : -1;

  var orderCache = _buildOrderCache(sheet);
  Logger.log('Order cache built: highest=' + orderCache.highestOrderNumber + ', orders=' + Object.keys(orderCache.orders).length);

  var tSetup = new Date();
  Logger.log('Setup time: ' + (tSetup - t0) + 'ms');

  var searchQuery = buildSearchQuery();
  Logger.log('Search query: ' + searchQuery);
  var threads = GmailApp.search(searchQuery, 0, 50);
  Logger.log('Found threads: ' + threads.length);

  var tSearch = new Date();
  Logger.log('Gmail search time: ' + (tSearch - tSetup) + 'ms');

  var processedCount = 0;
  var pdfTime = 0;
  var extractTime = 0;
  var writeTime = 0;
  var attachTime = 0;

  threads.forEach(function(thread, threadIndex) {
    var tThread = new Date();
    Logger.log('=== Processing thread ' + (threadIndex + 1) + '/' + threads.length + ' ===');
    var messages = thread.getMessages();
    var message = messages[0];

    var emailBody = message.getPlainBody();
    var subject = message.getSubject();
    var senderEmail = message.getFrom();
    Logger.log('Subject: ' + subject);
    Logger.log('From: ' + senderEmail);

    if (isExcludedSender(senderEmail)) {
      Logger.log('Excluded sender - skipping');
      thread.addLabel(getOrCreateLabel('PDF_Processed'));
      return;
    }

    var emailLink = (typeof getGmailMessageLink === 'function') ? getGmailMessageLink(message) : '';
    var attachments = message.getAttachments();
    Logger.log('Attachments: ' + attachments.length);

    // --- Extract text from PDF attachments ---
    var tPdf = new Date();
    var pdfText = '';
    attachments.forEach(function(attachment) {
      if (attachment.getContentType() === 'application/pdf') {
        try {
          var text = extractTextFromPDF(attachment);
          if (text && text.length > 0) pdfText += text + '\n\n';
        } catch (e) {
          if (typeof logError === 'function') {
            logError('extractTextFromPDF', e.toString(), { attachment: attachment.getName() });
          }
          Logger.log('Error extracting PDF text: ' + e);
        }
      }
    });
    pdfTime += (new Date() - tPdf);

    // --- Extract order data ---
    var tExtract = new Date();
    var extractedRecords = [];

    try {
      var result = extractOrderFromEmail(emailBody, pdfText, senderEmail, subject);

      if (result.success && result.orders.length > 0) {
        Logger.log('Extracted ' + result.orders.length + ' orders');
        if (result.flags.length > 0) Logger.log('Flags: ' + result.flags.join('; '));

        var senderName = getOriginalSender(message, emailBody);
        var groupKey = senderName + '|' + message.getId().substring(0, 8);

        extractedRecords = result.orders.map(function(order) {
          var displayOrderNum = order.sampleOrderNum || senderName;
          var orderGroupKey = order.sampleOrderNum
            ? (order.sampleOrderNum + '|' + message.getId().substring(0, 8))
            : groupKey;

          return {
            sampleOrderNum: displayOrderNum,
            _orderGroupKey: orderGroupKey,
            cargo: order.cargo || '',
            mark: order.mark || '',
            container: order.container || '',
            reference: order.reference || '',
            description: order.description || '',
            bagCount: order.bags || '',
            weight: '',
            sampleWeight: order.sampleSize || '5 LB',
            pNumber: order.reference || '',
            sNumber: '',
            warehouse: order.warehouse || '',
            shippingProcess: order.shipping || '',
            receiver: (order.receiver && typeof order.receiver === 'object')
              ? (order.receiver.name || order.receiver.address || '')
              : (order.receiver || ''),
            sender: order.client || getOriginalSender(message, emailBody),
            comments: (result.flags && result.flags.length > 0) ? result.flags.join('; ') : '',
            sourceEmail: senderEmail,
            emailLink: emailLink,
            _message: message,
            _isComplete: order.isComplete,
            _missingFields: order.missingFields
          };
        });
      } else {
        Logger.log('Extraction returned no orders. Flags: ' + (result.flags || []).join('; '));
      }
    } catch (e) {
      if (typeof logError === 'function') {
        logError('extractOrderFromEmail', e.toString(), { subject: subject, from: senderEmail });
      }
      Logger.log('Extraction error: ' + e);
    }
    extractTime += (new Date() - tExtract);

    if (extractedRecords.length === 0) {
      Logger.log('No records extracted - skipping');
    }

    // --- Deduplicate within this email ---
    var seen = {};
    var dedupedRecords = extractedRecords.filter(function(r) {
      var key = (r.container || '') + '|' + (r.mark || '') + '|' + (r.cargo || '') + '|' + (r.receiver || '');
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });

    // --- Write each record to sheet ---
    dedupedRecords.forEach(function(data, dataIndex) {
      Logger.log('Adding record ' + (dataIndex + 1) + '/' + dedupedRecords.length + ': ' + data.container);

      var dupCheck = checkForDuplicate(sheet, data);
      if (dupCheck.isDuplicate) {
        Logger.log('SKIPPING DUPLICATE: ' + data.container + ' | Already exists as ' + dupCheck.existingSample);
        if (typeof logWarning === 'function') {
          logWarning('processPDFsFromGmail', 'Skipped duplicate', {
            container: data.container,
            existingSample: dupCheck.existingSample
          });
        }
        return;
      }

      try {
        var msg = data._message;
        delete data._message;
        delete data._isComplete;
        delete data._missingFields;

        // --- Learning Engine: log raw data, then apply corrections ---
        if (typeof logExtraction === 'function') {
          logExtraction(data, { from: senderEmail, subject: subject });
        }
        if (typeof applyCorrections === 'function') {
          data = applyCorrections(data, senderEmail);
        }
        // --- End Learning Engine ---

        var tWrite = new Date();
        addDataToSheet(sheet, data, { skipLiveUpdate: true, skipAutoPrint: false, orderCache: orderCache });
        writeTime += (new Date() - tWrite);
        processedCount++;
        Logger.log('Successfully added');

        // --- Save attachments (PDFs and images) ---
        var tAtt = new Date();
        try {
          if (msg) {
            var lastRow = sheet.getLastRow();
            var sampleId = (sampleColIdx >= 0) ? sheet.getRange(lastRow, sampleColIdx + 1).getValue() : '';
            if (sampleId) {
              var savedPDFs = saveEmailPDFs(msg, String(sampleId));
              if (savedPDFs.length > 0) {
                var pdfUrls = savedPDFs.map(function(p) { return p.url; }).join('\n');
                if (attachColIdx >= 0) {
                  sheet.getRange(lastRow, attachColIdx + 1).setValue(pdfUrls);
                }
                Logger.log('Saved ' + savedPDFs.length + ' PDFs for ' + sampleId);
              }

              var savedImages = saveEmailImages(msg, String(sampleId));
              if (savedImages.length > 0) {
                var imgUrls = savedImages.map(function(p) { return p.url; }).join('\n');
                if (photoColIdx >= 0) {
                  sheet.getRange(lastRow, photoColIdx + 1).setValue(imgUrls);
                }
              }
            }
          }
        } catch (attachErr) {
          Logger.log('Attachment save failed (continuing): ' + attachErr);
        }
        attachTime += (new Date() - tAtt);

      } catch (e) {
        if (typeof logError === 'function') {
          logError('addDataToSheet', e.toString(), { container: data.container, mark: data.mark });
        }
        Logger.log('Failed to add record: ' + e);
        Logger.log('Error details: ' + e.stack);
      }
    });

    thread.addLabel(getOrCreateLabel('PDF_Processed'));
    Logger.log('Thread ' + (threadIndex + 1) + ' done in ' + (new Date() - tThread) + 'ms');
  });

  try {
    updateLiveOrdersView();
  } catch (e) {
    Logger.log('Live Orders update failed: ' + e);
  }

  var elapsed = ((new Date() - t0) / 1000).toFixed(1);
  Logger.log('=== PROCESSING COMPLETE ===');
  Logger.log('Total records added: ' + processedCount);
  Logger.log('PDF: ' + (pdfTime / 1000).toFixed(1) + 's | Extract: ' + (extractTime / 1000).toFixed(1) + 's | Write: ' + (writeTime / 1000).toFixed(1) + 's | Attach: ' + (attachTime / 1000).toFixed(1) + 's | Total: ' + elapsed + 's');
  SpreadsheetApp.getUi().alert('Processed ' + processedCount + ' records from ' + threads.length + ' emails in ' + elapsed + 's');
}


// ============================================================================
// SECTION 7: MAIN EXTRACTION ENTRY POINT
// ============================================================================

function extractOrderFromEmail(emailBody, pdfText, senderEmail, subject) {
  var result = {
    success: false,
    client: null,
    orders: [],
    rawData: {},
    flags: [],
    isComplete: true
  };

  if (isExcludedSender(senderEmail)) {
    result.flags.push('EXCLUDED: Internal email from ' + senderEmail);
    return result;
  }

  var cleanedPdfText = preprocessTableText(pdfText || '');
  var cleanedEmailBody = preprocessTableText(emailBody || '');
  var allText = cleanedEmailBody + '\n' + cleanedPdfText + '\n' + (subject || '');

  var hasSamplingInstructions = checkForSamplingInstructions(allText);
  result.client = identifyClient(senderEmail, allText);

  if (result.client && result.client.requiresSamplingInstructions && !hasSamplingInstructions) {
    result.flags.push('ARRIVAL_NOTICE_ONLY: No sampling instructions found');
    return result;
  }

  if (result.client) {
    result.orders = extractByClientType(result.client, cleanedEmailBody, cleanedPdfText, subject);
  } else {
    result.orders = extractGeneric(cleanedEmailBody, cleanedPdfText, subject);
  }

  // Post-process: validate references
  for (var i = 0; i < result.orders.length; i++) {
    result.orders[i].reference = validateReference(result.orders[i].reference, result.orders[i].container);
  }

  // Flag incomplete orders
  var missingContainerCount = 0;
  var missingReceiverCount = 0;

  for (var j = 0; j < result.orders.length; j++) {
    var ord = result.orders[j];
    var missingFields = [];

    if (!ord.container && !ord.mark && !ord.cargo) {
      missingFields.push('container/mark/cargo');
      missingContainerCount++;
    }
    if (!ord.receiver || (typeof ord.receiver === 'object' && !ord.receiver.address)) {
      missingFields.push('ship-to address');
      missingReceiverCount++;
    }

    ord.isComplete = missingFields.length === 0;
    ord.missingFields = missingFields;
    if (!ord.isComplete) result.isComplete = false;
  }

  if (missingContainerCount > 0 && missingContainerCount < result.orders.length) {
    result.flags.push(missingContainerCount + ' orders missing container/mark/cargo');
  }
  if (missingReceiverCount > 0 && missingReceiverCount < result.orders.length) {
    result.flags.push(missingReceiverCount + ' orders missing ship-to');
  }

  result.success = result.orders.length > 0;
  return result;
}


// ============================================================================
// SECTION 8: REFERENCE VALIDATION
// ============================================================================

function validateReference(ref, container) {
  if (!ref) return '';

  if (container && container.length >= 7) {
    var containerSuffixes = [
      container.substring(1),
      container.substring(2),
      container.substring(3),
      container.substring(4)
    ];
    for (var i = 0; i < containerSuffixes.length; i++) {
      if (ref === containerSuffixes[i] || ref === containerSuffixes[i] + '*') return '';
    }
  }

  if (ref === '.' || ref === '..' || ref.length < 2) return '';
  return ref;
}


// ============================================================================
// SECTION 9: TEXT PREPROCESSING & DETECTION
// ============================================================================

function preprocessTableText(text) {
  if (!text) return '';
  text = text.replace(/\b(C\d{6})(\d{2,3}[-\/]\d{3,4}[-\/]\d{3,5})/g, '$1 $2');
  text = text.replace(/(\d{3,5})\s*(\(['\d]+\)|\w+[A-Z][a-z]+)/g, '$1 $2');
  text = text.replace(/(\d+)\s*(g|lb|lbs)\s*(C\d{5,7})/gi, '$1$2 $3');
  text = text.replace(/([a-zA-Z]{3,})(\d+)(g|lb|lbs)\b/gi, '$1 $2$3');
  text = text.replace(/(\d)\s*([A-Z]{2})\s*(\d)/gi, '$1 $2 $3');
  text = text.replace(/([a-zA-Z])(\bC\d{5,7})/g, '$1 $2');
  text = text.replace(/\s{2,}/g, ' ');
  return text;
}

function checkForSamplingInstructions(text) {
  text = text.toLowerCase();
  for (var i = 0; i < SAMPLING_INDICATORS.length; i++) {
    var pattern = new RegExp(SAMPLING_INDICATORS[i], 'i');
    if (pattern.test(text)) return true;
  }
  return false;
}


// ============================================================================
// SECTION 10: CLIENT IDENTIFICATION
// ============================================================================

function identifyClient(senderEmail, allText) {
  senderEmail = (senderEmail || '').toLowerCase();
  allText = (allText || '').toLowerCase();

  // Check document content for brokered clients first
  var key, client;
  for (key in KNOWN_CLIENTS) {
    client = KNOWN_CLIENTS[key];
    if (client.broker) {
      for (var i = 0; i < client.keywords.length; i++) {
        if (allText.indexOf(client.keywords[i].toLowerCase()) >= 0) return client;
      }
    }
  }

  // Check delivery/account patterns
  var documentClient = null;
  var patterns = [
    /delivery to[^:]*:\s*([a-z][a-z\s\/]+)/i,
    /account of[:\s]*([a-z][a-z\s\/]+)/i,
    /bill to account[:\s]*([a-z][a-z\s\/]+)/i,
    /send samples? to[:\s]*([a-z][a-z\s]+)/i
  ];

  for (var p = 0; p < patterns.length; p++) {
    var match = allText.match(patterns[p]);
    if (match) { documentClient = match[1].trim(); break; }
  }

  if (documentClient) {
    for (key in KNOWN_CLIENTS) {
      client = KNOWN_CLIENTS[key];
      for (var j = 0; j < client.keywords.length; j++) {
        if (documentClient.indexOf(client.keywords[j].toLowerCase()) >= 0) return client;
      }
    }
  }

  // Check domain
  for (key in KNOWN_CLIENTS) {
    client = KNOWN_CLIENTS[key];
    for (var k = 0; k < client.domains.length; k++) {
      if (senderEmail.indexOf(client.domains[k].toLowerCase()) >= 0) {
        // NKG disambiguation
        if (client.domains[k] === '@nkg.coffee') {
          if (allText.indexOf('interamerican') >= 0) return KNOWN_CLIENTS['interamerican'];
          return KNOWN_CLIENTS['rothfos'];
        }
        // ECOM disambiguation
        if (client.domains[k] === '@ecomtrading.com') {
          if (allText.indexOf('atlantic specialty') >= 0 || allText.indexOf('ascilab') >= 0) return KNOWN_CLIENTS['atlantic_specialty'];
          if (allText.indexOf('atlantic (usa)') >= 0 || allText.indexOf('letter of entry') >= 0) return KNOWN_CLIENTS['atlantic_usa'];
          return null;
        }
        return client;
      }
    }
  }

  // Keyword fallback
  for (key in KNOWN_CLIENTS) {
    client = KNOWN_CLIENTS[key];
    if (client.broker) continue;
    for (var m = 0; m < client.keywords.length; m++) {
      if (allText.indexOf(client.keywords[m].toLowerCase()) >= 0) return client;
    }
  }

  return null;
}


// ============================================================================
// SECTION 11: CLIENT-SPECIFIC EXTRACTOR (ROUTER)
// All custom parser functions live in Parser_V3.gs.
// ============================================================================

function extractByClientType(client, emailBody, pdfText, subject) {
  if (client.useCustomParser) {
    if (client.name === 'Serengeti Trading Company')        return extractSerengetiOrders(emailBody, pdfText, subject);
    if (client.name === 'Atlantic (USA), LLC')              return extractAtlanticOrders(emailBody, pdfText, subject);
    if (client.name === 'Ally Coffee Trading S.A.')         return extractAllyOrders(emailBody, pdfText, subject);
    if (client.name === 'Armenia Coffee Corp')              return extractArmeniaOrders(emailBody, pdfText, subject);
    if (client.name === 'American Coffee Corporation')      return extractAmcofOrders(emailBody, pdfText, subject);
    if (client.name === 'Coffee America USA Corp')          return extractCoffeeAmericaOrders(emailBody, pdfText, subject);
    if (client.name === 'InterAmerican Coffee (NKG)')       return extractInterAmericanOrders(emailBody, pdfText, subject);
    if (client.name === 'International Coffee Corporation') return extractIccOrders(emailBody, pdfText, subject);
    if (client.name === 'Osito Coffee')                     return extractOsitoOrders(emailBody, pdfText, subject);
    if (client.name === 'Rothfos (Neumann Kaffee Gruppe)')  return extractRothfosOrders(emailBody, pdfText, subject);
    if (client.name === 'Louis Dreyfus Company')            return extractLdcOrders(emailBody, pdfText, subject);
    if (client.name === 'Sucafina NA')                      return extractSucafinaOrders(emailBody, pdfText, subject);
    if (client.name === 'The Coffee Source LLC')            return extractCoffeeSourceOrders(emailBody, pdfText, subject);

    Logger.log('useCustomParser=true but no parser matched for: ' + client.name);
  }

  // Generic extractor for Olam, Paragon, Excelco, Covoya, etc.
  var allText = (emailBody || '') + '\n' + (pdfText || '');
  var orders = [];

  var containers  = extractContainers(allText);
  var marks       = extractMarks(allText);
  var cargos      = extractCargos(allText);
  var references  = extractReferences(allText, client);
  var warehouse   = detectWarehouse(allText) || client.warehouse;
  var receivers   = extractReceivers(emailBody, pdfText, client);
  var bagCounts   = extractBagCounts(allText);

  var sampleSizeResult = extractSampleSize(allText, client.defaultSampleSize || '2 lb');
  var sampleSize = sampleSizeResult.size;
  var sampleSizeFlag = sampleSizeResult.flag;

  var shipping = extractShipping(allText, emailBody) || client.defaultShipping;
  var conditionNotes = extractConditionNotes(allText);

  var explicitOrderNum = extractSampleOrderNumber(allText, subject, references[0]);
  var originalSender = extractOriginalSender(emailBody, null);
  var sampleOrderNum = explicitOrderNum || originalSender || '';

  if (receivers.length === 0 && client.defaultReceiver) {
    receivers.push(client.defaultReceiver);
  }

  var orderCount = Math.max(containers.length, marks.length, cargos.length, 1);

  function getSampleSizeForReceiver(clientConfig, receiverName, defaultSize) {
    if (!clientConfig || !clientConfig.sampleSizeToSelf) return defaultSize;
    var receiverLower = (receiverName || '').toLowerCase();
    var clientKeyword = (clientConfig.name || '').toLowerCase().split(' ')[0];
    if (receiverLower.indexOf(clientKeyword) >= 0) return clientConfig.sampleSizeToSelf;
    return defaultSize;
  }

  if (containers.length > 0 || marks.length > 0 || cargos.length > 0) {
    var receiversToUse = receivers.length > 0 ? receivers : [null];

    for (var i = 0; i < orderCount; i++) {
      for (var r = 0; r < receiversToUse.length; r++) {
        var receiverName = getReceiverName(receiversToUse[r]);
        var orderSampleSize = getSampleSizeForReceiver(client, receiverName, sampleSize);

        var order = {
          client: client.name,
          sampleOrderNum: sampleOrderNum,
          container: containers[i] || '',
          mark: marks[i] || marks[0] || '',
          cargo: cargos[i] || '',
          reference: references[i] || references[0] || '',
          warehouse: warehouse,
          receiver: receiverName,
          sampleSize: orderSampleSize,
          shipping: shipping,
          fedexAccount: client.fedexAccount || extractFedExAccount(allText),
          bags: bagCounts[i] || '',
          origin: '',
          description: '',
          comments: ''
        };

        if (order.mark) order.origin = getOriginFromMark(order.mark);
        order.description = extractDescription(allText, order.origin);

        var commentParts = [];
        if (sampleSizeFlag) commentParts.push(sampleSizeFlag);
        if (conditionNotes) commentParts.push(conditionNotes);
        order.comments = commentParts.join('; ');

        orders.push(order);
      }
    }
  } else {
    var receiversForSingle = receivers.length > 0 ? receivers : [null];

    for (var rs = 0; rs < receiversForSingle.length; rs++) {
      var singleReceiverName = getReceiverName(receiversForSingle[rs]);
      var singleSampleSize = getSampleSizeForReceiver(client, singleReceiverName, sampleSize);

      var singleOrder = {
        client: client.name,
        sampleOrderNum: sampleOrderNum,
        container: '',
        mark: '',
        cargo: '',
        reference: references[0] || '',
        warehouse: warehouse,
        receiver: singleReceiverName,
        sampleSize: singleSampleSize,
        shipping: shipping,
        fedexAccount: client.fedexAccount || extractFedExAccount(allText),
        bags: '',
        origin: '',
        description: '',
        comments: ''
      };
      singleOrder.description = extractDescription(allText, singleOrder.origin);

      var commentParts2 = [];
      if (sampleSizeFlag) commentParts2.push(sampleSizeFlag);
      if (conditionNotes) commentParts2.push(conditionNotes);
      singleOrder.comments = commentParts2.join('; ');

      orders.push(singleOrder);
    }
  }

  return orders;
}


// ============================================================================
// SECTION 12: GENERIC EXTRACTOR
// ============================================================================

function extractGeneric(emailBody, pdfText, subject) {
  var allText = (emailBody || '') + '\n' + (pdfText || '') + '\n' + (subject || '');
  var orders = [];

  var containers = extractContainers(allText);
  var marks      = extractMarks(allText);
  var cargos     = extractCargos(allText);
  var references = extractReferences(allText, null);
  var warehouse  = detectWarehouse(allText);
  var shipping   = extractShipping(allText, emailBody);
  var receivers  = extractReceivers(emailBody, pdfText, null);

  var sampleSizeResult = extractSampleSize(allText, '2 lb');
  var sampleSize = sampleSizeResult.size;
  var sampleSizeFlag = sampleSizeResult.flag;

  var conditionNotes = extractConditionNotes(allText);

  var explicitOrderNum = extractSampleOrderNumber(allText, subject, references[0]);
  var originalSender = extractOriginalSender(emailBody, null);
  var sampleOrderNum = explicitOrderNum || originalSender || '';

  var orderCount = Math.max(containers.length, marks.length, cargos.length);

  if (orderCount > 0) {
    for (var i = 0; i < orderCount; i++) {
      var order = {
        client: 'Unknown',
        sampleOrderNum: sampleOrderNum,
        container: containers[i] || '',
        mark: marks[i] || '',
        cargo: cargos[i] || '',
        reference: references[i] || references[0] || '',
        warehouse: warehouse || 'Unknown',
        receiver: getReceiverName(receivers[i] || receivers[0]),
        sampleSize: sampleSize,
        shipping: shipping || 'FedEx',
        fedexAccount: extractFedExAccount(allText),
        bags: '',
        origin: '',
        description: '',
        comments: ''
      };
      if (order.mark) order.origin = getOriginFromMark(order.mark);
      order.description = extractDescription(allText, order.origin);

      var commentParts = [];
      if (sampleSizeFlag) commentParts.push(sampleSizeFlag);
      if (conditionNotes) commentParts.push(conditionNotes);
      order.comments = commentParts.join('; ');

      orders.push(order);
    }
  }

  mapReceiversToOrders(orders, receivers, allText);
  return orders;
}


// ============================================================================
// SECTION 13: FIELD EXTRACTORS
// ============================================================================

function extractContainers(text) {
  var containers = [];
  var patterns = [
    /\b([A-Z]{4}\d{7})\b/g,
    /\b([A-Z]{4}[-\s]?\d{6}[-\s]?\d)\b/g,
    /\b([A-Z]{4})\s+(\d{7})\b/g,
    /Container[#:\s]*([A-Z]{4}\d{7})/gi,
    /Container\s*No[.:\s]*([A-Z]{4}\d{7})/gi
  ];

  for (var p = 0; p < patterns.length; p++) {
    patterns[p].lastIndex = 0;
    var match;
    while ((match = patterns[p].exec(text)) !== null) {
      var container = match[2] ? (match[1] + match[2]).toUpperCase() : match[1].replace(/[-\s]/g, '').toUpperCase();
      if (containers.indexOf(container) === -1 && container.length === 11) {
        containers.push(container);
      }
    }
  }
  return containers;
}

function extractMarks(text) {
  var marks = [];
  var match;

  var pattern1 = /\b(\d{1,4})\s*[-\/]\s*(\d{2,5})\s*[-\/]\s*(\d{2,5})\b/g;
  pattern1.lastIndex = 0;
  while ((match = pattern1.exec(text)) !== null) {
    var part1 = match[1], part2 = match[2], part3 = match[3];
    var mark = part1 + '/' + part2 + '/' + part3;

    if (isLikelyDate(match[0])) continue;
    if (isLikelyPhone(match[0])) continue;
    if (/^20[0-3]\d$/.test(part1) || /^20[0-3]\d$/.test(part2) || /^20[0-3]\d$/.test(part3)) continue;

    if (part2.length >= 3 || part3.length >= 3) {
      if (marks.indexOf(mark) === -1) marks.push(mark);
    }
  }

  var pattern2 = /\b(\d{2,3})\s*[-\/]\s*([A-Z]{2})\s*[-\/]\s*(\d{2}[-]\d{2})\s*[-\/]\s*(\d{4})\b/g;
  pattern2.lastIndex = 0;
  while ((match = pattern2.exec(text)) !== null) {
    var mark2 = match[1] + '/' + match[2] + '/' + match[3] + '/' + match[4];
    if (marks.indexOf(mark2) === -1) marks.push(mark2);
  }

  return marks;
}

function isLikelyDate(str) {
  var parts = str.split(/[-\/]/);
  if (parts.length !== 3) return false;
  var p1 = parts[0], p2 = parts[1], p3 = parts[2];
  if (/^20[0-3]\d$/.test(p1) || /^20[0-3]\d$/.test(p2) || /^20[0-3]\d$/.test(p3)) return true;
  if (p1.length <= 2 && p2.length <= 2 && p3.length <= 2) {
    var n1 = parseInt(p1, 10), n2 = parseInt(p2, 10);
    if ((n1 >= 1 && n1 <= 12) && (n2 >= 1 && n2 <= 31)) return true;
    if ((n1 >= 1 && n1 <= 31) && (n2 >= 1 && n2 <= 12)) return true;
  }
  return false;
}

function isLikelyPhone(str) {
  return /^\d{3}[-]\d{3}[-]\d{4}$/.test(str);
}

function extractCargos(text) {
  var cargos = [];
  var patterns = [
    /\bC(\d{5,7})\b/gi,
    /Cargo[#:\s]*C?(\d{5,7})/gi,
    /Cargo\s*Ref[#:\s]*C?(\d{5,7})/gi,
    /Cargo#\s*(\d{5,7})/gi
  ];

  for (var p = 0; p < patterns.length; p++) {
    patterns[p].lastIndex = 0;
    var match;
    while ((match = patterns[p].exec(text)) !== null) {
      var cargo = 'C' + match[1];
      if (cargos.indexOf(cargo) === -1) cargos.push(cargo);
    }
  }

  if (/\bCARGO\b|Cargo\s*Nbr/i.test(text)) {
    var sixDigitPattern = /\b(\d{6})\b/g;
    sixDigitPattern.lastIndex = 0;
    var match2;
    while ((match2 = sixDigitPattern.exec(text)) !== null) {
      var num = match2[1];
      if (num.charAt(0) === '0') continue;
      if (/^(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])/.test(num)) continue;
      if (/^(0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])/.test(num)) continue;
      if (/202[4-9]/.test(num)) continue;
      var cargo2 = 'C' + num;
      if (cargos.indexOf(cargo2) === -1) cargos.push(cargo2);
    }
  }

  return cargos;
}

function extractReferences(text, client) {
  var refs = [];
  var patterns = [
    /\bS\d{5,12}\b/g,
    /\bR\d{9,12}\b/g,
    /\bP\d{5,9}\b/g,
    /\bS\d{5,7}[-]\d{1,2}\b/g,
    /\bAP\d{5}[A-Z]?\b/g,
    /\bBD\d{3}\b/g,
    /\bPO\d{5,6}\b/gi,
    /\b\d{7}\.\d\b/g,
    /Our Ref(?:erence)?[#:\s]+([A-Z0-9\-\/]{5,})/gi,
    /Sample Ref[#:\s]+([A-Z0-9]{5,})/gi,
    /\bSO\d{5,7}\b/g,
    /Purch[:\s]+(\d{5,7})/gi
  ];

  for (var p = 0; p < patterns.length; p++) {
    patterns[p].lastIndex = 0;
    var match;
    while ((match = patterns[p].exec(text)) !== null) {
      var ref = (match[1] || match[0]).trim();
      if (ref.length >= 5 &&
          refs.indexOf(ref) === -1 &&
          !/^[a-zA-Z]+$/.test(ref) &&
          !/^(Reference|Container|erence|rence|ontainer)$/i.test(ref)) {
        refs.push(ref);
      }
    }
  }
  return refs;
}

function detectWarehouse(text) {
  text = text.toLowerCase();
  // Continental — try specific address first, then general
  if (text.indexOf('300 middlesex') >= 0) return 'CONTINENTAL_300';
  if (text.indexOf('continental') >= 0 || text.indexOf(' cti') >= 0 || text.indexOf('\ncti') >= 0 ||
      text.indexOf('cti ') >= 0 || text.indexOf('200 middlesex') >= 0 || text.indexOf('carteret') >= 0) {
    return 'CONTINENTAL_200';
  }
  // RPM — try specific locations first, then general
  if (text.indexOf('2900 woodbridge') >= 0 || text.indexOf('rpm woodbridge') >= 0) return 'RPM_WOODBRIDGE';
  if (text.indexOf('talmage') >= 0 || text.indexOf('rpm talmage') >= 0) return 'RPM_TALMAGE';
  if (text.indexOf('1500 rahway') >= 0 || text.indexOf('avenel') >= 0 || text.indexOf('rpm avenel') >= 0) return 'RPM_AVENEL';
  if (text.indexOf('rpm') >= 0 || text.indexOf('edison') >= 0) return 'RPM_AVENEL';
  // Other warehouses
  if (text.indexOf('florence warehouse') >= 0 || text.indexOf('florence wh') >= 0) return 'Florence';
  if (text.indexOf('cadeco') >= 0) return 'Cadeco';
  if (text.indexOf('dupuy') >= 0) return 'Dupuy';
  if (text.indexOf('northdale') >= 0 || text.indexOf('rpm northdale') >= 0) return 'RPM_NORTHDALE';
  return null;
}

function extractSampleSize(text, defaultSize) {
  defaultSize = defaultSize || '2 lb';
  var unusualWeights = [];

  var explicitPatterns = [
    /Sample\s*Size[:\s]*(\d+)\s*(LB|lb|lbs|Lbs|g|grams?)/gi,
    /Sample\s*(?:Qty|Wt)[:\s]*(\d+)\s*(lb|lbs|Lbs|g|grams?)/gi,
    /(?:send|draw|pull)\s+(?:a\s+)?(\d+)\s*(lb|lbs|g|gram)s?\s+sample/gi,
    /(\d+)\s*(lb|lbs|pound|pounds)\s+sample/gi,
    /(\d+)\s*(g|gram|grams)\s+sample/gi,
    /Standard\s+(\d+)#/gi,
    /(\d+)#\s+sampling/gi
  ];

  for (var p = 0; p < explicitPatterns.length; p++) {
    explicitPatterns[p].lastIndex = 0;
    var match = explicitPatterns[p].exec(text);
    if (match) {
      var size = parseInt(match[1], 10);
      var unit = (match[2] || '').toLowerCase();

      if (unit === 'g' || unit === 'gram' || unit === 'grams') {
        if (size <= 1000) return { size: size + 'g', flag: null };
        else unusualWeights.push(size + 'g');
      } else {
        if (size <= 20) return { size: size + ' lb', flag: null };
        else unusualWeights.push(size + ' lb');
      }
    }
  }

  var conditionPattern = /(?:slack|stain|recondition|damage|wet|torn|ripped|short|sound)/i;
  var weightPatterns = [
    /Amount[:\s]*(\d+)\s*(Lb|lb|lbs|LB|g|grams?)/gi,
    /Allowance\s+(\d+)\s*(Lb|lb|lbs|LB|g|grams?)/gi
  ];

  for (var q = 0; q < weightPatterns.length; q++) {
    weightPatterns[q].lastIndex = 0;
    var match2;
    while ((match2 = weightPatterns[q].exec(text)) !== null) {
      var contextStart = Math.max(0, match2.index - 100);
      var contextEnd = Math.min(text.length, match2.index + match2[0].length + 100);
      var context = text.substring(contextStart, contextEnd);
      if (conditionPattern.test(context)) continue;

      var size2 = parseInt(match2[1], 10);
      var unit2 = (match2[2] || '').toLowerCase();

      if (unit2 === 'g' || unit2 === 'gram' || unit2 === 'grams') {
        if (size2 <= 1000) return { size: size2 + 'g', flag: null };
        else unusualWeights.push(size2 + 'g');
      } else {
        if (size2 <= 20) return { size: size2 + ' lb', flag: null };
        else unusualWeights.push(size2 + ' lb');
      }
    }
  }

  if (unusualWeights.length > 0) {
    var uniqueWeights = [];
    for (var w = 0; w < unusualWeights.length; w++) {
      if (uniqueWeights.indexOf(unusualWeights[w]) === -1) uniqueWeights.push(unusualWeights[w]);
    }
    return { size: defaultSize, flag: 'Check sample size - found ' + uniqueWeights.join(', ') + ' in email' };
  }

  return { size: defaultSize, flag: null };
}

function extractConditionNotes(text) {
  var notes = [];
  var cleanText = text.replace(/>\s*>/g, ' ').replace(/\n>\s*/g, '\n').replace(/\s+/g, ' ');

  var patterns = [
    /(\d+\s*STAINED?\s*BGS?\s*RECONDITIONED\s*INTO\s*\d+\s*SOUND\s*BGS?\s*\d*\s*SLACK\s*BGS?\s*@?\s*\d*\s*LBS?)/gi,
    /(\d+\s*SLACK\s*BGS?\s*@\s*\d+\s*LBS?)/gi,
    /(\d+\s*STAINED?\s*BGS?(?:\s*@\s*\d+\s*LBS?)?)/gi
  ];

  for (var p = 0; p < patterns.length; p++) {
    patterns[p].lastIndex = 0;
    var match;
    while ((match = patterns[p].exec(cleanText)) !== null) {
      var note = match[1].trim().replace(/\s+/g, ' ');
      if (note.length >= 8 && note.length <= 80 && notes.indexOf(note) === -1 &&
          !/[<>]/.test(note) && !/\b[A-Z]{4}\d{7}\b/.test(note) &&
          !/\bC\d{6}\b/.test(note) && !/\d{3}\/\d{4}\/\d{3}/.test(note)) {
        notes.push(note);
      }
    }
  }

  return notes.slice(0, 2).join('; ');
}

function extractSampleOrderNumber(text, subject, reference) {
  var allText = (subject || '') + '\n' + (text || '');
  var patterns = [
    /Sampling Order\s*#[:\s]*(\d+)/i,
    /Sampling Order[:\s]+(\d{5})/i,
    /Sampling\s+Order\s*#?\s*:?\s*(\d{5})/i,
    /Sample Order\s*#?[:\s]*(\d{5})/i,
    /Order\s*#\s*:?\s*(\d{5,})/i
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = allText.match(patterns[i]);
    if (match && match[1]) return match[1];
  }
  return null;
}

function extractOriginalSender(emailBody) {
  if (!emailBody) return null;

  var patterns = [
    /From:\s*([A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+)\s*<([^>]+)>/g,
    /From:\s*"([^"]+)"\s*<([^>]+)>/g,
    /From:\s*([A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+)\s*$/gm,
    /From:\s*([a-zA-Z][a-zA-Z.]+)@([a-zA-Z]+)\.(com|coffee|cr)\b/g
  ];

  var matches = [];

  for (var p = 0; p < patterns.length; p++) {
    patterns[p].lastIndex = 0;
    var match;
    while ((match = patterns[p].exec(emailBody)) !== null) {
      var sender = match[1].trim();
      var email = match[2] ? match[2].toLowerCase() : '';

      if (sender.indexOf('.') >= 0 && sender.indexOf(' ') === -1) {
        sender = sender.split('.').map(function(part) {
          return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
        }).join(' ');
      }

      var senderLower = sender.toLowerCase();
      var emailLower = email.toLowerCase();
      var isInternal =
          emailLower.indexOf('commoditysampler') >= 0 ||
          emailLower.indexOf('samplerdan') >= 0 ||
          emailLower.indexOf('danboy') >= 0 ||
          emailLower.indexOf('rpmwarehouse') >= 0 ||
          senderLower.indexOf('commodity') >= 0 ||
          senderLower.indexOf('sampler') >= 0 ||
          senderLower === 'daniel finkel' ||
          senderLower === 'dan finkel' ||
          senderLower === 'melissa finkel' ||
          senderLower === 'christine ingleton' ||
          senderLower.indexOf('paul napolitano') >= 0 ||
          senderLower.indexOf('rpm') >= 0;

      if (sender && !isInternal && sender.length > 3 && sender.length < 50) {
        matches.push(sender);
      }
    }
  }

  if (matches.length > 0) return matches[matches.length - 1];

  var emailPattern = /([a-zA-Z]+\.[a-zA-Z]+)@(?!commoditysampler|gmail|rpmwarehouse)([a-zA-Z]+)\./g;
  emailPattern.lastIndex = 0;
  var match2;
  while ((match2 = emailPattern.exec(emailBody)) !== null) {
    var namePart = match2[1];
    var name = namePart.split('.').map(function(part) {
      return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
    }).join(' ');
    if (name.length > 4 && name.split(' ').length >= 2) return name;
  }

  return null;
}

function extractShipping(allText, emailBody) {
  var text = ((emailBody || '') + '\n' + (allText || '')).toLowerCase();

  if (/standard\s+overnight/i.test(text))  return 'FedEx Standard Overnight';
  if (/stand\s+overnight/i.test(text))     return 'FedEx Standard Overnight';
  if (/priority\s+overnight/i.test(text))  return 'FedEx Priority';
  if (/fedex\s+ground/i.test(text))        return 'FedEx Ground';
  if (/fedex\s+2[\s-]?day/i.test(text))   return 'FedEx 2 Day';
  if (/(?:delivery|level of service).*2[\s-]?day/i.test(text)) return 'FedEx 2 Day';
  if (/ups\s+2nd\s+day/i.test(text))      return 'UPS 2nd Day';
  if (/ups\s+2[\s-]?day/i.test(text))     return 'UPS 2 Day';
  if (/ups\s+ground/i.test(text))         return 'UPS Ground';
  if (/fedex/i.test(text))                return 'FedEx';
  if (/ups/i.test(text))                  return 'UPS';
  return null;
}

function extractFedExAccount(text) {
  var patterns = [
    /(?:fedex|account)[:\s#]*(\d{4}[-]?\d{4}[-]?\d{1,4})/gi,
    /account\s*[-:#]?\s*(\d{9,12})/gi,
    /using our account\s*[-:#]?\s*(\d{9,12})/gi
  ];
  for (var p = 0; p < patterns.length; p++) {
    var match = patterns[p].exec(text);
    if (match) return match[1];
  }
  return null;
}

function extractBagCounts(text) {
  var bags = [];
  var patterns = [
    { regex: /(\d+)\s*SSACK/gi, suffix: 'ss' },
    { regex: /(\d+)\s*SACK(?!S)/gi, suffix: 's' },
    { regex: /(\d+)\s*BAGS/gi, suffix: '' },
    { regex: /(\d+)\s*BAG\b/gi, suffix: '' }
  ];

  var allMatches = [];
  for (var p = 0; p < patterns.length; p++) {
    var pattern = patterns[p];
    pattern.regex.lastIndex = 0;
    var match;
    while ((match = pattern.regex.exec(text)) !== null) {
      allMatches.push({ count: match[1] + pattern.suffix, position: match.index });
    }
  }

  allMatches.sort(function(a, b) { return a.position - b.position; });
  for (var i = 0; i < allMatches.length; i++) bags.push(allMatches[i].count);
  return bags;
}

function extractDescription(text, origin) {
  var description = origin || '';
  var gradePatterns = [
    /Quality[:\s]*([A-Za-z\s]+(?:GREEN COFFEE|ARABICA|ROBUSTA)[A-Za-z\s]*)/gi,
    /(?:Grade|Quality)[:\s]*([A-Za-z\s]+(?:EP|SHB|HB|FAQ|Organic|Washed|Natural|Honey)[A-Za-z\s]*)/gi,
    /(Organic[A-Za-z\s]*)/gi,
    /(Excelso[A-Za-z\s]*)/gi,
    /([A-Za-z]+\s+(?:Washed|Natural|Honey))/gi
  ];
  for (var p = 0; p < gradePatterns.length; p++) {
    var match = gradePatterns[p].exec(text);
    if (match) {
      var grade = match[1].trim();
      if (description.indexOf(grade) === -1) description += ' ' + grade;
      break;
    }
  }
  return description.trim();
}


// ============================================================================
// SECTION 14: RECEIVER EXTRACTION
// ============================================================================

function stripForwardedHeaders(text) {
  if (!text) return '';
  text = text.replace(/-{3,}\s*Forwarded\s+message\s*-{3,}[^\n]*\n(?:\s*(?:From|Date|Subject|To|Cc|Sent|Bcc)[:\s][^\n]*\n)*/gi, '\n');
  text = text.replace(/From:\s*[^\n]+\nSent:\s*[^\n]+\nTo:\s*[^\n]+\n(?:Cc:\s*[^\n]+\n)?Subject:\s*[^\n]+\n/gi, '\n');
  text = text.replace(/On\s+[A-Z][a-z]{2,8}\s+\d{1,2},?\s+\d{4},?\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)?[^\n]*wrote:\s*/gi, '\n');
  text = text.replace(/On\s+\d{1,2}\/\d{1,2}\/\d{2,4}\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)?[^\n]*wrote:\s*/gi, '\n');
  text = text.replace(/On\s+[A-Z][a-z]+,\s+[A-Z][a-z]+\s+\d{1,2},\s+\d{4}\s+at\s+\d{1,2}:\d{2}[^\n]*wrote:\s*/gi, '\n');
  text = text.replace(/\nDate:\s*[^\n]*\d{4}\s+at\s+\d[^\n]*/gi, '\n');
  text = text.replace(/\d{4}\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)?\s*Subject:\s*[^\n]*/gi, '\n');
  text = text.replace(/\d{4}\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)?\s*To:\s*[^\n]*/gi, '\n');
  text = text.replace(/\n(?:Thank you,?|Thanks,?|Best,?|Regards,?|Best regards,?)\s*\n[A-Z][a-z]+\s+[A-Z][a-z]+\n(?:[^\n]+\n){0,6}/gi, '\n');
  return text;
}

function extractReceivers(emailBody, pdfText, client) {
  var receivers = [];
  var cleanEmail = stripForwardedHeaders(emailBody || '');
  var allText = cleanEmail + '\n' + (pdfText || '');

  allText = allText.replace(/\r\n/g, '\n').replace(/[ \t]+/g, ' ');
  allText = allText.replace(/>\s*>/g, ' ').replace(/\s*>\s*/g, ' ');
  allText = allText.replace(/\*\s*/g, '').replace(/\s{2,}/g, ' ');

  var sendPatterns = [
    /Send\s+Samples?\s+to[:\s]+(.+?)(?=\s*Account)/i,
    /Send\s+Samples?\s+to[:\s]+([^\d\n]+?)(?=\s*\d)/i,
    /Send\s+Samples?\s+to[:\s]+([A-Z][A-Za-z\s&',\.]+?)\s*\n/i
  ];

  for (var sp = 0; sp < sendPatterns.length; sp++) {
    if (receivers.length > 0) break;
    var sendMatch = allText.match(sendPatterns[sp]);
    if (sendMatch) splitAndAddReceivers(receivers, sendMatch[1]);
  }

  if (receivers.length === 0) {
    var shipPatterns = [
      /Ship\s+to[:\s]+([^\d\n]+?)(?=\s*\d)/i,
      /Ship\s+to[:\s]+([A-Z][A-Za-z\s&',\.]+?)\s*\n/i
    ];
    for (var shp = 0; shp < shipPatterns.length; shp++) {
      var shipMatch = allText.match(shipPatterns[shp]);
      if (shipMatch) { splitAndAddReceivers(receivers, shipMatch[1]); break; }
    }
  }

  if (receivers.length === 0) {
    var deliverMatch = allText.match(/Deliver\s+to[:\s]+([^\d\n]+?)(?=\s*\d|\n)/i);
    if (deliverMatch) splitAndAddReceivers(receivers, deliverMatch[1]);
  }

  if (receivers.length === 0) {
    var subjectMatch = allText.match(/\d+\s+Samples?\s+(?:Please\s+)?[-\u2013]\s*([A-Za-z][A-Za-z\s&']+?)\s*[-\u2013]/i);
    if (subjectMatch) {
      var name4 = subjectMatch[1].trim();
      if (isValidReceiverName(name4)) receivers.push({ name: name4, company: name4 });
    }
  }

  if (receivers.length === 0) {
    var customerMatch = allText.match(/Customer:\s*([A-Z][A-Za-z\s&',\.]+?)(?=\s*Street|Recipient|$)/i);
    if (customerMatch) {
      var custName = customerMatch[1].trim();
      if (isValidReceiverName(custName)) receivers.push({ name: custName, company: custName });
    }
  }

  if (receivers.length === 0) {
    var accountMatch = allText.match(/Account\s+of[:\s]+([A-Za-z][A-Za-z\s&',\.]+?)(?=\s*\n|$)/i);
    if (accountMatch) {
      var name5 = accountMatch[1].trim();
      if (isValidReceiverName(name5)) receivers.push({ name: name5, company: name5 });
    }
  }

  return receivers;
}

function splitAndAddReceivers(receivers, rawName) {
  rawName = (rawName || '').trim();
  if (!rawName) return;

  var normalized = rawName.replace(/\s+/g, ' ');
  var parts = [];

  if (/ & [A-Za-z]/i.test(normalized)) parts = normalized.split(/ & /i);
  else if (/ and /i.test(normalized)) parts = normalized.split(/ and /i);
  else parts = [normalized];

  for (var i = 0; i < parts.length; i++) {
    var name = parts[i].trim();
    if (isValidReceiverName(name)) receivers.push({ name: name, company: name });
  }
}

function isValidReceiverName(name) {
  if (!name || name.length < 3 || name.length > 80) return false;
  if (!/[a-zA-Z]{2,}/.test(name)) return false;

  if (/^\d|^Suite|^Floor|Subject:|From:|To:|Day\s*>|FedEx|UPS|overnight|ground|priority|please\s|thank|regard|best|sincerely/i.test(name)) return false;
  if (/\d{1,2}:\d{2}/.test(name)) return false;
  if (/\d{1,2}\/\d{1,2}\/\d{2,4}/.test(name)) return false;
  if (/^Continental|^RPM|^Avenel|^Edison|^Carteret|^Keasbey|^Middlesex|Terminals?\s*Inc/i.test(name)) return false;
  if (/^\d+\s+\w+\s+(Street|Avenue|Road|Drive|Way|Blvd)/i.test(name)) return false;

  return true;
}

function getReceiverName(receiver) {
  if (!receiver) return '';
  if (typeof receiver === 'string') return receiver;
  if (typeof receiver === 'object') return receiver.name || receiver.company || '';
  return '';
}

function mapReceiversToOrders(orders, receivers, fullText) {
  if (!orders || !Array.isArray(orders) || !receivers || receivers.length === 0 || !fullText) return orders || [];

  var receiverPositions = receivers.map(function(r) {
    var searchText = r.name || r.company || r.address || '';
    return { receiver: r, position: fullText.indexOf(searchText) };
  }).filter(function(rp) { return rp.position >= 0; });

  if (receiverPositions.length === 0) {
    var fallbackReceiver = receivers[0];
    orders.forEach(function(order) {
      if (!order.receiver) order.receiver = fallbackReceiver.name || fallbackReceiver.company || '';
    });
    return orders;
  }

  orders.forEach(function(order) {
    if (order.receiver) return;
    var orderMarker = order.cargo || order.container || order.mark || '';
    var orderPos = orderMarker ? fullText.indexOf(orderMarker) : -1;

    if (orderPos < 0) {
      order.receiver = receivers[0].name || receivers[0].company || '';
      return;
    }

    var closestReceiver = null;
    var closestDistance = Infinity;
    receiverPositions.forEach(function(rp) {
      if (rp.position < orderPos) {
        var distance = orderPos - rp.position;
        if (distance < closestDistance) { closestDistance = distance; closestReceiver = rp.receiver; }
      }
    });

    if (closestReceiver) order.receiver = closestReceiver.name || closestReceiver.company || '';
    else if (receivers.length > 0) order.receiver = receivers[0].name || receivers[0].company || '';
  });

  return orders;
}


// ============================================================================
// SECTION 15: DRIVE / ATTACHMENT FUNCTIONS
// ============================================================================

function getGmailMessageLink(message) {
  if (!message) return '';
  try {
    var thread = message.getThread();
    return thread ? 'https://mail.google.com/mail/u/0/#inbox/' + thread.getId() : '';
  } catch (e) { return ''; }
}

function getOrCreateSampleFolder(sampleId) {
  var rootFolders = DriveApp.getFoldersByName('Coffee Sampling');
  var rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder('Coffee Sampling');

  var attachFolders = rootFolder.getFoldersByName('Attachments');
  var attachFolder = attachFolders.hasNext() ? attachFolders.next() : rootFolder.createFolder('Attachments');

  var sampleFolders = attachFolder.getFoldersByName(sampleId);
  return sampleFolders.hasNext() ? sampleFolders.next() : attachFolder.createFolder(sampleId);
}

function extractTextFromPDF(attachment) {
  try {
    var blob = attachment.copyBlob();
    var file = Drive.Files.insert(
      { title: blob.getName(), mimeType: 'application/pdf' },
      blob,
      { ocr: true, convert: true }
    );
    var doc = DocumentApp.openById(file.id);
    var text = doc.getBody().getText();
    Drive.Files.remove(file.id);
    return text;
  } catch (e) {
    Logger.log('Error extracting PDF: ' + e.toString());
    return '';
  }
}

function saveEmailPDFs(message, sampleId) {
  var saved = [];
  if (!message || !sampleId) return saved;
  try {
    var folder = getOrCreateSampleFolder(sampleId);
    var attachments = message.getAttachments();
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      if (att.getContentType() === 'application/pdf' || att.getName().toLowerCase().endsWith('.pdf')) {
        var fileName = sampleId + '_' + att.getName();
        var file = folder.createFile(att.copyBlob().setName(fileName));
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        saved.push({ name: att.getName(), url: file.getUrl(), size: att.getSize() });
      }
    }
  } catch (e) { Logger.log('saveEmailPDFs error: ' + e); }
  return saved;
}

function saveEmailImages(message, sampleId) {
  var saved = [];
  if (!message || !sampleId) return saved;
  try {
    var folder = getOrCreateSampleFolder(sampleId);
    var attachments = message.getAttachments();
    var imageTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/bmp'];
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      var contentType = att.getContentType();
      var name = att.getName().toLowerCase();
      if (imageTypes.indexOf(contentType) >= 0 ||
          name.endsWith('.jpg') || name.endsWith('.jpeg') || name.endsWith('.png') || name.endsWith('.gif')) {
        if (att.getSize() < 5000) continue;
        var fileName = sampleId + '_' + att.getName();
        var file = folder.createFile(att.copyBlob().setName(fileName));
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        saved.push({ name: att.getName(), url: file.getUrl(), size: att.getSize() });
      }
    }
  } catch (e) { Logger.log('saveEmailImages error: ' + e); }
  return saved;
}


// ============================================================================
// SECTION 16: UTILITY FUNCTIONS
// ============================================================================

function getKnownClientsList() {
  var list = [];
  for (var key in KNOWN_CLIENTS) {
    list.push({
      id: key,
      name: KNOWN_CLIENTS[key].name,
      domains: KNOWN_CLIENTS[key].domains.join(', '),
      defaultShipping: KNOWN_CLIENTS[key].defaultShipping,
      warehouse: KNOWN_CLIENTS[key].warehouse,
      isBroker: KNOWN_CLIENTS[key].isBroker || false,
      broker: KNOWN_CLIENTS[key].broker || null
    });
  }
  return list;
}

function getOriginalSender(message, emailBody) {
  if (!emailBody && message) {
    emailBody = message.getPlainBody();
  }
  if (!emailBody) return message ? message.getFrom() : '';

  var forwardPatterns = [
    /From:\s*([^<\n]+)\s*<?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})>?/i,
    /From:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i,
    /---------- Forwarded message ---------[\s\S]{0,200}From:\s*([^<\n]+)/i
  ];

  for (var fp = 0; fp < forwardPatterns.length; fp++) {
    var match = emailBody.match(forwardPatterns[fp]);
    if (match) {
      var sender = match[1].trim();
      sender = sender.replace(/^(From:|By:)\s*/i, '');
      if (sender.indexOf('@') >= 0) {
        var emailParts = sender.split('@')[0];
        sender = emailParts.replace(/[._-]/g, ' ');
        sender = sender.split(' ').map(function(word) {
          return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
        }).join(' ');
      }
      return sender;
    }
  }

  if (message) {
    var fromAddress = message.getFrom();
    var nameMatch = fromAddress.match(/^([^<]+)</);
    if (nameMatch) return nameMatch[1].trim();

    var emailMatch = fromAddress.match(/([a-zA-Z0-9._%+-]+)@/);
    if (emailMatch) {
      var namePart = emailMatch[1].replace(/[._-]/g, ' ');
      return namePart.split(' ').map(function(word) {
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      }).join(' ');
    }
    return fromAddress;
  }

  return '';
}

function cleanReceiverName(receiverName) {
  if (!receiverName) return '';
  receiverName = String(receiverName).trim();

  if (/^\d{4}\s+at\s+\d/i.test(receiverName)) return '';
  if (/^at\s+\d{1,2}:\d{2}/i.test(receiverName)) return '';
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(receiverName)) return '';
  if (/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)/i.test(receiverName)) return '';
  if (/^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d/i.test(receiverName)) return '';
  if (/^Date:/i.test(receiverName)) return '';
  if (/^From:/i.test(receiverName)) return '';
  if (/^Subject:/i.test(receiverName)) return '';

  receiverName = receiverName.replace(/<[^>]+>/g, '').replace(/[\w\.-]+@[\w\.-]+\.\w+/g, '');
  receiverName = receiverName.replace(/\s*Attn:.*$/i, '');
  receiverName = receiverName.replace(/\s*Attention:.*$/i, '');
  receiverName = receiverName.replace(/\s*C\/O.*$/i, '');
  receiverName = receiverName.replace(/\s*Mr\..*$/i, '');
  receiverName = receiverName.replace(/\s*Mrs\..*$/i, '');
  receiverName = receiverName.replace(/\s*Ms\..*$/i, '');

  var companyKeywords = ['Coffee', 'Trading', 'Corp', 'Corporation', 'Company', 'Inc', 'LLC', 'Ltd', 'Services', 'Roasters', 'Foods', 'Group'];
  var parts = receiverName.split(/[\n,]/);
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    for (var ck = 0; ck < companyKeywords.length; ck++) {
      if (part.indexOf(companyKeywords[ck]) >= 0 && part.length >= 5) {
        receiverName = part;
        break;
      }
    }
  }

  var words = receiverName.trim().split(/\s+/);
  if (words.length > 3) receiverName = words.slice(0, 3).join(' ');

  return receiverName.trim();
}

function escapeCSVField(field) {
  if (field == null || field === '') return '';
  var str = String(field);
  if (field instanceof Date) {
    str = Utilities.formatDate(field, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  if (str.indexOf(',') >= 0 || str.indexOf('"') >= 0 || str.indexOf('\n') >= 0) {
    str = '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function validateOrderData(data) {
  var errors = [];
  if (!data) return { valid: true, errors: [] };

  if (data.container) {
    var container = String(data.container).trim().toUpperCase();
    if (!/^[A-Z]{4}\d{7}$/.test(container)) {
      errors.push('Container format invalid: should be 4 letters + 7 digits');
    }
  }
  if (data.mark) {
    var mark = String(data.mark).trim();
    if (!/^\d{1,3}[\/\-]\d{2,5}[\/\-]\d{2,6}$/.test(mark)) {
      Logger.log('Mark format unusual: ' + mark);
    }
  }
  if (data.cargo) {
    var cargo = String(data.cargo).trim().toUpperCase();
    if (!/^C?\d{5,8}$/.test(cargo)) {
      errors.push('Cargo format unusual: expected C + 5-8 digits');
    }
  }
  if (data.sampleWeight) {
    var weight = String(data.sampleWeight).toLowerCase();
    if (weight && !/\d/.test(weight)) {
      errors.push('Sample weight should include a number');
    }
  }

  return { valid: errors.length === 0, errors: errors };
}

function checkForDuplicate(sheet, data) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { isDuplicate: false };

  var col = _getColumnMap(sheet);
  var numRows = Math.min(200, lastRow - 1);
  var startRow = lastRow - numRows + 1;
  var rows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  var newContainer = String(data.container || '').trim();
  var newMark = String(data.mark || '').trim();
  var newCargo = String(data.cargo || '').trim();
  var newReceiver = String(data.receiver || '').trim().toLowerCase();

  if (!newContainer && !newMark && !newCargo) return { isDuplicate: false };

  var receiverColIdx = col['Receiver'];

  for (var i = 0; i < rows.length; i++) {
    var rowContainer = String(rows[i][col['Container #']] || '').trim();
    var rowMark = String(rows[i][col['Mark #']] || '').trim();
    var rowCargo = String(rows[i][col['Cargo #']] || '').trim();
    var rowReceiver = (receiverColIdx !== undefined) ? String(rows[i][receiverColIdx] || '').trim().toLowerCase() : '';

    var isMatch = false;
    if (newContainer && newMark && rowContainer === newContainer && rowMark === newMark) {
      if (newReceiver === rowReceiver) isMatch = true;
    }
    else if (newContainer && newCargo && rowContainer === newContainer && rowCargo === newCargo) {
      if (newReceiver === rowReceiver) isMatch = true;
    }
    else if (newContainer && newMark && newCargo &&
             rowContainer === newContainer && rowMark === newMark && rowCargo === newCargo) {
      if (newReceiver === rowReceiver) isMatch = true;
    }

    if (isMatch) {
      return { isDuplicate: true, existingRow: startRow + i, existingSample: rows[i][col['CS Sample #']] };
    }
  }

  return { isDuplicate: false };
}
