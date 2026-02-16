// ============================================================
// RUNTESTORDER.gs - Example orders for testing
// Creates 15 test samples across 3 warehouses, 4 customers,
// and 5 order types. Remove after testing.
// ============================================================

function createExampleOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Run Setup > Initial Setup first.');
    return;
  }

  var response = ui.alert(
    'Create Example Orders',
    'This will add 15 test samples across 3 warehouses, 4 customers, and 5 order types.\n\n' +
    'Features tested:\n' +
    '- Multiple receivers on one order\n' +
    '- All 5 sample types\n' +
    '- Shipping lines, notes, container statuses, ETAs\n' +
    '- Field Report grouping: Warehouse > Customer > Order > Receiver\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var orderCache = _buildOrderCache(sheet);
  var batchOpts = { orderCache: orderCache, skipLiveUpdate: true, skipAutoPrint: true };

  // -- ORDER 1: StoneX / Hersheys -- Container Supervision + FCC --
  var stoneXSamples = [
    { container: 'ONEU1298426', receiver: 'HERSHEYS', sampleType: 'Container Supervision' },
    { container: 'ONEU1298426', receiver: 'HERSHEYS', sampleType: 'FCC Grading-Cocoa' },
    { container: 'ONEU1443191', receiver: 'HERSHEYS', sampleType: 'Container Supervision' },
    { container: 'ONEU1443191', receiver: 'HERSHEYS', sampleType: 'FCC Grading-Cocoa' }
  ];

  stoneXSamples.forEach(function(s) {
    addDataToSheet(sheet, {
      sender: 'STONEX COMMODITY SOLUTIONS, LLC',
      receiver: s.receiver,
      warehouse: 'Dependable-59th',
      description: 'ECUADOR COCOA',
      sampleOrderNum: 'SCS_173596',
      cargo: '',
      mark: '',
      container: s.container,
      reference: 'SCS_173596',
      bagCount: '362',
      weight: '',
      sampleWeight: '10 LB',
      pNumber: '',
      sNumber: '',
      shippingProcess: s.sampleType,
      comments: 'Example order - multi-type',
      sourceEmail: 'example@test.com',
      sampleType: s.sampleType,
      containerStatus: 'Discharged At Pier',
      containerETA: '02/15/2026',
      shippingLine: 'ONE',
      shippingNotes: 'GYEG00859900'
    }, batchOpts);
  });

  // -- ORDER 2: InterAmerican / Multiple receivers --
  var iacSamples = [
    { container: 'HLBU3479463', mark: '002/3220/0608', receiver: 'LA COLOMBE', desc: 'BRAZIL ARABICA NY 2/3 SCR 15/16 SSFC', ref: 'P200451#1' },
    { container: 'SMLU3122632', mark: '013/0055/0187', receiver: 'FERRIS COFFEE & NUT COMPANY, INC.', desc: 'HONDURAS ARABICA HG EP', ref: 'P200309#1' },
    { container: 'SMLU2591764', mark: '013/0055/0174', receiver: 'INTERAMERICAN COFFEE', desc: 'HONDURAS ARABICA SHG EP', ref: 'P200559#1' }
  ];

  iacSamples.forEach(function(s) {
    addDataToSheet(sheet, {
      sender: 'INTERAMERICAN COFFEE',
      receiver: s.receiver,
      warehouse: 'Continental NJ',
      description: s.desc,
      sampleOrderNum: '212113',
      cargo: '377' + Math.floor(Math.random() * 999),
      mark: s.mark,
      container: s.container,
      reference: s.ref,
      bagCount: '275',
      weight: '',
      sampleWeight: '4 LB',
      pNumber: s.ref,
      sNumber: '',
      shippingProcess: 'Warehouse Sample',
      comments: 'Example order - multi-receiver',
      sourceEmail: 'example@test.com',
      sampleType: 'Warehouse Sample',
      containerStatus: 'Afloat',
      containerETA: '02/11/2026',
      shippingLine: 'HAPAG-LLOYD',
      shippingNotes: ''
    }, batchOpts);
  });

  // -- ORDER 3: Cooperative Coffee -- Photos With Sample --
  var coopSamples = [
    { container: 'RFSU3052052', mark: '028/0007/24005', desc: 'RWANDA KIGEYO KODUKAK FTO', cargo: '376597', ref: 'RWC2501' },
    { container: 'RFSU3052052', mark: '028/0011/25015', desc: 'RWANDA MICROLOT GAKENKE HINGA KAWA WASHED FTO', cargo: '376598', ref: 'MLRA2501' }
  ];

  coopSamples.forEach(function(s) {
    addDataToSheet(sheet, {
      sender: 'COOPERATIVE COFFEE',
      receiver: 'COOPERATIVE COFFEE',
      warehouse: 'Continental NJ',
      description: s.desc,
      sampleOrderNum: '211622',
      cargo: s.cargo,
      mark: s.mark,
      container: s.container,
      reference: s.ref,
      bagCount: '240',
      weight: '',
      sampleWeight: '',
      pNumber: s.ref,
      sNumber: '',
      shippingProcess: 'Photos With Sample',
      comments: 'Example order - photos with sample',
      sourceEmail: 'example@test.com',
      sampleType: 'Photos With Sample',
      containerStatus: 'Picked Up By Warehouse',
      containerETA: '02/06/2026',
      shippingLine: 'MSC',
      shippingNotes: ''
    }, batchOpts);
  });

  // -- ORDER 4: COFCO / Exchange Samples at Cadeco Houston --
  var exchangeSamples = [
    { container: 'CAAU4734634', mark: 'KCEWR-CAHO-9674-0', cargo: '51', ref: '013/273/0113' },
    { container: 'ECMU5254704', mark: 'KCEWR-CAHO-9675-0', cargo: '52', ref: '013/273/0112' },
    { container: 'ECMU5254957', mark: 'KCEWR-CAHO-9676-0', cargo: '53', ref: '013/273/0114' }
  ];

  exchangeSamples.forEach(function(s) {
    addDataToSheet(sheet, {
      sender: 'COFCO AGRI-CT',
      receiver: 'ICE FUTURES US',
      warehouse: 'Cadeco Ind. Houston',
      description: 'Coffee Exchange Samples',
      sampleOrderNum: '212414',
      cargo: s.cargo,
      mark: s.mark,
      container: s.container,
      reference: s.ref,
      bagCount: '333',
      weight: '',
      sampleWeight: '5 LB',
      pNumber: '',
      sNumber: '',
      shippingProcess: 'Exchange Samples',
      comments: 'Example order - exchange samples',
      sourceEmail: 'example@test.com',
      sampleType: 'Exchange Samples',
      containerStatus: '',
      containerETA: '',
      shippingLine: '',
      shippingNotes: ''
    }, batchOpts);
  });

  // -- ORDER 5: Olam / RPM Unassigned -- multi-receiver same order --
  var olamSamples = [
    { container: 'MRSU0391982', mark: '003/1832/1068', receiver: 'COFFEE AMERICA', ref: 'P07084.005' },
    { container: 'TIIU2080777', mark: '003/1832/1069', receiver: 'COFFEE AMERICA', ref: 'P07084.005' },
    { container: 'MRKU9562661', mark: '016/2380/0001', receiver: 'REILY FOODS COMPANY', ref: 'S07348.001' }
  ];

  olamSamples.forEach(function(s) {
    addDataToSheet(sheet, {
      sender: 'OLAM AMERICA',
      receiver: s.receiver,
      warehouse: 'RPM Unassigned',
      description: 'COLOMBIAN ARABICA',
      sampleOrderNum: '212388',
      cargo: '',
      mark: s.mark,
      container: s.container,
      reference: s.ref,
      bagCount: '275',
      weight: '',
      sampleWeight: '2 lbs.',
      pNumber: s.ref,
      sNumber: '',
      shippingProcess: 'Warehouse Sample',
      comments: 'Example order - multi receiver same order',
      sourceEmail: 'example@test.com',
      sampleType: 'Warehouse Sample',
      containerStatus: 'Discharged At Pier',
      containerETA: '02/05/2026',
      shippingLine: 'MAERSK',
      shippingNotes: ''
    }, batchOpts);
  });

  updateLiveOrdersView();

  ui.alert(
    'Example Orders Created',
    '15 samples added across:\n\n' +
    'Dependable-59th - StoneX/Hersheys (Container Supervision + FCC)\n' +
    'Continental NJ - InterAmerican (multi-receiver) + Coop Coffee (Photos)\n' +
    'Cadeco Ind. Houston - COFCO (Exchange Samples)\n' +
    'RPM Unassigned - Olam (multi-receiver same order)\n\n' +
    'Now try:\n' +
    '- Reports > Field Report\n' +
    '- Check the sheet for new columns\n\n' +
    'Delete these rows when done testing.',
    ui.ButtonSet.OK
  );
}
