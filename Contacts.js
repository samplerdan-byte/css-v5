// ============================================================
// Contacts.gs — Contact management, customer & warehouse lists
// Addresses verified via web search Feb 2026
// Entries still missing street addresses are marked with // NEEDS ADDRESS
// ============================================================

// NOTE: getCustomerList() and getWarehouseList() are defined in
// CustomerConfig.gs (richer versions that pull from KNOWN_CLIENTS
// and WAREHOUSES in EmailExtraction_v3.gs).
// Do NOT re-define them here — duplicates cause project-wide compile failure.

function setupContacts() {
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  var isNew = !sheet;

  if (isNew) {
    sheet = ss.insertSheet('Contacts');
    var headers = ['Type', 'Company', 'Attention', 'Address', 'City', 'State', 'Zip', 'Phone', 'Email', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
    var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['Shipper', 'Receiver']).build();
    sheet.getRange(2, 1, 500, 1).setDataValidation(typeRule);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 220);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 50);
    sheet.setColumnWidth(7, 80);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 180);
    sheet.setColumnWidth(10, 150);
  } else {
    var ui = SpreadsheetApp.getUi();
    var resp = ui.alert('Contacts sheet exists', 'Add industry contacts to existing sheet?\n(Will skip duplicates)', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  // Build list of existing company names for duplicate check
  var existingNames = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < existing.length; i++) {
      var type = String(existing[i][0] || '').trim();
      var company = String(existing[i][1] || '').trim().toLowerCase();
      if (type && company) {
        var key = type + '|' + company;
        existingNames[key] = true;
      }
    }
  }

  // Format: [Company, Attention, Address, City, State, Zip]

  // ── TRADING HOUSES / IMPORTERS (Shippers) ──
  var tradingHouses = [
    ['Louis Dreyfus Company',             'Cupping Room', '40 Danbury Rd',                       'Wilton',           'CT', '06897'],
    ['Olam Food Ingredients',             'Cupping Room', '800 Westchester Ave',                 'Rye Brook',        'NY', '10573'],
    ['Olam Specialty Coffee',             'Cupping Room', '800 Westchester Ave',                 'Rye Brook',        'NY', '10573'],
    ['Volcafe',                           'Cupping Room', '2 Bridge Ave, Ste 111',               'Red Bank',         'NJ', '07701'],
    ['Volcafe Specialty',                 'Cupping Room', '3 W Main St, Ste 203',                'Irvington',        'NY', '10533'],
    ['Sucafina NA',                       'Cupping Room', '1825 Park Ave, Ste 604',              'New York',         'NY', '10035'],
    ['Sucafina Specialty',                'Cupping Room', '1825 Park Ave, Ste 604',              'New York',         'NY', '10035'],
    ['ECOM Trading',                      'Cupping Room', '',                                     'East Rutherford',  'NJ', '07073'],
    ['Neumann Kaffee Gruppe (NKG)',       'Cupping Room', '111 River St, Ste 1220',              'Hoboken',          'NJ', '07030'],
    ['InterAmerican Coffee',              'Cupping Room', '111 River St, Ste 1220',              'Hoboken',          'NJ', '07030'],
    ['Atlas Coffee Importers',            'Cupping Room', '201 5th Ave S, Ste 201',              'Edmonds',          'WA', '98020'],
    ['Mercon Coffee Group',               'Cupping Room', '2333 Ponce De Leon Blvd, Ste 600',    'Coral Gables',     'FL', '33134'],
    ['Mercon Specialty',                  'Cupping Room', '2333 Ponce De Leon Blvd, Ste 600',    'Coral Gables',     'FL', '33134'],
    ['Paragon Coffee Trading',            'Cupping Room', '',                                     '',                 '',   ''],
    ['Atlantic USA / Atlantic Coffee',    'Cupping Room', '',                                     '',                 '',   ''],
    ['Serengeti Trading',                 'Cupping Room', '',                                     '',                 '',   ''],
    ['The Coffee Source',                 'Cupping Room', '',                                     '',                 '',   ''],
    ['Cafe Imports',                      'Cupping Room', '2617 E Hennepin Ave',                  'Minneapolis',      'MN', '55413'],
    ['Royal New York',                    'Cupping Room', '661 Hadley Rd',                        'South Plainfield', 'NJ', '07080'],
    ['Royal Coffee',                      'Cupping Room', '3306 Powell St',                       'Emeryville',       'CA', '94608'],
    ['Balzac Brothers',                   'Cupping Room', '',                                     '',                 '',   ''],
    ['Ally Coffee',                       'Cupping Room', '1801 Rutherford Rd, Ste 112',          'Greenville',       'SC', '29609'],
    ['Sustainable Harvest',               'Cupping Room', '',                                     'Portland',         'OR', ''],
    ['Walker Coffee Trading',             'Cupping Room', '',                                     'Houston',          'TX', ''],
    ['Crop to Cup Coffee',                'Cupping Room', '',                                     'Brooklyn',         'NY', ''],
    ['Bodhi Leaf Coffee Traders',         'Cupping Room', '',                                     'Long Beach',       'CA', ''],
    ['Covoya Coffee',                     'Cupping Room', '',                                     '',                 '',   ''],
    ['Westrock Coffee',                   'Cupping Room', '',                                     'Conway',           'AR', '72032'],
    ['Bernhardt Coffee',                  'Cupping Room', '',                                     '',                 '',   ''],
    ['Swiss Water Decaffeinated Coffee',  'Cupping Room', '3131 Lake City Way',                   'Burnaby',          'BC', 'V5A 3A3'],
    ['Coffee Holding Co.',                'Cupping Room', '3475 Victory Blvd',                    'Staten Island',    'NY', '10314'],
    ['Rothfos Corporation',               'Cupping Room', '111 River St, Ste 1220',              'Hoboken',          'NJ', '07030'],
    ['Commodities USA',                   'Cupping Room', '',                                     '',                 '',   ''],
    ['Comexim USA',                       'Cupping Room', '',                                     '',                 '',   ''],
    ['Coex Coffee International',         'Cupping Room', '',                                     '',                 '',   ''],
    ['ED&F Man',                          'Cupping Room', '',                                     '',                 '',   ''],
    ['StoneX Coffee',                     'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Annex Coffee Importers',            'Cupping Room', '',                                     '',                 '',   ''],
    ['Dupuy Storage',                     'Cupping Room', '7703 Cannon St',                       'Houston',          'TX', '77021'],
    ['Tropiq / Volcafe Tropiq',           'Cupping Room', '',                                     '',                 '',   ''],
    ['Zephyr Green Coffee',               'Cupping Room', '',                                     '',                 '',   ''],
    ['Artisan Coffee Imports',            'Cupping Room', '',                                     '',                 '',   ''],
    ['San Cristobal Coffee Importers',    'Cupping Room', '',                                     '',                 '',   ''],
    ['Red Fox Coffee Merchants',          'Cupping Room', '',                                     'Portland',         'OR', ''],
    ['Caravela Coffee',                   'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Genuine Origin',                    'Cupping Room', '3 W Main St, Ste 203',                'Irvington',        'NY', '10533'],
    ['Theta Ridge Coffee',                'Cupping Room', '',                                     '',                 '',   ''],
    ['Cedro Coffee',                      'Cupping Room', '',                                     '',                 '',   ''],
    ['Meru Coffee Farms',                 'Cupping Room', '',                                     '',                 '',   ''],
    ['Cup to Cup',                        'Cupping Room', '',                                     '',                 '',   ''],
    ['Vournas Coffee Trading',            'Cupping Room', '',                                     '',                 '',   ''],
    ['Coffee America USA',                'Cupping Room', '',                                     '',                 '',   ''],
    ['Keffa Coffee',                      'Cupping Room', '',                                     '',                 '',   ''],
    ['Moledina Commodities',              'Cupping Room', '',                                     '',                 '',   ''],
    ['Third Wave Coffee Source',          'Cupping Room', '',                                     '',                 '',   ''],
    ['Onyx Coffee Importers',             'Cupping Room', '',                                     '',                 '',   ''],
    ['Xavier Imports',                    'Cupping Room', '',                                     '',                 '',   ''],
    ['Copag International',               'Cupping Room', '',                                     '',                 '',   ''],
    ['Optco (Organic Products Trading Co.)','Cupping Room','',                                    '',                 '',   ''],
    ['Gold Mountain Coffee Growers',      'Cupping Room', '',                                     '',                 '',   ''],
    ['Torch Coffee Labs',                 'Cupping Room', '',                                     '',                 '',   '']
  ];

  // ── ROASTERS / RECEIVERS ──
  var roasters = [
    ['Keurig Dr Pepper / Green Mountain', 'Cupping Room', '33 Coffee Ln',                        'Waterbury',        'VT', '05676'],
    ['Starbucks Coffee Company',          'Cupping Room', '2401 Utah Ave S',                     'Seattle',          'WA', '98134'],
    ['Peets Coffee',                      'Cupping Room', '1400 Park Ave',                       'Emeryville',       'CA', '94608'],
    ['JDE Peets',                         'Cupping Room', '',                                     '',                 '',   ''],
    ['Folgers Coffee / Smucker\'s',       'Cupping Room', '1 Strawberry Ln',                     'Orrville',         'OH', '44667'],
    ['Nestle USA / Nespresso',            'Cupping Room', '',                                     '',                 '',   ''],
    ['Lavazza',                           'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Massimo Zanetti USA',               'Cupping Room', '',                                     'Suffolk',          'VA', '23434'],
    ['S&D Coffee (Westrock)',             'Cupping Room', '',                                     'Concord',          'NC', '28027'],
    ['Eight O\'Clock Coffee',             'Cupping Room', '',                                     'Montvale',         'NJ', '07645'],
    ['Community Coffee Company',          'Cupping Room', '',                                     'Baton Rouge',      'LA', ''],
    ['Stumptown Coffee Roasters',         'Cupping Room', '',                                     'Portland',         'OR', '97214'],
    ['Intelligentsia Coffee',             'Cupping Room', '',                                     'Chicago',          'IL', ''],
    ['Blue Bottle Coffee',                'Cupping Room', '',                                     'Oakland',          'CA', ''],
    ['Counter Culture Coffee',            'Cupping Room', '',                                     'Durham',           'NC', '27705'],
    ['La Colombe Coffee Roasters',        'Cupping Room', '',                                     'Philadelphia',     'PA', ''],
    ['Caribou Coffee',                    'Cupping Room', '',                                     'Minneapolis',      'MN', ''],
    ['Tim Hortons',                       'Cupping Room', '',                                     '',                 '',   ''],
    ['Dunkin\' (Inspire Brands)',          'Cupping Room', '130 Royall St',                       'Canton',           'MA', '02021'],
    ['Revere Coffee',                     'Cupping Room', '',                                     '',                 '',   ''],
    ['Stone Street Coffee',               'Cupping Room', '',                                     'Brooklyn',         'NY', ''],
    ['Porto Rico Importing Co.',          'Cupping Room', '201 Bleecker St',                      'New York',         'NY', '10012'],
    ['Gillies Coffee Company',            'Cupping Room', '',                                     'Brooklyn',         'NY', ''],
    ['F. Gavina & Sons',                  'Cupping Room', '2700 Fruitland Ave',                   'Vernon',           'CA', '90058'],
    ['Rogers Family Coffee',              'Cupping Room', '',                                     'Lincoln',          'CA', ''],
    ['Boyd Coffee Company',               'Cupping Room', '',                                     'Portland',         'OR', ''],
    ['Farmer Brothers',                   'Cupping Room', '',                                     'Northlake',        'TX', ''],
    ['Hawaiian Isles Kona Coffee',        'Cupping Room', '',                                     'Honolulu',         'HI', ''],
    ['Cameron\'s Coffee',                 'Cupping Room', '',                                     'Shakopee',         'MN', ''],
    ['Dillanos Coffee Roasters',          'Cupping Room', '',                                     'Sumner',           'WA', ''],
    ['Nordstrom Coffee',                  'Cupping Room', '',                                     '',                 '',   ''],
    ['Irving Farm Coffee Roasters',       'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Joe Coffee Company',                'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Birch Coffee',                      'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Toby\'s Estate Coffee',             'Cupping Room', '',                                     'Brooklyn',         'NY', ''],
    ['Partners Coffee',                   'Cupping Room', '',                                     'Brooklyn',         'NY', ''],
    ['Black Fox Coffee',                  'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Illy Caffe North America',          'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Lavazza Professional NA',           'Cupping Room', '',                                     '',                 '',   ''],
    ['Mother Parker\'s Coffee',           'Cupping Room', '',                                     'Fort Worth',       'TX', ''],
    ['Club Coffee',                       'Cupping Room', '',                                     '',                 '',   ''],
    ['Trung Nguyen USA',                  'Cupping Room', '',                                     '',                 '',   ''],
    ['Mayorga Organics',                  'Cupping Room', '',                                     'Rockville',        'MD', ''],
    ['Larry\'s Coffee',                   'Cupping Room', '',                                     'Raleigh',          'NC', ''],
    ['Equator Coffees',                   'Cupping Room', '',                                     'San Rafael',       'CA', ''],
    ['Verve Coffee Roasters',             'Cupping Room', '',                                     'Santa Cruz',       'CA', ''],
    ['Onyx Coffee Lab',                   'Cupping Room', '',                                     'Rogers',           'AR', ''],
    ['George Howell Coffee',              'Cupping Room', '',                                     'Acton',            'MA', ''],
    ['Klatch Coffee',                     'Cupping Room', '',                                     'Rancho Cucamonga', 'CA', ''],
    ['PT\'s Coffee Roasting',             'Cupping Room', '',                                     'Topeka',           'KS', ''],
    ['Conduit Coffee',                    'Cupping Room', '',                                     '',                 '',   ''],
    ['Cafe Grumpy',                       'Cupping Room', '',                                     'Brooklyn',         'NY', ''],
    ['Think Coffee',                      'Cupping Room', '',                                     'New York',         'NY', ''],
    ['Gasoline Alley Coffee',             'Cupping Room', '',                                     'New York',         'NY', '']
  ];

  // ── Seed existing customers as Shippers too ──
  var customers = typeof getCustomerList === 'function' ? getCustomerList() : [];
  for (var c = 0; c < customers.length; c++) {
    var alreadyListed = false;
    for (var t = 0; t < tradingHouses.length; t++) {
      if (tradingHouses[t][0] === customers[c].name) { alreadyListed = true; break; }
    }
    if (customers[c].key !== 'other' && !alreadyListed) {
      tradingHouses.push([customers[c].name, 'Cupping Room', '', '', '', '']);
    }
  }

  // ── Build batch of new rows, skipping duplicates ──
  var addedCount = 0;
  var newRows = [];

  function addContact(type, contactArr) {
    var company   = String(contactArr[0] || '').trim();
    var attention = String(contactArr[1] || '').trim();
    var address   = String(contactArr[2] || '').trim();
    var city      = String(contactArr[3] || '').trim();
    var state     = String(contactArr[4] || '').trim();
    var zip       = String(contactArr[5] || '').trim();

    if (!company) return; // Skip empty company names

    // Normalize all contact data before storage (trim and collapse whitespace)
    company = company.replace(/\s+/g, ' ');
    attention = attention.replace(/\s+/g, ' ');
    address = address.replace(/\s+/g, ' ');
    city = city.replace(/\s+/g, ' ');
    state = state.replace(/\s+/g, ' ');
    zip = zip.replace(/\s+/g, ' ');

    // Duplicate check must be case-insensitive
    var key = type + '|' + company.toLowerCase();
    if (existingNames[key]) return;
    existingNames[key] = true;
    newRows.push([type, company, attention, address, city, state, zip, '', '', '']);
    addedCount++;
  }

  for (var s = 0; s < tradingHouses.length; s++) addContact('Shipper', tradingHouses[s]);
  for (var r = 0; r < roasters.length; r++) addContact('Receiver', roasters[r]);

  // Write all new rows in a single batch call
  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, 10).setValues(newRows);
    Logger.log('setupContacts: batch-wrote ' + newRows.length + ' rows starting at row ' + startRow);
  }

  sheet.autoResizeColumn(2);

  Logger.log('setupContacts: ' + (isNew ? 'created' : 'updated') + ' sheet, added ' + addedCount + ' contacts');
  SpreadsheetApp.getUi().alert(
    'Contacts ' + (isNew ? 'sheet created' : 'updated') + '!\n\n' +
    addedCount + ' contacts added (' + tradingHouses.length + ' shippers, ' + roasters.length + ' receivers).\n\n' +
    'Edit the Contacts sheet directly to add addresses,\n' +
    'phone numbers, or additional contacts.\n' +
    'New contacts added via order entry are saved here automatically.'
  );
  } catch (e) {
    Logger.log('setupContacts ERROR: ' + e.message);
    SpreadsheetApp.getUi().alert('Error setting up Contacts: ' + e.message);
  }
}

function getContactsList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contacts');
    if (!sheet) return { shippers: [], receivers: [] };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { shippers: [], receivers: [] };

    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var shippers = [];
    var receivers = [];

    for (var i = 0; i < data.length; i++) {
      var type = String(data[i][0] || '').trim();
      var company = String(data[i][1] || '').trim();
      if (!company) continue;

      var contact = {
        company: company,
        attention: String(data[i][2] || '').trim(),
        address: String(data[i][3] || '').trim(),
        city: String(data[i][4] || '').trim(),
        state: String(data[i][5] || '').trim(),
        zip: String(data[i][6] || '').trim(),
        phone: String(data[i][7] || '').trim(),
        email: String(data[i][8] || '').trim(),
        notes: String(data[i][9] || '').trim()
      };

      if (type === 'Shipper') shippers.push(contact);
      else if (type === 'Receiver') receivers.push(contact);
    }

    shippers.sort(function(a, b) { return a.company.localeCompare(b.company); });
    receivers.sort(function(a, b) { return a.company.localeCompare(b.company); });

    Logger.log('getContactsList: ' + shippers.length + ' shippers, ' + receivers.length + ' receivers');
    return { shippers: shippers, receivers: receivers };
  } catch (e) {
    Logger.log('getContactsList ERROR: ' + e.message);
    return { shippers: [], receivers: [] };
  }
}

function saveNewContact(type, contactData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contacts');
    if (!sheet) {
      setupContacts();
      sheet = ss.getSheetByName('Contacts');
      if (!sheet) return { success: false, message: 'Could not create Contacts sheet' };
    }

    var company = String(contactData.company || '').trim();
    if (!company) return { success: false, message: 'Company name is required' };

    // Check for duplicate
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < existing.length; i++) {
        var existingType = String(existing[i][0] || '').trim();
        var existingCompany = String(existing[i][1] || '').trim().toLowerCase();
        if (existingType === type && existingCompany === company.toLowerCase()) {
          return { success: true, message: 'Contact already exists' };
        }
      }
    }

    // Normalize all contact data before storage
    var normalizeField = function(s) { return String(s || '').trim().replace(/\s+/g, ' '); };

    sheet.appendRow([
      type,
      normalizeField(company),
      normalizeField(contactData.attention),
      normalizeField(contactData.address),
      normalizeField(contactData.city),
      normalizeField(contactData.state),
      normalizeField(contactData.zip),
      normalizeField(contactData.phone),
      normalizeField(contactData.email),
      normalizeField(contactData.notes)
    ]);

    return { success: true, message: company + ' saved as ' + type };
  } catch (e) {
    return { success: false, message: 'Error saving contact: ' + e.message };
  }
}
