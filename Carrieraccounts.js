// ============================================================================
// CarrierAccounts.gs - FedEx/UPS Account Management
// Commodity Sampler Services
// Transcribed from physical clipboard lists (Updated 07/31/2023 + handwritten additions)
// CSS FedEx Account: 127086486 | CSS UPS Account: 07A48A
// ============================================================================

// ============================================================
// STEP 1: Add carrier columns to Contacts sheet (run once)
// ============================================================

function addCarrierColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Contacts sheet not found. Run setupContacts() first.');
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Check if columns already exist
  if (headers.indexOf('FedEx Account') >= 0) {
    SpreadsheetApp.getUi().alert('Carrier columns already exist.');
    return;
  }
  
  var nextCol = headers.length + 1;
  sheet.getRange(1, nextCol).setValue('FedEx Account').setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
  sheet.getRange(1, nextCol + 1).setValue('UPS Account').setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
  sheet.getRange(1, nextCol + 2).setValue('Preferred Carrier').setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
  
  sheet.setColumnWidth(nextCol, 130);
  sheet.setColumnWidth(nextCol + 1, 110);
  sheet.setColumnWidth(nextCol + 2, 120);
  
  // Add dropdown validation for Preferred Carrier
  var carrierRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['FedEx', 'UPS', 'USPS', 'DHL', 'FedEx Ground', 'UPS Ground'])
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, nextCol + 2, 500, 1).setDataValidation(carrierRule);
  
  SpreadsheetApp.getUi().alert('✅ Added FedEx Account, UPS Account, and Preferred Carrier columns to Contacts sheet.');
}


// ============================================================
// STEP 2: Populate carrier accounts from clipboard data (run once)
// ============================================================

function populateCarrierAccounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet) { SpreadsheetApp.getUi().alert('Contacts sheet not found.'); return; }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var fedexCol = headers.indexOf('FedEx Account');
  var upsCol = headers.indexOf('UPS Account');
  var preferredCol = headers.indexOf('Preferred Carrier');
  
  if (fedexCol === -1 || upsCol === -1) {
    SpreadsheetApp.getUi().alert('Run addCarrierColumns() first.');
    return;
  }
  
  // ── FedEx Account Numbers (from clipboard photo) ──
  var fedexAccounts = [
    ['Ally Brazil', '280910333'],
    ['Agora Coffee', '285570611'],
    ['American', '160463996'],
    ['Armenia', '010086710'],
    ['Atlas Coffee', '255050397'],
    ['Atlantic USA', '122636097'],
    ['Atlantic Cocoa', '94150140'],
    ['Atlantic Specialty', '201238501'],
    ['Balzac Bros', '131987005'],
    ['Brauner', '10045894'],
    ['Caribou', '178424157'],
    ['Cape Horn Coffee', '202701116'],
    ['Cascade', '227528400'],
    ['Cardiff', '812095005'],
    ['Caturra', '231117334'],
    ['Cofco Agri', '802808747'],
    ['Coffee America', '670146952'],
    ['Covoya', '295818891'],
    ['Excelco', '670146952'],
    ['Falcon', '660151451'],
    ['Finagra', '837552621'],
    ['Globus', '200392639'],
    ['Global Coffee Trading', '673603181'],
    ['Monkey Bean', '190611760'],
    ['InterAmerican', '372518081'],
    ['ICT', '461992226'],
    ['ICC', '070111110'],
    ['Kencaf', '162548824'],
    ['Ken Gabbay', '257162958'],
    ['La Minita', '139794877'],
    ['List & Beisler', '763817172'],
    ['Louis Dreyfus', '281140086'],
    ['MCT USA', '678557358'],
    ['Mass Commodities', '468489988'],
    ['Melitta', '019104737'],
    ['Mercon', '104528660'],
    ['Mitsui', '112689125'],
    ['Mother Parker', '123217462'],
    ['NJ Douek', '124906806'],
    ['Olam Americas', '303421262'],
    ['Olam Canada', '820209815'],
    ['OPTCO', '156405574'],
    ['Osito Coffee', '625178762'],
    ['Pacific Blue', '484345066'],
    ['Pan American Costa Rica', '329393011'],
    ['Paragon Trading', '134765364'],
    ['Red Goni', '121600897'],
    ['RGC Coffee', '474826325'],
    ['Rothfos', '155854537'],
    ['Royal NY', '192984147'],
    ['Royal Pacific', '145752086'],
    ['S&D Coffee', '303127135'],
    ['Serengeti Trading', '238805902'],
    ['Sucafina', '245009941'],
    ['Sunwest', '120202391'],
    ['Swiss Water', '361265726'],
    ['T.K Company', '307324326'],
    ['Touton USA', '442134146'],
    ['Trabocca', '490037845'],
    ['The Coffee Source', '946265551'],
    ['The Salvage Group', '231944648'],
    ['Tristao', '112183922'],
    ['Volcafe Spec', '235952823'],
    ['Volcafe USA', '206201940'],
    ['Volcafe Swiss', '285570611'],
    ['White Coffee', '111633266'],
    ['Zephyr', '281140086'],
    ['Expocacer', '127086486'],
    ['Age of Coffee', '5671086124'],
    ['Yellow Rooster', '2069448628'],
    ['ONE ACRE Fund', '650600673'],
    ['Sopex', '200562811'],
    ['Sustainable Harvest', '363806341']
  ];
  
  // ── UPS Account Numbers (from clipboard photo) ──
  var upsAccounts = [
    ['32 Cup Spec', '17VE28'],
    ['All World', '660AX3'],
    ['American', '8A5E34'],
    ['Armenia', 'F038R6'],
    ['Atlantic Spec', '442EE1'],
    ['Atlantic USA', 'F0V832'],
    ['Atlas Coffee Import', '8E156Y'],
    ['Cafe Zinho', '89W54Y'],
    ['Cardiff', '194Y9A'],
    ['Coffee Holding', '58764V'],
    ['Collaborative', '0F59X8'],
    ['Cooperative', '04R1V4'],
    ['Crop to Cup', '781AR7'],
    ['Equal Exchange', '664404'],
    ['Excelco', 'W4X495'],
    ['Exotic Green', '2501F6'],
    ['Kraft Foods', 'V11021'],
    ['ICC', '7159E2'],
    ['ICT', '3W8F69'],
    ['LJ Cooper', '08F495'],
    ['List & Beisler', '0846RF'],
    ['Magnum', '8E37E8'],
    ['Mercon', 'R0188X'],
    ['Mercon Specialty', '8E1E56'],
    ['Olam America', 'V56E57'],
    ['OmniActive', 'F9898Y'],
    ['Opal Coffee', '75E5X0'],
    ['OPTCO', '410EW8'],
    ['Paragon Trading', '2E43W5'],
    ['Rothfos', '11776E'],
    ['Pro Coffee', '5003V3'],
    ['OCAFI', 'H7R311'],
    ['Cafe Lobo', 'W60736'],
    ['Red Fox', 'RA6886'],
    ['Royal NY', 'W021X1'],
    ['S&D Coffee', 'W1808Y'],
    ['Sustainable Harvest', '11A8E5'],
    ['Urban Shade', '63W0V0'],
    ['Volcafe Spec', 'RW4350'],
    ['Walker', '9VR018'],
    ['Westfeldt Bros', '010233'],
    ['Stonex', 'J290X5'],
    ['Yellow Rooster', '81918'],
    ['Higa', 'FB2443'],
    ['Cafe Kreyol', 'HE7436'],
    ['DHL Shared', 'J26607']
  ];
  
  // Read all contacts
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No contacts to update.'); return; }
  
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var companyCol = 1; // 0-indexed column B
  
  // Build lookup: lowercase company name -> row index
  var companyIndex = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][companyCol]).trim().toLowerCase();
    if (name) {
      if (!companyIndex[name]) companyIndex[name] = [];
      companyIndex[name].push(i);
    }
  }
  
  var matched = 0;
  var unmatched = [];
  var added = 0;
  
  // ── Match helper: try exact, then contains, then fuzzy ──
  function findContactRows(clipName) {
    var lower = clipName.toLowerCase().trim();
    
    // Exact match
    if (companyIndex[lower]) return companyIndex[lower];
    
    // Contains match (either direction)
    for (var key in companyIndex) {
      if (key.indexOf(lower) >= 0 || lower.indexOf(key) >= 0) {
        return companyIndex[key];
      }
    }
    
    // Partial word match (first significant word)
    var words = lower.split(/[\s\/&,]+/).filter(function(w) { return w.length > 3; });
    for (var w = 0; w < words.length; w++) {
      for (var key2 in companyIndex) {
        if (key2.indexOf(words[w]) >= 0) {
          return companyIndex[key2];
        }
      }
    }
    
    return null;
  }
  
  // ── Apply FedEx accounts ──
  for (var f = 0; f < fedexAccounts.length; f++) {
    var name = fedexAccounts[f][0];
    var acct = fedexAccounts[f][1];
    var rows = findContactRows(name);
    if (rows) {
      for (var r = 0; r < rows.length; r++) {
        var rowNum = rows[r] + 2; // 1-indexed, skip header
        var existing = sheet.getRange(rowNum, fedexCol + 1).getValue();
        if (!existing) {
          sheet.getRange(rowNum, fedexCol + 1).setValue(acct);
          matched++;
        }
      }
    } else {
      unmatched.push('FedEx: ' + name + ' = ' + acct);
      // Add as new contact
      var newRow = [
        'Receiver', name, '', '', '', '', '', '', '', '',
      ];
      // Pad to reach FedEx column
      while (newRow.length < fedexCol) newRow.push('');
      newRow.push(acct); // FedEx
      newRow.push('');    // UPS
      newRow.push('');    // Preferred
      sheet.appendRow(newRow);
      added++;
      // Update index
      var newIdx = sheet.getLastRow() - 2;
      companyIndex[name.toLowerCase()] = [newIdx];
    }
  }
  
  // ── Apply UPS accounts ──
  for (var u = 0; u < upsAccounts.length; u++) {
    var name2 = upsAccounts[u][0];
    var acct2 = upsAccounts[u][1];
    var rows2 = findContactRows(name2);
    if (rows2) {
      for (var r2 = 0; r2 < rows2.length; r2++) {
        var rowNum2 = rows2[r2] + 2;
        var existing2 = sheet.getRange(rowNum2, upsCol + 1).getValue();
        if (!existing2) {
          sheet.getRange(rowNum2, upsCol + 1).setValue(acct2);
          matched++;
        }
      }
    } else {
      unmatched.push('UPS: ' + name2 + ' = ' + acct2);
      var newRow2 = ['Receiver', name2, '', '', '', '', '', '', '', ''];
      while (newRow2.length < fedexCol) newRow2.push('');
      newRow2.push('');    // FedEx
      newRow2.push(acct2); // UPS
      newRow2.push('');    // Preferred
      sheet.appendRow(newRow2);
      added++;
      companyIndex[name2.toLowerCase()] = [sheet.getLastRow() - 2];
    }
  }
  
  // ── Auto-set preferred carrier based on what's available ──
  var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var p = 0; p < allData.length; p++) {
    var hasFedex = String(allData[p][fedexCol] || '').trim();
    var hasUps = String(allData[p][upsCol] || '').trim();
    var hasPref = String(allData[p][preferredCol] || '').trim();
    if (!hasPref && (hasFedex || hasUps)) {
      // Default to FedEx if they have it, otherwise UPS
      sheet.getRange(p + 2, preferredCol + 1).setValue(hasFedex ? 'FedEx' : 'UPS');
    }
  }
  
  var msg = '✅ Carrier accounts populated!\n\n' +
    'Matched & updated: ' + matched + ' entries\n' +
    'New contacts added: ' + added + '\n';
  if (unmatched.length > 0) {
    msg += '\nAdded as new contacts (verify names):\n' + unmatched.slice(0, 15).join('\n');
    if (unmatched.length > 15) msg += '\n... and ' + (unmatched.length - 15) + ' more';
  }
  
  SpreadsheetApp.getUi().alert(msg);
}


// ============================================================
// CARRIER LOOKUP (used by scan-out + coversheets)
// ============================================================

function lookupCarrierByCompany(companyName) {
  if (!companyName) return null;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet || sheet.getLastRow() < 2) return null;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var fedexCol = headers.indexOf('FedEx Account');
  var upsCol = headers.indexOf('UPS Account');
  var preferredCol = headers.indexOf('Preferred Carrier');
  if (fedexCol === -1 && upsCol === -1) return null;
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var searchName = String(companyName).trim().toLowerCase();
  
  // Try exact match on company name (col index 1)
  for (var i = 0; i < data.length; i++) {
    var company = String(data[i][1]).trim().toLowerCase();
    if (company === searchName) {
      return _buildCarrierResult(data[i], fedexCol, upsCol, preferredCol);
    }
  }
  
  // Try contains match
  for (var j = 0; j < data.length; j++) {
    var company2 = String(data[j][1]).trim().toLowerCase();
    if (company2.indexOf(searchName) >= 0 || searchName.indexOf(company2) >= 0) {
      return _buildCarrierResult(data[j], fedexCol, upsCol, preferredCol);
    }
  }
  
  // Try word-level match (for partial names like "Volcafe" matching "Volcafe Specialty")
  var words = searchName.split(/[\s\/&,\-]+/).filter(function(w) { return w.length > 3; });
  if (words.length > 0) {
    for (var k = 0; k < data.length; k++) {
      var company3 = String(data[k][1]).trim().toLowerCase();
      for (var w = 0; w < words.length; w++) {
        if (company3.indexOf(words[w]) >= 0) {
          var result = _buildCarrierResult(data[k], fedexCol, upsCol, preferredCol);
          if (result && (result.fedex || result.ups)) return result;
        }
      }
    }
  }
  
  return null;
}

function _buildCarrierResult(row, fedexCol, upsCol, preferredCol) {
  var fedex = fedexCol >= 0 ? String(row[fedexCol] || '').trim() : '';
  var ups = upsCol >= 0 ? String(row[upsCol] || '').trim() : '';
  var preferred = preferredCol >= 0 ? String(row[preferredCol] || '').trim() : '';
  if (!fedex && !ups) return null;
  return { fedex: fedex, ups: ups, preferred: preferred, company: String(row[1]).trim() };
}


// ============================================================
// UPDATED getContactsList (reads carrier columns)
// Call this instead of the old version in Main.gs
// ============================================================

function getContactsListWithCarrier() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet || sheet.getLastRow() < 2) return { shippers: [], receivers: [] };
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var fedexCol = headers.indexOf('FedEx Account');
  var upsCol = headers.indexOf('UPS Account');
  var preferredCol = headers.indexOf('Preferred Carrier');
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var shippers = [];
  var receivers = [];
  
  for (var i = 0; i < data.length; i++) {
    var type = String(data[i][0]).trim();
    var company = String(data[i][1]).trim();
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
      notes: String(data[i][9] || '').trim(),
      fedex: fedexCol >= 0 ? String(data[i][fedexCol] || '').trim() : '',
      ups: upsCol >= 0 ? String(data[i][upsCol] || '').trim() : '',
      preferredCarrier: preferredCol >= 0 ? String(data[i][preferredCol] || '').trim() : ''
    };
    
    if (type === 'Shipper') shippers.push(contact);
    else if (type === 'Receiver') receivers.push(contact);
  }
  
  shippers.sort(function(a, b) { return a.company.localeCompare(b.company); });
  receivers.sort(function(a, b) { return a.company.localeCompare(b.company); });
  
  return { shippers: shippers, receivers: receivers };
}