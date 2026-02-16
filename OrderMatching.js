// ============================================================================
// OrderMatching.gs — Order Matching & Duplicate Detection
// Commodity Sampler Services
//
// FIXED Feb 15, 2026:
//   - Changed all column references to match actual sheet headers:
//     'Container' → 'Container #', 'Mark' → 'Mark #',
//     'Customer'/'Client' → 'Sender', 'Date Received' → 'Timestamp',
//     'Bags' → 'Bag Count', 'Notes' → 'Comments'
//   - Removed MATCH_CONFIG.spreadsheetId (CONFIG doesn't define it;
//     now uses getActiveSpreadsheet() instead)
//   - Fixed sheetName from 'Live Orders' to 'All Orders' (matching
//     against Live Orders misses shipped/completed rows)
// ============================================================================


// ============================================================================
// CONFIGURATION — Uses same sheet names as Code.gs CONFIG
// ============================================================================

var MATCH_CONFIG = {
  sheetName: 'All Orders',                        // Search all orders, not just live
  
  // Matching thresholds
  containerMatchWeight: 30,
  markMatchWeight: 25,
  cargoMatchWeight: 20,
  clientMatchWeight: 15,
  dateProximityWeight: 10,
  
  minMatchScore: 60,                              // Minimum score to consider a match
  dateProximityDays: 7                            // Look within 7 days for matches
};


// ============================================================================
// findMatchingOrder — Look for existing order that matches new data
// Returns { found: boolean, row: number, score: number, existingData: object }
// ============================================================================

function findMatchingOrder(newOrder) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MATCH_CONFIG.sheetName);
  
  if (!sheet) {
    Logger.log('OrderMatching: Sheet not found: ' + MATCH_CONFIG.sheetName);
    return { found: false, row: -1, score: 0, existingData: null };
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { found: false, row: -1, score: 0, existingData: null };
  }
  
  var colMap = _getColumnMap(sheet);
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var bestMatch = { found: false, row: -1, score: 0, existingData: null };
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var score = 0;
    
    // Container match (strongest signal)
    if (newOrder.container && colMap['Container #'] !== undefined &&
        String(row[colMap['Container #']]).trim() === String(newOrder.container).trim()) {
      score += MATCH_CONFIG.containerMatchWeight;
    }
    
    // Mark match
    if (newOrder.mark && colMap['Mark #'] !== undefined &&
        String(row[colMap['Mark #']]).trim() === String(newOrder.mark).trim()) {
      score += MATCH_CONFIG.markMatchWeight;
    }
    
    // Cargo match
    if (newOrder.cargo && colMap['Cargo #'] !== undefined &&
        String(row[colMap['Cargo #']]).trim() === String(newOrder.cargo).trim()) {
      score += MATCH_CONFIG.cargoMatchWeight;
    }
    
    // Client/Sender match
    var existingSender = colMap['Sender'] !== undefined ? String(row[colMap['Sender']]).trim() : '';
    if (newOrder.client && existingSender &&
        existingSender.toLowerCase().indexOf(String(newOrder.client).toLowerCase()) >= 0) {
      score += MATCH_CONFIG.clientMatchWeight;
    }
    
    // Date proximity (using Timestamp column)
    if (colMap['Timestamp'] !== undefined && row[colMap['Timestamp']] && newOrder.dateReceived) {
      try {
        var existingDate = new Date(row[colMap['Timestamp']]);
        var newDate = new Date(newOrder.dateReceived);
        if (!isNaN(existingDate.getTime()) && !isNaN(newDate.getTime())) {
          var daysDiff = Math.abs((newDate - existingDate) / (1000 * 60 * 60 * 24));
          if (daysDiff <= MATCH_CONFIG.dateProximityDays) {
            score += MATCH_CONFIG.dateProximityWeight * (1 - daysDiff / MATCH_CONFIG.dateProximityDays);
          }
        }
      } catch (e) {
        // Skip date comparison on error
      }
    }
    
    // Track best match
    if (score > bestMatch.score && score >= MATCH_CONFIG.minMatchScore) {
      bestMatch = {
        found: true,
        row: i + 2,  // 1-indexed for sheet operations
        score: score,
        existingData: {
          container: colMap['Container #'] !== undefined ? String(row[colMap['Container #']]) : '',
          mark: colMap['Mark #'] !== undefined ? String(row[colMap['Mark #']]) : '',
          cargo: colMap['Cargo #'] !== undefined ? String(row[colMap['Cargo #']]) : '',
          sender: existingSender,
          csOrder: colMap['CS Order #'] !== undefined ? String(row[colMap['CS Order #']]) : ''
        }
      };
    }
  }
  
  return bestMatch;
}


// ============================================================================
// updateExistingOrder — Update an existing order row with new data
// Only updates fields that were empty before
// ============================================================================

function updateExistingOrder(rowNum, newData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MATCH_CONFIG.sheetName);
  
  if (!sheet || rowNum < 2) {
    return { success: false, message: 'Invalid sheet or row' };
  }
  
  var colMap = _getColumnMap(sheet);
  var range = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
  var existing = range.getValues()[0];
  var updated = false;
  
  // Map newData keys to actual sheet column names
  var fieldMapping = {
    'container': 'Container #',
    'mark': 'Mark #',
    'cargo': 'Cargo #',
    'bags': 'Bag Count',
    'reference': 'Reference',
    'warehouse': 'Warehouse'
  };
  
  for (var dataKey in fieldMapping) {
    var colName = fieldMapping[dataKey];
    var col = colMap[colName];
    
    if (col !== undefined && newData[dataKey] && !existing[col]) {
      sheet.getRange(rowNum, col + 1).setValue(newData[dataKey]);
      updated = true;
    }
  }
  
  // Always append to Comments
  if (newData.notes) {
    var commentsCol = colMap['Comments'];
    if (commentsCol !== undefined) {
      var existingComments = existing[commentsCol] || '';
      var newComments = existingComments + (existingComments ? '\n' : '') + newData.notes;
      sheet.getRange(rowNum, commentsCol + 1).setValue(newComments);
      updated = true;
    }
  }
  
  return { success: updated, message: updated ? 'Order updated' : 'No updates needed' };
}


// ============================================================================
// processEmailWithMatching — Process email and check for matches
// Wrapper that can be integrated into main processing pipeline
// ============================================================================

function processEmailWithMatching(emailData, extractedOrder) {
  var match = findMatchingOrder(extractedOrder);
  
  if (match.found) {
    Logger.log('Found matching order at row ' + match.row + ' with score ' + match.score);
    Logger.log('Existing: ' + JSON.stringify(match.existingData));
    
    var updateResult = updateExistingOrder(match.row, {
      container: extractedOrder.container,
      mark: extractedOrder.mark,
      cargo: extractedOrder.cargo,
      bags: extractedOrder.bags,
      reference: extractedOrder.reference,
      warehouse: extractedOrder.warehouse,
      notes: 'Updated from email: ' + emailData.subject + ' (' + new Date().toLocaleDateString() + ')'
    });
    
    return {
      action: 'updated',
      row: match.row,
      score: match.score,
      existingCSOrder: match.existingData.csOrder,
      updateResult: updateResult
    };
  }
  
  return {
    action: 'insert',
    row: -1,
    score: 0,
    existingCSOrder: null
  };
}


// ============================================================================
// Test function
// ============================================================================

function testOrderMatching() {
  var testOrder = {
    container: 'MSKU1234567',
    mark: '003/0123/00456',
    cargo: 'C123456',
    client: 'Louis Dreyfus',
    dateReceived: new Date()
  };
  
  var result = findMatchingOrder(testOrder);
  Logger.log('Match result: ' + JSON.stringify(result));
}
