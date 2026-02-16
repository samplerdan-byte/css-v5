// ============================================================================
// ErrorLog.gs — Error Logging Utility
// Commodity Sampler Services
//
// Usage: logError('functionName', 'what went wrong', {optional: 'details'})
// ============================================================================


var ERROR_LOG_SHEET = 'Error Log';


// ============================================================================
// logError — Log an error to the Error Log sheet
// ============================================================================

function logError(functionName, errorMessage, details) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheetByName(ERROR_LOG_SHEET);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(ERROR_LOG_SHEET);
      sheet.appendRow(['Timestamp', 'Function', 'Error', 'Details', 'Resolved']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(2, 150);
      sheet.setColumnWidth(3, 300);
      sheet.setColumnWidth(4, 300);
      sheet.setColumnWidth(5, 80);
    }
    
    // Format details as string
    var detailsStr = '';
    if (details) {
      if (typeof details === 'object') {
        detailsStr = JSON.stringify(details);
      } else {
        detailsStr = String(details);
      }
    }
    
    // Append error row
    sheet.appendRow([
      new Date(),
      functionName || 'Unknown',
      errorMessage || 'No message',
      detailsStr,
      ''  // Resolved column - mark manually when fixed
    ]);
    
  } catch (e) {
    // Fallback to Logger if sheet logging fails
    Logger.log('ERROR LOG FAILED: ' + e.message);
    Logger.log('Original error: ' + functionName + ' - ' + errorMessage);
  }
}


// ============================================================================
// logWarning — Log a warning (non-critical issue)
// ============================================================================

function logWarning(functionName, message, details) {
  logError(functionName, '⚠️ WARNING: ' + message, details);
}


// ============================================================================
// getRecentErrors — Get errors from last N hours (for digest email)
// ============================================================================

function getRecentErrors(hours) {
  hours = hours || 24;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ERROR_LOG_SHEET);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  
  var data = sheet.getDataRange().getValues();
  var cutoff = new Date();
  cutoff.setHours(cutoff.getHours() - hours);
  
  var recent = [];
  for (var i = 1; i < data.length; i++) {
    var timestamp = new Date(data[i][0]);
    if (timestamp >= cutoff && !data[i][4]) {  // Not resolved
      recent.push({
        timestamp: timestamp,
        functionName: data[i][1],
        error: data[i][2],
        details: data[i][3],
        row: i + 1
      });
    }
  }
  
  return recent;
}


// ============================================================================
// clearOldErrors — Remove errors older than N days (run weekly)
// ============================================================================

function clearOldErrors(daysToKeep) {
  daysToKeep = daysToKeep || 30;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ERROR_LOG_SHEET);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return;
  }
  
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysToKeep);
  
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  
  for (var i = data.length - 1; i >= 1; i--) {
    var timestamp = new Date(data[i][0]);
    if (timestamp < cutoff) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // Delete from bottom up
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }
  
  Logger.log('Cleared ' + rowsToDelete.length + ' old errors');
}


// ============================================================================
// Test function
// ============================================================================

function testErrorLog() {
  logError('testErrorLog', 'This is a test error', { test: true, value: 123 });
  logWarning('testErrorLog', 'This is a test warning', 'Some details here');
  Logger.log('Check your Error Log sheet');
}