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
    if (!ss) {
      Logger.log('ERROR LOG FAILED: No active spreadsheet');
      return;
    }

    var sheet = ss.getSheetByName(ERROR_LOG_SHEET);

    // Create sheet if it doesn't exist
    if (!sheet) {
      try {
        sheet = ss.insertSheet(ERROR_LOG_SHEET);
        sheet.appendRow(['Timestamp', 'Function', 'Error', 'Details', 'Resolved']);
        sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
        sheet.setColumnWidth(1, 150);
        sheet.setColumnWidth(2, 150);
        sheet.setColumnWidth(3, 300);
        sheet.setColumnWidth(4, 300);
        sheet.setColumnWidth(5, 80);
      } catch (e) {
        Logger.log('ERROR LOG FAILED to create sheet: ' + e.message);
        return;
      }
    }

    // Format details as string — JSON.stringify is wrapped separately so a
    // circular-reference or non-serializable value doesn't swallow the error.
    var detailsStr = '';
    if (details) {
      if (typeof details === 'object') {
        try {
          detailsStr = JSON.stringify(details);
        } catch (jsonErr) {
          detailsStr = '[unserializable: ' + String(details) + ']';
        }
      } else {
        detailsStr = String(details);
      }
    }

    // Append error row
    if (sheet) {
      sheet.appendRow([
        new Date(),
        functionName || 'Unknown',
        errorMessage || 'No message',
        detailsStr,
        ''  // Resolved column - mark manually when fixed
      ]);
    }

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
  try {
    hours = hours || 24;
    // Validate hours parameter
    if (typeof hours !== 'number' || hours < 0) hours = 24;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];

    var sheet = ss.getSheetByName(ERROR_LOG_SHEET);

    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return [];

    var cutoff = new Date();
    cutoff.setHours(cutoff.getHours() - hours);

    var recent = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i] || data[i].length < 4) continue;
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
  } catch (e) {
    Logger.log('getRecentErrors error: ' + e.message);
    return [];
  }
}


// ============================================================================
// clearOldErrors — Remove errors older than N days (run weekly)
// ============================================================================

function clearOldErrors(daysToKeep) {
  try {
    daysToKeep = daysToKeep || 30;
    // Validate daysToKeep parameter
    if (typeof daysToKeep !== 'number' || daysToKeep < 0) daysToKeep = 30;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var sheet = ss.getSheetByName(ERROR_LOG_SHEET);

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - daysToKeep);

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return;

    var rowsToDelete = [];

    // Collect rows to delete (iterate backwards to get correct indices)
    for (var i = data.length - 1; i >= 1; i--) {
      if (!data[i] || data[i].length < 1) continue;
      var timestamp = new Date(data[i][0]);
      if (timestamp < cutoff) {
        rowsToDelete.push(i + 1);
      }
    }

    // rowsToDelete was built by iterating i from high to low, so it is already
    // in descending order (largest row index first).  Deleting from index 0
    // (largest row number) down to the end preserves sheet row positions for
    // every subsequent deletion.
    for (var j = 0; j < rowsToDelete.length; j++) {
      if (rowsToDelete[j] > 1) {  // Never delete header row
        sheet.deleteRow(rowsToDelete[j]);
      }
    }

    Logger.log('Cleared ' + rowsToDelete.length + ' old errors');
  } catch (e) {
    Logger.log('clearOldErrors error: ' + e.message);
  }
}


// ============================================================================
// fetchWithRetry — Exponential backoff wrapper for UrlFetchApp.fetch()
//
// Retries on transient failures: HTTP 429, 500, 502, 503, 504, and network errors.
// Does NOT retry on 4xx client errors (except 429) — those are permanent failures.
//
// Usage:
//   var result = fetchWithRetry(url, options);   // returns UrlFetchApp response
//   // throws on exhausted retries or non-retryable errors
//
// Parameters:
//   url       {string}  The URL to fetch
//   options   {Object}  UrlFetchApp options (must include muteHttpExceptions: true)
//   maxTries  {number}  Max attempts, default 3
//   context   {string}  Label for error logging (e.g. 'EasyPost POST /shipments')
// ============================================================================

function fetchWithRetry(url, options, maxTries, context) {
  maxTries = maxTries || 3;
  context = context || 'fetchWithRetry';

  // Ensure muteHttpExceptions so we can inspect the status ourselves
  options = options || {};
  options.muteHttpExceptions = true;

  var RETRYABLE_CODES = { 429: true, 500: true, 502: true, 503: true, 504: true };
  var lastErr = null;

  for (var attempt = 1; attempt <= maxTries; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();

      if (!RETRYABLE_CODES[code]) {
        // Either success (2xx) or a permanent error (4xx) — return immediately
        return response;
      }

      // Retryable HTTP status
      lastErr = 'HTTP ' + code;
      Logger.log(context + ': retryable status ' + code + ' (attempt ' + attempt + '/' + maxTries + ')');

    } catch (networkErr) {
      // Network-level failure (DNS, timeout, etc.)
      lastErr = networkErr.message;
      Logger.log(context + ': network error on attempt ' + attempt + '/' + maxTries + ': ' + networkErr.message);
    }

    if (attempt < maxTries) {
      // Exponential backoff: 2s, 4s, 8s …  capped at 30s
      var delay = Math.min(Math.pow(2, attempt) * 1000, 30000);
      Logger.log(context + ': waiting ' + (delay / 1000) + 's before retry');
      Utilities.sleep(delay);
    }
  }

  var msg = context + ': all ' + maxTries + ' attempts failed. Last error: ' + lastErr;
  Logger.log(msg);
  if (typeof logError === 'function') {
    logError(context, 'All retries exhausted', { url: url, lastError: lastErr, maxTries: maxTries });
  }
  throw new Error(msg);
}


// ============================================================================
// Test function
// ============================================================================

function testErrorLog() {
  try {
    logError('testErrorLog', 'This is a test error', { test: true, value: 123 });
    logWarning('testErrorLog', 'This is a test warning', 'Some details here');
    Logger.log('Check your Error Log sheet');
  } catch (e) {
    Logger.log('testErrorLog error: ' + e.message);
  }
}