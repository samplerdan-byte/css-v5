// ============================================================================
// QuotaGuards.js — Google Apps Script Quota & Execution Limit Guards
// Commodity Sampler Services
// ============================================================================
// Tracks and prevents hitting these GAS hard limits:
//   • 6-minute execution limit per function
//   • Email send quota (100/day free, 1500/day Workspace)
//   • UrlFetch quota (20,000/day)
//   • Properties storage (9MB total, 500KB per value)
//   • CacheService limits (100KB per value, 25MB total)
//   • Sheet cell limit (10M cells)
//   • Trigger execution (90 min/day for simple triggers)
//
// All quota data stored in PropertiesService (Script Properties).
// Data resets daily at midnight UTC.
//
// Usage:
//   - Call _quotaCheckExecutionTime() at start of long-running functions
//   - Call _quotaLogEmailSend() after each email sent
//   - Call _quotaLogUrlFetch() after each API call
//   - Call _quotaCheckProperties() before large writes
//   - Use _quotaGetStatus() to display dashboard
//
// Integration:
//   - processPDFsFromGmail(): Already has 5-min guard
//   - EasyPost functions: Wrap in _quotaCheckUrlFetch before fetch calls
//   - Email functions: Wrap in _quotaCheckEmailQuota before GmailApp.sendEmail
//
// ============================================================================

// ============================================================================
// QUOTA KEYS IN PROPERTIES SERVICE
// ============================================================================

var QUOTA_KEYS = {
  EMAIL_SENT_TODAY: 'QUOTA_EMAIL_SENT_TODAY',
  EMAIL_LAST_RESET: 'QUOTA_EMAIL_LAST_RESET',
  URLFETCH_COUNT_TODAY: 'QUOTA_URLFETCH_COUNT_TODAY',
  URLFETCH_LAST_RESET: 'QUOTA_URLFETCH_LAST_RESET',
  TRIGGER_TIME_TODAY: 'QUOTA_TRIGGER_TIME_TODAY',
  TRIGGER_LAST_RESET: 'QUOTA_TRIGGER_LAST_RESET'
};

var QUOTA_LIMITS = {
  EXECUTION_TIME_MS: 360000,      // 6 minutes
  EXECUTION_SAFETY_MARGIN_MS: 300000, // Stop at 5 min to be safe
  EMAIL_QUOTA_FREE: 100,
  EMAIL_QUOTA_WORKSPACE: 1500,
  URLFETCH_QUOTA: 20000,
  PROPERTIES_TOTAL_BYTES: 9437184, // 9MB
  PROPERTIES_VALUE_BYTES: 524288,  // 500KB per value
  CACHE_VALUE_BYTES: 102400,       // 100KB per value
  CACHE_TOTAL_BYTES: 26214400      // 25MB
};

// ============================================================================
// EXECUTION TIME GUARD
// ============================================================================

/**
 * Check if we should stop to avoid 6-minute timeout.
 * Returns {shouldStop: boolean, timeElapsedMs: number, remainingMs: number}
 */
function _quotaCheckExecutionTime(startTimeMs) {
  var elapsedMs = Date.now() - startTimeMs;
  var remainingMs = QUOTA_LIMITS.EXECUTION_SAFETY_MARGIN_MS - elapsedMs;
  return {
    shouldStop: elapsedMs > QUOTA_LIMITS.EXECUTION_SAFETY_MARGIN_MS,
    timeElapsedMs: elapsedMs,
    remainingMs: Math.max(0, remainingMs)
  };
}

/**
 * Guard helper: stop processing if near timeout.
 * Use in loops:
 *   for (var i = 0; i < data.length; i++) {
 *     if (i % 10 === 0 && _quotaCheckExecutionTime(startTime).shouldStop) break;
 *     // process data[i]
 *   }
 */
function _quotaGuardExecutionTime(startTimeMs, iterationNumber, checkInterval) {
  // Check every N iterations (default 20) for performance
  checkInterval = checkInterval || 20;
  if (iterationNumber % checkInterval === 0) {
    return _quotaCheckExecutionTime(startTimeMs);
  }
  return { shouldStop: false };
}

// ============================================================================
// EMAIL QUOTA TRACKING
// ============================================================================

/**
 * Get daily email quota for this account.
 * Note: Detection is heuristic — actual limit depends on account type.
 * Workspace accounts typically have 1500/day; free accounts 100/day.
 */
function _quotaGetEmailLimit() {
  try {
    // Rough heuristic: check if this is a Workspace account
    var email = Session.getActiveUser().getEmail();
    if (email && email.indexOf('@gmail.com') === -1 && email.indexOf('@googlemail.com') === -1) {
      return QUOTA_LIMITS.EMAIL_QUOTA_WORKSPACE; // likely Workspace
    }
  } catch (e) {
    Logger.log('Could not determine email quota; assuming Workspace (1500/day)');
  }
  return QUOTA_LIMITS.EMAIL_QUOTA_WORKSPACE; // default to Workspace since that's what CSS uses
}

/**
 * Log an email send. Call AFTER GmailApp.sendEmail() succeeds.
 */
function _quotaLogEmailSend() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  var lastReset = props.getProperty(QUOTA_KEYS.EMAIL_LAST_RESET) || '';

  // Reset counter if day changed
  if (lastReset !== today) {
    props.setProperty(QUOTA_KEYS.EMAIL_SENT_TODAY, '0');
    props.setProperty(QUOTA_KEYS.EMAIL_LAST_RESET, today);
  }

  var count = parseInt(props.getProperty(QUOTA_KEYS.EMAIL_SENT_TODAY) || '0', 10);
  count++;
  props.setProperty(QUOTA_KEYS.EMAIL_SENT_TODAY, String(count));

  Logger.log('_quotaLogEmailSend: Count is now ' + count);
}

/**
 * Check email quota before sending.
 * Returns {canSend: boolean, countToday: number, limit: number, remaining: number}
 */
function _quotaCheckEmailQuota() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  var lastReset = props.getProperty(QUOTA_KEYS.EMAIL_LAST_RESET) || '';

  // Reset counter if day changed
  if (lastReset !== today) {
    props.setProperty(QUOTA_KEYS.EMAIL_SENT_TODAY, '0');
    props.setProperty(QUOTA_KEYS.EMAIL_LAST_RESET, today);
  }

  var limit = _quotaGetEmailLimit();
  var count = parseInt(props.getProperty(QUOTA_KEYS.EMAIL_SENT_TODAY) || '0', 10);
  var remaining = Math.max(0, limit - count);

  return {
    canSend: remaining > 0,
    countToday: count,
    limit: limit,
    remaining: remaining
  };
}

// ============================================================================
// URLFETCH QUOTA TRACKING
// ============================================================================

/**
 * Log a UrlFetch call. Call AFTER UrlFetchApp.fetch() succeeds.
 */
function _quotaLogUrlFetch() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  var lastReset = props.getProperty(QUOTA_KEYS.URLFETCH_LAST_RESET) || '';

  // Reset counter if day changed
  if (lastReset !== today) {
    props.setProperty(QUOTA_KEYS.URLFETCH_COUNT_TODAY, '0');
    props.setProperty(QUOTA_KEYS.URLFETCH_LAST_RESET, today);
  }

  var count = parseInt(props.getProperty(QUOTA_KEYS.URLFETCH_COUNT_TODAY) || '0', 10);
  count++;
  props.setProperty(QUOTA_KEYS.URLFETCH_COUNT_TODAY, String(count));

  Logger.log('_quotaLogUrlFetch: Count is now ' + count);
}

/**
 * Check UrlFetch quota before making API calls.
 * Returns {canFetch: boolean, countToday: number, limit: number, remaining: number}
 */
function _quotaCheckUrlFetchQuota() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  var lastReset = props.getProperty(QUOTA_KEYS.URLFETCH_LAST_RESET) || '';

  // Reset counter if day changed
  if (lastReset !== today) {
    props.setProperty(QUOTA_KEYS.URLFETCH_COUNT_TODAY, '0');
    props.setProperty(QUOTA_KEYS.URLFETCH_LAST_RESET, today);
  }

  var limit = QUOTA_LIMITS.URLFETCH_QUOTA;
  var count = parseInt(props.getProperty(QUOTA_KEYS.URLFETCH_COUNT_TODAY) || '0', 10);
  var remaining = Math.max(0, limit - count);

  return {
    canFetch: remaining > 0,
    countToday: count,
    limit: limit,
    remaining: remaining
  };
}

// ============================================================================
// PROPERTIES STORAGE GUARD
// ============================================================================

/**
 * Check if a value is safe to write to PropertiesService.
 * Returns {canWrite: boolean, valueBytes: number, maxBytes: number}
 */
function _quotaCheckProperties(value) {
  try {
    var valueStr = typeof value === 'string' ? value : JSON.stringify(value);
    var bytes = Utilities.getUuid().length; // rough estimate; better to use actual string bytes
    bytes = valueStr ? valueStr.length : 0;

    return {
      canWrite: bytes <= QUOTA_LIMITS.PROPERTIES_VALUE_BYTES,
      valueBytes: bytes,
      maxBytes: QUOTA_LIMITS.PROPERTIES_VALUE_BYTES
    };
  } catch (e) {
    Logger.log('_quotaCheckProperties error: ' + e);
    return { canWrite: false, valueBytes: 0, maxBytes: QUOTA_LIMITS.PROPERTIES_VALUE_BYTES };
  }
}

// ============================================================================
// CACHE SERVICE GUARD
// ============================================================================

/**
 * Check if a value is safe to write to CacheService.
 * Returns {canCache: boolean, valueBytes: number, maxBytes: number}
 */
function _quotaCheckCache(value) {
  try {
    var valueStr = typeof value === 'string' ? value : JSON.stringify(value);
    var bytes = valueStr ? valueStr.length : 0;

    return {
      canCache: bytes <= QUOTA_LIMITS.CACHE_VALUE_BYTES,
      valueBytes: bytes,
      maxBytes: QUOTA_LIMITS.CACHE_VALUE_BYTES
    };
  } catch (e) {
    Logger.log('_quotaCheckCache error: ' + e);
    return { canCache: false, valueBytes: 0, maxBytes: QUOTA_LIMITS.CACHE_VALUE_BYTES };
  }
}

// ============================================================================
// TRIGGER EXECUTION TIME TRACKING
// ============================================================================

/**
 * Log trigger execution time. Call at end of trigger-bound function.
 * @param {number} elapsedMs — Milliseconds spent in the trigger
 */
function _quotaLogTriggerTime(elapsedMs) {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  var lastReset = props.getProperty(QUOTA_KEYS.TRIGGER_LAST_RESET) || '';

  // Reset counter if day changed
  if (lastReset !== today) {
    props.setProperty(QUOTA_KEYS.TRIGGER_TIME_TODAY, '0');
    props.setProperty(QUOTA_KEYS.TRIGGER_LAST_RESET, today);
  }

  var totalMs = parseInt(props.getProperty(QUOTA_KEYS.TRIGGER_TIME_TODAY) || '0', 10);
  totalMs += elapsedMs;
  props.setProperty(QUOTA_KEYS.TRIGGER_TIME_TODAY, String(totalMs));

  Logger.log('_quotaLogTriggerTime: Total trigger time today: ' + (totalMs / 60000).toFixed(2) + ' min');
}

/**
 * Get trigger execution time budget.
 * Returns {totalMinutes: number, limit: number, remaining: number}
 */
function _quotaGetTriggerTimeStatus() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  var lastReset = props.getProperty(QUOTA_KEYS.TRIGGER_LAST_RESET) || '';

  // Reset if day changed
  if (lastReset !== today) {
    props.setProperty(QUOTA_KEYS.TRIGGER_TIME_TODAY, '0');
    props.setProperty(QUOTA_KEYS.TRIGGER_LAST_RESET, today);
  }

  var totalMs = parseInt(props.getProperty(QUOTA_KEYS.TRIGGER_TIME_TODAY) || '0', 10);
  var totalMin = totalMs / 60000;
  var limit = 90; // 90 min/day for simple triggers
  var remaining = Math.max(0, limit - totalMin);

  return {
    totalMinutes: totalMin,
    limit: limit,
    remaining: remaining
  };
}

// ============================================================================
// STATUS DASHBOARD
// ============================================================================

/**
 * Get overall quota status for logging/display.
 * Returns object with all quota metrics for today.
 */
function _quotaGetStatus() {
  var emailQuota = _quotaCheckEmailQuota();
  var urlQuota = _quotaCheckUrlFetchQuota();
  var triggerStatus = _quotaGetTriggerTimeStatus();

  return {
    timestamp: new Date().toISOString(),
    email: {
      sent: emailQuota.countToday,
      limit: emailQuota.limit,
      remaining: emailQuota.remaining,
      status: emailQuota.remaining > 0 ? 'OK' : 'LIMIT_REACHED'
    },
    urlfetch: {
      calls: urlQuota.countToday,
      limit: urlQuota.limit,
      remaining: urlQuota.remaining,
      status: urlQuota.remaining > 0 ? 'OK' : 'LIMIT_REACHED'
    },
    trigger: {
      minutesUsed: triggerStatus.totalMinutes.toFixed(2),
      minutesLimit: triggerStatus.limit,
      minutesRemaining: triggerStatus.remaining.toFixed(2),
      status: triggerStatus.remaining > 0 ? 'OK' : 'LIMIT_REACHED'
    }
  };
}

/**
 * Display quota status in UI and log.
 */
function showQuotaStatus() {
  var status = _quotaGetStatus();
  var ui = SpreadsheetApp.getUi();

  var msg = '📊 QUOTA STATUS — Today\n\n';
  msg += '📧 Email Sends\n';
  msg += '  Sent: ' + status.email.sent + '/' + status.email.limit + '\n';
  msg += '  Remaining: ' + status.email.remaining + '\n\n';

  msg += '🔗 UrlFetch API Calls\n';
  msg += '  Used: ' + status.urlfetch.calls + '/' + status.urlfetch.limit + '\n';
  msg += '  Remaining: ' + status.urlfetch.remaining + '\n\n';

  msg += '⏱️  Trigger Execution\n';
  msg += '  Used: ' + status.trigger.minutesUsed + '/' + status.trigger.minutesLimit + ' min\n';
  msg += '  Remaining: ' + status.trigger.minutesRemaining + ' min\n';

  Logger.log('QUOTA STATUS: ' + JSON.stringify(status));
  ui.alert(msg);
}

/**
 * Reset all quota counters (admin use only).
 */
function _quotaResetAll() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(QUOTA_KEYS.EMAIL_SENT_TODAY);
  props.deleteProperty(QUOTA_KEYS.EMAIL_LAST_RESET);
  props.deleteProperty(QUOTA_KEYS.URLFETCH_COUNT_TODAY);
  props.deleteProperty(QUOTA_KEYS.URLFETCH_LAST_RESET);
  props.deleteProperty(QUOTA_KEYS.TRIGGER_TIME_TODAY);
  props.deleteProperty(QUOTA_KEYS.TRIGGER_LAST_RESET);
  Logger.log('All quota counters reset.');
}

// ============================================================================
// END OF FILE
// ============================================================================
