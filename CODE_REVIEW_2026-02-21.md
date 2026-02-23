# CSS Apps Script - Full Codebase Review

**Date:** February 21, 2026
**Scope:** All 30 .js files and 4 .html files (~23,000 lines)
**Goal:** Identify bugs, dead code, silent failures, security risks, and sloppy patterns. No refactor suggestions unless something is actively broken.

---

## CRITICAL Issues (Fix Soon)

### 1. Hardcoded Sheet Name Mismatch — Menu.js & Setup.js

`unlockHeaders()` references `'ALL Orders'` but `CONFIG.mainSheetName` is `'All Orders'`. The sheet will never be found; unlocking silently fails.

**Files:** Menu.js ~line 458, Setup.js ~line 458

### 2. Regex Escape Bug in Tracking Number Detection — scanout.js

`looksLikeTracking()` uses double-escaped backslashes inside regex literals: `/(94|93|92|91|90)\\d{18,22}$/` and `/(^\\d{12,22}$)/`. The `\\d` is treated as a literal backslash + "d" instead of a digit class. USPS and FedEx tracking numbers won't be recognized.

**File:** scanout.js ~line 574-576

### 3. Duplicate Detection Only Checks Last 200 Rows — Emailextraction_v3.js

`checkForDuplicate()` only scans the last 200 rows. If the sheet has 1000+ rows and a duplicate exists beyond that window, a duplicate order is created silently.

**File:** Emailextraction_v3.js ~line 522

### 4. LearningEngine Auto-Apply After Only 2 Corrections

`minConfidence: 2` means after just 2 corrections, a rule auto-applies globally to all future orders from that sender. A single typo correction could propagate to every order from that trading house.

**File:** LearningEngine.js ~line 268

### 5. `processPendingMoves()` Leaves Orphaned Rows

After deleting rows by CS Sample #, the code only deletes the first match per sample. If duplicate CS Sample #s exist, subsequent duplicates remain as orphaned rows.

**File:** Orders.js ~line 1177

### 6. JSON Parse Without Try-Catch on API Responses — EasyPost.js

`_easyPostRequest()` and `_easyPostGet()` call `JSON.parse(response.getContentText())` with no error handling. If the API returns HTML (error page, timeout), this crashes with an unhelpful error.

**File:** EasyPost.js ~lines 138, 163

### 7. ManualOrderEntry.html — Code Runs Before DOM Ready

`populateDropdowns()` and `restoreState()` execute at script parse time, before DOM elements exist. `getElementById()` returns null, causing cascading silent failures on form initialization.

**File:** ManualOrderEntry.html ~lines 509-510

### 8. ManualOrderEntry.html — Server Template Injection

`var CONTACTS = <?!= contactsJson ?>;` injects raw JSON into a script tag. Malformed JSON or unescaped quotes break the page entirely.

**File:** ManualOrderEntry.html ~line 227

---

## HIGH Issues (Risky, Should Fix)

### 9. Unsafe Role Fallback — Menu.js

`_getUserRole()` defaults to `'editor'` if email detection fails. An unknown user hitting a simple trigger could get editor access.

**File:** Menu.js ~lines 105-125

### 10. Inconsistent Column Key Lookups — Labels.js

`printCheckedSamples()` uses `col['CS Sample']` while other functions use `col['CS Sample #']`. If the exact key doesn't exist, the value is `undefined` and labels print blank.

**File:** Labels.js ~lines 518, 599-614

### 11. Cross-File Function Dependencies Without Guards

Multiple files call functions defined in other files (`addDataToSheet()`, `_buildTestOrders()`, `escapeCSVField()`, `updateLiveOrdersView()`, `_findRowBySample()`, `_getColumnMap()`) with zero defensive checks. If any file fails to load, multiple modules crash.

**Files:** Export.js, InputManual.js, RunTestOrder.js, EasyPost.js, Emailextraction_v3.js

### 12. Silent Thread Access Failure — Orders.js

`repairReceivers()` calls `getThreadById()` with no null check. If the thread ID is stale or malformed, the try-catch at line 1822 swallows the error and the row is silently left unfixed.

**File:** Orders.js ~line 1353

### 13. Collision-Prone Message Key — Emailextraction_v3.js

`message.getId().substring(0, 8)` is used as a unique key, but 8 characters from a longer ID is collision-prone. Different emails could be grouped as the same order.

**File:** Emailextraction_v3.js ~lines 458-459

### 14. Unescaped User Input in Prompt — WebAppScanner.html

User input from `prompt()` is passed directly to server-side functions without sanitization or escaping.

**File:** WebAppScanner.html ~lines 1067-1071

### 15. Missing Base64 Validation — WebAppScanner.html

FileReader result is split on comma without checking the result contains one. Certain file types cause a TypeError crash.

**File:** WebAppScanner.html ~line 874

### 16. Incorrect href Escaping — Clientportal.html

The `esc()` function uses `textContent` (text-safe) but is applied to `href` attributes. `javascript:` URLs can still execute.

**File:** Clientportal.html ~line 359

### 17. Fire-and-Forget Contact Save — ManualOrderEntry.html

`saveNewContact()` calls lack `withFailureHandler()`. The UI updates optimistically but if the save fails, data is lost silently.

**File:** ManualOrderEntry.html ~lines 436, 440

---

## MEDIUM Issues (Should Track)

### 18. Hardcoded Order Number Base — Orders.js

`_buildOrderCache()` hardcodes `highestOrderNumber: 211656`. No validation that this hasn't already been assigned. Risk of duplicate order numbers if base drifts.

**File:** Orders.js ~line 1315

### 19. Silent Auto-Billing Failure — Billing.js

`_autoBillSamples()` silently returns if the rate sheet has no rates. If rates are accidentally deleted, all billing stops with zero notification.

**File:** Billing.js ~lines 198-199

### 20. Stale Shipping Fee on Recalculation — Billing.js

`_recalcInvoiceTotal()` only updates the SHIP-FEE line item if the new fee > 0. If the rate becomes 0%, the old fee value persists.

**File:** Billing.js ~line 1597

### 21. Fuzzy Company Match Returns First Hit — Billing.js & EasyPost.js

`_getCustomerAddress()` and `_getReceiverContact()` use `indexOf()` for fuzzy matching. "Cafe" matches whichever of "Cafe Kreyol" or "Cafe Lobo" comes first.

**Files:** Billing.js ~line 871, EasyPost.js ~line 594

### 22. Sender Pattern Matching Logic Bug — LearningEngine.js

Double-negative `indexOf()` logic: `senderLower.indexOf(rulePattern) === -1 && rulePattern.indexOf(senderLower) === -1`. "louis dreyfus" rule won't match "louis dreyfus company" because the first check succeeds and short-circuits.

**File:** LearningEngine.js ~line 345

### 23. Hardcoded Column Indices — Carrieraccounts.js & EasyPost.js

New contact row construction and receiver lookups use hardcoded column indices instead of header lookups. Any column reorder breaks these silently.

**Files:** Carrieraccounts.js ~lines 199, 261-267; EasyPost.js ~line 557

### 24. Formula Injection Risk — scanout.js

`setFormula()` interpolates `trackingNum` directly. Special characters or quotes in tracking numbers could break or inject formulas.

**File:** scanout.js ~line 182

### 25. Receiver Type Ambiguity — Emailextraction_v3.js

Parsers return `order.receiver` as either a string or an object `{name, address}`. No type enforcement; downstream code handles both but inconsistently.

**File:** Emailextraction_v3.js ~line 482

### 26. MorningBriefing Recipient Evaluated at Load Time

`BRIEFING_CONFIG.recipientEmail` calls `Session.getActiveUser().getEmail()` when the script loads, not when the briefing sends. Trigger-based execution sends to the wrong person.

**File:** MorningBriefing_V3.js ~line 27

### 27. Weekend Skip Uses Wrong Timezone Source

Weekend detection hardcodes `America/New_York` instead of using `BRIEFING_CONFIG.timezone`.

**File:** MorningBriefing_V3.js ~lines 82-83

### 28. Cache Expiration Without Notification — Labels.js

5-minute cache timeout. If label generation takes longer, `_getSamplesFromCache()` returns null and shows "Session expired" with no way to recover without restarting.

**File:** Labels.js ~line 726

### 29. Weak Portal Token Generation — ClientPortal.js

Uses `Math.random()` instead of `Utilities.getUuid()` or cryptographic random. Tokens are guessable.

**File:** ClientPortal.js ~lines 75-79

### 30. `SPREADSHEET_ID` Never Defined — Photos.js

Variable referenced but never assigned. Fallback to `getActiveSpreadsheet()` always triggers, which is unreliable in web app context.

**File:** Photos.js ~line 35

### 31. Error Swallowing Throughout — Multiple Files

`catch(ignore) {}` and `catch(e) { /* logged only */ }` patterns appear across scanout.js, Menu.js, WebApp.js, and Reports.js. Critical operations like `updateLiveOrdersView()` fail silently.

### 32. Unescaped Datalist Values — ManualOrderEntry.html

Company names injected into `<option>` values without HTML entity escaping. Malicious company names could inject HTML.

**File:** ManualOrderEntry.html ~lines 467-475

### 33. Auto-Refresh Without Backoff — Clientportal.html

Portal auto-refresh has no exponential backoff. If the server is slow, rapid retries compound the load.

---

## LOW Issues (Noted for Awareness)

### 34. Coffee Origin List Hardcoded — Dashboard.js ~line 745
New origins (e.g., Yemen) default to "Unknown". List needs manual maintenance.

### 35. Whitespace Duplicates in Contacts — Contacts.js ~line 198
Leading/trailing spaces create "duplicate" entries that pass the dedup check.

### 36. Export Files Go to Drive Root — Export.js ~lines 40, 72, 148
All exports land in root folder. No organization by date or type.

### 37. Errorlog Deletes Row-by-Row — Errorlog.js ~line 131
`deleteRow()` in a loop is slow with many rows. Could timeout on large error sheets.

### 38. `getMarketPredictions_()` Naming — marketpredictions.js ~line 54
Trailing underscore (private convention) but called from menu items as public.

### 39. Static Market Data — marketpredictions.js ~line 14
Hardcoded as of 2026-02-14. No refresh mechanism.

### 40. Accessibility Gaps — All HTML Files
Missing `type="button"` on non-submit buttons, missing `for` attributes on labels, generated images lack alt text, div-based buttons not keyboard accessible.

---

## Cross-File Dependency Map (Undocumented)

These cross-file calls have **no defensive guards** and will crash if the source file fails to load:

| Caller | Function Called | Expected Source |
|--------|----------------|-----------------|
| Emailextraction_v3.js | `_getColumnMap()`, `_buildOrderCache()`, `checkForDuplicate()` | Orders.js |
| EasyPost.js | `_findRowBySample()`, `_getColumnMap()` | scanout.js |
| InputManual.js | `addDataToSheet()` | Orders.js (presumed) |
| RunTestOrder.js | `_buildTestOrders()` | Unknown |
| Export.js | `escapeCSVField()` | Unknown |
| CustomerConfig.js | `KNOWN_CLIENTS` | Emailextraction_v3.js |
| Config.js | `KNOWN_CLIENTS` | Emailextraction_v3.js |
| QRcode.js | `getCountryFromICOMark()` | Unknown |

---

## HTML ↔ Server Function Cross-Reference

All `google.script.run` calls verified against server-side code. **No mismatches found.** All 11 unique calls resolve to existing functions.

---

## What's Clean

- **No dead functions detected.** Everything defined appears to be called.
- **Parser organization** is sound — one parser per trading company is the right pattern for this domain.
- **Dual caching in scanner modules** is well-implemented.
- **HTML files are mostly well-structured** with consistent patterns for `google.script.run` calls.
- **Error logging to sheet** (Errorlog.js) is a solid pattern for a GAS codebase.

---

*Generated by automated code review — line numbers are approximate due to file size. Verify against current source before making changes.*
