# HANDOFF.md — Current Session Context for Claude Code

> **Date:** 2026-02-21 (Session 2 — Parser Fixes)
> **Previous session:** Cowork deep review of entire CSS codebase with 4 specialized agents
> **This session:** Email parser bug fixes across multiple trading companies + infrastructure tweaks

---

## ⚠️ DEPLOYMENT STATUS: NOT YET PUSHED

All changes below are **local only**. Need `clasp push` to deploy, then run `reprocessEmailsFrom()` to test.

---

## 📝 Changes Made This Session

### Files Modified

**`Emailextraction_v3.js`** — Main extraction pipeline

1. **EXCLUDED_DOMAINS** — Added `@theice.com`, `do_not_reply@`, `do-not-reply@`, `donotreply@` to prevent ICE exchange emails and generic no-reply addresses from processing.

2. **`isAutoReply()` function** (new) — Detects OOO/auto-reply emails by checking subject patterns (out of office, automatic reply, etc.) and body patterns (will be back, away from office, etc.). Only triggers if the email also has no sampling keywords (container, cargo, sample, etc.). Called early in `extractOrderFromEmail` to skip these before any parsing.

3. **`NO_ORDER_DATA` check** in `extractOrderFromEmail` post-processing — After extraction, if ALL orders have no container, no mark, and no cargo, AND no sampling instructions found, drops the orders and flags as `NO_ORDER_DATA`. Catches conversational replies like Olam's "Great, thank you" that get client-matched by domain but produce empty rows.

4. **`extractMarks`** — Added decaf-style mark pattern `/(\\d{4,7})/` to catch marks like `/77050/` used for decaffeinated coffee. Filters out year-like numbers (2000-2039).

5. **`extractDescription`** — Added two new fallback patterns:
   - `Total Quantity: N bags GRADE` (Excelco structured format)
   - `Remarks:` section scanning for coffee grade keywords

6. **`extractReceivers`** — Strips `__` markdown bold markers. Added broader `(?:ship|send)\\s+[\\w\\s]+to` pattern to catch "ship priority overnight to:" phrasing (List+Beisler format).

7. **`extractReferences`** — Added `NP[A-Z]{2}-\\d{4,6}[A-Z]?` and `NS[A-Z]{2}-\\d{4,6}[A-Z]?` patterns for Sucafina NP/NS reference numbers with optional letter suffix.

---

**`Parser_V3.js`** — Custom per-client parsers

8. **Sucafina receiver regex** — Changed character class from `[A-Za-z\\s&',\\.]` to `[^\\n_]` to handle curly/smart quotes in company names like "Zingerman's".

9. **Sucafina description enrichment** — After fallback extraction, scans text for origin+grade pattern to enrich bare country descriptions (e.g., "Colombia" → "Colombia Asobombo Huila Organic"). Includes dedup fix so "Colombia Colombia Lowgrades" doesn't happen.

10. **Sucafina standard main loop** — Modified container+cargo pattern to accept bare numbers (no C prefix) with optional letter suffix via alternation: `\\b([A-Z]{4}\\d{7})(C\\d{5,7}[A-Z]?)\\b|\\b([A-Z]{4}\\d{7})\\s+(\\d{6,7}[A-Z]?)\\b`. Bare numbers get `C` prefixed automatically.

11. **Sucafina NP/NS references** — Added letter suffix support (`NPCO-32397A`), plus NS reference fallback (`NSPE-39116`). Applied in both Tastify and standard paths.

12. **Sucafina Tastify row splitter** (CRITICAL FIX) — Rewrote table row splitting logic. Old approach used `\\b` word boundary which failed when row numbers were jammed against prior cargo numbers (e.g., `C3726992CONTI` — the `2` is row 2, not part of cargo). New approach: finds ALL digit+warehouse matches, then walks them looking for the sequential chain (1, 2, 3, ...) to identify real row boundaries. Fixes missing samples.

13. **Rothfos/NKG parser** — Major enhancement to handle plain-text email format (not just structured PDF):
    - Reference from subject: `\\bP(\\d{4,7}[A-Z]?)\\b`
    - Sender ref from subject: `\\bS\\s*(\\d{5,12})\\b`
    - ICO mark: Added "ICO number" label pattern + standalone 3-part mark fallback
    - Container: Added `(?:number|#|No\\.?)` to label pattern
    - Cargo: Added `(?:number\\s+)?` to label pattern
    - Bags: Added `\\bBAGS\\b` label fallback
    - Description: Added QUALITY label fallback and subject-line origin abbreviation expansion (HOND→Honduras, etc.)
    - Receiver: Added broader `(?:send|ship)\\s+[\\w\\s]+to\\s+` pattern and company name block pattern

---

**`Config.js`** — Search query and utilities

14. **`buildSearchQuery`** — Added `after:2026/02/15` date filter so only emails from Feb 15 onward are processed. This prevents reprocessing the entire inbox.

15. **`reprocessEmailsFrom()`** (new utility) — One-time function to strip `PDF_Processed` label from all threads after 2/15. Run once from script editor to retest parser fixes, then delete.

---

## 🧪 Testing Plan

1. `clasp push` to deploy all changes
2. Run `reprocessEmailsFrom()` from script editor — strips labels from threads after 2/15
3. Run `processPDFsFromGmail` manually (or wait for 15-min trigger)
4. Check "All Orders" sheet for:
   - Sucafina/Zingerman's: receiver = "Zingerman's Coffee Company", full description
   - Excelco/Regal: receiver = "REGAL COMMODITIES", mark = "77050", cargo present
   - List+Beisler: receiver populated from "ship priority overnight to:" line
   - Rothfos/NKG: receiver = "Ellis Coffee", mark/reference/description present
   - Sucafina/Gavina: 4 samples (not 2), cargos present, NP reference with letter suffix
   - Sucafina/Roosevelt: 3 samples (not 2), receiver = "ROOSEVELT COFFEE LLC"
   - Olam "thank you": NOT processed (NO_ORDER_DATA flag)
   - Erica Arevalos OOO: NOT processed (AUTO_REPLY flag)
   - ICE do_not_reply: NOT processed (EXCLUDED)
5. Clean up duplicate rows from reprocessing

⚠️ **After testing:** Consider removing the `after:2026/02/15` date filter from `Config.js` line 102 and deleting `reprocessEmailsFrom()` — both are temporary.

---

## 🔥 URGENT Items (Carried Forward from Session 1)

**Triggers may still be dead.** Run `createTimeTriggers()` from Menu.js if the 15-min email processing isn't running automatically.

**Backup cleanup bug** in `Export.js` line 183 — exact name match vs. date-suffixed filenames. Old backups accumulate in Drive.

---

## 📋 Review Findings (Carried Forward from Session 1)

### Critical
1. Email parsing failures are silent/unrecoverable — need "Failed Emails" quarantine sheet
2. No `LockService` concurrency protection — race conditions on sheet writes
3. Triggers silently disappear — need startup verification

### High Priority
4. Billing not idempotent — duplicate line items on rerun
5. Web app endpoints have no server-side auth
6. No exponential backoff on EasyPost API calls
7. Error log grows unbounded — `clearOldErrors()` has no trigger
8. Schema validation missing — no column existence checks

### Medium Priority
9. No offline queue for warehouse scanning
10. No data access abstraction layer
11. Customer/parser config is hardcoded
12. Learning Engine can propagate correction errors

---

## 📌 Pending Work

- [ ] Verify Roosevelt Tastify receiver extraction works (Customer: line match)
- [ ] `clasp push` all changes
- [ ] Run `reprocessEmailsFrom()` and validate
- [ ] Clean up test duplicates
- [ ] Remove `after:2026/02/15` filter and `reprocessEmailsFrom()` when done testing
- [ ] Commit all working changes to git

---

## 📌 Dan's Working Style (from CLAUDE.md)

- Hands-on, technical. Skip theory, show code.
- Diffs only unless file is new.
- Don't refactor what you weren't asked to touch.
- Flag risks with `⚠️ HEADS UP:` — he'll decide.
- When in doubt, stop and ask.

---

*Generated by Cowork session 2026-02-21 (Session 2). This file supplements CLAUDE.md — read both.*
