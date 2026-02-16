# CLAUDE.md — Commodity Sampler Services (CSS) Apps Script Project

> **Read this entire file before writing a single line of code.**

---

## 🔒 PRIME DIRECTIVE

**NEVER BREAK EXISTING WORKING CODE.**

This is a live production system running a real business. Every function, every quirk, every "weird" pattern exists because it works and someone depends on it. Treat the codebase the way a surgeon treats a patient — precise cuts, minimal disruption, full awareness of what's connected to what.

### Rules of Engagement

1. **No rewrites unless explicitly asked.** If Dan says "fix X," fix X. Don't refactor the file. Don't "improve" adjacent code. Don't rename things for consistency. Touch only what you're told to touch.

2. **Diffs only.** Show the exact lines changing. Never output an entire file unless Dan asks for one or the file is brand new.

3. **Preserve function signatures.** If a function is called from other files, HTML templates, menu items, triggers, or `google.script.run`, its name and parameter list are frozen. Changing them breaks things you can't see.

4. **When in doubt, stop and ask Dan.** If something looks wrong, smells fragile, or has unclear dependencies — flag it. Don't guess. Don't "fix" it silently. A 30-second question beats a 3-hour debugging session.

5. **Work at whatever pace the task demands.** Simple rename? Fast. Refactoring a parser that touches 5 files? Slow, methodical, one piece at a time. There is no rush. Getting it right matters more than getting it done quickly.

6. **Use smart functions when they help, brute force when they don't.** Regex is great for structured text. Don't use it to parse messy PDFs when line-by-line iteration is more reliable. Match the tool to the problem.

7. **Call on Dan to fix things that are worth it.** If you spot a bug that's outside the current task scope but genuinely matters — flag it clearly with `⚠️ HEADS UP:` and explain what you found and what Dan should test. Don't silently fix side issues.

---

## 📋 Project Overview

**Business:** Commodity Sampler Services (CSS) — family-owned third-party coffee and cocoa sampling company.  
**Tagline:** "Integrity Through Independence"  
**Platform:** Google Apps Script, built on Google Sheets  
**Script ID:** `1CpNGoW3LYgUXoXD0odbSXpI502ffkf1koMEo4WrLe6ctwcSEjb0miGIM`  
**Local Dev Path:** `C:\Users\dan\CSS-AppsScript`  
**Backup Path:** `C:\Users\dan\CSS-AppsScript-BACKUP`

### What the system does

- **Email processing:** Reads incoming emails from major coffee trading houses (Louis Dreyfus, Olam, Volcafe, Sucafina, Serengeti, Atlantic USA, Armenia Coffee Corp, and others). Custom parsers extract order data from various PDF and email formats.
- **PDF data extraction:** Pulls structured data (contract numbers, marks, weights, container numbers, etc.) from attached PDFs in multiple formats per client.
- **Warehouse operations:** Barcode/QR scanning interfaces for receiving, sampling, and shipping. Label printing. Inventory tracking.
- **Shipping workflows:** Carrier integration (FedEx, UPS), shipping label generation, tracking, scanner validation with history displays.
- **Container tracking:** Bulk container status lookups via carrier/shipping line data.
- **Billing integration:** Generates QuickBooks-compatible billing from completed sampling work.
- **Field reporting:** Multi-user reporting with email archive access.
- **Photo/attachment linking:** Auto-saves PDFs and images to Google Drive, links them back to order records.

### Key technical patterns

- Google Sheets as database (multiple sheets within workbooks)
- HTML templates served via `HtmlService` for all UI (sidebars, modals, web apps)
- `google.script.run` bridges between client HTML and server-side `.gs` functions
- `PropertiesService` and `CacheService` for state and performance
- `GmailApp` / `Gmail API` for email processing
- `DriveApp` for file storage
- Custom email parsers per trading company — each has its own quirks and format
- Dual caching strategy in scanning modules to reduce server round-trips

---

## 📁 File Map

> **TODO:** Dan needs to paste all project files so this section can be populated with one-line descriptions per file. Until then, known file categories:

### Server-side (.gs → .js locally via clasp)
- `Main.gs` — Core entry point. **Planned split into:** Scanner.gs, Reports.gs, Export.gs, Views.gs, Setup.gs
- Various parser files per trading company
- Shipping/carrier integration
- QuickBooks billing functions
- Email processing and archiving
- Container tracking

### Client-side (.html)
- Scanner UI templates
- Shipping interfaces
- Report viewers
- Sidebar/modal templates
- Inline CSS and JS (being extracted to separate template files)

**To populate this section:** Paste all filenames (or the full codebase) and I'll generate one-line descriptions for each.

---

## 🔧 Development Workflow

### Setup (already done)
```
Node.js v20.18.1 installed
npm install -g @google/clasp
clasp login (authenticated)
Project cloned to C:\Users\dan\CSS-AppsScript
Git initialized with initial commit
```

### Daily workflow
```bash
# 1. Sync latest from Google (in case of live edits)
clasp pull

# 2. Branch for the task
git checkout -b fix/description-of-change

# 3. Make edits (locally or via Claude)

# 4. Push to Apps Script and test LIVE
clasp push

# 5. If it works → commit and merge
git add -A
git commit -m "fix: description of what changed and why"
git checkout main
git merge fix/description-of-change

# 6. If it breaks → revert instantly
git checkout .
clasp push
```

### Emergency revert
```bash
git checkout .        # undo all local changes
clasp push            # restore last known good to Apps Script
```

### Giving code to Claude
- Copy/paste relevant files (or whole codebase) into the chat
- Always include this CLAUDE.md at the start of the conversation
- For targeted fixes: paste only the affected file(s) + describe the problem
- For broad work: paste everything so Claude has full context

---

## 📌 Current Backlog

### Ready
- [ ] Split `Main.gs` into `Scanner.gs`, `Reports.gs`, `Export.gs`, `Views.gs`, `Setup.gs`
- [ ] Container tracking via carrier emails
- [ ] Faster shipping scanner
- [ ] Extraction testing (PDF parser validation)
- [ ] Populate file map in this document (paste codebase to Claude)

### In Progress
- [ ] Email/PDF/photo linking capabilities
- [ ] Bulk container tracking interface

### Done
- [x] clasp + git local dev setup
- [x] Initial project clone and backup
- [x] PowerShell execution policy configured
- [x] Consolidated email parsers for trading clients
- [x] Dual caching in scanning modules
- [x] Email attachment auto-save to Google Drive
- [x] Carrier account info transcribed (FedEx, UPS)
- [x] Fuzzy matching for contact population
- [x] Field reporting functionality
- [x] Email archive for multi-user access
- [x] Inline HTML extraction to template files (ongoing)

---

## ⚠️ Known Landmines

Things that can break silently if you're not careful:

1. **`google.script.run` calls** — If you rename a server function that's called from HTML, the UI breaks with no useful error. Always search `.html` files for the function name before renaming.

2. **Trigger-bound functions** — Some functions run on time-based or form-submit triggers. Renaming them breaks the trigger silently. Ask Dan which functions have triggers before touching them.

3. **PropertiesService keys** — String keys stored in Script/User/Document properties. Changing a key name orphans the stored data. Always search for the key string across all files.

4. **Sheet column order** — Many functions reference columns by index (e.g., `row[4]`). Adding/removing columns in the spreadsheet will break these silently.

5. **PDF parser fragility** — Each trading company's parser is tuned to their specific PDF format. Even minor format changes from the trading house can break extraction. These are inherently brittle and that's OK — just don't make them more brittle.

6. **CacheService expiry** — Cached data expires. Code that assumes cache hits without fallback will fail intermittently.

---

## 🗣️ Communication Style

- Dan is hands-on and technical. Skip the theory, show the code.
- He prefers practical, immediately deployable solutions.
- When showing changes: exact diffs with file name and line context.
- Flag risks with `⚠️ HEADS UP:` — he'll decide whether to act on them.
- Don't over-explain. Don't add disclaimers. Don't suggest improvements he didn't ask for.
- If a task is too big for one session, say so and outline the pieces.

---

*Last updated: February 15, 2026*
