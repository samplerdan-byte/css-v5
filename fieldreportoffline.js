// ============================================================
// FIELD REPORT — OFFLINE MODE
// ============================================================
// Generates a standalone HTML file that works without internet.
// Edits are tracked locally and synced back when online.
// ============================================================

function generateOfflineFieldReport() {
  var reportData = getFieldReportData();
  var perms = getEditorPermissions();

  var warehouseNames = Object.keys(reportData.warehouses).sort();
  if (warehouseNames.length === 0) {
    return { success: false, html: '', message: 'No active orders found.' };
  }

  // Try to get the deployed web app URL for auto-sync
  var syncUrl = '';
  try {
    syncUrl = ScriptApp.getService().getUrl() || '';
  } catch (e) { /* not deployed as web app yet */ }

  var today = new Date();
  var dateStr = today.toLocaleDateString('en-US', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });
  var reportId = 'css_fr_' + today.toISOString().slice(0, 10);

  // Embed data as JSON
  var dataJson = JSON.stringify(reportData);
  var permsJson = JSON.stringify(perms);

  var html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n' +
    '<meta charset="utf-8">\n' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">\n' +
    '<title>CSS Field Report — ' + dateStr + ' (Offline)</title>\n' +
    '<style>\n' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }\n' +
    'body { font-family: Arial, sans-serif; display: flex; height: 100vh; background: #f5f5f5; }\n' +
    '.sidebar { width: 250px; min-width: 250px; background: #1a3a25; color: #fff; overflow-y: auto; padding: 12px 0; display: flex; flex-direction: column; }\n' +
    '.shdr { padding: 0 12px 12px; border-bottom: 1px solid rgba(255,255,255,0.1); margin-bottom: 8px; }\n' +
    '.stitle { font-size: 14px; font-weight: bold; color: #8cc99e; }\n' +
    '.ssub { font-size: 10px; color: #6b9e7d; margin-top: 2px; }\n' +
    '.slbl { font-size: 9px; color: #6b9e7d; text-transform: uppercase; padding: 6px 12px 3px; font-weight: bold; }\n' +
    '.si { padding: 8px 12px; cursor: pointer; border-left: 3px solid transparent; }\n' +
    '.si:hover { background: rgba(255,255,255,0.08); border-left-color: #4caf50; }\n' +
    '.sn { font-size: 12px; font-weight: bold; }\n' +
    '.ss { font-size: 9px; color: #8cc99e; margin-top: 1px; }\n' +
    '.sfoot { margin-top: auto; padding: 12px; border-top: 1px solid rgba(255,255,255,0.1); }\n' +
    '.btn { width: 100%; padding: 8px; border: none; border-radius: 5px; font-size: 12px; font-weight: bold; cursor: pointer; margin-top: 5px; }\n' +
    '.btn-sync { background: #e67e22; color: #fff; }\n' +
    '.btn-sync:hover { background: #d35400; }\n' +
    '.btn-sync:disabled { background: #666; cursor: default; }\n' +
    '.btn-copy { background: #2980b9; color: #fff; }\n' +
    '.btn-copy:hover { background: #1f6da1; }\n' +
    '.btn-print { background: #4A7C59; color: #fff; }\n' +
    '#pendingBadge { display: inline-block; background: #e74c3c; color: #fff; border-radius: 10px; padding: 1px 7px; font-size: 10px; margin-left: 4px; }\n' +
    '#connStatus { padding: 4px 8px; border-radius: 3px; font-size: 10px; text-align: center; margin-top: 8px; font-weight: bold; }\n' +
    '.conn-online { background: rgba(76,175,80,0.2); color: #4caf50; }\n' +
    '.conn-offline { background: rgba(231,76,60,0.2); color: #e74c3c; }\n' +
    '.main { flex: 1; overflow-y: auto; padding: 15px; }\n' +
    '.rpth { text-align: center; margin-bottom: 15px; padding: 15px; background: #fff; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }\n' +
    '.rptt { font-size: 18px; font-weight: bold; color: #2E5339; letter-spacing: 1px; }\n' +
    '.rpttag { font-size: 9px; color: #888; font-style: italic; }\n' +
    '.rpts { font-size: 14px; color: #2E5339; margin-top: 6px; }\n' +
    '.rptm { font-size: 11px; color: #666; margin-top: 3px; }\n' +
    '.ws { background: #fff; border-radius: 6px; margin-bottom: 15px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }\n' +
    '.wh { background: #2E5339; color: #fff; padding: 12px 16px; font-size: 16px; font-weight: bold; letter-spacing: 0.5px; }\n' +
    '.cb { border-bottom: 2px solid #ddd; }\n' +
    '.cb:last-child { border-bottom: none; }\n' +
    '.ch { padding: 6px 16px; font-weight: bold; font-size: 13px; color: #222; background: #f0f0f0; border: 1px solid #ccc; display: inline-block; margin: 10px 0 0 16px; border-radius: 3px; }\n' +
    '.ob { padding: 4px 16px 8px; border-bottom: 1px solid #eee; }\n' +
    '.ob:last-child { border-bottom: none; }\n' +
    '.ol { display: flex; justify-content: space-between; align-items: baseline; margin: 1px 0 5px; font-size: 11px; }\n' +
    '.on { font-weight: bold; color: #333; margin-right: 6px; }\n' +
    '.olink { color: #1a73e8; text-decoration: none; }\n' +
    '.olink:hover { text-decoration: underline; }\n' +
    '.ec { cursor: pointer; border-bottom: 1px dashed #2E5339; padding: 0 2px; }\n' +
    '.ec:hover { background: #e8f5e9; border-radius: 2px; }\n' +
    '.ec.edited { background: #fff3cd; border-bottom-color: #e67e22; }\n' +
    '.ep { color: #aaa; font-size: 9px; font-style: italic; }\n' +
    '.eip { border: 1px solid #2E5339; border-radius: 3px; padding: 2px 4px; font-size: 10px; font-family: Arial, sans-serif; outline: none; min-width: 60px; }\n' +
    '.eip:focus { box-shadow: 0 0 0 2px rgba(46,83,57,0.3); }\n' +
    '.si2 { color: #555; flex: 1; }\n' +
    '.ei { color: #c0392b; font-weight: bold; white-space: nowrap; }\n' +
    '.rh { padding: 4px 0 2px; font-size: 11px; font-weight: bold; color: #2E5339; border-top: 1px dashed #bbb; margin-top: 4px; }\n' +
    'table { width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 3px; }\n' +
    'th { background: #e8f5e9; color: #2E5339; padding: 3px 4px; text-align: left; font-size: 9px; border-bottom: 1px solid #c8e6c9; white-space: nowrap; }\n' +
    'td { padding: 3px 4px; border-bottom: 1px solid #f0f0f0; vertical-align: top; font-size: 10px; }\n' +
    '.dc { max-width: 220px; word-wrap: break-word; }\n' +
    '.dh { min-width: 120px; }\n' +
    '#toast { position: fixed; bottom: 15px; right: 15px; padding: 10px 18px; border-radius: 6px; font-weight: bold; font-size: 12px; z-index: 1000; display: none; box-shadow: 0 3px 10px rgba(0,0,0,0.2); }\n' +
    '.tok { background: #d4edda; color: #155724; display: block !important; }\n' +
    '.terr { background: #f8d7da; color: #721c24; display: block !important; }\n' +
    '.twait { background: #d1ecf1; color: #0c5460; display: block !important; }\n' +
    '@media print { .sidebar { display: none !important; } body { display: block; } .main { padding: 0; overflow: visible; } .ws { break-inside: avoid; box-shadow: none; border: 1px solid #999; } .rpth { box-shadow: none; border: 1px solid #999; } #toast { display: none !important; } .ec { border-bottom: none; cursor: default; } .ep { display: none; } .ec.edited { background: none; } }\n' +
    '</style>\n</head>\n<body>\n' +

    // ── SCRIPT WITH EMBEDDED DATA ──
    '<script>\n' +
    'var REPORT_DATA = ' + dataJson + ';\n' +
    'var PERMS = ' + permsJson + ';\n' +
    'var SYNC_URL = ' + JSON.stringify(syncUrl) + ';\n' +
    'var SYNC_TOKEN = ' + JSON.stringify(PropertiesService.getScriptProperties().getProperty('WEBAPP_ACTION_TOKEN') || '') + ';\n' +
    'var REPORT_ID = ' + JSON.stringify(reportId) + ';\n' +
    'var REPORT_DATE = ' + JSON.stringify(dateStr) + ';\n' +
    '\n' +
    '// ── Pending edits ──\n' +
    'var pendingEdits = [];\n' +
    'try { pendingEdits = JSON.parse(localStorage.getItem(REPORT_ID) || "[]"); } catch(e) {}\n' +
    'function saveEdits() {\n' +
    '  try { localStorage.setItem(REPORT_ID, JSON.stringify(pendingEdits)); } catch(e) {}\n' +
    '  updateBadge();\n' +
    '}\n' +
    '\n' +
    '// ── Connection status ──\n' +
    'function updateConnStatus() {\n' +
    '  var el = document.getElementById("connStatus");\n' +
    '  if (navigator.onLine) {\n' +
    '    el.className = "conn-online"; el.textContent = "🟢 Online";\n' +
    '  } else {\n' +
    '    el.className = "conn-offline"; el.textContent = "🔴 Offline";\n' +
    '  }\n' +
    '}\n' +
    'window.addEventListener("online", function() { updateConnStatus(); autoSync(); });\n' +
    'window.addEventListener("offline", updateConnStatus);\n' +
    '\n' +
    '// ── Badge ──\n' +
    'function updateBadge() {\n' +
    '  var el = document.getElementById("pendingBadge");\n' +
    '  if (pendingEdits.length > 0) {\n' +
    '    el.textContent = pendingEdits.length;\n' +
    '    el.style.display = "inline-block";\n' +
    '  } else {\n' +
    '    el.style.display = "none";\n' +
    '  }\n' +
    '  var btn = document.getElementById("syncBtn");\n' +
    '  if (btn) btn.disabled = pendingEdits.length === 0;\n' +
    '}\n' +
    '\n' +
    '// ── Toast ──\n' +
    'function toast(c, m) {\n' +
    '  var el = document.getElementById("toast");\n' +
    '  el.className = c; el.textContent = m;\n' +
    '  if (c !== "twait") setTimeout(function() { el.className = ""; el.style.display = "none"; }, 5000);\n' +
    '}\n' +
    '\n' +
    '// ── Edit cell ──\n' +
    'function canEditCol(colName) {\n' +
    '  if (!PERMS.canEdit) return false;\n' +
    '  if (PERMS.columns === "ALL") return true;\n' +
    '  return PERMS.columns.indexOf(colName) !== -1;\n' +
    '}\n' +
    '\n' +
    'function editCell(el, row, col) {\n' +
    '  if (el.querySelector("input")) return;\n' +
    '  var cur = el.querySelector("em") ? "" : (el.textContent || "").trim();\n' +
    '  var inp = document.createElement("input");\n' +
    '  inp.type = "text"; inp.className = "eip"; inp.value = cur;\n' +
    '  inp.style.width = Math.max(60, el.offsetWidth) + "px";\n' +
    '  el.textContent = ""; el.appendChild(inp); inp.focus(); inp.select();\n' +
    '  function save() {\n' +
    '    var nv = inp.value.trim();\n' +
    '    el.textContent = nv || "";\n' +
    '    if (!nv) el.innerHTML = \'<em class="ep">+ \' + col.toLowerCase().replace(/[^a-z]/g, " ").trim() + \'</em>\';\n' +
    '    if (nv !== cur) {\n' +
    '      el.classList.add("edited");\n' +
    '      pendingEdits.push({ row: row, column: col, value: nv, old: cur, ts: new Date().toISOString() });\n' +
    '      saveEdits();\n' +
    '      toast("tok", "Edit saved locally (" + pendingEdits.length + " pending)");\n' +
    '    }\n' +
    '  }\n' +
    '  inp.addEventListener("blur", save);\n' +
    '  inp.addEventListener("keydown", function(e) {\n' +
    '    if (e.key === "Enter") { e.preventDefault(); inp.blur(); }\n' +
    '    if (e.key === "Escape") { inp.value = cur; inp.blur(); }\n' +
    '  });\n' +
    '}\n' +
    '\n' +
    '// ── Sync edits ──\n' +
    'function syncEdits() {\n' +
    '  if (pendingEdits.length === 0) { toast("tok", "Nothing to sync"); return; }\n' +
    '  if (!navigator.onLine) { toast("terr", "No internet — edits saved locally, will sync when online"); return; }\n' +
    '  if (SYNC_URL) {\n' +
    '    postSync();\n' +
    '  } else {\n' +
    '    copyEditsToClipboard();\n' +
    '  }\n' +
    '}\n' +
    '\n' +
    'function postSync() {\n' +
    '  toast("twait", "Syncing " + pendingEdits.length + " edit(s)...");\n' +
    '  fetch(SYNC_URL, {\n' +
    '    method: "POST",\n' +
    '    headers: { "Content-Type": "text/plain;charset=utf-8" },\n' +
    '    body: JSON.stringify({ action: "syncFieldReport", token: SYNC_TOKEN, edits: pendingEdits })\n' +
    '  }).then(function(resp) {\n' +
    '    if (resp.ok) {\n' +
    '      return resp.json().then(function(data) {\n' +
    '        pendingEdits = [];\n' +
    '        saveEdits();\n' +
    '        toast("tok", "✅ " + (data.message || "All edits synced!"));\n' +
    '      });\n' +
    '    } else {\n' +
    '      toast("terr", "Server returned error " + resp.status + " — edits still saved locally. Use Copy to Clipboard.");\n' +
    '    }\n' +
    '  }).catch(function(err) {\n' +
    '    toast("terr", "Can\\x27t reach server (CORS/network) — edits still saved locally. Use Copy to Clipboard.");\n' +
    '  });\n' +
    '}\n' +
    '\n' +
    'function autoSync() {\n' +
    '  if (pendingEdits.length > 0 && SYNC_URL && navigator.onLine) {\n' +
    '    toast("twait", "Back online — auto-syncing " + pendingEdits.length + " edit(s)...");\n' +
    '    setTimeout(postSync, 1500);\n' +
    '  }\n' +
    '}\n' +
    '\n' +
    'function copyEditsToClipboard() {\n' +
    '  if (pendingEdits.length === 0) { toast("tok", "Nothing to copy"); return; }\n' +
    '  var json = JSON.stringify(pendingEdits, null, 2);\n' +
    '  navigator.clipboard.writeText(json).then(function() {\n' +
    '    toast("tok", "✅ " + pendingEdits.length + " edit(s) copied! Paste via CSS System → Reports → Import Offline Edits");\n' +
    '  }).catch(function() {\n' +
    '    // Fallback: show in a textarea\n' +
    '    prompt("Copy this text, then paste via Import Offline Edits:", json);\n' +
    '  });\n' +
    '}\n' +
    '\n' +
    'function clearEdits() {\n' +
    '  if (pendingEdits.length === 0) return;\n' +
    '  if (!confirm("Discard " + pendingEdits.length + " unsaved edit(s)? This cannot be undone.")) return;\n' +
    '  pendingEdits = [];\n' +
    '  saveEdits();\n' +
    '  toast("tok", "Edits cleared");\n' +
    '}\n' +
    '\n' +
    '// ── Scroll ──\n' +
    'function scrollTo(n) {\n' +
    '  var e = document.getElementById("wh-" + n.replace(/[^a-zA-Z0-9]/g, "_"));\n' +
    '  if (e) e.scrollIntoView({ behavior: "smooth", block: "start" });\n' +
    '}\n' +
    '\n' +
    '// ── Render ──\n' +
    'function esc(s) { var d = document.createElement("div"); d.textContent = s; return d.innerHTML; }\n' +
    '\n' +
    'function edt(row, colName, val, placeholder) {\n' +
    '  if (!canEditCol(colName)) return esc(val || "");\n' +
    '  if (val) return \'<span class="ec" onclick="editCell(this,\' + row + \',&#39;\' + colName + \'&#39;)" title="Click to edit \' + colName + \'">\' + esc(val) + \'</span>\';\n' +
    '  return \'<span class="ec" onclick="editCell(this,\' + row + \',&#39;\' + colName + \'&#39;)" title="Click to edit \' + colName + \'"><em class="ep">+ \' + (placeholder || colName.toLowerCase()) + \'</em></span>\';\n' +
    '}\n' +
    '\n' +
    'function renderReport() {\n' +
    '  var whs = REPORT_DATA.warehouses;\n' +
    '  var whNames = Object.keys(whs).sort();\n' +
    '  var grandTotal = 0;\n' +
    '\n' +
    '  // Sidebar\n' +
    '  var sbHtml = "";\n' +
    '  whNames.forEach(function(wh) {\n' +
    '    var custs = whs[wh]; var count = 0; var cc = Object.keys(custs).length;\n' +
    '    for (var c in custs) for (var o in custs[c]) { var ord = custs[c][o]; for (var r in ord.receivers) count += ord.receivers[r].length; }\n' +
    '    grandTotal += count;\n' +
    '    sbHtml += \'<div class="si" onclick="scrollTo(&#39;\' + wh + \'&#39;)">\' +\n' +
    '      \'<div class="sn">\' + esc(wh) + \'</div>\' +\n' +
    '      \'<div class="ss">\' + count + \' samples &middot; \' + cc + \' clients</div></div>\';\n' +
    '  });\n' +
    '\n' +
    '  document.getElementById("whList").innerHTML = sbHtml;\n' +
    '  document.getElementById("totalSamples").textContent = grandTotal + " active samples";\n' +
    '\n' +
    '  // Main content\n' +
    '  var mainHtml = "";\n' +
    '  whNames.forEach(function(wh) {\n' +
    '    var whId = wh.replace(/[^a-zA-Z0-9]/g, "_");\n' +
    '    var custNames = Object.keys(whs[wh]).sort();\n' +
    '    var custHtml = "";\n' +
    '\n' +
    '    custNames.forEach(function(cust) {\n' +
    '      var orders = whs[wh][cust];\n' +
    '      var orderKeys = Object.keys(orders).sort(function(a, b) { return (orders[a].sortDate || 0) - (orders[b].sortDate || 0); });\n' +
    '      var ordHtml = "";\n' +
    '\n' +
    '      orderKeys.forEach(function(orderNum) {\n' +
    '        var order = orders[orderNum];\n' +
    '        var recvNames = Object.keys(order.receivers);\n' +
    '        var multi = recvNames.length > 1;\n' +
    '        var oRow = order.firstRow;\n' +
    '\n' +
    '        var onHtml = order.emailLink ?\n' +
    '          \'<a href="\' + order.emailLink + \'" target="_blank" class="on olink">\' + esc(orderNum) + \'</a>\' :\n' +
    '          \'<span class="on">\' + esc(orderNum) + \'</span>\';\n' +
    '\n' +
    '        var parts = [];\n' +
    '        var sl = edt(oRow, "Shipping Line", order.shippingLine, "line"); if (sl) parts.push(sl);\n' +
    '        var sn = edt(oRow, "Shipping Notes", order.shippingNotes, "notes"); if (sn) parts.push(sn);\n' +
    '        var vs = edt(oRow, "Container Status", order.vesselStatus, "status"); if (vs) parts.push(vs);\n' +
    '        var shipStr = parts.join(" ");\n' +
    '        var etaH = edt(oRow, "Container ETA", order.eta, "eta");\n' +
    '\n' +
    '        var recvHtml = "";\n' +
    '        recvNames.forEach(function(recv) {\n' +
    '          var samples = order.receivers[recv];\n' +
    '          if (multi) recvHtml += \'<div class="rh">→ \' + esc(recv) + \'</div>\';\n' +
    '          var rows = "";\n' +
    '          samples.forEach(function(s) {\n' +
    '            var r = s.rowNumber;\n' +
    '            var rc = "";\n' +
    '            if (!multi) { rc = "<td>" + edt(r, "Receiver", recv !== "Unassigned" ? recv : "", "receiver") + "</td>"; }\n' +
    '            rows += "<tr>" +\n' +
    '              "<td>" + edt(r, "Reference", s.reference, "ref") + "</td>" +\n' +
    '              \'<td class="dc">\' + edt(r, "Description", s.description, "desc") + "</td>" +\n' +
    '              "<td>" + edt(r, "Mark #", s.mark, "mark") + "</td>" +\n' +
    '              "<td>" + edt(r, "Container #", s.container, "container") + "</td>" +\n' +
    '              "<td>" + edt(r, "Cargo #", s.cargo, "cargo") + "</td>" +\n' +
    '              rc +\n' +
    '              "<td>" + edt(r, "Bag Count", s.bagCount, "bags") + "</td>" +\n' +
    '              "<td>" + edt(r, "Sample Weight", s.sampleWeight, "wt") + "</td>" +\n' +
    '            "</tr>";\n' +
    '          });\n' +
    '          var rth = !multi ? "<th>Receiver</th>" : "";\n' +
    '          recvHtml += \'<table><thead><tr><th>Ref</th><th class="dh">Description</th><th>Mark</th><th>Container</th><th>Cargo</th>\' + rth + \'<th>Bags</th><th>Wt</th></tr></thead><tbody>\' + rows + \'</tbody></table>\';\n' +
    '        });\n' +
    '\n' +
    '        ordHtml += \'<div class="ob"><div class="ol">\' + onHtml +\n' +
    '          (shipStr ? \' <span class="si2">\' + shipStr + \'</span>\' : "") +\n' +
    '          \'<span class="ei">ETA: \' + etaH + \'</span></div>\' + recvHtml + \'</div>\';\n' +
    '      });\n' +
    '\n' +
    '      custHtml += \'<div class="cb"><div class="ch">\' + esc(cust) + \'</div>\' + ordHtml + \'</div>\';\n' +
    '    });\n' +
    '\n' +
    '    mainHtml += \'<div class="ws" id="wh-\' + whId + \'"><div class="wh">\' + esc(wh) + \'</div>\' + custHtml + \'</div>\';\n' +
    '  });\n' +
    '\n' +
    '  document.getElementById("reportBody").innerHTML = mainHtml;\n' +
    '}\n' +
    '\n' +
    '// ── Init ──\n' +
    'document.addEventListener("DOMContentLoaded", function() {\n' +
    '  renderReport();\n' +
    '  updateBadge();\n' +
    '  updateConnStatus();\n' +
    '});\n' +
    '</script>\n' +

    // ── SIDEBAR HTML ──
    '<div class="sidebar">\n' +
    '<div class="shdr">\n' +
    '  <div class="stitle">🏭 Field Report</div>\n' +
    '  <div class="ssub">' + dateStr + '</div>\n' +
    '  <div class="ssub" id="totalSamples"></div>\n' +
    '  <div class="ssub" style="color:#e67e22;margin-top:4px;">📥 Offline Mode</div>\n' +
    (perms.canEdit ? '  <div class="ssub" style="color:#4caf50;">✏️ Edit enabled' + (perms.columns === 'ALL' ? ' (full)' : '') + '</div>\n' : '') +
    '  <div id="connStatus"></div>\n' +
    '</div>\n' +
    '<div class="slbl">WAREHOUSES</div>\n' +
    '<div id="whList"></div>\n' +
    '<div class="sfoot">\n' +
    '  <div style="font-size:10px;color:#8cc99e;margin-bottom:6px;">Pending edits: <span id="pendingBadge">0</span></div>\n' +
    '  <button class="btn btn-sync" id="syncBtn" onclick="syncEdits()" disabled>🔄 Sync Edits</button>\n' +
    '  <button class="btn btn-copy" onclick="copyEditsToClipboard()">📋 Copy Edits to Clipboard</button>\n' +
    '  <button class="btn btn-print" onclick="window.print()">🖨️ Print Report</button>\n' +
    '  <button class="btn" style="background:rgba(255,255,255,0.1);color:#ccc;" onclick="clearEdits()">🗑️ Discard Edits</button>\n' +
    '</div>\n' +
    '</div>\n' +

    // ── MAIN CONTENT ──
    '<div class="main">\n' +
    '<div class="rpth">\n' +
    '  <div class="rptt">COMMODITY SAMPLER SERVICES</div>\n' +
    '  <div class="rpttag">"Integrity Through Independence"</div>\n' +
    '  <div class="rpts">🏭 Field Report — Offline Copy</div>\n' +
    '  <div class="rptm">' + dateStr + '</div>\n' +
    '</div>\n' +
    '<div id="reportBody"></div>\n' +
    '</div>\n' +
    '<div id="toast"></div>\n' +
    '</body>\n</html>';

  return { success: true, html: html, message: 'Offline report ready' };
}


// ============================================================
// IMPORT OFFLINE EDITS
// ============================================================
function showImportOfflineEdits() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: Arial, sans-serif; padding: 16px; background: #f5f5f5; }' +
    'h2 { color: #2E5339; font-size: 16px; margin-bottom: 10px; }' +
    'p { font-size: 12px; color: #666; margin-bottom: 10px; }' +
    'textarea { width: 100%; height: 280px; border: 1px solid #ccc; border-radius: 6px; padding: 10px; font-family: monospace; font-size: 11px; resize: vertical; }' +
    'textarea:focus { outline: none; border-color: #2E5339; box-shadow: 0 0 0 2px rgba(46,83,57,0.2); }' +
    '.btn { width: 100%; padding: 12px; border: none; border-radius: 6px; font-size: 14px; font-weight: bold; cursor: pointer; margin-top: 8px; }' +
    '.btn-go { background: #2E5339; color: #fff; }' +
    '.btn-go:hover { background: #3a6249; }' +
    '.btn-go:disabled { background: #ccc; }' +
    '#result { padding: 10px; border-radius: 6px; margin-top: 10px; font-size: 12px; display: none; }' +
    '.rok { background: #d4edda; color: #155724; display: block !important; }' +
    '.rerr { background: #f8d7da; color: #721c24; display: block !important; }' +
    '</style>' +
    '<h2>📥 Import Offline Edits</h2>' +
    '<p>Paste the JSON copied from the offline field report:</p>' +
    '<textarea id="editJson" placeholder=\'[{ "row": 5, "column": "Container Status", "value": "Discharged At Pier" }, ...]\'></textarea>' +
    '<button class="btn btn-go" id="goBtn" onclick="doImport()">Apply Edits</button>' +
    '<div id="result"></div>' +
    '<script>' +
    'function doImport() {' +
    '  var json = document.getElementById("editJson").value.trim();' +
    '  if (!json) { show("rerr", "Paste the JSON first"); return; }' +
    '  var btn = document.getElementById("goBtn");' +
    '  btn.disabled = true; btn.textContent = "Applying...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r) {' +
    '      btn.disabled = false; btn.textContent = "Apply Edits";' +
    '      show(r.success ? "rok" : "rerr", r.message);' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      btn.disabled = false; btn.textContent = "Apply Edits";' +
    '      show("rerr", "Error: " + e.message);' +
    '    })' +
    '    .importOfflineEdits(json);' +
    '}' +
    'function show(cls, msg) { var el = document.getElementById("result"); el.className = cls; el.textContent = msg; }' +
    '</script>'
  ).setWidth(480).setHeight(520).setTitle('Import Offline Edits');

  SpreadsheetApp.getUi().showModalDialog(html, '📥 Import Offline Edits');
}

function importOfflineEdits(jsonStr) {
  var perms = getEditorPermissions();
  if (!perms.canEdit) return { success: false, message: 'You do not have edit access.' };

  var edits;
  try {
    edits = JSON.parse(jsonStr);
  } catch (e) {
    return { success: false, message: 'Invalid JSON: ' + e.message };
  }

  if (!Array.isArray(edits) || edits.length === 0) {
    return { success: false, message: 'No edits found. Expected a JSON array.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.mainSheetName);
  if (!sheet) return { success: false, message: 'Main sheet not found.' };

  var col = _getColumnMap(sheet);
  var applied = 0;
  var skipped = 0;
  var errors = [];

  for (var i = 0; i < edits.length; i++) {
    var edit = edits[i];
    if (!edit.row || !edit.column) {
      skipped++;
      continue;
    }

    // Check column permission
    if (perms.columns !== 'ALL' && perms.columns.indexOf(edit.column) === -1) {
      errors.push('Row ' + edit.row + ': no permission for "' + edit.column + '"');
      skipped++;
      continue;
    }

    var colIdx = col[edit.column];
    if (colIdx === undefined) {
      errors.push('Column "' + edit.column + '" not found');
      skipped++;
      continue;
    }

    try {
      sheet.getRange(edit.row, colIdx + 1).setValue(edit.value || '');
      applied++;
    } catch (e) {
      errors.push('Row ' + edit.row + ': ' + e.message);
      skipped++;
    }
  }

  var msg = '✅ Applied ' + applied + ' edit(s)';
  if (skipped > 0) msg += ', skipped ' + skipped;
  if (errors.length > 0) msg += '\n\nIssues:\n' + errors.slice(0, 5).join('\n');

  return { success: true, message: msg };
}


// ============================================================
// WEB APP — doPost handler for offline sync
// ============================================================
// Note: doGet already exists in inputManual.gs for DataEntry.
// doPost handles sync requests from offline field reports.
// The project must be deployed as a web app for this to work.
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // Auth check — require token for all POST actions
    var token = payload.token || '';
    var expectedToken = PropertiesService.getScriptProperties().getProperty('WEBAPP_ACTION_TOKEN');
    if (!expectedToken || token !== expectedToken) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'syncFieldReport' && Array.isArray(payload.edits)) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.mainSheetName);
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Sheet not found' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var col = _getColumnMap(sheet);
      var applied = 0;

      for (var i = 0; i < payload.edits.length; i++) {
        var edit = payload.edits[i];
        if (!edit.row || !edit.column) continue;
        var colIdx = col[edit.column];
        if (colIdx === undefined) continue;
        try {
          sheet.getRange(edit.row, colIdx + 1).setValue(edit.value || '');
          applied++;
        } catch (err) { /* skip individual failures */ }
      }

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: 'Applied ' + applied + ' of ' + payload.edits.length + ' edit(s)'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
