// ============================================================
// 📸 PHOTO MODULE — Sample Photo Upload & Management
// ============================================================
// Server-side handlers for web app photo feature
// Photos stored in Google Drive under CSS_Photos/{CS_Sample_#}/
// Photo links written back to sample row in "Photos" column
//
// UPDATED Feb 15, 2026:
//   - Lookup by CS Sample #, CS Order #, Cargo #, or Container #
//   - Multi-match returns list for user to pick from
//   - Bulk upload to all samples in an order
// ============================================================

var PHOTO_CONFIG = {
  rootFolderName: 'CSS_Photos',
  photoColumnName: 'Photos',
  thumbnailSize: 200,
  // Columns to search (in priority order)
  searchColumns: ['CS Sample #', 'CS Order #', 'Cargo #', 'Container #']
};


// ============================================================
// Flexible lookup — find samples by any identifier
// Returns { success, matches: [{ sampleId, order, cargo, container, desc, status }] }
// ============================================================

function webAppFindSamples(query) {
  try {
    if (!query || !String(query).trim()) {
      return { success: false, message: 'Enter a sample #, order #, cargo #, or container #' };
    }
    query = String(query).trim();

    var ss = (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID)
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    var mainName = (typeof CONFIG !== 'undefined' && CONFIG.mainSheetName) ? CONFIG.mainSheetName : 'All Orders';
    var completedName = (typeof CONFIG !== 'undefined' && CONFIG.completedOrdersSheetName) ? CONFIG.completedOrdersSheetName : 'Completed Orders';
    var sheets = [
      ss.getSheetByName(mainName),
      ss.getSheetByName(completedName)
    ];

    var matches = [];
    var seenSamples = {};
    var queryLower = query.toLowerCase();

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      if (!sheet || sheet.getLastRow() < 2) continue;

      var col = _getColumnMap(sheet);
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

      for (var i = 0; i < data.length; i++) {
        var sampleId = col['CS Sample #'] !== undefined ? String(data[i][col['CS Sample #']] || '').trim() : '';
        if (!sampleId || seenSamples[sampleId]) continue;

        var hit = false;

        // Check each searchable column (exact match first, then contains)
        for (var c = 0; c < PHOTO_CONFIG.searchColumns.length; c++) {
          var colName = PHOTO_CONFIG.searchColumns[c];
          if (col[colName] === undefined) continue;
          var cellVal = String(data[i][col[colName]] || '').trim();
          if (!cellVal) continue;
          var cellLower = cellVal.toLowerCase();
          // Exact match OR cell contains query
          if (cellLower === queryLower || cellLower.indexOf(queryLower) !== -1) {
            hit = true;
            break;
          }
        }

        if (hit) {
          seenSamples[sampleId] = true;

          // Check Photos column for existing link (fast — no Drive API call)
          var photoLink = '';
          if (col[PHOTO_CONFIG.photoColumnName] !== undefined) {
            photoLink = String(data[i][col[PHOTO_CONFIG.photoColumnName]] || '').trim();
          }

          matches.push({
            sampleId: sampleId,
            order: col['CS Order #'] !== undefined ? String(data[i][col['CS Order #']] || '').trim() : '',
            cargo: col['Cargo #'] !== undefined ? String(data[i][col['Cargo #']] || '').trim() : '',
            container: col['Container #'] !== undefined ? String(data[i][col['Container #']] || '').trim() : '',
            mark: col['Mark #'] !== undefined ? String(data[i][col['Mark #']] || '').trim() : '',
            description: col['Description'] !== undefined ? String(data[i][col['Description']] || '').trim() : '',
            sender: col['Sender'] !== undefined ? String(data[i][col['Sender']] || '').trim() : '',
            status: col['Status'] !== undefined ? String(data[i][col['Status']] || '').trim() : '',
            warehouse: col['Warehouse'] !== undefined ? String(data[i][col['Warehouse']] || '').trim() : '',
            hasPhotos: !!photoLink,
            row: i + 2
          });
        }
      }
    }

    if (matches.length === 0) {
      return { success: false, message: '❌ No samples found matching "' + query + '"' };
    }

    return { success: true, matches: matches, query: query };

  } catch (e) {
    Logger.log('webAppFindSamples error: ' + e);
    return { success: false, message: '❌ Error: ' + e.message };
  }
}


// ============================================================
// Load photos for a specific sample (by CS Sample #)
// ============================================================

function webAppLoadSamplePhotos(sampleId) {
  try {
    if (!sampleId) return { success: false, message: 'No sample ID provided' };
    sampleId = String(sampleId).trim();

    var photos = [];
    var folder = _getPhotoFolder(sampleId, false);
    if (folder) {
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var mimeType = file.getMimeType();
        if (mimeType && mimeType.indexOf('image') !== -1) {
          photos.push({
            name: file.getName(),
            url: file.getUrl(),
            thumbnail: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w' + PHOTO_CONFIG.thumbnailSize,
            id: file.getId(),
            date: file.getDateCreated().toLocaleDateString()
          });
        }
      }
    }

    return {
      success: true,
      sampleId: sampleId,
      photos: photos,
      folderUrl: folder ? folder.getUrl() : null
    };

  } catch (e) {
    Logger.log('webAppLoadSamplePhotos error: ' + e);
    return { success: false, message: '❌ Error: ' + e.message };
  }
}


// ============================================================
// Upload a photo to one or more samples
// sampleIds can be a single string or an array
// ============================================================

function webAppUploadPhoto(sampleIds, base64Data, fileName, mimeType) {
  try {
    if (!sampleIds || !base64Data) {
      return { success: false, message: 'Missing sample ID or photo data' };
    }

    // Normalize to array
    if (typeof sampleIds === 'string') sampleIds = [sampleIds];

    mimeType = mimeType || 'image/jpeg';
    var decoded = Utilities.base64Decode(base64Data);

    var ext = '.jpg';
    if (mimeType.indexOf('png') !== -1) ext = '.png';
    else if (mimeType.indexOf('webp') !== -1) ext = '.webp';

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var results = [];

    for (var i = 0; i < sampleIds.length; i++) {
      var sampleId = String(sampleIds[i]).trim();
      if (!sampleId) continue;

      try {
        var folder = _getPhotoFolder(sampleId, true);
        var cleanName = sampleId + '_' + timestamp + ext;
        var blob = Utilities.newBlob(decoded, mimeType, cleanName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        _updatePhotoLink(sampleId, folder);

        results.push({
          sampleId: sampleId,
          success: true,
          fileUrl: file.getUrl(),
          fileName: cleanName
        });
      } catch (e) {
        results.push({
          sampleId: sampleId,
          success: false,
          message: e.message
        });
      }
    }

    var okCount = results.filter(function(r) { return r.success; }).length;
    return {
      success: okCount > 0,
      message: '✅ Photo saved to ' + okCount + ' of ' + sampleIds.length + ' sample(s)',
      results: results
    };

  } catch (e) {
    Logger.log('webAppUploadPhoto error: ' + e);
    if (typeof logError === 'function') logError('webAppUploadPhoto', e.message, { stack: e.stack, sampleCount: (sampleIds || []).length });
    return { success: false, message: '❌ Upload error: ' + e.message };
  }
}


// ============================================================
// Delete a photo by file ID
// ============================================================

function webAppDeletePhoto(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var name = file.getName();
    file.setTrashed(true);
    return { success: true, message: '🗑️ Deleted: ' + name };
  } catch (e) {
    return { success: false, message: '❌ Delete error: ' + e.message };
  }
}


// ============================================================
// HELPER: Get or create the photo folder for a sample
// ============================================================

function _getPhotoFolder(sampleId, create) {
  var rootFolder = _getPhotoRootFolder(create);
  if (!rootFolder) return null;

  var folders = rootFolder.getFoldersByName(sampleId);
  if (folders.hasNext()) return folders.next();

  if (create) {
    var folder = rootFolder.createFolder(sampleId);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return folder;
  }

  return null;
}

function _getPhotoRootFolder(create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssFile = DriveApp.getFileById(ss.getId());
  var parents = ssFile.getParents();
  var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  var folders = parentFolder.getFoldersByName(PHOTO_CONFIG.rootFolderName);
  if (folders.hasNext()) return folders.next();

  if (create) {
    return parentFolder.createFolder(PHOTO_CONFIG.rootFolderName);
  }
  return null;
}


// ============================================================
// HELPER: Update the "Photos" column with Drive folder link
// ============================================================

function _updatePhotoLink(sampleId, folder) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainName = (typeof CONFIG !== 'undefined' && CONFIG.mainSheetName) ? CONFIG.mainSheetName : 'All Orders';
  var completedName = (typeof CONFIG !== 'undefined' && CONFIG.completedOrdersSheetName) ? CONFIG.completedOrdersSheetName : 'Completed Orders';
  var sheets = [
    ss.getSheetByName(mainName),
    ss.getSheetByName(completedName)
  ];

  var folderUrl = folder.getUrl();

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (!sheet || sheet.getLastRow() < 2) continue;

    var col = _getColumnMap(sheet);
    var sampleIdx = col['CS Sample #'];
    var photoIdx = col[PHOTO_CONFIG.photoColumnName];

    if (sampleIdx === undefined) continue;

    // If no Photos column, create it
    if (photoIdx === undefined) {
      var lastCol = sheet.getLastColumn();
      sheet.getRange(1, lastCol + 1).setValue(PHOTO_CONFIG.photoColumnName)
        .setFontWeight('bold').setBackground('#2E5339').setFontColor('#fff');
      col = _getColumnMap(sheet);
      photoIdx = col[PHOTO_CONFIG.photoColumnName];
    }

    var data = sheet.getRange(2, sampleIdx + 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === sampleId) {
        sheet.getRange(i + 2, photoIdx + 1).setValue(folderUrl);
        return;
      }
    }
  }
}


// ============================================================
// HELPER: Get current Photos column value for a sample
// ============================================================

function _getPhotoColumnValue(sampleId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainName = (typeof CONFIG !== 'undefined' && CONFIG.mainSheetName) ? CONFIG.mainSheetName : 'All Orders';
  var completedName = (typeof CONFIG !== 'undefined' && CONFIG.completedOrdersSheetName) ? CONFIG.completedOrdersSheetName : 'Completed Orders';
  var sheets = [
    ss.getSheetByName(mainName),
    ss.getSheetByName(completedName)
  ];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (!sheet || sheet.getLastRow() < 2) continue;

    var col = _getColumnMap(sheet);
    var sampleIdx = col['CS Sample #'];
    var photoIdx = col[PHOTO_CONFIG.photoColumnName];
    if (sampleIdx === undefined || photoIdx === undefined) continue;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][sampleIdx]).trim() === sampleId) {
        return String(data[i][photoIdx] || '').trim();
      }
    }
  }
  return '';
}


// ============================================================
// SIDEBAR LAUNCHER — Photo Upload UI
// ============================================================

function showPhotoUpload() {
  var html = HtmlService.createHtmlOutput(_getPhotoUploadHTML())
    .setWidth(420)
    .setHeight(700)
    .setTitle('📸 Sample Photos');
  SpreadsheetApp.getUi().showSidebar(html);
}


// ============================================================
// PHOTO UPLOAD SIDEBAR HTML
// ============================================================

function _getPhotoUploadHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; padding: 12px; background: #f5f5f5; }
    h2 { text-align: center; color: #2E5339; margin-bottom: 10px; font-size: 18px; }

    /* SEARCH */
    .search-box { display: flex; gap: 6px; margin-bottom: 10px; }
    .search-box input {
      flex: 1; padding: 12px; font-size: 16px; text-align: center;
      border: 3px solid #2E5339; border-radius: 10px; background: #fff;
    }
    .search-box input:focus { outline: none; border-color: #4A7C59; box-shadow: 0 0 0 3px rgba(74,124,89,0.25); }
    .search-box button {
      padding: 12px 16px; font-size: 16px; font-weight: bold;
      background: #2E5339; color: #fff; border: none; border-radius: 10px; cursor: pointer;
    }
    .search-hint { text-align: center; font-size: 10px; color: #999; margin-bottom: 10px; }

    /* STATUS */
    #status {
      padding: 8px; border-radius: 6px; text-align: center;
      font-size: 13px; font-weight: bold; margin-bottom: 10px; min-height: 32px;
    }
    .s-ready { background: #e8f5e9; color: #2E5339; }
    .s-loading { background: #e3f2fd; color: #1565c0; }
    .s-success { background: #d4edda; color: #155724; }
    .s-error { background: #f8d7da; color: #721c24; }

    /* MATCHES LIST */
    #matchSection { display: none; }
    .match-header {
      display: flex; justify-content: space-between; align-items: center;
      padding: 6px 0; border-bottom: 1px solid #ddd; margin-bottom: 6px;
    }
    .match-header h3 { font-size: 12px; color: #333; }
    .match-header span { font-size: 10px; color: #999; }
    .match-item {
      display: flex; align-items: center; padding: 8px; margin-bottom: 4px;
      background: #fff; border-radius: 6px; border: 2px solid #e0e0e0;
      cursor: pointer; transition: all 0.15s; font-size: 11px;
    }
    .match-item:hover { border-color: #4A7C59; background: #f0faf0; }
    .match-item.selected { border-color: #2E5339; background: #e8f5e9; }
    .match-cb { margin-right: 8px; width: 16px; height: 16px; accent-color: #2E5339; }
    .match-info { flex: 1; line-height: 1.4; }
    .match-sample { font-weight: bold; color: #1e3c72; }
    .match-detail { color: #666; font-size: 9px; }
    .match-photos { font-size: 9px; color: #4A7C59; font-weight: bold; }
    .select-all-row { margin-bottom: 6px; }
    .select-all-row label { font-size: 10px; color: #666; cursor: pointer; }

    /* UPLOAD AREA */
    #uploadSection { display: none; }
    .upload-target {
      border: 3px dashed #ccc; border-radius: 10px; padding: 30px 15px;
      text-align: center; cursor: pointer; transition: all 0.2s;
      margin-bottom: 10px; background: #fff;
    }
    .upload-target:hover, .upload-target.dragover {
      border-color: #4A7C59; background: #f0faf0;
    }
    .upload-target .icon { font-size: 40px; margin-bottom: 6px; }
    .upload-target .label { font-size: 13px; color: #666; }
    .upload-target .sublabel { font-size: 10px; color: #999; margin-top: 3px; }
    #fileInput { display: none; }

    /* PREVIEW */
    #previewSection { display: none; margin-bottom: 10px; }
    .preview-img {
      max-width: 100%; max-height: 200px; border-radius: 8px;
      display: block; margin: 0 auto 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }
    .preview-name { text-align: center; font-size: 10px; color: #666; margin-bottom: 8px; }
    .preview-actions { display: flex; gap: 6px; }
    .btn-upload {
      flex: 1; padding: 12px; font-size: 14px; font-weight: bold;
      background: #2E5339; color: #fff; border: none; border-radius: 8px; cursor: pointer;
    }
    .btn-upload:disabled { opacity: 0.5; cursor: not-allowed; }
    .btn-cancel {
      padding: 12px 16px; font-size: 14px;
      background: #f5f5f5; color: #666; border: 1px solid #ddd; border-radius: 8px; cursor: pointer;
    }

    /* EXISTING PHOTOS */
    #existingSection { display: none; margin-top: 12px; }
    .photos-header {
      display: flex; justify-content: space-between; align-items: center;
      padding: 6px 0; border-bottom: 1px solid #ddd; margin-bottom: 6px;
    }
    .photos-header h3 { font-size: 12px; color: #333; }
    .photos-header a { font-size: 10px; color: #1976d2; }
    .photo-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
    .photo-thumb {
      position: relative; border-radius: 6px; overflow: hidden;
      background: #eee; aspect-ratio: 1;
    }
    .photo-thumb img {
      width: 100%; height: 100%; object-fit: cover; display: block; cursor: pointer;
    }
    .photo-thumb .delete-btn {
      position: absolute; top: 3px; right: 3px;
      background: rgba(0,0,0,0.6); color: #fff; border: none;
      border-radius: 50%; width: 20px; height: 20px; font-size: 11px;
      cursor: pointer; display: none; line-height: 20px; text-align: center;
    }
    .photo-thumb:hover .delete-btn { display: block; }
    .photo-date { font-size: 8px; color: #999; text-align: center; margin-top: 2px; }

    /* CAMERA */
    .camera-btn {
      width: 100%; padding: 10px; font-size: 13px; font-weight: bold;
      background: #fff; color: #2E5339; border: 2px solid #2E5339;
      border-radius: 8px; cursor: pointer; margin-bottom: 8px;
    }
    .camera-btn:hover { background: #e8f5e9; }

    /* PROGRESS */
    .progress-bar {
      height: 4px; background: #e0e0e0; border-radius: 2px;
      margin-bottom: 8px; overflow: hidden; display: none;
    }
    .progress-bar .fill {
      height: 100%; background: #4A7C59; border-radius: 2px;
      transition: width 0.3s ease;
    }
  </style>
</head>
<body>

  <h2>📸 Sample Photos</h2>

  <!-- SEARCH -->
  <form onsubmit="doSearch(); return false;" autocomplete="off">
    <div class="search-box">
      <input type="text" id="queryBox" placeholder="Scan or type ID..."
             autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
      <button type="submit" id="searchBtn">🔍</button>
    </div>
  </form>
  <div class="search-hint">Sample #, Order #, Cargo #, or Container #</div>

  <div id="status" class="s-ready">Ready — scan or type an identifier</div>
  <div class="progress-bar" id="progressBar"><div class="fill" id="progressFill"></div></div>

  <!-- MATCHES -->
  <div id="matchSection">
    <div class="match-header">
      <h3>Matching Samples</h3>
      <span id="matchCount"></span>
    </div>
    <div class="select-all-row" id="selectAllRow" style="display:none;">
      <label><input type="checkbox" id="selectAll" onchange="toggleSelectAll()"> Select all — upload photo to entire group</label>
    </div>
    <div id="matchList"></div>
  </div>

  <!-- UPLOAD AREA -->
  <div id="uploadSection">
    <button class="camera-btn" id="cameraBtn" onclick="openCamera()">📷 Take Photo</button>
    <div class="upload-target" id="dropZone" onclick="document.getElementById('fileInput').click()">
      <div class="icon">📁</div>
      <div class="label">Tap to choose file or drag & drop</div>
      <div class="sublabel">JPG, PNG, WEBP</div>
    </div>
    <input type="file" id="fileInput" accept="image/*" capture="environment" onchange="handleFile(this.files[0])">
  </div>

  <!-- PREVIEW -->
  <div id="previewSection">
    <img id="previewImg" class="preview-img">
    <div id="previewName" class="preview-name"></div>
    <div class="preview-actions">
      <button class="btn-upload" id="uploadBtn" onclick="doUpload()">📤 Upload</button>
      <button class="btn-cancel" onclick="cancelPreview()">✕</button>
    </div>
  </div>

  <!-- EXISTING PHOTOS -->
  <div id="existingSection">
    <div class="photos-header">
      <h3 id="existingTitle">Photos</h3>
      <a id="folderLink" href="#" target="_blank">Open folder ↗</a>
    </div>
    <div class="photo-grid" id="photoGrid"></div>
  </div>

<script>
/* ============================================================
   STATE
   ============================================================ */
var selectedSamples = [];   /* sample IDs checked for upload */
var allMatches = [];        /* full match data from search */
var pendingFile = null;     /* { base64, name, mimeType } */
var currentViewSample = null;

/* ============================================================
   SEARCH
   ============================================================ */
function doSearch() {
  var q = document.getElementById("queryBox").value.trim();
  if (!q) return;

  setStatus("s-loading", "Searching...");
  hideAll();

  google.script.run
    .withSuccessHandler(function(r) {
      if (!r.success) {
        setStatus("s-error", r.message);
        return;
      }
      allMatches = r.matches;

      if (allMatches.length === 1) {
        /* Single match — go straight to upload */
        selectedSamples = [allMatches[0].sampleId];
        currentViewSample = allMatches[0].sampleId;
        setStatus("s-success", "✓ Found: " + allMatches[0].sampleId);
        showUpload();
        loadExistingPhotos(allMatches[0].sampleId);
      } else {
        /* Multiple matches — show picker */
        setStatus("s-success", "Found " + allMatches.length + " samples");
        showMatches();
      }
    })
    .withFailureHandler(function(e) { setStatus("s-error", "Error: " + e.message); })
    .webAppFindSamples(q);
}

/* ============================================================
   MATCH LIST
   ============================================================ */
function showMatches() {
  var el = document.getElementById("matchList");
  document.getElementById("matchSection").style.display = "block";
  document.getElementById("matchCount").textContent = allMatches.length + " found";

  if (allMatches.length > 1) {
    document.getElementById("selectAllRow").style.display = "block";
  }

  var h = "";
  for (var i = 0; i < allMatches.length; i++) {
    var m = allMatches[i];
    var details = [];
    if (m.order) details.push("Order: " + m.order);
    if (m.cargo) details.push("Cargo: " + m.cargo);
    if (m.container) details.push("Ctr: " + m.container);
    if (m.sender) details.push(m.sender);
    if (m.warehouse) details.push(m.warehouse);

    h += '<div class="match-item" onclick="toggleMatch(' + i + ',this)">' +
      '<input type="checkbox" class="match-cb" id="cb' + i + '" data-idx="' + i + '">' +
      '<div class="match-info">' +
        '<div class="match-sample">' + m.sampleId + '</div>' +
        '<div class="match-detail">' + details.join(' · ') + '</div>' +
        (m.description ? '<div class="match-detail">' + m.description + '</div>' : '') +
        (m.hasPhotos ? '<div class="match-photos">📷 Has photos</div>' : '') +
      '</div></div>';
  }
  el.innerHTML = h;
}

function toggleMatch(idx, el) {
  var cb = document.getElementById("cb" + idx);
  cb.checked = !cb.checked;
  el.classList.toggle("selected", cb.checked);
  updateSelectedSamples();
}

function toggleSelectAll() {
  var checked = document.getElementById("selectAll").checked;
  for (var i = 0; i < allMatches.length; i++) {
    var cb = document.getElementById("cb" + i);
    if (cb) {
      cb.checked = checked;
      cb.closest(".match-item").classList.toggle("selected", checked);
    }
  }
  updateSelectedSamples();
}

function updateSelectedSamples() {
  selectedSamples = [];
  for (var i = 0; i < allMatches.length; i++) {
    var cb = document.getElementById("cb" + i);
    if (cb && cb.checked) {
      selectedSamples.push(allMatches[i].sampleId);
    }
  }
  if (selectedSamples.length > 0) {
    showUpload();
    /* Load photos for the first selected */
    currentViewSample = selectedSamples[0];
    loadExistingPhotos(currentViewSample);
  } else {
    document.getElementById("uploadSection").style.display = "none";
    document.getElementById("existingSection").style.display = "none";
  }
}

/* ============================================================
   FILE HANDLING
   ============================================================ */
function showUpload() {
  document.getElementById("uploadSection").style.display = "block";
}

function openCamera() {
  var input = document.getElementById("fileInput");
  input.setAttribute("capture", "environment");
  input.click();
}

function handleFile(file) {
  if (!file) return;
  if (!file.type.match(/^image\\//)) {
    setStatus("s-error", "Not an image file");
    return;
  }

  var reader = new FileReader();
  reader.onload = function(e) {
    var base64 = e.target.result.split(",")[1];
    pendingFile = { base64: base64, name: file.name, mimeType: file.type };

    document.getElementById("previewImg").src = e.target.result;
    document.getElementById("previewName").textContent = file.name +
      (selectedSamples.length > 1 ? " → " + selectedSamples.length + " samples" : " → " + (selectedSamples[0] || ""));
    document.getElementById("previewSection").style.display = "block";
    document.getElementById("uploadSection").style.display = "none";

    setStatus("s-ready", "Preview — tap Upload to save");
  };
  reader.readAsDataURL(file);
}

function cancelPreview() {
  pendingFile = null;
  document.getElementById("previewSection").style.display = "none";
  document.getElementById("uploadSection").style.display = "block";
  document.getElementById("fileInput").value = "";
  setStatus("s-ready", "Upload cancelled");
}

/* Drag and drop */
(function() {
  var dz = document.getElementById("dropZone");
  dz.addEventListener("dragover", function(e) { e.preventDefault(); dz.classList.add("dragover"); });
  dz.addEventListener("dragleave", function() { dz.classList.remove("dragover"); });
  dz.addEventListener("drop", function(e) {
    e.preventDefault(); dz.classList.remove("dragover");
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
})();

/* ============================================================
   UPLOAD
   ============================================================ */
function doUpload() {
  if (!pendingFile || selectedSamples.length === 0) {
    setStatus("s-error", "Select sample(s) and a photo first");
    return;
  }

  var btn = document.getElementById("uploadBtn");
  btn.disabled = true;
  btn.textContent = "Uploading...";
  setStatus("s-loading", "Saving to " + selectedSamples.length + " sample(s)...");
  showProgress(true);

  google.script.run
    .withSuccessHandler(function(r) {
      btn.disabled = false;
      btn.textContent = "📤 Upload";
      showProgress(false);

      if (r.success) {
        setStatus("s-success", r.message);
        pendingFile = null;
        document.getElementById("previewSection").style.display = "none";
        document.getElementById("uploadSection").style.display = "block";
        document.getElementById("fileInput").value = "";
        /* Refresh photos */
        if (currentViewSample) loadExistingPhotos(currentViewSample);
      } else {
        setStatus("s-error", r.message);
      }
    })
    .withFailureHandler(function(e) {
      btn.disabled = false;
      btn.textContent = "📤 Upload";
      showProgress(false);
      setStatus("s-error", "Error: " + e.message);
    })
    .webAppUploadPhoto(selectedSamples, pendingFile.base64, pendingFile.name, pendingFile.mimeType);
}

/* ============================================================
   EXISTING PHOTOS
   ============================================================ */
function loadExistingPhotos(sampleId) {
  google.script.run
    .withSuccessHandler(function(r) {
      var section = document.getElementById("existingSection");
      if (!r.success || !r.photos || r.photos.length === 0) {
        section.style.display = "none";
        return;
      }

      document.getElementById("existingTitle").textContent = "📷 " + sampleId + " (" + r.photos.length + " photo" + (r.photos.length !== 1 ? "s" : "") + ")";

      if (r.folderUrl) {
        var link = document.getElementById("folderLink");
        link.href = r.folderUrl;
        link.style.display = "inline";
      }

      var h = "";
      for (var i = 0; i < r.photos.length; i++) {
        var p = r.photos[i];
        h += '<div class="photo-thumb">' +
          '<img src="' + p.thumbnail + '" onclick="window.open(\\'' + p.url + '\\',\\'_blank\\')" title="' + p.name + '">' +
          '<button class="delete-btn" onclick="deletePhoto(\\'' + p.id + '\\',\\'' + sampleId + '\\')" title="Delete">✕</button>' +
          (p.date ? '<div class="photo-date">' + p.date + '</div>' : '') +
          '</div>';
      }
      document.getElementById("photoGrid").innerHTML = h;
      section.style.display = "block";
    })
    .withFailureHandler(function() {
      document.getElementById("existingSection").style.display = "none";
    })
    .webAppLoadSamplePhotos(sampleId);
}

function deletePhoto(fileId, sampleId) {
  if (!confirm("Delete this photo?")) return;
  setStatus("s-loading", "Deleting...");
  google.script.run
    .withSuccessHandler(function(r) {
      setStatus(r.success ? "s-success" : "s-error", r.message);
      if (r.success) loadExistingPhotos(sampleId);
    })
    .withFailureHandler(function(e) { setStatus("s-error", "Error: " + e.message); })
    .webAppDeletePhoto(fileId);
}

/* ============================================================
   HELPERS
   ============================================================ */
function setStatus(cls, msg) {
  var el = document.getElementById("status");
  el.className = cls;
  el.textContent = msg;
}

function hideAll() {
  document.getElementById("matchSection").style.display = "none";
  document.getElementById("uploadSection").style.display = "none";
  document.getElementById("previewSection").style.display = "none";
  document.getElementById("existingSection").style.display = "none";
  document.getElementById("selectAllRow").style.display = "none";
  selectedSamples = [];
  allMatches = [];
  pendingFile = null;
  currentViewSample = null;
}

function showProgress(show) {
  document.getElementById("progressBar").style.display = show ? "block" : "none";
  if (show) document.getElementById("progressFill").style.width = "80%";
}

/* Focus search on load */
document.getElementById("queryBox").focus();
</script>
</body>
</html>`;
}
