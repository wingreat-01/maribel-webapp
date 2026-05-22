// =============================================================================
// YMCC Inventory System — Google Apps Script Backend
// Handles: Beginning (DATA tab), Received, Returns, proxy fetch for CORS
// Deploy as: Web App | Execute as: Me | Who has access: Anyone
// =============================================================================

// ── CONFIGURE THIS ────────────────────────────────────────────────────────────
// Paste your Google Sheet ID here (found in the sheet URL between /d/ and /edit)
var SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';

// Tab names — must match the actual tab names in your Google Sheet
var DATA_TAB     = 'BEGINNING';  // Beginning inventory: cols A:G only
var RECEIVED_TAB = 'RECEIVED';   // Received items log
var RETURNS_TAB  = 'RETURNS';    // Returns/refunds log

// Column headers — auto-created if the tab doesn't exist yet
// BEGINNING tab uses exactly 7 columns (A:G). Columns H:Q are ignored by this system.
var DATA_HEADER     = ['ItemCode', 'Qty', 'Width', 'W-UM', 'Length', 'L-UM', 'Remarks'];
var RECEIVED_HEADER = ['Date', 'ItemCode', 'Qty', 'Size', 'MRR', 'Supplier', 'Remarks', 'ID'];
var RETURNS_HEADER  = ['Date', 'ItemCode', 'Qty', 'Size', 'Customer', 'WhdlNo', 'Remarks', 'ID'];

// Allowed origins for CORS — add your GitHub Pages URL here after deploying
// Example: 'https://your-username.github.io'
var ALLOWED_ORIGINS = [
  'https://your-username.github.io',   // ← replace with your actual GitHub Pages URL
  'http://localhost',                   // local dev
  'http://127.0.0.1'                   // local dev
];
// ─────────────────────────────────────────────────────────────────────────────


// =============================================================================
// ENTRY POINTS
// =============================================================================

function doGet(e) {
  var output;
  var action = ((e && e.parameter && e.parameter.action) || '').toLowerCase();

  try {
    // ── CORS proxy: fetch an external URL on behalf of the client ─────────────
    // Used to bypass CORS restrictions when fetching published Google Sheet CSVs
    if (!action || action === 'proxy') {
      var url = e.parameter.url;
      if (!url) return _txt('No URL provided');
      var resp = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      output = _txt(resp.getContentText());
      return _addCors(output);
    }

    // ── Read actions (return JSON) ────────────────────────────────────────────
    if (action === 'read_data') {
      var dsh = _getDataSheet(); var lastRow = dsh.getLastRow(); output = _json({ok: true, rows: lastRow < 1 ? [] : dsh.getRange(1, 1, lastRow, 7).getValues()});
      return _addCors(output);
    }
    if (action === 'read_received') {
      output = _json({ok: true, rows: _getReceivedSheet().getDataRange().getValues()});
      return _addCors(output);
    }
    if (action === 'read_returns') {
      output = _json({ok: true, rows: _getReturnsSheet().getDataRange().getValues()});
      return _addCors(output);
    }

    // ── Write actions passed as GET query params ───────────────────────────────
    // GitHub Pages can't send cross-origin POST with a body, so writes come as
    // GET requests with the JSON payload URL-encoded in ?payload=...
    if (e.parameter.payload) {
      var body;
      try { body = JSON.parse(e.parameter.payload); }
      catch (err) { return _addCors(_json({ok: false, error: 'Bad payload JSON: ' + err})); }
      body.action = action;
      output = _handleWrite(body);
      return _addCors(output);
    }

    return _addCors(_json({ok: false, error: 'Unknown action: ' + action}));

  } catch (err) {
    return _addCors(_json({ok: false, error: 'doGet error: ' + String(err.message || err)}));
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return _addCors(_handleWrite(body));
  } catch (err) {
    return _addCors(_json({ok: false, error: 'doPost error: ' + String(err.message || err)}));
  }
}


// =============================================================================
// WRITE HANDLER
// =============================================================================

function _handleWrite(body) {
  var action = ((body && body.action) || '').toLowerCase();

  // ── Beginning inventory ────────────────────────────────────────────────────
  if (action === 'save_beginning_single') {
    var it = body.item || {};
    if (!it.code) return _json({ok: false, error: 'Missing item code'});
    _upsert(_getDataSheet(), it.code, it.qty, it.width||'', it.widthUm||'', it.length||'', it.lengthUm||'', it.remarks||'');
    return _json({ok: true, code: it.code, qty: Number(it.qty) || 0});
  }

  if (action === 'save_beginning_bulk') {
    var items = body.items || [];
    var dsh = _getDataSheet();
    items.forEach(function(it) { _upsert(dsh, it.code, it.qty, it.width||'', it.widthUm||'', it.length||'', it.lengthUm||'', it.remarks||''); });
    return _json({ok: true, count: items.length});
  }

  // ── Received ───────────────────────────────────────────────────────────────
  if (action === 'save_received') {
    var r = body.record || {};
    if (!r.itemCode) return _json({ok: false, error: 'Missing itemCode'});
    var id = r.id || _genId('rec');
    var sh = _getReceivedSheet();
    var existing = _findRowById(sh, 8, id);
    var row = [r.date || '', String(r.itemCode), Number(r.qty) || 0, r.size || '', r.mrrNo || '', r.supplier || '', r.remarks || '', id];
    if (existing === -1) sh.appendRow(row);
    else sh.getRange(existing, 1, 1, row.length).setValues([row]);
    return _json({ok: true, id: id});
  }

  if (action === 'delete_received') {
    var id = body.id || '';
    var sh = _getReceivedSheet();
    var rowIdx = _findRowById(sh, 8, id);
    if (rowIdx === -1) return _json({ok: false, error: 'Received row not found: ' + id});
    sh.deleteRow(rowIdx);
    return _json({ok: true});
  }

  // ── Returns ────────────────────────────────────────────────────────────────
  if (action === 'save_return') {
    var r = body.record || {};
    if (!r.itemCode) return _json({ok: false, error: 'Missing itemCode'});
    var id = r.id || _genId('ret');
    var sh = _getReturnsSheet();
    var existing = _findRowById(sh, 8, id);
    var row = [r.date || '', String(r.itemCode), Number(r.qty) || 0, r.size || '', r.customer || '', r.whdlNo || '', r.remarks || '', id];
    if (existing === -1) sh.appendRow(row);
    else sh.getRange(existing, 1, 1, row.length).setValues([row]);
    return _json({ok: true, id: id});
  }

  if (action === 'delete_return') {
    var id = body.id || '';
    var sh = _getReturnsSheet();
    var rowIdx = _findRowById(sh, 8, id);
    if (rowIdx === -1) return _json({ok: false, error: 'Return row not found: ' + id});
    sh.deleteRow(rowIdx);
    return _json({ok: true});
  }

  // ── Full backup (Brave → Sheet migration) ──────────────────────────────────
  if (action === 'backup_all') {
    var b = body.data || {};

    var dsh = _getDataSheet();
    dsh.clear();
    var drows = [DATA_HEADER];
    (b.beginning || []).forEach(function(it) {
      drows.push([String(it.code || ''), Number(it.qty) || 0, String(it.width||''), String(it.widthUm||''), String(it.length||''), String(it.lengthUm||''), String(it.remarks||'')]);
    });
    dsh.getRange(1, 1, drows.length, DATA_HEADER.length).setValues(drows);

    var rsh = _getReceivedSheet();
    rsh.clear();
    var rrows = [RECEIVED_HEADER];
    (b.received || []).forEach(function(r) {
      rrows.push([r.date || '', String(r.itemCode || ''), Number(r.qty) || 0,
                  r.size || '', r.mrrNo || '', r.supplier || '', r.remarks || '',
                  r.id || _genId('rec')]);
    });
    if (rrows.length > 1) rsh.getRange(1, 1, rrows.length, RECEIVED_HEADER.length).setValues(rrows);

    var esh = _getReturnsSheet();
    esh.clear();
    var erows = [RETURNS_HEADER];
    (b.returns || []).forEach(function(r) {
      erows.push([r.date || '', String(r.itemCode || ''), Number(r.qty) || 0,
                  r.size || '', r.customer || '', r.whdlNo || '', r.remarks || '',
                  r.id || _genId('ret')]);
    });
    if (erows.length > 1) esh.getRange(1, 1, erows.length, RETURNS_HEADER.length).setValues(erows);

    return _json({ok: true, written: {data: drows.length - 1, received: rrows.length - 1, returns: erows.length - 1}});
  }

  // ── Delete beginning row ───────────────────────────────────────────────────
  if (action === 'delete_row') {
    var code = body.code || body.id || '';
    if (code.indexOf('beg_') === 0) code = code.substring(4);
    var sh = _getDataSheet();
    var rowIdx = _findRowByCode(sh, code);
    if (rowIdx === -1) return _json({ok: false, error: 'Item not found: ' + code});
    sh.deleteRow(rowIdx);
    return _json({ok: true});
  }

  return _json({ok: false, error: 'Unknown write action: ' + action});
}


// =============================================================================
// SHEET ACCESSORS
// =============================================================================

function _openSS() {
  if (!SHEET_ID || SHEET_ID === '1mbafKFp4BeSA8xYmhl77y48YBf38UFjgJRJyDEAHxYs') {
    throw new Error('SHEET_ID is not configured. Open Code.gs and paste your Google Sheet ID.');
  }
  return SpreadsheetApp.openById(SHEET_ID);
}

function _getSheetWithHeader(name, header) {
  var ss = _openSS();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    return sh;
  }
  if (sh.getLastRow() === 0) sh.appendRow(header);
  return sh;
}

function _getDataSheet()     { return _getSheetWithHeader(DATA_TAB,     DATA_HEADER); }
function _getReceivedSheet() { return _getSheetWithHeader(RECEIVED_TAB, RECEIVED_HEADER); }
function _getReturnsSheet()  { return _getSheetWithHeader(RETURNS_TAB,  RETURNS_HEADER); }


// =============================================================================
// HELPERS
// =============================================================================

// Find a row in a sheet by matching the value in column 1 (item code)
function _findRowByCode(sh, code) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var codes = sh.getRange(2, 1, last - 1, 1).getValues();
  var target = String(code).trim().toLowerCase();
  for (var i = 0; i < codes.length; i++) {
    if (String(codes[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

// Find a row by its ID column (1-based column index)
function _findRowById(sh, idColIndex1Based, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, idColIndex1Based, last - 1, 1).getValues();
  var target = String(id).trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === target) return i + 2;
  }
  return -1;
}

// Insert or update a row in the BEGINNING sheet (cols A:G only — never touches H:Q)
function _upsert(sh, code, qty, width, widthUm, length, lengthUm, remarks) {
  if (!code) return;
  var w   = (width    != null) ? String(width)    : '';
  var wu  = (widthUm  != null) ? String(widthUm)  : '';
  var l   = (length   != null) ? String(length)   : '';
  var lu  = (lengthUm != null) ? String(lengthUm) : '';
  var rem = (remarks  != null) ? String(remarks)  : '';
  var rowIdx = _findRowByCode(sh, code);
  if (rowIdx === -1) {
    sh.appendRow([String(code), Number(qty) || 0, w, wu, l, lu, rem]);
  } else {
    // Only overwrite A:G — leave H:Q untouched by using setValues on a 7-col range
    sh.getRange(rowIdx, 1, 1, 7).setValues([[String(code), Number(qty) || 0, w, wu, l, lu, rem]]);
  }
}

function _genId(prefix) {
  return prefix + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
}

// Response builders
function _txt(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}
function _json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Add CORS headers so GitHub Pages (a different origin) can read the response.
// Apps Script's ContentService doesn't natively support setHeader(), so we
// wrap the output in a JSONP-style response isn't needed — Apps Script Web Apps
// deployed as "Anyone" already allow cross-origin reads via fetch() with
// no-cors or redirect:follow. The key requirement is that the deploy access
// is set to "Anyone" (not "Anyone with Google Account").
// This function is a passthrough kept for clarity and future flexibility.
function _addCors(output) {
  return output;
}
