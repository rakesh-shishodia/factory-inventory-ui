/***** CONFIG (pre-filled) *****/
const CONFIG = {
  // File IDs (from your message)
  SOURCE_SPREADSHEET_ID: '1DQ9s7xhHJH8OSQ14EbyGRmtqpr5zC_8997QeEeTrHKQ', // Stock-v4.0.1 - 250825
  TARGET_SPREADSHEET_ID: '1EUcYQ2RvljVTcLdbBS8AbCij0-F_QpaLhZULdr42CrM',  // Ecwid Inventory Mgt Sheet

  // Tabs and header rows
  SOURCE_TAB_NAME: 'Stock Sheet',
  TARGET_TAB_NAME: 'Products',
  SOURCE_HEADER_ROW: 3,   // headers at row 3 in source
  TARGET_HEADER_ROW: 1,   // headers at row 1 in target

  // Key header
  KEY_HEADER: 'sku',      // case-insensitive; we trim & uppercase for matching

  // Source → Target mapping (case-insensitive header matching)
  FIELD_MAP: {
    'unit': 'unit',
    'stock location': 'location',
    'min lvl': 'min_stock',
    'balance': 'current_stock',
  },

  // Helper sheet names
  PREVIEW_SHEET: 'SyncPreview',      // in TARGET file
  LOG_SHEET: 'SyncLog',              // in TARGET file
  ONLY_IN_SOURCE_SHEET: 'OnlyInSource', // in SOURCE file
};

/***** MENU *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Sync')
    .addItem('Preview Sync', 'previewSync')
    .addItem('Apply Updates', 'applyUpdates')
    .addItem('Run Diagnostics', 'diagnoseAccess')
    .addToUi();
}

/***** ENTRY POINTS *****/
function previewSync() {
  const srcSS = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
  const tgtSS = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  const src = srcSS.getSheetByName(CONFIG.SOURCE_TAB_NAME);
  const tgt = tgtSS.getSheetByName(CONFIG.TARGET_TAB_NAME);

  if (!src) throw new Error(`Source tab "${CONFIG.SOURCE_TAB_NAME}" not found.`);
  if (!tgt) throw new Error(`Target tab "${CONFIG.TARGET_TAB_NAME}" not found.`);

  const source = readSheetByHeader(src, CONFIG.SOURCE_HEADER_ROW);
  const target = readSheetByHeader(tgt, CONFIG.TARGET_HEADER_ROW);

  const sCols = headerMap(source.headers);
  const tCols = headerMap(target.headers);

  // Validate required headers exist on both sides
  assertHeadersExist(sCols, [CONFIG.KEY_HEADER, ...Object.keys(CONFIG.FIELD_MAP)]);
  assertHeadersExist(tCols, [CONFIG.KEY_HEADER, ...Object.values(CONFIG.FIELD_MAP)]);

  // Build SKU -> last row (source)
  const srcMap = buildSourceMap(source.rows, sCols[CONFIG.KEY_HEADER]);

  // Build SKU -> array of target row indexes
  const { map: tgtMap } = buildTargetIndex(target.rows, tCols[CONFIG.KEY_HEADER]);

  // SKUs present in source but not in target → OnlyInSource (in SOURCE file)
  const onlyInSource = [];
  Object.keys(srcMap).forEach(sku => {
    if (!tgtMap[sku]) onlyInSource.push(srcMap[sku]);
  });
  writeOnlyInSourceSheet(onlyInSource, srcSS, source.headers);

  // Prepare preview rows for target
  const previewRows = [];
  let countNoChange = 0;
  let countOK = 0;
  let countMissingInSource = 0;

  for (let i = 0; i < target.rows.length; i++) {
    const row = target.rows[i];
    const sku = normalizeSku(row[tCols[CONFIG.KEY_HEADER]]);
    if (!sku) continue; // skip blank SKU rows

    const srcRow = srcMap[sku];
    if (!srcRow) {
      previewRows.push(buildPreviewLine(target, i, null, tCols, 'SKU not found in source'));
      countMissingInSource++;
      continue;
    }

    // Compute proposed update according to rules
    const proposed = computeProposedUpdate(row, srcRow, tCols, sCols);

    const diffs = compareFields(row, proposed, tCols, Object.values(CONFIG.FIELD_MAP));
    const status = diffs.hasChange ? 'OK' : 'No change';
    if (status === 'OK') countOK++; else countNoChange++;

    previewRows.push(buildPreviewLine(target, i, proposed, tCols, status));
  }

  writePreviewSheet(tgtSS, previewRows);

  console.log(JSON.stringify({
    previewSummary: {
      previewRows: previewRows.length,
      ok: countOK,
      noChange: countNoChange,
      missingInSource: countMissingInSource,
      onlyInSource: onlyInSource.length
    }
  }, null, 2));

  SpreadsheetApp.getActive().toast('Preview ready in "SyncPreview".', 'Inventory Sync', 5);
}

function applyUpdates() {
  const tgtSS = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  const tgt = tgtSS.getSheetByName(CONFIG.TARGET_TAB_NAME);
  const preview = tgtSS.getSheetByName(CONFIG.PREVIEW_SHEET);

  if (!tgt) throw new Error(`Target tab "${CONFIG.TARGET_TAB_NAME}" not found.`);
  if (!preview) throw new Error(`"${CONFIG.PREVIEW_SHEET}" not found. Run Preview first.`);

  // Backup target tab
  createBackup(tgtSS, tgt);

  // Read target again to get fresh header map
  const target = readSheetByHeader(tgt, CONFIG.TARGET_HEADER_ROW);
  const tCols = headerMap(target.headers);

  const pData = preview.getDataRange().getValues();
  if (pData.length < 2) {
    SpreadsheetApp.getActive().toast('No rows in SyncPreview to apply.', 'Inventory Sync', 5);
    return;
  }

  // Preview headers
  const pHeaders = pData[0].map(h => String(h).trim().toLowerCase());
  const idxStatus = pHeaders.indexOf('status');
  const idxSheetRow = pHeaders.indexOf('sheet_row');
  const newColIndexByTargetField = {};
  Object.values(CONFIG.FIELD_MAP).forEach(tf => {
    const colName = `${tf}_new`;
    const idx = pHeaders.indexOf(colName);
    if (idx === -1) throw new Error(`SyncPreview missing column: ${colName}`);
    newColIndexByTargetField[tf] = idx;
  });

  // Collect writes: sheetRow -> { targetColIndex: value }
  const writesByRow = new Map();
  let updates = 0;

  for (let r = 1; r < pData.length; r++) {
    const status = String(pData[r][idxStatus] || '').trim();
    if (status !== 'OK') continue;

    const sheetRow = Number(pData[r][idxSheetRow]);
    if (!sheetRow || isNaN(sheetRow)) continue;

    const perRow = {};
    Object.entries(newColIndexByTargetField).forEach(([tf, idx]) => {
      const val = pData[r][idx];
      if (val !== undefined) perRow[tCols[tf]] = val; // 0-based index
    });

    if (Object.keys(perRow).length) {
      writesByRow.set(sheetRow, perRow);
      updates++;
    }
  }

  if (updates === 0) {
    SpreadsheetApp.getActive().toast('No Status=OK rows to update.', 'Inventory Sync', 5);
    return;
  }

  batchWriteByRow(tgt, writesByRow);
  appendLog(tgtSS, { when: new Date(), rowsUpdated: updates, note: 'Applied from SyncPreview' });

  SpreadsheetApp.getActive().toast(`Applied ${updates} updates. See SyncLog.`, 'Inventory Sync', 6);
}

/***** HELPERS *****/
function readSheetByHeader(sheet, headerRow) {
  const rng = sheet.getDataRange().getValues();
  if (rng.length < headerRow) throw new Error(`Header row ${headerRow} beyond data.`);
  const headers = (rng[headerRow - 1] || []).map(s => String(s || '').trim());
  const rows = [];
  for (let r = headerRow; r < rng.length; r++) rows.push(rng[r]);
  return { headers, rows };
}

function headerMap(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).trim().toLowerCase()] = i);
  return map;
}

function assertHeadersExist(map, requiredList) {
  const missing = requiredList.filter(h => map[String(h).toLowerCase()] === undefined);
  if (missing.length) {
    throw new Error('Missing required headers: ' + missing.join(', '));
  }
}

function normalizeSku(val) {
  const s = safeString(val);
  return s ? s.trim().toUpperCase() : '';
}

function safeString(v) {
  if (v === null || v === undefined) return '';
  return String(v);
}

function coerceNumberOrZero(raw) {
  if (raw === '' || raw === null || raw === undefined) return 0;     // blank → 0
  const n = Number(raw);
  return isNaN(n) ? 0 : n;                                          // non-numeric → 0
}

function buildSourceMap(rows, keyIdx) {
  const map = {};
  const dups = new Set();
  for (let i = 0; i < rows.length; i++) {
    const sku = normalizeSku(rows[i][keyIdx]);
    if (!sku) continue;
    if (map[sku]) dups.add(sku);    // last wins
    map[sku] = rows[i];
  }
  if (dups.size) console.log('Duplicate SKUs in source (last occurrence wins):', Array.from(dups).join(', '));
  return map;
}

function buildTargetIndex(rows, keyIdx) {
  const map = {};
  for (let i = 0; i < rows.length; i++) {
    const sku = normalizeSku(rows[i][keyIdx]);
    if (!sku) continue;
    if (!map[sku]) map[sku] = [];
    map[sku].push(i);
  }
  return { map };
}

function computeProposedUpdate(tgtRow, srcRow, tCols, sCols) {
  const out = {};

  // current_stock from Balance (blank/non-numeric => 0)
  if (sCols['balance'] !== undefined) {
    out[tCols['current_stock']] = coerceNumberOrZero(srcRow[sCols['balance']]);
  }

  // min_stock from Min LVL (blank => leave-as-is)
  if (sCols['min lvl'] !== undefined) {
    const raw = srcRow[sCols['min lvl']];
    if (String(raw).trim() !== '') {
      const n = Number(raw);
      out[tCols['min_stock']] = isNaN(n) ? 0 : n;
    }
  }

  // location from Stock Location (blank => leave-as-is)
  if (sCols['stock location'] !== undefined) {
    const loc = srcRow[sCols['stock location']];
    if (String(loc).trim() !== '') out[tCols['location']] = loc;
  }

  // unit from Unit (blank => leave-as-is)
  if (sCols['unit'] !== undefined) {
    const unit = srcRow[sCols['unit']];
    if (String(unit).trim() !== '') out[tCols['unit']] = unit;
  }

  return out;
}

function compareFields(tgtRow, proposed, tCols, targetFields) {
  let hasChange = false;
  targetFields.forEach(tf => {
    const colIdx = tCols[tf];
    if (proposed[colIdx] === undefined) return;     // leave-as-is case
    const oldVal = tgtRow[colIdx];
    const newVal = proposed[colIdx];
    if (!valuesEqual(oldVal, newVal)) hasChange = true;
  });
  return { hasChange };
}

function valuesEqual(a, b) {
  if (a === b) return true;
  // numeric loose compare
  const na = (a === '' || a === null || a === undefined) ? NaN : Number(a);
  const nb = (b === '' || b === null || b === undefined) ? NaN : Number(b);
  if (!isNaN(na) && !isNaN(nb)) return na === nb;
  return String(a) === String(b);
}

function buildPreviewLine(target, targetRowIdx, proposed, tCols, status) {
  const sheetRowNumber = CONFIG.TARGET_HEADER_ROW + 1 + targetRowIdx;

  const getOld = (tf) => target.rows[targetRowIdx][tCols[tf]];
  const getNew = (tf) => (proposed && proposed[tCols[tf]] !== undefined) ? proposed[tCols[tf]] : getOld(tf);

  const sku = normalizeSku(target.rows[targetRowIdx][tCols[CONFIG.KEY_HEADER]]);

  return [
    sku,
    sheetRowNumber,
    getOld('current_stock'), getNew('current_stock'),
    getOld('min_stock'),     getNew('min_stock'),
    getOld('location'),      getNew('location'),
    getOld('unit'),          getNew('unit'),
    status,
  ];
}

function writePreviewSheet(tgtSS, previewRows) {
  const headers = [
    'sku',
    'sheet_row',
    'current_stock_old', 'current_stock_new',
    'min_stock_old',     'min_stock_new',
    'location_old',      'location_new',
    'unit_old',          'unit_new',
    'Status',
  ];

  let sh = tgtSS.getSheetByName(CONFIG.PREVIEW_SHEET);
  if (!sh) sh = tgtSS.insertSheet(CONFIG.PREVIEW_SHEET);
  sh.clearContents();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (previewRows.length) {
    sh.getRange(2, 1, previewRows.length, headers.length).setValues(previewRows);
  }
  sh.autoResizeColumns(1, headers.length);
}

function writeOnlyInSourceSheet(rowsInSourceNotInTarget, srcSS, sourceHeaders) {
  let sh = srcSS.getSheetByName(CONFIG.ONLY_IN_SOURCE_SHEET);
  if (!sh) sh = srcSS.insertSheet(CONFIG.ONLY_IN_SOURCE_SHEET);
  sh.clearContents();

  const headers = sourceHeaders.slice();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (rowsInSourceNotInTarget.length) {
    sh.getRange(2, 1, rowsInSourceNotInTarget.length, headers.length)
      .setValues(rowsInSourceNotInTarget);
  }
  sh.autoResizeColumns(1, headers.length);
}

function createBackup(tgtSS, tgtSheet) {
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const base = `AppBackup_${dateStr}`;
  let name = base, i = 1;
  while (tgtSS.getSheetByName(name)) name = `${base}_${++i}`;
  tgtSheet.copyTo(tgtSS).setName(name);
}

function batchWriteByRow(sheet, writesByRow) {
  // Simple and reliable per-cell writes (fine for a few thousand rows)
  writesByRow.forEach((perRow, sheetRow) => {
    Object.entries(perRow).forEach(([colIdxStr, value]) => {
      const col = Number(colIdxStr) + 1; // 0-based → 1-based
      sheet.getRange(sheetRow, col).setValue(value);
    });
  });
}

function appendLog(tgtSS, entry) {
  const headers = ['timestamp', 'rowsUpdated', 'note'];
  let sh = tgtSS.getSheetByName(CONFIG.LOG_SHEET);
  if (!sh) {
    sh = tgtSS.insertSheet(CONFIG.LOG_SHEET);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sh.appendRow([entry.when, entry.rowsUpdated, entry.note || '']);
}

/***** DIAGNOSTICS *****/
function diagnoseAccess() {
  const ui = SpreadsheetApp.getUi();
  const msgs = [];

  // Spreadsheet-only check for source spreadsheet
  let srcSS;
  try {
    srcSS = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    msgs.push('Source spreadsheet opened: ' + srcSS.getName());
  } catch (e) {
    ui.alert('ERROR: Spreadsheets service could not open the Source spreadsheet.\n' +
             'ID: ' + CONFIG.SOURCE_SPREADSHEET_ID + '\n' +
             'Tip: Open https://docs.google.com/spreadsheets/d/' + CONFIG.SOURCE_SPREADSHEET_ID + '/edit in your browser and ensure this account has access.\n\n' +
             'Details: ' + e.message);
    return;
  }

  // Spreadsheet-only check for target spreadsheet
  let tgtSS;
  try {
    tgtSS = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
    msgs.push('Target spreadsheet opened: ' + tgtSS.getName());
  } catch (e) {
    ui.alert('ERROR: Spreadsheets service could not open the Target spreadsheet.\n' +
             'ID: ' + CONFIG.TARGET_SPREADSHEET_ID + '\n' +
             'Ensure this script is bound to the target file or that you have access.\n\n' +
             'Details: ' + e.message);
    return;
  }

  // Validate tab existence
  const srcTab = srcSS.getSheetByName(CONFIG.SOURCE_TAB_NAME);
  const tgtTab = tgtSS.getSheetByName(CONFIG.TARGET_TAB_NAME);
  if (!srcTab) {
    ui.alert('ERROR: Source tab not found.\n' +
             'Expected tab name: "' + CONFIG.SOURCE_TAB_NAME + '"');
    return;
  }
  if (!tgtTab) {
    ui.alert('ERROR: Target tab not found.\n' +
             'Expected tab name: "' + CONFIG.TARGET_TAB_NAME + '"');
    return;
  }

  // Peek headers
  const srcHeaders = (srcTab.getRange(CONFIG.SOURCE_HEADER_ROW, 1, 1, srcTab.getLastColumn()).getValues()[0] || [])
    .map(h => String(h || '').trim());
  const tgtHeaders = (tgtTab.getRange(CONFIG.TARGET_HEADER_ROW, 1, 1, tgtTab.getLastColumn()).getValues()[0] || [])
    .map(h => String(h || '').trim());

  msgs.push('Source headers @ row ' + CONFIG.SOURCE_HEADER_ROW + ': ' + JSON.stringify(srcHeaders));
  msgs.push('Target headers @ row ' + CONFIG.TARGET_HEADER_ROW + ': ' + JSON.stringify(tgtHeaders));

  // Check required headers
  const srcNeeded = [CONFIG.KEY_HEADER].concat(Object.keys(CONFIG.FIELD_MAP));
  const tgtNeeded = [CONFIG.KEY_HEADER].concat(Object.values(CONFIG.FIELD_MAP));

  const missingSrc = srcNeeded.filter(h => srcHeaders.map(x => x.toLowerCase()).indexOf(h.toLowerCase()) === -1);
  const missingTgt = tgtNeeded.filter(h => tgtHeaders.map(x => x.toLowerCase()).indexOf(h.toLowerCase()) === -1);

  if (missingSrc.length) {
    ui.alert('ERROR: Missing required headers in SOURCE tab @ row ' + CONFIG.SOURCE_HEADER_ROW + ':\n' +
             missingSrc.join(', ') + '\n\nFound headers:\n' + JSON.stringify(srcHeaders, null, 2));
    return;
  }
  if (missingTgt.length) {
    ui.alert('ERROR: Missing required headers in TARGET tab @ row ' + CONFIG.TARGET_HEADER_ROW + ':\n' +
             missingTgt.join(', ') + '\n\nFound headers:\n' + JSON.stringify(tgtHeaders, null, 2));
    return;
  }

  // Everything looks reachable
  ui.alert('Diagnostics OK ✅\n\n' + msgs.join('\n'));
}