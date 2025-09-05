/**** CONFIG ****/
const SHEET_PRODUCTS = 'Products';
const SHEET_TX       = 'Transactions';
const SHEET_SYNC     = 'SyncQueue';
const SHEET_EMPLOYEES = 'Employees';

// Cache settings for getProduct (per-SKU). TTL 10 minutes.
const CACHE_TTL_SEC    = 600;
const CACHE_KEY_PREFIX = 'prod:'; // key = 'prod:' + sku
const CACHE_EMP_KEY    = 'employees:list';
const CACHE_EMP_TTLSEC = 3600; // 1 hour

/**** WEB ENTRYPOINTS ****/
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;
  if (action === 'getProduct') return getProductHandler(e); // keep HTTP GET for testing
  if (action === 'getEmployees') return getEmployeesHandler(e);

  // Serve templated HTML so <?!= ... ?> is processed
  const t = HtmlService.createTemplateFromFile('mobile');
  return t.evaluate()
    .setTitle('Factory Inventory — Mobile')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // optional, OK to keep/remove
}
function doPost(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : null;
    if (action !== 'createTx') {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action (POST)' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Accept JSON or form-encoded (or text/plain)
    let data = {};
    if (e && e.postData) {
      const ct   = String(e.postData.type || '').toLowerCase();
      const body = e.postData.contents || '';

      if (ct.includes('application/json')) {
        // Frontend sent JSON
        data = JSON.parse(body || '{}');
      } else if (ct.includes('application/x-www-form-urlencoded') || ct.includes('text/plain')) {
        // Frontend sent form data (no preflight) or simple text
        data = {
          user:       e.parameter.user,
          sku:        e.parameter.sku,
          location:   e.parameter.location,
          qty_change: e.parameter.qty_change,
          reason:     e.parameter.reason,
          note:       e.parameter.note
        };
      } else {
        // Fallback: try JSON, else use parameters
        try {
          data = JSON.parse(body || '{}');
        } catch (_) {
          data = {
            user:       e.parameter.user,
            sku:        e.parameter.sku,
            location:   e.parameter.location,
            qty_change: e.parameter.qty_change,
            reason:     e.parameter.reason,
            note:       e.parameter.note
          };
        }
      }
    }

    const resp = coreCreateTx_(data);
    return ContentService.createTextOutput(JSON.stringify(resp))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      error: 'Server error in doPost',
      details: String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
/**** CORE LOGIC (shared by HTTP + RPC) ****/
function coreGetProductBySku_(sku) {
  if (!sku) throw new Error('Missing SKU');

  // Normalize and try cache first
  const keySku = String(sku).trim();
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_PREFIX + keySku);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (_) {
      // fall through to sheet lookup on parse error
    }
  }

  // Fallback: read from sheet
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PRODUCTS);
  if (!sheet) throw new Error(`Missing sheet "${SHEET_PRODUCTS}"`);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const skuIndex = headers.indexOf('sku');
  if (skuIndex < 0) throw new Error('Header "sku" not found in Products');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][skuIndex]).trim() === keySku) {
      const row = {};
      headers.forEach((h, j) => row[h] = data[i][j]);
      try {
        cache.put(CACHE_KEY_PREFIX + keySku, JSON.stringify(row), CACHE_TTL_SEC);
      } catch (_) { /* ignore cache quota errors */ }
      return row;
    }
  }
  throw new Error('SKU not found');
}

function coreGetEmployees_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_EMP_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) { /* fall through */ }
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
  if (!sh) throw new Error(`Missing sheet "${SHEET_EMPLOYEES}"`);
  const values = sh.getDataRange().getValues();
  if (!values.length) return [];
  const hdr = values[0].map(v => String(v||'').toLowerCase());
  const nameIdx = hdr.indexOf('name');
  if (nameIdx < 0) throw new Error('Employees sheet must have a "name" header');
  const names = [];
  for (let r = 1; r < values.length; r++) {
    const n = String(values[r][nameIdx] || '').trim();
    if (n) names.push(n);
  }
  try { cache.put(CACHE_EMP_KEY, JSON.stringify(names), CACHE_EMP_TTLSEC); } catch (_) {}
  return names;
}

function coreCreateTx_(data) {
  const ss        = SpreadsheetApp.getActive();
  const txSheet   = ss.getSheetByName(SHEET_TX);
  const prodSheet = ss.getSheetByName(SHEET_PRODUCTS);
  const syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!txSheet || !prodSheet || !syncSheet) {
    throw new Error(`Missing sheet(s) — Products:${!!prodSheet} Transactions:${!!txSheet} SyncQueue:${!!syncSheet}`);
  }

  const ts = new Date();
  const { user, sku, location, qty_change, reason, note } = data || {};
  if (!sku) throw new Error('Missing SKU in payload');

  // 1) Append transaction
  txSheet.appendRow([
    ts,
    user || '',
    sku,
    location || 'MAIN',
    Number(qty_change) || 0,
    reason || '',
    note || '',
    'mobile'
  ]);

  // 2) Update product stock
  const prodData   = prodSheet.getDataRange().getValues();
  const headers    = prodData[0] || [];
  const skuIndex   = headers.indexOf('sku');
  const stockIndex = headers.indexOf('current_stock');
  if (skuIndex < 0 || stockIndex < 0) {
    throw new Error('Products needs headers "sku" and "current_stock"');
  }

  let newStock = null;
  for (let i = 1; i < prodData.length; i++) {
    if (prodData[i][skuIndex] === sku) {
      newStock = (Number(prodData[i][stockIndex]) || 0) + (Number(qty_change) || 0);
      prodSheet.getRange(i + 1, stockIndex + 1).setValue(newStock);
      break;
    }
  }
  if (newStock === null) throw new Error('SKU not found in Products');

  // 3) Queue Ecwid sync
  const allowRaise = (Number(qty_change) || 0) > 0; // positive transactions allow raises
  syncSheet.appendRow([ts, sku, newStock, allowRaise, 'queued', '']);

  return { ok: true, sku, new_stock: newStock, warning: newStock < 0 ? 'Stock is negative!' : null };
}

/**** RPC ENTRYPOINTS (for google.script.run from the HTML) ****/
function api_getProduct(sku)   { return coreGetProductBySku_(sku); }
function api_createTx(payload) { return coreCreateTx_(payload); }

/**** OPTIONAL: HTTP HANDLERS (for direct URL tests) ****/
function getProductHandler(e) {
  try {
    const row = coreGetProductBySku_(e.parameter && e.parameter.sku);
    return ContentService.createTextOutput(JSON.stringify(row)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getEmployeesHandler(e) {
  try {
    const names = coreGetEmployees_();
    return ContentService.createTextOutput(JSON.stringify({ names }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createTxHandler(data) {
  try {
    const resp = coreCreateTx_(data);
    return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Column names in Products (exact, case-insensitive ok)
const COL_SKU           = 'sku';
const COL_PROD_ID       = 'product_id';
const COL_COMB_ID       = 'combination_id';

// Script Properties keys (set once under: Extensions → Apps Script → Project Settings → Script properties)
const SP_ECWID_STORE_ID = 'ECWID_STORE_ID';
const SP_ECWID_TOKEN    = 'ECWID_TOKEN';
/**** MENU ****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ecwid Sync')
    .addItem('Push queued to Ecwid',  'ecwid_pushQueued')
    .addSeparator()
    .addItem('Refresh Catalog from Ecwid', 'ecwid_refreshCatalog')
    .addItem('Import product by ID…',    'ecwid_importProductById')
    .addToUi();
}

/**** PUBLIC ACTIONS ****/


// 2) Push queued SyncQueue rows (status="queued") to Ecwid
function ecwid_pushQueued() {
  const ss         = SpreadsheetApp.getActive();
  const prodSheet  = ss.getSheetByName(SHEET_PRODUCTS);
  const syncSheet  = ss.getSheetByName(SHEET_SYNC);

  if (!prodSheet || !syncSheet) throw new Error('Missing Products or SyncQueue sheet.');

  const prodRows   = prodSheet.getDataRange().getValues();
  const prodHdr    = prodRows[0];
  const pidx       = indexer_(prodHdr, [COL_SKU, COL_PROD_ID, COL_COMB_ID]);

  const syncRows   = syncSheet.getDataRange().getValues();
  const shdr       = syncRows[0];
  const sidx       = indexer_(shdr, ['ts','sku','target_stock','allow_raise','status','last_error']);

  let ok = 0, skipped = 0, errors = 0;

  for (let r = 1; r < syncRows.length; r++) {
    const status = String(syncRows[r][sidx['status']] || '').trim();
    if (status !== 'queued') continue;

    const sku         = String(syncRows[r][sidx['sku']] || '').trim();
    const targetStock = Number(syncRows[r][sidx['target_stock']]);
    const allowRaw    = sidx['allow_raise'] >= 0 ? syncRows[r][sidx['allow_raise']] : false;
    const allowStr    = String(allowRaw).toLowerCase();
    const allowRaise  = (allowRaw === true) || allowStr === 'true' || allowStr === '1' || allowStr === 'yes' || allowStr === 'y';

    // 1) Find product row by SKU
    const prow = findRowBySku_(prodRows, pidx[COL_SKU], sku);
    if (prow < 0) {
      setSyncRow_(syncSheet, r, sidx, 'error', `SKU ${sku} not found in Products`);
      errors++; continue;
    }

    const productId     = prodRows[prow][pidx[COL_PROD_ID]];
    const combinationId = prodRows[prow][pidx[COL_COMB_ID]];

    if (!productId) {
      setSyncRow_(syncSheet, r, sidx, 'error', `Missing product_id for SKU ${sku}`);
      errors++; continue;
    }

    // 2) Read current quantity from Ecwid (product or combination)
    const cur = ecwid_getCurrentQty_(productId, combinationId);
    if (cur == null) {
      setSyncRow_(syncSheet, r, sidx, 'error', `Unable to read current qty for ${sku}`);
      errors++; continue;
    }

    // 3) Compute target per policy
    let target;
    if (allowRaise) {
      // additions: allow raise or lower to match sheet (clamped at 0)
      target = Math.max(0, Number(targetStock));
    } else {
      // default min-rule: never raise Ecwid, only lower (clamped at 0)
      target = Math.max(0, Math.min(Number(targetStock), Number(cur)));
    }

    // Ecwid inventory endpoint expects quantityDelta
    const delta = target - Number(cur);
    if (!Number.isFinite(delta)) {
      setSyncRow_(syncSheet, r, sidx, 'error', `Invalid target/current for ${sku}: ${targetStock}/${cur}`);
      errors++; continue;
    }
    if (delta === 0) {
      setSyncRow_(syncSheet, r, sidx, 'done', '');
      skipped++; continue;
    }

    // 4) Push delta
    const pushed = ecwid_adjustQtyDelta_(productId, combinationId, delta);
    if (pushed) {
      setSyncRow_(syncSheet, r, sidx, 'done', '');
      ok++;
    } else {
      setSyncRow_(syncSheet, r, sidx, 'error', `Failed to push delta ${delta} for ${sku}`);
      errors++;
    }

    Utilities.sleep(250); // gentle rate limiting
  }

  SpreadsheetApp.getUi().alert(`Sync complete.\nOK: ${ok}\nSkipped (already at target): ${skipped}\nErrors: ${errors}`);
}

/**** HELPERS ****/

function indexer_(header, keys) {
  const map = {};
  const lower = header.map(h => String(h || '').toLowerCase());
  keys.forEach(k => { map[k] = lower.indexOf(String(k).toLowerCase()); });
  return map;
}

function findRowBySku_(rows, skuIdx, sku) {
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][skuIdx]).trim() === sku) return i;
  }
  return -1;
}

function setSyncRow_(sheet, r, sidx, status, errMsg) {
  if (sidx['status'] >= 0)    sheet.getRange(r + 1, sidx['status'] + 1).setValue(status);
  if (sidx['last_error'] >= 0) sheet.getRange(r + 1, sidx['last_error'] + 1).setValue(errMsg || '');
}

/**** ECWID API LAYER ****/



function ecwid_getCurrentQty_(productId, combinationId) {
  const { storeId, token } = ecwid_creds_();
  let url;
  if (combinationId) {
    url = `https://app.ecwid.com/api/v3/${storeId}/products/${productId}/combinations/${combinationId}`;
  } else {
    url = `https://app.ecwid.com/api/v3/${storeId}/products/${productId}`;
  }
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) return null;
  const data = JSON.parse(resp.getContentText() || '{}');
  if (data.unlimited === true) return 0; // cannot adjust; treat as 0 to compute delta; you may want to skip
  return Number(data.quantity);
}

function ecwid_adjustQtyDelta_(productId, combinationId, delta) {
  const { storeId, token } = ecwid_creds_();
  let url;
  if (combinationId) {
    url = `https://app.ecwid.com/api/v3/${storeId}/products/${productId}/combinations/${combinationId}/inventory`;
  } else {
    url = `https://app.ecwid.com/api/v3/${storeId}/products/${productId}/inventory`;
  }
  const resp = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ quantityDelta: Number(delta) }),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  if (code === 200) return true;

  // capture error text in logs
  const body = resp.getContentText();
  console.warn('Ecwid adjust error', code, body);
  return false;
}

function ecwid_creds_() {
  const props   = PropertiesService.getScriptProperties();
  const storeId = props.getProperty(SP_ECWID_STORE_ID);
  const token   = props.getProperty(SP_ECWID_TOKEN);
  if (!storeId || !token) throw new Error('Missing ECWID_STORE_ID or ECWID_TOKEN in Script Properties.');
  return { storeId, token };
}
function ecwid_refreshCatalog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const prodSheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (!prodSheet) { ui.alert('Products sheet not found'); return; }

  // 1) Read existing Products to preserve user columns by SKU
  const existing = prodSheet.getDataRange().getValues();
  const existingHdr = existing[0] || [];
  const existingIdx = indexer_(existingHdr, existingHdr); // identity map name->index
  const skuCol = existingIdx['sku'];
  if (skuCol < 0) { ui.alert('Products sheet must include a "sku" header'); return; }

  const keepBySku = {}; // sku -> row object of existing values
  for (let r = 1; r < existing.length; r++) {
    const row = existing[r];
    const sku = String(row[skuCol] || '').trim();
    if (!sku) continue;
    const obj = {};
    for (let c = 0; c < existingHdr.length; c++) obj[String(existingHdr[c]||'').toLowerCase()] = row[c];
    keepBySku[sku] = obj;
  }

  // 2) Fetch full Ecwid catalog (paged)
  const { storeId, token } = ecwid_creds_();
  const pageLimit = 100;
  let offset = 0, total = null;
  const ecwidRows = []; // rows with import fields only

  do {
    const url = `https://app.ecwid.com/api/v3/${storeId}/products?showVariants=true&limit=${pageLimit}&offset=${offset}`;
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) { ui.alert(`Failed to fetch catalog (HTTP ${resp.getResponseCode()})`); return; }
    const data = JSON.parse(resp.getContentText() || '{}');
    total = data.total || 0;
    const items = data.items || [];

    for (const p of items) {
      const pThumb = p.thumbnailUrl || p.imageUrl || '';
      // product-level SKU
      if (p.sku && String(p.sku).trim()) {
        ecwidRows.push({
          product_id: p.id,
          combination_id: '',
          sku: p.sku,
          name: p.name,
          option_values: '',
          enabled: p.enabled,
          unlimited: p.unlimited,
          image_url: pThumb
        });
      }
      // variant-level SKUs
      if (Array.isArray(p.combinations)) {
        for (const c of p.combinations) {
          if (!c || !c.sku || !String(c.sku).trim()) continue;
          const cThumb = c.thumbnailUrl || c.imageUrl || pThumb || '';
          ecwidRows.push({
            product_id: p.id,
            combination_id: c.id,
            sku: c.sku,
            name: p.name,
            option_values: ecwid_formatOptionValues_(c.optionValues),
            enabled: c.enabled,
            unlimited: c.unlimited,
            image_url: cThumb
          });
        }
      }
    }
    offset += items.length;
  } while (offset < total);

  // 3) Build final header: start from existing headers, ensure required import fields exist
  const required = ['product_id','combination_id','sku','name','option_values','enabled','unlimited','image_url'];
  const finalHdr = existingHdr.slice();
  for (const h of required) {
    if (finalHdr.map(x=>String(x||'').toLowerCase()).indexOf(h) === -1) finalHdr.push(h);
  }

  // 4) Compose full table in memory, preserving existing per SKU for non-import columns
  const finalRows = [];
  finalRows.push(finalHdr); // header row
  const importSet = new Set(required);

  for (const r of ecwidRows) {
    const sku = r.sku;
    const keep = keepBySku[sku] || {};
    const rowArr = new Array(finalHdr.length).fill('');
    for (let i = 0; i < finalHdr.length; i++) {
      const colName = String(finalHdr[i]||'').toLowerCase();
      if (importSet.has(colName)) {
        rowArr[i] = r[colName] != null ? r[colName] : '';
      } else {
        rowArr[i] = keep[colName] != null ? keep[colName] : '';
      }
    }
    finalRows.push(rowArr);
  }

  // 5) Write to TMP in bulk, then swap back to Products
  const TMP = 'CatalogImport_TMP';
  let tmp = ss.getSheetByName(TMP);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(TMP);

  const chunk = 1000; // rows per chunk to avoid setValues limits
  for (let start = 0; start < finalRows.length; start += chunk) {
    const block = finalRows.slice(start, Math.min(start + chunk, finalRows.length));
    tmp.getRange(start + 1, 1, block.length, finalHdr.length).setValues(block);
  }

  // Clear Products and bulk write composed table back
  prodSheet.clearContents();
  for (let start = 0; start < finalRows.length; start += chunk) {
    const block = finalRows.slice(start, Math.min(start + chunk, finalRows.length));
    prodSheet.getRange(start + 1, 1, block.length, finalHdr.length).setValues(block);
  }

  // Optional: remove TMP sheet after success
  ss.deleteSheet(tmp);

  ui.alert(`Catalog refresh complete. Rows written: ${finalRows.length - 1}`);
}

function ecwid_importProductById() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Import product by ID', 'Enter Ecwid productId (numeric):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const pid = String(resp.getResponseText()||'').trim();
  if (!pid || !/^\d+$/.test(pid)) { ui.alert('Invalid productId'); return; }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sheet) { ui.alert('Products sheet not found'); return; }

  // Read existing table once
  const existing = sheet.getDataRange().getValues();
  const hdr = existing[0] || [];
  const idx = indexer_(hdr, hdr);
  const skuCol = idx['sku'];
  if (skuCol < 0) { ui.alert('Products sheet must include a "sku" header'); return; }

  // Map existing rows by SKU
  const keepBySku = {};
  for (let r = 1; r < existing.length; r++) {
    const row = existing[r];
    const sku = String(row[skuCol] || '').trim();
    if (!sku) continue;
    const obj = {};
    for (let c = 0; c < hdr.length; c++) obj[String(hdr[c]||'').toLowerCase()] = row[c];
    keepBySku[sku] = obj;
  }

  // Fetch product (with variants)
  const prod = ecwid_fetchProductById_(pid);
  if (!prod) { ui.alert('Product not found or API error'); return; }

  // Build rows for this product
  const p = prod;
  const pThumb = p.thumbnailUrl || p.imageUrl || '';
  const rows = [];
  if (p.sku && String(p.sku).trim()) {
    rows.push({ product_id: p.id, combination_id: '', sku: p.sku, name: p.name, option_values: '', enabled: p.enabled, unlimited: p.unlimited, image_url: pThumb });
  }
  if (Array.isArray(p.combinations)) {
    for (const c of p.combinations) {
      if (!c || !c.sku || !String(c.sku).trim()) continue;
      const cThumb = c.thumbnailUrl || c.imageUrl || pThumb || '';
      rows.push({ product_id: p.id, combination_id: c.id, sku: c.sku, name: p.name, option_values: ecwid_formatOptionValues_(c.optionValues), enabled: c.enabled, unlimited: c.unlimited, image_url: cThumb });
    }
  }

  // Ensure required headers exist
  const required = ['product_id','combination_id','sku','name','option_values','enabled','unlimited','image_url'];
  const finalHdr = hdr.slice();
  for (const h of required) {
    if (finalHdr.map(x=>String(x||'').toLowerCase()).indexOf(h) === -1) finalHdr.push(h);
  }
  if (finalHdr.length !== hdr.length) {
    // expand sheet headers if new columns were added
    sheet.insertColumnsAfter(hdr.length, finalHdr.length - hdr.length);
    sheet.getRange(1, 1, 1, finalHdr.length).setValues([finalHdr]);
  }

  // Upsert rows: update existing SKUs in place, collect new SKUs to append
  const importSet = new Set(required);
  const newRows = [];

  // Build a helper to render a row array for a given SKU record
  function renderRow(obj, keepMap) {
    const sku = obj.sku;
    const keep = keepMap[sku] || {};
    const arr = new Array(finalHdr.length).fill('');
    for (let i = 0; i < finalHdr.length; i++) {
      const col = String(finalHdr[i]||'').toLowerCase();
      if (importSet.has(col)) arr[i] = obj[col] != null ? obj[col] : '';
      else arr[i] = keep[col] != null ? keep[col] : '';
    }
    return arr;
  }

  // Update existing
  const nameToIndex = indexer_(hdr, hdr);
  const skuToRow = {};
  for (let r = 1; r < existing.length; r++) {
    const s = String(existing[r][skuCol]||'').trim();
    if (s) skuToRow[s] = r;
  }
  for (const obj of rows) {
    if (skuToRow[obj.sku] != null) {
      const rowIdx = skuToRow[obj.sku];
      const arr = renderRow(obj, keepBySku);
      sheet.getRange(rowIdx + 1, 1, 1, finalHdr.length).setValues([arr]);
    } else {
      newRows.push(renderRow(obj, keepBySku));
    }
  }

  // Append new rows in bulk
  if (newRows.length) {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, newRows.length, finalHdr.length).setValues(newRows);
  }

  ui.alert(`Imported product ${pid}. Updated: ${rows.length - newRows.length}, Added: ${newRows.length}`);
}

function ecwid_fetchProductById_(productId) {
  const { storeId, token } = ecwid_creds_();
  const url = `https://app.ecwid.com/api/v3/${storeId}/products/${productId}?showVariants=true`;
  const resp = UrlFetchApp.fetch(url, { method: 'get', headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return null;
  return JSON.parse(resp.getContentText() || '{}');
}

// Format optionValues array to a string
function ecwid_formatOptionValues_(optionValues) {
  if (!Array.isArray(optionValues) || !optionValues.length) return '';
  return optionValues.map(function(opt) {
    if (opt.optionName && opt.optionValue) {
      return opt.optionName + ':' + opt.optionValue;
    }
    return '';
  }).filter(Boolean).join('; ');
}