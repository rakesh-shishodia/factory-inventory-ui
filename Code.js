/**** CONFIG ****/
const SHEET_PRODUCTS = 'Products';
const SHEET_TX       = 'Transactions';
const SHEET_SYNC     = 'SyncQueue';

/**** WEB ENTRYPOINTS ****/
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;
  if (action === 'getProduct') return getProductHandler(e); // keep HTTP GET for testing

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
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PRODUCTS);
  if (!sheet) throw new Error(`Missing sheet "${SHEET_PRODUCTS}"`);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const skuIndex = headers.indexOf('sku');
  if (skuIndex < 0) throw new Error('Header "sku" not found in Products');
  for (let i = 1; i < data.length; i++) {
    if (data[i][skuIndex] === sku) {
      const row = {};
      headers.forEach((h, j) => row[h] = data[i][j]);
      return row;
    }
  }
  throw new Error('SKU not found');
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
  syncSheet.appendRow([ts, sku, newStock, 'queued', '']);

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
    .addItem('1) Backfill IDs (by SKU)', 'ecwid_backfillIdsBySku')
    .addItem('2) Push queued to Ecwid',  'ecwid_pushQueued')
    .addSeparator()
    .addItem('Test Auth',                'ecwid_testAuth')
    .addToUi();
}

/**** PUBLIC ACTIONS ****/

// 1) Resolve product_id / combination_id for all SKUs in Products
function ecwid_backfillIdsBySku() {
  const ss     = SpreadsheetApp.getActive();
  const sheet  = ss.getSheetByName(SHEET_PRODUCTS);
  const rows   = sheet.getDataRange().getValues();
  const header = rows[0];
  const idx = indexer_(header, [COL_SKU, COL_PROD_ID, COL_COMB_ID]);

  // Build a map of SKU -> row index
  const updates = [];
  for (let r = 1; r < rows.length; r++) {
    const sku = String(rows[r][idx[COL_SKU]] || '').trim();
    if (!sku) continue;

    const hasProdId = rows[r][idx[COL_PROD_ID]];
    // Always (re)lookup if missing product_id
    if (!hasProdId) {
      const info = ecwid_lookupBySku_(sku);
      if (info) {
        updates.push({ r, productId: info.productId, combinationId: info.combinationId || '' });
      }
    }
  }

  // Write back
  updates.forEach(u => {
    if (idx[COL_PROD_ID] >= 0) sheet.getRange(u.r + 1, idx[COL_PROD_ID] + 1).setValue(u.productId);
    if (idx[COL_COMB_ID] >= 0) sheet.getRange(u.r + 1, idx[COL_COMB_ID] + 1).setValue(u.combinationId);
  });

  SpreadsheetApp.getUi().alert(`Backfill complete.\nResolved ${updates.length} SKU(s).`);
}

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
  const sidx       = indexer_(shdr, ['ts','sku','target_stock','status','last_error']);

  let ok = 0, skipped = 0, errors = 0;

  for (let r = 1; r < syncRows.length; r++) {
    const status = String(syncRows[r][sidx['status']] || '').trim();
    if (status !== 'queued') continue;

    const sku         = String(syncRows[r][sidx['sku']] || '').trim();
    const targetStock = Number(syncRows[r][sidx['target_stock']]);

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

    // 3) Compute delta (Ecwid inventory endpoint expects quantityDelta)
    const delta = Number(targetStock) - Number(cur);
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

function ecwid_testAuth() {
  const { storeId, token } = ecwid_creds_();
  const url = `https://app.ecwid.com/api/v3/${storeId}/profile`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  SpreadsheetApp.getUi().alert(`Auth test HTTP ${code}`);
}

function ecwid_lookupBySku_(sku) {
  const { storeId, token } = ecwid_creds_();
  const base = `https://app.ecwid.com/api/v3/${storeId}/products`;
  const url = `${base}?sku=${encodeURIComponent(sku)}&showVariants=true`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) return null;

  const json = JSON.parse(resp.getContentText() || '{}');
  const items = json.items || [];

  // A) product-level SKU
  for (const p of items) {
    if (String(p.sku || '') === sku) {
      return { productId: p.id, combinationId: null };
    }
  }
  // B) variant-level SKU
  for (const p of items) {
    const combs = p.combinations || [];
    for (const c of combs) {
      if (String(c.sku || '') === sku) {
        return { productId: p.id, combinationId: c.id };
      }
    }
  }
  return null;
}

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