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
  // Optional: keep HTTP POST route for external tools (curl/Postman)
  try {
    const action = e && e.parameter ? e.parameter.action : null;
    if (action === 'createTx') {
      const body = e.postData && e.postData.contents ? e.postData.contents : '{}';
      const data = JSON.parse(body);
      return createTxHandler(data);
    }
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action (POST)' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Server error in doPost', details: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
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