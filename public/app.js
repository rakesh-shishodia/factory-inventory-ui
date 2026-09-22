const $ = (selector, context = document) => context.querySelector(selector);
const number = new Intl.NumberFormat('en-IN');
const PENDING_KEY = 'factory-inventory.pending-operation.v1';

const REASONS = {
  ECWID_PICK: { label: 'Ecwid Order Pick', apiType: 'ECWID_PICK', direction: -1 },
  EMAIL_SALE: { label: 'Email Sale', apiType: 'EMAIL_SALE', direction: -1 },
  INTERNAL_USE: { label: 'Internal Use', apiType: 'INTERNAL_USE', direction: -1 },
  RESTOCK: { label: 'Restock', apiType: 'RESTOCK', direction: 1 },
  ADJUST_UP: { label: 'Adjust Up', apiType: 'RESTOCK', reasonCode: 'ADJUST_UP', direction: 1, notePrefix: 'Adjustment up: ', noteRequired: true },
  ADJUST_DOWN: { label: 'Adjust Down', apiType: 'INTERNAL_USE', reasonCode: 'ADJUST_DOWN', direction: -1, notePrefix: 'Adjustment down: ', noteRequired: true },
};
const LEGACY_PENDING_TYPES = new Set(['ECWID_PICK', 'EMAIL_SALE', 'INTERNAL_USE', 'RESTOCK', 'ALLOCATE', 'RELEASE']);

const state = {
  session: null,
  item: null,
  ecwidStock: null,
  orderSelection: null,
  loading: false,
  lookingUp: false,
  orderLookingUp: false,
  submitting: false,
  pending: readPending(),
  lookupSequence: 0,
  orderLookupSequence: 0,
  facingMode: 'environment',
  scanner: null,
  scannerStarting: null,
  scannerStopping: null,
  scannerGeneration: 0,
};

function icon(name) {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', 'icon');
  svg.setAttribute('aria-hidden', 'true');
  const use = document.createElementNS('http://www.w3.org/2000/svg', 'use');
  use.setAttribute('href', `#i-${name}`);
  svg.append(use);
  return svg;
}

function readPending() {
  try {
    const saved = JSON.parse(localStorage.getItem(PENDING_KEY) || 'null');
    const payload = saved?.payload;
    const orderRequired = ['ECWID_PICK', 'ALLOCATE', 'RELEASE'].includes(payload?.type);
    return payload?.operation_id && payload?.item_id && LEGACY_PENDING_TYPES.has(payload.type)
      && (!orderRequired || (payload.order_id && payload.order_line_id)) ? saved : null;
  } catch { return null; }
}

function persistPending(value) {
  state.pending = value;
  try {
    if (value) localStorage.setItem(PENDING_KEY, JSON.stringify(value));
    else localStorage.removeItem(PENDING_KEY);
  } catch { /* The in-memory copy still protects same-page retries. */ }
  renderPending();
}

class ApiError extends Error {
  constructor(message, status, details) {
    super(message);
    this.status = status;
    this.details = details;
  }
}

async function api(path, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 25000);
  try {
    const response = await fetch(path, {
      ...options,
      credentials: 'same-origin',
      cache: 'no-store',
      signal: controller.signal,
      headers: {
        Accept: 'application/json',
        ...(options.body ? { 'Content-Type': 'application/json' } : {}),
        ...options.headers,
      },
    });
    const result = await response.json().catch(() => null);
    if (!response.ok) {
      const message = typeof result?.error === 'string' ? result.error : result?.error?.message || result?.message;
      throw new ApiError(message || `Request could not be completed (${response.status}).`, response.status, result);
    }
    if (!result) throw new Error('The server response could not be read.');
    return result;
  } catch (error) {
    if (error.name === 'AbortError') throw new Error('The request timed out. Check the connection and try again.');
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function normalizeScan(raw) {
  const trimmed = raw.trim();
  if (!trimmed) throw new Error('Scan an item or enter its SKU first.');
  if (!trimmed.includes('|')) return trimmed;
  const parts = trimmed.split('|');
  if (parts.length > 3) throw new Error('This QR format is not supported. Enter the item SKU instead.');
  if (parts[1]?.trim() && !/^[1-9]\d{0,30}$/.test(parts[1].trim())) {
    throw new Error('This older QR does not identify the exact variation. Enter its unique SKU instead.');
  }
  if (!parts[0]?.trim()) throw new Error('The QR does not contain an item SKU.');
  return trimmed;
}

function value(input) {
  return input === null || input === undefined ? '—' : number.format(Number(input));
}

function remaining(line) {
  return Math.max(0, Number(line?.remaining_qty ?? Number(line?.ordered_qty || 0) - Number(line?.picked_qty || 0)));
}

function supplierItem(item) {
  return item?.inventory_mode === 'SUPPLIER_BACKED_UNLIMITED';
}

function operationalItem(item) {
  return item?.active === 1 && !(supplierItem(item) && item.opening_verified !== 1);
}

function availableStock(item) {
  return Number((supplierItem(item) ? item?.free : item?.available) ?? 0);
}

function pickableQuantity(line) {
  if (!line || line.management_mode === 'WORKBOOK' || line.fulfillment_state === 'REVIEW') return 0;
  if (Number.isFinite(line.pickable_qty)) return Math.max(0, Math.min(remaining(line), Number(line.pickable_qty)));
  if (supplierItem(line)) return Math.max(0, Math.min(remaining(line), Number(line.allocated_qty || 0)));
  return remaining(line);
}

function canPickOrder(order) {
  return String(order?.payment_status).toUpperCase() === 'PAID'
    && ['AWAITING_PROCESSING', 'PROCESSING'].includes(String(order?.fulfillment_status).toUpperCase())
    && !order?.needs_review;
}

function orderNumber(order) {
  return String(order?.order_number || order?.number || order?.ecwid_order_number || order?.id || '');
}

function orderCustomer(order) {
  return order?.customer_name || order?.customer?.name || order?.customer_email || order?.email || 'Customer';
}

function reasonConfig(reason = $('#reason')?.value) {
  return REASONS[reason] || REASONS.ECWID_PICK;
}

function adjustedNote(reason, note) {
  const config = reasonConfig(reason);
  const trimmed = note.trim();
  return config.notePrefix ? `${config.notePrefix}${trimmed}` : trimmed;
}

function setError(message = '') {
  const node = $('#form-error');
  node.textContent = message;
  node.hidden = !message;
}

function showGlobalError(message = '') {
  const node = $('#global-error');
  node.textContent = message;
  node.hidden = !message;
}

let toastTimer;
function toast(message, error = false) {
  clearTimeout(toastTimer);
  const node = $('#toast');
  node.textContent = message;
  node.classList.toggle('error', error);
  node.hidden = false;
  toastTimer = setTimeout(() => { node.hidden = true; }, error ? 9000 : 6000);
}

function renderConnection() {
  const status = $('#connection-status');
  const online = navigator.onLine;
  status.classList.toggle('offline', !online);
  $('span:last-child', status).textContent = !online ? 'Offline' : state.session ? 'Ready' : 'Connecting';
  $('#paused-banner').hidden = state.session?.inventory_enabled !== false;
  updateEnabled();
}

function renderPending() {
  const banner = $('#pending-banner');
  banner.hidden = !state.pending || state.submitting;
  if (state.pending) {
    const savedReason = REASONS[state.pending.display_reason]?.label;
    const fallback = Object.values(REASONS).find(reason => reason.apiType === state.pending.payload.type)?.label || 'Stock change';
    $('#pending-description').textContent = `${savedReason || fallback} · ${state.pending.item_name || state.pending.sku || 'Item'} · ${value(state.pending.payload.quantity)} units. Retry uses the same safe submission ID.`;
  }
  updateEnabled();
}

function renderItem() {
  const host = $('#item-result');
  host.hidden = !state.item;
  if (!state.item) return;
  $('#item-sku').textContent = state.item.sku;
  $('#item-name').textContent = state.item.name;
  $('#item-location').textContent = state.item.location || 'Not set';
  $('#item-on-hand').textContent = value(state.item.on_hand);
  $('#item-available').textContent = value(supplierItem(state.item) ? state.item.free : state.item.available);
  $('#item-ecwid').textContent = state.ecwidStock
    ? state.ecwidStock.unlimited ? 'Unlimited' : value(state.ecwidStock.quantity)
    : supplierItem(state.item) ? 'Unlimited' : value(state.item.last_ecwid_quantity);
  const checked = $('#item-stock-checked');
  if (checked) {
    checked.textContent = state.ecwidStock?.checked_at
      ? `Online stock checked ${new Intl.DateTimeFormat('en-IN', { hour: 'numeric', minute: '2-digit' }).format(new Date(state.ecwidStock.checked_at))}`
      : 'Showing the latest recorded online stock.';
  }
}

function clearItem() {
  state.lookupSequence++;
  state.item = null;
  state.ecwidStock = null;
  $('#item-result').hidden = true;
  clearOrder(true);
  setError();
  updateEnabled();
}

function clearOrder(clearInput = false) {
  state.orderLookupSequence++;
  state.orderSelection = null;
  $('#order-result').hidden = true;
  $('#order-result').replaceChildren();
  if (clearInput) $('#order-id').value = '';
  updateQuantityLimit();
  updateEnabled();
}

function renderOrderSelection() {
  const host = $('#order-result');
  host.hidden = !state.orderSelection;
  if (!state.orderSelection) return;
  const { order, line } = state.orderSelection;
  const title = document.createElement('strong');
  title.textContent = `Order #${orderNumber(order)} confirmed`;
  const detail = document.createElement('span');
  detail.textContent = `${orderCustomer(order)} · ${value(remaining(line))} left to pick`;
  host.replaceChildren(title, detail);
}

function updateQuantityLimit() {
  const quantity = $('#quantity');
  const reason = $('#reason').value;
  let maximum = null;
  if (reason === 'ECWID_PICK') {
    if (state.orderSelection) maximum = pickableQuantity(state.orderSelection.line);
  } else if (reasonConfig(reason).direction < 0 && state.item) {
    maximum = availableStock(state.item);
  }
  if (maximum === null) quantity.removeAttribute('max');
  else quantity.max = String(Math.max(0, maximum));
  $('#quantity-help').textContent = maximum === null
    ? 'Enter a positive whole number.'
    : `Maximum currently available: ${value(maximum)}.`;
}

function updateReason() {
  const config = reasonConfig();
  const orderRequired = $('#reason').value === 'ECWID_PICK';
  $('#order-field').hidden = !orderRequired;
  $('#order-id').required = orderRequired;
  $('#notes').required = !!config.noteRequired;
  $('#notes-label').replaceChildren(document.createTextNode('Notes '), document.createElement('span'));
  $('span', $('#notes-label')).textContent = config.noteRequired ? '(required)' : '(optional)';
  $('#notes-help').hidden = !config.noteRequired;
  updateQuantityLimit();
  setError();
  updateEnabled();
}

function updateEnabled() {
  const locked = state.loading || state.lookingUp || state.orderLookingUp || state.submitting || !!state.pending || !navigator.onLine
    || !state.session;
  const recordingPaused = state.session?.inventory_enabled === false;
  const itemReady = operationalItem(state.item);
  const config = reasonConfig();
  const orderReady = $('#reason')?.value !== 'ECWID_PICK' || !!state.orderSelection;
  const noteReady = !config.noteRequired || !!$('#notes')?.value.trim();
  // Camera and SKU lookup stay available during a safe production preview;
  // every control that can lead to a write remains disabled while paused.
  $('#open-scanner').disabled = locked;
  $('#sku').disabled = locked;
  $('#fetch-item').disabled = locked || !$('#sku').value.trim();
  $('#quantity').disabled = locked || recordingPaused;
  $('#quantity-minus').disabled = locked || recordingPaused;
  $('#quantity-plus').disabled = locked || recordingPaused;
  $('#reason').disabled = locked || recordingPaused;
  $('#order-id').disabled = locked || recordingPaused || !itemReady;
  $('#fetch-order').disabled = locked || recordingPaused || !itemReady || !$('#order-id').value.trim();
  $('#notes').disabled = locked || recordingPaused;
  $('#submit-movement').disabled = locked || recordingPaused || !itemReady || !orderReady || !noteReady;
  $('#retry-pending').disabled = state.submitting || !navigator.onLine || recordingPaused;
}

async function lookupItem() {
  if (state.lookingUp || state.submitting || state.pending || !navigator.onLine) return;
  const raw = $('#sku').value;
  clearItem();
  const sequence = ++state.lookupSequence;
  try {
    const code = normalizeScan(raw);
    state.lookingUp = true;
    $('#fetch-item').textContent = 'Fetching…';
    updateEnabled();
    const data = await api(`/api/items/lookup?code=${encodeURIComponent(code)}`);
    if (sequence !== state.lookupSequence || $('#sku').value !== raw) return;
    if (!data.item?.id) throw new Error('No inventory item matches this SKU.');
    state.item = data.item;
    state.ecwidStock = data.ecwid_stock || null;
    renderItem();
    if (!operationalItem(state.item)) {
      setError(supplierItem(state.item) && state.item.opening_verified !== 1
        ? 'This item needs a verified opening count before it can move.'
        : 'This item is not active in the pilot. Continue using the workbook for it.');
    } else {
      $('#quantity').focus();
      $('#quantity').select();
    }
  } catch (error) {
    if (sequence === state.lookupSequence) setError(error.message);
  } finally {
    state.lookingUp = false;
    $('#fetch-item').textContent = 'Fetch';
    updateEnabled();
  }
}

async function lookupOrder() {
  if (state.orderLookingUp || state.submitting || state.pending || !operationalItem(state.item)) return;
  const raw = $('#order-id').value.trim();
  clearOrder();
  if (!/^[A-Za-z0-9_-]{1,100}$/.test(raw)) {
    setError('Enter the exact Ecwid Order ID from the pick list.');
    return;
  }
  const sequence = ++state.orderLookupSequence;
  state.orderLookingUp = true;
  $('#fetch-order').textContent = 'Finding…';
  updateEnabled();
  try {
    const data = await api(`/api/orders/${encodeURIComponent(raw)}/refresh`, { method: 'POST' });
    if (sequence !== state.orderLookupSequence || $('#order-id').value.trim() !== raw) return;
    const order = data.order;
    if (!canPickOrder(order)) throw new Error('Only a paid order awaiting processing can be picked.');
    const matches = (order.lines || []).filter(line => String(line.item_id) === String(state.item.id)
      && line.management_mode !== 'WORKBOOK' && pickableQuantity(line) > 0);
    if (!matches.length) throw new Error('This order does not contain this SKU, it is already fully picked, or the line is not ready to pick.');
    if (matches.length > 1) throw new Error('This SKU appears on more than one line in this order. Ask an administrator to review the order before picking.');
    state.orderSelection = { order, line: matches[0] };
    renderOrderSelection();
    updateQuantityLimit();
  } catch (error) {
    if (sequence === state.orderLookupSequence) setError(error.message);
  } finally {
    state.orderLookingUp = false;
    $('#fetch-order').textContent = 'Find';
    updateEnabled();
  }
}

async function refreshSelectedOrder(selection) {
  const data = await api(`/api/orders/${encodeURIComponent(selection.order.id)}/refresh`, { method: 'POST' });
  const order = data.order;
  if (!canPickOrder(order)) throw new Error('This order is no longer paid and ready for processing.');
  const line = (order.lines || []).find(candidate => String(candidate.id) === String(selection.line.id)
    && String(candidate.item_id) === String(state.item.id));
  if (!line || pickableQuantity(line) < 1) throw new Error('This item is no longer available to pick on the selected order.');
  return { order, line };
}

function validQuantity(quantity, maximum = null) {
  return Number.isSafeInteger(quantity) && quantity >= 1 && (maximum === null || quantity <= maximum);
}

async function submitMovement() {
  if (state.submitting || state.pending) return;
  setError();
  if (!navigator.onLine) { setError('Reconnect to the internet before submitting.'); return; }
  if (!operationalItem(state.item)) { setError('Fetch an active pilot item before submitting.'); return; }
  const reason = $('#reason').value;
  const config = reasonConfig(reason);
  const quantity = Number($('#quantity').value);
  const note = $('#notes').value.trim();
  if (config.noteRequired && !note) { setError('Enter a reason in Notes for this adjustment.'); $('#notes').focus(); return; }
  if (!validQuantity(quantity)) { setError('Quantity must be a positive whole number.'); $('#quantity').focus(); return; }

  let orderSelection = null;
  try {
    if (reason === 'ECWID_PICK') {
      const selected = state.orderSelection;
      if (!selected) throw new Error('Find and confirm the exact Ecwid Order ID first.');
      $('#submit-movement span').textContent = 'Checking Order…';
      state.submitting = true;
      updateEnabled();
      orderSelection = await refreshSelectedOrder(selected);
      const maximum = pickableQuantity(orderSelection.line);
      if (!validQuantity(quantity, maximum)) throw new Error(`Only ${value(maximum)} units can currently be picked on this order line.`);
    } else if (config.direction < 0) {
      const maximum = availableStock(state.item);
      if (!validQuantity(quantity, maximum)) throw new Error(`Only ${value(maximum)} uncommitted units are currently available.`);
    }

    const payload = {
      operation_id: crypto.randomUUID(),
      type: config.apiType,
      item_id: state.item.id,
      quantity,
      note: adjustedNote(reason, note),
      reason_code: config.reasonCode || '',
      ...(orderSelection ? { order_id: orderSelection.order.id, order_line_id: orderSelection.line.id } : {}),
    };
    persistPending({
      payload,
      display_reason: reason,
      item_name: state.item.name,
      sku: state.item.sku,
      created_at: new Date().toISOString(),
    });
  } catch (error) {
    setError(error.message);
  } finally {
    state.submitting = false;
    $('#submit-movement span').textContent = 'Submit Stock Change';
    updateEnabled();
  }
  if (state.pending) await sendPending();
}

function matchesMovementConfirmation(movement, payload) {
  return movement && String(movement.id).toLowerCase() === String(payload.operation_id).toLowerCase()
    && movement.type === payload.type
    && String(movement.item_id) === String(payload.item_id)
    && Number(movement.quantity) === Number(payload.quantity)
    && String(movement.order_id || '') === String(payload.order_id || '')
    && String(movement.order_line_id || '') === String(payload.order_line_id || '')
    && String(movement.note || '') === String(payload.note || '')
    && String(movement.reason_code || '') === String(payload.reason_code || '');
}

function matchesAllocationConfirmation(allocation, payload) {
  return allocation && String(allocation.id).toLowerCase() === String(payload.operation_id).toLowerCase()
    && allocation.type === payload.type
    && String(allocation.item_id) === String(payload.item_id)
    && String(allocation.order_id) === String(payload.order_id)
    && String(allocation.order_line_id) === String(payload.order_line_id)
    && Number(allocation.quantity) === Number(payload.quantity)
    && String(allocation.note || '') === String(payload.note || '');
}

async function sendPending() {
  if (!state.pending || state.submitting || !navigator.onLine) return;
  state.submitting = true;
  renderPending();
  const intent = state.pending;
  const allocation = ['ALLOCATE', 'RELEASE'].includes(intent.payload.type);
  $('#submit-movement span').textContent = 'Submitting…';
  try {
    const result = await api(allocation ? '/api/allocations' : '/api/movements', {
      method: 'POST',
      body: JSON.stringify(intent.payload),
    });
    const confirmation = allocation ? result.allocation : result.movement;
    if (!confirmation) throw new Error('The server confirmation could not be read.');
    if (allocation && !matchesAllocationConfirmation(confirmation, intent.payload)) {
      throw new Error('The server confirmation did not match the saved assignment.');
    }
    if (!allocation && !matchesMovementConfirmation(confirmation, intent.payload)) {
      throw new Error('The server confirmation did not match the saved submission.');
    }
    persistPending(null);
    toast(result.duplicate ? 'This stock change was already recorded. Nothing was recorded twice.' : 'Stock change recorded successfully.');
    resetForm();
  } catch (error) {
    const definitive = error instanceof ApiError && error.status >= 400 && error.status < 500
      && ![401, 403, 408, 429].includes(error.status);
    if (definitive) {
      persistPending(null);
      setError(error.message);
    } else {
      state.pending.status_message = error.message;
      persistPending(state.pending);
      toast('Confirmation was not received. Use Retry safely; do not enter the movement again.', true);
    }
  } finally {
    state.submitting = false;
    $('#submit-movement span').textContent = 'Submit Stock Change';
    renderPending();
  }
}

function resetForm() {
  $('#sku').value = '';
  $('#quantity').value = '1';
  $('#notes').value = '';
  $('#order-id').value = '';
  clearItem();
  updateReason();
  $('#sku').focus();
}

function stepQuantity(change) {
  const current = Number($('#quantity').value);
  const next = Math.max(1, (Number.isSafeInteger(current) ? current : 1) + change);
  const maximum = Number($('#quantity').max);
  $('#quantity').value = String(Number.isFinite(maximum) && maximum > 0 ? Math.min(next, maximum) : next);
}

async function startCamera() {
  await stopScanner(false);
  const dialog = $('#scanner-dialog');
  if (!dialog.open) return;
  const generation = ++state.scannerGeneration;
  $('#scanner-help').textContent = `Starting the ${state.facingMode === 'environment' ? 'back' : 'front'} camera…`;
  $('#switch-camera').disabled = true;
  if (!window.isSecureContext) {
    $('#scanner-help').textContent = 'Camera access requires HTTPS. Close the camera and enter the SKU instead.';
    return;
  }
  if (!window.Html5Qrcode) {
    $('#scanner-help').textContent = 'The scanner could not load. Close the camera and enter the SKU instead.';
    return;
  }
  try {
    const scanner = new window.Html5Qrcode('qr-reader');
    state.scanner = scanner;
    let accepted = false;
    state.scannerStarting = scanner.start(
      { facingMode: state.facingMode },
      { fps: 8, qrbox: { width: 230, height: 230 } },
      async decoded => {
        if (accepted || generation !== state.scannerGeneration || !dialog.open) return;
        accepted = true;
        await stopScanner(true);
        $('#sku').value = decoded;
        clearItem();
        updateEnabled();
        await lookupItem();
      },
    );
    await state.scannerStarting;
    state.scannerStarting = null;
    if (generation !== state.scannerGeneration || !dialog.open) {
      try { if (scanner.isScanning) await scanner.stop(); } catch { /* The scanner already stopped. */ }
      return;
    }
    $('#scanner-help').textContent = 'Point the camera at the QR code on the item box.';
    $('#switch-camera').disabled = false;
  } catch {
    state.scannerStarting = null;
    if (generation !== state.scannerGeneration) return;
    $('#scanner-help').textContent = 'Camera access is unavailable. Allow camera access, switch camera, or enter the SKU instead.';
    $('#switch-camera').disabled = false;
  }
}

async function openScanner() {
  if ($('#open-scanner').disabled) return;
  state.facingMode = 'environment';
  $('#switch-camera').replaceChildren(icon('camera-switch'), document.createTextNode('Use front camera'));
  const dialog = $('#scanner-dialog');
  if (!dialog.open) dialog.showModal();
  await startCamera();
}

async function switchCamera() {
  state.facingMode = state.facingMode === 'environment' ? 'user' : 'environment';
  $('#switch-camera').replaceChildren(icon('camera-switch'), document.createTextNode(state.facingMode === 'environment' ? 'Use front camera' : 'Use back camera'));
  await startCamera();
}

async function stopScanner(closeDialog = true) {
  state.scannerGeneration++;
  if (state.scannerStopping) {
    await state.scannerStopping;
    if (closeDialog && $('#scanner-dialog').open) $('#scanner-dialog').close();
    return;
  }
  const scanner = state.scanner;
  const starting = state.scannerStarting;
  state.scanner = null;
  state.scannerStarting = null;
  state.scannerStopping = (async () => {
    if (starting) { try { await starting; } catch { /* Camera start may be cancelled or denied. */ } }
    if (scanner) {
      try { if (scanner.isScanning) await scanner.stop(); } catch { /* Camera may already be stopped. */ }
      try { scanner.clear(); } catch { /* An unstarted scanner has no surface to clear. */ }
    }
  })();
  try { await state.scannerStopping; }
  finally {
    state.scannerStopping = null;
    if (closeDialog && $('#scanner-dialog').open) $('#scanner-dialog').close();
  }
}

async function initialize() {
  state.loading = true;
  renderConnection();
  renderPending();
  try {
    state.session = await api('/api/session');
  } catch (error) {
    showGlobalError(`The inventory workspace could not be opened. ${error.message}`);
  } finally {
    state.loading = false;
    renderConnection();
    updateReason();
    renderPending();
  }
}

$('#movement-form').addEventListener('submit', event => { event.preventDefault(); submitMovement(); });
$('#open-scanner').addEventListener('click', openScanner);
$('#close-scanner').addEventListener('click', () => stopScanner(true));
$('#switch-camera').addEventListener('click', switchCamera);
$('#scanner-manual').addEventListener('click', async () => { await stopScanner(true); $('#sku').focus(); });
$('#scanner-dialog').addEventListener('cancel', event => { event.preventDefault(); stopScanner(true); });
$('#sku').addEventListener('input', () => { clearItem(); updateEnabled(); });
$('#sku').addEventListener('keydown', event => { if (event.key === 'Enter') { event.preventDefault(); lookupItem(); } });
$('#fetch-item').addEventListener('click', lookupItem);
$('#quantity-minus').addEventListener('click', () => stepQuantity(-1));
$('#quantity-plus').addEventListener('click', () => stepQuantity(1));
$('#quantity').addEventListener('input', updateEnabled);
$('#reason').addEventListener('change', updateReason);
$('#order-id').addEventListener('input', () => { clearOrder(); setError(); updateEnabled(); });
$('#order-id').addEventListener('keydown', event => { if (event.key === 'Enter') { event.preventDefault(); lookupOrder(); } });
$('#fetch-order').addEventListener('click', lookupOrder);
$('#notes').addEventListener('input', updateEnabled);
$('#retry-pending').addEventListener('click', sendPending);
window.addEventListener('online', renderConnection);
window.addEventListener('offline', renderConnection);
window.addEventListener('pagehide', () => { stopScanner(true); });
document.addEventListener('visibilitychange', () => { if (document.hidden) stopScanner(true); });

initialize();
