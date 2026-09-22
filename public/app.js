const $ = (selector, context = document) => context.querySelector(selector);
const $$ = (selector, context = document) => [...context.querySelectorAll(selector)];
const number = new Intl.NumberFormat('en-IN');
const PENDING_KEY = 'factory-inventory.pending-operation.v1';
const MOVEMENT_NAMES = { ECWID_PICK: 'Order pick', EMAIL_SALE: 'Email sale', INTERNAL_USE: 'In-house use', RESTOCK: 'Restock' };
const ALLOCATION_NAMES = { ALLOCATE: 'Assign to order', RELEASE: 'Release assignment', PICK: 'Assignment picked', CANCEL_RELEASE: 'Cancelled assignment released' };
const VIEW_COPY = {
  pick: ['Pick orders', 'Ready to pick.', 'The right item, the right quantity, the right order.'],
  movement: ['Stock movement', 'Keep things moving.', 'Record what comes in and what goes out.'],
  inventory: ['Inventory', 'A place for every item.', 'See what is on hand, committed, and available.'],
  activity: ['Activity & sync', 'Every movement, accounted for.', 'A shared record of your stock and Ecwid updates.'],
};

const state = {
  session: null, dashboard: {}, items: [], orders: [], movements: [], allocations: [], allocationForms: [], outbox: [], issues: [],
  inboxCounts: [], syncState: [], syncCheckedAt: null, syncRefreshing: false, syncSequence: 0,
  view: 'pick', orderStatus: 'PAID', orderSearch: '', inventorySearch: '', selectedOrderId: null,
  selectedOrder: null, movementType: 'EMAIL_SALE', loading: false, submitting: false,
  form: null, lookupSequence: 0, orderSequence: 0, refreshSequence: 0, inventorySequence: 0, inventoryResults: null,
  pending: readPending(), scanner: null, scannerStarting: null, scannerStopping: null, scannerGeneration: 0, scannerCallback: null, scannerOrderMode: false,
};

function readPending() {
  try {
    const value = JSON.parse(localStorage.getItem(PENDING_KEY) || 'null');
    const payload = value?.payload;
    const validType = payload && (Object.hasOwn(MOVEMENT_NAMES, payload.type) || (['ALLOCATE', 'RELEASE'].includes(payload.type) && payload.order_id && payload.order_line_id));
    return payload?.operation_id && payload?.item_id && validType ? value : null;
  } catch { return null; }
}

function persistPending(value) {
  state.pending = value;
  try {
    if (value) localStorage.setItem(PENDING_KEY, JSON.stringify(value));
    else localStorage.removeItem(PENDING_KEY);
  } catch { /* Same-tab retries still retain the operation ID when storage is unavailable. */ }
  renderPending();
}

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined && text !== null) node.textContent = String(text);
  return node;
}

function icon(name, className = '') {
  const node = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  node.setAttribute('class', `icon ${className}`.trim());
  node.setAttribute('aria-hidden', 'true');
  const use = document.createElementNS('http://www.w3.org/2000/svg', 'use');
  use.setAttribute('href', `#i-${name}`);
  node.append(use);
  return node;
}

function button(text, className, iconName) {
  const node = element('button', `button ${className || ''}`);
  node.type = 'button';
  if (iconName) node.append(icon(iconName));
  node.append(document.createTextNode(text));
  return node;
}

function badge(text, tone = 'neutral') { return element('span', `badge badge-${tone}`, text); }
function variationLabel(item) {
  if (!item.ecwid_combination_id) return '';
  try {
    const options = JSON.parse(item.ecwid_option_signature || '[]');
    return options.map(option => `${option.name}: ${option.value}`).join(' · ') || 'Independently stocked variation';
  } catch { return 'Independently stocked variation'; }
}
function value(n) { return number.format(Number(n || 0)); }
function knownValue(n) { return n === null || n === undefined ? '—' : value(n); }
function supplierItem(item) { return item?.inventory_mode === 'SUPPLIER_BACKED_UNLIMITED'; }
function unknownOpening(item) { return supplierItem(item) && item.opening_verified !== 1; }
function operationalItem(item) { return item?.active === 1 && !unknownOpening(item); }
function supplierLineEligible(line) { return !unknownOpening(line) && line.item_active === 1 && !['REVIEW', 'CLOSED'].includes(line.fulfillment_state); }
function allocationIntent(intent) { return ['ALLOCATE', 'RELEASE'].includes(intent?.payload?.type); }
function matchesAllocationConfirmation(event, payload) {
  return event && String(event.id).toLowerCase() === String(payload.operation_id).toLowerCase()
    && event.type === payload.type && sameId(event.item_id, payload.item_id)
    && sameId(event.order_id, payload.order_id) && sameId(event.order_line_id, payload.order_line_id)
    && event.quantity === payload.quantity && event.note === payload.note;
}
function sameId(a, b) { return String(a) === String(b); }
function payment(order) { return String(order?.payment_status || '').toUpperCase(); }
function remaining(line) { return Math.max(0, Number(line.remaining_qty ?? Number(line.ordered_qty || 0) - Number(line.picked_qty || 0))); }
function orderRemaining(order) { return order.lines?.reduce((sum, line) => sum + remaining(line), 0) ?? Number(order.remaining_qty || 0); }
function orderNumber(order) { return String(order.order_number || order.number || order.ecwid_order_number || order.id); }
function orderCustomer(order) { return order.customer_name || order.customer?.name || order.customer_email || order.email || 'Ecwid customer'; }
function canPick(order) {
  const fulfillment = String(order?.fulfillment_status || '').toUpperCase();
  return payment(order) === 'PAID' && !order.needs_review && ['AWAITING_PROCESSING', 'PROCESSING'].includes(fulfillment) && orderRemaining(order) > 0;
}
function canAllocate(order) {
  return !order.needs_review && ['PAID', 'AWAITING_PAYMENT'].includes(payment(order))
    && ['AWAITING_PROCESSING', 'PROCESSING'].includes(String(order.fulfillment_status).toUpperCase());
}
function pickableQuantity(line) {
  if (line.fulfillment_state === 'REVIEW') return 0;
  if (supplierItem(line) && !supplierLineEligible(line)) return 0;
  if (Number.isFinite(line.pickable_qty)) return Math.max(0, Math.min(remaining(line), line.pickable_qty));
  return supplierItem(line) ? Math.max(0, Math.min(remaining(line), Number(line.allocated_qty || 0))) : remaining(line);
}
function supplierLineStatus(line) {
  if (unknownOpening(line)) return 'Opening count required';
  if (line.fulfillment_state === 'REVIEW') return 'Needs review';
  if (line.fulfillment_state === 'CLOSED') return Number(line.allocated_qty) > 0 ? 'Needs review — assignment retained' : 'Closed — no picking';
  if (line.item_active !== 1) return 'Inactive — awaiting activation';
  if (remaining(line) === 0) return 'Fully picked';
  if (Number(line.unallocated_qty) > 0) {
    return Number(line.free_qty) > 0 ? 'Free stock available — assign explicitly' : 'Unassigned demand — awaiting supplier';
  }
  return 'Assigned to this order';
}
function actorName(actor) { return typeof actor === 'string' ? actor : actor?.name || actor?.email || actor?.id || 'Inventory team'; }
function timestamp(date, short = false) {
  if (!date) return '—';
  const parsed = new Date(date);
  if (Number.isNaN(parsed.getTime())) return '—';
  return new Intl.DateTimeFormat('en-IN', short
    ? { day: 'numeric', month: 'short' }
    : { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' }).format(parsed);
}

class ApiError extends Error {
  constructor(message, status, details) { super(message); this.status = status; this.details = details; }
}

async function api(path, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 25000);
  try {
    const response = await fetch(path, {
      ...options, credentials: 'same-origin', cache: 'no-store', signal: controller.signal,
      headers: { Accept: 'application/json', ...(options.body ? { 'Content-Type': 'application/json' } : {}), ...options.headers },
    });
    const result = await response.json().catch(() => null);
    if (!response.ok) {
      const message = typeof result?.error === 'string' ? result.error : result?.error?.message || result?.message;
      throw new ApiError(message || `Request could not be completed (${response.status}).`, response.status, result);
    }
    if (!result) throw new Error('The server response could not be read.');
    return result;
  } catch (error) {
    if (error.name === 'AbortError') throw new Error('The request timed out.');
    throw error;
  } finally { clearTimeout(timeout); }
}

let toastTimer;
function toast(message, error = false) {
  clearTimeout(toastTimer);
  const node = $('#toast');
  node.textContent = message;
  node.classList.toggle('error', error);
  node.hidden = false;
  toastTimer = setTimeout(() => { node.hidden = true; }, error ? 9000 : 6500);
}

function showGlobalError(message) {
  const node = $('#global-error');
  node.textContent = message || '';
  node.hidden = !message;
}

function renderSession() {
  if (!state.session) return;
  const name = actorName(state.session.actor);
  $('#actor-name').textContent = name;
  $('#actor-name').title = name;
  $('#actor-role').textContent = state.session.role || 'Picker';
  $('#actor-initials').textContent = name.split(/[\s@._-]+/).filter(Boolean).slice(0, 2).map(part => part[0]).join('').toUpperCase();
  $('#demo-banner').hidden = state.session.mode !== 'demo';
  $('#inventory-paused-banner').hidden = state.session.inventory_enabled !== false;
  renderConnection();
}

function renderConnection() {
  const online = navigator.onLine;
  document.body.classList.toggle('offline', !online);
  const mode = state.session?.mode === 'demo' ? 'Local demo' : state.session ? 'Live workspace' : 'Connecting';
  const pill = $('#mode-pill');
  pill.replaceChildren(element('span', 'status-dot'), document.createTextNode(online ? mode : 'Offline'));
  $('#connection-label').textContent = !online ? 'Offline · recording paused'
    : state.session?.mode === 'demo' ? 'Local demo connected'
      : state.session?.live_sync_enabled ? 'Ecwid sync enabled' : 'Live Ecwid sync paused';
  $('#connection-dot').classList.toggle('offline', !online);
  updateFormEnabled();
}

function renderStats() {
  const d = state.dashboard;
  $('#stat-physical').textContent = value(d.physical_units ?? state.items.reduce((s, x) => s + Number(x.on_hand || 0), 0));
  $('#stat-reserved').textContent = value(d.reserved_units ?? state.items.reduce((s, x) => s + Number(x.reserved || 0), 0));
  $('#stat-available').textContent = value(d.available_units ?? state.items.reduce((s, x) => s + Number(x.available || 0), 0));
  $('#stat-pending').textContent = value(d.pending_sync ?? state.outbox.filter(x => ['QUEUED', 'PENDING', 'SENDING', 'PROCESSING', 'RETRYABLE'].includes(String(x.status).toUpperCase())).length);
  const issues = Number(d.attention_count ?? state.issues.length);
  $('#stat-sync-caption').textContent = issues ? `${value(issues)} ${issues === 1 ? 'issue needs' : 'issues need'} attention` : state.session?.mode === 'demo' ? 'Demo updates stay in this workspace' : 'Tracked updates to your Ecwid stock';
  $('#nav-pick-count').textContent = value(d.pickable_orders ?? state.orders.filter(canPick).length);
  $('#nav-issue-count').textContent = value(issues);
  $('#nav-issue-count').hidden = issues === 0;
}

async function refreshWorkspace() {
  if (state.loading) return;
  state.loading = true;
  const sequence = ++state.refreshSequence;
  state.syncSequence++;
  const inventorySequence = state.inventorySequence;
  const refresh = $('#refresh-all');
  refresh.disabled = true;
  $('.icon', refresh).classList.add('spin');
  updateFormEnabled();
  showGlobalError('');
  const paths = ['/api/session', '/api/dashboard', '/api/items', '/api/orders?status=ALL', '/api/movements', '/api/sync', '/api/allocations'];
  if (state.inventorySearch.trim()) paths.push(`/api/items?search=${encodeURIComponent(state.inventorySearch)}`);
  try {
    const results = await Promise.allSettled(paths.map(path => api(path)));
    if (sequence !== state.refreshSequence) return;
    const errors = [];
    results.forEach((result, index) => {
      if (result.status === 'rejected') { errors.push(result.reason.message); return; }
      const data = result.value;
      if (index === 0) state.session = data;
      if (index === 1) state.dashboard = data.dashboard || data;
      if (index === 2) state.items = data.items || [];
      if (index === 3) state.orders = data.orders || [];
      if (index === 4) state.movements = data.movements || [];
      if (index === 5) applySyncData(data);
      if (index === 6) state.allocations = data.allocations || [];
      if (index === 7 && inventorySequence === state.inventorySequence) state.inventoryResults = data.items || [];
    });
    if (errors.length) showGlobalError(`Some workspace data could not be loaded. ${[...new Set(errors)].join(' ')} Use Refresh to try again.`);
    renderSession();
    renderStats();
    renderOrders();
    renderInventory();
    renderActivity();
    if (state.selectedOrderId !== null) {
      const order = state.orders.find(candidate => sameId(candidate.id, state.selectedOrderId));
      if (order && order.lines) { state.selectedOrder = order; renderOrderDetail(); }
      else await selectOrder(state.selectedOrderId, false);
    }
    if (!state.form || state.view === 'movement') renderMovementForm();
    if (!errors.length) $('#updated-at').textContent = `Updated ${new Intl.DateTimeFormat('en-IN', { hour: 'numeric', minute: '2-digit' }).format(new Date())}`;
  } finally {
    state.loading = false;
    refresh.disabled = false;
    $('.icon', refresh).classList.remove('spin');
    updateFormEnabled();
  }
}

function renderOrders() {
  const search = state.orderSearch.toLowerCase();
  const filtered = state.orders.filter(order => payment(order) === state.orderStatus)
    .filter(order => state.orderStatus !== 'PAID' || (orderRemaining(order) > 0 && ['AWAITING_PROCESSING', 'PROCESSING'].includes(String(order.fulfillment_status).toUpperCase())))
    .filter(order => `${orderNumber(order)} ${orderCustomer(order)}`.toLowerCase().includes(search));
  $('#order-count').textContent = value(filtered.length);
  const list = $('#order-list');
  list.replaceChildren();
  if (!filtered.length) {
    const empty = element('div', 'empty-state');
    empty.append(element('h3', '', search ? 'No matching orders' : state.orderStatus === 'PAID' ? 'All caught up.' : 'No orders awaiting payment'), element('p', '', search ? 'Try an order number or customer name.' : state.orderStatus === 'PAID' ? 'Paid orders with items left to pick will appear here.' : 'Awaiting-payment orders reserve stock until paid or cancelled.'));
    list.append(empty);
    return;
  }
  for (const order of filtered) {
    const tile = element('button', `order-tile${sameId(state.selectedOrderId, order.id) ? ' selected' : ''}`);
    tile.type = 'button';
    tile.setAttribute('aria-pressed', String(sameId(state.selectedOrderId, order.id)));
    const top = element('div', 'order-tile-top');
    top.append(element('h3', '', `#${orderNumber(order)}`), order.needs_review ? badge('Needs review', 'danger') : payment(order) === 'PAID' ? badge('Paid', 'success') : badge('Awaiting payment', 'warning'));
    const bottom = element('div', 'order-tile-bottom');
    const lineCount = order.lines?.length ?? order.items_count;
    bottom.append(element('span', '', `${value(orderRemaining(order))} units left${lineCount ? ` · ${value(lineCount)} ${lineCount === 1 ? 'item' : 'items'}` : ''}`), icon('chevron'));
    tile.append(top, element('p', 'order-tile-meta', orderCustomer(order)), bottom);
    tile.addEventListener('click', () => selectOrder(order.id, true));
    list.append(tile);
  }
}

async function selectOrder(id, scroll) {
  const sequence = ++state.orderSequence;
  state.selectedOrderId = id;
  state.selectedOrder = null;
  state.form = null;
  state.lookupSequence++;
  await stopScanner();
  renderOrders();
  $('#order-detail').replaceChildren(element('div', 'loading-state', 'Loading order…'));
  try {
    const data = await api(`/api/orders/${encodeURIComponent(id)}`);
    if (sequence !== state.orderSequence || !sameId(state.selectedOrderId, id)) return;
    state.selectedOrder = data.order;
    if (!state.selectedOrder) throw new Error('Order data was not returned.');
    const index = state.orders.findIndex(order => sameId(order.id, id));
    if (index >= 0) state.orders[index] = data.order;
    renderOrderDetail();
    renderOrders();
    if (scroll && window.innerWidth <= 900) $('#order-detail').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    if (sequence !== state.orderSequence) return;
    const empty = element('div', 'empty-state');
    const retry = button('Try again', 'button-outline', 'refresh');
    retry.addEventListener('click', () => selectOrder(id, false));
    empty.append(element('h3', '', 'Could not load this order'), element('p', '', error.message), retry);
    $('#order-detail').replaceChildren(empty);
  }
}

function renderOrderDetail() {
  const order = state.selectedOrder;
  if (!order) return;
  state.allocationForms = [];
  const host = $('#order-detail');
  host.replaceChildren();
  const heading = element('div', 'detail-heading');
  const copy = element('div');
  copy.append(element('div', 'eyebrow', 'ORDER DETAILS'), element('h3', '', `Order #${orderNumber(order)}`), element('p', '', orderCustomer(order)));
  const right = element('div', 'detail-heading-right');
  right.append(order.needs_review ? badge('Needs review', 'danger') : payment(order) === 'PAID' ? badge('Paid', 'success') : badge(payment(order).replaceAll('_', ' '), 'warning'));
  const refresh = button('Refresh order', 'button-quiet', 'refresh');
  refresh.addEventListener('click', async () => {
    refresh.disabled = true;
    try {
      await api(`/api/orders/${encodeURIComponent(order.id)}/refresh`, { method: 'POST' });
      if (!sameId(state.selectedOrderId, order.id)) return;
      await selectOrder(order.id, false);
      await refreshWorkspace();
      toast('Order refreshed.');
    } catch (error) { toast(error.message, true); refresh.disabled = false; }
  });
  right.append(refresh);
  heading.append(copy, right);
  host.append(heading);
  const linesHeading = element('div', 'order-lines-heading');
  linesHeading.append(element('span', '', 'ITEMS TO PICK'), element('span', '', 'PICKED / ORDERED'));
  host.append(linesHeading);
  const lines = element('div', 'order-lines');
  for (const line of order.lines || []) {
    const entry = element('div', 'order-line-entry');
    const row = element('button', 'order-line');
    row.type = 'button';
    row.dataset.lineId = line.id;
    row.disabled = !canPick(order) || pickableQuantity(line) === 0;
    const symbol = element('span', 'item-symbol');
    symbol.append(icon(remaining(line) === 0 ? 'check' : 'box'));
    const description = element('span', 'item-description');
    description.append(element('strong', '', line.name || line.sku), element('small', '', line.sku));
    if (variationLabel(line)) description.append(element('small', '', variationLabel(line)));
    const qty = element('span', 'line-quantity', `${value(line.picked_qty)} / ${value(line.ordered_qty)}`);
    qty.append(element('small', '', remaining(line) === 0 ? 'Complete' : `${value(remaining(line))} remaining`));
    row.append(symbol, description, qty);
    row.addEventListener('click', () => {
      if (!state.form || state.pending || state.submitting) return;
      state.form.code.value = line.sku;
      clearLookup(state.form);
      lookupItem(state.form, line.id);
    });
    entry.append(row);
    if (supplierItem(line)) entry.append(createSupplierLineControls(order, line));
    lines.append(entry);
  }
  host.append(lines);
  const ordered = (order.lines || []).reduce((sum, line) => sum + Number(line.ordered_qty || 0), 0);
  const picked = (order.lines || []).reduce((sum, line) => sum + Number(line.picked_qty || 0), 0);
  const progress = element('div', 'order-progress');
  const track = element('div', 'progress-track');
  const bar = element('div', 'progress-bar');
  bar.style.width = `${Math.min(100, ordered ? picked / ordered * 100 : 0)}%`;
  track.append(bar);
  progress.append(track, element('span', 'progress-caption', `${value(picked)} of ${value(ordered)} units picked`));
  host.append(progress);
  if ((order.lines || []).some(line => supplierItem(line) && remaining(line) > 0)) {
    host.append(element('p', 'supplier-order-note', 'Supplier items need received stock assigned to their exact order line. On Paid orders, other eligible lines may be picked. An order is not fully picked until every required line is picked.'));
  }
  if (!canPick(order)) {
    const message = order.needs_review ? 'This order needs review before picking. Check Activity & sync for the reason.'
      : payment(order) !== 'PAID' ? 'Picking is locked until this order is marked Paid in Ecwid. Received supplier stock can be assigned without picking it.'
        : orderRemaining(order) === 0 ? 'All items are picked. Shipping is managed separately in Ecwid.'
          : 'This order cannot be picked in its current status.';
    host.append(element('div', 'form-readonly', message));
    if (state.view === 'pick') state.form = null;
    queueMicrotask(updateFormEnabled);
    return;
  }
  const wrap = element('div', 'pick-form-wrap');
  const title = element('div', 'form-eyebrow');
  title.append(icon('scan'), document.createTextNode('VERIFY & PICK AN ITEM'));
  wrap.append(title, createMovementForm('ECWID_PICK', order));
  host.append(wrap);
}

function createSupplierLineControls(order, line) {
  const panel = element('section', 'supplier-line-panel');
  panel.setAttribute('aria-label', `Supplier stock for ${line.sku}`);
  const top = element('div', 'supplier-line-heading');
  top.append(badge('Supplier-backed', 'neutral'), element('span', '', supplierLineStatus(line)));
  const counts = element('div', 'supplier-line-counts');
  for (const [label, count] of [['On shelf', unknownOpening(line) ? null : line.on_hand], ['Assigned here', line.allocated_qty], ['Free, all orders', unknownOpening(line) ? null : line.free_qty], ['Unassigned here', line.unallocated_qty]]) {
    const part = element('span', '', label); part.append(element('strong', '', knownValue(count))); counts.append(part);
  }
  panel.append(top, counts);
  if (!unknownOpening(line) && Number(line.free_qty) > 0 && Number(line.unallocated_qty) > 0) {
    panel.append(element('p', 'field-hint', 'Free stock is shared across orders. It is not assigned to this line until you confirm.'));
  }
  if (!canAllocate(order) || !line.item_id || remaining(line) === 0 || !supplierLineEligible(line)) return panel;
  const form = element('form', 'allocation-form');
  form.noValidate = true;
  const quantityLabel = element('label', 'field-label', 'Assignment quantity');
  const quantity = element('input', 'quantity-input');
  quantity.type = 'number'; quantity.min = '1'; quantity.step = '1'; quantity.inputMode = 'numeric'; quantity.value = '1';
  quantity.setAttribute('aria-label', `Assignment quantity for ${line.sku}`);
  quantityLabel.append(quantity);
  const noteLabel = element('label', 'field-label', 'Note (optional)');
  const note = element('input', 'text-input');
  note.type = 'text'; note.maxLength = 500; note.placeholder = 'Assignment or release reason';
  note.setAttribute('aria-label', `Assignment note for ${line.sku}`);
  noteLabel.append(note);
  const controls = element('div', 'allocation-actions');
  const assign = button('Assign to this order', 'button-outline', 'plus');
  const release = button('Release assignment', 'button-quiet');
  const assignMax = Math.max(0, Math.min(Number(line.free_qty || 0), Number(line.unallocated_qty || 0)));
  const releaseMax = Math.max(0, Number(line.allocated_qty || 0));
  quantity.max = String(Math.max(assignMax, releaseMax));
  const error = element('div', 'form-error'); error.hidden = true; error.setAttribute('role', 'alert');
  const controller = { node: form, order, line, quantity, note, assign, release, error, assignMax, releaseMax };
  state.allocationForms.push(controller);
  assign.addEventListener('click', () => submitAllocation(controller, 'ALLOCATE'));
  release.addEventListener('click', () => submitAllocation(controller, 'RELEASE'));
  form.addEventListener('submit', event => event.preventDefault());
  controls.append(assign, release);
  form.append(quantityLabel, noteLabel, controls, error);
  panel.append(form, element('p', 'field-hint', 'Assignment holds shelf stock for this order. It does not pick, ship or change Ecwid.'));
  return panel;
}

async function submitAllocation(form, type) {
  if (!form.node.isConnected || state.loading || state.submitting || state.pending) return;
  setFormError(form, '');
  if (!navigator.onLine || state.session?.inventory_enabled === false) {
    setFormError(form, 'Recording is paused. Reconnect and check the workspace status.'); return;
  }
  if (!canAllocate(form.order) || !supplierLineEligible(form.line)) { setFormError(form, 'This order line cannot accept stock assignments. Refresh and check its opening count and activation.'); return; }
  const quantity = Number(form.quantity.value);
  const limit = type === 'ALLOCATE' ? form.assignMax : form.releaseMax;
  if (!Number.isSafeInteger(quantity) || quantity < 1 || quantity > limit) {
    setFormError(form, `Enter a whole quantity from 1 to ${value(limit)} ${type === 'ALLOCATE' ? 'to assign' : 'to release'}.`); return;
  }
  const payload = { operation_id: crypto.randomUUID(), type, item_id: form.line.item_id,
    order_id: form.order.id, order_line_id: form.line.id, quantity, note: form.note.value.trim() };
  persistPending({ payload, item_name: form.line.name, sku: form.line.sku, created_at: new Date().toISOString() });
  await sendPending();
}

function renderMovementForm() {
  if (state.view !== 'movement' && state.form) return;
  const host = $('#movement-form-host');
  host.replaceChildren(createMovementForm(state.movementType));
}

function createMovementForm(type, order = null) {
  const node = element('form', 'stock-form');
  node.noValidate = true;
  const formId = type === 'ECWID_PICK' ? 'pick' : 'movement';
  const label = element('label', 'field-label', 'Item QR or SKU');
  label.htmlFor = `${formId}-code`;
  const lookupRow = element('div', 'lookup-row');
  const code = element('input', 'text-input');
  code.id = `${formId}-code`; code.type = 'text'; code.placeholder = 'Scan a QR or enter SKU'; code.autocomplete = 'off'; code.spellcheck = false; code.maxLength = 250;
  const lookup = button('Find item', 'button-outline');
  const scan = button('', 'button-outline scan-button', 'scan');
  scan.setAttribute('aria-label', 'Scan item QR with camera');
  scan.title = 'Scan item QR with camera';
  lookupRow.append(code, lookup, scan);
  const hint = element('p', 'field-hint', 'Use the exact item or variation SKU. Each size has its own stock.');
  const error = element('div', 'form-error');
  error.hidden = true; error.setAttribute('role', 'alert');
  const confirmation = element('div');
  const bottom = element('div', 'form-bottom');
  const quantityWrap = element('div');
  const quantityLabel = element('label', 'field-label', 'Quantity');
  quantityLabel.htmlFor = `${formId}-quantity`;
  const quantity = element('input', 'quantity-input');
  quantity.id = `${formId}-quantity`; quantity.type = 'number'; quantity.min = '1'; quantity.step = '1'; quantity.inputMode = 'numeric'; quantity.value = '1';
  quantityWrap.append(quantityLabel, quantity);
  const noteWrap = element('div', 'form-note');
  const noteLabel = element('label', 'field-label', type === 'EMAIL_SALE' ? 'Invoice / reference (optional)' : 'Note (optional)');
  noteLabel.htmlFor = `${formId}-note`;
  const note = element('input', 'text-input');
  note.id = `${formId}-note`; note.type = 'text'; note.maxLength = 500; note.placeholder = type === 'EMAIL_SALE' ? 'e.g. Invoice INV-1042' : type === 'INTERNAL_USE' ? 'e.g. Assembly bench' : type === 'RESTOCK' ? 'e.g. Supplier delivery' : 'Anything the team should know';
  noteWrap.append(noteLabel, note);
  bottom.append(quantityWrap, noteWrap);
  const submitRow = element('div', 'submit-row');
  const explanation = element('p', 'submit-explanation');
  const submit = button(type === 'ECWID_PICK' ? 'Confirm pick' : type === 'RESTOCK' ? 'Record restock' : 'Record movement', 'button-dark', 'check');
  submit.type = 'submit'; submit.disabled = true;
  submitRow.append(explanation, submit);
  node.append(label, lookupRow, hint, error, confirmation, bottom, submitRow);
  const form = { node, code, lookup, scan, quantity, quantityLabel, note, noteLabel, explanation, error, confirmation, submit, type, order, item: null, line: null, lookupBusy: false };
  renderMovementExplanation(form);
  if ((type === 'ECWID_PICK' && state.view === 'pick') || (type !== 'ECWID_PICK' && state.view === 'movement')) state.form = form;
  code.addEventListener('input', () => clearLookup(form));
  code.addEventListener('keydown', event => {
    if (event.key === 'Enter') { event.preventDefault(); lookupItem(form); }
  });
  lookup.addEventListener('click', () => lookupItem(form));
  scan.addEventListener('click', () => openScanner(async decoded => {
    if (state.form !== form) return;
    form.code.value = decoded;
    clearLookup(form);
    await lookupItem(form);
  }));
  node.addEventListener('submit', event => { event.preventDefault(); submitForm(form); });
  queueMicrotask(updateFormEnabled);
  return node;
}

function setFormError(form, message) {
  form.error.textContent = message || '';
  form.error.hidden = !message;
}

function renderMovementExplanation(form) {
  const supplier = supplierItem(form.item);
  form.quantityLabel.textContent = form.type === 'RESTOCK' ? 'Quantity received' : 'Quantity';
  form.noteLabel.textContent = form.type === 'RESTOCK' ? 'Delivery reference / note (optional)'
    : form.type === 'EMAIL_SALE' ? 'Invoice / reference (optional)' : 'Note (optional)';
  if (form.type === 'RESTOCK') {
    form.explanation.textContent = supplier
      ? 'Record the entire accepted delivery, not just surplus. Adds factory stock only; Ecwid stays unlimited. Assign received units in Pick orders.'
      : form.item ? 'Record the entire accepted delivery. Adds physical stock and queues an Ecwid increase.'
        : 'Record the entire accepted delivery, including units bought for an order. Verify the item to see its Ecwid policy.';
    form.submit.replaceChildren(icon('check'), document.createTextNode(supplier ? 'Record receipt' : 'Record restock'));
  } else if (form.type === 'ECWID_PICK') {
    form.explanation.textContent = supplier ? 'Picks only units assigned to this line. Reduces factory stock, with no Ecwid stock write.'
      : 'Updates physical stock only. Ecwid-order picks never send a second stock deduction.';
  } else {
    form.explanation.textContent = supplier ? 'Uses only free, unassigned factory stock. No Ecwid stock write; its unlimited setting stays unchanged.'
      : form.item ? 'Removes uncommitted physical stock and queues an Ecwid decrease.' : 'Verify the item to see its available stock and Ecwid policy.';
  }
}

function clearLookup(form) {
  state.lookupSequence++;
  form.item = null; form.line = null; form.lookupBusy = false;
  form.lookup.textContent = 'Find item';
  form.confirmation.replaceChildren();
  setFormError(form, '');
  form.quantity.removeAttribute('max');
  renderMovementExplanation(form);
  $$('.order-line').forEach(row => row.classList.remove('selected'));
  updateFormEnabled();
}

function normalizeScan(raw) {
  const trimmed = raw.trim();
  if (!trimmed) throw new Error('Enter or scan an item SKU first.');
  if (!trimmed.includes('|')) return trimmed;
  const parts = trimmed.split('|');
  if (parts.length > 3) throw new Error('This QR format is not supported. Enter the item SKU manually.');
  if (parts[1]?.trim() && !/^[1-9]\d{0,30}$/.test(parts[1].trim())) throw new Error('This legacy QR has a variation label instead of an ID. Enter the variation’s unique SKU.');
  if (!parts[0]?.trim()) throw new Error('The QR does not contain an item SKU.');
  return trimmed;
}

async function lookupItem(form, preferredLineId) {
  if (state.form !== form || state.submitting || state.pending || !navigator.onLine) return;
  const raw = form.code.value;
  clearLookup(form);
  const sequence = ++state.lookupSequence;
  try {
    const code = normalizeScan(raw);
    form.lookupBusy = true;
    form.lookup.textContent = 'Finding…';
    updateFormEnabled();
    const data = await api(`/api/items/lookup?code=${encodeURIComponent(code)}`);
    if (sequence !== state.lookupSequence || state.form !== form || form.code.value !== raw) return;
    const item = data.item;
    if (!item?.id) throw new Error('No inventory item matches this code.');
    let line = null;
    if (form.order) {
      if (!operationalItem(item)) throw new Error(unknownOpening(item) ? 'A verified opening count is required before this supplier item can be picked.' : 'This item is inactive. Ask an administrator to review its activation.');
      const matches = (form.order.lines || []).filter(candidate => sameId(candidate.item_id, item.id) && remaining(candidate) > 0);
      if (matches.length > 1 && preferredLineId === undefined) throw new Error('This SKU appears on more than one order line. Select the exact line above before picking.');
      line = (preferredLineId !== undefined && matches.find(candidate => sameId(candidate.id, preferredLineId))) || matches[0];
      if (!line) throw new Error('This item is not needed on this order, or its full quantity has already been picked.');
      if (!canPick(form.order)) throw new Error('This order is no longer eligible for picking. Refresh the order.');
      if (supplierItem(item) && pickableQuantity(line) === 0) throw new Error('Assign received stock to this order line before picking. If no free stock is available, record the supplier receipt first.');
    }
    form.item = item; form.line = line;
    const confirmation = element('div', 'item-confirmation');
    const symbol = element('span', 'item-symbol'); symbol.append(icon('box'));
    const description = element('div', 'item-description');
    description.append(element('strong', '', item.name), element('small', '', `${item.sku}${item.location ? ` · ${item.location}` : ''}`));
    if (variationLabel(item)) description.append(element('small', '', variationLabel(item)));
    if (supplierItem(item)) description.append(element('small', 'supplier-policy', `Supplier-backed${item.supplier_name ? ` · ${item.supplier_name}` : ''} · Ecwid unlimited`));
    if (!operationalItem(item)) description.append(element('small', 'inventory-hold', unknownOpening(item) ? 'Opening count required · inactive' : 'Inactive — recording disabled'));
    confirmation.append(symbol, description, icon('check'));
    const stock = element('div', 'stock-snapshot');
    const stockCounts = supplierItem(item) ? [['On shelf', unknownOpening(item) ? null : item.on_hand], ['Assigned, all orders', item.allocated], ['Free', unknownOpening(item) ? null : item.free], ['Uncovered demand, all orders', unknownOpening(item) ? null : item.uncovered_demand]]
      : [['On hand', item.on_hand], ['Reserved', item.reserved], ['Available', item.available]];
    for (const [label, count] of stockCounts) {
      const fragment = element('span', '', label); fragment.append(element('strong', '', knownValue(count))); stock.append(fragment);
    }
    if (line) {
      const fragment = element('span', '', 'Left to pick'); fragment.append(element('strong', '', value(remaining(line)))); stock.append(fragment);
      if (supplierItem(item)) {
        const assigned = element('span', '', 'Assigned to this line'); assigned.append(element('strong', '', knownValue(line.allocated_qty))); stock.append(assigned);
      }
      form.quantity.max = String(pickableQuantity(line));
      $$('.order-line').forEach(row => row.classList.toggle('selected', sameId(row.dataset.lineId, line.id)));
    } else if (form.type !== 'RESTOCK') form.quantity.max = String(Math.max(0, Number((supplierItem(item) ? item.free : item.available) || 0)));
    form.confirmation.replaceChildren(confirmation, stock);
    renderMovementExplanation(form);
    if (!operationalItem(item)) setFormError(form, unknownOpening(item) ? 'Opening count required. Unknown factory stock is not zero. Recording stays disabled until this item is verified and activated.' : 'This item is inactive. Recording stays disabled until an administrator activates it.');
    form.quantity.focus(); form.quantity.select();
  } catch (error) {
    if (sequence === state.lookupSequence && state.form === form) setFormError(form, error.message);
  } finally {
    if (sequence === state.lookupSequence && state.form === form) {
      form.lookupBusy = false;
      form.lookup.textContent = 'Find item';
      updateFormEnabled();
    }
  }
}

function updateFormEnabled() {
  $('#refresh-all').disabled = state.loading || state.submitting || state.syncRefreshing;
  $('#quick-sync-status').disabled = state.loading || state.syncRefreshing || !navigator.onLine;
  $('#refresh-sync-health').disabled = state.loading || state.syncRefreshing || !navigator.onLine;
  $('#retry-pending').disabled = state.submitting || state.loading || !navigator.onLine;
  const locked = state.loading || state.submitting || !!state.pending || !navigator.onLine || !state.session || state.session.inventory_enabled === false;
  for (const allocation of state.allocationForms) {
    const disabled = locked || !canAllocate(allocation.order) || !supplierLineEligible(allocation.line);
    allocation.quantity.disabled = disabled;
    allocation.note.disabled = disabled;
    allocation.assign.disabled = disabled || allocation.assignMax < 1;
    allocation.release.disabled = disabled || allocation.releaseMax < 1;
  }
  $$('.movement-type').forEach(control => { control.disabled = state.submitting || !!state.pending; });
  const form = state.form;
  if (!form) return;
  for (const control of [form.code, form.scan, form.quantity, form.note]) control.disabled = locked;
  form.lookup.disabled = locked || form.lookupBusy;
  form.submit.disabled = locked || !operationalItem(form.item) || form.lookupBusy;
}

async function submitForm(form) {
  if (state.form !== form || state.loading || state.submitting || state.pending) return;
  setFormError(form, '');
  if (!navigator.onLine) { setFormError(form, 'Reconnect to the internet before recording stock.'); return; }
  if (!form.item) { setFormError(form, 'Find and verify the item before recording it.'); return; }
  if (!operationalItem(form.item)) { setFormError(form, unknownOpening(form.item) ? 'A verified opening count and activation are required before recording this item.' : 'This item is inactive. Recording is disabled.'); return; }
  const quantity = Number(form.quantity.value);
  if (!Number.isSafeInteger(quantity) || quantity <= 0) { setFormError(form, 'Enter a positive whole-number quantity.'); return; }
  if (form.line && quantity > remaining(form.line)) { setFormError(form, `Only ${value(remaining(form.line))} units remain to be picked on this order line.`); return; }
  if (form.line && quantity > pickableQuantity(form.line)) { setFormError(form, supplierItem(form.item) ? 'This quantity exceeds stock assigned to this exact order line. Assign free stock first, then refresh.' : 'This quantity exceeds the stock currently available for this pick. Refresh the order.'); return; }
  const free = supplierItem(form.item) ? form.item.free : form.item.available;
  if (form.type !== 'RESTOCK' && form.type !== 'ECWID_PICK' && quantity > Number(free || 0)) { setFormError(form, 'This quantity would use stock committed to orders. Check free stock.'); return; }
  const payload = {
    operation_id: crypto.randomUUID(), type: form.type, item_id: form.item.id,
    quantity, note: form.note.value.trim(),
    ...(form.order ? { order_id: form.order.id, order_line_id: form.line.id } : {}),
  };
  persistPending({ payload, item_name: form.item.name, sku: form.item.sku, created_at: new Date().toISOString() });
  await sendPending();
}

function renderPending() {
  $('#pending-banner').hidden = !state.pending || state.submitting;
  $('#pending-description').textContent = state.pending ? `${MOVEMENT_NAMES[state.pending.payload.type] || ALLOCATION_NAMES[state.pending.payload.type]} · ${state.pending.item_name || state.pending.sku || 'Item'} · ${value(state.pending.payload.quantity)} units${state.pending.payload.order_id ? ` · Order #${state.pending.payload.order_id}` : ''}${state.pending.status_message ? ` · ${state.pending.status_message}` : ''}` : '';
  updateFormEnabled();
}

async function sendPending() {
  if (!state.pending || state.submitting || !navigator.onLine) return;
  state.submitting = true;
  const intent = state.pending;
  renderPending();
  const submit = state.form?.submit;
  const originalLabel = submit?.textContent;
  if (submit) submit.textContent = 'Recording…';
  try {
    const isAllocation = allocationIntent(intent);
    const result = await api(isAllocation ? '/api/allocations' : '/api/movements', { method: 'POST', body: JSON.stringify(intent.payload) });
    if (isAllocation ? !result.allocation : !result.movement) throw new Error('The recording confirmation could not be read.');
    if (isAllocation && !matchesAllocationConfirmation(result.allocation, intent.payload)) throw new Error('The assignment confirmation did not match the saved submission.');
    persistPending(null);
    const status = String(result.sync_status || result.movement?.sync_status || '').toUpperCase();
    const text = result.duplicate ? 'This submission was already recorded. Nothing was recorded twice.'
      : isAllocation ? `${value(intent.payload.quantity)} units ${intent.payload.type === 'ALLOCATE' ? 'assigned to this order' : 'released to free stock'}. No shelf quantity or Ecwid stock change.`
      : intent.payload.type === 'ECWID_PICK' ? `${value(intent.payload.quantity)} ${intent.payload.quantity === 1 ? 'unit' : 'units'} picked. Physical stock updated.`
        : status === 'NOT_REQUIRED' ? 'Factory stock updated. No Ecwid stock write is required.'
        : ['UNKNOWN', 'FAILED', 'REJECTED', 'BLOCKED'].includes(status) ? 'Movement recorded. Ecwid sync needs attention—check Activity & sync.'
          : ['APPLIED', 'SYNCED', 'SIMULATED', 'DEMO_APPLIED', 'APPLIED_DEMO'].includes(status) ? 'Movement recorded and stock updated.'
            : 'Movement recorded. Ecwid update is pending.';
    toast(text);
    if (state.form) { state.form.code.value = ''; state.form.quantity.value = '1'; state.form.note.value = ''; clearLookup(state.form); }
    await refreshWorkspace();
  } catch (error) {
    // A definitive validation response confirms no operation was accepted. Ambiguous transport or
    // server failures retain exactly the same operation ID and body for a safe retry.
    if (error instanceof ApiError && error.status >= 400 && error.status < 500 && ![401, 403, 408, 429].includes(error.status) && error.details?.code !== 'IDEMPOTENCY_CONFLICT') {
      persistPending(null);
      if (state.form) setFormError(state.form, error.message);
      toast(error.message, true);
    } else {
      const message = error instanceof ApiError && [401, 403].includes(error.status)
        ? 'Sign in with the same account, then retry this saved submission.'
        : error.details?.code === 'IDEMPOTENCY_CONFLICT'
          ? 'This operation was already used with different details. Ask an administrator to review it before recording more stock.'
          : 'Confirmation was interrupted. Use Retry submission to check or record the same saved operation safely.';
      persistPending({ ...intent, status_message: message });
      toast(message, true);
    }
  } finally {
    state.submitting = false;
    if (submit?.isConnected && originalLabel) submit.textContent = originalLabel;
    renderPending();
    updateFormEnabled();
  }
}

function renderInventory() {
  const search = state.inventorySearch.toLowerCase();
  const items = state.inventoryResults ?? state.items.filter(item => `${item.name} ${item.sku} ${item.location || ''}`.toLowerCase().includes(search));
  const total = Number(state.dashboard.items_count ?? state.items.length);
  $('#inventory-count').textContent = !search && total > items.length ? `${value(items.length)} of ${value(total)}` : value(items.length);
  const body = $('#inventory-rows'); body.replaceChildren();
  $('#inventory-empty').hidden = items.length > 0;
  for (const item of items) {
    const row = element('tr');
    const first = element('td');
    const itemRow = element('div', 'table-item');
    const symbol = element('span', 'item-symbol'); symbol.append(icon('box'));
    const description = element('div', 'item-description');
    description.append(element('strong', '', item.name), element('small', '', item.sku));
    if (variationLabel(item)) description.append(element('small', '', variationLabel(item)));
    if (supplierItem(item)) description.append(element('small', 'supplier-policy', `Supplier-backed${item.supplier_name ? ` · ${item.supplier_name}` : ''}`));
    if (!operationalItem(item)) description.append(element('small', 'inventory-hold', unknownOpening(item) ? 'Opening count required · inactive' : 'Inactive'));
    itemRow.append(symbol, description); first.append(itemRow);
    const free = supplierItem(item) ? unknownOpening(item) ? null : item.free : item.available;
    const held = supplierItem(item) ? item.allocated : item.reserved;
    const demand = element('td', 'supplier-demand-cell');
    if (supplierItem(item)) {
      demand.append(element('strong', '', `${knownValue(item.unallocated_demand)} unassigned`),
        element('small', '', unknownOpening(item) ? 'Coverage unknown — opening count required' : `${knownValue(item.uncovered_demand)} not covered by shelf stock`),
        element('small', '', `Total demand: ${knownValue(item.paid_demand)} paid · ${knownValue(item.awaiting_payment_demand)} awaiting payment`));
    } else demand.textContent = '—';
    row.append(first, element('td', '', item.location || '—'), element('td', 'numeric', knownValue(unknownOpening(item) ? null : item.on_hand)), element('td', 'numeric', knownValue(held)), element('td', `numeric available-number${free !== null && Number(free) <= 0 ? ' low-stock' : ''}`, knownValue(free)), demand,
      element('td', 'numeric', supplierItem(item) ? 'Unlimited policy' : knownValue(item.last_ecwid_quantity)));
    body.append(row);
  }
}

function syncBadge(status) {
  const normalized = String(status || 'PENDING').toUpperCase();
  if (normalized === 'NOT_REQUIRED') return badge('No stock write', 'neutral');
  if (['APPLIED', 'SYNCED'].includes(normalized)) return badge('Synced', 'success');
  if (['SIMULATED', 'DEMO_APPLIED', 'APPLIED_DEMO'].includes(normalized)) return badge('Demo applied', 'success');
  if (['UNKNOWN', 'FAILED', 'REJECTED', 'BLOCKED', 'NEEDS_REVIEW'].includes(normalized)) return badge(normalized === 'UNKNOWN' ? 'Uncertain · review' : 'Needs attention', 'danger');
  if (['SENDING', 'PROCESSING', 'IN_FLIGHT'].includes(normalized)) return badge('Sending', 'warning');
  return badge(normalized === 'RETRYABLE' ? 'Retry pending' : 'Pending', 'warning');
}

function renderActivity() {
  renderSyncHealth();
  const body = $('#activity-rows'); body.replaceChildren();
  $('#activity-empty').hidden = state.movements.length > 0;
  $('#movement-count').textContent = `${value(state.movements.length)} RECENT MOVEMENTS`;
  for (const movement of state.movements) {
    const row = element('tr');
    const type = element('td', 'movement-cell');
    type.append(element('strong', '', MOVEMENT_NAMES[movement.type] || movement.type.replaceAll('_', ' ')), element('small', '', movement.note || (movement.order_id ? `Order #${movement.order_id}` : 'Stock movement')));
    const item = element('td', 'movement-cell'); item.append(element('strong', '', movement.name || movement.sku || movement.item_id), element('small', '', movement.sku || ''));
    const delta = movement.quantity_delta ?? (movement.type === 'RESTOCK' ? Number(movement.quantity) : -Number(movement.quantity));
    const sync = element('td'); sync.append(syncBadge(movement.sync_status || (movement.type === 'ECWID_PICK' ? 'NOT_REQUIRED' : 'PENDING')));
    row.append(type, item, element('td', `numeric ${delta > 0 ? 'quantity-positive' : 'quantity-negative'}`, `${delta > 0 ? '+' : ''}${value(delta)}`), element('td', '', actorName(movement.actor)), sync, element('td', '', timestamp(movement.created_at)));
    body.append(row);
  }
  const allocationBody = $('#allocation-rows'); allocationBody.replaceChildren();
  $('#allocation-count').textContent = value(state.allocations.length);
  $('#allocation-empty').hidden = state.allocations.length > 0;
  for (const event of state.allocations) {
    const row = element('tr');
    const action = element('td', 'movement-cell');
    action.append(element('strong', '', ALLOCATION_NAMES[event.type] || String(event.type || 'Assignment').replaceAll('_', ' ')), element('small', '', event.note || 'Supplier-stock assignment'));
    const item = element('td', 'movement-cell');
    item.append(element('strong', '', event.sku || event.name || event.item_id), element('small', '', `Order #${event.order_id} · ${event.order_line_id}`));
    const delta = Number(event.quantity_delta ?? (event.type === 'ALLOCATE' ? event.quantity : -event.quantity));
    row.append(action, item, element('td', 'numeric', `${delta > 0 ? '+' : ''}${value(delta)}`), element('td', '', actorName(event.actor)), element('td', '', timestamp(event.created_at)));
    allocationBody.append(row);
  }
  $('#sync-count').textContent = value(state.outbox.length);
  const sync = $('#sync-list'); sync.replaceChildren();
  if (!state.outbox.length) sync.append(element('p', 'sync-empty', 'No queued updates. Order picks and supplier-backed stock movements do not send Ecwid stock writes.'));
  for (const job of state.outbox) {
    const row = element('div', 'sync-item'); const top = element('div', 'sync-item-top');
    const delta = Number(job.quantity_delta || 0);
    top.append(element('strong', '', `${job.sku || job.name || job.item_id} · ${delta > 0 ? '+' : ''}${value(delta)} units`), syncBadge(job.status));
    row.append(top, element('p', '', job.last_error || `${timestamp(job.created_at)}${job.attempts ? ` · ${value(job.attempts)} ${job.attempts === 1 ? 'attempt' : 'attempts'}` : ''}`));
    sync.append(row);
  }
  const openIssues = state.issues.filter(issue => !issue.resolved_at && String(issue.status).toUpperCase() !== 'RESOLVED');
  $('#issues-count').textContent = value(openIssues.length);
  const issues = $('#issues-list'); issues.replaceChildren();
  if (!openIssues.length) issues.append(element('p', 'sync-empty', 'Nothing needs attention. Stock exceptions and uncertain sync results will appear here.'));
  for (const issue of openIssues) {
    const row = element('div', 'sync-item'); const top = element('div', 'sync-item-top');
    top.append(element('strong', '', String(issue.kind || 'Inventory issue').replaceAll('_', ' ')), badge('Review', 'warning'));
    row.append(top, element('p', '', issue.message || 'Review this item or order before continuing.'), element('p', '', `${issue.order_id ? `Order #${issue.order_id} · ` : ''}${timestamp(issue.created_at)}`));
    issues.append(row);
  }
}

function applySyncData(data) {
  state.outbox = data.outbox || [];
  state.issues = data.issues || [];
  state.inboxCounts = data.inbox_counts || [];
  state.syncState = data.sync_state || [];
  state.syncCheckedAt = new Date().toISOString();
  const statuses = new Map(state.outbox.map(job => [String(job.id), job.status]));
  for (const movement of state.movements) {
    if (statuses.has(String(movement.id))) movement.sync_status = statuses.get(String(movement.id));
  }
}

function renderSyncHealth() {
  const live = state.session?.mode === 'live';
  const loaded = state.syncCheckedAt !== null;
  const latestPoll = state.syncState.find(entry => entry.key === 'orders_last_full_poll');
  const pendingIncoming = state.inboxCounts.filter(entry => ['PENDING', 'PROCESSING'].includes(String(entry.status).toUpperCase()))
    .reduce((sum, entry) => sum + Number(entry.count || 0), 0);
  const blockedIncoming = state.inboxCounts.filter(entry => ['BLOCKED', 'UNKNOWN', 'FAILED'].includes(String(entry.status).toUpperCase()))
    .reduce((sum, entry) => sum + Number(entry.count || 0), 0);
  const openIssues = Number(state.dashboard.attention_count ?? state.issues.filter(issue => !issue.resolved_at && String(issue.status).toUpperCase() !== 'RESOLVED').length);
  const pendingStock = Number(state.dashboard.pending_sync ?? state.outbox.filter(job => ['PENDING', 'PROCESSING'].includes(String(job.status).toUpperCase())).length);
  const orderText = !loaded ? '—' : live ? latestPoll ? timestamp(latestPoll.value || latestPoll.updated_at) : 'Initial order import not complete' : 'Local demo data';
  $('#sync-health-orders').textContent = orderText;
  $('#sync-health-orders').classList.toggle('sync-health-attention', loaded && live && !latestPoll);
  $('#sync-health-orders-note').textContent = live ? 'Most recent complete check of Ecwid orders' : 'Live order imports are not used in demo';
  $('#sync-health-incoming').textContent = loaded ? value(pendingIncoming) : '—';
  $('#sync-health-exceptions').textContent = loaded ? value(openIssues) : '—';
  $('#sync-health-exceptions').classList.toggle('sync-health-attention', openIssues > 0 || blockedIncoming > 0);
  $('#sync-health-exceptions-note').textContent = blockedIncoming ? `${value(blockedIncoming)} incoming ${blockedIncoming === 1 ? 'update is' : 'updates are'} blocked` : 'Stock and order exceptions';
  $('#sync-health-checked').textContent = state.syncRefreshing ? 'Checking the latest status…' : loaded ? `Status checked ${timestamp(state.syncCheckedAt)}${live && !state.session.live_sync_enabled ? ' · Ecwid stock updates paused' : ''}` : 'Waiting for sync status';
  const quick = $('#quick-sync-status');
  const text = !loaded ? 'Checking sync' : openIssues ? `${value(openIssues)} ${openIssues === 1 ? 'issue' : 'issues'}` : blockedIncoming ? `${value(blockedIncoming)} blocked`
    : live && !latestPoll ? 'Import pending' : pendingStock + pendingIncoming > 0 ? `${value(pendingStock + pendingIncoming)} pending` : live ? state.session.live_sync_enabled ? 'Sync clear' : 'Sync paused' : 'Demo sync';
  $('span', quick).textContent = state.syncRefreshing ? 'Checking…' : text;
  quick.classList.toggle('has-attention', loaded && (openIssues > 0 || blockedIncoming > 0 || (live && !latestPoll)));
  quick.setAttribute('aria-label', `${text}. Refresh sync status without changing your current form.`);
  quick.disabled = state.loading || state.syncRefreshing || !navigator.onLine;
  $('#refresh-sync-health').disabled = state.loading || state.syncRefreshing || !navigator.onLine;
  $('.icon', $('#refresh-sync-health')).classList.toggle('spin', state.syncRefreshing);
}

async function refreshSyncHealth() {
  if (state.loading || state.syncRefreshing || !navigator.onLine) return;
  state.syncRefreshing = true;
  const sequence = ++state.syncSequence;
  renderSyncHealth();
  updateFormEnabled();
  try {
    const results = await Promise.allSettled([api('/api/sync'), api('/api/dashboard')]);
    if (sequence !== state.syncSequence) return;
    if (results[0].status === 'fulfilled') applySyncData(results[0].value);
    if (results[1].status === 'fulfilled') state.dashboard = results[1].value.dashboard || results[1].value;
    renderStats();
    renderActivity();
    const failed = results.find(result => result.status === 'rejected');
    if (failed) toast(`Could not fully refresh sync status. ${failed.reason.message}`, true);
    else if (state.view !== 'activity') toast('Sync status refreshed. Your current form is unchanged.');
  } finally {
    state.syncRefreshing = false;
    renderSyncHealth();
    updateFormEnabled();
  }
}

async function openScanner(onScan, orderMode = false) {
  if (state.pending || state.submitting) return;
  await stopScanner();
  const dialog = $('#scanner-dialog');
  const generation = ++state.scannerGeneration;
  state.scannerCallback = onScan;
  state.scannerOrderMode = orderMode;
  $('.scanner-heading h2').textContent = orderMode ? 'Scan the Order ID' : 'Scan the box QR';
  $('.scanner-heading .eyebrow').textContent = orderMode ? 'FIND ORDER' : 'IDENTIFY ITEM';
  $('#scanner-manual').textContent = orderMode ? 'Enter Order ID manually instead' : 'Enter SKU manually instead';
  $('#scanner-help').textContent = 'Starting the camera…';
  dialog.showModal();
  if (!window.isSecureContext) { $('#scanner-help').textContent = 'Camera access requires HTTPS or localhost. Close this window and enter the SKU manually.'; return; }
  if (!window.Html5Qrcode) { $('#scanner-help').textContent = 'The QR scanner could not load. Close this window and enter the SKU manually.'; return; }
  try {
    const scanner = new window.Html5Qrcode('qr-reader');
    state.scanner = scanner;
    let accepted = false;
    state.scannerStarting = scanner.start({ facingMode: 'environment' }, { fps: 8, qrbox: { width: 210, height: 210 } }, async decoded => {
      if (accepted || generation !== state.scannerGeneration || !dialog.open) return;
      accepted = true;
      const callback = state.scannerCallback;
      await stopScanner();
      await callback?.(decoded);
    }, () => {});
    await state.scannerStarting;
    state.scannerStarting = null;
    if (generation === state.scannerGeneration && dialog.open) $('#scanner-help').textContent = orderMode ? 'Scan a QR containing the exact Ecwid Order ID.' : 'Point your camera at the QR code on the item box.';
    else { try { await scanner.stop(); } catch { /* Already stopped. */ } }
  } catch (error) {
    state.scannerStarting = null;
    if (generation !== state.scannerGeneration) return;
    $('#scanner-help').textContent = 'Camera access is unavailable. Allow camera access in your browser, or enter the SKU manually.';
  }
}

async function stopScanner() {
  state.scannerGeneration++;
  if (state.scannerStopping) return state.scannerStopping;
  const scanner = state.scanner;
  const starting = state.scannerStarting;
  state.scanner = null; state.scannerCallback = null;
  const dialog = $('#scanner-dialog');
  if (dialog.open) dialog.close();
  state.scannerStopping = (async () => {
    if (starting) { try { await starting; } catch { /* Camera initialization can be cancelled or denied. */ } }
    if (scanner) {
      try { if (scanner.isScanning) await scanner.stop(); } catch { /* The camera may already be stopped. */ }
      try { scanner.clear(); } catch { /* An unstarted scanner has no surface to clear. */ }
    }
    $('#qr-reader').replaceChildren();
  })();
  try { await state.scannerStopping; } finally { state.scannerStopping = null; }
}

function navigate() {
  const candidate = window.location.hash.slice(1);
  const view = VIEW_COPY[candidate] ? candidate : 'pick';
  state.view = view;
  state.lookupSequence++;
  state.form = null;
  stopScanner();
  $$('.view-section').forEach(node => { node.hidden = node.id !== `view-${view}`; });
  $$('[data-view]').forEach(node => {
    const active = node.dataset.view === view;
    node.classList.toggle('active', active);
    if (active) node.setAttribute('aria-current', 'page'); else node.removeAttribute('aria-current');
  });
  const [label, title, description] = VIEW_COPY[view];
  $('#breadcrumb-current').textContent = label;
  $('#page-title').textContent = title;
  $('#page-description').textContent = description;
  document.title = `${label} · Factory Inventory`;
  if (view === 'movement') renderMovementForm();
  if (view === 'pick' && state.selectedOrder) renderOrderDetail();
  renderPending();
}

$('#refresh-all').addEventListener('click', refreshWorkspace);
$('#refresh-sync-health').addEventListener('click', refreshSyncHealth);
$('#quick-sync-status').addEventListener('click', refreshSyncHealth);
$('#retry-pending').addEventListener('click', sendPending);
$('#order-search').addEventListener('input', event => { state.orderSearch = event.target.value; renderOrders(); });
function openOrderId(raw) {
  const id = raw.trim().replace(/^#/, '');
  if (!id || id.length > 200 || /\s/.test(id)) { toast('Enter the exact Ecwid Order ID to open it.', true); return; }
  selectOrder(id, true);
}
$('#open-order-id').addEventListener('click', () => openOrderId($('#order-search').value));
$('#order-search').addEventListener('keydown', event => { if (event.key === 'Enter') { event.preventDefault(); openOrderId(event.target.value); } });
$('#scan-order-id').addEventListener('click', () => openScanner(decoded => openOrderId(decoded), true));
let inventoryTimer;
$('#inventory-search').addEventListener('input', event => {
  state.inventorySearch = event.target.value;
  state.inventoryResults = null;
  const sequence = ++state.inventorySequence;
  const search = state.inventorySearch;
  clearTimeout(inventoryTimer);
  renderInventory();
  if (!search.trim()) return;
  inventoryTimer = setTimeout(async () => {
    try {
      const result = await api(`/api/items?search=${encodeURIComponent(search)}`);
      if (sequence !== state.inventorySequence) return;
      state.inventoryResults = result.items || [];
      renderInventory();
    } catch (error) { if (sequence === state.inventorySequence) toast(error.message, true); }
  }, 250);
});
$$('[data-order-status]').forEach(control => control.addEventListener('click', () => {
  if (state.orderStatus === control.dataset.orderStatus) return;
  state.orderStatus = control.dataset.orderStatus;
  state.orderSequence++;
  state.lookupSequence++;
  state.selectedOrderId = null;
  state.selectedOrder = null;
  if (state.view === 'pick') state.form = null;
  stopScanner();
  const empty = element('div', 'empty-state large');
  const symbol = element('div', 'empty-icon');
  symbol.append(icon(state.orderStatus === 'PAID' ? 'pick' : 'clock'));
  empty.append(symbol,
    element('h3', '', state.orderStatus === 'PAID' ? 'One order at a time.' : 'Waiting for payment.'),
    element('p', '', state.orderStatus === 'PAID' ? 'Select an order to view its items and start picking.' : 'Select an order to review its items. Picking becomes available after it is marked Paid.'));
  $('#order-detail').replaceChildren(empty);
  $$('[data-order-status]').forEach(node => {
    const selected = node.dataset.orderStatus === state.orderStatus;
    node.classList.toggle('selected', selected); node.setAttribute('aria-pressed', String(selected));
  });
  renderOrders();
  renderPending();
}));
$$('[data-movement-type]').forEach(control => control.addEventListener('click', () => {
  if (state.pending || state.submitting) return;
  state.movementType = control.dataset.movementType;
  $$('[data-movement-type]').forEach(node => {
    const selected = node.dataset.movementType === state.movementType;
    node.classList.toggle('selected', selected); node.setAttribute('aria-pressed', String(selected));
  });
  state.lookupSequence++;
  renderMovementForm();
}));
$('#close-scanner').addEventListener('click', stopScanner);
$('#scanner-manual').addEventListener('click', async () => {
  const orderMode = state.scannerOrderMode;
  await stopScanner();
  if (orderMode) $('#order-search').focus(); else state.form?.code.focus();
});
$('#scanner-dialog').addEventListener('cancel', event => { event.preventDefault(); stopScanner(); });
document.addEventListener('visibilitychange', () => { if (document.hidden) stopScanner(); });
window.addEventListener('hashchange', navigate);
window.addEventListener('offline', () => { renderConnection(); toast('You are offline. Recording is paused until you reconnect.', true); });
window.addEventListener('online', () => { renderConnection(); toast('Connection restored. Refresh to see current stock.'); });
window.addEventListener('pagehide', () => { stopScanner(); });
$('#today').textContent = new Intl.DateTimeFormat('en-IN', { weekday: 'short', day: 'numeric', month: 'short' }).format(new Date());
navigate();
refreshWorkspace().catch(error => showGlobalError(error.message));
