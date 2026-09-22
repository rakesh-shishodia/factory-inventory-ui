import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { alignCutoverRow, beginCutoverAlignment, finishAndActivateCutover, recoverAndActivateCutover,
  type CutoverAlignmentPolicy, type CutoverAlignmentRequest } from '../src/cutover-alignment';
import { previewOpeningCutover, stageOpeningCutover, type OpeningCutoverRequest } from '../src/opening-cutover';
import type { EcwidFetch, EcwidOrder } from '../src/ecwid';
import { applyMigrations, sqliteD1 } from './d1';

const at = '2026-09-22T12:01:00.000Z';
const frozen = '2026-09-22T12:00:00.000Z';
const policy: CutoverAlignmentPolicy = { storeId: '12345', actor: 'admin@example.com', token: 'private-test-secret', mode: 'live',
  inventoryEnabled: 'false', liveSyncEnabled: 'false', orderSyncEnabled: 'false', now: at };
let sqlite: Database.Database; let db: D1Database;
let request: CutoverAlignmentRequest; let staged: OpeningCutoverRequest; let itemId: string;
let quantity: number; let productOverrides: Record<string, unknown>; let currentOrders: EcwidOrder[];
let fetcher: ReturnType<typeof vi.fn<EcwidFetch>>;

function input(): OpeningCutoverRequest {
  const target = { id: '123', sku: 'BOLT-1', name: 'Bolt', quantity: 12, unlimited: false, hasOptions: false,
    hasVariations: false, combinationId: null, variationOptions: [], enabled: true, eligibilityVerified: true,
    hasExtraOptions: false, hasBundleRelationships: false };
  return { operation_id: crypto.randomUUID(), expected_hash: '', confirm_staging: true, physical_counts_confirmed: true,
    freeze: { confirmed: true, started_at: frozen },
    input: { store_id: policy.storeId, source_ref: 'Verified workbook', snapshot_source_hash: 'a'.repeat(64),
      balance_meaning: 'PHYSICAL_ON_HAND', reservations_confirmed: true,
      rows: [{ sku: 'BOLT-1', balance: 10, name: 'Bolt', single_unit_confirmed: true }],
      reservations: [{ sku: 'BOLT-1', quantity: 3 }],
      catalog: { kind: 'READONLY_CATALOGUE', schema_version: 1, dry_run: true, complete: true, store_id: policy.storeId,
        started_at: '2026-09-22T12:00:01.000Z', completed_at: '2026-09-22T12:00:30.000Z',
        product_count: 1, stock_target_count: 1, products: [target], stock_targets: [target], reservations_confirmed: false } },
    scope: [{ sku: 'BOLT-1', ecwid_product_id: '123', ecwid_combination_id: null, ecwid_option_signature: '[]' }],
    orders: { kind: 'READONLY_ORDERS', schema_version: 1, dry_run: true, complete: true, store_id: policy.storeId,
      started_at: '2026-09-22T12:00:01.000Z', completed_at: '2026-09-22T12:00:30.000Z',
      creation_cutoff: Date.parse('2026-09-22T12:00:01.000Z') / 1000, orders_checked: 1, pending_order_count: 1, line_count: 1,
      orders: [{ id: '12', paymentStatus: 'PAID', fulfillmentStatus: 'AWAITING_PROCESSING', updatedAt: '2026-09-22T11:30:00.000Z',
        createdAt: '2026-09-22T11:00:00.000Z', items: [{ id: '111', productId: '123', sku: 'BOLT-1', name: 'Bolt', quantity: 3,
          combinationId: null, selectedOptions: [], digital: false, trackQuantity: true }] }] },
    line_confirmations: [{ order_id: '12', ecwid_line_id: '111', previously_picked_quantity: 0 }], workbook_scope: [] };
}
function wire(order: EcwidOrder) {
  return { id: order.id, paymentStatus: order.paymentStatus, fulfillmentStatus: order.fulfillmentStatus,
    updateTimestamp: Date.parse(order.updatedAt) / 1000,
    ...(order.createdAt ? { createTimestamp: Date.parse(order.createdAt) / 1000 } : {}), items: order.items };
}
function product() { return { id: 123, sku: 'BOLT-1', name: 'Bolt', quantity, unlimited: false, enabled: true,
  options: [], combinations: [], ...productOverrides }; }
async function transport(input: string | URL | Request, init?: RequestInit): Promise<Response> {
  const url = new URL(typeof input === 'string' || input instanceof URL ? input : input.url);
  expect(url.origin).toBe('https://app.ecwid.com');
  expect(url.pathname).toMatch(/^\/api\/v3\/12345\//);
  if (url.pathname.endsWith('/orders')) {
    expect(init?.method).toBe('GET');
    const offset = Number(url.searchParams.get('offset'));
    return Response.json({ total: currentOrders.length, count: currentOrders.slice(offset, offset + 100).length, offset,
      items: currentOrders.slice(offset, offset + 100).map(wire) });
  }
  if (init?.method === 'PUT') {
    expect(sqlite.prepare('SELECT alignment_status FROM opening_cutover_rows WHERE item_id=?').pluck().get(itemId)).toBe('PROCESSING');
    expect(sqlite.prepare('SELECT active FROM items WHERE id=?').pluck().get(itemId)).toBe(0);
    expect(url.pathname).toMatch(/^\/api\/v3\/12345\/products\/123(?:\/combinations\/501)?$/);
    const body = JSON.parse(String(init.body)); expect(Object.keys(body)).toEqual(['quantity']);
    quantity = body.quantity;
    return Response.json({ updateCount: 1 });
  }
  expect(url.pathname).toBe('/api/v3/12345/products/123');
  return Response.json(product());
}
async function stage(custom = input()) {
  staged = custom;
  staged.expected_hash = (await previewOpeningCutover(staged, policy)).review_hash;
  await stageOpeningCutover(db, staged, policy);
  request = { operation_id: staged.operation_id, expected_hash: staged.expected_hash, freeze: staged.freeze };
  itemId = sqlite.prepare('SELECT item_id FROM opening_cutover_rows WHERE operation_id=?').pluck().get(staged.operation_id) as string;
  currentOrders = structuredClone(staged.orders.orders);
}
async function begin() { return beginCutoverAlignment(db, request, policy, { fetcher }); }
async function align() { return alignCutoverRow(db, { ...request, item_id: itemId }, policy, { fetcher }); }
async function finish() { return finishAndActivateCutover(db, request, policy, { fetcher }); }
function state() { return sqlite.prepare('SELECT state FROM opening_cutover_batches').pluck().get(); }
function row() { return sqlite.prepare('SELECT alignment_status,before_quantity,after_quantity,last_error FROM opening_cutover_rows').get(); }
function writes() { return fetcher.mock.calls.filter(([, init]) => init?.method === 'PUT'); }

beforeEach(() => {
  sqlite = new Database(':memory:'); applyMigrations(sqlite); db = sqliteD1(sqlite);
  quantity = 12; productOverrides = {}; currentOrders = [];
  fetcher = vi.fn<EcwidFetch>(transport);
  vi.stubGlobal('fetch', vi.fn(() => { throw new Error('Real network is forbidden in cutover tests'); }));
});
afterEach(() => { sqlite.close(); vi.unstubAllGlobals(); });

describe('frozen opening alignment lifecycle', () => {
  it('claims before a quantity-only write and atomically activates verified shelf stock and commitments', async () => {
    await stage();
    expect(await begin()).toMatchObject({ state: 'ALIGNING', activated: false, verified_count: 0 });
    expect(await align()).toMatchObject({ state: 'ALIGNING', alignment_status: 'VERIFIED', activated: false, no_op: false });
    expect(row()).toMatchObject({ before_quantity: 12, after_quantity: 7, alignment_status: 'VERIFIED' });
    expect(await finish()).toMatchObject({ state: 'ACTIVE', activated: true, verified_count: 1 });
    expect(sqlite.prepare('SELECT active,on_hand,reserved,available,last_ecwid_quantity FROM item_stock').get())
      .toEqual({ active: 1, on_hand: 10, reserved: 3, available: 7, last_ecwid_quantity: 7 });
    expect(sqlite.prepare('SELECT status FROM sync_issues').pluck().get()).toBe('RESOLVED');
    expect(sqlite.prepare("SELECT value FROM sync_state WHERE key='orders_tracking_started'").pluck().get()).toBe(frozen);
    expect(sqlite.prepare('SELECT COUNT(*) FROM movements').pluck().get()).toBe(0);
    expect(sqlite.prepare('SELECT COUNT(*) FROM outbox').pluck().get()).toBe(0);
    expect(sqlite.pragma('foreign_key_check')).toEqual([]);
    expect(writes()).toHaveLength(1); expect(fetch).not.toHaveBeenCalled();
  });

  it('does not write when the freshly verified expected quantity already equals target', async () => {
    const req = input();
    const catalogue = req.input.catalog as { products: Array<{ quantity: number }>; stock_targets: Array<{ quantity: number }> };
    catalogue.products[0].quantity = 7; catalogue.stock_targets[0].quantity = 7;
    quantity = 7;
    await stage(req); await begin();
    expect(await align()).toMatchObject({ no_op: true, alignment_status: 'VERIFIED' });
    await finish(); expect(writes()).toHaveLength(0);
  });

  it('makes verified-row and active-batch replays no-ops', async () => {
    await stage(); await begin(); await align();
    fetcher.mockClear();
    expect(await align()).toMatchObject({ duplicate: true }); expect(fetcher).not.toHaveBeenCalled();
    await finish(); fetcher.mockClear();
    expect(await finish()).toMatchObject({ duplicate: true, activated: true }); expect(fetcher).not.toHaveBeenCalled();
  });

  it('returns completed begin and verified-row receipts after the freeze expires, without requests or mutations', async () => {
    await stage(); await begin(); await align();
    const expired = { ...policy, now: '2026-09-22T13:00:00.000Z' };
    const beforeBatch = sqlite.prepare('SELECT * FROM opening_cutover_batches').get();
    const beforeRow = sqlite.prepare('SELECT * FROM opening_cutover_rows').get();
    fetcher.mockClear();
    expect(await beginCutoverAlignment(db, request, expired, { fetcher })).toMatchObject({ duplicate: true, state: 'ALIGNING' });
    expect(await alignCutoverRow(db, { ...request, item_id: itemId }, expired, { fetcher }))
      .toMatchObject({ duplicate: true, alignment_status: 'VERIFIED' });
    expect(fetcher).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT * FROM opening_cutover_batches').get()).toEqual(beforeBatch);
    expect(sqlite.prepare('SELECT * FROM opening_cutover_rows').get()).toEqual(beforeRow);
  });

  it('returns an ACTIVE receipt after freeze expiry without refreshing or mutating stock', async () => {
    await stage(); await begin(); await align(); await finish();
    const beforeBatch = sqlite.prepare('SELECT * FROM opening_cutover_batches').get();
    const beforeItem = sqlite.prepare('SELECT * FROM items').get();
    fetcher.mockClear();
    expect(await finishAndActivateCutover(db, request, { ...policy, now: '2026-09-22T13:00:00.000Z' }, { fetcher }))
      .toMatchObject({ duplicate: true, state: 'ACTIVE', activated: true });
    expect(fetcher).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT * FROM opening_cutover_batches').get()).toEqual(beforeBatch);
    expect(sqlite.prepare('SELECT * FROM items').get()).toEqual(beforeItem);
  });

  it('supports an independently stocked variation without touching its parent stock', async () => {
    const req = input(); const options = [{ name: 'Size', value: 'M8' }];
    const catalogue = req.input.catalog as { products: Record<string, unknown>[]; stock_targets: Record<string, unknown>[] };
    Object.assign(catalogue.stock_targets[0], { combinationId: '501', variationOptions: options, hasOptions: true });
    req.scope[0].ecwid_combination_id = '501'; req.scope[0].ecwid_option_signature = JSON.stringify(options);
    req.orders.orders[0].items[0].combinationId = '501'; req.orders.orders[0].items[0].selectedOptions = options;
    await stage(req); await begin();
    fetcher.mockImplementation(async (input, init) => {
      if (init?.method === 'PUT' || String(input).includes('/orders?')) return transport(input, init);
      return Response.json({ ...product(), quantity: 999, options: [{ name: 'Size', type: 'SELECT', choices: [{ text: 'M8' }] }],
        combinations: [{ id: 501, sku: 'BOLT-1', quantity, unlimited: false, options }] });
    });
    await align(); await finish();
    expect(String(writes()[0][0]).endsWith('/products/123/combinations/501')).toBe(true);
    expect(quantity).toBe(7);
  });

  it.each(['inventoryEnabled', 'liveSyncEnabled', 'orderSyncEnabled'] as const)('requires %s disabled before even reading Ecwid', async key => {
    await stage();
    await expect(beginCutoverAlignment(db, request, { ...policy, [key]: 'true' }, { fetcher }))
      .rejects.toMatchObject({ code: 'CUTOVER_FLAGS_UNSAFE' });
    expect(state()).toBe('STAGED'); expect(fetcher).not.toHaveBeenCalled();
  });

  it('treats concurrent begin requests as the same approved batch, without a spurious hold', async () => {
    await stage();
    const results = await Promise.all([begin(), begin()]);
    expect(results.map(result => result.duplicate).sort()).toEqual([false, true]);
    expect(state()).toBe('ALIGNING'); expect(writes()).toHaveLength(0);
  });

  it.each([
    { mode: 'demo' }, { token: '' }, { storeId: '999' }, { actor: 'other@example.com' },
  ])('requires exact live store and reviewed administrator configuration %s', async change => {
    await stage();
    await expect(beginCutoverAlignment(db, request, { ...policy, ...change }, { fetcher })).rejects.toThrow();
    expect(state()).toBe('STAGED'); expect(fetcher).not.toHaveBeenCalled();
  });

  it.each([
    { expected_hash: 'b'.repeat(64) }, { operation_id: crypto.randomUUID() },
    { freeze: { confirmed: false, started_at: frozen } },
    { freeze: { confirmed: true, started_at: '2026-09-22T12:00:01.000Z' } },
    { unexpected: true },
  ])('rejects a changed operation, hash or freeze %s', async change => {
    await stage(); await expect(beginCutoverAlignment(db, { ...request, ...change }, policy, { fetcher })).rejects.toThrow();
    expect(fetcher).not.toHaveBeenCalled(); expect(state()).toBe('STAGED');
  });

  it.each(['2026-09-22T11:59:59.000Z', '2026-09-22T12:15:00.000Z', '2026-09-22T12:15:01.000Z'])('fails closed on an invalid or expired freeze at %s', async now => {
    await stage();
    await expect(beginCutoverAlignment(db, request, { ...policy, now }, { fetcher })).rejects.toThrow();
    expect(fetcher).not.toHaveBeenCalled(); expect(state()).toBe('STAGED');
  });

  it.each(['pending_row', 'final_activation'])('rejects new %s work after freeze expiry without requests or mutations', async step => {
    await stage(); await begin();
    if (step === 'final_activation') await align();
    const beforeBatch = sqlite.prepare('SELECT * FROM opening_cutover_batches').get();
    const beforeRow = sqlite.prepare('SELECT * FROM opening_cutover_rows').get();
    const beforeItem = sqlite.prepare('SELECT * FROM items').get();
    const expired = { ...policy, now: '2026-09-22T12:15:00.000Z' };
    fetcher.mockClear();
    const result = step === 'pending_row'
      ? alignCutoverRow(db, { ...request, item_id: itemId }, expired, { fetcher })
      : finishAndActivateCutover(db, request, expired, { fetcher });
    await expect(result).rejects.toMatchObject({ code: 'CUTOVER_FREEZE_EXPIRED' });
    expect(fetcher).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT * FROM opening_cutover_batches').get()).toEqual(beforeBatch);
    expect(sqlite.prepare('SELECT * FROM opening_cutover_rows').get()).toEqual(beforeRow);
    expect(sqlite.prepare('SELECT * FROM items').get()).toEqual(beforeItem);
  });

  it('rechecks the lease after slow preflight and never writes after expiry', async () => {
    await stage(); await begin(); let now = at;
    fetcher.mockImplementation(async (input, init) => { const result = await transport(input, init); now = '2026-09-22T12:15:00.000Z'; return result; });
    await expect(alignCutoverRow(db, { ...request, item_id: itemId }, policy, { fetcher, clock: () => now }))
      .rejects.toMatchObject({ code: 'CUTOVER_REVIEW_REQUIRED' });
    expect(writes()).toHaveLength(0); expect(row()).toMatchObject({ alignment_status: 'BLOCKED' }); expect(state()).toBe('REVIEW');
  });
});

describe('fresh complete order verification', () => {
  it('checks all unfiltered pages with a current fixed creation cutoff', async () => {
    const req = input(); req.orders.orders_checked = 102; await stage(req);
    for (let index = 0; index < 101; index++) currentOrders.push({ ...currentOrders[0], id: String(1000 + index), fulfillmentStatus: 'SHIPPED' });
    await begin();
    const calls = fetcher.mock.calls.map(([url]) => new URL(String(url)));
    expect(calls.map(url => url.searchParams.get('offset'))).toEqual(['0', '100', '0']);
    expect(calls.every(url => url.searchParams.get('createdTo') === String(Date.parse(at) / 1000))).toBe(true);
    expect(calls.every(url => !url.searchParams.has('paymentStatus') && !url.searchParams.has('fulfillmentStatus'))).toBe(true);
  });

  it.each(['payment', 'fulfillment', 'timestamp', 'quantity', 'sku', 'new_pending', 'new_terminal', 'missing'])('blocks changed open order evidence: %s', async scenario => {
    await stage();
    if (scenario === 'payment') currentOrders[0].paymentStatus = 'AWAITING_PAYMENT';
    if (scenario === 'fulfillment') currentOrders[0].fulfillmentStatus = 'PROCESSING';
    if (scenario === 'timestamp') currentOrders[0].updatedAt = '2026-09-22T11:31:00.000Z';
    if (scenario === 'quantity') currentOrders[0].items[0].quantity = 4;
    if (scenario === 'sku') currentOrders[0].items[0].sku = 'CHANGED';
    if (scenario === 'new_pending' || scenario === 'new_terminal') currentOrders.push({ ...currentOrders[0], id: '13',
      fulfillmentStatus: scenario === 'new_terminal' ? 'SHIPPED' : 'AWAITING_PROCESSING' });
    if (scenario === 'missing') currentOrders = [];
    await expect(begin()).rejects.toMatchObject({ code: 'CUTOVER_ORDERS_CHANGED' });
    expect(state()).toBe('REVIEW'); expect(writes()).toHaveLength(0);
  });

  it.each(['duplicate', 'count', 'offset', 'over_limit', 'empty_page', 'changing_total'])('rejects inconsistent order pagination: %s', async scenario => {
    const req = input(); req.orders.orders_checked = scenario === 'duplicate' || scenario === 'empty_page' || scenario === 'changing_total' ? 2 : 1;
    await stage(req); let call = 0;
    fetcher.mockImplementation(async () => {
      call++;
      if (scenario === 'duplicate') return Response.json({ total: 2, count: 2, offset: 0, items: [wire(currentOrders[0]), wire(currentOrders[0])] });
      if (scenario === 'count') return Response.json({ total: 1, count: 0, offset: 0, items: [wire(currentOrders[0])] });
      if (scenario === 'offset') return Response.json({ total: 1, count: 1, offset: 1, items: [wire(currentOrders[0])] });
      if (scenario === 'over_limit') return Response.json({ total: 50_001, count: 1, offset: 0, items: [wire(currentOrders[0])] });
      if (scenario === 'empty_page') return Response.json({ total: 2, count: 0, offset: 0, items: [] });
      return Response.json({ total: call === 1 ? 2 : 1, count: 1, offset: call - 1, items: [wire({ ...currentOrders[0], id: String(call) })] });
    });
    await expect(begin()).rejects.toThrow(); expect(state()).toBe('REVIEW'); expect(writes()).toHaveLength(0);
  });

  it('detects a new terminal order inserted after the full scan using the final count probe', async () => {
    await stage(); let call = 0;
    fetcher.mockImplementation(async (input, init) => {
      call++;
      if (call === 2) currentOrders.push({ ...currentOrders[0], id: '13', fulfillmentStatus: 'SHIPPED' });
      return transport(input, init);
    });
    await expect(begin()).rejects.toMatchObject({ code: 'CUTOVER_ORDERS_CHANGED' }); expect(state()).toBe('REVIEW');
  });

  it('does not accept a replacement order created after the original cutoff even if count stays unchanged', async () => {
    const req = input(); req.orders.orders_checked = 2; await stage(req);
    currentOrders.push({ ...currentOrders[0], id: '99', fulfillmentStatus: 'SHIPPED', createdAt: '2026-09-22T12:00:10.000Z' });
    await expect(begin()).rejects.toMatchObject({ code: 'CUTOVER_ORDERS_CHANGED' });
  });
});

describe('one-row journal safety', () => {
  it.each([
    { quantity: 13 }, { unlimited: true }, { enabled: false }, { sku: 'OTHER' }, { id: 999 },
    { compositeComponents: [{ id: 2 }] }, { options: [{ name: 'Text', type: 'TEXTFIELD' }] },
    { defaultCombinationId: 501 },
  ])('rejects drifted stock identity, policy or expected quantity %s', async overrides => {
    await stage(); await begin(); productOverrides = overrides;
    await expect(align()).rejects.toMatchObject({ code: 'CUTOVER_REVIEW_REQUIRED' });
    expect(writes()).toHaveLength(0); expect(row()).toMatchObject({ alignment_status: 'BLOCKED' }); expect(state()).toBe('REVIEW');
  });

  it.each(['timeout', '500', 'invalid_json', 'missing_confirmation', 'readback_failure', 'readback_mismatch'])('holds uncertain writes and never retries: %s', async scenario => {
    await stage(); await begin(); let put = false;
    fetcher.mockImplementation(async (input, init) => {
      if (init?.method === 'PUT') {
        put = true;
        if (scenario === 'timeout') throw new Error('Transport may contain private-test-secret');
        if (scenario === '500') return new Response('', { status: 500 });
        if (scenario === 'invalid_json') return new Response('private-test-secret');
        if (scenario === 'missing_confirmation') return Response.json({});
        return transport(input, init);
      }
      if (put && scenario === 'readback_failure') throw new Error('Read failed');
      if (put && scenario === 'readback_mismatch') return Response.json({ ...product(), quantity: 9 });
      return transport(input, init);
    });
    await expect(align()).rejects.toMatchObject({ code: 'CUTOVER_REVIEW_REQUIRED' });
    expect(row()).toMatchObject({ alignment_status: 'UNKNOWN' }); expect(JSON.stringify(row())).not.toContain(policy.token);
    expect(state()).toBe('REVIEW'); expect(writes()).toHaveLength(1);
    await expect(align()).rejects.toThrow(); expect(writes()).toHaveLength(1);
  });

  it.each([400, 403, 429])('holds a definitively rejected HTTP %s write without retry', async status => {
    await stage(); await begin();
    fetcher.mockImplementation(async (input, init) => init?.method === 'PUT' ? new Response('', { status }) : transport(input, init));
    await expect(align()).rejects.toThrow(); expect(row()).toMatchObject({ alignment_status: 'BLOCKED' });
    expect(writes()).toHaveLength(1); expect(state()).toBe('REVIEW');
  });

  it('will not overwrite an unexpected preflight target just because it already equals the desired quantity', async () => {
    await stage(); await begin(); quantity = 7;
    await expect(align()).rejects.toThrow(); expect(writes()).toHaveLength(0); expect(state()).toBe('REVIEW');
  });

  it('never retries a persisted PROCESSING claim after a crash', async () => {
    await stage(); await begin();
    sqlite.prepare("UPDATE opening_cutover_rows SET alignment_status='PROCESSING',before_quantity=12,attempted_at=?").run(at);
    fetcher.mockClear(); await expect(align()).rejects.toMatchObject({ code: 'CUTOVER_ROW_HELD' });
    expect(fetcher).not.toHaveBeenCalled();
  });

  it('claims concurrent requests once, before the winner sends its PUT', async () => {
    await stage(); await begin();
    const results = await Promise.allSettled([align(), align()]);
    expect(results.filter(result => result.status === 'fulfilled')).toHaveLength(1);
    expect(writes()).toHaveLength(1); expect(row()).toMatchObject({ alignment_status: 'VERIFIED' });
  });

  it('does not permit selecting another batch item by ID', async () => {
    await stage(); await begin(); fetcher.mockClear();
    await expect(alignCutoverRow(db, { ...request, item_id: crypto.randomUUID() }, policy, { fetcher }))
      .rejects.toMatchObject({ code: 'CUTOVER_ITEM_INVALID' }); expect(fetcher).not.toHaveBeenCalled();
  });
});

describe('activation verification and rollback', () => {
  it('refuses activation until every row is verified', async () => {
    await stage(); await begin(); fetcher.mockClear();
    await expect(finish()).rejects.toMatchObject({ code: 'CUTOVER_NOT_VERIFIED' });
    expect(fetcher).not.toHaveBeenCalled(); expect(sqlite.prepare('SELECT active FROM items').pluck().get()).toBe(0);
  });

  it.each(['physical', 'reservation', 'picked', 'payment', 'active', 'quantity', 'issue', 'tracking'])('blocks changed database evidence: %s', async scenario => {
    await stage(); await begin(); await align();
    if (scenario === 'physical') sqlite.prepare('UPDATE items SET on_hand=11').run();
    if (scenario === 'reservation') sqlite.prepare('UPDATE order_lines SET ordered_qty=4').run();
    if (scenario === 'picked') sqlite.prepare('UPDATE order_lines SET picked_qty=1').run();
    if (scenario === 'payment') sqlite.prepare("UPDATE orders SET payment_status='CANCELLED'").run();
    if (scenario === 'active') sqlite.prepare('UPDATE items SET active=1').run();
    if (scenario === 'quantity') sqlite.prepare('UPDATE items SET last_ecwid_quantity=9').run();
    if (scenario === 'issue') sqlite.prepare("INSERT INTO sync_issues(id,item_id,kind,message,status,created_at) VALUES('issue',?,'UNRELATED','Review','OPEN',?)").run(itemId, at);
    if (scenario === 'tracking') sqlite.prepare("INSERT INTO sync_state VALUES('orders_tracking_started','2026-09-21T00:00:00.000Z',?)").run(at);
    await expect(finish()).rejects.toMatchObject({ code: 'CUTOVER_DATABASE_CHANGED' }); expect(state()).toBe('REVIEW');
    expect(sqlite.prepare("SELECT status FROM sync_issues WHERE kind='OPENING_CUTOVER_STAGED'").pluck().get()).toBe('OPEN');
  });

  it('rechecks remote quantities after all per-row confirmations', async () => {
    await stage(); await begin(); await align(); quantity = 6;
    await expect(finish()).rejects.toMatchObject({ code: 'CUTOVER_TARGET_CHANGED' });
    expect(state()).toBe('REVIEW'); expect(sqlite.prepare('SELECT active FROM items').pluck().get()).toBe(0);
    expect(writes()).toHaveLength(1);
  });

  it('rechecks the full pending order snapshot at activation', async () => {
    await stage(); await begin(); await align(); currentOrders[0].paymentStatus = 'CANCELLED';
    await expect(finish()).rejects.toMatchObject({ code: 'CUTOVER_ORDERS_CHANGED' });
    expect(state()).toBe('REVIEW'); expect(sqlite.prepare('SELECT active FROM items').pluck().get()).toBe(0);
  });

  it('atomically rejects a database race during the last remote read', async () => {
    await stage(); await begin(); await align(); let reads = 0;
    fetcher.mockImplementation(async (input, init) => {
      const response = await transport(input, init);
      if (++reads === 3) sqlite.prepare('UPDATE items SET on_hand=11').run();
      return response;
    });
    await expect(finish()).rejects.toMatchObject({ code: 'CUTOVER_REVIEW_REQUIRED' });
    expect(state()).toBe('REVIEW'); expect(sqlite.prepare('SELECT active,last_ecwid_quantity FROM items').get()).toEqual({ active: 0, last_ecwid_quantity: 12 });
    expect(sqlite.prepare("SELECT status FROM sync_issues WHERE kind='OPENING_CUTOVER_STAGED'").pluck().get()).toBe('OPEN');
    expect(sqlite.prepare('SELECT COUNT(*) FROM sync_state').pluck().get()).toBe(0);
  });

  it('rolls back every activation mutation if a late statement fails', async () => {
    await stage(); await begin(); await align();
    sqlite.exec("CREATE TRIGGER activation_test_failure BEFORE INSERT ON sync_state BEGIN SELECT RAISE(ABORT,'TEST_FAILURE'); END;");
    await expect(finish()).rejects.toMatchObject({ code: 'CUTOVER_REVIEW_REQUIRED' });
    expect(state()).toBe('REVIEW'); expect(sqlite.prepare('SELECT active,last_ecwid_quantity FROM items').get()).toEqual({ active: 0, last_ecwid_quantity: 12 });
    expect(sqlite.prepare('SELECT status FROM sync_issues').pluck().get()).toBe('OPEN');
  });
});

describe('single-session partial alignment recovery', () => {
  const recovery = () => ({ operation_id: request.operation_id, expected_hash: request.expected_hash,
    recovery_id: '50000000-0000-4000-8000-000000000005',
    recovery_freeze: { confirmed: true as const, started_at: '2026-09-22T12:20:00.000Z' } });

  it('validates preserved VERIFIED stock before any recovery write and leaves the batch ALIGNING on mismatch', async () => {
    await stage(); await begin(); await align(); quantity = 6; fetcher.mockClear();
    await expect(recoverAndActivateCutover(db, recovery(), { ...policy, now: '2026-09-22T12:20:30.000Z' }, { fetcher }))
      .rejects.toMatchObject({ code: 'CUTOVER_TARGET_CHANGED' });
    expect(writes()).toHaveLength(0); expect(state()).toBe('ALIGNING');
    expect(row()).toMatchObject({ alignment_status: 'VERIFIED', after_quantity: 7 });
  });

  it('requires a new explicit recovery freeze and keeps every live flag disabled', async () => {
    await stage(); await begin(); await align(); fetcher.mockClear();
    const now = { ...policy, now: '2026-09-22T12:20:30.000Z' };
    await expect(recoverAndActivateCutover(db, { operation_id: request.operation_id, expected_hash: request.expected_hash,
      recovery_id: recovery().recovery_id, freeze: { confirmed: true, started_at: frozen } }, now, { fetcher }))
      .rejects.toMatchObject({ code: 'CUTOVER_RECOVERY_REQUEST_INVALID' });
    for (const key of ['inventoryEnabled', 'liveSyncEnabled', 'orderSyncEnabled'] as const) {
      await expect(recoverAndActivateCutover(db, recovery(), { ...now, [key]: 'true' }, { fetcher }))
        .rejects.toMatchObject({ code: 'CUTOVER_FLAGS_UNSAFE' });
    }
    expect(fetcher).not.toHaveBeenCalled(); expect(state()).toBe('ALIGNING');
  });

  it('does not turn a pure recovery lease expiry into REVIEW', async () => {
    await stage(); await begin(); await align(); fetcher.mockClear();
    await expect(recoverAndActivateCutover(db, recovery(), { ...policy, now: '2026-09-22T12:50:00.000Z' }, { fetcher }))
      .rejects.toMatchObject({ code: 'CUTOVER_FREEZE_EXPIRED' });
    expect(fetcher).not.toHaveBeenCalled(); expect(state()).toBe('ALIGNING');
    expect(row()).toMatchObject({ alignment_status: 'VERIFIED' });
  });
});
