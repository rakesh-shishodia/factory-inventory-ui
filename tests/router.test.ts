import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as auth from '../src/auth';
import { EcwidError } from '../src/ecwid';
import worker from '../src/index';
import { previewImport } from '../src/opening-import';
import * as sync from '../src/sync';
import { sqliteD1, applyMigrations } from './d1';

vi.mock('../src/sync', () => ({
  enqueueSync: vi.fn().mockResolvedValue(true),
  ingestWebhook: vi.fn().mockResolvedValue(Response.json({ received: true }, { status: 202 })),
  pollOrders: vi.fn().mockResolvedValue({ processed: 0, complete: true }),
  processSyncMessage: vi.fn().mockResolvedValue(undefined),
  pumpSync: vi.fn().mockResolvedValue({ queued: 0 }),
  refreshOrder: vi.fn().mockResolvedValue({ id: 'ORDER-1', needs_review: false }),
}));

let sqlite: Database.Database;
let env: Env;
let ctx: ExecutionContext;
let background: Promise<unknown>[];
const origin = 'http://127.0.0.1:8787';
const timestamp = '2026-09-22T08:00:00.000Z';

function request(path: string, method = 'GET', body?: unknown, headers: Record<string, string> = {}) {
  return new Request(`${origin}${path}`, { method,
    headers: { ...(body === undefined ? {} : { 'Content-Type': 'application/json', Origin: origin }), ...headers },
    body: body === undefined ? undefined : JSON.stringify(body) });
}

function pick(quantity = 1) {
  return { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'item-1', quantity,
    order_id: 'ORDER-1', order_line_id: 'LINE-1', note: '' };
}

function authorizeLive() {
  env.ECWID_MODE = 'live';
  env.INVENTORY_ENABLED = 'true';
  env.ORDER_SYNC_ENABLED = 'true';
  vi.spyOn(auth, 'authenticate').mockResolvedValue({ actor: 'demo@local', role: 'admin' });
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(sync.refreshOrder).mockResolvedValue({ id: 'ORDER-1', needs_review: false });
  // Any unexpected external network access fails the test rather than touching a store.
  vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('External network is forbidden in router tests.')));
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
    VALUES('item-1','NUT','Nut','BIN-NUT','Shelf A','1001',10,7)`).run();
  sqlite.prepare(`INSERT INTO orders(id,payment_status,remote_updated_at,updated_at)
    VALUES('ORDER-1','PAID',?,?)`).run(timestamp, timestamp);
  sqlite.prepare(`INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
    VALUES('LINE-1','ORDER-1','REMOTE-LINE','item-1','NUT','Nut',3)`).run();
  env = { DB: sqliteD1(sqlite), ECWID_MODE: 'demo', INVENTORY_ENABLED: 'false', LIVE_SYNC_ENABLED: 'false', ORDER_SYNC_ENABLED: 'false',
    ECWID_STORE_ID: '', ECWID_TOKEN: '', ECWID_CLIENT_SECRET: '', ACCESS_TEAM_DOMAIN: '', ACCESS_AUD: '', ADMIN_EMAILS: '', STAFF_EMAILS: '',
    ASSETS: { fetch: vi.fn().mockResolvedValue(new Response('<html>app</html>', { headers: { 'Content-Type': 'text/html' } })) } as unknown as Fetcher,
    SYNC_QUEUE: { send: vi.fn() } as unknown as Queue };
  background = [];
  // The route only uses waitUntil; platform lifecycle fields are not accessed.
  ctx = { waitUntil: (promise: Promise<unknown>) => background.push(promise) } as unknown as ExecutionContext;
});

afterEach(async () => {
  await Promise.all(background);
  sqlite.close();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe('picker HTTP contract', () => {
  it('forbids awaiting-payment picks without changing physical stock', async () => {
    sqlite.prepare("UPDATE orders SET payment_status='AWAITING_PAYMENT'").run();
    const response = await worker.fetch(request('/api/movements', 'POST', pick()), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
    expect(sqlite.prepare('SELECT on_hand FROM items').get()).toEqual({ on_hand: 10 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 0 });
  });

  it('records a paid partial pick once and never queues an Ecwid deduction', async () => {
    const input = pick(2);
    const response = await worker.fetch(request('/api/movements', 'POST', input), env, ctx);
    expect(response.status).toBe(201);
    expect(await response.json()).toMatchObject({ duplicate: false, sync_status: 'NOT_REQUIRED' });
    expect(sqlite.prepare('SELECT on_hand,reserved,available FROM item_stock').get()).toEqual({ on_hand: 8, reserved: 1, available: 7 });
    expect(sync.enqueueSync).not.toHaveBeenCalled();
    expect(background).toHaveLength(0);
  });

  it('rejects over-picking with a usable conflict response', async () => {
    const response = await worker.fetch(request('/api/movements', 'POST', pick(4)), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'PICK_QUANTITY_EXCEEDED' });
    expect(sqlite.prepare('SELECT picked_qty FROM order_lines').get()).toEqual({ picked_qty: 0 });
  });

  it('replays a successful pick after cancellation without fetching Ecwid again', async () => {
    const input = pick();
    expect((await worker.fetch(request('/api/movements', 'POST', input), env, ctx)).status).toBe(201);
    sqlite.prepare("UPDATE orders SET payment_status='CANCELLED'").run();
    authorizeLive();
    env.INVENTORY_ENABLED = 'false';
    vi.mocked(sync.refreshOrder).mockRejectedValue(new EcwidError('Store unavailable', 'RETRYABLE'));
    const replay = await worker.fetch(request('/api/movements', 'POST', input), env, ctx);
    expect(replay.status).toBe(200);
    expect(await replay.json()).toMatchObject({ duplicate: true, sync_status: 'NOT_REQUIRED' });
    expect(sync.refreshOrder).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT on_hand FROM items').get()).toEqual({ on_hand: 9 });
  });

  it('rejects a changed duplicate payload without refreshing the order', async () => {
    const input = pick();
    await worker.fetch(request('/api/movements', 'POST', input), env, ctx);
    authorizeLive();
    const response = await worker.fetch(request('/api/movements', 'POST', { ...input, quantity: 2 }), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'IDEMPOTENCY_CONFLICT' });
    expect(sync.refreshOrder).not.toHaveBeenCalled();
  });

  it('refreshes a new live pick before recording it and handles unavailable Ecwid', async () => {
    authorizeLive();
    vi.mocked(sync.refreshOrder).mockRejectedValue(new EcwidError('Read timed out', 'RETRYABLE'));
    const response = await worker.fetch(request('/api/movements', 'POST', pick()), env, ctx);
    expect(response.status).toBe(503);
    expect(await response.json()).toMatchObject({ code: 'ECWID_UNAVAILABLE' });
    expect(sync.refreshOrder).toHaveBeenCalledWith(env, 'ORDER-1');
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 0 });
  });

  it('uses the refreshed payment state when accepting a new live pick', async () => {
    authorizeLive();
    vi.mocked(sync.refreshOrder).mockImplementation(async () => {
      sqlite.prepare("UPDATE orders SET payment_status='AWAITING_PAYMENT'").run();
      return { id: 'ORDER-1', needs_review: false };
    });
    const response = await worker.fetch(request('/api/movements', 'POST', pick()), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
    expect(sqlite.prepare('SELECT on_hand FROM items').get()).toEqual({ on_hand: 10 });
  });

  it('queues a committed non-order movement with waitUntil', async () => {
    const id = crypto.randomUUID();
    const response = await worker.fetch(request('/api/movements', 'POST', {
      operation_id: id, type: 'INTERNAL_USE', item_id: 'item-1', quantity: 1,
    }), env, ctx);
    expect(response.status).toBe(201);
    expect(await response.json()).toMatchObject({ sync_status: 'PENDING' });
    expect(sync.enqueueSync).toHaveBeenCalledWith(env, { kind: 'outbox', id });
    expect(background).toHaveLength(1);
  });
});

describe('configuration, origin, and routing boundaries', () => {
  it('exposes accurate demo and live session flags', async () => {
    const demo = await worker.fetch(request('/api/session'), env, ctx);
    expect(await demo.json()).toMatchObject({ actor: 'demo@local', role: 'admin', mode: 'demo', inventory_enabled: true, live_sync_enabled: false });
    authorizeLive();
    env.INVENTORY_ENABLED = 'false';
    const live = await worker.fetch(request('/api/session'), env, ctx);
    expect(await live.json()).toMatchObject({ mode: 'live', inventory_enabled: false, live_sync_enabled: false });
  });

  it('blocks new live movements until cutover before fetching Ecwid', async () => {
    authorizeLive();
    env.INVENTORY_ENABLED = 'false';
    const response = await worker.fetch(request('/api/movements', 'POST', pick()), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'CUTOVER_REQUIRED' });
    expect(sync.refreshOrder).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 0 });
  });

  it.each<Record<string, string>>([
    { Origin: 'https://attacker.example' },
    { 'Sec-Fetch-Site': 'cross-site' },
    { 'Sec-Fetch-Site': 'same-site' },
  ])('blocks cross-origin writes (%j)', async headers => {
    const response = await worker.fetch(request('/api/movements', 'POST', pick(), headers), env, ctx);
    expect(response.status).toBe(403);
    expect(await response.json()).toMatchObject({ code: 'CROSS_ORIGIN' });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 0 });
  });

  it('fails closed for live API access without staff sign-in configuration', async () => {
    env.ECWID_MODE = 'live';
    const response = await worker.fetch(request('/api/session'), env, ctx);
    expect(response.status).toBe(503);
    expect(await response.json()).toMatchObject({ code: 'ACCESS_NOT_CONFIGURED' });
    expect(fetch).not.toHaveBeenCalled();
  });

  it('keeps the health probe available without authorizing inventory access', async () => {
    env.ECWID_MODE = 'live';
    const response = await worker.fetch(request('/api/health'), env, ctx);
    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({ ok: true });
  });

  it.each(['/api/does-not-exist', '/api/items/does-not-exist', '/api'])('returns JSON 404 for unknown API route %s', async path => {
    const response = await worker.fetch(request(path), env, ctx);
    expect(response.status).toBe(404);
    expect(response.headers.get('content-type')).toContain('application/json');
    expect(await response.json()).toMatchObject({ code: 'NOT_FOUND' });
    expect(env.ASSETS.fetch).not.toHaveBeenCalled();
  });

  it('delegates the app document to static assets', async () => {
    const response = await worker.fetch(request('/'), env, ctx);
    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toBe('text/html');
    expect(env.ASSETS.fetch).toHaveBeenCalledTimes(1);
  });

  it.each(['/api/sync/pump', '/api/orders/poll', '/api/import/preview', '/api/import/stage'])('requires an administrator for %s', async path => {
    vi.spyOn(auth, 'authenticate').mockResolvedValue({ actor: 'picker@example.com', role: 'picker' });
    const response = await worker.fetch(request(path, 'POST', {}), env, ctx);
    expect(response.status).toBe(403);
    expect(await response.json()).toMatchObject({ code: 'ADMIN_REQUIRED' });
    expect(sync.pumpSync).not.toHaveBeenCalled();
    expect(sync.pollOrders).not.toHaveBeenCalled();
  });

  it('returns a client error for malformed encoded order IDs', async () => {
    const response = await worker.fetch(request('/api/orders/%E0%A4%A'), env, ctx);
    expect(response.status).toBe(400);
    expect(await response.json()).toMatchObject({ code: 'INVALID_ORDER_ID' });
  });

  it('rejects a preview for another store when a store is configured', async () => {
    env.ECWID_STORE_ID = '2442119';
    const response = await worker.fetch(request('/api/import/preview', 'POST', { store_id: '987654' }), env, ctx);
    expect(response.status).toBe(400);
    expect(await response.json()).toMatchObject({ code: 'CATALOGUE_STORE_MISMATCH' });
  });

  it('includes the actual order tracking checkpoint in sync status', async () => {
    sqlite.prepare('INSERT INTO sync_state(key,value,updated_at) VALUES(?,?,?)')
      .run('orders_tracking_started', timestamp, timestamp);
    const response = await worker.fetch(request('/api/sync'), env, ctx);
    expect(await response.json()).toMatchObject({ sync_state: [{ key: 'orders_tracking_started', value: timestamp }] });
  });
});

describe('opening staging HTTP boundary', () => {
  async function stagingRequest() {
    const target = { id: '7001', combinationId: null, sku: 'PILOT-NUT', name: 'Pilot nut', quantity: 4,
      unlimited: false, enabled: true, hasOptions: false, hasVariations: false, variationOptions: [],
      eligibilityVerified: true, hasBundleRelationships: false, hasExtraOptions: false };
    const input = { store_id: '2442119', source_ref: 'Reviewed pilot workbook', snapshot_source_hash: 'a'.repeat(64),
      balance_meaning: 'PHYSICAL_ON_HAND', reservations_confirmed: true, reservations: [],
      rows: [{ sku: 'PILOT-NUT', balance: 8, single_unit_confirmed: true }],
      catalog: { kind: 'READONLY_CATALOGUE', schema_version: 1, dry_run: true, complete: true,
        store_id: '2442119', reservations_confirmed: false, started_at: timestamp, completed_at: timestamp,
        product_count: 1, stock_target_count: 1, products: [target], stock_targets: [target] } };
    return { operation_id: crypto.randomUUID(), expected_hash: (await previewImport(input)).source_hash,
      confirm_staging: true, input, scope: [{ sku: 'PILOT-NUT', ecwid_product_id: '7001',
        ecwid_combination_id: null, ecwid_option_signature: '[]' }] };
  }
  it('requires an explicitly configured store even in demo mode', async () => {
    const response = await worker.fetch(request('/api/import/stage', 'POST', await stagingRequest()), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'IMPORT_STORE_REQUIRED' });
  });
  it.each(['INVENTORY_ENABLED', 'LIVE_SYNC_ENABLED'] as const)('requires %s to be disabled', async flag => {
    env.ECWID_STORE_ID = '2442119'; env[flag] = 'true';
    const response = await worker.fetch(request('/api/import/stage', 'POST', await stagingRequest()), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'IMPORT_REQUIRES_DISABLED_LIVE_FLAGS' });
  });
  it('stages inactive stock once, preserves existing stock and never contacts Ecwid', async () => {
    env.ECWID_STORE_ID = '2442119';
    const body = await stagingRequest();
    const response = await worker.fetch(request('/api/import/stage', 'POST', body), env, ctx);
    expect(response.status).toBe(201);
    expect(await response.json()).toMatchObject({ status: 'STAGED', duplicate: false, row_count: 1,
      activated: false, reservations_loaded: false, ecwid_changed: false });
    const replay = await worker.fetch(request('/api/import/stage', 'POST', body), env, ctx);
    expect(replay.status).toBe(200);
    expect(await replay.json()).toMatchObject({ duplicate: true });
    const staged = sqlite.prepare("SELECT id,on_hand,active FROM items WHERE sku='PILOT-NUT'").get() as { id: string; on_hand: number; active: number };
    expect(staged).toMatchObject({ on_hand: 8, active: 0 });
    expect(sqlite.prepare("SELECT on_hand FROM items WHERE id='item-1'").get()).toEqual({ on_hand: 10 });
    const move = await worker.fetch(request('/api/movements', 'POST', { operation_id: crypto.randomUUID(),
      type: 'RESTOCK', item_id: staged.id, quantity: 1 }), env, ctx);
    expect(move.status).toBe(409);
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM outbox').get()).toEqual({ count: 0 });
    expect(fetch).not.toHaveBeenCalled();
    expect(sync.enqueueSync).not.toHaveBeenCalled();
  });
  it('does not permit a cross-origin opening import', async () => {
    env.ECWID_STORE_ID = '2442119';
    const response = await worker.fetch(request('/api/import/stage', 'POST', await stagingRequest(), { Origin: 'https://attacker.example' }), env, ctx);
    expect(response.status).toBe(403);
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM opening_import_batches').get()).toEqual({ count: 0 });
  });
});

describe('legacy QR compatibility', () => {
  it.each(['NUT', 'BIN-NUT', 'NUT||Shelf A', ' nut ||ignored location'])('accepts simple item code %s', async code => {
    const response = await worker.fetch(request(`/api/items/lookup?code=${encodeURIComponent(code)}`), env, ctx);
    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({ item: { id: 'item-1', sku: 'NUT', location: 'Shelf A' } });
  });

  it.each(['NUT|VARIANT|Shelf A', 'NUT|||extra', 'NUT|0', 'NUT|-1'])('rejects malformed variation QR %s', async code => {
    const response = await worker.fetch(request(`/api/items/lookup?code=${encodeURIComponent(code)}`), env, ctx);
    expect(response.status).toBe(400);
    expect(await response.json()).toMatchObject({ code: 'INVALID_VARIATION_QR' });
  });

  it('never treats a base product as a supplied variation', async () => {
    const response = await worker.fetch(request('/api/items/lookup?code=NUT%7C1'), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'VARIATION_QR_MISMATCH' });
  });
});

describe('independent variation lookup', () => {
  beforeEach(() => {
    const insert = sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,on_hand)
      VALUES(?,?,?,?,?,?,?,?)`);
    insert.run('v-small', 'NUT-M3', 'Nut', 'BIN-M3', '2001', '3001', '[{"name":"Size","value":"M3"}]', 30);
    insert.run('v-large', 'NUT-M5', 'Nut', 'BIN-M5', '2001', '3002', '[{"name":"Size","value":"M5"}]', 50);
  });

  it.each(['NUT-M3', 'BIN-M3', 'NUT-M3||Shelf A', 'NUT-M3|3001|Shelf A'])('returns the exact mapped size for %s', async code => {
    const response = await worker.fetch(request(`/api/items/lookup?code=${encodeURIComponent(code)}`), env, ctx);
    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({ item: { id: 'v-small', on_hand: 30,
      ecwid_product_id: '2001', ecwid_combination_id: '3001', ecwid_option_signature: '[{"name":"Size","value":"M3"}]' } });
  });

  it('rejects a sibling variation ID even under the same parent', async () => {
    const response = await worker.fetch(request('/api/items/lookup?code=NUT-M3%7C3002'), env, ctx);
    expect(response.status).toBe(409);
    expect(await response.json()).toMatchObject({ code: 'VARIATION_QR_MISMATCH' });
    expect(sqlite.prepare("SELECT on_hand FROM items WHERE id IN ('v-small','v-large') ORDER BY id").all())
      .toEqual([{ on_hand: 50 }, { on_hand: 30 }]);
  });
});
