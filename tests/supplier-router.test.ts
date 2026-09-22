import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as auth from '../src/auth';
import { EcwidError } from '../src/ecwid';
import worker from '../src/index';
import * as sync from '../src/sync';
import { applyMigrations, sqliteD1 } from './d1';

vi.mock('../src/sync', () => ({
  enqueueSync: vi.fn().mockResolvedValue(true),
  ingestWebhook: vi.fn(), pollOrders: vi.fn(), processSyncMessage: vi.fn(), pumpSync: vi.fn(),
  refreshOrder: vi.fn().mockResolvedValue({ id: 'SUP-ORDER', needs_review: false }),
}));

let sqlite: Database.Database;
let env: Env;
let ctx: ExecutionContext;
const origin = 'http://127.0.0.1:8793';

function request(path: string, body?: unknown, headers: Record<string, string> = {}) {
  return new Request(`${origin}${path}`, { method: body === undefined ? 'GET' : 'POST',
    headers: { Origin: origin, 'Content-Type': 'application/json', ...headers },
    body: body === undefined ? undefined : JSON.stringify(body) });
}
function allocation(quantity = 2) {
  return { operation_id: crypto.randomUUID(), type: 'ALLOCATE', item_id: 'supplier',
    order_id: 'SUP-ORDER', order_line_id: 'SUP-LINE', quantity, note: 'Exact order allocation' };
}
function live(enabled = true) {
  env.ECWID_MODE = 'live';
  env.INVENTORY_ENABLED = String(enabled);
  vi.spyOn(auth, 'authenticate').mockResolvedValue({ actor: 'demo@local', role: 'picker' });
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(sync.refreshOrder).mockResolvedValue({ id: 'SUP-ORDER', needs_review: false });
  vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('External calls forbidden in supplier router tests.')));
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  sqlite.exec(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,active,inventory_mode,supplier_name)
    VALUES('supplier','SUP-TEST','Fictitious supplier item','SUP-TEST','9001',0,'SUPPLIER_BACKED_UNLIMITED','Test supplier');
    INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
    VALUES('opening','supplier',5,'Explicit fictitious shelf count','demo@local','2026-09-22T08:00:00Z');
    UPDATE items SET active=1 WHERE id='supplier';
    INSERT INTO orders(id,payment_status,remote_updated_at,updated_at)
    VALUES('SUP-ORDER','PAID','2026-09-22T08:00:00Z','2026-09-22T08:00:00Z');
    INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
    VALUES('SUP-LINE','SUP-ORDER','line','supplier','SUP-TEST','Fictitious supplier item',4);`);
  env = { DB: sqliteD1(sqlite), ECWID_MODE: 'demo', INVENTORY_ENABLED: 'false', LIVE_SYNC_ENABLED: 'false',
    ECWID_STORE_ID: '', ECWID_TOKEN: '', ECWID_CLIENT_SECRET: '', ACCESS_TEAM_DOMAIN: '', ACCESS_AUD: '', ADMIN_EMAILS: '',
    ASSETS: { fetch: vi.fn() } as unknown as Fetcher, SYNC_QUEUE: { send: vi.fn() } as unknown as Queue };
  ctx = { waitUntil: vi.fn() } as unknown as ExecutionContext;
});
afterEach(() => { sqlite.close(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('supplier allocation HTTP boundary', () => {
  it('records, lists and releases an allocation without shelf or Ecwid changes', async () => {
    const input = allocation();
    const result = await worker.fetch(request('/api/allocations', input), env, ctx);
    expect(result.status).toBe(201);
    expect(await result.json()).toMatchObject({ duplicate: false, allocation: { id: input.operation_id, quantity: 2 } });
    expect(await (await worker.fetch(request('/api/allocations'), env, ctx)).json())
      .toMatchObject({ allocations: [{ id: input.operation_id, type: 'ALLOCATE', actor: 'demo@local' }] });
    const release = await worker.fetch(request('/api/allocations', { ...input, operation_id: crypto.randomUUID(), type: 'RELEASE', quantity: 1 }), env, ctx);
    expect(release.status).toBe(201);
    expect(sqlite.prepare('SELECT on_hand,allocated,free FROM item_stock').get()).toEqual({ on_hand: 5, allocated: 1, free: 4 });
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM movements').get()).toEqual({ n: 0 });
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM outbox').get()).toEqual({ n: 0 });
    expect(sync.enqueueSync).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
  });

  it('allows awaiting-payment assignment but blocks picking until paid', async () => {
    sqlite.exec("UPDATE orders SET payment_status='AWAITING_PAYMENT'");
    expect((await worker.fetch(request('/api/allocations', allocation()), env, ctx)).status).toBe(201);
    const input = { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'supplier', quantity: 2,
      order_id: 'SUP-ORDER', order_line_id: 'SUP-LINE' };
    const blocked = await worker.fetch(request('/api/movements', input), env, ctx);
    expect(blocked.status).toBe(409);
    expect(await blocked.json()).toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
    sqlite.exec("UPDATE orders SET payment_status='PAID'");
    const picked = await worker.fetch(request('/api/movements', input), env, ctx);
    expect(picked.status).toBe(201);
    expect(await picked.json()).toMatchObject({ sync_status: 'NOT_REQUIRED' });
    expect(sqlite.prepare('SELECT on_hand,allocated,free FROM item_stock').get()).toEqual({ on_hand: 3, allocated: 0, free: 3 });
    expect(ctx.waitUntil).not.toHaveBeenCalled();
  });

  it('records the entire supplier receipt without scheduling an online stock write', async () => {
    const result = await worker.fetch(request('/api/movements', {
      operation_id: crypto.randomUUID(), type: 'RESTOCK', item_id: 'supplier', quantity: 7,
      note: 'Whole receipt including extras',
    }), env, ctx);
    expect(result.status).toBe(201);
    expect(await result.json()).toMatchObject({ sync_status: 'NOT_REQUIRED', movement: { ecwid_quantity_delta: 0 } });
    expect(sqlite.prepare('SELECT on_hand FROM items').get()).toEqual({ on_hand: 12 });
    expect(sync.enqueueSync).not.toHaveBeenCalled();
  });

  it('replays a saved allocation after cancellation and disabled cutover without a remote request', async () => {
    const input = allocation();
    await worker.fetch(request('/api/allocations', input), env, ctx);
    sqlite.exec("UPDATE orders SET payment_status='CANCELLED'");
    live(false);
    vi.mocked(sync.refreshOrder).mockRejectedValue(new EcwidError('Offline', 'RETRYABLE'));
    const replay = await worker.fetch(request('/api/allocations', input), env, ctx);
    expect(replay.status).toBe(200);
    expect(await replay.json()).toMatchObject({ duplicate: true });
    expect(sync.refreshOrder).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT allocated FROM item_stock').get()).toEqual({ allocated: 0 });
    const changed = await worker.fetch(request('/api/allocations', { ...input, quantity: 3 }), env, ctx);
    expect(changed.status).toBe(409);
    expect(await changed.json()).toMatchObject({ code: 'IDEMPOTENCY_CONFLICT' });
  });

  it('rejects another actor reusing the saved operation ID', async () => {
    const input = allocation();
    await worker.fetch(request('/api/allocations', input), env, ctx);
    vi.spyOn(auth, 'authenticate').mockResolvedValue({ actor: 'different@example.com', role: 'picker' });
    const replay = await worker.fetch(request('/api/allocations', input), env, ctx);
    expect(replay.status).toBe(409);
    expect(await replay.json()).toMatchObject({ code: 'IDEMPOTENCY_CONFLICT' });
  });

  it('blocks live allocation before cutover without refreshing orders', async () => {
    live(false);
    const result = await worker.fetch(request('/api/allocations', allocation()), env, ctx);
    expect(result.status).toBe(409);
    expect(await result.json()).toMatchObject({ code: 'CUTOVER_REQUIRED' });
    expect(sync.refreshOrder).not.toHaveBeenCalled();
  });

  it('fails closed when a fresh live order cannot be confirmed', async () => {
    live();
    vi.mocked(sync.refreshOrder).mockRejectedValue(new EcwidError('Offline', 'RETRYABLE'));
    const result = await worker.fetch(request('/api/allocations', allocation()), env, ctx);
    expect(result.status).toBe(503);
    expect(await result.json()).toMatchObject({ code: 'ECWID_UNAVAILABLE' });
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM supplier_allocation_events').get()).toEqual({ n: 0 });
  });

  it('uses a fresh live status before accepting an allocation', async () => {
    live();
    vi.mocked(sync.refreshOrder).mockImplementation(async () => {
      sqlite.exec("UPDATE orders SET payment_status='CANCELLED'");
      return { id: 'SUP-ORDER', needs_review: false };
    });
    const result = await worker.fetch(request('/api/allocations', allocation()), env, ctx);
    expect(result.status).toBe(409);
    expect(await result.json()).toMatchObject({ code: 'ORDER_NOT_ALLOCATABLE' });
    expect(sync.refreshOrder).toHaveBeenCalledWith(env, 'SUP-ORDER');
  });

  it('enforces same-origin and sign-in controls on allocations', async () => {
    const cross = await worker.fetch(request('/api/allocations', allocation(), { Origin: 'https://elsewhere.invalid' }), env, ctx);
    expect(cross.status).toBe(403);
    expect(await cross.json()).toMatchObject({ code: 'CROSS_ORIGIN' });
    env.ECWID_MODE = 'live';
    const unauthenticated = await worker.fetch(request('/api/allocations'), env, ctx);
    expect(unauthenticated.status).toBe(503);
    expect(await unauthenticated.json()).toMatchObject({ code: 'ACCESS_NOT_CONFIGURED' });
  });

  it.each([
    { quantity: 0, code: 'INVALID_QUANTITY' },
    { quantity: 5, code: 'ALLOCATION_QUANTITY_EXCEEDED' },
    { quantity: 6, code: 'INSUFFICIENT_FREE_SUPPLIER_STOCK' },
    { order_line_id: 'wrong', code: 'ALLOCATION_LINE_MISMATCH' },
    { type: 'PICK', code: 'INVALID_ALLOCATION_TYPE' },
  ])('returns a useful allocation rejection for $code', async ({ code, ...change }) => {
    const result = await worker.fetch(request('/api/allocations', { ...allocation(), ...change }), env, ctx);
    expect(result.status).toBeLessThan(500);
    expect(await result.json()).toMatchObject({ code });
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM supplier_allocation_events').get()).toEqual({ n: 0 });
  });
});
