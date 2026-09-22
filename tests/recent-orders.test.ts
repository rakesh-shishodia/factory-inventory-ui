import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EcwidClient, parseOrder, type EcwidFetch } from '../src/ecwid';
import { createMovement } from '../src/inventory';
import { pollOrders, pumpSync, upsertOrderSnapshot, type SyncEnv } from '../src/sync';
import { applyMigrations, sqliteD1 } from './d1';

const clock = Date.parse('2030-01-01T12:00:00.000Z');
const seconds = clock / 1000;
let sqlite: Database.Database;
let env: SyncEnv;
let queries: number;

function state(key: string): string | null {
  return (sqlite.prepare('SELECT value FROM sync_state WHERE key=?').get(key) as { value: string } | undefined)?.value ?? null;
}
function setState(key: string, value: string) {
  sqlite.prepare(`INSERT INTO sync_state(key,value,updated_at) VALUES(?,?,?)
    ON CONFLICT(key) DO UPDATE SET value=excluded.value`).run(key, value, new Date().toISOString());
}
function checkpoint() { return JSON.parse(state('orders_recent_poll_cursor') || 'null'); }
function wire(id: string, updateTimestamp = seconds - 30, paymentStatus = 'PAID', fulfillmentStatus = 'AWAITING_PROCESSING') {
  return { id, paymentStatus, fulfillmentStatus, createTimestamp: seconds - 86400 * 100, updateTimestamp,
    items: [{ id: 'line', productId: 1001, sku: 'NUT', name: 'Nut', quantity: 2 }] };
}
type WireOrder = ReturnType<typeof wire>;
function catalogue(rows: WireOrder[], pageMaximum = 100) {
  return vi.fn<EcwidFetch>(async input => {
    const query = new URL(String(input)).searchParams;
    const from = query.get('updatedFrom');
    const to = query.get('updatedTo');
    const eligible = rows.filter(row => (from === null || row.updateTimestamp >= Number(from))
      && (to === null || row.updateTimestamp <= Number(to)));
    const offset = Number(query.get('offset'));
    const limit = Math.min(pageMaximum, Number(query.get('limit')));
    const items = eligible.slice(offset, offset + limit);
    return Response.json({ offset, total: eligible.length, count: items.length, items });
  });
}

beforeEach(() => {
  vi.useFakeTimers();
  vi.setSystemTime(clock);
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand,last_ecwid_quantity)
    VALUES('nut','NUT','Nut','NUT','1001',100,100)`).run();
  const real = sqliteD1(sqlite);
  queries = 0;
  const db = new Proxy(real, { get(target, key) {
    if (key === 'prepare') return (sql: string) => { queries++; return target.prepare(sql); };
    const value = Reflect.get(target, key);
    return typeof value === 'function' ? value.bind(target) : value;
  } });
  env = { DB: db, ECWID_MODE: 'live', ECWID_STORE_ID: '123', ECWID_TOKEN: 'test', ORDER_SYNC_ENABLED: 'true', LIVE_SYNC_ENABLED: 'false',
    SYNC_QUEUE: { send: vi.fn().mockResolvedValue(undefined) } as unknown as Queue };
  setState('orders_tracking_started', new Date(clock - 600_000).toISOString());
});
afterEach(() => { sqlite.close(); vi.useRealTimers(); vi.restoreAllMocks(); });

describe('recent updated-order feed', () => {
  it('imports a new order before history, including an old-created order recently paid', async () => {
    const fetcher = catalogue([wire('NEW'), wire('OLD-RECENTLY-PAID')]);
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 2, complete: true });
    expect(state('orders_recent_watermark')).toBe(new Date(clock - 5_000).toISOString());
    expect(state('orders_recent_status')).toBe('CURRENT');
    expect(state('orders_last_full_poll')).toBeNull();
    expect(sqlite.prepare('SELECT count(*) AS n FROM orders').get()).toEqual({ n: 2 });
    expect(sqlite.prepare('SELECT reserved FROM item_stock').get()).toEqual({ reserved: 4 });
    for (const [input, init] of fetcher.mock.calls) {
      const query = new URL(String(input)).searchParams;
      expect(query.get('updatedFrom')).toBe(String(seconds - 600));
      expect(query.get('updatedTo')).toBe(String(seconds - 5));
      expect(query.has('createdTo')).toBe(false);
      expect(query.has('paymentStatus')).toBe(false);
      expect(query.has('fulfillmentStatus')).toBe(false);
      expect(init?.method).toBe('GET');
    }
    expect(fetcher).toHaveBeenCalledTimes(2);
  });

  it('imports payment changes and cancellations without changing shelf stock', async () => {
    await upsertOrderSnapshot(env.DB, parseOrder(wire('ORDER', seconds - 500, 'AWAITING_PAYMENT')));
    await pollOrders(env, catalogue([wire('ORDER', seconds - 30, 'PAID')]));
    expect(sqlite.prepare('SELECT payment_status FROM orders').get()).toEqual({ payment_status: 'PAID' });
    vi.setSystemTime(clock + 60_000);
    await pollOrders(env, catalogue([wire('ORDER', seconds + 30, 'CANCELLED')]));
    expect(sqlite.prepare('SELECT payment_status FROM orders').get()).toEqual({ payment_status: 'CANCELLED' });
    expect(sqlite.prepare('SELECT on_hand,reserved FROM item_stock').get()).toEqual({ on_hand: 100, reserved: 0 });
    expect(sqlite.prepare('SELECT count(*) AS n FROM movements').get()).toEqual({ n: 0 });
  });

  it.each(['SHIPPED', 'READY_FOR_PICKUP', 'DELIVERED'])('does not hide a recently %s old order without recorded picks', async fulfillment => {
    await pollOrders(env, catalogue([wire('OLD', seconds - 30, 'PAID', fulfillment)]));
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 1 });
    expect(sqlite.prepare("SELECT count(*) AS n FROM sync_issues WHERE status='OPEN'").get()).toEqual({ n: 1 });
  });

  it('retains fixed bounds through pages and only advances after complete verification', async () => {
    const fetcher = catalogue(Array.from({ length: 7 }, (_, i) => wire(`ORDER${i}`)));
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 3, complete: false });
    expect(checkpoint()).toMatchObject({ phase: 'APPLY', offset: 3, updatedTo: seconds - 5 });
    expect(state('orders_recent_watermark')).toBeNull();
    expect(state('orders_recent_status')).toBe('PENDING');
    vi.setSystemTime(clock + 60_000);
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 3, complete: false });
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 1, complete: true });
    expect(state('orders_recent_watermark')).toBe(new Date(clock - 5_000).toISOString());
    expect(fetcher.mock.calls.map(([input]) => new URL(String(input)).searchParams.get('offset'))).toEqual(['0', '3', '6', '0']);
    expect(new Set(fetcher.mock.calls.map(([input]) => new URL(String(input)).searchParams.get('updatedTo'))).size).toBe(1);
  });

  it('uses an overlap after completed windows without duplicating lines or reservations', async () => {
    const fetcher = catalogue([wire('ORDER')]);
    await pollOrders(env, fetcher);
    vi.setSystemTime(clock + 60_000);
    await pollOrders(env, fetcher);
    expect(new URL(String(fetcher.mock.calls[2][0])).searchParams.get('updatedFrom')).toBe(String(seconds - 125));
    expect(sqlite.prepare('SELECT count(*) AS n FROM order_lines').get()).toEqual({ n: 1 });
    expect(sqlite.prepare('SELECT reserved FROM item_stock').get()).toEqual({ reserved: 2 });
  });

  it('accepts legal short pages and resumes verification across invocations', async () => {
    const fetcher = catalogue([wire('A'), wire('B'), wire('C')], 1);
    for (let i = 0; i < 4; i++) expect((await pollOrders(env, fetcher)).complete).toBe(false);
    expect(checkpoint()).toMatchObject({ phase: 'VERIFY', offset: 2 });
    expect(state('orders_recent_watermark')).toBeNull();
    expect((await pollOrders(env, fetcher)).complete).toBe(true);
  });

  it('retains the checkpoint and watermark after a page failure then safely resumes', async () => {
    const fetcher = catalogue(Array.from({ length: 4 }, (_, i) => wire(String(i))));
    await pollOrders(env, fetcher);
    const previous = state('orders_recent_poll_cursor');
    await expect(pollOrders(env, async () => new Response('', { status: 503 }))).rejects.toThrow();
    expect(state('orders_recent_poll_cursor')).toBe(previous);
    expect(state('orders_recent_watermark')).toBeNull();
    expect(state('orders_recent_status')).toBe('ERROR');
    expect((await pollOrders(env, fetcher)).complete).toBe(true);
    expect(state('orders_recent_error')).toBe('');
    expect(sqlite.prepare('SELECT count(*) AS n FROM orders').get()).toEqual({ n: 4 });
  });

  it('restarts the same window when membership changes across pages', async () => {
    const original = [wire('A'), wire('B'), wire('C'), wire('D')];
    await pollOrders(env, catalogue(original));
    await expect(pollOrders(env, catalogue(original.slice(1)))).rejects.toThrow(/pages changed/);
    expect(checkpoint()).toMatchObject({ offset: 0, phase: 'APPLY', updatedTo: seconds - 5, seen: [] });
    expect(state('orders_recent_watermark')).toBeNull();
    expect((await pollOrders(env, catalogue(original.slice(1)))).complete).toBe(true);
  });

  it('detects same-size shifted pages in its second pass instead of silently losing an order', async () => {
    const rows = [wire('A'), wire('B'), wire('C'), wire('D')];
    await pollOrders(env, catalogue(rows));
    const moved = [wire('E'), wire('B'), wire('C'), wire('D')];
    await expect(pollOrders(env, catalogue(moved))).rejects.toThrow(/verification changed/);
    expect(state('orders_recent_watermark')).toBeNull();
    expect(checkpoint().offset).toBe(0);
    await pollOrders(env, catalogue(moved));
    expect((await pollOrders(env, catalogue(moved))).complete).toBe(true);
    expect(sqlite.prepare("SELECT id FROM orders WHERE id='E'").get()).toEqual({ id: 'E' });
  });

  it('detects status changes even with unchanged timestamps and line hashes', async () => {
    let call = 0;
    const fetcher: EcwidFetch = async input => catalogue([wire('A', seconds - 30, ++call === 1 ? 'PAID' : 'CANCELLED')])(input);
    await expect(pollOrders(env, fetcher)).rejects.toThrow(/verification changed/);
    expect(state('orders_recent_watermark')).toBeNull();
  });

  it('replays a partly applied page safely when a later snapshot transaction fails', async () => {
    sqlite.exec(`CREATE TRIGGER reject_second_test_order BEFORE INSERT ON orders WHEN NEW.id='B'
      BEGIN SELECT RAISE(ABORT,'synthetic database failure'); END`);
    const fetcher = catalogue([wire('A'), wire('B'), wire('C')]);
    await expect(pollOrders(env, fetcher)).rejects.toThrow(/synthetic database failure/);
    expect(checkpoint()).toMatchObject({ offset: 0, phase: 'APPLY', seen: [] });
    expect(state('orders_recent_watermark')).toBeNull();
    expect(sqlite.prepare('SELECT count(*) AS n FROM orders').get()).toEqual({ n: 1 });
    sqlite.exec('DROP TRIGGER reject_second_test_order');
    expect((await pollOrders(env, fetcher)).complete).toBe(true);
    expect(sqlite.prepare('SELECT count(*) AS n FROM order_lines').get()).toEqual({ n: 3 });
    expect(sqlite.prepare('SELECT reserved FROM item_stock').get()).toEqual({ reserved: 6 });
  });

  it('retains the prior watermark on failure in a subsequent overlapped window', async () => {
    await pollOrders(env, catalogue([wire('A')]));
    const previous = state('orders_recent_watermark');
    vi.setSystemTime(clock + 60_000);
    await expect(pollOrders(env, async () => new Response('', { status: 503 }))).rejects.toThrow();
    expect(state('orders_recent_watermark')).toBe(previous);
    expect(checkpoint().updatedFrom).toBe(seconds - 125);
  });

  it('quarantines a changed order line without overwriting its original commitment', async () => {
    await upsertOrderSnapshot(env.DB, parseOrder(wire('EDITED', seconds - 500)));
    const changed = wire('EDITED');
    changed.items[0].quantity = 7;
    await pollOrders(env, catalogue([changed]));
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 1 });
    expect(sqlite.prepare('SELECT ordered_qty FROM order_lines').get()).toEqual({ ordered_qty: 2 });
  });

  it.each([
    { total: 4, offset: 0, count: 0, items: [] },
    { total: 1, offset: 1, count: 1, items: [wire('A')] },
    { total: 2, offset: 0, count: 2, items: [wire('A'), wire('A')] },
    { total: 1, offset: 0, count: 1, items: [wire('OUTSIDE', seconds + 1)] },
  ])('rejects malformed or incomplete windows without advancement (%j)', async page => {
    await expect(pollOrders(env, async () => Response.json(page))).rejects.toThrow();
    expect(state('orders_recent_watermark')).toBeNull();
    expect(state('orders_recent_status')).toBe('ERROR');
  });

  it('makes overload explicit and retains the original window', async () => {
    await expect(pollOrders(env, async () => Response.json({ total: 1001, offset: 0, count: 1, items: [wire('A')] }))).rejects.toThrow(/exceeds/);
    expect(state('orders_recent_status')).toBe('OVERLOADED');
    expect(state('orders_recent_watermark')).toBeNull();
    expect(checkpoint().updatedTo).toBe(seconds - 5);
  });

  it('fails closed without the approved tracking baseline', async () => {
    sqlite.prepare("DELETE FROM sync_state WHERE key='orders_tracking_started'").run();
    const fetcher = catalogue([]);
    await expect(pollOrders(env, fetcher)).rejects.toThrow(/baseline/);
    expect(fetcher).not.toHaveBeenCalled();
    expect(state('orders_tracking_started')).toBeNull();
    expect(state('orders_recent_watermark')).toBeNull();
  });

  it('does not advance a corrupt checkpoint', async () => {
    setState('orders_recent_poll_cursor', '{bad');
    const fetcher = catalogue([]);
    await expect(pollOrders(env, fetcher)).rejects.toThrow(/checkpoint/);
    expect(fetcher).not.toHaveBeenCalled();
    expect(state('orders_recent_watermark')).toBeNull();
  });

  it('serializes simultaneous invocations with an atomic lease', async () => {
    let release!: () => void;
    const gate = new Promise<void>(resolve => { release = resolve; });
    const fetcher: EcwidFetch = async input => { await gate; return catalogue([wire('A')])(input); };
    const first = pollOrders(env, fetcher);
    while (!state('orders_poll_lease')) await Promise.resolve();
    expect(await pollOrders(env, catalogue([]))).toEqual({ processed: 0, complete: false, busy: true });
    release();
    expect((await first).complete).toBe(true);
    expect(state('orders_poll_lease')).toBeNull();
  });

  it('fences an expired poller from overwriting a replacement checkpoint or watermark', async () => {
    const replacement = JSON.stringify({ token: 'replacement', expiresAt: clock + 240_000 });
    const fetcher: EcwidFetch = async input => {
      setState('orders_poll_lease', replacement);
      setState('orders_recent_status', 'PENDING');
      return catalogue([wire('A')])(input);
    };
    await expect(pollOrders(env, fetcher)).rejects.toThrow(/lease expired/);
    expect(state('orders_recent_watermark')).toBeNull();
    expect(state('orders_recent_status')).toBe('PENDING');
    expect(state('orders_poll_lease')).toBe(replacement);
  });

  it('resumes the saved window after a crashed lease expires', async () => {
    const fetcher = catalogue([wire('A'), wire('B'), wire('C'), wire('D')]);
    await pollOrders(env, fetcher);
    setState('orders_poll_lease', JSON.stringify({ token: 'crashed', expiresAt: clock - 1 }));
    expect((await pollOrders(env, fetcher)).complete).toBe(true);
    expect(state('orders_recent_watermark')).toBe(new Date(clock - 5_000).toISOString());
  });

  it('cannot advance after its lease time expires even if nobody has replaced it', async () => {
    const fetcher: EcwidFetch = async input => {
      vi.setSystemTime(clock + 121_000);
      return catalogue([wire('A')])(input);
    };
    await expect(pollOrders(env, fetcher)).rejects.toThrow(/lease expired/);
    expect(state('orders_recent_watermark')).toBeNull();
    expect(state('orders_recent_status')).toBe('PENDING');
  });

  it('waits for the settle interval immediately after activation', async () => {
    setState('orders_tracking_started', new Date(clock).toISOString());
    const fetcher = catalogue([]);
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 0, complete: false });
    expect(fetcher).not.toHaveBeenCalled();
    expect(state('orders_recent_status')).toBe('PENDING');
    expect(state('orders_recent_watermark')).toBeNull();
  });

  it('runs bounded historical reconciliation only after a zero-application verified recent window', async () => {
    const fetcher = catalogue([wire('ANCIENT-UNPAID', seconds - 86400, 'AWAITING_PAYMENT')]);
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 0, complete: true, historical_processed: 1 });
    expect(state('orders_last_full_poll')).not.toBeNull();
    expect(sqlite.prepare('SELECT id FROM orders').get()).toEqual({ id: 'ANCIENT-UNPAID' });
    expect(new URL(String(fetcher.mock.calls.at(-1)![0])).searchParams.get('limit')).toBe('2');
  });

  it('does not invalidate a verified recent window if historical reconciliation fails', async () => {
    const fetcher: EcwidFetch = async input => new URL(String(input)).searchParams.has('createdTo')
      ? new Response('', { status: 503 }) : Response.json({ total: 0, count: 0, offset: 0, items: [] });
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 0, complete: true, historical_error: true });
    expect(state('orders_recent_status')).toBe('CURRENT');
    expect(state('orders_last_full_poll')).toBeNull();
  });

  it.each([true, false])('stays below 50 D1 statements including scheduled pumpSync (recent=%s)', async recent => {
    const rows = Array.from({ length: recent ? 3 : 10 }, (_, i) => wire(String(i), recent ? seconds - 30 : seconds - 86400, 'AWAITING_PAYMENT'));
    await Promise.all([pollOrders(env, catalogue(rows)), pumpSync(env)]);
    expect(queries).toBeLessThanOrEqual(40);
  });

  it.each([true, false])('bounds combined D1, Ecwid and queue subrequests below 50 (recent=%s)', async recent => {
    for (let i = 0; i < 5; i++) {
      sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand,last_ecwid_quantity)
        VALUES(?,?,?,?,?,10,10)`).run(`extra${i}`, `EXTRA${i}`, 'Extra', `BIN${i}`, String(2000 + i));
      await createMovement(env.DB, { operation_id: crypto.randomUUID(), type: 'INTERNAL_USE', item_id: `extra${i}`, quantity: 1 }, 'test');
      sqlite.prepare(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,received_at,updated_at)
        VALUES(?,'order.updated','A','123','{}',?,?)`).run(String(i), new Date().toISOString(), new Date().toISOString());
    }
    queries = 0;
    const rows = Array.from({ length: 3 }, (_, i) => wire(String(i), recent ? seconds - 30 : seconds - 86400, 'AWAITING_PAYMENT'));
    const fetcher = catalogue(rows, recent ? 100 : 1);
    await Promise.all([pollOrders(env, fetcher), pumpSync(env)]);
    expect(env.SYNC_QUEUE.send).toHaveBeenCalledTimes(10);
    expect(queries + fetcher.mock.calls.length + vi.mocked(env.SYNC_QUEUE.send).mock.calls.length).toBeLessThanOrEqual(50);
  });

  it('caps each queue category to five durable hints per scheduled invocation', async () => {
    for (let i = 0; i < 30; i++) sqlite.prepare(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,received_at,updated_at)
      VALUES(?,'order.updated','A','123','{}',?,?)`).run(String(i), new Date().toISOString(), new Date().toISOString());
    await pumpSync(env);
    expect(env.SYNC_QUEUE.send).toHaveBeenCalledTimes(5);
  });

  it.each([undefined, 'false', 'TRUE'])('does not fetch while order-sync flag is %s', async flag => {
    env.ORDER_SYNC_ENABLED = flag;
    const fetcher = catalogue([]);
    await expect(pollOrders(env, fetcher)).rejects.toMatchObject({ code: 'ORDER_SYNC_PAUSED' });
    expect(fetcher).not.toHaveBeenCalled();
    expect(queries).toBe(0);
  });
});

describe('small Ecwid search pages', () => {
  it.each([0, 101, -1, 1.1, NaN])('rejects invalid limit %s before fetching', async limit => {
    const fetcher = catalogue([]);
    await expect(new EcwidClient({ storeId: '123', token: 'test' }, fetcher).listOrders({ limit })).rejects.toThrow(/pagination/);
    expect(fetcher).not.toHaveBeenCalled();
  });
});
