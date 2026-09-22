import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EcwidClient, parseOrder, parseWebhook, verifyWebhookSignature, type EcwidOrder } from '../src/ecwid';
import { claimOutbox, ingestWebhook, pollOrders, processOutbox, recoverStaleProcessing, upsertOrderSnapshot, type SyncEnv } from '../src/sync';
import { createMovement } from '../src/inventory';
import { applyMigrations } from './d1';

function d1Adapter(sqlite: Database.Database): D1Database {
  class Statement {
    constructor(readonly sql: string, readonly values: unknown[] = []) {}
    bind(...values: unknown[]) { return new Statement(this.sql, values); }
    execute() {
      const statement = sqlite.prepare(this.sql);
      if (statement.reader) return { success: true, results: statement.all(...this.values), meta: { changes: 0 } };
      const result = statement.run(...this.values);
      return { success: true, results: [], meta: { changes: result.changes } };
    }
    async run() { return this.execute(); }
    async all() { return this.execute(); }
    async first(column?: string) {
      const row = sqlite.prepare(this.sql).get(...this.values) as Record<string, unknown> | undefined;
      return row ? (column ? row[column] : row) : null;
    }
  }
  return {
    prepare: (sql: string) => new Statement(sql),
    batch: async (statements: Statement[]) => sqlite.transaction(() => statements.map(statement => statement.execute()))(),
  } as unknown as D1Database;
}

let sqlite: Database.Database;
let db: D1Database;
let env: SyncEnv;
beforeEach(() => {
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand,last_ecwid_quantity)
    VALUES('item-1','NUT-M8','M8 nut','BIN-M8','1001',20,20)`).run();
  db = d1Adapter(sqlite);
  env = { DB: db, ECWID_MODE: 'live', ECWID_STORE_ID: '123', ECWID_TOKEN: 'test-token',
    ECWID_CLIENT_SECRET: 'test-client-secret', LIVE_SYNC_ENABLED: 'true', ORDER_SYNC_ENABLED: 'true',
    SYNC_QUEUE: { send: vi.fn().mockResolvedValue(undefined) } as unknown as Queue };
});
afterEach(() => { sqlite.close(); vi.restoreAllMocks(); });

function order(payment = 'AWAITING_PAYMENT', revision = 1): EcwidOrder {
  return { id: 'ORD1', paymentStatus: payment, fulfillmentStatus: 'AWAITING_PROCESSING',
    updatedAt: `2026-09-22T08:00:0${revision}.000Z`, items: [{ id: 'line1', productId: '1001', sku: 'NUT-M8',
      name: 'M8 nut', quantity: 3, combinationId: null, selectedOptions: [], digital: false, trackQuantity: true }] };
}

async function outgoing(quantity = 1) {
  const id = crypto.randomUUID();
  await createMovement(db, { operation_id: id, type: 'INTERNAL_USE', item_id: 'item-1', quantity }, 'picker');
  return id;
}

async function signature(eventId: string, eventCreated: string, secret = 'test-client-secret') {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const result = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(`${eventCreated}.${eventId}`));
  return btoa(String.fromCharCode(...new Uint8Array(result)));
}

describe('Ecwid wire contract', () => {
  it('sends a signed stock delta, never a stale absolute quantity', async () => {
    const fetcher = vi.fn().mockResolvedValue(Response.json({ updateCount: 1 }));
    await new EcwidClient({ storeId: '123', token: 'token' }, fetcher).adjustStock('1001', -2);
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(fetcher.mock.calls[0][0]).toBe('https://app.ecwid.com/api/v3/123/products/1001/inventory');
    expect(fetcher.mock.calls[0][1]).toMatchObject({ method: 'PUT', body: '{"quantityDelta":-2}', redirect: 'error' });
  });

  it.each([500, 502, 503, 408])('classifies HTTP %s after a write as uncertain', async status => {
    const fetcher = vi.fn().mockResolvedValue(new Response('failed', { status }));
    await expect(new EcwidClient({ storeId: '123', token: 'token' }, fetcher).adjustStock('1001', 1))
      .rejects.toMatchObject({ outcome: 'UNKNOWN' });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('does not retry a timed-out stock write', async () => {
    const fetcher = vi.fn((_input, init?: RequestInit) => new Promise<Response>((_resolve, reject) => {
      init?.signal?.addEventListener('abort', () => reject(new Error('timeout')), { once: true });
    }));
    await expect(new EcwidClient({ storeId: '123', token: 'token' }, fetcher, 5).adjustStock('1001', 1))
      .rejects.toMatchObject({ outcome: 'UNKNOWN' });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it.each([null, {}, { updateCount: 7 }, 'not json'])('requires a valid positive write confirmation (%s)', async body => {
    const fetcher = vi.fn().mockResolvedValue(typeof body === 'string' ? new Response(body) : Response.json(body));
    await expect(new EcwidClient({ storeId: '123', token: 'token' }, fetcher).adjustStock('1001', 1))
      .rejects.toMatchObject({ outcome: 'UNKNOWN' });
  });

  it('recognizes documented ignored 429 requests as safe to retry later', async () => {
    const fetcher = vi.fn().mockResolvedValue(new Response('', { status: 429, headers: { 'Retry-After': '120' } }));
    await expect(new EcwidClient({ storeId: '123', token: 'token' }, fetcher).adjustStock('1001', 1))
      .rejects.toMatchObject({ outcome: 'RETRYABLE', retryAfter: 120 });
  });

  it('rejects malformed quantities and duplicate order lines', () => {
    const raw = { id: 'A', paymentStatus: 'PAID', fulfillmentStatus: 'AWAITING_PROCESSING', updateTimestamp: 1,
      items: [{ id: 1, quantity: 0.5 }] };
    expect(() => parseOrder(raw)).toThrow('quantity');
    expect(() => parseOrder({ ...raw, items: [{ id: 1, quantity: 1 }, { id: 1, quantity: 1 }] })).toThrow('Duplicate');
  });
});

describe('durable outbound delivery', () => {
  it('claims an outbox row once across concurrent and duplicate messages', async () => {
    const id = await outgoing();
    const results = await Promise.all([claimOutbox(db, id), claimOutbox(db, id)]);
    expect(results.filter(Boolean)).toHaveLength(1);
    expect(await claimOutbox(db, id)).toBeNull();
  });

  it('applies a demo stock delta exactly once', async () => {
    const id = await outgoing(2);
    env.ECWID_MODE = 'demo';
    const fetcher = vi.fn();
    await Promise.all([processOutbox(env, id, fetcher), processOutbox(env, id, fetcher)]);
    await processOutbox(env, id, fetcher);
    expect(fetcher).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT last_ecwid_quantity FROM items').get()).toEqual({ last_ecwid_quantity: 18 });
    expect(sqlite.prepare('SELECT status,attempts FROM outbox').get()).toEqual({ status: 'APPLIED', attempts: 1 });
  });

  it('blocks newer item adjustments after an ambiguous write', async () => {
    const first = await outgoing(1);
    const second = await outgoing(2);
    sqlite.prepare('UPDATE outbox SET created_at=? WHERE id=?').run('2026-09-22T00:00:00.000Z', first);
    const fetcher = vi.fn().mockRejectedValue(new Error('connection lost after send'));
    await processOutbox(env, first, fetcher);
    await processOutbox(env, first, fetcher);
    await processOutbox(env, second, fetcher);
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(sqlite.prepare('SELECT status FROM outbox WHERE id=?').get(first)).toEqual({ status: 'UNKNOWN' });
    expect(sqlite.prepare('SELECT status,attempts FROM outbox WHERE id=?').get(second)).toEqual({ status: 'PENDING', attempts: 0 });
    expect(sqlite.prepare("SELECT count(*) AS count FROM sync_issues WHERE status='OPEN'").get()).toEqual({ count: 1 });
  });

  it('recovers a crashed in-flight stock adjustment into manual review, never pending', async () => {
    const id = await outgoing();
    await claimOutbox(db, id);
    sqlite.prepare("UPDATE outbox SET updated_at='2020-01-01T00:00:00.000Z'").run();
    await recoverStaleProcessing(db);
    expect(sqlite.prepare('SELECT status FROM outbox').get()).toEqual({ status: 'UNKNOWN' });
    expect(await claimOutbox(db, id)).toBeNull();
  });

  it('does not resend a confirmed write when the stock read-back fails', async () => {
    const id = await outgoing();
    const fetcher = vi.fn().mockResolvedValueOnce(Response.json({ updateCount: 1 })).mockRejectedValue(new Error('read failed'));
    await processOutbox(env, id, fetcher);
    await processOutbox(env, id, fetcher);
    expect(fetcher).toHaveBeenCalledTimes(2);
    expect(sqlite.prepare('SELECT status FROM outbox').get()).toEqual({ status: 'APPLIED' });
    expect(sqlite.prepare('SELECT last_ecwid_quantity FROM items').get()).toEqual({ last_ecwid_quantity: null });
  });

  it('leaves stock queued when live writes are disabled', async () => {
    const id = await outgoing();
    env.LIVE_SYNC_ENABLED = 'false';
    const fetcher = vi.fn();
    await processOutbox(env, id, fetcher);
    expect(fetcher).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT status,attempts FROM outbox').get()).toEqual({ status: 'PENDING', attempts: 0 });
  });
});

describe('order snapshot reconciliation', () => {
  it('does not mistake low-stock notification settings for untracked inventory', async () => {
    const simple = order('PAID');
    simple.items[0].trackQuantity = false;
    await upsertOrderSnapshot(db, simple);
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 0 });
    expect(sqlite.prepare('SELECT item_id FROM order_lines').get()).toEqual({ item_id: 'item-1' });
  });

  it('reserves awaiting-payment units, changes to paid without another reservation, and preserves partial picks', async () => {
    await upsertOrderSnapshot(db, order());
    expect(sqlite.prepare('SELECT on_hand,reserved FROM item_stock').get()).toEqual({ on_hand: 20, reserved: 3 });
    await upsertOrderSnapshot(db, order('PAID', 2));
    await createMovement(db, { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'item-1', quantity: 1,
      order_id: 'ORD1', order_line_id: 'ORD1:line1' }, 'picker');
    await upsertOrderSnapshot(db, order('PAID', 3));
    await upsertOrderSnapshot(db, order('AWAITING_PAYMENT', 1));
    expect(sqlite.prepare('SELECT on_hand,reserved FROM item_stock').get()).toEqual({ on_hand: 19, reserved: 2 });
    expect(sqlite.prepare('SELECT payment_status,needs_review FROM orders').get()).toEqual({ payment_status: 'PAID', needs_review: 0 });
    expect(sqlite.prepare('SELECT count(*) AS count FROM outbox').get()).toEqual({ count: 0 });
  });

  it('releases a cancellation before picking without changing physical stock', async () => {
    await upsertOrderSnapshot(db, order());
    await upsertOrderSnapshot(db, order('CANCELLED', 2));
    expect(sqlite.prepare('SELECT on_hand,reserved FROM item_stock').get()).toEqual({ on_hand: 20, reserved: 0 });
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 0 });
  });

  it('flags cancellation after a partial pick and blocks the affected stock', async () => {
    await upsertOrderSnapshot(db, order('PAID'));
    await createMovement(db, { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'item-1', quantity: 1,
      order_id: 'ORD1', order_line_id: 'ORD1:line1' }, 'picker');
    await upsertOrderSnapshot(db, order('CANCELLED', 2));
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 1 });
    expect(sqlite.prepare('SELECT picked_qty FROM order_lines').get()).toEqual({ picked_qty: 1 });
    await expect(outgoing()).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
  });

  it('flags changed lines without overwriting their audited quantity', async () => {
    await upsertOrderSnapshot(db, order('PAID'));
    const edited = order('PAID', 2);
    edited.items[0].quantity = 5;
    await upsertOrderSnapshot(db, edited);
    expect(sqlite.prepare('SELECT ordered_qty FROM order_lines').get()).toEqual({ ordered_qty: 3 });
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 1 });
  });

  it('quarantines an unmapped variation without blocking an independent parent target', async () => {
    const variation = order('PAID');
    variation.items[0].combinationId = '456';
    variation.items[0].sku = 'VARIANT-M8';
    await upsertOrderSnapshot(db, variation);
    expect(sqlite.prepare('SELECT item_id FROM order_lines').get()).toEqual({ item_id: null });
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 1 });
    await expect(outgoing()).resolves.toEqual(expect.any(String));
  });

  it('flags conflicting statuses sharing one remote timestamp rather than unblocking picking', async () => {
    await upsertOrderSnapshot(db, order());
    await upsertOrderSnapshot(db, order('PAID'));
    expect(sqlite.prepare('SELECT payment_status,needs_review FROM orders').get()).toEqual({ payment_status: 'AWAITING_PAYMENT', needs_review: 1 });
  });

  it.each(['SHIPPED', 'READY_FOR_PICKUP'])('does not import historical %s orders as physical reservations', async fulfillmentStatus => {
    const historic = { ...order('PAID'), fulfillmentStatus };
    expect(await upsertOrderSnapshot(db, historic, { skipUntrackedTerminal: true })).toMatchObject({ skipped: true });
    expect(sqlite.prepare('SELECT count(*) AS count FROM orders').get()).toEqual({ count: 0 });
    await upsertOrderSnapshot(db, historic);
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 1 });
  });

  it.each(['SHIPPED', 'READY_FOR_PICKUP'])('quarantines tracked %s orders with missing picks without changing stock', async fulfillmentStatus => {
    await upsertOrderSnapshot(db, order('PAID'));
    await createMovement(db, { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'item-1', quantity: 1,
      order_id: 'ORD1', order_line_id: 'ORD1:line1' }, 'picker');
    const result = await upsertOrderSnapshot(db, { ...order('PAID', 2), fulfillmentStatus }, { skipUntrackedTerminal: true });
    expect(result).toEqual({ id: 'ORD1', needs_review: true });
    expect(sqlite.prepare('SELECT fulfillment_status FROM orders').get()).toEqual({ fulfillment_status: fulfillmentStatus });
    expect(sqlite.prepare('SELECT on_hand,reserved,available,last_ecwid_quantity FROM item_stock').get())
      .toEqual({ on_hand: 19, reserved: 2, available: 17, last_ecwid_quantity: 20 });
    expect(sqlite.prepare('SELECT picked_qty FROM order_lines').get()).toEqual({ picked_qty: 1 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 1 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM outbox').get()).toEqual({ count: 0 });
    await expect(outgoing()).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
  });

  it.each(['SHIPPED', 'READY_FOR_PICKUP'])('accepts %s after all picks without double-deducting or flagging delivery', async fulfillmentStatus => {
    await upsertOrderSnapshot(db, order('PAID'));
    await createMovement(db, { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'item-1', quantity: 3,
      order_id: 'ORD1', order_line_id: 'ORD1:line1' }, 'picker');
    expect(await upsertOrderSnapshot(db, { ...order('PAID', 2), fulfillmentStatus })).toEqual({ id: 'ORD1', needs_review: false });
    expect(sqlite.prepare('SELECT fulfillment_status FROM orders').get()).toEqual({ fulfillment_status: fulfillmentStatus });
    expect(sqlite.prepare('SELECT on_hand,reserved,last_ecwid_quantity FROM item_stock').get())
      .toEqual({ on_hand: 17, reserved: 0, last_ecwid_quantity: 20 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 1 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM outbox').get()).toEqual({ count: 0 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM sync_issues').get()).toEqual({ count: 0 });
  });
});

describe('webhook authentication and durable ingestion', () => {
  it('verifies the documented eventCreated.eventId HMAC and rejects changes', async () => {
    const event = parseWebhook({ eventId: 'event-1', eventCreated: 1720000000, storeId: 123,
      eventType: 'order.updated', entityId: 999, data: { orderId: 'ORD1' } });
    const signed = await signature('event-1', '1720000000');
    expect(event.entityId).toBe('ORD1');
    expect(await verifyWebhookSignature(event, signed, 'test-client-secret')).toBe(true);
    expect(await verifyWebhookSignature({ ...event, eventId: 'changed' }, signed, 'test-client-secret')).toBe(false);
    expect(await verifyWebhookSignature(event, null, 'test-client-secret')).toBe(false);
    expect(await verifyWebhookSignature(event, signed, 'wrong')).toBe(false);
  });

  it('stores a webhook exactly once and acknowledges even if queue delivery is unavailable', async () => {
    env.SYNC_QUEUE.send = vi.fn().mockRejectedValue(new Error('queue offline'));
    const body = { eventId: 'event-1', eventCreated: 1720000000, storeId: 123,
      eventType: 'order.updated', entityId: 999, data: { orderId: 'ORD1' } };
    const signed = await signature('event-1', '1720000000');
    const request = () => new Request('https://inventory.example/api/webhooks/ecwid', { method: 'POST',
      body: JSON.stringify(body), headers: { 'X-Ecwid-Webhook-Signature': signed } });
    expect((await ingestWebhook(request(), env)).status).toBe(202);
    expect((await ingestWebhook(request(), env)).status).toBe(202);
    expect(sqlite.prepare('SELECT event_id,status FROM webhook_events').all()).toEqual([{ event_id: 'event-1', status: 'PENDING' }]);
  });

  it('rejects a signed webhook from another store without storing it', async () => {
    const body = { eventId: 'event-1', eventCreated: 1720000000, storeId: 999, eventType: 'product.updated', entityId: 1001 };
    const request = new Request('https://inventory.example/api/webhooks/ecwid', { method: 'POST', body: JSON.stringify(body),
      headers: { 'X-Ecwid-Webhook-Signature': await signature('event-1', '1720000000') } });
    expect((await ingestWebhook(request, env)).status).toBe(401);
    expect(sqlite.prepare('SELECT count(*) AS count FROM webhook_events').get()).toEqual({ count: 0 });
  });
});

describe('complete order polling', () => {
  it('persists pagination rather than repeatedly inspecting only the first page', async () => {
    // One result per page is legal even when limit=100. This makes the checkpoint visible.
    const fetcher = vi.fn((input: string | URL | Request) => {
      const offset = Number(new URL(String(input)).searchParams.get('offset'));
      return Promise.resolve(Response.json({ total: 3, count: 1, offset, items: [{ id: `ORD${offset}`,
        paymentStatus: 'AWAITING_PAYMENT', fulfillmentStatus: 'AWAITING_PROCESSING', createTimestamp: 1, updateTimestamp: 2,
        items: [{ id: 1, productId: 1001, sku: 'NUT-M8', name: 'Nut', quantity: 1, trackQuantity: true }] }] }));
    });
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 2, complete: false });
    expect(await pollOrders(env, fetcher)).toEqual({ processed: 1, complete: true });
    expect(fetcher.mock.calls.map(call => new URL(String(call[0])).searchParams.get('offset'))).toEqual(['0', '1', '2']);
    expect(sqlite.prepare('SELECT reserved FROM item_stock').get()).toEqual({ reserved: 3 });
  });
});
