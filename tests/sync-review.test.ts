import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createMovement } from '../src/inventory';
import { pollOrders, processOutbox, processWebhook, pumpSync, upsertOrderSnapshot, type SyncEnv } from '../src/sync';
import type { EcwidOrder } from '../src/ecwid';
import { applyMigrations } from './d1';

function adapter(sqlite: Database.Database): D1Database {
  class Statement {
    constructor(readonly sql: string, readonly values: unknown[] = []) {}
    bind(...values: unknown[]) { return new Statement(this.sql, values); }
    execute() {
      const statement = sqlite.prepare(this.sql);
      if (statement.reader) return { success: true, results: statement.all(...this.values), meta: { changes: 0 } };
      return { success: true, results: [], meta: { changes: statement.run(...this.values).changes } };
    }
    async all() { return this.execute(); }
    async run() { return this.execute(); }
    async first(column?: string) {
      const row = sqlite.prepare(this.sql).get(...this.values) as Record<string, unknown> | undefined;
      return row ? (column ? row[column] : row) : null;
    }
  }
  // Test-only adapter implements the D1 subset used by this app. Real SQLite
  // executes every statement and transaction; unrelated D1 methods are unused.
  return {
    prepare: (sql: string) => new Statement(sql),
    batch: async (statements: Statement[]) => sqlite.transaction(() => statements.map(statement => statement.execute()))(),
  } as unknown as D1Database;
}

let sqlite: Database.Database;
let db: D1Database;
let env: SyncEnv;
const trackingStart = '2026-09-22T08:00:00.000Z';

beforeEach(() => {
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand,last_ecwid_quantity)
    VALUES('item-1','NUT','Nut','BIN-NUT','1001',100,100),('item-2','BOLT','Bolt','BIN-BOLT','1002',100,100)`).run();
  db = adapter(sqlite);
  env = { DB: db, ECWID_MODE: 'live', ECWID_STORE_ID: '123', ECWID_TOKEN: 'test-token',
    ECWID_CLIENT_SECRET: 'test-secret', LIVE_SYNC_ENABLED: 'true',
    SYNC_QUEUE: { send: vi.fn().mockResolvedValue(undefined) } as unknown as Queue };
});
afterEach(() => { sqlite.close(); vi.restoreAllMocks(); });

async function outgoing(item = 'item-1') {
  const id = crypto.randomUUID();
  await createMovement(db, { operation_id: id, type: 'INTERNAL_USE', item_id: item, quantity: 1 }, 'picker');
  return id;
}

function snapshot(paymentStatus: string, updatedAt: string): EcwidOrder {
  return { id: 'ORDER', paymentStatus, fulfillmentStatus: 'AWAITING_PROCESSING', updatedAt,
    items: [{ id: 'line', productId: '1001', sku: 'NUT', name: 'Nut', quantity: 3,
      combinationId: null, selectedOptions: [], digital: false, trackQuantity: false }] };
}

describe('independent integration regressions', () => {
  it('keeps the newer paid snapshot when first imports race in either order', async () => {
    await Promise.all([
      upsertOrderSnapshot(db, snapshot('PAID', '2026-09-22T08:00:02.000Z')),
      upsertOrderSnapshot(db, snapshot('AWAITING_PAYMENT', '2026-09-22T08:00:01.000Z')),
    ]);
    expect(sqlite.prepare('SELECT payment_status,needs_review FROM orders').get()).toEqual({ payment_status: 'PAID', needs_review: 0 });
    expect(sqlite.prepare('SELECT count(*) AS count FROM order_lines').get()).toEqual({ count: 1 });
    expect(sqlite.prepare("SELECT reserved FROM item_stock WHERE id='item-1'").get()).toEqual({ reserved: 3 });
  });

  it('does not skip an old order fulfilled after stock tracking began', async () => {
    sqlite.prepare('INSERT INTO sync_state(key,value,updated_at) VALUES(?,?,?)')
      .run('orders_tracking_started', trackingStart, trackingStart);
    const start = Date.parse(trackingStart) / 1000;
    const fetcher = vi.fn().mockResolvedValue(Response.json({ total: 1, count: 1, offset: 0, items: [{
      id: 'OLDER-ORDER', paymentStatus: 'PAID', fulfillmentStatus: 'SHIPPED',
      createTimestamp: start - 86400, updateTimestamp: start + 60,
      items: [{ id: 1, productId: 1001, sku: 'NUT', name: 'Nut', quantity: 2 }],
    }] }));
    await pollOrders(env, fetcher);
    expect(sqlite.prepare('SELECT id,needs_review FROM orders').get()).toEqual({ id: 'OLDER-ORDER', needs_review: 1 });
    await expect(outgoing()).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
  });

  it('delivers another item when the oldest 25 pending rows are blocked', async () => {
    for (let i = 0; i < 25; i++) await outgoing();
    sqlite.prepare(`INSERT INTO sync_issues(id,item_id,kind,message,created_at)
      VALUES('review','item-1','REVIEW','Count needed',?)`).run(trackingStart);
    const deliverable = await outgoing('item-2');
    await pumpSync(env);
    expect(env.SYNC_QUEUE.send).toHaveBeenCalledWith({ kind: 'outbox', id: deliverable });
    expect(env.SYNC_QUEUE.send).toHaveBeenCalledTimes(1);
  });

  it('flags a mapped product whose SKU identity changed', async () => {
    sqlite.prepare(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,received_at,updated_at)
      VALUES('product-event','product.updated','1001','123','{}',?,?)`).run(trackingStart, trackingStart);
    const fetcher = vi.fn().mockResolvedValue(Response.json({ id: 1001, sku: 'DIFFERENT-PART', name: 'Different part',
      quantity: 100, unlimited: false, enabled: true, options: [], combinations: [] }));
    await processWebhook(env, 'product-event', fetcher);
    expect(sqlite.prepare("SELECT item_id,status FROM sync_issues WHERE item_id='item-1'").get())
      .toEqual({ item_id: 'item-1', status: 'OPEN' });
    await expect(outgoing()).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
  });

  it('blocks a mapped product deleted before an update webhook is processed', async () => {
    sqlite.prepare(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,received_at,updated_at)
      VALUES('missing-product','product.updated','1001','123','{}',?,?)`).run(trackingStart, trackingStart);
    await processWebhook(env, 'missing-product', vi.fn().mockResolvedValue(new Response('', { status: 404 })));
    expect(sqlite.prepare("SELECT item_id,status FROM sync_issues WHERE item_id='item-1'").get())
      .toEqual({ item_id: 'item-1', status: 'OPEN' });
    await expect(outgoing()).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
  });

  it('treats a negative-stock warning as applied and blocks further movements', async () => {
    const id = await outgoing();
    const fetcher = vi.fn().mockResolvedValueOnce(Response.json({ updateCount: 1, warning: 'Product stock became negative' }))
      .mockResolvedValueOnce(Response.json({ id: 1001, sku: 'NUT', name: 'Nut', quantity: -1, unlimited: false, enabled: true }));
    await processOutbox(env, id, fetcher);
    await processOutbox(env, id, fetcher);
    expect(fetcher).toHaveBeenCalledTimes(2);
    expect(sqlite.prepare('SELECT status FROM outbox WHERE id=?').get(id)).toEqual({ status: 'APPLIED' });
    expect(sqlite.prepare("SELECT kind,status FROM sync_issues WHERE item_id='item-1'").get())
      .toEqual({ kind: 'NEGATIVE_ECWID_STOCK', status: 'OPEN' });
    await expect(outgoing()).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
  });
});
