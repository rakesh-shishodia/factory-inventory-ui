import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createMovement, dashboard, getItem, getOrder, listOrders } from '../src/inventory';
import { validateMovementInput } from '../src/domain';
import { applyMigrations } from './d1';

// The domain only needs D1's prepared statements and atomic batch. Real SQLite
// executes the migration and triggers, rather than mocking their stock effects.
function d1Adapter(sqlite: Database.Database): D1Database {
  class Statement {
    constructor(readonly sql: string, readonly values: unknown[] = []) {}
    bind(...values: unknown[]) { return new Statement(this.sql, values); }
    execute() {
      const stmt = sqlite.prepare(this.sql);
      if (stmt.reader) return { success: true, results: stmt.all(...this.values), meta: { changes: 0 } };
      const result = stmt.run(...this.values);
      return { success: true, results: [], meta: { changes: result.changes } };
    }
    async all() { return this.execute(); }
    async run() { return this.execute(); }
    async first(column?: string) {
      const row = sqlite.prepare(this.sql).get(...this.values) as Record<string, unknown> | undefined;
      return row ? (column ? row[column] : row) : null;
    }
  }
  return {
    prepare: (sql: string) => new Statement(sql),
    batch: async (statements: Statement[]) => sqlite.transaction(() => statements.map(stmt => stmt.execute()))(),
  } as unknown as D1Database;
}

let sqlite: Database.Database;
let db: D1Database;
const actor = 'picker@example.com';
const timestamp = '2026-09-22T08:00:00.000Z';

function movement(type: 'ECWID_PICK' | 'EMAIL_SALE' | 'INTERNAL_USE' | 'RESTOCK', quantity: number, extra = {}) {
  return { operation_id: crypto.randomUUID(), type, item_id: 'item-1', quantity, ...extra };
}

function addOrder(id: string, quantity: number, payment = 'PAID', fulfillment = 'AWAITING_PROCESSING') {
  sqlite.prepare(`INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
    VALUES(?,?,?,?,?)`).run(id, payment, fulfillment, timestamp, timestamp);
  sqlite.prepare(`INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
    VALUES(?,?,?,?,?,?,?)`).run(`line-${id}`, id, 'line-1', 'item-1', 'NUT-M8', 'M8 nut', quantity);
}

function pick(order: string, quantity: number) {
  return movement('ECWID_PICK', quantity, { order_id: order, order_line_id: `line-${order}` });
}

beforeEach(() => {
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
    VALUES('item-1','NUT-M8','M8 nut','BIN-M8','Shelf A','1001',10,10)`).run();
  db = d1Adapter(sqlite);
});

afterEach(() => sqlite.close());

describe('physical inventory and order reservations', () => {
  it('reserves awaiting-payment stock, then permits picking only after payment', async () => {
    addOrder('order-1', 3, 'AWAITING_PAYMENT');
    expect(await getItem(db, ' nut-m8 ')).toMatchObject({ on_hand: 10, reserved: 3, available: 7 });
    await expect(createMovement(db, pick('order-1', 2), actor)).rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE', status: 409 });
    expect(await listOrders(db)).toHaveLength(0);
    sqlite.prepare("UPDATE orders SET payment_status = 'PAID' WHERE id = 'order-1'").run();
    expect(await getItem(db, 'BIN-M8')).toMatchObject({ on_hand: 10, reserved: 3, available: 7 });
    expect(await listOrders(db)).toHaveLength(1);
    const result = await createMovement(db, pick('order-1', 2), actor);
    expect(result).toMatchObject({ duplicate: false, sync_status: 'NOT_REQUIRED' });
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 8, reserved: 1, available: 7, last_ecwid_quantity: 10 });
    expect((await getOrder(db, 'order-1')).lines[0]).toMatchObject({ ordered_qty: 3, picked_qty: 2, remaining_qty: 1 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM outbox').get()).toEqual({ count: 0 });
    await createMovement(db, pick('order-1', 1), actor);
    expect(await listOrders(db)).toHaveLength(0);
  });

  it('releases cancelled reservations without changing physical stock', async () => {
    addOrder('order-1', 3, 'AWAITING_PAYMENT');
    sqlite.prepare("UPDATE orders SET payment_status = 'CANCELLED' WHERE id = 'order-1'").run();
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 10, reserved: 0, available: 10 });
    await expect(createMovement(db, pick('order-1', 1), actor)).rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
  });

  it('prevents over-picking and picking a line belonging to another order', async () => {
    addOrder('order-1', 2);
    addOrder('order-2', 1);
    await expect(createMovement(db, pick('order-1', 3), actor)).rejects.toMatchObject({ code: 'PICK_QUANTITY_EXCEEDED' });
    await expect(createMovement(db, { ...pick('order-1', 1), order_line_id: 'line-order-2' }, actor))
      .rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 10, reserved: 3 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 0 });
  });

  it('rejects reviewed, shipped and picked-and-packed pickup orders even when paid', async () => {
    addOrder('review', 2);
    addOrder('shipped', 2, 'PAID', 'SHIPPED');
    addOrder('pickup', 2, 'PAID', 'READY_FOR_PICKUP');
    sqlite.prepare("UPDATE orders SET needs_review = 1 WHERE id = 'review'").run();
    for (const order of ['review', 'shipped', 'pickup']) {
      await expect(createMovement(db, pick(order, 1), actor)).rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
    }
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 10, reserved: 6 });
    expect(await listOrders(db)).toEqual([]);
    expect(await dashboard(db)).toMatchObject({ pickable_orders: 0 });
  });

  it('preserves reservations when recording email sales and internal use', async () => {
    addOrder('order-1', 8);
    await expect(createMovement(db, movement('EMAIL_SALE', 3), actor)).rejects.toMatchObject({ code: 'INSUFFICIENT_AVAILABLE_STOCK' });
    await createMovement(db, movement('INTERNAL_USE', 2), actor);
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 8, reserved: 8, available: 0 });
    const queued = sqlite.prepare('SELECT quantity_delta,status FROM outbox').all();
    expect(queued).toEqual([{ quantity_delta: -2, status: 'PENDING' }]);
  });

  it('records restocks and outgoing movements with signed Ecwid deltas', async () => {
    await createMovement(db, movement('RESTOCK', 5), actor);
    sqlite.prepare("UPDATE outbox SET status='APPLIED'").run();
    await createMovement(db, movement('EMAIL_SALE', 2), actor);
    sqlite.prepare("UPDATE outbox SET status='APPLIED'").run();
    await createMovement(db, movement('INTERNAL_USE', 1), actor);
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 12, available: 12 });
    expect(sqlite.prepare('SELECT quantity_delta FROM outbox ORDER BY rowid').all()).toEqual([
      { quantity_delta: 5 }, { quantity_delta: -2 }, { quantity_delta: -1 },
    ]);
    expect(await dashboard(db)).toMatchObject({ physical_units: 12, pending_sync: 1, attention_count: 0 });
  });

  it('waits for the previous Ecwid delta before accepting another for the same item', async () => {
    const first = movement('EMAIL_SALE', 1);
    await createMovement(db, first, actor);
    await expect(createMovement(db, movement('INTERNAL_USE', 1), actor))
      .rejects.toMatchObject({ code: 'STOCK_SYNC_PENDING', status: 409 });
    expect(await createMovement(db, first, actor)).toMatchObject({ duplicate: true, sync_status: 'PENDING' });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 1 });
    sqlite.prepare("UPDATE outbox SET status='APPLIED'").run();
    await expect(createMovement(db, movement('RESTOCK', 1), actor)).resolves.toMatchObject({ duplicate: false });
  });

  it('stores adjustment reasons explicitly and requires a note and matching direction', async () => {
    const up = { ...movement('RESTOCK', 2), reason_code: 'ADJUST_UP' as const, note: 'Adjustment up: Cycle count' };
    const result = await createMovement(db, up, actor);
    expect(result.movement).toMatchObject({ type: 'RESTOCK', reason_code: 'ADJUST_UP', quantity_delta: 2 });
    sqlite.prepare("UPDATE outbox SET status='APPLIED'").run();
    await expect(createMovement(db, { ...movement('INTERNAL_USE', 1), reason_code: 'ADJUST_DOWN', note: '' }, actor))
      .rejects.toMatchObject({ code: 'ADJUSTMENT_NOTE_REQUIRED', status: 400 });
    await expect(createMovement(db, { ...movement('RESTOCK', 1), reason_code: 'ADJUST_DOWN', note: 'Wrong way' }, actor))
      .rejects.toMatchObject({ code: 'INVALID_ADJUSTMENT_DIRECTION', status: 400 });
  });

  it('never lets two requests spend the same physical units', async () => {
    const outcomes = await Promise.allSettled([
      createMovement(db, movement('INTERNAL_USE', 7), actor),
      createMovement(db, movement('EMAIL_SALE', 7), actor),
    ]);
    expect(outcomes.filter(result => result.status === 'fulfilled')).toHaveLength(1);
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 3 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 1 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM outbox').get()).toEqual({ count: 1 });
  });

  it('never lets two requests pick the same final order quantity', async () => {
    addOrder('order-1', 2);
    const outcomes = await Promise.allSettled([
      createMovement(db, pick('order-1', 2), actor),
      createMovement(db, pick('order-1', 2), actor),
    ]);
    expect(outcomes.filter(result => result.status === 'fulfilled')).toHaveLength(1);
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 8, reserved: 0 });
  });
});

describe('durability and idempotency', () => {
  it('replays the exact UUID without a second balance or outbox change', async () => {
    const input = movement('EMAIL_SALE', 10);
    const results = await Promise.all([createMovement(db, input, actor), createMovement(db, input, actor)]);
    expect(results.map(result => result.duplicate).sort()).toEqual([false, true]);
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 0 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM outbox').get()).toEqual({ count: 1 });
    // Replaying remains safe after the old movement is pending operator review.
    sqlite.prepare("UPDATE outbox SET status = 'UNKNOWN'").run();
    expect(await createMovement(db, input, actor)).toMatchObject({ duplicate: true, sync_status: 'UNKNOWN' });
  });

  it('rejects reuse of an operation ID with a changed payload or actor', async () => {
    const input = movement('RESTOCK', 2);
    await createMovement(db, input, actor);
    await expect(createMovement(db, { ...input, quantity: 3 }, actor)).rejects.toMatchObject({ code: 'IDEMPOTENCY_CONFLICT', status: 409 });
    await expect(createMovement(db, input, 'different@example.com')).rejects.toMatchObject({ code: 'IDEMPOTENCY_CONFLICT' });
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 12 });
  });

  it('rolls back the movement and stock when creating its outbox entry fails', async () => {
    sqlite.exec("CREATE TRIGGER simulate_outbox_failure BEFORE INSERT ON outbox BEGIN SELECT RAISE(ABORT, 'outbox unavailable'); END;");
    await expect(createMovement(db, movement('EMAIL_SALE', 2), actor)).rejects.toThrow('outbox unavailable');
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 10 });
    expect(sqlite.prepare('SELECT COUNT(*) AS count FROM movements').get()).toEqual({ count: 0 });
  });

  it('blocks new movements when stock sync needs review', async () => {
    const input = movement('EMAIL_SALE', 1);
    await createMovement(db, input, actor);
    sqlite.prepare("UPDATE outbox SET status = 'UNKNOWN'").run();
    await expect(createMovement(db, movement('RESTOCK', 1), actor)).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
    sqlite.prepare("UPDATE outbox SET status = 'APPLIED'").run();
    sqlite.prepare(`INSERT INTO sync_issues(id,item_id,kind,message,created_at)
      VALUES('issue-1','item-1','DRIFT','Please check the stock',?)`).run(timestamp);
    await expect(createMovement(db, movement('INTERNAL_USE', 1), actor)).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
  });

  it('requires verified mapping for a movement that adjusts Ecwid', async () => {
    sqlite.prepare('UPDATE items SET ecwid_product_id = NULL').run();
    await expect(createMovement(db, movement('EMAIL_SALE', 1), actor)).rejects.toMatchObject({ code: 'ECWID_MAPPING_REQUIRED' });
  });

  it('keeps the movement audit immutable', async () => {
    await createMovement(db, movement('RESTOCK', 2), actor);
    expect(() => sqlite.prepare("UPDATE movements SET note = 'changed'").run()).toThrow('MOVEMENT_IMMUTABLE');
    expect(() => sqlite.prepare('DELETE FROM movements').run()).toThrow('MOVEMENT_IMMUTABLE');
  });

  it('records one immutable opening balance before operational movements', async () => {
    sqlite.prepare("UPDATE items SET on_hand = 0 WHERE id = 'item-1'").run();
    const insert = sqlite.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
      VALUES(?,'item-1',15,'stock-sheet:approved-run-1',?,?)`);
    insert.run('opening-1', actor, timestamp);
    expect(await getItem(db, 'item-1')).toMatchObject({ on_hand: 15 });
    expect(() => insert.run('opening-2', actor, timestamp)).toThrow();
    expect(() => sqlite.prepare('UPDATE opening_balances SET on_hand = 20').run()).toThrow('OPENING_BALANCE_IMMUTABLE');
    expect(() => sqlite.prepare('DELETE FROM opening_balances').run()).toThrow('OPENING_BALANCE_IMMUTABLE');
  });

  it.each([0, -1, 1.5, '2', null, 1000001])('rejects invalid quantity %s before writing', quantity => {
    expect(() => validateMovementInput({ ...movement('RESTOCK', 1), quantity })).toThrow('Quantity must be a whole number');
  });

  it('rejects lowercase, duplicate and missing stock identifiers in the database', () => {
    expect(() => sqlite.prepare("INSERT INTO items(id,sku,name,scan_code) VALUES('b','nut-m8','Nut','OTHER')").run()).toThrow();
    expect(() => sqlite.prepare("INSERT INTO items(id,sku,name,scan_code) VALUES('b','OTHER','Nut','BIN-M8')").run()).toThrow();
    expect(() => sqlite.prepare("INSERT INTO items(id,sku,name,scan_code) VALUES('b','','Nut','OTHER')").run()).toThrow();
    expect(() => sqlite.prepare("INSERT INTO items(id,sku,name,scan_code) VALUES('b','BIN-M8','Nut','OTHER')").run()).toThrow('AMBIGUOUS_ITEM_CODE');
  });
});
