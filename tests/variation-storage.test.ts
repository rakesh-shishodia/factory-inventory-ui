import { readFileSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';
import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EcwidClient, EcwidError, type EcwidOrder, type EcwidProduct } from '../src/ecwid';
import { createMovement } from '../src/inventory';
import { claimOutbox, processOutbox, processWebhook, upsertOrderSnapshot, type SyncEnv } from '../src/sync';
import { applyMigrations, sqliteD1 } from './d1';

const timestamp = '2026-09-22T08:00:01.000Z';
const signature = (size: string) => JSON.stringify([{ name: 'Size', value: size }]);
let sqlite: Database.Database;
let db: D1Database;
let env: SyncEnv;

function insertVariation(id: string, combination: string, size: string, quantity: number) {
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,on_hand,last_ecwid_quantity)
    VALUES(?,?,?,?,?,?,?,?,?)`).run(id, `BOLT-${size}`, `Bolt ${size}`, `BIN-${size}`, '1001', combination, signature(size), quantity, quantity);
}

beforeEach(() => {
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  insertVariation('small', '201', 'M3', 20);
  insertVariation('large', '202', 'M8', 40);
  insertVariation('other', '203', 'M10', 60);
  db = sqliteD1(sqlite);
  env = { DB: db, ECWID_MODE: 'live', ECWID_STORE_ID: '123', ECWID_TOKEN: 'test-token',
    ECWID_CLIENT_SECRET: 'test-secret', LIVE_SYNC_ENABLED: 'true',
    SYNC_QUEUE: { send: vi.fn().mockResolvedValue(undefined) } as unknown as Queue };
});
afterEach(() => { sqlite.close(); vi.restoreAllMocks(); });

function order(payment = 'AWAITING_PAYMENT', revision = 1, target = 'small'): EcwidOrder {
  const size = target === 'small' ? 'M3' : 'M8';
  return { id: 'ORDER', paymentStatus: payment, fulfillmentStatus: 'AWAITING_PROCESSING',
    updatedAt: `2026-09-22T08:00:0${revision}.000Z`, items: [{ id: 'line', productId: '1001', sku: `BOLT-${size}`,
      name: `Bolt ${size}`, quantity: 3, combinationId: target === 'small' ? '201' : '202',
      selectedOptions: [{ name: 'Size', value: size }], digital: false, trackQuantity: false }] };
}

async function movement(item: string, type: 'INTERNAL_USE' | 'RESTOCK' | 'ECWID_PICK' = 'INTERNAL_USE', quantity = 1) {
  const id = crypto.randomUUID();
  await createMovement(db, { operation_id: id, type, item_id: item, quantity,
    ...(type === 'ECWID_PICK' ? { order_id: 'ORDER', order_line_id: 'ORDER:line' } : {}) }, 'picker');
  return id;
}

function target(combination: string, size: string, quantity: number): EcwidProduct {
  return { id: '1001', combinationId: combination, sku: `BOLT-${size}`, name: `Bolt ${size}`, quantity,
    unlimited: false, enabled: true, hasOptions: true, hasVariations: false,
    variationOptions: [{ name: 'Size', value: size }], hasBundleRelationships: false,
    hasExtraOptions: false, eligibilityVerified: true };
}

function webhook(type = 'product.updated') {
  sqlite.prepare(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,received_at,updated_at)
    VALUES('event',?,'1001','123','{}',?,?)`).run(type, timestamp, timestamp);
}

describe('variation storage migration', () => {
  it('adds a clearly separated variation demo and reseeds without changing either ledger', async () => {
    const old = new Database(':memory:');
    try {
      applyMigrations(old);
      const adapter = sqliteD1(old);
      old.exec(readFileSync(new NodeURL('../fixtures/demo.sql', import.meta.url), 'utf8'));
      await createMovement(adapter, { operation_id: crypto.randomUUID(), type: 'INTERNAL_USE', item_id: 'item-nut-m3', quantity: 1 }, 'picker');
      const originalItems = old.prepare('SELECT * FROM items ORDER BY id').all();
      const originalMovements = old.prepare('SELECT * FROM movements ORDER BY id').all();
      const seed = readFileSync(new NodeURL('../fixtures/variation-demo.sql', import.meta.url), 'utf8');
      old.exec(seed);
      expect(old.prepare("SELECT id,on_hand,reserved,available FROM item_stock WHERE id LIKE 'DEMO-VAR-%' ORDER BY id").all()).toEqual([
        { id: 'DEMO-VAR-BOLT-M6-20', on_hand: 30, reserved: 4, available: 26 },
        { id: 'DEMO-VAR-BOLT-M6-30', on_hand: 50, reserved: 5, available: 45 },
      ]);
      expect(old.prepare("SELECT * FROM items WHERE id NOT LIKE 'DEMO-VAR-%' ORDER BY id").all()).toEqual(originalItems);
      expect(old.prepare('SELECT * FROM movements ORDER BY id').all()).toEqual(originalMovements);
      await createMovement(adapter, { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'DEMO-VAR-BOLT-M6-20',
        quantity: 1, order_id: 'DEMO-VARIATION-PAID', order_line_id: 'DEMO-VARIATION-PAID:1' }, 'picker');
      await expect(createMovement(adapter, { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: 'DEMO-VAR-BOLT-M6-30',
        quantity: 1, order_id: 'DEMO-VARIATION-AWAITING', order_line_id: 'DEMO-VARIATION-AWAITING:1' }, 'picker'))
        .rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
      old.prepare("UPDATE orders SET payment_status='PAID' WHERE id='DEMO-VARIATION-AWAITING'").run();
      await createMovement(adapter, { operation_id: crypto.randomUUID(), type: 'RESTOCK', item_id: 'DEMO-VAR-BOLT-M6-30', quantity: 2 }, 'picker');
      const tables = ['items', 'orders', 'order_lines', 'opening_balances', 'movements', 'outbox'];
      const beforeReseed = tables.map(table => old.prepare(`SELECT * FROM ${table} ORDER BY id`).all());
      old.exec(seed);
      tables.forEach((table, index) => expect(old.prepare(`SELECT * FROM ${table} ORDER BY id`).all()).toEqual(beforeReseed[index]));
      expect(old.prepare("SELECT id,on_hand,reserved FROM item_stock WHERE id LIKE 'DEMO-VAR-%' ORDER BY id").all()).toEqual([
        { id: 'DEMO-VAR-BOLT-M6-20', on_hand: 29, reserved: 3 },
        { id: 'DEMO-VAR-BOLT-M6-30', on_hand: 52, reserved: 5 },
      ]);
      expect(old.prepare("SELECT ecwid_product_id,ecwid_combination_id,quantity_delta FROM outbox WHERE item_id='DEMO-VAR-BOLT-M6-30'").all())
        .toEqual([{ ecwid_product_id: '900001', ecwid_combination_id: '910002', quantity_delta: 2 }]);
      expect(old.pragma('foreign_key_check')).toEqual([]);
    } finally { old.close(); }
  });

  it('preserves a populated old database and all immutable audit rows with foreign keys on', async () => {
    const old = new Database(':memory:');
    try {
      old.pragma('foreign_keys = ON');
      old.exec(readFileSync(new NodeURL('../migrations/0001_inventory.sql', import.meta.url), 'utf8'));
      // A populated v1 database with an opening ledger, reservation, pick and pending write.
      old.exec(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id) VALUES('legacy','LEGACY','Legacy','LEGACY','9001');
        INSERT INTO opening_balances VALUES('opening','legacy',10,'approved-source','picker','2026-09-22');
        INSERT INTO orders(id,payment_status,remote_updated_at,updated_at) VALUES('legacy-order','PAID','2026-09-22','2026-09-22');
        INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
          VALUES('legacy-line','legacy-order','1','legacy','LEGACY','Legacy',3);`);
      // Historical schema fixture: do not route v1 data through today's API,
      // which correctly requires fields introduced by later migrations.
      old.exec(`INSERT INTO movements(id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,order_id,order_line_id,actor,created_at)
        VALUES('legacy-pick','legacy-pick','ECWID_PICK','legacy',1,-1,0,'legacy-order','legacy-line','picker','2026-09-22');
        INSERT INTO movements(id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,actor,created_at)
        VALUES('legacy-restock','legacy-restock','RESTOCK','legacy',2,2,2,'picker','2026-09-22');`);
      const tables = ['items', 'orders', 'order_lines', 'opening_balances', 'movements', 'outbox'];
      const before = tables.map(table => old.prepare(`SELECT * FROM ${table}`).all());
      old.transaction(() => old.exec(readFileSync(new NodeURL('../migrations/0002_variation_inventory.sql', import.meta.url), 'utf8')))();
      tables.forEach((table, index) => {
        const after = old.prepare(`SELECT * FROM ${table}`).all() as Record<string, unknown>[];
        for (const row of after) { delete row.ecwid_combination_id; delete row.ecwid_option_signature; }
        expect(after).toEqual(before[index]);
      });
      expect(old.pragma('foreign_key_check')).toEqual([]);
      expect(old.pragma('foreign_keys', { simple: true })).toBe(1);
      expect(old.prepare('SELECT on_hand,reserved,available FROM item_stock').get()).toEqual({ on_hand: 11, reserved: 2, available: 9 });
      expect(() => old.prepare('DELETE FROM items').run()).toThrow('FOREIGN KEY');
      expect(() => old.prepare('DELETE FROM movements').run()).toThrow('MOVEMENT_IMMUTABLE');
      expect(() => old.prepare('UPDATE opening_balances SET on_hand=99').run()).toThrow('OPENING_BALANCE_IMMUTABLE');
      old.exec(`INSERT INTO movements(id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,actor,created_at)
        VALUES('legacy-use','legacy-use','INTERNAL_USE','legacy',1,-1,-1,'picker','2026-09-22');`);
      expect(old.prepare('SELECT on_hand FROM items').get()).toEqual({ on_hand: 10 });
    } finally { old.close(); }
  });

  it('keeps demo seed idempotent after partial picks and a migration', async () => {
    const old = new Database(':memory:');
    try {
      old.pragma('foreign_keys = ON');
      old.exec(readFileSync(new NodeURL('../migrations/0001_inventory.sql', import.meta.url), 'utf8'));
      // Capture only the original simple rows; newer demo variants are v2-only.
      old.exec(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id) VALUES('item-nut-m3','NUT-M3','M3 hex nut','NUT-M3','10001');
        INSERT INTO opening_balances VALUES('demo-opening:item-nut-m3','item-nut-m3',200,'LOCAL DEMO','demo','2026-09-22');
        INSERT INTO orders(id,payment_status,remote_updated_at,updated_at) VALUES('DEMO-1001','PAID','2026-09-22','2026-09-22');
        INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
          VALUES('DEMO-1001:1','DEMO-1001','1','item-nut-m3','NUT-M3','M3 hex nut',12);`);
      old.exec(`INSERT INTO movements(id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,order_id,order_line_id,actor,created_at)
        VALUES('legacy-demo-pick','legacy-demo-pick','ECWID_PICK','item-nut-m3',2,-2,0,'DEMO-1001','DEMO-1001:1','picker','2026-09-22');`);
      old.transaction(() => old.exec(readFileSync(new NodeURL('../migrations/0002_variation_inventory.sql', import.meta.url), 'utf8')))();
      const seed = readFileSync(new NodeURL('../fixtures/demo.sql', import.meta.url), 'utf8');
      old.exec(seed);
      const first = old.prepare('SELECT * FROM items ORDER BY id').all();
      old.exec(seed);
      expect(old.prepare('SELECT * FROM items ORDER BY id').all()).toEqual(first);
      expect(old.prepare("SELECT on_hand FROM items WHERE id='item-nut-m3'").get()).toEqual({ on_hand: 198 });
      expect(old.prepare("SELECT picked_qty FROM order_lines WHERE id='DEMO-1001:1'").get()).toEqual({ picked_qty: 2 });
      expect(old.pragma('foreign_key_check')).toEqual([]);
    } finally { old.close(); }
  });

  it('allows independent siblings but rejects duplicate targets and global SKU collisions', () => {
    expect(sqlite.prepare('SELECT count(*) AS count FROM items WHERE ecwid_product_id=?').get('1001')).toEqual({ count: 3 });
    expect(() => insertVariation('duplicate', '201', 'M4', 10)).toThrow('UNIQUE');
    expect(() => insertVariation('duplicate', '204', 'M3', 10)).toThrow('UNIQUE');
    expect(() => sqlite.prepare("UPDATE items SET ecwid_product_id=NULL WHERE id='small'").run()).toThrow('CHECK');
  });

  it.each(['opening', 'movement', 'order'])('locks identity after %s starts', async source => {
    if (source === 'opening') {
      sqlite.prepare("UPDATE items SET on_hand=0 WHERE id='small'").run();
      sqlite.prepare(`INSERT INTO opening_balances VALUES('opening','small',20,'source','picker',?)`).run(timestamp);
    } else if (source === 'movement') await movement('small');
    else await upsertOrderSnapshot(db, order());
    for (const change of ["ecwid_product_id='1002'", "ecwid_combination_id='299'", "sku='OTHER-SKU'", `ecwid_option_signature='${signature('M4')}'`]) {
      expect(() => sqlite.prepare(`UPDATE items SET ${change} WHERE id='small'`).run()).toThrow('ITEM_MAPPING_IMMUTABLE');
    }
    expect(() => sqlite.prepare("UPDATE items SET location='New shelf' WHERE id='small'").run()).not.toThrow();
  });
});

describe('variation orders and outbox', () => {
  it('reserves the exact size, permits only Paid picking, and never sends a second Ecwid deduction', async () => {
    await upsertOrderSnapshot(db, order());
    expect(sqlite.prepare('SELECT id,reserved FROM item_stock ORDER BY id').all()).toEqual([
      { id: 'large', reserved: 0 }, { id: 'other', reserved: 0 }, { id: 'small', reserved: 3 },
    ]);
    await expect(movement('small', 'ECWID_PICK')).rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
    await upsertOrderSnapshot(db, order('PAID', 2));
    await expect(movement('large', 'ECWID_PICK')).rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
    await movement('small', 'ECWID_PICK');
    await upsertOrderSnapshot(db, order('PAID', 3));
    expect(sqlite.prepare("SELECT on_hand,reserved,available FROM item_stock WHERE id='small'").get())
      .toEqual({ on_hand: 19, reserved: 2, available: 17 });
    expect(sqlite.prepare('SELECT count(*) AS count FROM outbox').get()).toEqual({ count: 0 });
    await upsertOrderSnapshot(db, order('CANCELLED', 4));
    await expect(movement('small')).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
    await expect(movement('large')).resolves.toEqual(expect.any(String));
    expect(sqlite.prepare("SELECT on_hand FROM items WHERE id='small'").get()).toEqual({ on_hand: 19 });
  });

  it('blocks the old and edited variation but leaves another sibling usable', async () => {
    await upsertOrderSnapshot(db, order('PAID'));
    await upsertOrderSnapshot(db, order('PAID', 2, 'large'));
    expect(sqlite.prepare('SELECT item_id FROM sync_issues WHERE item_id IS NOT NULL ORDER BY item_id').all())
      .toEqual([{ item_id: 'large' }, { item_id: 'small' }]);
    expect(sqlite.prepare('SELECT item_id,ordered_qty FROM order_lines').get()).toEqual({ item_id: 'small', ordered_qty: 3 });
    await expect(movement('other')).resolves.toEqual(expect.any(String));
  });

  it('canonicalizes selection order without falsely flagging an unchanged order', async () => {
    const expected = [{ name: 'Length', value: '12mm' }, { name: 'Size', value: 'M3' }];
    sqlite.prepare("UPDATE items SET ecwid_option_signature=? WHERE id='small'").run(JSON.stringify(expected));
    const first = order('PAID');
    first.items[0].selectedOptions = [...expected].reverse();
    await upsertOrderSnapshot(db, first);
    const next = order('PAID', 2);
    next.items[0].selectedOptions = expected;
    await upsertOrderSnapshot(db, next);
    expect(sqlite.prepare('SELECT item_id FROM order_lines').get()).toEqual({ item_id: 'small' });
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 0 });
  });

  it.each(['wrong-size', 'extra-option', 'missing-option', 'wrong-sku', 'missing-combination'])('rejects %s order metadata without mapping to another size', async fault => {
    const snapshot = order('PAID');
    if (fault === 'wrong-size') snapshot.items[0].selectedOptions = [{ name: 'Size', value: 'M8' }];
    if (fault === 'extra-option') snapshot.items[0].selectedOptions.push({ name: 'Add washer', value: 'Yes' });
    if (fault === 'missing-option') snapshot.items[0].selectedOptions = [];
    if (fault === 'wrong-sku') snapshot.items[0].sku = 'BOLT-M8';
    if (fault === 'missing-combination') snapshot.items[0].combinationId = null;
    await upsertOrderSnapshot(db, snapshot);
    expect(sqlite.prepare('SELECT item_id FROM order_lines').get()).toEqual({ item_id: null });
    expect(sqlite.prepare('SELECT needs_review FROM orders').get()).toEqual({ needs_review: 1 });
    expect(sqlite.prepare("SELECT reserved FROM item_stock WHERE id='large'").get()).toEqual({ reserved: 0 });
    expect(sqlite.prepare("SELECT item_id FROM sync_issues WHERE item_id='small'").get()).toEqual({ item_id: 'small' });
    if (fault === 'wrong-sku') {
      expect(sqlite.prepare("SELECT item_id FROM sync_issues WHERE item_id='large'").get()).toEqual({ item_id: 'large' });
    }
    await expect(movement('other')).resolves.toEqual(expect.any(String));
  });

  it('snapshots each combination and sends stock deltas to its own inventory endpoint', async () => {
    const small = await movement('small');
    const large = await movement('large', 'RESTOCK', 5);
    expect(sqlite.prepare('SELECT item_id,ecwid_product_id,ecwid_combination_id,quantity_delta FROM outbox ORDER BY rowid').all()).toEqual([
      { item_id: 'small', ecwid_product_id: '1001', ecwid_combination_id: '201', quantity_delta: -1 },
      { item_id: 'large', ecwid_product_id: '1001', ecwid_combination_id: '202', quantity_delta: 5 },
    ]);
    expect(() => sqlite.prepare("UPDATE outbox SET ecwid_combination_id='202' WHERE id=?").run(small)).toThrow('OUTBOX_TARGET_IMMUTABLE');
    const readBack = vi.spyOn(EcwidClient.prototype, 'getProductStock')
      .mockResolvedValueOnce(target('201', 'M3', 19)).mockResolvedValueOnce(target('202', 'M8', 45));
    const fetcher = vi.fn().mockImplementation(() => Promise.resolve(Response.json({ updateCount: 1 })));
    await processOutbox(env, small, fetcher);
    await processOutbox(env, large, fetcher);
    expect(fetcher.mock.calls.map(call => String(call[0]))).toEqual([
      'https://app.ecwid.com/api/v3/123/products/1001/combinations/201/inventory',
      'https://app.ecwid.com/api/v3/123/products/1001/combinations/202/inventory',
    ]);
    expect(readBack.mock.calls).toEqual([['1001', '201'], ['1001', '202']]);
    expect(sqlite.prepare('SELECT id,on_hand,last_ecwid_quantity FROM items WHERE id!=? ORDER BY id').all('other')).toEqual([
      { id: 'large', on_hand: 45, last_ecwid_quantity: 45 }, { id: 'small', on_hand: 19, last_ecwid_quantity: 19 },
    ]);
  });

  it('an uncertain write blocks that variation only, never resends and allows its sibling', async () => {
    const small = await movement('small');
    const large = await movement('large');
    const fetcher = vi.fn().mockRejectedValue(new Error('ambiguous'));
    await processOutbox(env, small, fetcher);
    await processOutbox(env, small, fetcher);
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(await claimOutbox(db, small)).toBeNull();
    expect(await claimOutbox(db, large)).toMatchObject({ ecwid_combination_id: '202' });
  });
});

describe('parent product events fan out safely', () => {
  it('quarantines all mapped siblings when a received variation stock value is invalid', async () => {
    webhook();
    const raw = { id: 1001, sku: 'CONTAINER', name: 'Bolt', quantity: 999, unlimited: false, enabled: true,
      options: [{ name: 'Size', type: 'SELECT', choices: [{ text: 'M3' }, { text: 'M8' }, { text: 'M10' }] }],
      combinations: [
        { id: 201, sku: 'BOLT-M3', quantity: 1.5, unlimited: false, options: [{ name: 'Size', value: 'M3' }] },
        { id: 202, sku: 'BOLT-M8', quantity: 40, unlimited: false, options: [{ name: 'Size', value: 'M8' }] },
        { id: 203, sku: 'BOLT-M10', quantity: 60, unlimited: false, options: [{ name: 'Size', value: 'M10' }] },
      ] };
    const fetcher = vi.fn().mockResolvedValue(Response.json(raw));
    await processWebhook(env, 'event', fetcher);
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(sqlite.prepare('SELECT status,last_error FROM webhook_events').get()).toEqual({
      status: 'BLOCKED', last_error: 'Ecwid product contains invalid stock or variation metadata.',
    });
    expect(sqlite.prepare('SELECT item_id,kind FROM sync_issues WHERE item_id IS NOT NULL ORDER BY item_id').all()).toEqual([
      { item_id: 'large', kind: 'PRODUCT_REVIEW' }, { item_id: 'other', kind: 'PRODUCT_REVIEW' },
      { item_id: 'small', kind: 'PRODUCT_REVIEW' },
    ]);
    expect(sqlite.prepare('SELECT count(*) AS count FROM items WHERE last_ecwid_quantity IS NULL').get()).toEqual({ count: 3 });
    for (const item of ['small', 'large', 'other']) {
      await expect(movement(item)).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
    }
    expect(sqlite.prepare('SELECT count(*) AS count FROM outbox').get()).toEqual({ count: 0 });
    expect(sqlite.prepare("SELECT on_hand FROM items WHERE id='small'").get()).toEqual({ on_hand: 20 });
  });

  it('retains retryable transport failure behavior without treating a network outage as invalid metadata', async () => {
    webhook();
    await expect(processWebhook(env, 'event', vi.fn().mockRejectedValue(new Error('network unavailable'))))
      .rejects.toMatchObject({ outcome: 'RETRYABLE' });
    expect(sqlite.prepare('SELECT status FROM webhook_events').get()).toEqual({ status: 'PENDING' });
    expect(sqlite.prepare('SELECT count(*) AS count FROM sync_issues').get()).toEqual({ count: 0 });
    expect(sqlite.prepare("SELECT last_ecwid_quantity FROM items WHERE id='small'").get()).toEqual({ last_ecwid_quantity: 20 });
  });

  it('refreshes all sibling mirrors from one parent fetch without changing physical stock', async () => {
    webhook();
    const read = vi.spyOn(EcwidClient.prototype, 'getProductStockTargets').mockResolvedValue([
      target('201', 'M3', 15), target('202', 'M8', 35), target('203', 'M10', 55),
    ]);
    await processWebhook(env, 'event');
    expect(read).toHaveBeenCalledTimes(1);
    expect(sqlite.prepare('SELECT id,on_hand,last_ecwid_quantity FROM items ORDER BY id').all()).toEqual([
      { id: 'large', on_hand: 40, last_ecwid_quantity: 35 }, { id: 'other', on_hand: 60, last_ecwid_quantity: 55 },
      { id: 'small', on_hand: 20, last_ecwid_quantity: 15 },
    ]);
    expect(sqlite.prepare('SELECT count(*) AS count FROM sync_issues').get()).toEqual({ count: 0 });
  });

  it('a removed variation blocks only its mapping, without falling back to a parent or sibling', async () => {
    webhook();
    vi.spyOn(EcwidClient.prototype, 'getProductStockTargets').mockResolvedValue([
      target('202', 'M8', 35), target('203', 'M10', 55),
    ]);
    await processWebhook(env, 'event');
    expect(sqlite.prepare('SELECT item_id FROM sync_issues').all()).toEqual([{ item_id: 'small' }]);
    expect(sqlite.prepare("SELECT last_ecwid_quantity FROM items WHERE id='small'").get()).toEqual({ last_ecwid_quantity: null });
    await expect(movement('small')).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
    await expect(movement('large')).resolves.toEqual(expect.any(String));
  });

  it.each(['product.deleted', '404'])('%s blocks every sibling and clears stale remote mirrors', async event => {
    webhook(event === '404' ? 'product.updated' : event);
    const read = vi.spyOn(EcwidClient.prototype, 'getProductStockTargets')
      .mockRejectedValue(new EcwidError('Missing', 'REJECTED', 404));
    await processWebhook(env, 'event');
    expect(sqlite.prepare('SELECT item_id FROM sync_issues WHERE item_id IS NOT NULL ORDER BY item_id').all())
      .toEqual([{ item_id: 'large' }, { item_id: 'other' }, { item_id: 'small' }]);
    expect(sqlite.prepare('SELECT count(*) AS count FROM items WHERE last_ecwid_quantity IS NULL').get()).toEqual({ count: 3 });
    expect(read).toHaveBeenCalledTimes(event === '404' ? 1 : 0);
  });

  it.each(['untracked', 'bundle', 'extra-options', 'identity', 'unknown-eligibility'])('blocks a changed %s target without assigning its quantity', async fault => {
    webhook();
    const changed = target('201', 'M3', 999);
    if (fault === 'untracked') changed.unlimited = true;
    if (fault === 'bundle') changed.hasBundleRelationships = true;
    if (fault === 'extra-options') changed.hasExtraOptions = true;
    if (fault === 'identity') changed.sku = 'OTHER';
    if (fault === 'unknown-eligibility') changed.eligibilityVerified = false;
    vi.spyOn(EcwidClient.prototype, 'getProductStockTargets').mockResolvedValue([
      changed, target('202', 'M8', 35), target('203', 'M10', 55),
    ]);
    await processWebhook(env, 'event');
    expect(sqlite.prepare('SELECT item_id FROM sync_issues').all()).toEqual([{ item_id: 'small' }]);
    expect(sqlite.prepare("SELECT on_hand,last_ecwid_quantity FROM items WHERE id='small'").get())
      .toEqual({ on_hand: 20, last_ecwid_quantity: null });
  });
});
