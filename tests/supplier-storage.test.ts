import Database from 'better-sqlite3';
import { readFileSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { applyMigrations, sqliteD1 } from './d1';

let sql: Database.Database;
const now = '2026-09-22T00:00:00.000Z';
const supplier = 'SUPPLIER_BACKED_UNLIMITED';
const read = (path: string) => readFileSync(new NodeURL(`../${path}`, import.meta.url), 'utf8');
const migrate = (name: string) => sql.transaction(() => sql.exec(read(`migrations/${name}`)))();
const snapshot = (table: string) => sql.prepare(`SELECT * FROM ${table} ORDER BY rowid`).all();
function item(id = 'part', count: number | null = 0, mode = supplier) {
  sql.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,inventory_mode,supplier_name,active)
    VALUES(?,?,?,?,?,?,?,0)`).run(id,id.toUpperCase(),id,id.toUpperCase(),`100-${id}`,mode,'Supplier');
  if (count !== null) {
    sql.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
      VALUES(?,?,?,'counted test','counter',?)`).run(`opening:${id}`,id,count,now);
    sql.prepare('UPDATE items SET active=1 WHERE id=?').run(id);
  }
}
function order(id = 'paid', qty = 20, payment = 'PAID', itemId = 'part') {
  sql.prepare(`INSERT INTO orders(id,payment_status,remote_updated_at,updated_at) VALUES(?,?,?,?)
    ON CONFLICT(id) DO NOTHING`).run(id,payment,now,now);
  sql.prepare(`INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
    SELECT ?,?,?,id,sku,name,? FROM items WHERE id=?`).run(`${id}:${itemId}`,id,itemId,qty,itemId);
}
function allocate(qty: number, orderId = 'paid', id = crypto.randomUUID(), type = 'ALLOCATE', itemId = 'part') {
  sql.prepare(`INSERT INTO supplier_allocation_events
    (id,fingerprint,type,item_id,order_id,order_line_id,quantity,quantity_delta,actor,created_at)
    VALUES(?,?,?,?,?,?,?,?,?,?)`).run(id,`fingerprint:${id}`,type,itemId,orderId,`${orderId}:${itemId}`,qty,
      type === 'ALLOCATE' ? qty : -qty,'picker',now);
  return id;
}
function move(qty: number, type = 'RESTOCK', orderId = 'paid', id = crypto.randomUUID(), itemId = 'part', mode = supplier) {
  sql.prepare(`INSERT INTO movements
    (id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,order_id,order_line_id,actor,created_at,inventory_mode)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)`).run(id,`fingerprint:${id}`,type,itemId,qty,type === 'RESTOCK' ? qty : -qty,
      mode === supplier || type === 'ECWID_PICK' ? 0 : type === 'RESTOCK' ? qty : -qty,
      type === 'ECWID_PICK' ? orderId : null,type === 'ECWID_PICK' ? `${orderId}:${itemId}` : null,'picker',now,mode);
  return id;
}
function stock(itemId = 'part') {
  return sql.prepare('SELECT on_hand,allocated,free,supplier_demand,unallocated_demand,uncovered_demand FROM item_stock WHERE id=?').get(itemId);
}
beforeEach(() => { sql = new Database(':memory:'); applyMigrations(sql); });
afterEach(() => sql.close());

describe('supplier schema and opening stock', () => {
  it('requires inactive registration and an explicit opening count including zero', () => {
    expect(() => sql.exec(`INSERT INTO items(id,sku,name,scan_code,inventory_mode) VALUES('bad','BAD','bad','BAD','${supplier}')`))
      .toThrow('SUPPLIER_OPENING_REQUIRED');
    item('part',null);
    expect(sql.prepare("SELECT active,opening_verified FROM item_stock WHERE id='part'").get()).toEqual({active:0,opening_verified:0});
    expect(() => sql.exec("UPDATE items SET active=1 WHERE id='part'")).toThrow('SUPPLIER_OPENING_REQUIRED');
    expect(() => move(2)).toThrow('ITEM_UNAVAILABLE');
    sql.prepare(`INSERT INTO opening_balances VALUES('count','part',0,'confirmed zero','counter',?)`).run(now);
    sql.exec("UPDATE items SET active=1 WHERE id='part'");
    expect(sql.prepare("SELECT on_hand,active,opening_verified FROM item_stock WHERE id='part'").get())
      .toEqual({on_hand:0,active:1,opening_verified:1});
  });
  it('does not allow initial on-hand without its opening ledger', () => {
    expect(() => sql.exec(`INSERT INTO items(id,sku,name,scan_code,inventory_mode,active,on_hand)
      VALUES('bad','BAD','bad','BAD','${supplier}',0,5)`)).toThrow('SUPPLIER_OPENING_REQUIRED');
  });
  it('requires supplier mapping at activation and prevents reinterpretation of historical mode', () => {
    item();
    expect(() => sql.exec("UPDATE items SET inventory_mode='STOCK_LIMITED' WHERE id='part'")).toThrow('INVENTORY_MODE_IMMUTABLE');
    expect(() => sql.exec("UPDATE items SET on_hand=3 WHERE id='part'")).toThrow('SUPPLIER_BALANCE_REQUIRES_LEDGER');
    expect(() => sql.exec("UPDATE items SET ecwid_product_id='other' WHERE id='part'")).toThrow('ITEM_MAPPING_IMMUTABLE');
    item('missing',null);
    sql.exec("UPDATE items SET ecwid_product_id=NULL WHERE id='missing'");
    sql.prepare(`INSERT INTO opening_balances VALUES('count','missing',0,'confirmed zero','counter',?)`).run(now);
    expect(() => sql.exec("UPDATE items SET active=1 WHERE id='missing'")).toThrow('SUPPLIER_MAPPING_REQUIRED');
  });
});

describe('supplier atomic movement and allocation ledger', () => {
  beforeEach(() => { item(); order(); });
  it('receive exactly, allocate, then pick conserves stock and creates no outbox', () => {
    move(20); allocate(20); move(20,'ECWID_PICK');
    expect(stock()).toEqual({on_hand:0,allocated:0,free:0,supplier_demand:0,unallocated_demand:0,uncovered_demand:0});
    expect(sql.prepare('SELECT type,quantity_delta FROM supplier_allocation_events ORDER BY rowid').all())
      .toEqual([{type:'ALLOCATE',quantity_delta:20},{type:'PICK',quantity_delta:-20}]);
    expect(snapshot('outbox')).toEqual([]);
    expect(sql.prepare('SELECT picked_qty FROM order_lines').get()).toEqual({picked_qty:20});
    expect(sql.pragma('foreign_key_check')).toEqual([]);
  });
  it('keeps surplus free and does not allocate receipts implicitly', () => {
    move(25);
    expect(stock()).toEqual({on_hand:25,allocated:0,free:25,supplier_demand:20,unallocated_demand:20,uncovered_demand:0});
    allocate(20);
    expect(stock()).toEqual({on_hand:25,allocated:20,free:5,supplier_demand:20,unallocated_demand:0,uncovered_demand:0});
    move(20,'ECWID_PICK');
    expect(stock()).toEqual({on_hand:5,allocated:0,free:5,supplier_demand:0,unallocated_demand:0,uncovered_demand:0});
  });
  it('handles partial receipt and partial picks with uncovered demand remaining', () => {
    move(8); allocate(8); move(3,'ECWID_PICK');
    expect(stock()).toEqual({on_hand:5,allocated:5,free:0,supplier_demand:17,unallocated_demand:12,uncovered_demand:12});
    expect(() => move(6,'ECWID_PICK')).toThrow('INSUFFICIENT_PHYSICAL_STOCK');
    expect(() => allocate(1)).toThrow('INSUFFICIENT_FREE_SUPPLIER_STOCK');
  });
  it('rejects order over-allocation and picking another line allocation', () => {
    order('second'); move(30); allocate(20,'second');
    expect(() => move(1,'ECWID_PICK')).toThrow('INSUFFICIENT_LINE_ALLOCATION');
    expect(() => allocate(1,'second')).toThrow('ALLOCATION_QUANTITY_EXCEEDED');
    expect(() => allocate(21)).toThrow('INSUFFICIENT_FREE_SUPPLIER_STOCK');
  });
  it('release changes only assignment and cannot exceed assigned quantity', () => {
    move(20); allocate(15); allocate(4,'paid',crypto.randomUUID(),'RELEASE');
    expect(stock()).toEqual({on_hand:20,allocated:11,free:9,supplier_demand:20,unallocated_demand:9,uncovered_demand:0});
    expect(() => allocate(12,'paid',crypto.randomUUID(),'RELEASE')).toThrow('INSUFFICIENT_LINE_ALLOCATION');
    expect(snapshot('movements')).toHaveLength(1);
  });
  it.each(['EMAIL_SALE','INTERNAL_USE'])('%s spends only free supplier stock without an Ecwid write', type => {
    move(25); allocate(20); move(5,type);
    expect(() => move(1,type)).toThrow('INSUFFICIENT_AVAILABLE_STOCK');
    expect(stock()).toMatchObject({on_hand:20,allocated:20,free:0});
    expect(snapshot('outbox')).toEqual([]);
  });
  it('keeps Awaiting Payment allocation but blocks picks until Paid', () => {
    sql.exec("UPDATE orders SET payment_status='AWAITING_PAYMENT'");
    move(20); allocate(20);
    expect(sql.prepare('SELECT supplier_paid_demand,supplier_awaiting_payment_demand FROM item_stock').get())
      .toEqual({supplier_paid_demand:0,supplier_awaiting_payment_demand:20});
    expect(() => move(1,'ECWID_PICK')).toThrow('ORDER_NOT_PICKABLE');
    sql.exec("UPDATE orders SET payment_status='PAID'");
    move(20,'ECWID_PICK');
    expect(stock()).toMatchObject({on_hand:0,allocated:0});
  });
  it('rolls back a receipt when allocation fails within the same D1 transaction', async () => {
    const db=sqliteD1(sql);
    const receipt=db.prepare(`INSERT INTO movements(id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,actor,created_at,inventory_mode)
      VALUES('receipt','receipt','RESTOCK','part',8,8,0,'picker',?,'SUPPLIER_BACKED_UNLIMITED')`).bind(now);
    const allocation=db.prepare(`INSERT INTO supplier_allocation_events(id,fingerprint,type,item_id,order_id,order_line_id,quantity,quantity_delta,actor,created_at)
      VALUES('assign','assign','ALLOCATE','part','paid','paid:part',9,9,'picker',?)`).bind(now);
    await expect(db.batch([receipt,allocation])).rejects.toThrow('INSUFFICIENT_FREE_SUPPLIER_STOCK');
    expect(stock()).toMatchObject({on_hand:0,allocated:0});
    expect(snapshot('movements')).toEqual([]);
  });
  it('immutable ledgers and guards prohibit fabricated supplier writes', () => {
    const receipt=move(20); allocate(20);
    expect(() => sql.prepare('UPDATE movements SET quantity=1 WHERE id=?').run(receipt)).toThrow('MOVEMENT_IMMUTABLE');
    expect(() => sql.exec('DELETE FROM supplier_allocation_events')).toThrow('ALLOCATION_IMMUTABLE');
    expect(() => sql.exec('UPDATE supplier_allocation_events SET quantity=1')).toThrow('ALLOCATION_IMMUTABLE');
    expect(() => sql.exec("UPDATE order_lines SET ordered_qty=30")).toThrow('ALLOCATED_ORDER_LINE_IMMUTABLE');
    expect(() => sql.exec("UPDATE order_lines SET picked_qty=1")).toThrow('SUPPLIER_PICK_REQUIRES_MOVEMENT');
    expect(() => move(1,'RESTOCK','paid',crypto.randomUUID(),'part','STOCK_LIMITED')).toThrow('INVENTORY_MODE_MISMATCH');
    expect(() => sql.prepare(`INSERT INTO outbox(id,item_id,ecwid_product_id,quantity_delta,created_at,updated_at)
      VALUES(?,'part','100-part',1,?,?)`).run(receipt,now,now)).toThrow('SUPPLIER_ECWID_WRITE_FORBIDDEN');
  });
  it.each(['item','order','flag'])('blocks allocation and picking with an unresolved %s issue', kind => {
    move(20); allocate(20);
    if(kind==='flag') sql.exec("UPDATE orders SET needs_review=1");
    else sql.prepare(`INSERT INTO sync_issues(id,item_id,order_id,kind,message,created_at)
      VALUES('issue',?,?,'REVIEW','review',?)`).run(kind==='item'?'part':null,kind==='order'?'paid':null,now);
    expect(() => allocate(1,'paid',crypto.randomUUID(),'RELEASE')).toThrow(kind==='flag'?'ORDER_NOT_ALLOCATABLE':'ITEM_NEEDS_REVIEW');
    expect(() => move(1,'ECWID_PICK')).toThrow(kind==='flag'?'ORDER_NOT_PICKABLE':'ITEM_NEEDS_REVIEW');
    expect(stock()).toMatchObject({on_hand:20,allocated:20});
  });
});

describe('supplier cancellation and serialized contenders', () => {
  beforeEach(() => { item(); order('paid',2); order('second',2); move(2); });
  it.each(['CANCELLED','REFUNDED','INCOMPLETE'])('%s before picks releases only allocation, once', status => {
    allocate(2);
    sql.prepare('UPDATE orders SET payment_status=? WHERE id=?').run(status,'paid');
    sql.prepare('UPDATE orders SET payment_status=? WHERE id=?').run(status,'paid');
    expect(stock()).toMatchObject({on_hand:2,allocated:0,free:2,supplier_demand:2});
    expect(snapshot('supplier_allocation_events')).toHaveLength(2);
    expect(() => allocate(1)).toThrow('ORDER_NOT_ALLOCATABLE');
    expect(() => move(1,'ECWID_PICK')).toThrow('ORDER_NOT_PICKABLE');
    expect(snapshot('movements')).toHaveLength(1);
  });
  it('cancellation releases accumulated assignments above the per-action 1m limit', () => {
    sql.exec("UPDATE order_lines SET ordered_qty=1500000 WHERE id='paid:part'");
    move(1_000_000); move(500_000);
    allocate(1_000_000); allocate(500_000);
    const physicalBefore=(stock() as {on_hand:number}).on_hand;
    sql.exec("UPDATE orders SET payment_status='CANCELLED' WHERE id='paid'");
    expect(stock()).toMatchObject({on_hand:physicalBefore,allocated:0,free:physicalBefore});
    expect(sql.prepare("SELECT quantity,quantity_delta FROM supplier_allocation_events WHERE type='CANCEL_RELEASE'").all())
      .toEqual([{quantity:1_500_000,quantity_delta:-1_500_000}]);
    sql.exec("UPDATE orders SET payment_status='CANCELLED' WHERE id='paid'");
    expect(snapshot('supplier_allocation_events')).toHaveLength(3);
    expect(()=>allocate(1_000_001,'second')).toThrow();
    expect(()=>move(1_000_001)).toThrow();
    expect(snapshot('outbox')).toEqual([]);
  });
  it('partial-picked cancellation retains remaining allocation and creates no stock return', () => {
    allocate(2); move(1,'ECWID_PICK');
    sql.exec("UPDATE orders SET payment_status='CANCELLED' WHERE id='paid'");
    expect(sql.prepare("SELECT needs_review FROM orders WHERE id='paid'").get()).toEqual({needs_review:1});
    expect(stock()).toMatchObject({on_hand:1,allocated:1,free:0});
    expect(snapshot('supplier_allocation_events')).toHaveLength(2);
    expect(() => allocate(1,'paid',crypto.randomUUID(),'RELEASE')).toThrow('ORDER_NOT_ALLOCATABLE');
  });
  it('a picked stock-limited line also prevents releasing supplier allocations in a mixed cancellation', () => {
    item('stock',3,'STOCK_LIMITED'); order('paid',2,'PAID','stock'); allocate(2);
    move(1,'ECWID_PICK','paid',crypto.randomUUID(),'stock','STOCK_LIMITED');
    sql.exec("UPDATE orders SET payment_status='CANCELLED' WHERE id='paid'");
    expect(stock()).toMatchObject({on_hand:2,allocated:2});
    expect(sql.prepare("SELECT needs_review FROM orders WHERE id='paid'").get()).toEqual({needs_review:1});
  });
  it.each(['READY_FOR_PICKUP','SHIPPED','DELIVERED'])('terminal %s missing local picks holds assignments for review', status => {
    allocate(2);
    sql.prepare("UPDATE orders SET fulfillment_status=? WHERE id='paid'").run(status);
    expect(stock()).toMatchObject({on_hand:2,allocated:2});
    expect(sql.prepare("SELECT needs_review FROM orders WHERE id='paid'").get()).toEqual({needs_review:1});
    expect(() => move(1,'ECWID_PICK')).toThrow('ORDER_NOT_PICKABLE');
  });
  it('does not release when cancellation is accompanied by identity review', () => {
    allocate(2);
    sql.exec("UPDATE orders SET payment_status='CANCELLED',needs_review=1 WHERE id='paid'");
    expect(stock()).toMatchObject({on_hand:2,allocated:2});
  });
  it('serialized competing allocation inserts cannot claim the same final unit', async () => {
    move(1,'INTERNAL_USE');
    const attempts=await Promise.allSettled([Promise.resolve().then(()=>allocate(1)),Promise.resolve().then(()=>allocate(1,'second'))]);
    expect(attempts.filter(result=>result.status==='fulfilled')).toHaveLength(1);
    expect(attempts.filter(result=>result.status==='rejected')).toHaveLength(1);
    expect(stock()).toMatchObject({on_hand:1,allocated:1,free:0});
  });
  it('serialized competing pick inserts consume the same allocation at most once', async () => {
    allocate(1);
    const attempts=await Promise.allSettled([Promise.resolve().then(()=>move(1,'ECWID_PICK')),Promise.resolve().then(()=>move(1,'ECWID_PICK'))]);
    expect(attempts.filter(result=>result.status==='fulfilled')).toHaveLength(1);
    expect(stock()).toMatchObject({on_hand:1,allocated:0,free:1});
    expect(sql.prepare("SELECT picked_qty FROM order_lines WHERE id='paid:part'").get()).toEqual({picked_qty:1});
  });
  it.each(['cancel-first','allocate-first'])('allocation/cancellation race %s leaves no stranded allocations', ordering => {
    if(ordering==='allocate-first') allocate(2);
    sql.exec("UPDATE orders SET payment_status='CANCELLED' WHERE id='paid'");
    if(ordering==='cancel-first') expect(()=>allocate(2)).toThrow('ORDER_NOT_ALLOCATABLE');
    expect(stock()).toMatchObject({on_hand:2,allocated:0});
  });
  it.each(['cancel-first','pick-first'])('picking/cancellation race %s never invents stock', ordering => {
    allocate(2);
    if(ordering==='pick-first') move(1,'ECWID_PICK');
    sql.exec("UPDATE orders SET payment_status='CANCELLED' WHERE id='paid'");
    if(ordering==='cancel-first') expect(()=>move(1,'ECWID_PICK')).toThrow('ORDER_NOT_PICKABLE');
    expect(stock()).toMatchObject(ordering==='pick-first'?{on_hand:1,allocated:1}:{on_hand:2,allocated:0});
  });
  it('duplicate IDs cannot apply ledger effects twice', () => {
    const id=allocate(1); const picked=move(1,'ECWID_PICK');
    expect(()=>allocate(1,'paid',id)).toThrow();
    expect(()=>move(1,'ECWID_PICK','paid',picked)).toThrow();
    expect(stock()).toMatchObject({on_hand:1,allocated:0});
    expect(snapshot('movements')).toHaveLength(2);
  });
  it('one operation ID cannot be reused across physical and allocation ledgers', () => {
    const receipt=move(1);
    expect(()=>allocate(1,'paid',receipt)).toThrow('IDEMPOTENCY_CONFLICT');
    const assignment=allocate(1);
    expect(()=>move(1,'RESTOCK','paid',assignment)).toThrow('IDEMPOTENCY_CONFLICT');
    expect(stock()).toMatchObject({on_hand:3,allocated:1});
    expect(snapshot('movements')).toHaveLength(2);
    expect(snapshot('supplier_allocation_events')).toHaveLength(1);
  });
});

describe('forward migration and demo fixture preservation', () => {
  it('preserves all historical audit references, data and stock-limited semantics', () => {
    sql.close(); sql=new Database(':memory:'); sql.pragma('foreign_keys=ON');
    for(const name of ['0001_inventory.sql','0002_variation_inventory.sql','0003_pickup_fulfillment.sql','0004_opening_import_batches.sql']) migrate(name);
    sql.exec(read('fixtures/demo.sql'));
    sql.prepare(`INSERT INTO movements(id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,actor,created_at)
      VALUES('old-move','old','RESTOCK','item-nut-m3',2,2,2,'staff',?)`).run(now);
    const tables=['items','opening_balances','orders','order_lines','movements','outbox','sync_issues'];
    const before=tables.map(snapshot);
    const stockBefore=sql.prepare('SELECT id,reserved,available FROM item_stock ORDER BY id').all();
    migrate('0005_supplier_inventory.sql');
    const after=tables.map(table=>snapshot(table).map(value=>{
      const row={...(value as Record<string,unknown>)};
      if(table==='items'){expect(row.inventory_mode).toBe('STOCK_LIMITED');expect(row.supplier_name).toBe('');delete row.inventory_mode;delete row.supplier_name;}
      if(table==='movements'){expect(row.inventory_mode).toBe('STOCK_LIMITED');delete row.inventory_mode;}
      return row;
    }));
    expect(after).toEqual(before);
    expect(sql.pragma('foreign_key_check')).toEqual([]);
    expect(sql.prepare('SELECT id,reserved,available FROM item_stock ORDER BY id').all()).toEqual(stockBefore);
    expect(()=>sql.exec("DELETE FROM movements WHERE id='old-move'")).toThrow('MOVEMENT_IMMUTABLE');
    expect(()=>sql.exec("UPDATE outbox SET ecwid_product_id='wrong' WHERE id='old-move'")).toThrow('OUTBOX_TARGET_IMMUTABLE');
  });
  it('standalone fabricated demo has explicit zero counts and preserves progress on rerun', () => {
    sql.exec(read('fixtures/supplier-demo.sql'));
    expect(sql.prepare("SELECT count(*) AS n FROM items WHERE inventory_mode='SUPPLIER_BACKED_UNLIMITED' AND on_hand=0 AND active=1").get()).toEqual({n:2});
    expect(sql.prepare("SELECT count(*) AS n FROM opening_balances WHERE on_hand=0").get()).toEqual({n:2});
    move(25,'RESTOCK','SUP-DEMO-PAID',crypto.randomUUID(),'SUP-DEMO-ITEM-A');
    sql.exec("UPDATE orders SET payment_status='PAID' WHERE id='SUP-DEMO-AWAITING'");
    sql.exec("UPDATE items SET active=0 WHERE id='SUP-DEMO-ITEM-B'");
    const tables=['items','opening_balances','orders','order_lines','movements','outbox','sync_state'];
    const before=tables.map(snapshot);
    sql.exec(read('fixtures/supplier-demo.sql'));
    expect(tables.map(snapshot)).toEqual(before);
    expect(sql.pragma('foreign_key_check')).toEqual([]);
  });
});
