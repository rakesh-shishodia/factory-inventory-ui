import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createMovement, getItem, getOrder } from '../src/inventory';
import { createAllocation, listAllocations, validateAllocationInput } from '../src/supplier';
import { applyMigrations, sqliteD1 } from './d1';

let sqlite: Database.Database;
let db: D1Database;
const actor = 'receiver@example.test';
const timestamp = '2026-09-22T08:00:00.000Z';

function order(id = 'order-1',quantity = 20,payment = 'PAID') {
  sqlite.prepare(`INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
    VALUES(?,?,'AWAITING_PROCESSING',?,?)`).run(id,payment,timestamp,timestamp);
  sqlite.prepare(`INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
    VALUES(?,?,'1','supplier','MISUMI-NUT','Supplier nut',?)`).run(`${id}:1`,id,quantity);
}
function movement(type: 'RESTOCK'|'ECWID_PICK'|'INTERNAL_USE'|'EMAIL_SALE',quantity: number,extra = {}) {
  return {operation_id:crypto.randomUUID(),type,item_id:'supplier',quantity,...extra};
}
function allocation(quantity: number,type = 'ALLOCATE',id = 'order-1') {
  return {operation_id:crypto.randomUUID(),type,item_id:'supplier',order_id:id,order_line_id:`${id}:1`,quantity,note:'delivery 1'};
}
function pick(quantity: number,id = 'order-1') {
  return movement('ECWID_PICK',quantity,{order_id:id,order_line_id:`${id}:1`});
}

beforeEach(() => {
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  db = sqliteD1(sqlite);
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,inventory_mode,supplier_name,active)
    VALUES('supplier','MISUMI-NUT','Supplier nut','MISUMI-NUT','100','SUPPLIER_BACKED_UNLIMITED','MISUMI',0)`).run();
  sqlite.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
    VALUES('opening','supplier',0,'verified-local-fixture',?,?)`).run(actor,timestamp);
  sqlite.prepare("UPDATE items SET active=1 WHERE id='supplier'").run();
});
afterEach(() => sqlite.close());

describe('supplier receipts, allocation and physical picking',() => {
  it('keeps unknown opening stock inactive and blocks operations instead of inventing zero',async () => {
    sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,inventory_mode,supplier_name,active)
      VALUES('unknown','UNKNOWN','Unknown shelf count','UNKNOWN','101','SUPPLIER_BACKED_UNLIMITED','MISUMI',0)`).run();
    expect(() => sqlite.prepare("UPDATE items SET active=1 WHERE id='unknown'").run()).toThrow();
    await expect(createMovement(db,{...movement('RESTOCK',1),item_id:'unknown'},actor)).rejects.toMatchObject({status:409});
    expect(await getItem(db,'unknown')).toMatchObject({active:0,opening_verified:0});
  });

  it('receives exact demand, allocates explicitly, then picks without any Ecwid outbox',async () => {
    order();
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:0,allocated:0,free:0,unallocated_demand:20,uncovered_demand:20});
    expect((await getOrder(db,'order-1')).lines[0]).toMatchObject({fulfillment_state:'AWAITING_SUPPLIER',pickable_qty:0});
    await createMovement(db,movement('RESTOCK',20),actor);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:0,free:20,uncovered_demand:0});
    expect((await getOrder(db,'order-1')).lines[0]).toMatchObject({fulfillment_state:'AWAITING_ASSIGNMENT',allocated_qty:0,pickable_qty:0});
    await createAllocation(db,allocation(20),actor);
    expect((await getOrder(db,'order-1')).lines[0]).toMatchObject({fulfillment_state:'READY_TO_PICK',allocated_qty:20,pickable_qty:20});
    await createMovement(db,pick(20),actor);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:0,allocated:0,free:0,unallocated_demand:0,uncovered_demand:0});
    expect((await getOrder(db,'order-1')).lines[0]).toMatchObject({fulfillment_state:'PICKED',picked_qty:20});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM outbox').get()).toEqual({n:0});
    expect(sqlite.prepare('SELECT ecwid_quantity_delta,inventory_mode FROM movements').all()).toEqual([
      {ecwid_quantity_delta:0,inventory_mode:'SUPPLIER_BACKED_UNLIMITED'},
      {ecwid_quantity_delta:0,inventory_mode:'SUPPLIER_BACKED_UNLIMITED'},
    ]);
    expect((await listAllocations(db)).map(event => event.type).sort()).toEqual(['ALLOCATE','PICK']);
  });

  it('leaves surplus free and never auto-assigns received stock',async () => {
    order();
    await createMovement(db,movement('RESTOCK',25),actor);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:25,allocated:0,free:25,unallocated_demand:20,uncovered_demand:0});
    await createAllocation(db,allocation(20),actor);
    await createMovement(db,pick(20),actor);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:5,allocated:0,free:5});
  });

  it('permits partial receipt and pick without treating outstanding demand as physical stock',async () => {
    order();
    await createMovement(db,movement('RESTOCK',8),actor);
    await expect(createAllocation(db,allocation(9),actor)).rejects.toMatchObject({status:409});
    await createAllocation(db,allocation(8),actor);
    await expect(createMovement(db,pick(9),actor)).rejects.toMatchObject({status:409});
    await createMovement(db,pick(8),actor);
    expect((await getOrder(db,'order-1')).lines[0]).toMatchObject({picked_qty:8,remaining_qty:12,unallocated_qty:12,fulfillment_state:'AWAITING_SUPPLIER'});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:0,uncovered_demand:12});
  });

  it('holds unpaid allocation but gates picking on payment',async () => {
    order('order-1',20,'AWAITING_PAYMENT');
    await createMovement(db,movement('RESTOCK',20),actor);
    await createAllocation(db,allocation(20),actor);
    expect(await getItem(db,'supplier')).toMatchObject({paid_demand:0,awaiting_payment_demand:20,free:0});
    expect((await getOrder(db,'order-1')).lines[0]).toMatchObject({allocated_qty:20,pickable_qty:0,fulfillment_state:'AWAITING_PAYMENT'});
    await expect(createMovement(db,pick(1),actor)).rejects.toMatchObject({code:'ORDER_NOT_PICKABLE'});
    sqlite.prepare("UPDATE orders SET payment_status='PAID' WHERE id='order-1'").run();
    expect((await getOrder(db,'order-1')).lines[0]).toMatchObject({pickable_qty:20});
    await createMovement(db,pick(20),actor);
  });

  it('protects allocated stock from email sales and internal use without Ecwid adjustments',async () => {
    order();
    await createMovement(db,movement('RESTOCK',25),actor);
    await createAllocation(db,allocation(20),actor);
    await expect(createMovement(db,movement('EMAIL_SALE',6),actor)).rejects.toMatchObject({code:'INSUFFICIENT_AVAILABLE_STOCK'});
    await createMovement(db,movement('EMAIL_SALE',3),actor);
    await createMovement(db,movement('INTERNAL_USE',2),actor);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:20,free:0});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM outbox').get()).toEqual({n:0});
  });
});

describe('supplier concurrency, retries and allocation guards',() => {
  it('allows exactly one order to allocate the final free unit',async () => {
    order('first',1); order('second',1);
    await createMovement(db,movement('RESTOCK',1),actor);
    const outcomes = await Promise.allSettled([createAllocation(db,allocation(1,'ALLOCATE','first'),actor),createAllocation(db,allocation(1,'ALLOCATE','second'),actor)]);
    expect(outcomes.filter(value => value.status==='fulfilled')).toHaveLength(1);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:1,allocated:1,free:0,uncovered_demand:1});
  });

  it('allows exactly one picker to consume the final assigned unit',async () => {
    order('order-1',1);
    await createMovement(db,movement('RESTOCK',1),actor);
    await createAllocation(db,allocation(1),actor);
    const outcomes = await Promise.allSettled([createMovement(db,pick(1),actor),createMovement(db,pick(1),actor)]);
    expect(outcomes.filter(value => value.status==='fulfilled')).toHaveLength(1);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:0,allocated:0});
  });

  it('replays receipts, assignments and picks exactly and rejects changed payload or actor',async () => {
    order();
    const receipt = movement('RESTOCK',20);
    const assignment = allocation(20);
    const picking = pick(20);
    for (const input of [receipt,assignment,picking]) {
      const execute = input.type==='ALLOCATE' ? createAllocation : createMovement;
      expect(await execute(db,input,actor)).toMatchObject({duplicate:false});
      expect(await execute(db,input,actor.toUpperCase())).toMatchObject({duplicate:true});
      await expect(execute(db,{...input,quantity:19},actor)).rejects.toMatchObject({code:'IDEMPOTENCY_CONFLICT'});
      await expect(execute(db,input,'another@example.test')).rejects.toMatchObject({code:'IDEMPOTENCY_CONFLICT'});
    }
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:0,allocated:0});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM movements').get()).toEqual({n:2});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM supplier_allocation_events').get()).toEqual({n:2});
  });

  it('releases assignment with an immutable audit and no physical movement',async () => {
    order();
    await createMovement(db,movement('RESTOCK',20),actor);
    await createAllocation(db,allocation(20),actor);
    await createAllocation(db,allocation(5,'RELEASE'),actor);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:15,free:5,unallocated_demand:5,uncovered_demand:0});
    await expect(createAllocation(db,allocation(16,'RELEASE'),actor)).rejects.toMatchObject({status:409});
    expect(() => sqlite.exec('DELETE FROM supplier_allocation_events')).toThrow();
    expect(() => sqlite.exec("UPDATE supplier_allocation_events SET note='edited'")).toThrow();
  });

  it('rejects reusing an operation ID across receiving and allocation workflows',async () => {
    order();
    const receipt = movement('RESTOCK',20);
    await createMovement(db,receipt,actor);
    await expect(createAllocation(db,{...allocation(20),operation_id:receipt.operation_id},actor))
      .rejects.toMatchObject({code:'IDEMPOTENCY_CONFLICT',status:409});
    const assignment = allocation(20);
    await createAllocation(db,assignment,actor);
    await expect(createMovement(db,{...movement('RESTOCK',5),operation_id:assignment.operation_id},actor))
      .rejects.toMatchObject({code:'IDEMPOTENCY_CONFLICT',status:409});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:20,free:0});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM movements').get()).toEqual({n:1});
  });

  it('rejects wrong line, over-demand, closed and reviewed orders',async () => {
    order('order-1',2); order('other',2);
    await createMovement(db,movement('RESTOCK',20),actor);
    await expect(createAllocation(db,{...allocation(1),order_line_id:'other:1'},actor)).rejects.toMatchObject({status:409});
    await expect(createAllocation(db,allocation(3),actor)).rejects.toMatchObject({status:409});
    sqlite.prepare("UPDATE orders SET needs_review=1 WHERE id='order-1'").run();
    await expect(createAllocation(db,allocation(1),actor)).rejects.toMatchObject({status:409});
    sqlite.prepare("UPDATE orders SET fulfillment_status='SHIPPED' WHERE id='other'").run();
    await expect(createAllocation(db,allocation(1,'ALLOCATE','other'),actor)).rejects.toMatchObject({status:409});
  });

  it.each([0,-1,1.5,'2',null,1_000_001])('rejects invalid allocation quantity %s',quantity => {
    expect(() => validateAllocationInput({...allocation(1),quantity})).toThrow();
  });
  it('rejects missing targets, fake operation IDs and internal event types',() => {
    for (const change of [{order_line_id:''},{item_id:''},{operation_id:'x'},{type:'PICK'},{type:'CANCEL_RELEASE'}]) {
      expect(() => validateAllocationInput({...allocation(1),...change})).toThrow();
    }
  });
});
