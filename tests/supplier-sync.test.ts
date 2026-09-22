import Database from 'better-sqlite3';
import { afterEach,beforeEach,describe,expect,it,vi } from 'vitest';
import { createMovement,getItem,getOrder } from '../src/inventory';
import { createAllocation } from '../src/supplier';
import { claimOutbox,processOutbox,processWebhook,upsertOrderSnapshot,type SyncEnv } from '../src/sync';
import type { EcwidOrder } from '../src/ecwid';
import { applyMigrations,sqliteD1 } from './d1';

let sqlite: Database.Database;
let db: D1Database;
let env: SyncEnv;
const stamp = '2026-09-22T08:00:00.000Z';
const actor = 'picker@example.test';
const supplierLine = {id:'s',productId:'100',sku:'SUP-NUT',name:'Supplier nut',quantity:20,
  combinationId:null,selectedOptions:[],digital:false,trackQuantity:false};
const stockLine = {id:'l',productId:'200',sku:'LOCAL-BOLT',name:'Shelf bolt',quantity:2,
  combinationId:null,selectedOptions:[],digital:false,trackQuantity:true};
function order(paymentStatus = 'PAID',revision = 1): EcwidOrder {
  return {id:'ORDER',paymentStatus,fulfillmentStatus:'AWAITING_PROCESSING',updatedAt:`2026-09-22T08:00:0${revision}.000Z`,
    items:[{...supplierLine}]};
}
async function receive(quantity = 20) {
  return createMovement(db,{operation_id:crypto.randomUUID(),type:'RESTOCK',item_id:'supplier',quantity},actor);
}
async function assign(quantity = 20) {
  return createAllocation(db,{operation_id:crypto.randomUUID(),type:'ALLOCATE',item_id:'supplier',order_id:'ORDER',order_line_id:'ORDER:s',quantity},actor);
}
async function pick(quantity = 1,item = 'supplier',line = 's') {
  return createMovement(db,{operation_id:crypto.randomUUID(),type:'ECWID_PICK',item_id:item,order_id:'ORDER',order_line_id:`ORDER:${line}`,quantity},actor);
}
function product(overrides: Record<string,unknown> = {}) {
  return {id:100,sku:'SUP-NUT',name:'Supplier nut',enabled:true,unlimited:true,
    options:[],combinations:[],compositeParents:[],compositeComponents:[],...overrides};
}
function webhook(id = 'product',productId = '100') {
  sqlite.prepare(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,received_at,updated_at)
    VALUES(?,'product.updated',?,'123','{}',?,?)`).run(id,productId,stamp,stamp);
}
beforeEach(() => {
  sqlite = new Database(':memory:');applyMigrations(sqlite);db = sqliteD1(sqlite);
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,inventory_mode,supplier_name,active)
    VALUES('supplier','SUP-NUT','Supplier nut','SUP-NUT','100','SUPPLIER_BACKED_UNLIMITED','MISUMI',0)`).run();
  sqlite.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
    VALUES('opening','supplier',0,'verified fixture','admin',?)`).run(stamp);
  sqlite.prepare("UPDATE items SET active=1 WHERE id='supplier'").run();
  sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand,last_ecwid_quantity)
    VALUES('local','LOCAL-BOLT','Shelf bolt','LOCAL-BOLT','200',10,10)`).run();
  env = {DB:db,ECWID_MODE:'live',LIVE_SYNC_ENABLED:'true',ORDER_SYNC_ENABLED:'true',ECWID_STORE_ID:'123',ECWID_TOKEN:'test-token',
    SYNC_QUEUE:{send:vi.fn().mockResolvedValue(undefined)} as unknown as Queue};
});
afterEach(() => {sqlite.close();vi.restoreAllMocks();});

describe('supplier order sync and mixed orders',() => {
  it('imports supplier demand without physical reservation or integrity review',async () => {
    expect(await upsertOrderSnapshot(db,order())).toMatchObject({needs_review:false});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:0,reserved:0,available:0,unallocated_demand:20,uncovered_demand:20});
    expect((await getOrder(db,'ORDER')).lines[0]).toMatchObject({item_id:'supplier',fulfillment_state:'AWAITING_SUPPLIER'});
  });

  it('allows local lines in a mixed Paid order while keeping supplier lines incomplete',async () => {
    await upsertOrderSnapshot(db,{...order(),items:[{...supplierLine},{...stockLine}]});
    await pick(2,'local','l');
    const result = await getOrder(db,'ORDER');
    expect(result.needs_review).toBe(0);
    expect(result.lines.find(line => line.item_id==='supplier')).toMatchObject({remaining_qty:20,picked_qty:0,fulfillment_state:'AWAITING_SUPPLIER'});
    expect(result.lines.find(line => line.item_id==='local')).toMatchObject({remaining_qty:0,picked_qty:2,fulfillment_state:'PICKED'});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM outbox').get()).toEqual({n:0});
  });

  it('keeps unsupported mixed orders in whole-order review',async () => {
    await upsertOrderSnapshot(db,{...order(),items:[{...supplierLine},{...stockLine},{...stockLine,id:'unmapped',sku:'OTHER',productId:'300'}]});
    expect((await getOrder(db,'ORDER')).needs_review).toBe(1);
    await expect(pick(1,'local','l')).rejects.toMatchObject({code:'ITEM_NEEDS_REVIEW'});
  });

  it.each(['CANCELLED','REFUNDED','INCOMPLETE'])('releases unpicked %s allocations without creating shelf stock',async payment => {
    await upsertOrderSnapshot(db,order());await receive();await assign();
    await upsertOrderSnapshot(db,order(payment,2));
    await upsertOrderSnapshot(db,order(payment,2));
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:0,free:20,unallocated_demand:0});
    expect(sqlite.prepare("SELECT quantity FROM supplier_allocation_events WHERE type='CANCEL_RELEASE'").all()).toEqual([{quantity:20}]);
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM movements').get()).toEqual({n:1});
  });

  it('retains allocation on cancellation after partial picking without fabricating a return',async () => {
    await upsertOrderSnapshot(db,order());await receive();await assign();await pick(5);
    expect(await upsertOrderSnapshot(db,order('CANCELLED',2))).toMatchObject({needs_review:true});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:15,allocated:15,free:0});
    expect((await getOrder(db,'ORDER')).lines[0]).toMatchObject({picked_qty:5,fulfillment_state:'REVIEW'});
    expect(sqlite.prepare("SELECT COUNT(*) AS n FROM supplier_allocation_events WHERE type='CANCEL_RELEASE'").get()).toEqual({n:0});
  });

  it('retains supplier allocation when a different line in the mixed order was picked before cancellation',async () => {
    const mixed = {...order(),items:[{...supplierLine},{...stockLine}]};
    await upsertOrderSnapshot(db,mixed);await receive();await assign();await pick(2,'local','l');
    await upsertOrderSnapshot(db,{...mixed,paymentStatus:'CANCELLED',updatedAt:order('CANCELLED',2).updatedAt});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:20});
    expect((await getOrder(db,'ORDER')).needs_review).toBe(1);
  });

  it.each(['SHIPPED','READY_FOR_PICKUP'])('holds %s with missing supplier picks without synthetic movements',async fulfillmentStatus => {
    await upsertOrderSnapshot(db,order());await receive();await assign();
    expect(await upsertOrderSnapshot(db,{...order('PAID',2),fulfillmentStatus})).toMatchObject({needs_review:true});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:20});
    expect((await getOrder(db,'ORDER')).lines[0]).toMatchObject({picked_qty:0,fulfillment_state:'REVIEW'});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM movements').get()).toEqual({n:1});
  });

  it('serializes cancellation races with allocation and with picking',async () => {
    await upsertOrderSnapshot(db,order());await receive();
    await Promise.allSettled([assign(),upsertOrderSnapshot(db,order('CANCELLED',2))]);
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:20,allocated:0,free:20});
    // A fresh order can legitimately become Paid again; no old allocation is revived.
    await upsertOrderSnapshot(db,order('PAID',3));await assign();
    await Promise.allSettled([pick(5),upsertOrderSnapshot(db,order('CANCELLED',4))]);
    const item = await getItem(db,'supplier');const current = await getOrder(db,'ORDER');
    if (current.lines[0].picked_qty===0) expect(item).toMatchObject({on_hand:20,allocated:0});
    else expect(item).toMatchObject({on_hand:15,allocated:15});
    expect(item.on_hand).toBe(20-current.lines[0].picked_qty);
  });
});

describe('supplier catalogue policy and outbound fail-closed behavior',() => {
  it('accepts explicit unlimited policy without changing physical count or mode',async () => {
    webhook();await receive(5);
    const fetcher = vi.fn().mockResolvedValue(Response.json(product()));
    await processWebhook(env,'product',fetcher);
    expect(await getItem(db,'supplier')).toMatchObject({inventory_mode:'SUPPLIER_BACKED_UNLIMITED',on_hand:5,last_ecwid_quantity:null});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM sync_issues').get()).toEqual({n:0});
    expect(fetcher.mock.calls[0][1]).toMatchObject({method:'GET'});
  });

  it('flags unexpected finite stock for a supplier target without automatically switching it',async () => {
    await upsertOrderSnapshot(db,order());
    webhook();await processWebhook(env,'product',vi.fn().mockResolvedValue(Response.json(product({unlimited:false,quantity:4797}))));
    expect(sqlite.prepare("SELECT kind,status FROM sync_issues WHERE item_id='supplier'").get()).toEqual({kind:'PRODUCT_REVIEW',status:'OPEN'});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:0,inventory_mode:'SUPPLIER_BACKED_UNLIMITED'});
    await expect(receive()).rejects.toMatchObject({code:'ITEM_NEEDS_REVIEW'});
    expect((await getOrder(db,'ORDER')).lines[0]).toMatchObject({pickable_qty:0,fulfillment_state:'REVIEW'});
  });

  it('does not automatically classify a stock-limited target as supplier-backed when Ecwid becomes unlimited',async () => {
    webhook('local-product','200');
    await processWebhook(env,'local-product',vi.fn().mockResolvedValue(Response.json(product({id:200,sku:'LOCAL-BOLT'}))));
    expect(await getItem(db,'local')).toMatchObject({inventory_mode:'STOCK_LIMITED',on_hand:10});
    expect(sqlite.prepare("SELECT kind FROM sync_issues WHERE item_id='local'").get()).toEqual({kind:'PRODUCT_REVIEW'});
  });

  it('preserves exact variation and option identity for unlimited supplier targets',async () => {
    sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,inventory_mode,supplier_name,active)
      VALUES('variation','SUP-M5','Supplier M5','SUP-M5','300','501','[{"name":"Thread","value":"M5"}]','SUPPLIER_BACKED_UNLIMITED','MISUMI',0)`).run();
    sqlite.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
      VALUES('variant-opening','variation',0,'verified fixture','admin',?)`).run(stamp);
    sqlite.prepare("UPDATE items SET active=1 WHERE id='variation'").run();
    const variant = product({id:300,sku:'PARENT',options:[{name:'Thread',type:'SELECT',choices:[{text:'M5'},{text:'M6'}]}],
      combinations:[{id:501,sku:'SUP-M5',unlimited:true,options:[{name:'Thread',value:'M5'}]}]});
    webhook('variation','300');await processWebhook(env,'variation',vi.fn().mockResolvedValue(Response.json(variant)));
    expect(sqlite.prepare("SELECT COUNT(*) AS n FROM sync_issues WHERE item_id='variation'").get()).toEqual({n:0});
    webhook('changed-variation','300');
    const changed = {...variant,combinations:[{id:501,sku:'SUP-M5',unlimited:true,options:[{name:'Thread',value:'M6'}]}]};
    await processWebhook(env,'changed-variation',vi.fn().mockResolvedValue(Response.json(changed)));
    expect(sqlite.prepare("SELECT kind FROM sync_issues WHERE item_id='variation'").get()).toEqual({kind:'PRODUCT_REVIEW'});
  });

  it.each(['demo','live'])('blocks a malformed supplier outbox row in %s without a remote call',async mode => {
    // Simulates legacy corruption by bypassing only the insertion defense in
    // this isolated test DB. Runtime claims and processors must still refuse it.
    sqlite.exec('DROP TRIGGER supplier_outbox_guard');
    const receipt = await receive(5);
    sqlite.prepare(`INSERT INTO outbox(id,item_id,ecwid_product_id,quantity_delta,created_at,updated_at)
      VALUES(?,'supplier','100',5,?,?)`).run(receipt.movement.id,stamp,stamp);
    const fetcher = vi.fn();env.ECWID_MODE=mode;
    expect(await claimOutbox(db,receipt.movement.id)).toBeNull();
    await processOutbox(env,receipt.movement.id,fetcher);
    expect(fetcher).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT status FROM outbox').get()).toEqual({status:'BLOCKED'});
    expect(await getItem(db,'supplier')).toMatchObject({on_hand:5,last_ecwid_quantity:null});
  });

  it('rejects a supplier movement outbox even if its item pointer was corrupted to a stock-limited item',async () => {
    sqlite.exec('DROP TRIGGER supplier_outbox_guard');
    const receipt = await receive(5);
    sqlite.prepare(`INSERT INTO outbox(id,item_id,ecwid_product_id,quantity_delta,created_at,updated_at)
      VALUES(?,'local','200',5,?,?)`).run(receipt.movement.id,stamp,stamp);
    const fetcher = vi.fn();
    expect(await claimOutbox(db,receipt.movement.id)).toBeNull();
    await processOutbox(env,receipt.movement.id,fetcher);
    expect(fetcher).not.toHaveBeenCalled();
    expect(sqlite.prepare('SELECT status FROM outbox').get()).toEqual({status:'BLOCKED'});
    expect(await getItem(db,'local')).toMatchObject({on_hand:10,last_ecwid_quantity:10});
  });
});
