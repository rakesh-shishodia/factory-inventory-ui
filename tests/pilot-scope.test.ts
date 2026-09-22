import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { registerWorkbookTargets, validateWorkbookTargets, workbookTargetId, type WorkbookTarget } from '../src/pilot-scope';
import { createMovement, getOrder, listOrders } from '../src/inventory';
import { createAllocation } from '../src/supplier';
import { ingestWebhook, pollOrders, processWebhook, refreshOrder, upsertOrderSnapshot, type SyncEnv } from '../src/sync';
import type { EcwidOrder } from '../src/ecwid';
import { applyMigrations, sqliteD1 } from './d1';

let sqlite: Database.Database;
let db: D1Database;
const review = {reference:'reviewed-pilot-exclusions',actor:'admin@example.test',timestamp:'2026-09-22T08:00:00.000Z'};
const target: WorkbookTarget = {ecwid_product_id:'200',ecwid_combination_id:null,ecwid_option_signature:'[]',sku:'EXTERNAL',name:'Workbook item'};
function order(revision = 1): EcwidOrder {
  return {id:'ORDER',paymentStatus:'PAID',fulfillmentStatus:'AWAITING_PROCESSING',updatedAt:`2026-09-22T08:00:0${revision}.000Z`,items:[
    {id:'app',productId:'100',combinationId:null,sku:'APP',name:'App item',quantity:2,selectedOptions:[],digital:false,trackQuantity:true},
    {id:'wb',productId:'200',combinationId:null,sku:'EXTERNAL',name:'Workbook item',quantity:3,selectedOptions:[],digital:false,trackQuantity:true},
  ]};
}
async function pick() {
  return createMovement(db,{operation_id:crypto.randomUUID(),type:'ECWID_PICK',item_id:'app',quantity:2,order_id:'ORDER',order_line_id:'ORDER:app'},review.actor);
}
beforeEach(() => {
  sqlite = new Database(':memory:'); applyMigrations(sqlite); db=sqliteD1(sqlite);
  sqlite.exec(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand) VALUES('app','APP','App item','APP','100',10)`);
});
afterEach(() => {sqlite.close(); vi.restoreAllMocks();});

describe('reviewed workbook target registry', () => {
  it('normalizes reviewed SKU and keeps exact canonical option identity', () => {
    expect(validateWorkbookTargets([{...target,sku:' external '}])).toEqual([target]);
    expect(workbookTargetId(target)).toBe('workbook:200:simple');
    expect(validateWorkbookTargets([{...target,ecwid_option_signature:'[{"name":"Cut length","value":"1.5m"}]'}])).toHaveLength(1);
  });
  it.each([null,{},[{}],[{...target,sku:''}],[{...target,ecwid_combination_id:''}],
    [{...target,ecwid_option_signature:'null'}],[{...target,ecwid_option_signature:'[{}]'}],
    [{...target,ecwid_option_signature:'[ {"name":"Size","value":"M3"} ]'}],[target,target],
    [target,{...target,ecwid_product_id:'201'}],Array(501).fill(target)])('rejects invalid or conflicting review input %j', value => {
    expect(() => validateWorkbookTargets(value)).toThrow();
  });
  it('requires administrator review metadata and makes exact retries idempotent', async () => {
    await expect(registerWorkbookTargets(db,[target],{...review,reference:''})).rejects.toMatchObject({code:'WORKBOOK_REVIEW_REQUIRED'});
    await registerWorkbookTargets(db,[target],review);
    await registerWorkbookTargets(db,[target],{...review,timestamp:'2026-09-23'});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM workbook_managed_targets').get()).toEqual({n:1});
    await expect(registerWorkbookTargets(db,[{...target,name:'Changed'}],review)).rejects.toThrow('WORKBOOK_TARGET_CONFLICT');
    expect(() => sqlite.exec("UPDATE workbook_managed_targets SET name='Changed'")).toThrow('WORKBOOK_TARGET_IMMUTABLE');
    expect(() => sqlite.exec('DELETE FROM workbook_managed_targets')).toThrow('WORKBOOK_TARGET_IMMUTABLE');
  });
  it.each([{...target,sku:'APP'},{...target,ecwid_product_id:'100'}])('rejects existing app SKU or exact-target collision', async row => {
    await expect(registerWorkbookTargets(db,[row],review)).rejects.toThrow('WORKBOOK_APP_IDENTITY_CONFLICT');
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM workbook_managed_targets').get()).toEqual({n:0});
  });
  it('rolls back the whole registry batch on a later identity collision', async () => {
    await expect(registerWorkbookTargets(db,[target,{...target,ecwid_product_id:'100',sku:'CONFLICT'}],review)).rejects.toThrow('WORKBOOK_APP_IDENTITY_CONFLICT');
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM workbook_managed_targets').get()).toEqual({n:0});
  });
  it('also rejects future app identity collisions while permitting independent siblings', async () => {
    await registerWorkbookTargets(db,[target],review);
    expect(() => sqlite.exec("INSERT INTO items(id,sku,name,scan_code,ecwid_product_id) VALUES('bad','EXTERNAL','bad','bad','300')")).toThrow('WORKBOOK_APP_IDENTITY_CONFLICT');
    expect(() => sqlite.exec("UPDATE items SET ecwid_product_id='200' WHERE id='app'")).toThrow('WORKBOOK_APP_IDENTITY_CONFLICT');
    expect(() => sqlite.exec(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,ecwid_combination_id,ecwid_option_signature)
      VALUES('sibling','SIBLING','sibling','SIBLING','200','20','[{"name":"Size","value":"M3"}]')`)).not.toThrow();
  });
});

describe('mixed app and workbook orders', () => {
  beforeEach(async () => {await registerWorkbookTargets(db,[target],review);});
  it('permits Paid app picks without reserving workbook stock or marking the whole order complete', async () => {
    expect(await upsertOrderSnapshot(db,order())).toMatchObject({needs_review:false});
    const current = await getOrder(db,'ORDER');
    expect(current.lines.find(line => line.sku==='EXTERNAL')).toMatchObject({management_mode:'WORKBOOK',item_id:null,
      picked_qty:0,pickable_qty:0,fulfillment_state:'WORKBOOK_MANAGED'});
    expect(await listOrders(db)).toHaveLength(1);
    await pick();
    expect(await listOrders(db)).toHaveLength(0);
    const after = await getOrder(db,'ORDER');
    expect(after.fulfillment_status).toBe('AWAITING_PROCESSING');
    expect(after.lines.find(line => line.sku==='EXTERNAL')).toMatchObject({picked_qty:0,remaining_qty:3});
    expect(sqlite.prepare('SELECT on_hand FROM items').get()).toEqual({on_hand:8});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM outbox').get()).toEqual({n:0});
  });
  it('still blocks Awaiting Payment picking in mixed orders', async () => {
    await upsertOrderSnapshot(db,{...order(),paymentStatus:'AWAITING_PAYMENT'});
    await expect(pick()).rejects.toMatchObject({code:'ORDER_NOT_PICKABLE'});
    await upsertOrderSnapshot(db,order(2));
    await expect(pick()).resolves.toMatchObject({sync_status:'NOT_REQUIRED'});
  });
  it.each(['SHIPPED','READY_FOR_PICKUP'])('accepts external %s after every APP pick, without a workbook pseudo-pick', async fulfillmentStatus => {
    await upsertOrderSnapshot(db,order()); await pick();
    expect(await upsertOrderSnapshot(db,{...order(2),fulfillmentStatus})).toMatchObject({needs_review:false});
    expect(sqlite.prepare("SELECT picked_qty FROM order_lines WHERE management_mode='WORKBOOK'").get()).toEqual({picked_qty:0});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM movements').get()).toEqual({n:1});
  });
  it('still quarantines terminal orders with missing APP picks', async () => {
    await upsertOrderSnapshot(db,order());
    expect(await upsertOrderSnapshot(db,{...order(2),fulfillmentStatus:'READY_FOR_PICKUP'})).toMatchObject({needs_review:true});
  });
  it('keeps an exactly reviewed workbook line external even when Ecwid reports downloadable attachments',async()=>{
    const input=order();input.items[1].digital=true;
    expect(await upsertOrderSnapshot(db,input)).toMatchObject({needs_review:false});
    expect((await getOrder(db,'ORDER')).lines.find(row=>row.id==='ORDER:wb')).toMatchObject({management_mode:'WORKBOOK',item_id:null,pickable_qty:0});
  });
  it.each(['wrong-sku','wrong-parent','wrong-combination','wrong-options','nonboolean-digital','unknown'])('never exempts a %s line by absence from the app', async fault => {
    const input = order(); const line = input.items[1];
    if (fault==='wrong-sku') line.sku='APP';
    if (fault==='wrong-parent') line.productId='100';
    if (fault==='wrong-combination') line.combinationId='99';
    if (fault==='wrong-options') line.selectedOptions=[{name:'Size',value:'M3'}];
    if (fault==='nonboolean-digital') Object.assign(line,{digital:'true'});
    if (fault==='unknown') {line.productId='300';line.sku='UNREVIEWED';}
    expect(await upsertOrderSnapshot(db,input)).toMatchObject({needs_review:true});
    expect((await getOrder(db,'ORDER')).lines.find(row=>row.id==='ORDER:wb')).toMatchObject({management_mode:'APP',item_id:null,fulfillment_state:'REVIEW'});
    await expect(pick()).rejects.toMatchObject({code:'ITEM_NEEDS_REVIEW'});
  });
  it('holds edits to a workbook line for review rather than silently reclassifying it', async () => {
    await upsertOrderSnapshot(db,order());
    const changed=order(2);changed.items[1].quantity=4;
    expect(await upsertOrderSnapshot(db,changed)).toMatchObject({needs_review:true});
    expect(sqlite.prepare("SELECT ordered_qty FROM order_lines WHERE management_mode='WORKBOOK'").get()).toEqual({ordered_qty:3});
  });
  it('enforces outside-app handling even against direct SQL mutation or wrong movement references', async () => {
    await upsertOrderSnapshot(db,order());
    for (const assignment of ["picked_qty=1","item_id='app'","sku='APP'","ordered_qty=4"]) {
      expect(() => sqlite.exec(`UPDATE order_lines SET ${assignment} WHERE management_mode='WORKBOOK'`)).toThrow('WORKBOOK_LINE_HANDLED_EXTERNALLY');
    }
    expect(() => sqlite.exec("UPDATE order_lines SET management_mode='APP' WHERE management_mode='WORKBOOK'")).toThrow('ORDER_LINE_MANAGEMENT_IMMUTABLE');
    await expect(createMovement(db,{operation_id:crypto.randomUUID(),type:'ECWID_PICK',item_id:'app',quantity:1,
      order_id:'ORDER',order_line_id:'ORDER:wb'},review.actor)).rejects.toMatchObject({code:'WORKBOOK_LINE_HANDLED_EXTERNALLY'});
  });
  it('permits supplier allocation in known workbook mixed orders but not unknown mixed orders', async () => {
    sqlite.exec(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,inventory_mode,supplier_name,active)
      VALUES('supplier','SUPPLIER','Supplier','SUPPLIER','400','SUPPLIER_BACKED_UNLIMITED','Demo',0);
      INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at) VALUES('opening','supplier',5,'verified','admin','2026-09-22');
      UPDATE items SET active=1 WHERE id='supplier';`);
    const input=order();input.items[0]={...input.items[0],productId:'400',sku:'SUPPLIER'};
    await upsertOrderSnapshot(db,input);
    await expect(createAllocation(db,{operation_id:crypto.randomUUID(),type:'ALLOCATE',item_id:'supplier',quantity:2,
      order_id:'ORDER',order_line_id:'ORDER:app'},review.actor)).resolves.toMatchObject({duplicate:false});
    await createMovement(db,{operation_id:crypto.randomUUID(),type:'ECWID_PICK',item_id:'supplier',quantity:2,
      order_id:'ORDER',order_line_id:'ORDER:app'},review.actor);
    expect(await upsertOrderSnapshot(db,{...input,updatedAt:order(2).updatedAt,fulfillmentStatus:'SHIPPED'})).toMatchObject({needs_review:false});
  });
});

describe('order-sync cutover safety switch', () => {
  it.each([undefined,'false','TRUE','1',''])('does no live ingestion, fetch, polling or event claim when flag=%s', async flag => {
    const env: SyncEnv={DB:db,ECWID_MODE:'live',LIVE_SYNC_ENABLED:'true',ORDER_SYNC_ENABLED:flag,ECWID_STORE_ID:'123',
      ECWID_TOKEN:'test',ECWID_CLIENT_SECRET:'test',SYNC_QUEUE:{send:vi.fn()} as unknown as Queue};
    const fetcher=vi.fn();
    await expect(refreshOrder(env,'ORDER',fetcher)).rejects.toMatchObject({code:'ORDER_SYNC_PAUSED'});
    await expect(pollOrders(env,fetcher)).rejects.toMatchObject({code:'ORDER_SYNC_PAUSED'});
    const response=await ingestWebhook(new Request('https://example.test/api/webhooks/ecwid',{method:'POST',body:'{}'}),env);
    expect(response.status).toBe(503);
    sqlite.exec(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,received_at,updated_at)
      VALUES('event','order.updated','ORDER','123','{}','2026-09-22','2026-09-22')`);
    await processWebhook(env,'event',fetcher);
    expect(sqlite.prepare('SELECT status,attempts FROM webhook_events').get()).toEqual({status:'PENDING',attempts:0});
    expect(sqlite.prepare('SELECT COUNT(*) AS n FROM sync_state').get()).toEqual({n:0});
    expect(fetcher).not.toHaveBeenCalled();expect(env.SYNC_QUEUE.send).not.toHaveBeenCalled();
  });
});
