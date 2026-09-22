import Database from 'better-sqlite3';
import { readFileSync, readdirSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { canonicalVariationOptions, parseOrder, type EcwidOrder, type EcwidProduct } from '../src/ecwid';
import { opaqueOptionsHash, sanitizedOpeningOrder, stockOptionSubset, workbookLineMatches } from '../src/workbook-identity';
import { registerWorkbookTargets, validateWorkbookTargets, type WorkbookTarget } from '../src/pilot-scope';
import { orderSnapshotHash, upsertOrderSnapshot } from '../src/sync';
import { previewOpeningCutover, stageOpeningCutover, type OpeningCutoverRequest } from '../src/opening-cutover';
import { beginCutoverAlignment, alignCutoverRow, finishAndActivateCutover, type CutoverAlignmentPolicy } from '../src/cutover-alignment';
import { applyMigrations, sqliteD1 } from './d1';

const at='2026-09-22T12:01:00.000Z';
const policy:CutoverAlignmentPolicy={storeId:'2442119',actor:'admin@example.test',token:'test-only',mode:'live',
  inventoryEnabled:'false',liveSyncEnabled:'false',orderSyncEnabled:'false',now:at};
const stock=[{name:'Length',value:'4000'}];
const cutSecret='Private custom cutting instructions and contact details';
const fileSecret='https://private.example/customer-file';
const target:WorkbookTarget={ecwid_product_id:'200',ecwid_combination_id:'201',sku:'PROFILE',name:'Workbook profile',
  ecwid_option_signature:JSON.stringify(stock),sku_source:'PARENT_IF_VARIATION_BLANK',option_policy:'STOCK_SELECTION_PLUS_OPAQUE_EXTRAS'};
const review={reference:'Explicit workbook-only approval',actor:policy.actor,timestamp:at};
const app:EcwidProduct={id:'100',sku:'APP',name:'Pilot app item',quantity:12,unlimited:false,enabled:true,hasOptions:false,
  hasVariations:false,combinationId:null,variationOptions:[],hasExtraOptions:false,hasBundleRelationships:false,eligibilityVerified:true};
function wbCatalog(ownSku=false):EcwidProduct {
  return {...app,id:'200',sku:ownSku?'ASSEMBLY':'',name:target.name,combinationId:'201',variationOptions:stock,
    quantity:null,unlimited:true,hasOptions:true,hasExtraOptions:true,eligibilityVerified:false};
}
function rawOrder(extraChoices=false):EcwidOrder {
  return {id:'12',createdAt:'2026-09-22T11:00:00.000Z',updatedAt:'2026-09-22T11:30:00.000Z',paymentStatus:'PAID',fulfillmentStatus:'AWAITING_PROCESSING',items:[
    {id:'1',productId:'100',combinationId:null,sku:'APP',name:app.name,quantity:3,digital:false,trackQuantity:true,selectedOptions:[]},
    {id:'2',productId:'200',combinationId:'201',sku:extraChoices?'ASSEMBLY':'PROFILE',name:target.name,quantity:2,digital:false,trackQuantity:false,
      selectedOptions:[{...stock[0],type:'CHOICE'}, {name:'End Tapping',value:'Yes',type:'CHOICE'},
        ...(extraChoices?[{name:'Motor',value:'NEMA23',type:'CHOICE'}]:[
          {name:'Cut Details',value:cutSecret,type:'TEXT'}, {name:'Drawing',value:'File',type:'FILE',files:[{url:fileSecret}]}])]},
  ]};
}
async function opening(ownSku=false):Promise<{request:OpeningCutoverRequest;raw:EcwidOrder}> {
  const raw=rawOrder(ownSku); const wb=wbCatalog(ownSku);
  const parent={...wb,combinationId:null,sku:ownSku?'PARENT-ASSEMBLY':'PROFILE',hasVariations:true,variationOptions:[],variations:[wb]};
  return {raw,request:{operation_id:crypto.randomUUID(),expected_hash:'',confirm_staging:true,physical_counts_confirmed:true,
    freeze:{confirmed:true,started_at:'2026-09-22T12:00:00.000Z'},input:{store_id:policy.storeId,source_ref:'Fresh verified workbook',
      snapshot_source_hash:'a'.repeat(64),balance_meaning:'PHYSICAL_ON_HAND',reservations_confirmed:true,
      rows:[{sku:'APP',name:app.name,balance:10,single_unit_confirmed:true}],reservations:[{sku:'APP',quantity:3}],
      catalog:{kind:'READONLY_CATALOGUE',schema_version:1,dry_run:true,complete:true,store_id:policy.storeId,
        started_at:'2026-09-22T12:00:01.000Z',completed_at:'2026-09-22T12:00:30.000Z',product_count:2,stock_target_count:2,
        products:[app,parent],stock_targets:[app,wb],reservations_confirmed:false}},
    scope:[{sku:'APP',ecwid_product_id:'100',ecwid_combination_id:null,ecwid_option_signature:'[]'}],
    orders:{kind:'READONLY_ORDERS',schema_version:2,dry_run:true,complete:true,store_id:policy.storeId,
      started_at:'2026-09-22T12:00:01.000Z',completed_at:'2026-09-22T12:00:30.000Z',creation_cutoff:Date.parse('2026-09-22T12:00:01.000Z')/1000,
      orders_checked:1,pending_order_count:1,line_count:2,orders:[await sanitizedOpeningOrder(raw,[app,wb])]},
    line_confirmations:[{order_id:'12',ecwid_line_id:'1',previously_picked_quantity:0},{order_id:'12',ecwid_line_id:'2',previously_picked_quantity:0}],
    workbook_scope:[ownSku?{...target,sku:'ASSEMBLY',sku_source:'TARGET'}:{...target}]}};
}
let sqlite:Database.Database;let db:D1Database;
beforeEach(()=>{sqlite=new Database(':memory:');applyMigrations(sqlite);db=sqliteD1(sqlite);});
afterEach(()=>{sqlite.close();vi.restoreAllMocks();});

describe('privacy-preserving workbook-only identities',()=>{
  it.each([false,true])('matches fresh raw and sanitized hashes, extra choices=%s',async ownSku=>{
    const {request,raw}=await opening(ownSku);const sanitized=request.orders.orders[0];
    expect(sanitized.items[1].selectedOptions).toEqual(stock);
    expect(sanitized.items[1].workbookOptionsEvidence).toMatchObject({kind:'OPAQUE_OPTIONS_V1',sha256:expect.stringMatching(/^[a-f0-9]{64}$/)});
    expect(JSON.stringify(sanitized)).not.toContain(cutSecret);expect(JSON.stringify(sanitized)).not.toContain(fileSecret);
    expect(JSON.stringify(sanitized)).not.toContain('End Tapping');expect(JSON.stringify(sanitized)).not.toContain('NEMA23');
    expect(await orderSnapshotHash(raw,{workbookTargets:request.workbook_scope}))
      .toBe(await orderSnapshotHash(sanitized,{workbookTargets:request.workbook_scope,allowSnapshotEvidence:true}));
    expect(workbookLineMatches(raw.items[1],request.workbook_scope[0])).toBe(true);
    expect(workbookLineMatches(sanitized.items[1],request.workbook_scope[0],true)).toBe(true);
  });
  it('keeps legacy supported APP hash material unchanged',async()=>{
    const order=rawOrder();order.items=[order.items[0]];
    const tuple=order.items.map(l=>[l.id,l.productId,l.sku,l.quantity,l.combinationId,canonicalVariationOptions(l.selectedOptions),l.digital]);
    const old=Buffer.from(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(tuple)))).toString('hex');
    expect(await orderSnapshotHash(order)).toBe(old);
  });
  it('canonicalizes supported choice metadata consistently, but detects changed free-text or extra choices',async()=>{
    expect(await opaqueOptionsHash([{name:'A',value:'B',type:'CHOICE'}])).toBe(await opaqueOptionsHash([{value:'B',name:'A'}]));
    const raw=rawOrder();const prior=await orderSnapshotHash(raw,{workbookTargets:[target]});
    (raw.items[1].selectedOptions[2] as Record<string,unknown>).value='Different cutting instructions';
    expect(await orderSnapshotHash(raw,{workbookTargets:[target]})).not.toBe(prior);
    const choices=rawOrder(true);const own={...target,sku:'ASSEMBLY',sku_source:'TARGET' as const};
    const previous=await orderSnapshotHash(choices,{workbookTargets:[own]});
    (choices.items[1].selectedOptions[1] as Record<string,unknown>).value='No';
    expect(await orderSnapshotHash(choices,{workbookTargets:[own]})).not.toBe(previous);
  });
  it.each(['missing','duplicate','changed','nonboolean-digital','wrong-parent','wrong-combination','wrong-sku'])('fails exact stock identity on %s',fault=>{
    const line=rawOrder().items[1];
    if(fault==='missing')line.selectedOptions.shift();if(fault==='duplicate')line.selectedOptions.push(stock[0]);
    if(fault==='changed')line.selectedOptions[0]={name:'Length',value:'20000'};
    if(fault==='nonboolean-digital')Object.assign(line,{digital:'true'});if(fault==='wrong-parent')line.productId='999';
    if(fault==='wrong-combination')line.combinationId='202';if(fault==='wrong-sku')line.sku='OTHER';
    expect(workbookLineMatches(line,target)).toBe(false);
  });
  it('rejects untrusted redacted evidence in live ingestion and ignores wire-supplied claims',async()=>{
    const {request}=await opening();const sanitized=request.orders.orders[0];
    await expect(orderSnapshotHash(sanitized)).rejects.toThrow('Untrusted redacted');
    await registerWorkbookTargets(db,[target],review);
    await expect(upsertOrderSnapshot(db,sanitized)).rejects.toThrow('Untrusted redacted');
    expect(sqlite.prepare('SELECT count(*) FROM orders').pluck().get()).toBe(0);
    const parsed=parseOrder({...sanitized,updateTimestamp:Date.parse(sanitized.updatedAt)/1000});
    expect(parsed.items[1].workbookOptionsEvidence).toBeUndefined();
    expect(await orderSnapshotHash(parsed,{workbookTargets:[target]})).not.toBe(await orderSnapshotHash(sanitized,{allowSnapshotEvidence:true,workbookTargets:[target]}));
  });
  it.each(['malformed-hash','extra-evidence-key','extra-choice','v1','app-claim','unreviewed'])('rejects unsafe opening claim %s',async fault=>{
    const {request}=await opening();const line=request.orders.orders[0].items[1];
    if(fault==='malformed-hash')line.workbookOptionsEvidence!.sha256='bad';
    if(fault==='extra-evidence-key')Object.assign(line.workbookOptionsEvidence!,{instructions:cutSecret});
    if(fault==='extra-choice')line.selectedOptions.push({name:'Extra',value:'Invalid'});
    if(fault==='v1')request.orders.schema_version=1;
    if(fault==='app-claim')request.orders.orders[0].items[0].workbookOptionsEvidence=line.workbookOptionsEvidence;
    if(fault==='unreviewed')request.workbook_scope=[];
    await expect(previewOpeningCutover(request,policy)).rejects.toBeInstanceOf(Error);
  });
  it('does not convert unsupported lines to an approval and cannot recover an old REDACTED sentinel',async()=>{
    const raw=rawOrder();raw.items[1].combinationId='999';
    const sanitized=await sanitizedOpeningOrder(raw,[wbCatalog()]);
    expect(sanitized.items[1].workbookOptionsEvidence).toBeUndefined();
    expect(canonicalVariationOptions(sanitized.items[1].selectedOptions)).toBeNull();
    expect(stockOptionSubset([{name:'Length',value:'4000',type:'REDACTED'}],stock)).toBeNull();
  });
});

describe('explicit workbook scope and exact catalogue proof',()=>{
  it('allows same-parent inherited display SKUs only under explicit paired policies',async()=>{
    const sibling={...target,ecwid_combination_id:'202',ecwid_option_signature:'[{"name":"Length","value":"20000"}]'};
    expect(validateWorkbookTargets([target,sibling])).toHaveLength(2);
    await registerWorkbookTargets(db,[target,sibling],review);
    expect(sqlite.prepare('SELECT count(*) FROM workbook_managed_targets').pluck().get()).toBe(2);
    await registerWorkbookTargets(db,[target,sibling],review);
    for(const row of [{...sibling,ecwid_product_id:'300'}, {...sibling,sku_source:'TARGET' as const},
      {...sibling,option_policy:'EXACT' as const}, {...sibling,ecwid_combination_id:null}]){
      expect(()=>validateWorkbookTargets([target,row])).toThrow();
    }
    await expect(registerWorkbookTargets(db,[{...sibling,ecwid_product_id:'300'}],review)).rejects.toThrow('WORKBOOK_TARGET_CONFLICT');
    expect(()=>sqlite.exec("INSERT INTO items(id,sku,name,scan_code,ecwid_product_id) VALUES('bad','PROFILE','Bad','BAD','300')")).toThrow('WORKBOOK_APP_IDENTITY_CONFLICT');
  });
  it.each([false,true])('accepts catalogue-proven workbook policy with own SKU=%s without app eligibility relaxation',async ownSku=>{
    const {request}=await opening(ownSku);const result=await previewOpeningCutover(request,policy);
    expect(result).toMatchObject({row_count:1,line_count:2,workbook_line_count:1});
  });
  it.each(['missing-policy','nonblank-child','parent-sku','parent-not-variant','stock-option','own-sku-mismatch'])('rejects incomplete catalogue proof %s',async fault=>{
    const {request}=await opening(fault==='own-sku-mismatch');
    const catalog=request.input.catalog as {products:EcwidProduct[];stock_targets:EcwidProduct[]};
    if(fault==='missing-policy'){delete request.workbook_scope[0].sku_source;delete request.workbook_scope[0].option_policy;}
    if(fault==='nonblank-child')catalog.stock_targets[1].sku='SOMETHING';
    if(fault==='parent-sku')catalog.products[1].sku='WRONG';
    if(fault==='parent-not-variant')catalog.products[1].hasVariations=false;
    if(fault==='stock-option')catalog.stock_targets[1].variationOptions=[{name:'Length',value:'5000'}];
    if(fault==='own-sku-mismatch')catalog.stock_targets[1].sku='DIFFERENT';
    await expect(previewOpeningCutover(request,policy)).rejects.toBeInstanceOf(Error);
  });
});

describe('opening/live synchronization equivalence',()=>{
  it('preserves downloadable-attachment evidence on exact reviewed workbook lines, including opaque extras',async()=>{
    const {request,raw}=await opening(true);raw.items[1].digital=true;
    request.orders.orders=[await sanitizedOpeningOrder(raw,[app,wbCatalog(true)])];
    expect(request.orders.orders[0].items[1].digital).toBe(true);
    request.expected_hash=(await previewOpeningCutover(request,policy)).review_hash;
    await stageOpeningCutover(db,request,policy);
    expect(await upsertOrderSnapshot(db,raw)).toMatchObject({needs_review:false});
    expect(sqlite.prepare("SELECT management_mode,item_id,picked_qty FROM order_lines WHERE ecwid_line_id='2'").get())
      .toEqual({management_mode:'WORKBOOK',item_id:null,picked_qty:0});
    expect(JSON.parse(sqlite.prepare('SELECT orders_json FROM opening_cutover_batches').pluck().get() as string).orders[0].items[1].digital).toBe(true);
    raw.updatedAt=at;raw.items[1].digital=false;
    expect(await upsertOrderSnapshot(db,raw)).toMatchObject({needs_review:true});
  });
  it.each([undefined,null,'true',1])('rejects malformed explicit opening digital evidence %j',async digital=>{
    const {request}=await opening();Object.assign(request.orders.orders[0].items[1],{digital});
    await expect(previewOpeningCutover(request,policy)).rejects.toMatchObject({code:'INVALID_OPENING_ORDER_LINE'});
  });
  it.each([null,'true',1])('rejects a malformed live wire digital flag %j',digital=>{
    const raw=rawOrder();Object.assign(raw.items[1],{digital});
    expect(()=>parseOrder({...raw,updateTimestamp:Date.parse(raw.updatedAt)/1000})).toThrow('Invalid Ecwid digital flag');
  });
  it('does not relax APP digital eligibility in either opening or live ingestion',async()=>{
    const {request,raw}=await opening();request.orders.orders[0].items[0].digital=true;
    await expect(previewOpeningCutover(request,policy)).rejects.toMatchObject({code:'OPENING_ORDER_LINE_UNMAPPED'});
    await registerWorkbookTargets(db,[target],review);
    sqlite.exec("INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand) VALUES('app','APP','App','APP','100',10)");
    raw.items[0].digital=true;
    expect(await upsertOrderSnapshot(db,raw)).toMatchObject({needs_review:true});
    expect(sqlite.prepare("SELECT management_mode,item_id FROM order_lines WHERE ecwid_line_id='1'").get()).toEqual({management_mode:'APP',item_id:null});
  });
  it('keeps two approved inherited-SKU variants separate through stage and raw live ingestion',async()=>{
    const {request,raw}=await opening();
    const siblingStock=[{name:'Length',value:'20000'}];
    const sibling={...target,ecwid_combination_id:'202',ecwid_option_signature:JSON.stringify(siblingStock)};
    const siblingCatalog={...wbCatalog(),combinationId:'202',variationOptions:siblingStock};
    const catalog=request.input.catalog as {products:EcwidProduct[];stock_targets:EcwidProduct[];stock_target_count:number};
    catalog.stock_targets.push(siblingCatalog);catalog.stock_target_count++;
    catalog.products[1].variations!.push(siblingCatalog);
    raw.items.push({...structuredClone(raw.items[1]),id:'3',combinationId:'202',quantity:1,
      selectedOptions:[{...siblingStock[0],type:'CHOICE'},{name:'Cut Details',type:'TEXT',value:'Different external instructions'}]});
    request.workbook_scope.push(sibling);request.orders.line_count++;
    request.orders.orders=[await sanitizedOpeningOrder(raw,catalog.stock_targets)];
    request.line_confirmations.push({order_id:'12',ecwid_line_id:'3',previously_picked_quantity:0});
    request.expected_hash=(await previewOpeningCutover(request,policy)).review_hash;
    await stageOpeningCutover(db,request,policy);
    expect(await upsertOrderSnapshot(db,raw)).toMatchObject({needs_review:false});
    expect(sqlite.prepare("SELECT sku,workbook_target_id,ordered_qty FROM order_lines WHERE management_mode='WORKBOOK' ORDER BY ecwid_line_id").all())
      .toEqual([{sku:'PROFILE',workbook_target_id:'workbook:200:201',ordered_qty:2},{sku:'PROFILE',workbook_target_id:'workbook:200:202',ordered_qty:1}]);
    expect(sqlite.prepare('SELECT on_hand,reserved FROM item_stock').get()).toEqual({on_hand:10,reserved:3});
  });
  it.each([false,true])('stages opaque workbook lines and ingests the same raw snapshot without review, own SKU=%s',async ownSku=>{
    const {request,raw}=await opening(ownSku);request.expected_hash=(await previewOpeningCutover(request,policy)).review_hash;
    await stageOpeningCutover(db,request,policy);
    expect(await upsertOrderSnapshot(db,raw)).toMatchObject({needs_review:false});
    expect(sqlite.prepare('SELECT sku,management_mode,item_id,picked_qty FROM order_lines WHERE ecwid_line_id=?').get('2'))
      .toEqual({sku:ownSku?'ASSEMBLY':'PROFILE',management_mode:'WORKBOOK',item_id:null,picked_qty:0});
    expect(sqlite.prepare('SELECT on_hand,reserved FROM item_stock').get()).toEqual({on_hand:10,reserved:3});
    const persisted=JSON.stringify(sqlite.prepare('SELECT * FROM opening_cutover_batches').all());
    expect(persisted).not.toContain(cutSecret);expect(persisted).not.toContain(fileSecret);
    expect(sqlite.prepare('SELECT count(*) FROM movements').pluck().get()).toBe(0);
    expect(sqlite.prepare('SELECT count(*) FROM outbox').pluck().get()).toBe(0);
  });
  it.each(['quantity','text','choice','stock-selection','combination'])('quarantines later %s drift without changing the workbook line or app stock',async change=>{
    const {request,raw}=await opening();request.expected_hash=(await previewOpeningCutover(request,policy)).review_hash;
    await stageOpeningCutover(db,request,policy);raw.updatedAt=at;
    if(change==='quantity')raw.items[1].quantity++;
    if(change==='text')(raw.items[1].selectedOptions[2] as Record<string,unknown>).value='Changed';
    if(change==='choice')(raw.items[1].selectedOptions[1] as Record<string,unknown>).value='No';
    if(change==='stock-selection')raw.items[1].selectedOptions[0]={name:'Length',value:'20000',type:'CHOICE'};
    if(change==='combination')raw.items[1].combinationId='202';
    expect(await upsertOrderSnapshot(db,raw)).toMatchObject({needs_review:true});
    expect(sqlite.prepare("SELECT ordered_qty,picked_qty FROM order_lines WHERE management_mode='WORKBOOK'").get()).toEqual({ordered_qty:2,picked_qty:0});
    expect(sqlite.prepare('SELECT on_hand FROM items').pluck().get()).toBe(10);
  });
  it.each([false,true])('checks the full raw live hash before alignment; tampered evidence=%s',async tampered=>{
    const {request,raw}=await opening(true);
    raw.items[1].digital=true;request.orders.orders[0].items[1].digital=true;
    if(tampered)request.orders.orders[0].items[1].workbookOptionsEvidence!.sha256='0'.repeat(64);
    request.expected_hash=(await previewOpeningCutover(request,policy)).review_hash;await stageOpeningCutover(db,request,policy);
    let quantity=12;
    const fetcher=vi.fn(async(input:string|URL|Request,init?:RequestInit)=>{
      const url=new URL(String(input));
      if(url.pathname.endsWith('/orders'))return Response.json({total:1,count:1,offset:0,items:[{...raw,updateTimestamp:Date.parse(raw.updatedAt)/1000,createTimestamp:Date.parse(raw.createdAt!)/1000}]});
      expect(url.pathname).toBe('/api/v3/2442119/products/100');
      if(init?.method==='PUT'){quantity=JSON.parse(String(init.body)).quantity;return Response.json({updateCount:1});}
      return Response.json({id:100,sku:'APP',name:app.name,quantity,unlimited:false,enabled:true,options:[],combinations:[]});
    });
    const value={operation_id:request.operation_id,expected_hash:request.expected_hash,freeze:request.freeze};
    if(tampered){await expect(beginCutoverAlignment(db,value,policy,{fetcher})).rejects.toBeInstanceOf(Error);
      expect(sqlite.prepare('SELECT state FROM opening_cutover_batches').pluck().get()).toBe('REVIEW');
      expect(fetcher.mock.calls.filter(([,init])=>init?.method==='PUT')).toHaveLength(0);return;}
    await beginCutoverAlignment(db,value,policy,{fetcher});
    const item_id=sqlite.prepare('SELECT item_id FROM opening_cutover_rows').pluck().get() as string;
    await alignCutoverRow(db,{...value,item_id},policy,{fetcher});
    expect(await finishAndActivateCutover(db,value,policy,{fetcher})).toMatchObject({state:'ACTIVE'});
    expect(quantity).toBe(7);expect(fetcher.mock.calls.filter(([,init])=>init?.method==='PUT')).toHaveLength(1);
    expect(sqlite.prepare('SELECT count(*) FROM items').pluck().get()).toBe(1);
  });
});

describe('populated immutable registry migration',()=>{
  function beforeMigration(){
    sqlite.close();sqlite=new Database(':memory:');sqlite.pragma('foreign_keys=ON');
    const directory=new NodeURL('../migrations/',import.meta.url);
    for(const file of readdirSync(directory).filter(name=>/^000[1-7].*\.sql$/.test(name)).sort())
      sqlite.transaction(()=>sqlite.exec(readFileSync(new NodeURL(file,directory),'utf8')))();
    sqlite.exec(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id,on_hand) VALUES('app','APP','App','APP','100',10);
      INSERT INTO workbook_managed_targets VALUES('workbook:200:simple','200',NULL,'[]','EXTERNAL','External','approved','admin','2026-09-22');
      INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at) VALUES('order','PAID','AWAITING_PROCESSING','old','old');
      INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty,picked_qty,management_mode,workbook_target_id)
        VALUES('line','order','line',NULL,'EXTERNAL','External',3,0,'WORKBOOK','workbook:200:simple');`);
    return readFileSync(new NodeURL('0008_workbook_opaque_options.sql',directory),'utf8');
  }
  it('preserves populated line FKs and review identity while retaining immutable/no-pick guards',()=>{
    const sql=beforeMigration();const original=sqlite.prepare('SELECT * FROM workbook_managed_targets').get();
    sqlite.transaction(()=>sqlite.exec(sql))();
    expect(sqlite.prepare('SELECT * FROM workbook_managed_targets').get()).toEqual({...original as object,sku_source:'TARGET',option_policy:'EXACT'});
    expect(sqlite.pragma('foreign_key_check')).toEqual([]);
    expect((sqlite.pragma('foreign_key_list(order_lines)') as {table:string}[]).some(row=>row.table==='workbook_managed_targets')).toBe(true);
    expect(()=>sqlite.exec("UPDATE workbook_managed_targets SET name='changed'")).toThrow('WORKBOOK_TARGET_IMMUTABLE');
    expect(()=>sqlite.exec('DELETE FROM workbook_managed_targets')).toThrow('WORKBOOK_TARGET_IMMUTABLE');
    expect(()=>sqlite.exec("UPDATE order_lines SET picked_qty=1 WHERE id='line'")).toThrow('WORKBOOK_LINE_HANDLED_EXTERNALLY');
    expect(()=>sqlite.exec("UPDATE items SET sku='EXTERNAL' WHERE id='app'")).toThrow('WORKBOOK_APP_IDENTITY_CONFLICT');
    expect(sqlite.prepare('SELECT workbook_target_id FROM order_lines').pluck().get()).toBe('workbook:200:simple');
  });
  it('rolls back a failed populated upgrade including old schema, unique constraint, triggers and rows',()=>{
    const sql=beforeMigration();const previous=sqlite.prepare("SELECT type,name,sql FROM sqlite_master WHERE name NOT LIKE 'sqlite_%' ORDER BY name").all();
    expect(()=>sqlite.transaction(()=>{sqlite.exec(sql);throw new Error('Injected failure after migration');})()).toThrow('Injected failure');
    expect(sqlite.prepare("SELECT type,name,sql FROM sqlite_master WHERE name NOT LIKE 'sqlite_%' ORDER BY name").all()).toEqual(previous);
    expect(sqlite.pragma('foreign_key_check')).toEqual([]);
    expect(()=>sqlite.exec("INSERT INTO workbook_managed_targets VALUES('second','201',NULL,'[]','EXTERNAL','Other','approved','admin','2026-09-22')")).toThrow();
    expect(()=>sqlite.exec('DELETE FROM workbook_managed_targets')).toThrow('WORKBOOK_TARGET_IMMUTABLE');
    expect(sqlite.prepare('SELECT count(*) FROM workbook_managed_targets').pluck().get()).toBe(1);
    expect(sqlite.prepare('SELECT count(*) FROM order_lines').pluck().get()).toBe(1);
  });
});
