import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as auth from '../src/auth';
import worker from '../src/index';
import type { OpeningCutoverRequest } from '../src/opening-cutover';
import { applyMigrations, sqliteD1 } from './d1';

const origin='https://inventory.example.test';
const admin='admin@example.test';
const routes=['/api/cutover/preview','/api/cutover/stage'];
let sqlite: Database.Database;
let env: Env;
let ctx: ExecutionContext;
let background: Promise<unknown>[];

function fixture(): OpeningCutoverRequest {
  const current=Date.now();
  const time=(offset:number)=>new Date(current+offset).toISOString();
  const start=time(-2000); const completed=time(-1000);
  const target={id:'123',sku:'BOLT-1',name:'Bolt',quantity:12,unlimited:false,enabled:true,
    hasOptions:false,hasVariations:false,combinationId:null,variationOptions:[],eligibilityVerified:true,
    hasExtraOptions:false,hasBundleRelationships:false};
  return {operation_id:crypto.randomUUID(),expected_hash:'',confirm_staging:true,physical_counts_confirmed:true,
    freeze:{confirmed:true,started_at:time(-3000)},
    input:{store_id:'2442119',source_ref:'Synthetic verified workbook',snapshot_source_hash:'a'.repeat(64),
      balance_meaning:'PHYSICAL_ON_HAND',reservations_confirmed:true,
      rows:[{sku:'BOLT-1',name:'Bolt',balance:10,single_unit_confirmed:true}],reservations:[{sku:'BOLT-1',quantity:3}],
      catalog:{kind:'READONLY_CATALOGUE',schema_version:1,dry_run:true,complete:true,store_id:'2442119',
        started_at:start,completed_at:completed,product_count:1,stock_target_count:1,products:[target],stock_targets:[target],reservations_confirmed:false}},
    scope:[{sku:'BOLT-1',ecwid_product_id:'123',ecwid_combination_id:null,ecwid_option_signature:'[]'}],
    orders:{kind:'READONLY_ORDERS',schema_version:1,dry_run:true,complete:true,store_id:'2442119',started_at:start,
      completed_at:completed,creation_cutoff:Math.floor(Date.parse(start)/1000),orders_checked:1,pending_order_count:1,line_count:1,
      orders:[{id:'12',paymentStatus:'PAID',fulfillmentStatus:'AWAITING_PROCESSING',updatedAt:time(-60000),createdAt:time(-120000),
        items:[{id:'111',productId:'123',sku:'BOLT-1',name:'Bolt',quantity:3,combinationId:null,selectedOptions:[],digital:false,trackQuantity:true}]}]},
    line_confirmations:[{order_id:'12',ecwid_line_id:'111',previously_picked_quantity:0}],workbook_scope:[]};
}

function request(path:string,body?:unknown,headers:Record<string,string>={}) {
  return new Request(origin+path,{method:body===undefined?'GET':'POST',
    headers:{...(body===undefined?{}:{'Content-Type':'application/json',Origin:origin}),...headers},
    body:body===undefined?undefined:JSON.stringify(body)});
}
function call(path:string,body?:unknown,headers:Record<string,string>={}) {
  return worker.fetch(request(path,body,headers),env,ctx);
}
function counts() {
  return Object.fromEntries(['items','opening_balances','orders','order_lines','movements','outbox','sync_issues',
    'opening_cutover_batches','opening_cutover_rows','opening_cutover_orders','workbook_managed_targets','sync_state']
    .map(table=>[table,sqlite.prepare(`SELECT COUNT(*) AS n FROM ${table}`).pluck().get()]));
}
async function approved(input=fixture()) {
  const response=await call('/api/cutover/preview',input);
  expect(response.status).toBe(200);
  const result=await response.json<{review_hash:string}>();
  input.expected_hash=result.review_hash;
  return input;
}

beforeEach(()=>{
  sqlite=new Database(':memory:');applyMigrations(sqlite);
  env={DB:sqliteD1(sqlite),ECWID_MODE:'live',ECWID_STORE_ID:'2442119',ECWID_TOKEN:'test-only-token',ECWID_CLIENT_SECRET:'',
    INVENTORY_ENABLED:'false',LIVE_SYNC_ENABLED:'false',ORDER_SYNC_ENABLED:'false',
    ACCESS_TEAM_DOMAIN:'https://example.cloudflareaccess.com',ACCESS_AUD:'test-audience',ADMIN_EMAILS:admin,STAFF_EMAILS:admin,
    ASSETS:{fetch:vi.fn()} as unknown as Fetcher,SYNC_QUEUE:{send:vi.fn()} as unknown as Queue};
  background=[];
  ctx={waitUntil:(promise:Promise<unknown>)=>background.push(promise)} as unknown as ExecutionContext;
  vi.spyOn(auth,'authenticate').mockResolvedValue({actor:admin,role:'admin'});
  vi.stubGlobal('fetch',vi.fn(()=>{throw new Error('No real network is permitted in cutover router tests.');}));
});
afterEach(async()=>{
  await Promise.all(background);
  expect(fetch).not.toHaveBeenCalled();
  expect(env.SYNC_QUEUE.send).not.toHaveBeenCalled();
  sqlite.close();vi.restoreAllMocks();vi.unstubAllGlobals();
});

describe('cutover HTTP authorization and fail-closed configuration',()=>{
  it.each(routes)('requires authenticated administrator for %s',async path=>{
    vi.mocked(auth.authenticate).mockResolvedValue({actor:'picker@example.test',role:'picker'});
    const response=await call(path,fixture());
    expect(response.status).toBe(403);expect(await response.json()).toMatchObject({code:'ADMIN_REQUIRED'});
    expect(counts().items).toBe(0);
  });
  it.each(routes)('requires a real sign-in before %s',async path=>{
    vi.mocked(auth.authenticate).mockRestore();
    const response=await call(path,fixture());
    expect(response.status).toBe(401);expect(await response.json()).toMatchObject({code:'SIGN_IN_REQUIRED'});
    expect(counts().opening_cutover_batches).toBe(0);
  });
  it.each(routes.flatMap(path=>[
    {path,headers:{Origin:'https://foreign.example.test'}},
    {path,headers:{'Sec-Fetch-Site':'cross-site'}},
    {path,headers:{'Sec-Fetch-Site':'same-site'}},
  ]))('rejects cross-origin submission to $path with $headers',async({path,headers})=>{
    const response=await call(path,fixture(),headers);
    expect(response.status).toBe(403);expect(await response.json()).toMatchObject({code:'CROSS_ORIGIN'});
    expect(counts().items).toBe(0);
  });
  const flagCases=routes.flatMap(path=>['INVENTORY_ENABLED','LIVE_SYNC_ENABLED','ORDER_SYNC_ENABLED'].flatMap(flag=>
    ['true','FALSE','',undefined,false].map(value=>({path,flag,value}))));
  it.each(flagCases)('requires exact false for $flag=$value on $path',async({path,flag,value})=>{
    Object.assign(env,{[flag]:value});
    const response=await call(path,fixture());
    expect(response.status).toBe(409);expect(await response.json()).toMatchObject({code:'CUTOVER_REQUIRES_DISABLED_FLAGS'});
    expect(counts().opening_cutover_batches).toBe(0);
  });
  it.each(routes)('requires configured store for %s',async path=>{
    env.ECWID_STORE_ID='';
    const response=await call(path,fixture());
    expect(response.status).toBe(409);expect(await response.json()).toMatchObject({code:'IMPORT_STORE_REQUIRED'});
  });
  it.each(routes)('rejects another store in %s',async path=>{
    const input=fixture();input.input.store_id='999';
    const response=await call(path,input);
    expect(response.status).toBe(400);expect(await response.json()).toMatchObject({code:'OPENING_STORE_MISMATCH'});
    expect(counts().opening_cutover_batches).toBe(0);
  });
});

describe('cutover HTTP preview and atomic staging',()=>{
  it('previews real source-derived reservations without writing to the database',async()=>{
    const initial=counts();const response=await call('/api/cutover/preview',fixture());
    expect(response.status).toBe(200);
    expect(response.headers.get('Cache-Control')).toBe('no-store');
    expect(await response.json()).toMatchObject({dry_run:true,row_count:1,order_count:1,line_count:1,
      reservations_loaded:false,activated:false,ecwid_changed:false,
      rows:[{sku:'BOLT-1',physical_on_hand:10,reserved:3,desired_ecwid_quantity:7}]});
    expect(counts()).toEqual(initial);
  });
  it.each(['missing','wrong','changed-after-preview'])('rejects %s review approval',async fault=>{
    const input=await approved();
    if(fault==='missing') input.expected_hash='';
    if(fault==='wrong') input.expected_hash='f'.repeat(64);
    if(fault==='changed-after-preview') (input.input.rows as {balance:number}[])[0].balance=11;
    const response=await call('/api/cutover/stage',input);
    expect(response.status).toBe(409);expect(await response.json()).toMatchObject({code:'OPENING_REVIEW_CHANGED'});
    expect(counts().items).toBe(0);
  });
  it('stages one inactive pilot and exact order once, returning HTTP 200 for identical retries',async()=>{
    const input=await approved();const response=await call('/api/cutover/stage',input);
    expect(response.status).toBe(201);
    const result=await response.json();
    expect(result).toMatchObject({status:'STAGED',duplicate:false,row_count:1,order_count:1,line_count:1,
      reservations_loaded:true,activated:false,ecwid_changed:false});
    expect(sqlite.prepare('SELECT on_hand,reserved,available,active FROM item_stock').get()).toEqual({on_hand:10,reserved:3,available:7,active:0});
    expect(sqlite.prepare('SELECT ordered_qty,picked_qty,management_mode FROM order_lines').get()).toEqual({ordered_qty:3,picked_qty:0,management_mode:'APP'});
    expect(sqlite.prepare('SELECT physical,unpicked,target_quantity,alignment_status FROM opening_cutover_rows').get()).toEqual({physical:10,unpicked:3,target_quantity:7,alignment_status:'PENDING'});
    expect(counts()).toMatchObject({items:1,orders:1,order_lines:1,opening_balances:1,movements:0,outbox:0,sync_issues:1});
    const initial=counts();const retry=await call('/api/cutover/stage',input);
    expect(retry.status).toBe(200);expect(await retry.json()).toEqual({...result as object,duplicate:true});
    expect(counts()).toEqual(initial);expect(sqlite.pragma('foreign_key_check')).toEqual([]);expect(background).toHaveLength(0);
  });
  it('rolls back every insert when a late staging statement fails',async()=>{
    const input=await approved();const initial=counts();
    sqlite.exec("CREATE TRIGGER reject_cutover_issue BEFORE INSERT ON sync_issues BEGIN SELECT RAISE(ABORT,'TEST_LATE_FAILURE'); END");
    vi.spyOn(console,'error').mockImplementation(()=>{});
    const response=await call('/api/cutover/stage',input);
    expect(response.status).toBe(500);expect(await response.json()).toMatchObject({code:'INTERNAL_ERROR'});
    expect(counts()).toEqual(initial);
  });
  it('reads a staged live order locally while order sync is disabled and refuses explicit refresh',async()=>{
    const input=await approved();expect((await call('/api/cutover/stage',input)).status).toBe(201);
    const response=await call('/api/orders/12');
    expect(response.status).toBe(200);
    expect(await response.json()).toMatchObject({order:{id:'12',payment_status:'PAID',needs_review:0,
      lines:[{sku:'BOLT-1',ordered_qty:3,picked_qty:0,item_active:0,pickable_qty:0}]}});
    const refresh=await call('/api/orders/12/refresh',{});
    expect(refresh.status).toBe(409);expect(await refresh.json()).toMatchObject({code:'ORDER_SYNC_PAUSED'});
  });
});
