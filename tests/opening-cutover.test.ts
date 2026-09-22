import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { previewOpeningCutover, stageOpeningCutover, type OpeningCutoverRequest } from '../src/opening-cutover';
import { createMovement } from '../src/inventory';
import { upsertOrderSnapshot } from '../src/sync';
import { applyMigrations, sqliteD1 } from './d1';

type Row = Record<string, unknown>;
const policy = { storeId: '2442119', actor: 'admin@example.com', now: '2026-09-22T12:01:00.000Z' };
let sqlite: Database.Database; let db: D1Database;
function input(): OpeningCutoverRequest {
  const target = { id: '123', sku: 'BOLT-1', name: 'Bolt', quantity: 12, unlimited: false, hasOptions: false,
    hasVariations: false, combinationId: null, variationOptions: [], enabled: true, eligibilityVerified: true,
    hasExtraOptions: false, hasBundleRelationships: false };
  return { operation_id: crypto.randomUUID(), expected_hash: '', confirm_staging: true, physical_counts_confirmed: true,
    freeze: { confirmed: true, started_at: '2026-09-22T12:00:00.000Z' },
    input: { store_id: policy.storeId, source_ref: 'Verified factory workbook', snapshot_source_hash: 'a'.repeat(64),
      balance_meaning: 'PHYSICAL_ON_HAND', reservations_confirmed: true,
      rows: [{ sku: 'BOLT-1', balance: 10, name: 'Bolt', single_unit_confirmed: true, source_row: 12, source_sheet: 'Fasteners' }],
      reservations: [{ sku: 'BOLT-1', quantity: 3 }],
      catalog: { kind: 'READONLY_CATALOGUE', schema_version: 1, dry_run: true, complete: true, store_id: policy.storeId,
        started_at: '2026-09-22T12:00:01.000Z', completed_at: '2026-09-22T12:00:30.000Z',
        product_count: 1, stock_target_count: 1, products: [target], stock_targets: [target], reservations_confirmed: false } },
    scope: [{ sku: 'BOLT-1', ecwid_product_id: '123', ecwid_combination_id: null, ecwid_option_signature: '[]' }],
    orders: { kind: 'READONLY_ORDERS', schema_version: 1, dry_run: true, complete: true, store_id: policy.storeId,
      started_at: '2026-09-22T12:00:01.000Z', completed_at: '2026-09-22T12:00:30.000Z',
      creation_cutoff: Date.parse('2026-09-22T12:00:01.000Z') / 1000, orders_checked: 100,
      pending_order_count: 1, line_count: 1,
      orders: [{ id: '12', paymentStatus: 'PAID', fulfillmentStatus: 'AWAITING_PROCESSING', updatedAt: '2026-09-22T11:30:00.000Z',
        items: [{ id: '111', productId: '123', sku: 'BOLT-1', name: 'Bolt', quantity: 3, combinationId: null,
          selectedOptions: [], digital: false, trackQuantity: true }] }] },
    line_confirmations: [{ order_id: '12', ecwid_line_id: '111', previously_picked_quantity: 0 }], workbook_scope: [] };
}
async function approved(req = input()) {
  req.expected_hash = (await previewOpeningCutover(req, policy)).review_hash;
  return req;
}
const source = (req: OpeningCutoverRequest) => req.input.rows as Row[];
const catalog = (req: OpeningCutoverRequest) => req.input.catalog as Row;
const targets = (req: OpeningCutoverRequest) => catalog(req).stock_targets as Row[];
function counts() {
  return Object.fromEntries(['items','opening_balances','orders','order_lines','movements','outbox','sync_issues',
    'opening_cutover_batches','opening_cutover_rows','opening_cutover_orders','workbook_managed_targets'].map(table =>
    [table, sqlite.prepare(`SELECT count(*) FROM ${table}`).pluck().get()]));
}
function addWorkbookLine(req: OpeningCutoverRequest) {
  req.workbook_scope.push({ sku: 'WORKBOOK-ONLY', name: 'Workbook item', ecwid_product_id: '456', ecwid_combination_id: null, ecwid_option_signature: '[]' });
  req.orders.orders[0].items.push({ id: '112', productId: '456', sku: 'WORKBOOK-ONLY', name: 'Workbook item', quantity: 4,
    combinationId: null, selectedOptions: [], digital: false, trackQuantity: false });
  req.orders.line_count++;
  req.line_confirmations.push({ order_id: '12', ecwid_line_id: '112', previously_picked_quantity: 0 });
  const workbook = { ...targets(req)[0], id: '456', sku: 'WORKBOOK-ONLY', name: 'Workbook item', unlimited: true, quantity: null };
  targets(req).push(workbook); catalog(req).products = targets(req);
  catalog(req).product_count = catalog(req).stock_target_count = 2;
}
beforeEach(() => {
  sqlite = new Database(':memory:'); applyMigrations(sqlite); db = sqliteD1(sqlite);
  vi.stubGlobal('fetch', vi.fn(() => { throw new Error('Opening import must not call any network'); }));
});
afterEach(() => { sqlite.close(); vi.unstubAllGlobals(); });

describe('atomic opening cutover staging', () => {
  it('loads exact physical counts and unpicked commitments without enabling inventory or touching Ecwid', async () => {
    const req = await approved(); const result = await stageOpeningCutover(db, req, policy);
    expect(result).toMatchObject({ status: 'STAGED', duplicate: false, row_count: 1, order_count: 1, line_count: 1,
      reservations_loaded: true, activated: false, ecwid_changed: false });
    expect(sqlite.prepare('SELECT on_hand,reserved,available,active FROM item_stock').get()).toEqual({ on_hand: 10, reserved: 3, available: 7, active: 0 });
    expect(sqlite.prepare('SELECT physical,unpicked,target_quantity,expected_ecwid_quantity,alignment_status FROM opening_cutover_rows').get())
      .toEqual({ physical: 10, unpicked: 3, target_quantity: 7, expected_ecwid_quantity: 12, alignment_status: 'PENDING' });
    expect(sqlite.prepare('SELECT state,review_hash,source_hash,actor FROM opening_cutover_batches').get())
      .toEqual({ state: 'STAGED', review_hash: req.expected_hash, source_hash: 'a'.repeat(64), actor: policy.actor });
    expect(counts()).toEqual({ items: 1, opening_balances: 1, orders: 1, order_lines: 1, movements: 0, outbox: 0,
      sync_issues: 1, opening_cutover_batches: 1, opening_cutover_rows: 1, opening_cutover_orders: 1, workbook_managed_targets: 0 });
    expect(sqlite.pragma('foreign_key_check')).toEqual([]); expect(fetch).not.toHaveBeenCalled();
  });

  it('includes explicitly workbook-managed lines without reserving or inventing physical app stock for them', async () => {
    const req = input(); addWorkbookLine(req);
    await stageOpeningCutover(db, await approved(req), policy);
    expect(sqlite.prepare("SELECT management_mode,item_id,workbook_target_id,picked_qty FROM order_lines WHERE sku='WORKBOOK-ONLY'").get())
      .toEqual({ management_mode: 'WORKBOOK', item_id: null, workbook_target_id: 'workbook:456:simple', picked_qty: 0 });
    expect(sqlite.prepare('SELECT needs_review FROM orders').pluck().get()).toBe(0);
    expect(counts().items).toBe(1); expect(counts().order_lines).toBe(2);
    // The imported hash is identical to sync's canonical hash, not a new hash convention.
    sqlite.exec('UPDATE items SET active=1');
    await upsertOrderSnapshot(db, req.orders.orders[0]);
    expect(sqlite.prepare('SELECT needs_review FROM orders').pluck().get()).toBe(0);
  });

  it('reserves Awaiting Payment lines but keeps picking blocked', async () => {
    const req = input(); req.orders.orders[0].paymentStatus = 'AWAITING_PAYMENT';
    await stageOpeningCutover(db, await approved(req), policy);
    expect(sqlite.prepare('SELECT reserved FROM item_stock').pluck().get()).toBe(3);
    sqlite.exec("UPDATE items SET active=1; UPDATE sync_issues SET status='RESOLVED'");
    const itemId = sqlite.prepare('SELECT id FROM items').pluck().get() as string;
    await expect(createMovement(db, { operation_id: crypto.randomUUID(), type: 'ECWID_PICK', item_id: itemId, quantity: 1,
      order_id: '12', order_line_id: '12:111' }, policy.actor)).rejects.toMatchObject({ code: 'ORDER_NOT_PICKABLE' });
  });

  it('allows an explicitly empty complete order snapshot and a confirmed zero opening balance', async () => {
    const req = input(); source(req)[0].balance = 0; req.input.reservations = [];
    req.orders.orders = []; req.orders.pending_order_count = req.orders.line_count = 0; req.line_confirmations = [];
    await stageOpeningCutover(db, await approved(req), policy);
    expect(sqlite.prepare('SELECT physical,unpicked,target_quantity FROM opening_cutover_rows').get()).toEqual({ physical: 0, unpicked: 0, target_quantity: 0 });
    expect(counts().orders).toBe(0);
  });

  it('imports independent variation identities without changing the parent or siblings', async () => {
    const req = input(); const options = [{ name: 'Length', value: '20 mm' }];
    Object.assign(targets(req)[0], { combinationId: '987', variationOptions: options, hasOptions: true });
    Object.assign(req.scope[0], { ecwid_combination_id: '987', ecwid_option_signature: JSON.stringify(options) });
    Object.assign(req.orders.orders[0].items[0], { combinationId: '987', selectedOptions: options });
    await stageOpeningCutover(db, await approved(req), policy);
    expect(sqlite.prepare('SELECT ecwid_product_id,ecwid_combination_id FROM items').get()).toEqual({ ecwid_product_id: '123', ecwid_combination_id: '987' });
  });

  it('replays the same approved UUID without expiry failures or duplicate opening ledger entries', async () => {
    const req = await approved(); const first = await stageOpeningCutover(db, req, policy); const initial = counts();
    expect(await stageOpeningCutover(db, req, { ...policy, now: '2026-09-23T00:00:00.000Z' })).toEqual({ ...first, duplicate: true });
    expect(counts()).toEqual(initial);
  });

  it('handles concurrent duplicate staging and rejects actor/content reuse', async () => {
    const req = await approved(); const results = await Promise.all([stageOpeningCutover(db,req,policy),stageOpeningCutover(db,req,policy)]);
    expect(results.map(row => row.duplicate).sort()).toEqual([false,true]);
    await expect(stageOpeningCutover(db,req,{ ...policy,actor:'other@example.com' })).rejects.toMatchObject({ code:'OPENING_OPERATION_REUSED' });
    const changed = input(); changed.operation_id=req.operation_id; source(changed)[0].balance=11;
    await expect(stageOpeningCutover(db,await approved(changed),policy)).rejects.toMatchObject({ code:'OPENING_OPERATION_REUSED' });
    expect(counts().opening_balances).toBe(1);
  });

  it.each(['SKU','order','late'])('rolls back the entire atomic import on %s conflict/failure', async kind => {
    if (kind==='SKU') sqlite.exec("INSERT INTO items(id,sku,name,scan_code) VALUES('old','BOLT-1','Old','OLD')");
    if (kind==='order') sqlite.exec("INSERT INTO orders(id,payment_status,remote_updated_at,updated_at) VALUES('12','PAID','old','old')");
    if (kind==='late') sqlite.exec("CREATE TRIGGER fail_opening_issue BEFORE INSERT ON sync_issues BEGIN SELECT RAISE(ABORT,'TEST_LATE_FAILURE'); END");
    const initial=counts(); const req=input(); addWorkbookLine(req);
    await expect(stageOpeningCutover(db,await approved(req),policy)).rejects.toThrow();
    expect(counts()).toEqual(initial);
  });

  it.each(['opening_cutover_batches','opening_cutover_rows','opening_cutover_orders'])('keeps %s audit immutable', async table => {
    await stageOpeningCutover(db,await approved(),policy);
    expect(() => sqlite.exec(`DELETE FROM ${table}`)).toThrow('OPENING_CUTOVER_IMMUTABLE');
    expect(() => sqlite.exec(`UPDATE ${table} SET operation_id='changed'`)).toThrow('OPENING_CUTOVER_IMMUTABLE');
  });

  it('keeps staged items inactive and blocks manual activation from bypassing quarantine', async () => {
    await stageOpeningCutover(db,await approved(),policy);
    const itemId=sqlite.prepare('SELECT id FROM items').pluck().get() as string;
    const movement={operation_id:crypto.randomUUID(),type:'RESTOCK',item_id:itemId,quantity:1};
    await expect(createMovement(db,movement,policy.actor)).rejects.toMatchObject({code:'ITEM_UNAVAILABLE'});
    sqlite.exec('UPDATE items SET active=1');
    await expect(createMovement(db,movement,policy.actor)).rejects.toMatchObject({code:'ITEM_NEEDS_REVIEW'});
  });
});

describe('opening cutover validation', () => {
  it.each(['physical','freeze','staging'])('requires explicit %s confirmation', async field => {
    const req = input();
    if(field==='physical') Object.assign(req,{physical_counts_confirmed:false});
    if(field==='freeze') Object.assign(req.freeze,{confirmed:false});
    if(field==='staging') Object.assign(req,{confirm_staging:false});
    await expect(previewOpeningCutover(req,policy)).rejects.toBeInstanceOf(Error); expect(counts().items).toBe(0);
  });

  it.each(['incomplete','wrong-store','wrong-count','missing-line-count','bad-cutoff','future-order','duplicate-order','duplicate-line','no-lines'])
  ('rejects deficient complete-order evidence: %s', async kind => {
    const req=input();
    if(kind==='incomplete') Object.assign(req.orders,{complete:false});
    if(kind==='wrong-store') req.orders.store_id='999';
    if(kind==='wrong-count') req.orders.pending_order_count=2;
    if(kind==='missing-line-count') Object.assign(req.orders,{line_count:undefined});
    if(kind==='bad-cutoff') req.orders.creation_cutoff--;
    if(kind==='future-order') req.orders.orders[0].updatedAt='2026-09-23T00:00:00.000Z';
    if(kind==='duplicate-order') {req.orders.orders.push(req.orders.orders[0]);req.orders.pending_order_count=2;}
    if(kind==='duplicate-line') req.orders.orders[0].items.push(req.orders.orders[0].items[0]);
    if(kind==='no-lines') req.orders.orders[0].items=[];
    await expect(previewOpeningCutover(req,policy)).rejects.toBeInstanceOf(Error);
  });

  it.each(['old','before-freeze','future','catalogue-old'])('requires fresh snapshots within the confirmed freeze: %s', async kind => {
    const req=input(); const p={...policy};
    if(kind==='old') p.now='2026-09-22T12:16:00.000Z';
    if(kind==='before-freeze') req.freeze.started_at='2026-09-22T12:00:02.000Z';
    if(kind==='future') req.orders.completed_at='2026-09-22T12:02:00.000Z';
    if(kind==='catalogue-old') catalog(req).started_at='2026-09-22T11:59:00.000Z';
    await expect(previewOpeningCutover(req,p)).rejects.toMatchObject({code:'OPENING_SNAPSHOT_STALE'});
  });

  it.each(['CANCELLED','REFUNDED','INCOMPLETE','PARTIALLY_REFUNDED'])('rejects ambiguous/nonpending payment status %s', async status => {
    const req=input();req.orders.orders[0].paymentStatus=status;
    await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_ORDER_STATUS_UNSUPPORTED'});
  });
  it.each(['SHIPPED','READY_FOR_PICKUP','DELIVERED','UNKNOWN'])('rejects terminal or unknown opening fulfillment %s', async status => {
    const req=input();req.orders.orders[0].fulfillmentStatus=status;
    await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_ORDER_STATUS_UNSUPPORTED'});
  });

  it.each(['prior-pick','missing','duplicate','wrong-line'])('requires exact zero previous-pick confirmation per line: %s', async kind => {
    const req=input();
    if(kind==='prior-pick') Object.assign(req.line_confirmations[0],{previously_picked_quantity:1});
    if(kind==='missing') req.line_confirmations=[];
    if(kind==='duplicate') req.line_confirmations.push(req.line_confirmations[0]);
    if(kind==='wrong-line') req.line_confirmations[0].ecwid_line_id='999';
    await expect(previewOpeningCutover(req,policy)).rejects.toBeInstanceOf(Error);
  });

  it.each(['missing','wrong-total','foreign-summary'])('rejects reservation summaries not backed by exact pilot lines: %s', async kind => {
    const req=input();
    if(kind==='missing') req.input.reservations=[];
    if(kind==='wrong-total') req.input.reservations=[{sku:'BOLT-1',quantity:2}];
    if(kind==='foreign-summary') req.input.reservations=[{sku:'BOLT-1',quantity:3},{sku:'IGNORED',quantity:1}];
    await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_RESERVATIONS_MISMATCH'});
  });

  it.each(['unknown','product-mismatch','options-mismatch','variation-mismatch'])('never silently drops unresolved order lines: %s', async kind => {
    const req=input();const line=req.orders.orders[0].items[0];
    if(kind==='unknown') line.sku='UNKNOWN';
    if(kind==='product-mismatch') line.productId='999';
    if(kind==='options-mismatch') line.selectedOptions=[{name:'Size',value:'M6'}];
    if(kind==='variation-mismatch') line.combinationId='999';
    await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_ORDER_LINE_UNMAPPED'});
  });

  it.each(['digital','missing-digital','malformed-options','fractional-quantity'])('rejects malformed line evidence: %s', async kind => {
    const req=input();const line=req.orders.orders[0].items[0];
    if(kind==='digital') line.digital=true;
    if(kind==='missing-digital') Object.assign(line,{digital:undefined});
    if(kind==='malformed-options') line.selectedOptions=[{name:'text',value:'free',type:'TEXT'}];
    if(kind==='fractional-quantity') line.quantity=1.5;
    await expect(previewOpeningCutover(req,policy)).rejects.toBeInstanceOf(Error);
  });

  it('rejects supplier-backed/unlimited stock instead of coercing it to stock-limited', async () => {
    const req=input();targets(req)[0].unlimited=true;
    await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_PREVIEW_BLOCKED'});
  });
  it('rejects workbook and app target overlap', async () => {
    const req=input();req.workbook_scope=[{...req.scope[0],name:'Bolt'}];
    await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_SCOPE_CONFLICT'});
  });
  it.each(['missing','wrong-options','duplicate-sku'])('rejects invented or ambiguous workbook catalogue mappings: %s', async kind => {
    const req=input();addWorkbookLine(req);
    if(kind==='missing') targets(req)[1].id='999';
    if(kind==='wrong-options') req.workbook_scope[0].ecwid_option_signature='[{"name":"Length","value":"100"}]';
    if(kind==='duplicate-sku') {targets(req).push({...targets(req)[1],id:'999'});catalog(req).product_count=catalog(req).stock_target_count=3;}
    await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_WORKBOOK_CATALOGUE_MISMATCH'});
  });
  it('hash-binds workbook balance, complete orders, scope, confirmation and administrator', async () => {
    const req=await approved();source(req)[0].balance=20;
    await expect(stageOpeningCutover(db,req,policy)).rejects.toMatchObject({code:'OPENING_REVIEW_CHANGED'});
  });
  it('rejects target scopes that omit, duplicate or alter approved identities', async () => {
    for (const scope of [[],[input().scope[0],input().scope[0]],[{...input().scope[0],ecwid_product_id:'999'}]]) {
      const req=input();req.scope=scope;
      await expect(previewOpeningCutover(req,policy)).rejects.toMatchObject({code:'OPENING_SCOPE_MISMATCH'});
    }
  });
});

describe('alignment journal safety', () => {
  it.each(['before','after','verified_at'])('does not let SQLite NULL CHECK semantics bypass VERIFIED %s evidence', async missing => {
    await stageOpeningCutover(db,await approved(),policy);
    sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNING'");
    if(missing==='before') {
      expect(() => sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='VERIFIED',after_quantity=7,verified_at='now'")).toThrow();
    } else {
      sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='PROCESSING',before_quantity=12,attempted_at='now'");
      const assignment=missing==='after'?"verified_at='now'":'after_quantity=7';
      expect(() => sqlite.exec(`UPDATE opening_cutover_rows SET alignment_status='VERIFIED',${assignment}`)).toThrow();
    }
  });
  it('only permits direct PENDING verification for a confirmed no-op', async () => {
    await stageOpeningCutover(db,await approved(),policy);
    sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNING'");
    expect(() => sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='VERIFIED',before_quantity=12,after_quantity=7,verified_at='now'"))
      .toThrow('OPENING_CUTOVER_STATE_INVALID');
    sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='VERIFIED',before_quantity=7,after_quantity=7,verified_at='now'");
    expect(sqlite.prepare('SELECT alignment_status FROM opening_cutover_rows').pluck().get()).toBe('VERIFIED');
  });
  it('cannot rewrite the recorded pre-write quantity or attempt time after claiming a row', async () => {
    await stageOpeningCutover(db,await approved(),policy);
    sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNING'");
    sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='PROCESSING',before_quantity=12,attempted_at='now'");
    expect(() => sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='VERIFIED',before_quantity=7,after_quantity=7,verified_at='now'"))
      .toThrow('OPENING_CUTOVER_IMMUTABLE');
    expect(() => sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='UNKNOWN',attempted_at='later'"))
      .toThrow('OPENING_CUTOVER_IMMUTABLE');
  });
  it('cannot skip review states or activate without every verified target', async () => {
    const req=await approved();await stageOpeningCutover(db,req,policy);
    expect(() => sqlite.exec("UPDATE opening_cutover_batches SET state='ACTIVE'")).toThrow('OPENING_CUTOVER_STATE_INVALID');
    sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNING'");
    expect(() => sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNED'")).toThrow('OPENING_CUTOVER_NOT_VERIFIED');
    sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='PROCESSING',before_quantity=12,attempted_at='now'");
    expect(() => sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='VERIFIED',after_quantity=8,verified_at='now'")).toThrow();
    sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='VERIFIED',after_quantity=7,verified_at='now'");
    sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNED'");
    sqlite.exec("UPDATE opening_cutover_batches SET state='ACTIVE'");
    expect(sqlite.prepare('SELECT state FROM opening_cutover_batches').pluck().get()).toBe('ACTIVE');
    expect(await stageOpeningCutover(db,req,policy)).toMatchObject({status:'ACTIVE',duplicate:true,activated:true,ecwid_changed:true});
  });
  it('moves the batch to REVIEW on uncertain writes and disallows blind retry', async () => {
    const req=await approved();await stageOpeningCutover(db,req,policy);
    sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNING'");
    sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='PROCESSING',before_quantity=12,attempted_at='now'");
    sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='UNKNOWN',last_error='Timed out'");
    expect(sqlite.prepare('SELECT state FROM opening_cutover_batches').pluck().get()).toBe('REVIEW');
    expect(() => sqlite.exec("UPDATE opening_cutover_rows SET alignment_status='PROCESSING'")).toThrow('OPENING_CUTOVER_STATE_INVALID');
    expect(() => sqlite.exec("UPDATE opening_cutover_batches SET state='ALIGNING'")).toThrow('OPENING_CUTOVER_STATE_INVALID');
    expect(await stageOpeningCutover(db,req,policy)).toMatchObject({status:'REVIEW',duplicate:true,activated:false,ecwid_changed:null});
  });
});
