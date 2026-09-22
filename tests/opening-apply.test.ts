import Database from 'better-sqlite3';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { stageOpeningImport } from '../src/opening-apply';
import { previewImport } from '../src/opening-import';
import { createMovement } from '../src/inventory';
import { applyMigrations, sqliteD1 } from './d1';

type Row = Record<string, unknown>;
const policy = { storeId: '2442119', actor: 'admin@example.com' };
let sqlite: Database.Database;
let db: D1Database;
function stockInput(): Row {
  const target = {
    id: '123', sku: '001-BOLT', name: 'M3 bolt', quantity: 7, unlimited: false,
    hasOptions: false, hasVariations: false, combinationId: null, variationOptions: [], enabled: true,
    eligibilityVerified: true, hasExtraOptions: false, hasBundleRelationships: false
  };
  return {
    store_id: policy.storeId, source_ref: 'Factory stock workbook, reviewed 22 September 2026',
    snapshot_source_hash: 'a'.repeat(64), balance_meaning: 'PHYSICAL_ON_HAND', reservations_confirmed: true,
    rows: [{ sku: '001-BOLT', balance: 10, name: 'M3 bolt', location: 'Shelf A', single_unit_confirmed: true,
      source_row: 12, source_sheet: 'Fasteners' }],
    catalog: {
      kind: 'READONLY_CATALOGUE', schema_version: 1, dry_run: true, complete: true, store_id: policy.storeId,
      started_at: '2026-09-22T08:00:00.000Z', completed_at: '2026-09-22T08:00:10.000Z',
      product_count: 1, stock_target_count: 1, reservations_confirmed: false,
      products: [target], stock_targets: [target]
    },
    reservations: []
  };
}
async function request(input = stockInput()) {
  const preview = await previewImport(input);
  return {
    operation_id: crypto.randomUUID(), expected_hash: preview.source_hash, confirm_staging: true, input,
    scope: preview.rows.map(row => ({ sku: row.sku, ecwid_product_id: row.ecwid_product_id,
      ecwid_combination_id: row.ecwid_combination_id, ecwid_option_signature: row.ecwid_option_signature }))
  };
}
function counts() {
  return Object.fromEntries(['items', 'opening_balances', 'opening_import_batches', 'opening_import_rows', 'sync_issues', 'movements', 'outbox', 'orders', 'order_lines']
    .map(table => [table, sqlite.prepare(`SELECT count(*) FROM ${table}`).pluck().get()]));
}
function catalogOf(input: Row): Row { return input.catalog as Row; }
function targetsOf(input: Row): Row[] { return catalogOf(input).stock_targets as Row[]; }
function sourceOf(input: Row): Row[] { return input.rows as Row[]; }
function addSecond(input: Row) {
  const second = { ...targetsOf(input)[0], id: '124', sku: '002-BOLT' };
  const catalog = catalogOf(input);
  catalog.stock_targets = [...targetsOf(input), second];
  catalog.products = catalog.stock_targets;
  catalog.product_count = catalog.stock_target_count = 2;
  input.rows = [...sourceOf(input), { ...sourceOf(input)[0], sku: '002-BOLT' }];
}

beforeEach(() => {
  sqlite = new Database(':memory:');
  applyMigrations(sqlite);
  db = sqliteD1(sqlite);
  vi.stubGlobal('fetch', vi.fn(() => { throw new Error('Opening staging must not access any network'); }));
});
afterEach(() => { sqlite.close(); vi.unstubAllGlobals(); });

describe('DB-only immutable opening-stock staging', () => {
  it('stages an inactive item and immutable physical opening ledger with provenance, without orders or Ecwid writes', async () => {
    const req = await request();
    const staged = await stageOpeningImport(db, req, policy);
    expect(staged).toMatchObject({ status: 'STAGED', duplicate: false, activated: false, reservations_loaded: false,
      ecwid_changed: false, store_id: policy.storeId, row_count: 1, preview_hash: req.expected_hash });
    expect(counts()).toEqual({ items: 1, opening_balances: 1, opening_import_batches: 1, opening_import_rows: 1,
      sync_issues: 1, movements: 0, outbox: 0, orders: 0, order_lines: 0 });
    expect(sqlite.prepare('SELECT sku,on_hand,active,last_ecwid_quantity,ecwid_product_id,ecwid_combination_id FROM items').get())
      .toEqual({ sku: '001-BOLT', on_hand: 10, active: 0, last_ecwid_quantity: 7, ecwid_product_id: '123', ecwid_combination_id: null });
    expect(sqlite.prepare('SELECT source_hash,preview_hash,actor,status FROM opening_import_batches').get())
      .toEqual({ source_hash: 'a'.repeat(64), preview_hash: req.expected_hash, actor: policy.actor, status: 'STAGED' });
    expect(sqlite.prepare('SELECT catalogue_hash FROM opening_import_batches').pluck().get()).toMatch(/^[a-f0-9]{64}$/);
    expect(sqlite.prepare('SELECT source_row,source_sheet,on_hand,ecwid_quantity FROM opening_import_rows').get())
      .toEqual({ source_row: 12, source_sheet: 'Fasteners', on_hand: 10, ecwid_quantity: 7 });
    expect(fetch).not.toHaveBeenCalled();
    expect(sqlite.pragma('foreign_key_check')).toEqual([]);
  });

  it('preserves unrelated existing data and allows legitimate zero opening stock', async () => {
    sqlite.prepare("INSERT INTO items(id,sku,name,scan_code,on_hand) VALUES('demo','DEMO','Demo','DEMO-QR',42)").run();
    const old = sqlite.prepare("SELECT * FROM items WHERE id='demo'").get();
    const input = stockInput(); sourceOf(input)[0].balance = 0;
    await stageOpeningImport(db, await request(input), policy);
    expect(sqlite.prepare("SELECT * FROM items WHERE id='demo'").get()).toEqual(old);
    expect(sqlite.prepare("SELECT on_hand FROM items WHERE sku='001-BOLT'").pluck().get()).toBe(0);
  });

  it('stages distinct variations of one parent with exact option identities', async () => {
    const input = stockInput(); addSecond(input);
    targetsOf(input).forEach((target, i) => Object.assign(target, {
      id: '123', combinationId: String(456 + i), hasOptions: true,
      variationOptions: [{ name: 'Length', value: `${10 + i * 10} mm` }]
    }));
    await stageOpeningImport(db, await request(input), policy);
    expect(sqlite.prepare('SELECT ecwid_product_id,ecwid_combination_id,ecwid_option_signature FROM items ORDER BY sku').all()).toEqual([
      { ecwid_product_id: '123', ecwid_combination_id: '456', ecwid_option_signature: '[{"name":"Length","value":"10 mm"}]' },
      { ecwid_product_id: '123', ecwid_combination_id: '457', ecwid_option_signature: '[{"name":"Length","value":"20 mm"}]' }
    ]);
  });

  it('returns an identical receipt on repeat and concurrent same-content operations', async () => {
    const req = await request();
    const pair = await Promise.all([stageOpeningImport(db, req, policy), stageOpeningImport(db, req, policy)]);
    expect(pair.map(value => value.duplicate).sort()).toEqual([false, true]);
    expect(pair[0].created_at).toBe(pair[1].created_at);
    const again = await stageOpeningImport(db, req, policy);
    expect(again).toEqual({ ...pair[0], duplicate: true });
    expect(counts().opening_balances).toBe(1);
  });

  it('rejects the same operation key with changed content or administrator', async () => {
    const first = await request(); await stageOpeningImport(db, first, policy);
    const input = stockInput(); sourceOf(input)[0].balance = 11;
    const changed = { ...await request(input), operation_id: first.operation_id };
    await expect(stageOpeningImport(db, changed, policy)).rejects.toMatchObject({ code: 'OPENING_OPERATION_REUSED' });
    await expect(stageOpeningImport(db, first, { ...policy, actor: 'other@example.com' })).rejects.toMatchObject({ code: 'OPENING_OPERATION_REUSED' });
    expect(sqlite.prepare('SELECT on_hand FROM items').pluck().get()).toBe(10);
  });

  it('handles concurrent same-key different-content requests without double loading', async () => {
    const first = await request(); const input = stockInput(); sourceOf(input)[0].balance = 11;
    const changed = { ...await request(input), operation_id: first.operation_id };
    const results = await Promise.allSettled([stageOpeningImport(db, first, policy), stageOpeningImport(db, changed, policy)]);
    expect(results.filter(value => value.status === 'fulfilled')).toHaveLength(1);
    expect(results.filter(value => value.status === 'rejected')).toHaveLength(1);
    expect(counts().opening_import_batches).toBe(1);
    expect(counts().opening_balances).toBe(1);
  });

  it('binds a staged database to one store even if later configuration changes', async () => {
    await stageOpeningImport(db, await request(), policy);
    const input = stockInput(); input.store_id = '999'; catalogOf(input).store_id = '999';
    targetsOf(input)[0].id = '456'; targetsOf(input)[0].sku = 'OTHER-STORE'; sourceOf(input)[0].sku = 'OTHER-STORE';
    const original = counts();
    await expect(stageOpeningImport(db, await request(input), { ...policy, storeId: '999' }))
      .rejects.toMatchObject({ code: 'OPENING_DATABASE_STORE_MISMATCH' });
    expect(counts()).toEqual(original);
  });

  it('atomically rejects a conflicting store when first batches race', async () => {
    const first = await request();
    const other = stockInput(); other.store_id = '999'; catalogOf(other).store_id = '999';
    targetsOf(other)[0].id = '456'; targetsOf(other)[0].sku = 'OTHER-STORE'; sourceOf(other)[0].sku = 'OTHER-STORE';
    const second = await request(other);
    const raced = await Promise.allSettled([
      stageOpeningImport(db, first, policy), stageOpeningImport(db, second, { ...policy, storeId: '999' })
    ]);
    expect(raced.filter(value => value.status === 'fulfilled')).toHaveLength(1);
    const rejected = raced.find(value => value.status === 'rejected');
    expect(rejected?.status === 'rejected' ? rejected.reason : null).toMatchObject({ code: 'OPENING_DATABASE_STORE_MISMATCH' });
    expect(counts().opening_balances).toBe(1);
    expect(counts().opening_import_batches).toBe(1);
  });

  it.each([
    ['SKU', "INSERT INTO items(id,sku,name,scan_code) VALUES('old','001-BOLT','Existing','OLD-QR')"],
    ['QR namespace', "INSERT INTO items(id,sku,name,scan_code) VALUES('old','OLD','Existing','001-BOLT')"],
    ['target identity', "INSERT INTO items(id,sku,name,scan_code,ecwid_product_id) VALUES('old','OLD','Existing','OLD','123')"]
  ])('rolls back the whole batch on an existing %s conflict', async (_name, sql) => {
    sqlite.exec(sql);
    const original = counts();
    await expect(stageOpeningImport(db, await request(), policy)).rejects.toMatchObject({ code: 'OPENING_ITEM_CONFLICT' });
    expect(counts()).toEqual(original);
  });

  it('rolls back previously inserted rows when a later row conflicts', async () => {
    sqlite.prepare("INSERT INTO items(id,sku,name,scan_code) VALUES('old','002-BOLT','Existing','OTHER')").run();
    const original = counts(); const input = stockInput(); addSecond(input);
    await expect(stageOpeningImport(db, await request(input), policy)).rejects.toMatchObject({ code: 'OPENING_ITEM_CONFLICT' });
    expect(counts()).toEqual(original);
  });

  it('rolls back items, opening balances and audits if the final quarantine write fails', async () => {
    sqlite.exec("CREATE TRIGGER reject_stage_issue BEFORE INSERT ON sync_issues BEGIN SELECT RAISE(ABORT,'TEST_LATE_FAILURE'); END");
    await expect(stageOpeningImport(db, await request(), policy)).rejects.toThrow('TEST_LATE_FAILURE');
    expect(Object.values(counts()).every(value => value === 0)).toBe(true);
  });

  it('keeps staged stock unavailable and quarantined even if active is manually changed', async () => {
    await stageOpeningImport(db, await request(), policy);
    const itemId = sqlite.prepare('SELECT id FROM items').pluck().get() as string;
    const movement = { operation_id: crypto.randomUUID(), type: 'RESTOCK', item_id: itemId, quantity: 1 };
    await expect(createMovement(db, movement, policy.actor)).rejects.toMatchObject({ code: 'ITEM_UNAVAILABLE' });
    sqlite.prepare('UPDATE items SET active=1 WHERE id=?').run(itemId);
    await expect(createMovement(db, movement, policy.actor)).rejects.toMatchObject({ code: 'ITEM_NEEDS_REVIEW' });
    expect(counts().outbox).toBe(0);
  });

  it.each(['opening_import_batches', 'opening_import_rows'])('prevents mutation or deletion of %s audit', async table => {
    await stageOpeningImport(db, await request(), policy);
    await expect(db.prepare(`DELETE FROM ${table}`).run()).rejects.toThrow('OPENING_IMPORT_IMMUTABLE');
    await expect(db.prepare(`UPDATE ${table} SET operation_id=operation_id`).run()).rejects.toThrow('OPENING_IMPORT_IMMUTABLE');
  });

  it('requires an explicit matching SHA-256 review, not client READY labels', async () => {
    const stale = await request(); sourceOf(stale.input)[0].balance = 999;
    await expect(stageOpeningImport(db, stale, policy)).rejects.toMatchObject({ code: 'OPENING_PREVIEW_CHANGED' });
    const blocked = stockInput(); sourceOf(blocked)[0].single_unit_confirmed = false;
    const untrusted = { ...await request(blocked), ready_count: 1, status: 'READY' };
    await expect(stageOpeningImport(db, untrusted, policy)).rejects.toMatchObject({ code: 'INVALID_OPENING_STAGE' });
    await expect(stageOpeningImport(db, await request(blocked), policy)).rejects.toMatchObject({ code: 'OPENING_PREVIEW_BLOCKED' });
    expect(counts().items).toBe(0);
  });

  it.each([undefined, false, 'true', 1])('requires explicit boolean staging confirmation: %s', async confirm_staging => {
    await expect(stageOpeningImport(db, { ...await request(), confirm_staging }, policy)).rejects.toMatchObject({ code: 'INVALID_OPENING_STAGE' });
  });

  it.each(['', '2442118', 2442119])('rejects an input store that is not pinned to configuration: %s', async store_id => {
    const input = { ...stockInput(), store_id };
    const req = { ...await request(), input };
    await expect(stageOpeningImport(db, req, policy)).rejects.toMatchObject({ code: 'OPENING_STORE_MISMATCH' });
  });

  it('requires a configured store, authenticated actor and complete store-bound catalogue', async () => {
    const req = await request();
    await expect(stageOpeningImport(db, req, { ...policy, storeId: '' })).rejects.toMatchObject({ code: 'OPENING_STORE_NOT_CONFIGURED' });
    await expect(stageOpeningImport(db, req, { ...policy, actor: '' })).rejects.toMatchObject({ code: 'INVALID_OPENING_ACTOR' });
    const input = stockInput(); input.catalog = targetsOf(input);
    await expect(stageOpeningImport(db, { ...req, input }, policy)).rejects.toMatchObject({ code: 'INVALID_OPENING_STAGE' });
    const wrong = stockInput(); catalogOf(wrong).store_id = '99';
    await expect(stageOpeningImport(db, { ...req, input: wrong }, policy)).rejects.toMatchObject({ code: 'OPENING_STORE_MISMATCH' });
    const incomplete = stockInput(); catalogOf(incomplete).complete = false;
    await expect(stageOpeningImport(db, { ...req, input: incomplete }, policy)).rejects.toMatchObject({ code: 'INVALID_CATALOGUE_SNAPSHOT' });
  });

  it.each([undefined, '', 'unknown', 'a'.repeat(63)])('requires original workbook hash provenance: %s', async snapshot_source_hash => {
    const req = await request(); req.input.snapshot_source_hash = snapshot_source_hash;
    await expect(stageOpeningImport(db, req, policy)).rejects.toMatchObject({ code: 'OPENING_SOURCE_HASH_REQUIRED' });
  });

  it.each(['001-BOLT', 'OTHER-SKU'])('rejects nonzero reservation summaries rather than silently losing order commitments for %s', async sku => {
    const input = { ...stockInput(), reservations: [{ sku, quantity: 2 }] };
    await expect(stageOpeningImport(db, await request(input), policy)).rejects.toMatchObject({ code: 'OPENING_RESERVATIONS_UNSUPPORTED' });
    expect(counts().items).toBe(0);
  });

  it('allows explicit zero reservations and hashes their reviewed presence', async () => {
    const input = { ...stockInput(), reservations: [{ sku: '001-BOLT', quantity: 0 }] };
    await expect(stageOpeningImport(db, await request(input), policy)).resolves.toMatchObject({ status: 'STAGED' });
  });

  it.each(['missing', 'extra', 'wrong-product', 'wrong-combination', 'wrong-options', 'case-changed', 'duplicate'])
  ('requires an exact, unique target scope: %s', async change => {
    const input = stockInput(); if (change === 'duplicate') addSecond(input);
    const req = await request(input);
    if (change === 'missing') req.scope = [];
    if (change === 'extra') req.scope.push(req.scope[0]);
    if (change === 'wrong-product') req.scope[0].ecwid_product_id = '999';
    if (change === 'wrong-combination') req.scope[0].ecwid_combination_id = '456';
    if (change === 'wrong-options') req.scope[0].ecwid_option_signature = '[{"name":"Size","value":"M3"}]';
    if (change === 'case-changed') req.scope[0].sku = '001-bolt';
    if (change === 'duplicate') req.scope[1] = req.scope[0];
    await expect(stageOpeningImport(db, req, policy)).rejects.toMatchObject({ code: 'OPENING_SCOPE_MISMATCH' });
    expect(counts().items).toBe(0);
  });

  it('binds SQL values and preserves names containing quotes and SQL text', async () => {
    const input = stockInput(); sourceOf(input)[0].name = "Bolt'); DROP TABLE items; --";
    await stageOpeningImport(db, await request(input), policy);
    expect(sqlite.prepare('SELECT name FROM items').pluck().get()).toBe(sourceOf(input)[0].name);
  });

  it('bounds set-based batches to 200 explicitly reviewed rows', async () => {
    const input = stockInput(); const original = targetsOf(input)[0];
    const targets = Array.from({ length: 201 }, (_, i) => ({ ...original, id: String(1000 + i), sku: `BOLT-${i}` }));
    Object.assign(catalogOf(input), { products: targets, stock_targets: targets, product_count: 201, stock_target_count: 201 });
    input.rows = targets.map(target => ({ sku: target.sku, balance: 10, single_unit_confirmed: true }));
    await expect(stageOpeningImport(db, await request(input), policy)).rejects.toMatchObject({ code: 'OPENING_STAGE_TOO_LARGE' });
    expect(counts().items).toBe(0);
  });
});
