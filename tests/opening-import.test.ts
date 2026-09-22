import { describe, expect, it } from 'vitest';
import { spawnSync } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseStockCsv, preparePreviewInput, previewImport } from '../src/opening-import';

const input = () => ({
  source_ref: 'Stock sheet export reviewed 2026-09-22',
  balance_meaning: 'PHYSICAL_ON_HAND', reservations_confirmed: true,
  rows: [{ sku: '001-BOLT', balance: '10', name: 'M3 bolt', single_unit_confirmed: true }],
  catalog: [{ id: '123', sku: '001-BOLT', name: 'M3 bolt', quantity: 7, unlimited: false, hasOptions: false, hasVariations: false,
    combinationId: null as string | null, variationOptions: [] as {name: string; value: string}[], enabled: true,
    eligibilityVerified: true, hasExtraOptions: false, hasBundleRelationships: false }],
  reservations: [{ sku: '001-BOLT', quantity: 2 }]
});

const catalogSnapshot = () => ({
  kind: 'READONLY_CATALOGUE', schema_version: 1, dry_run: true, complete: true, store_id: '2442119',
  started_at: '2026-09-22T08:00:00.000Z', completed_at: '2026-09-22T08:00:10.000Z',
  product_count: 1, stock_target_count: 1, reservations_confirmed: false,
  products: input().catalog, stock_targets: input().catalog
});

describe('opening stock preview', () => {
  it('accepts a complete live catalogue envelope and retains store and interval in the hashed preview', async () => {
    const first = await previewImport({ ...input(), catalog: catalogSnapshot(), store_id: '2442119' });
    expect(first.ready_count).toBe(1);
    expect(first.catalogue_snapshot).toEqual({ kind: 'READONLY_CATALOGUE', schema_version: 1, store_id: '2442119',
      started_at: '2026-09-22T08:00:00.000Z', completed_at: '2026-09-22T08:00:10.000Z', product_count: 1, stock_target_count: 1 });
    const later = await previewImport({ ...input(), catalog: { ...catalogSnapshot(), completed_at: '2026-09-22T08:00:11.000Z' }, store_id: '2442119' });
    expect(later.source_hash).not.toBe(first.source_hash);
    expect((await previewImport({ ...input(), catalog: { ...catalogSnapshot(), store_id: '999' } })).source_hash).not.toBe(first.source_hash);
  });
  it.each([
    { kind: 'APPROVED' }, { schema_version: 2 }, { dry_run: false }, { complete: false }, { reservations_confirmed: true },
    { store_id: 2442119 }, { store_id: '' }, { product_count: 2 }, { stock_target_count: 2 },
    { products: null }, { stock_targets: null }, { started_at: '' }, { completed_at: 'yesterday' },
    { completed_at: '2026-09-22T07:00:00.000Z' },
  ])('rejects incomplete or invalid catalogue snapshots: %o', async change => {
    await expect(previewImport({ ...input(), catalog: { ...catalogSnapshot(), ...change } })).rejects.toThrow();
  });
  it('rejects incompatible catalogue store IDs in direct and prepared previews', async () => {
    await expect(previewImport({ ...input(), catalog: catalogSnapshot(), store_id: '999' })).rejects.toThrow(/does not match/);
    const prepared = preparePreviewInput('sku,balance,single_unit_confirmed\n001-BOLT,10,true\n', 'CSV', catalogSnapshot(), {
      ...input(), store_id: '999'
    });
    await expect(previewImport(prepared)).rejects.toThrow(/does not match/);
    await expect(previewImport({ ...input(), store_id: 2442119 })).rejects.toThrow(/positive numeric string/);
  });
  it('uses physical stock less outstanding reservations, preserving leading-zero SKU', async () => {
    const result = await previewImport(input());
    expect(result.dry_run).toBe(true);
    expect(result.ready_count).toBe(1);
    expect(result.rows[0]).toMatchObject({ sku: '001-BOLT', physical_on_hand: 10, reserved: 2,
      desired_ecwid_quantity: 8, ecwid_quantity: 7, difference: 1, status: 'READY' });
  });
  it('cannot infer whether an unconfirmed balance includes reservations', async () => {
    const result = await previewImport({ ...input(), balance_meaning: 'UNKNOWN', reservations_confirmed: false });
    expect(result.global_errors).toHaveLength(2);
    expect(result.ready_count).toBe(0);
  });
  it.each(['', ' ', '-1', '1.5', 'NaN', '1,000', null, true])('blocks invalid balance %s instead of changing it to zero', async balance => {
    const result = await previewImport({ ...input(), rows: [{ sku: '001-BOLT', balance }] });
    expect(result.rows[0].physical_on_hand).toBeNull();
    expect(result.blocked_count).toBe(1);
  });
  it('accepts a legitimate zero opening balance', async () => {
    const result = await previewImport({ ...input(), rows: [{ ...input().rows[0], balance: '0' }], reservations: [] });
    expect(result.rows[0]).toMatchObject({ physical_on_hand: 0, desired_ecwid_quantity: 0, status: 'READY' });
  });
  it('rejects case-colliding source SKUs and shared QR namespaces', async () => {
    const result = await previewImport({ ...input(), rows: [{ sku: '001-bolt ', balance: 10 }, { sku: '001-BOLT', balance: 10 }] });
    expect(result.blocked_count).toBe(2);
    const qr = await previewImport({ ...input(), rows: [{ sku: '001-BOLT', balance: 10 }, { sku: 'OTHER', scan_code: '001-BOLT', balance: 20 }] });
    expect(qr.rows.every(row => row.errors.some(error => error.includes('another source row')))).toBe(true);
  });
  it('requires one Ecwid product per SKU and rejects duplicated IDs', async () => {
    const duplicateSku = await previewImport({ ...input(), catalog: [...input().catalog, { ...input().catalog[0], id: '456' }] });
    expect(duplicateSku.blocked_count).toBe(1);
    const duplicateId = await previewImport({ ...input(), catalog: [...input().catalog, { ...input().catalog[0], sku: 'OTHER' }] });
    expect(duplicateId.blocked_count).toBe(1);
  });
  it('blocks variations and untracked inventory', async () => {
    for (const change of [{ hasVariations: true }, { hasOptions: true }, { unlimited: true }]) {
      const result = await previewImport({ ...input(), catalog: [{ ...input().catalog[0], ...change }] });
      expect(result.blocked_count).toBe(1);
    }
  });
  it('accepts independently tracked variants and records exact target identity', async () => {
    const variant = { ...input().catalog[0], combinationId: '456', hasOptions: true, hasVariations: false,
      variationOptions: [{ name: 'Length', value: '12 mm' }] };
    const result = await previewImport({ ...input(), catalog: [variant] });
    expect(result.rows[0]).toMatchObject({ status: 'READY', ecwid_product_id: '123', ecwid_combination_id: '456',
      ecwid_option_signature: '[{"name":"Length","value":"12 mm"}]' });
  });
  it('permits multiple variants of one parent but never two targets sharing a SKU', async () => {
    const variant = { ...input().catalog[0], combinationId: '456', hasOptions: true, hasVariations: false,
      variationOptions: [{ name: 'Size', value: 'M3' }] };
    const second = { ...variant, combinationId: '457', sku: '002-BOLT', variationOptions: [{ name: 'Size', value: 'M4' }] };
    expect((await previewImport({ ...input(), catalog: [variant, second] })).ready_count).toBe(1);
    expect((await previewImport({ ...input(), catalog: [variant, { ...second, sku: variant.sku }] })).ready_count).toBe(0);
    expect((await previewImport({ ...input(), catalog: [variant, { ...second, combinationId: variant.combinationId }] })).ready_count).toBe(0);
  });
  it('ignores a variation container sharing its default child SKU', async () => {
    const parent = { ...input().catalog[0], hasOptions: true, hasVariations: true };
    const variant = { ...parent, combinationId: '456', hasVariations: false, variationOptions: [{ name: 'Size', value: 'M3' }] };
    const result = await previewImport({ ...input(), catalog: [parent, variant] });
    expect(result.rows[0]).toMatchObject({ status: 'READY', ecwid_combination_id: '456' });
  });
  it.each([
    { enabled: false }, { eligibilityVerified: undefined }, { hasExtraOptions: true }, { hasBundleRelationships: true },
    { variationOptions: [] }, { variationOptions: [{ name: 'Size', value: '' }] },
    { variationOptions: [{ name: 'Size', value: 'M3' }, { name: 'Size', value: 'M4' }] },
    { combinationId: '0' }, { combinationId: undefined }, { unlimited: true }, { quantity: null },
  ])('blocks a variant with unverified or unsafe eligibility %o', async change => {
    const variant = { ...input().catalog[0], combinationId: '456', hasOptions: true, hasVariations: false,
      variationOptions: [{ name: 'Size', value: 'M3' }], ...change };
    expect((await previewImport({ ...input(), catalog: [variant] })).ready_count).toBe(0);
  });
  it('requires per-SKU physical-unit confirmation, not a guessed unit from SKU or name', async () => {
    for (const value of [undefined, false, 'yes', 1]) {
      const report = await previewImport({ ...input(), rows: [{ ...input().rows[0], single_unit_confirmed: value }] });
      expect(report.rows[0].status).toBe('BLOCKED');
      expect(report.rows[0].errors.some(error => error.includes('one physical piece'))).toBe(true);
    }
  });
  it('blocks mismatched pre-existing variant mappings and includes options in hash', async () => {
    const variant = { ...input().catalog[0], combinationId: '456', hasOptions: true, hasVariations: false,
      variationOptions: [{ name: 'Size', value: 'M3' }] };
    const good = await previewImport({ ...input(), catalog: [variant] });
    const wrong = await previewImport({ ...input(), catalog: [variant], rows: [{ ...input().rows[0], ecwid_combination_id: '457' }] });
    expect(wrong.ready_count).toBe(0);
    expect((await previewImport({ ...input(), catalog: [{ ...variant, variationOptions: [{ name: 'Size', value: 'M4' }] }] })).source_hash).not.toBe(good.source_hash);
  });
  it('does not treat a catalog with omitted option/variation fields as verified simple stock', async () => {
    const result = await previewImport({ ...input(), catalog: [{ id: '123', sku: '001-BOLT', quantity: 7, unlimited: false }] });
    expect(result.blocked_count).toBe(1);
  });
  it('blocks negative availability and duplicate reservation summaries', async () => {
    const negative = await previewImport({ ...input(), reservations: [{ sku: '001-BOLT', quantity: 11 }] });
    expect(negative.blocked_count).toBe(1);
    const duplicate = await previewImport({ ...input(), reservations: [...input().reservations, ...input().reservations] });
    expect(duplicate.global_errors[0]).toContain('Duplicate reservation');
  });
  it('reports unmatched rows and leaves Ecwid-only products out of the plan', async () => {
    const result = await previewImport({ ...input(), rows: [{ sku: 'MISSING', balance: 10 }] });
    expect(result.rows[0].status).toBe('BLOCKED');
    expect(result.ecwid_only_skus).toEqual(['001-BOLT']);
  });
  it('uses a content hash stable across previews and different for changed data', async () => {
    const first = await previewImport(input());
    expect((await previewImport(input())).source_hash).toBe(first.source_hash);
    expect((await previewImport({ ...input(), rows: [{ sku: '001-BOLT', balance: 11 }] })).source_hash).not.toBe(first.source_hash);
  });
  it('preserves source workbook coordinates separately from the normalized preview row', async () => {
    const result = await previewImport({ ...input(), rows: [{ ...input().rows[0], source_row: 417, source_sheet: 'Stock Sheet' }] });
    expect(result.rows[0]).toMatchObject({ row: 2, source_row: 417, source_sheet: 'Stock Sheet', status: 'READY' });
    const withoutCoordinates = await previewImport(input());
    expect(withoutCoordinates.rows[0]).not.toHaveProperty('source_row');
    expect(withoutCoordinates.rows[0]).not.toHaveProperty('source_sheet');
  });
  it.each([0, -1, 2.5, '417', '', null, true, Number.MAX_SAFE_INTEGER + 1])('blocks invalid source row %s without guessing coordinates', async source_row => {
    const result = await previewImport({ ...input(), rows: [{ ...input().rows[0], source_row }] });
    expect(result.rows[0].status).toBe('BLOCKED');
    expect(result.rows[0].errors).toContain('Source row must be a positive whole-number row index, supplied as a number.');
    expect(result.rows[0]).not.toHaveProperty('source_row');
  });
  it.each([42, null, true, 'x'.repeat(201)])('blocks invalid source sheet %s', async source_sheet => {
    const result = await previewImport({ ...input(), rows: [{ ...input().rows[0], source_sheet }] });
    expect(result.rows[0].status).toBe('BLOCKED');
    expect(result.rows[0].errors).toContain('Source sheet must be a string of at most 200 characters.');
  });
  it('accepts a 200-character source sheet name', async () => {
    const result = await previewImport({ ...input(), rows: [{ ...input().rows[0], source_sheet: 'x'.repeat(200) }] });
    expect(result.rows[0].status).toBe('READY');
  });
  it('keeps the snapshot hash separate and includes it in the preview-content hash', async () => {
    const first = await previewImport({ ...input(), snapshot_source_hash: 'a'.repeat(64) });
    const second = await previewImport({ ...input(), snapshot_source_hash: 'b'.repeat(64) });
    expect(first.snapshot_source_hash).toBe('a'.repeat(64));
    expect(first.source_hash).not.toBe(first.snapshot_source_hash);
    expect(second.source_hash).not.toBe(first.source_hash);
    expect((await previewImport(input()))).not.toHaveProperty('snapshot_source_hash');
  });
  it('blocks an invalid snapshot hash in a direct preview request', async () => {
    const result = await previewImport({ ...input(), snapshot_source_hash: 'not-a-sha256' });
    expect(result.ready_count).toBe(0);
    expect(result.global_errors).toContain('Snapshot source hash must be a SHA-256 hexadecimal digest.');
  });
});

describe('staged source-candidate input', () => {
  const candidate = () => ({
    kind: 'SOURCE_CANDIDATES', dry_run: true, source_ref: 'Drive workbook snapshot', source_hash: 'a'.repeat(64),
    balance_meaning: 'PHYSICAL_ON_HAND', ecwid_checked: false, reservations_confirmed: false,
    rows: [{ ...input().rows[0], source_row: 417, source_sheet: 'Stock Sheet' }]
  });
  const manifest = () => ({
    source_ref: 'Drive workbook snapshot', balance_meaning: 'PHYSICAL_ON_HAND',
    reservations_confirmed: true, reservations: input().reservations
  });
  const prepare = (source: unknown = candidate(), metadata: unknown = manifest()) =>
    preparePreviewInput(JSON.stringify(source), 'SOURCE_CANDIDATES', input().catalog, metadata);

  it('joins source coordinates to a separately supplied catalog and reservation manifest', async () => {
    const result = await previewImport(prepare());
    expect(result).toMatchObject({ source_ref: 'Drive workbook snapshot', snapshot_source_hash: 'a'.repeat(64), ready_count: 1, dry_run: true });
    expect(result.rows[0]).toMatchObject({ row: 2, source_row: 417, source_sheet: 'Stock Sheet', physical_on_hand: 10, reserved: 2, desired_ecwid_quantity: 8 });
  });
  it('allows an omitted or blank manifest reference, retaining the actual source reference', () => {
    for (const source_ref of [undefined, '', '  ', ' Drive workbook snapshot ']) {
      expect(prepare(candidate(), { ...manifest(), source_ref }).source_ref).toBe('Drive workbook snapshot');
    }
  });
  it.each(['Other workbook', 123, null])('rejects conflicting or invalid manifest source reference %s', source_ref => {
    expect(() => prepare(candidate(), { ...manifest(), source_ref })).toThrow(/source_ref does not match/);
  });
  it('cannot auto-confirm reservations through the source file', async () => {
    const result = await previewImport(prepare(candidate(), { ...manifest(), reservations_confirmed: false }));
    expect(result.ready_count).toBe(0);
    expect(result.global_errors).toContain('Confirm the outstanding unpicked quantities of all open Ecwid orders before alignment.');
    expect(() => prepare({ ...candidate(), reservations_confirmed: true })).toThrow(/Expected dry-run/);
  });
  it.each([
    { kind: 'READY_TO_IMPORT' }, { dry_run: false }, { ecwid_checked: true },
    { balance_meaning: 'AVAILABLE_TO_SELL' }, { catalog: input().catalog }, { reservations: [] },
    { csv: 'sku,balance\nOTHER,1\n' }, { extra: 'unexpected field' }
  ])('rejects candidate fields that could change the staging contract: %o', change => {
    expect(() => prepare({ ...candidate(), ...change })).toThrow(/Expected dry-run/);
  });
  it.each([{ source_ref: '' }, { source_ref: 123 }, { source_hash: 'abc' }, { source_hash: null }])('requires valid snapshot identity: %o', change => {
    expect(() => prepare({ ...candidate(), ...change })).toThrow(/nonblank source_ref and a SHA-256/);
  });
  it.each([null, [], 'stock data'])('rejects a non-object candidate payload', source => {
    expect(() => prepare(source)).toThrow(/must contain an object/);
  });
  it('rejects malformed candidate JSON and malformed rows', () => {
    expect(() => preparePreviewInput('{oops', 'SOURCE_CANDIDATES', [], manifest())).toThrow(/valid JSON/);
    expect(() => prepare({ ...candidate(), rows: ['not a row'] })).toThrow(/rows must be an array/);
  });
  it('never changes the physical balance meaning to satisfy the manifest', () => {
    expect(() => prepare(candidate(), { ...manifest(), balance_meaning: 'AVAILABLE_TO_SELL' })).toThrow(/PHYSICAL_ON_HAND/);
    expect(prepare(candidate(), { ...manifest(), balance_meaning: undefined }).balance_meaning).toBe('PHYSICAL_ON_HAND');
  });
  it('still applies ordinary preview row validation after accepting the staging envelope', async () => {
    const result = await previewImport(prepare({ ...candidate(), rows: [{ sku: '001-BOLT', balance: -2, source_row: '417' }] }));
    expect(result.blocked_count).toBe(1);
    expect(result.rows[0].errors).toContain('Balance must be a nonblank, nonnegative whole number.');
    expect(result.rows[0].errors).toContain('Source row must be a positive whole-number row index, supplied as a number.');
  });
  it('ignores manifest data overrides and uses the independently supplied source and catalog', async () => {
    const result = await previewImport(prepare(candidate(), {
      ...manifest(), rows: [], catalog: [], csv: 'sku,balance\nOTHER,1000\n', snapshot_source_hash: 'b'.repeat(64)
    }));
    expect(result.snapshot_source_hash).toBe('a'.repeat(64));
    expect(result.rows[0]).toMatchObject({ sku: '001-BOLT', physical_on_hand: 10, ecwid_product_id: '123' });
  });
  it('retains plain CSV compatibility and prevents manifest csv from replacing the source file', async () => {
    const report = await previewImport(preparePreviewInput('sku,balance,single_unit_confirmed\n001-BOLT,10,true\n', 'CSV', input().catalog, {
      ...manifest(), csv: 'sku,balance\nOTHER,999\n'
    }));
    expect(report.ready_count).toBe(1);
    expect(report.rows[0]).toMatchObject({ sku: '001-BOLT', physical_on_hand: 10, reserved: 2 });
    expect(report.rows[0]).not.toHaveProperty('source_row');
  });
  it('runs the preview CLI for candidate JSON and CSV with distinct safe exit codes', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'inventory-preview-test-'));
    try {
      const stockPath = join(directory, 'candidates.json');
      const csvPath = join(directory, 'stock.csv');
      const catalogPath = join(directory, 'catalog.json');
      const manifestPath = join(directory, 'reservations.json');
      await Promise.all([
        writeFile(stockPath, JSON.stringify(candidate())),
        writeFile(csvPath, 'sku,balance,single_unit_confirmed\n001-BOLT,10,true\n'),
        writeFile(catalogPath, JSON.stringify(input().catalog)),
        writeFile(manifestPath, JSON.stringify(manifest()))
      ]);
      const run = (path: string) => spawnSync(process.execPath,
        ['--import', 'tsx', 'scripts/preview-import.ts', path, catalogPath, manifestPath],
        { encoding: 'utf8', timeout: 10_000 });
      const staged = run(stockPath);
      expect(staged.status, staged.stderr).toBe(0);
      expect(JSON.parse(staged.stdout)).toMatchObject({ ready_count: 1, snapshot_source_hash: 'a'.repeat(64), rows: [{ source_row: 417 }] });
      const csv = run(csvPath);
      expect(csv.status, csv.stderr).toBe(0);
      expect(JSON.parse(csv.stdout).ready_count).toBe(1);

      await writeFile(catalogPath, JSON.stringify(catalogSnapshot()));
      const envelope = run(csvPath);
      expect(envelope.status, envelope.stderr).toBe(0);
      expect(JSON.parse(envelope.stdout)).toMatchObject({ ready_count: 1, catalogue_snapshot: { store_id: '2442119' } });
      await writeFile(manifestPath, JSON.stringify({ ...manifest(), store_id: '999' }));
      const wrongStore = run(csvPath);
      expect(wrongStore.status).toBe(1);
      expect(wrongStore.stderr).toContain('does not match');

      await writeFile(manifestPath, JSON.stringify({ ...manifest(), reservations_confirmed: false }));
      const blocked = run(stockPath);
      expect(blocked.status, blocked.stderr).toBe(2);
      expect(JSON.parse(blocked.stdout).blocked_count).toBe(1);

      await writeFile(manifestPath, JSON.stringify({ ...manifest(), source_ref: 'Different workbook' }));
      const mismatched = run(stockPath);
      expect(mismatched.status).toBe(1);
      expect(mismatched.stdout).toBe('');
      expect(mismatched.stderr).toContain('source_ref does not match');

      await writeFile(csvPath, 'x'.repeat(2_000_001));
      const oversized = run(csvPath);
      expect(oversized.status).toBe(1);
      expect(oversized.stdout).toBe('');
      expect(oversized.stderr).toContain('size limit');
      await writeFile(csvPath, Buffer.from([0xff]));
      const invalidUtf8 = run(csvPath);
      expect(invalidUtf8.status).toBe(1);
      expect(invalidUtf8.stdout).toBe('');
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });
});

describe('stock CSV parsing', () => {
  it('supports BOM, commas, multiline quotes, escaped quotes, CRLF and leading zeros', () => {
    expect(parseStockCsv('\uFEFFSKU,Name,Balance\r\n001,"Bolt, \"\"long\"\"\nM3",10\r\n')).toEqual([
      { sku: '001', name: 'Bolt, "long"\nM3', balance: '10' }
    ]);
  });
  it('does not drop a row with missing balance', () => {
    expect(parseStockCsv('SKU,Balance\nA,\n')).toEqual([{ sku: 'A', balance: '' }]);
  });
  it.each(['SKU,SKU,Balance\nA,A,1', 'SKU,Balance\nA,1,2', 'SKU,Balance\nA,"1', 'Name,Balance\nBolt,1'])('rejects malformed export', csv => {
    expect(() => parseStockCsv(csv)).toThrow();
  });
});
