import { describe, expect, it } from 'vitest';
import { reviewStockSource, type StockSnapshot, type StockSourceRow } from '../src/source-review';

function sourceRow(change: Partial<StockSourceRow> = {}): StockSourceRow {
  return {
    source_row: 4, status: 'Active', item_type: 'Single', store_id: 'STORE-001', name: 'M3 bolt',
    sku: '001-bolt', unit: 'Pcs', location: 'Rack A', minimum: 2, opening: 8, inbound: 4, outbound: 2,
    balance: 10, formulas: { opening: null, inbound: '=SUMIF(Inward!B:B,D4,Inward!E:E)', outbound: '=SUMIF(Outward!B:B,D4,Outward!E:E)', balance: '=M4+N4-O4' },
    audit_issues: [], ...change,
  };
}

function snapshot(rows: StockSourceRow[] = [sourceRow()]): StockSnapshot {
  return { schema_version: 1, source: { file_name: 'synthetic-stock.xlsm', sha256: 'a'.repeat(64),
    sheet_name: 'Stock Sheet', header_row: 3, source_ref: 'synthetic-test-fixture',
    source_modified_at: '2026-09-20T09:00:00Z', extracted_at: '2026-09-22T09:00:00Z' }, rows };
}

const codes = (row: StockSourceRow) => reviewStockSource(snapshot([row])).rows[0].reasons.map(issue => issue.code);

describe('read-only source stock review', () => {
  it('labels valid simple stock as a candidate, never ready or already app-owned', () => {
    const input = snapshot();
    const result = reviewStockSource(input);
    expect(result).toMatchObject({ kind: 'SOURCE_REVIEW', dry_run: true, source: input.source,
      balance_meaning: 'PHYSICAL_ON_HAND', ecwid_checked: false, reservations_confirmed: false,
      ready_count: 0, counts: { total: 1, candidate: 1, keep_workbook: 0, review: 0 } });
    expect(result.rows[0]).toMatchObject({ source_row: 4, source_sheet: 'Stock Sheet', source_status: 'Active',
      source_sku: '001-bolt', sku: '001-BOLT', balance: 10, status: 'CANDIDATE', current_authority: 'WORKBOOK',
      proposed_authority: 'APP_AFTER_APPROVAL', formulas: input.rows[0].formulas });
    expect(result.rows[0].reasons[0].message).toContain('outstanding orders');
    expect(result.candidate_rows).toEqual([{ source_row: 4, source_sheet: 'Stock Sheet', sku: '001-BOLT',
      scan_code: '001-BOLT', balance: 10, name: 'M3 bolt', location: 'Rack A' }]);
    expect(result.candidate_rows[0]).not.toHaveProperty('ecwid_product_id');
    expect(result.candidate_rows[0]).not.toHaveProperty('reserved');
    expect(Date.parse(result.generated_at)).not.toBeNaN();
  });

  it('normalizes eligibility case and spaces while retaining original cell values and leading zeros', () => {
    const row = sourceRow({ status: ' active ', item_type: ' SINGLE ', unit: ' pcs ', sku: ' 0001 ', balance: ' 12.0 ', name: ' Bolt ', location: ' A ' });
    const result = reviewStockSource(snapshot([row]));
    expect(result.rows[0]).toMatchObject({ source_status: ' active ', source_sku: ' 0001 ', item_type: ' SINGLE ', balance: ' 12.0 ', sku: '0001' });
    expect(result.candidate_rows[0]).toMatchObject({ sku: '0001', scan_code: '0001', balance: 12, name: 'Bolt', location: 'A' });
  });

  it.each([0, '0', 2147483647])('allows valid boundary balance %s', balance => {
    expect(reviewStockSource(snapshot([sourceRow({ balance })])).counts.candidate).toBe(1);
  });

  it.each(['NA', 'N/A', ' na ', ' n/a ', '', '  ', null])('keeps unmapped SKU %s workbook-owned without claiming it is absent from Ecwid', sku => {
    const result = reviewStockSource(snapshot([sourceRow({ sku })]));
    expect(result.rows[0]).toMatchObject({ status: 'KEEP_WORKBOOK', current_authority: 'WORKBOOK', proposed_authority: 'WORKBOOK' });
    expect(result.rows[0].reasons).toContainEqual({ code: 'SKU_UNMAPPED', message: expect.stringContaining('Ecwid listing status has not been checked') });
    expect(result.candidate_rows).toEqual([]);
  });

  it('does not treat repeated placeholder SKUs as duplicate product mappings', () => {
    const result = reviewStockSource(snapshot([
      sourceRow({ sku: 'NA' }),
      sourceRow({ source_row: 5, store_id: 'STORE-002', sku: ' na ' }),
      sourceRow({ source_row: 6, store_id: 'STORE-003', sku: '' }),
      sourceRow({ source_row: 7, store_id: 'STORE-004', sku: null }),
    ]));
    expect(result.counts).toEqual({ total: 4, candidate: 0, keep_workbook: 4, review: 0 });
    expect(result.rows.flatMap(row => row.reasons).some(issue => issue.code === 'DUPLICATE_SKU')).toBe(false);
  });

  it.each([
    { status: 'Inactive' }, { status: null }, { item_type: 'Composite' }, { item_type: 'Unknown' },
    { item_type: null }, { unit: 'Meters', balance: 2.25 }, { unit: 'Kg', balance: '.5' }, { unit: null },
  ])('defers unsupported or unconfirmed scope without rounding quantities: %j', change => {
    const result = reviewStockSource(snapshot([sourceRow(change)]));
    expect(result.rows[0].status).toBe('KEEP_WORKBOOK');
    expect(result.rows[0].proposed_authority).toBe('WORKBOOK');
    expect(result.candidate_rows).toEqual([]);
    expect(result.rows[0].balance).toBe(change.balance ?? 10);
  });

  it('checks SKU uniqueness across all rows, including deferred ones', () => {
    const result = reviewStockSource(snapshot([
      sourceRow(), sourceRow({ source_row: 5, store_id: 'STORE-002', sku: ' 001-BOLT ', unit: 'Meters' }),
    ]));
    expect(result.counts).toEqual({ total: 2, candidate: 0, keep_workbook: 0, review: 2 });
    for (const row of result.rows) expect(row.reasons.map(issue => issue.code)).toContain('DUPLICATE_SKU');
    expect(result.rows[0].balance).toBe(10);
    expect(result.rows[1].balance).toBe(10);
  });

  it('detects Store ID collisions across scope and normalizes case/space and numeric representation', () => {
    const result = reviewStockSource(snapshot([
      sourceRow({ store_id: 123 }), sourceRow({ source_row: 5, store_id: ' 123 ', sku: 'NA', item_type: 'Composite' }),
      sourceRow({ source_row: 6, store_id: ' Store-A ', sku: 'NEW-1' }),
      sourceRow({ source_row: 7, store_id: 'store-a', sku: 'NEW-2' }),
    ]));
    expect(result.counts.review).toBe(4);
    for (const row of result.rows) expect(row.reasons.map(issue => issue.code)).toContain('DUPLICATE_STORE_ID');
  });

  it.each([null, '', 'NA', 'N/A', false, 0, -1, 1.5, 'A|B', 'A\nB'])('flags invalid Store ID %s', store_id => {
    expect(codes(sourceRow({ store_id }))).toContain('INVALID_STORE_ID');
  });

  it.each([null, '', ' ', false, 123, 'bad\u0000name', 'a'.repeat(1001)])('flags an invalid item name', name => {
    expect(codes(sourceRow({ name }))).toContain('INVALID_NAME');
  });

  it.each([123, 0, 1.5, false])('blocks a nontext SKU %s, without inventing lost zeros', sku => {
    const result = reviewStockSource(snapshot([sourceRow({ sku })]));
    expect(result.rows[0]).toMatchObject({ status: 'REVIEW', source_sku: sku, sku: '' });
    expect(result.rows[0].reasons.map(issue => issue.code)).toContain('SKU_NOT_TEXT');
  });

  it('also blocks text SKU colliding with an ineligible numeric SKU', () => {
    const result = reviewStockSource(snapshot([
      sourceRow({ sku: 123 }), sourceRow({ source_row: 5, store_id: 'STORE-002', sku: '123' }),
    ]));
    expect(result.counts.review).toBe(2);
    expect(result.rows[1].reasons.map(issue => issue.code)).toContain('DUPLICATE_SKU');
  });

  it.each(['ABC|DEF', 'ABC\nDEF', 'ABC\u0000DEF', 'a'.repeat(201)])('blocks an invalid scan/SKU namespace', sku => {
    expect(codes(sourceRow({ sku }))).toContain('INVALID_SKU');
  });

  it.each(['', ' ', null, true, NaN, Infinity, 'NaN', '#VALUE!', '1,000', '1e3'])('flags invalid balance %s instead of using zero', balance => {
    const result = reviewStockSource(snapshot([sourceRow({ balance })]));
    expect(result.rows[0].status).toBe('REVIEW');
    expect(result.rows[0].reasons.map(issue => issue.code)).toContain('INVALID_BALANCE');
  });

  it.each([-1, '-0.5'])('requires reconciliation for negative balance %s, even outside initial scope', balance => {
    expect(codes(sourceRow({ balance, unit: 'Meters', item_type: 'Composite', sku: 'NA' }))).toContain('NEGATIVE_BALANCE');
    expect(reviewStockSource(snapshot([sourceRow({ balance, unit: 'Meters' })])).rows[0].status).toBe('REVIEW');
  });

  it.each([1.5, '.5', '10.25', '10.00000000000000001', `0.${'0'.repeat(400)}1`])('flags fractional Pcs balance %s without rounding', balance => {
    const result = reviewStockSource(snapshot([sourceRow({ balance })]));
    expect(result.rows[0].balance).toBe(balance);
    expect(result.rows[0].reasons.map(issue => issue.code)).toContain('FRACTIONAL_PCS_BALANCE');
  });

  it.each([2147483648, '2147483648', Number.MAX_SAFE_INTEGER, '2147483647.000000000000001'])('rejects overflow balance %s', balance => {
    expect(codes(sourceRow({ balance }))).toContain('BALANCE_TOO_LARGE');
  });

  it('does not silently underflow tiny negative text balances to zero', () => {
    expect(codes(sourceRow({ balance: `-0.${'0'.repeat(400)}1` }))).toContain('NEGATIVE_BALANCE');
  });

  it('preserves extractor audit findings and requires review even for otherwise deferred rows', () => {
    const audit = { code: 'WRONG_MOVEMENT_RANGE', message: 'Outbound formula refers to the wrong movement column.' };
    const result = reviewStockSource(snapshot([sourceRow({ sku: 'NA', unit: 'Meters', audit_issues: [audit] })]));
    expect(result.rows[0]).toMatchObject({ status: 'REVIEW', audit_issues: [audit], proposed_authority: 'WORKBOOK' });
    expect(result.rows[0].reasons).toContainEqual(audit);
  });

  it('accounts for every row across candidate, workbook and review classifications', () => {
    const result = reviewStockSource(snapshot([
      sourceRow(), sourceRow({ source_row: 5, store_id: 'STORE-002', sku: 'NA' }),
      sourceRow({ source_row: 6, store_id: 'STORE-003', sku: 'BOLT2', balance: -2 }),
    ]));
    expect(result.counts).toEqual({ total: 3, candidate: 1, keep_workbook: 1, review: 1 });
    expect(result.rows.map(row => row.source_row)).toEqual([4, 5, 6]);
    expect(result.rows.every(row => row.current_authority === 'WORKBOOK')).toBe(true);
  });

  it('does not mutate input data or return mutable references to its nested provenance', () => {
    const input = snapshot([sourceRow({ audit_issues: [{ code: 'AUDIT', message: 'Check balance.' }] })]);
    const before = JSON.stringify(input);
    const result = reviewStockSource(input);
    result.source.file_name = 'different';
    result.rows[0].formulas.balance = 'changed';
    result.rows[0].audit_issues[0].message = 'changed';
    result.rows[0].reasons[0].message = 'changed';
    expect(JSON.stringify(input)).toBe(before);
  });

  it('accepts an empty source without manufacturing ready items', () => {
    expect(reviewStockSource(snapshot([]))).toMatchObject({ ready_count: 0, candidate_rows: [], rows: [], counts: { total: 0, candidate: 0, keep_workbook: 0, review: 0 } });
  });
});

describe('source snapshot validation and bounds', () => {
  it.each([null, [], {}, { schema_version: 2 }, { ...snapshot(), rows: 'rows' }])('rejects malformed snapshot', input => {
    expect(() => reviewStockSource(input)).toThrow();
  });

  it.each([
    { sha256: 'not-a-hash' }, { file_name: '' }, { sheet_name: null }, { source_ref: ' ' },
    { header_row: 0 }, { header_row: '3' }, { header_row: 1.5 }, { header_row: 1048576 },
    { extracted_at: '2026-09-22' }, { source_modified_at: 'not-a-date' },
  ])('rejects malformed metadata %j', change => {
    const input = snapshot();
    expect(() => reviewStockSource({ ...input, source: { ...input.source, ...change } })).toThrow();
  });

  it.each([undefined, null, '4', 0, 3, 4.5, 1048577])('does not silently drop invalid source row %s', source_row => {
    expect(() => reviewStockSource({ ...snapshot(), rows: [{ ...sourceRow(), source_row }] })).toThrow();
  });

  it('rejects repeated worksheet rows even when the SKU differs', () => {
    expect(() => reviewStockSource(snapshot([sourceRow(), sourceRow({ sku: 'OTHER', store_id: 'STORE-002' })]))).toThrow('Duplicate source_row');
  });

  it.each([
    { name: { value: 'Bolt' } }, { balance: [] }, { location: undefined }, { formulas: null },
    { formulas: { ...sourceRow().formulas, balance: undefined } }, { audit_issues: null },
    { audit_issues: [{ code: '', message: 'Missing code' }] }, { audit_issues: [{ code: 'X', message: {} }] },
  ])('rejects malformed row payload %j', change => {
    expect(() => reviewStockSource({ ...snapshot(), rows: [{ ...sourceRow(), ...change }] })).toThrow();
  });

  it('bounds row count, individual text and number of audit entries', () => {
    expect(() => reviewStockSource(snapshot(Array.from({ length: 10001 }, (_, index) => sourceRow({ source_row: index + 4 }))))).toThrow('10,000');
    expect(() => reviewStockSource(snapshot([sourceRow({ name: 'a'.repeat(10001) })]))).toThrow('10000');
    expect(() => reviewStockSource(snapshot([sourceRow({ audit_issues: Array.from({ length: 101 }, () => ({ code: 'X', message: 'x' })) })]))).toThrow('at most 100');
  });

  it('bounds accumulated text and collision messages', () => {
    const longRows = Array.from({ length: 801 }, (_, index) => sourceRow({ source_row: index + 4, name: 'a'.repeat(10000) }));
    expect(() => reviewStockSource(snapshot(longRows))).toThrow('review size limit');
    const duplicateRows = Array.from({ length: 100 }, (_, index) => sourceRow({ source_row: index + 4 }));
    const result = reviewStockSource(snapshot(duplicateRows));
    expect(result.counts.review).toBe(100);
    expect(result.rows[0].reasons.find(issue => issue.code === 'DUPLICATE_SKU')?.message).toContain('89 more');
    expect(result.rows[0].reasons.every(issue => issue.message.length < 500)).toBe(true);
  });
});
