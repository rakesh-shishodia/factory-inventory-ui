import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import { renderSourceReport } from '../scripts/source-report';
import type { ReviewedStockRow, StockSourceReview } from '../src/source-review';

function stockRow(overrides: Partial<ReviewedStockRow> = {}): ReviewedStockRow {
  return {
    source_row: 4, source_sheet: 'Stock Sheet', source_status: 'Active', item_type: 'Hardware',
    store_id: '0001', name: 'M3 bolt', source_sku: '001-bolt ', sku: '001-BOLT', unit: 'Pcs',
    location: 'Rack A', minimum: 2, opening: 10, inbound: 4, outbound: 1, balance: 13,
    formulas: { opening: null, inbound: null, outbound: null, balance: '=O4+P4-Q4' },
    audit_issues: [], status: 'CANDIDATE', current_authority: 'WORKBOOK',
    proposed_authority: 'APP_AFTER_APPROVAL',
    reasons: [{ code: 'SOURCE_CANDIDATE_ONLY', message: 'Confirm the Ecwid product and open orders next.' }],
    ...overrides,
  };
}

function report(): StockSourceReview {
  return {
    kind: 'SOURCE_REVIEW', dry_run: true, balance_meaning: 'PHYSICAL_ON_HAND',
    ecwid_checked: false, reservations_confirmed: false, ready_count: 0,
    source: { file_name: 'Stock review.xlsm', sha256: 'a'.repeat(64), sheet_name: 'Stock Sheet',
      header_row: 3, source_ref: 'https://docs.google.com/spreadsheets/d/example/edit#gid=100',
      source_modified_at: '2026-09-22T10:00:00Z', extracted_at: '2026-09-22T10:10:00Z' },
    generated_at: '2026-09-22T10:20:00Z',
    counts: { total: 3, candidate: 1, keep_workbook: 1, review: 1 },
    rows: [stockRow(), stockRow({ source_row: 5, store_id: '0002', name: 'Steel rod', source_sku: 'ROD', sku: 'ROD',
      status: 'KEEP_WORKBOOK', proposed_authority: 'WORKBOOK', unit: 'Mtr', balance: 2.5,
      reasons: [{ code: 'NON_UNIT', message: 'Meter-based items stay in the workbook for now.' }] }),
    stockRow({ source_row: 6, store_id: '0003', name: 'M4 bolt', source_sku: '004', sku: '004',
      status: 'REVIEW', proposed_authority: 'WORKBOOK', balance: -2,
      reasons: [{ code: 'NEGATIVE_BALANCE', message: 'Check the negative physical balance.' }] })],
    candidate_rows: [],
  };
}

describe('source stock report', () => {
  it('renders a self-contained source-only review with all counts and no approval claim', () => {
    const html = renderSourceReport(report());
    expect(html).toContain('<title>Opening stock review</title>');
    expect(html).toContain('data-count="total">3<');
    for (const name of ['candidate', 'review', 'keep-workbook']) expect(html).toContain(`data-count="${name}">1<`);
    expect(html).toContain('No stock has been changed. Every item is still workbook-managed.');
    expect(html).toContain('A candidate has passed source-data checks only.');
    expect(html).toContain('Balance means physical stock held in the factory');
    expect(html).toContain('Ecwid has not been checked');
    expect(html).toContain('outstanding unpicked orders');
    expect(html).not.toMatch(/data-status="READY"|>Ready<|>Imported<|Import complete/i);
    expect(html).not.toMatch(/<(?:script|link|img)\b[^>]*(?:src|href)=/i);
    expect(html).not.toMatch(/fetch\(|XMLHttpRequest|WebSocket|@import|url\(/);
    expect(html).toContain("connect-src 'none'");
    expect(html).toContain('overflow:auto');
  });

  it('preserves leading-zero source identifiers, source row references, and normalized SKU separately', () => {
    const html = renderSourceReport(report());
    expect(html).toContain('Stock Sheet · row 4');
    expect(html).toContain('Stock Sheet · row 6');
    expect(html).toContain('Store ID <code>0001</code>');
    expect(html).toContain('<code class="raw-sku">001-bolt </code>');
    expect(html).toContain('App match: <code>001-BOLT</code>');
    expect(html).toContain('<code class="raw-sku">004</code>');
    expect(html).toContain('<strong>2.5</strong>');
    expect(html).toContain('<strong>-2</strong>');
    expect(html).toContain('<code>' + 'a'.repeat(64) + '</code>');
    expect(html).toContain('2026-09-22T10:10:00Z');
    expect(html).toContain('Continue using workbook');
    expect(html).toContain('Proposed: app after matching and approval');
  });

  it('escapes workbook text in both table contents and search attributes, never in executable script', () => {
    const input = report();
    const attack = '</script><img src=x onerror="alert(1)"> & \' onclick="alert(2)';
    input.rows[0] = stockRow({ name: attack, source_sku: attack, sku: attack, store_id: attack,
      location: attack, unit: attack, item_type: attack, source_status: attack, balance: attack,
      reasons: [{ code: attack, message: attack }] });
    input.source.file_name = attack;
    input.source.sheet_name = attack;
    const html = renderSourceReport(input);
    expect(html).not.toContain(attack);
    expect(html).toContain('&lt;/script&gt;&lt;img src=x onerror=&quot;alert(1)&quot;&gt; &amp; &#39; onclick=&quot;alert(2)');
    expect(html.match(/<script>/g)).toHaveLength(1);
    expect(html.match(/<\/script>/g)).toHaveLength(1);
    expect(html).not.toContain('<img');
    const executable = html.match(/<script>([\s\S]*?)<\/script>/)![1];
    expect(executable).not.toContain('alert(');
    expect(executable).not.toContain('innerHTML');
  });

  it.each([
    'https://docs.google.com/spreadsheets/d/example/edit',
    'https://drive.google.com/file/d/example/view',
  ])('links only allowed Google source URL %s', sourceRef => {
    const input = report();
    input.source.source_ref = sourceRef;
    expect(renderSourceReport(input)).toContain(`href="${sourceRef}" target="_blank" rel="noopener noreferrer"`);
  });

  it.each([
    'javascript:alert(1)', 'data:text/html,<script>alert(1)</script>', 'http://docs.google.com/example',
    'https://docs.google.com.evil.example/example', 'https://drive.google.com@evil.example/example',
    'https://user:secret@docs.google.com/example', '//docs.google.com/example', 'file:///private/stock.xlsm',
    'https://example.com/report', '<a href="https://docs.google.com">source</a>',
  ])('leaves unsafe or unsupported source references unlinked: %s', sourceRef => {
    const input = report();
    input.source.source_ref = sourceRef;
    const html = renderSourceReport(input);
    expect(html).not.toContain('<a href=');
    expect(html).not.toContain('<script>alert(1)</script>');
  });

  it('defaults to review rows and includes accessible search, filter and scroll controls', () => {
    const html = renderSourceReport(report());
    expect(html).toContain('data-filter="REVIEW" aria-pressed="true"');
    expect(html).toContain('data-filter="ALL" aria-pressed="false"');
    expect(html).toContain('data-filter="CANDIDATE" aria-pressed="false"');
    expect(html).toContain('data-filter="KEEP_WORKBOOK" aria-pressed="false"');
    expect(html).toContain('id="stock-search"');
    expect(html).toContain('role="status" aria-live="polite"');
    expect(html).toContain('role="region" aria-label="Stock review table; scroll horizontally for all columns" tabindex="0"');
    const tags = html.match(/<tr data-stock-row[^>]*>/g)!;
    expect(tags).toHaveLength(3);
    expect(tags[0]).toMatch(/ hidden>$/);
    expect(tags[1]).toMatch(/ hidden>$/);
    expect(tags[2]).not.toContain(' hidden');
  });

  it('combines status filtering with case-insensitive SKU, Store ID, name and location search', () => {
    const html = renderSourceReport(report());
    const executable = html.match(/<script>([\s\S]*?)<\/script>/)![1];
    const events = new Map<string, () => void>();
    const search = { value: '', addEventListener: (name: string, callback: () => void) => events.set(name, callback) };
    const rows = [
      { dataset: { status: 'CANDIDATE', search: '001-bolt 001-BOLT M3 bolt 0001 Rack A' }, hidden: false },
      { dataset: { status: 'KEEP_WORKBOOK', search: 'ROD Steel rod 0002 Rack B' }, hidden: false },
      { dataset: { status: 'REVIEW', search: '004 M4 bolt 0003 Rack C' }, hidden: false },
    ];
    const filters = ['ALL', 'CANDIDATE', 'REVIEW', 'KEEP_WORKBOOK'].map(status => ({
      dataset: { filter: status }, attributes: {} as Record<string, string>,
      setAttribute(name: string, value: string) { this.attributes[name] = value; },
      addEventListener(_name: string, callback: () => void) { events.set(status, callback); },
    }));
    const visibleCount = { textContent: '' };
    const empty = { hidden: false };
    runInNewContext(executable, { document: {
      getElementById: (id: string) => ({ 'stock-search': search, 'visible-count': visibleCount, 'empty-state': empty })[id],
      querySelectorAll: (selector: string) => selector === '[data-stock-row]' ? rows : filters,
    } });
    expect(rows.map(row => row.hidden)).toEqual([true, true, false]);
    expect(visibleCount.textContent).toBe('1 item shown');
    events.get('ALL')!();
    expect(rows.every(row => !row.hidden)).toBe(true);
    expect(visibleCount.textContent).toBe('3 items shown');
    for (const query of ['001-bolt', 'm3 BOLT', '0001', 'rAcK a']) {
      search.value = query;
      events.get('input')!();
      expect(rows.map(row => row.hidden)).toEqual([false, true, true]);
    }
    events.get('REVIEW')!();
    expect(rows.every(row => row.hidden)).toBe(true);
    expect(empty.hidden).toBe(false);
    search.value = '';
    events.get('input')!();
    expect(empty.hidden).toBe(true);
    expect(filters[2].attributes['aria-pressed']).toBe('true');
    expect(filters[0].attributes['aria-pressed']).toBe('false');
  });
});
