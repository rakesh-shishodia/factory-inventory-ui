import { DomainError, normalizeCode } from './domain';

export interface SourceIssue { code: string; message: string }

export interface StockSourceRow {
  source_row: number;
  status: unknown;
  item_type: unknown;
  store_id: unknown;
  name: unknown;
  sku: unknown;
  unit: unknown;
  location: unknown;
  minimum: unknown;
  opening: unknown;
  inbound: unknown;
  outbound: unknown;
  balance: unknown;
  formulas: { opening: string | null; inbound: string | null; outbound: string | null; balance: string | null };
  audit_issues: SourceIssue[];
}

export interface StockSnapshot {
  schema_version: 1;
  source: {
    file_name: string;
    sha256: string;
    sheet_name: string;
    header_row: number;
    source_ref: string;
    source_modified_at?: string;
    extracted_at: string;
  };
  rows: StockSourceRow[];
}

export interface ReviewedStockRow extends Omit<StockSourceRow, 'status' | 'sku'> {
  source_sheet: string;
  source_status: unknown;
  source_sku: unknown;
  sku: string;
  status: 'CANDIDATE' | 'KEEP_WORKBOOK' | 'REVIEW';
  current_authority: 'WORKBOOK';
  proposed_authority: 'APP_AFTER_APPROVAL' | 'WORKBOOK';
  reasons: SourceIssue[];
}

export interface StockCandidateRow {
  sku: string;
  balance: number;
  name: string;
  location: string;
  scan_code: string;
  source_row: number;
  source_sheet: string;
}

export interface StockSourceReview {
  kind: 'SOURCE_REVIEW';
  dry_run: true;
  source: StockSnapshot['source'];
  generated_at: string;
  balance_meaning: 'PHYSICAL_ON_HAND';
  ecwid_checked: false;
  reservations_confirmed: false;
  ready_count: 0;
  counts: { total: number; candidate: number; keep_workbook: number; review: number };
  rows: ReviewedStockRow[];
  candidate_rows: StockCandidateRow[];
}

const MAX_ROWS = 10000;
const MAX_TOTAL_CHARS = 8_000_000;
const MAX_QUANTITY = 2147483647;
const MAX_SHEET_ROW = 1048576;
const CELL_FIELDS = ['status', 'item_type', 'store_id', 'name', 'sku', 'unit', 'location', 'minimum', 'opening', 'inbound', 'outbound', 'balance'] as const;
const FORMULA_FIELDS = ['opening', 'inbound', 'outbound', 'balance'] as const;
const placeholder = (value: string) => !value || value === 'NA' || value === 'N/A';
const normalizedText = (value: unknown) => typeof value === 'string' ? normalizeCode(value) : '';

function invalid(message: string): never {
  throw new DomainError(400, 'INVALID_SOURCE_SNAPSHOT', message);
}

function record(value: unknown, label: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) invalid(`${label} must be an object.`);
  return value as Record<string, unknown>;
}

/** Validate the complete snapshot before classifying anything; no malformed rows are dropped. */
function parseSnapshot(input: unknown): StockSnapshot {
  const snapshot = record(input, 'Snapshot');
  if (snapshot.schema_version !== 1) invalid('Snapshot schema_version must be 1.');
  const source = record(snapshot.source, 'Snapshot source');
  let totalChars = 0;
  function text(value: unknown, label: string, max: number, allowBlank = false): string {
    if (typeof value !== 'string' || (!allowBlank && !value.trim())) invalid(`${label} must be a ${allowBlank ? '' : 'nonblank '}string.`);
    if (value.length > max) throw new DomainError(413, 'SOURCE_SNAPSHOT_TOO_LARGE', `${label} exceeds ${max} characters.`);
    totalChars += value.length;
    if (totalChars > MAX_TOTAL_CHARS) throw new DomainError(413, 'SOURCE_SNAPSHOT_TOO_LARGE', 'The stock snapshot exceeds the review size limit.');
    return value;
  }
  const fileName = text(source.file_name, 'Source file_name', 1000);
  const sheetName = text(source.sheet_name, 'Source sheet_name', 200);
  const sourceRef = text(source.source_ref, 'Source source_ref', 4000);
  const sha256 = text(source.sha256, 'Source sha256', 64);
  if (!/^[0-9a-f]{64}$/i.test(sha256)) invalid('Source sha256 must be a SHA-256 hex digest of the original workbook.');
  const headerRow = source.header_row;
  if (!Number.isInteger(headerRow) || (headerRow as number) < 1 || (headerRow as number) >= MAX_SHEET_ROW) invalid('Source header_row must be a valid worksheet row number.');
  function timestamp(value: unknown, label: string): string {
    const result = text(value, label, 100);
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/.test(result) || !Number.isFinite(Date.parse(result))) invalid(`${label} must be an ISO timestamp with a timezone.`);
    return result;
  }
  const extractedAt = timestamp(source.extracted_at, 'Source extracted_at');
  const modifiedAt = source.source_modified_at === undefined ? undefined : timestamp(source.source_modified_at, 'Source source_modified_at');
  if (!Array.isArray(snapshot.rows)) invalid('Snapshot rows must be an array.');
  if (snapshot.rows.length > MAX_ROWS) throw new DomainError(413, 'SOURCE_SNAPSHOT_TOO_LARGE', 'Split snapshots larger than 10,000 rows before reviewing.');
  const sourceRows = new Set<number>();
  const rows = snapshot.rows.map((value, index): StockSourceRow => {
    const row = record(value, `Snapshot row ${index + 1}`);
    const sourceRow = row.source_row;
    if (!Number.isInteger(sourceRow) || (sourceRow as number) <= (headerRow as number) || (sourceRow as number) > MAX_SHEET_ROW) invalid(`Snapshot row ${index + 1} needs a source_row after the header and within the worksheet.`);
    if (sourceRows.has(sourceRow as number)) invalid(`Duplicate source_row ${sourceRow}; every worksheet row must be represented only once.`);
    sourceRows.add(sourceRow as number);
    for (const field of CELL_FIELDS) {
      const cell = row[field];
      if (cell !== null && !['string', 'number', 'boolean'].includes(typeof cell)) invalid(`Source row ${sourceRow} ${field} must be a scalar cell value or null.`);
      if (typeof cell === 'string') text(cell, `Source row ${sourceRow} ${field}`, 10000, true);
    }
    const formulas = record(row.formulas, `Source row ${sourceRow} formulas`);
    for (const field of FORMULA_FIELDS) {
      if (formulas[field] !== null) text(formulas[field], `Source row ${sourceRow} ${field} formula`, 10000);
    }
    if (!Array.isArray(row.audit_issues) || row.audit_issues.length > 100) invalid(`Source row ${sourceRow} audit_issues must be an array of at most 100 issues.`);
    const auditIssues = row.audit_issues.map((issue, issueIndex): SourceIssue => {
      const entry = record(issue, `Source row ${sourceRow} audit issue ${issueIndex + 1}`);
      return {
        code: text(entry.code, `Source row ${sourceRow} audit issue code`, 200),
        message: text(entry.message, `Source row ${sourceRow} audit issue message`, 4000),
      };
    });
    return {
      source_row: sourceRow as number,
      ...Object.fromEntries(CELL_FIELDS.map(field => [field, row[field]])) as Pick<StockSourceRow, typeof CELL_FIELDS[number]>,
      formulas: Object.fromEntries(FORMULA_FIELDS.map(field => [field, formulas[field]])) as StockSourceRow['formulas'],
      audit_issues: auditIssues,
    };
  });
  return {
    schema_version: 1,
    source: { file_name: fileName, sha256, sheet_name: sheetName, header_row: headerRow as number,
      source_ref: sourceRef, extracted_at: extractedAt, ...(modifiedAt ? { source_modified_at: modifiedAt } : {}) },
    rows,
  };
}

function storeId(value: unknown): string | null {
  if (typeof value === 'number') return Number.isSafeInteger(value) && value > 0 ? String(value) : null;
  const result = normalizedText(value);
  return placeholder(result) || result.length > 200 || /[|\x00-\x1f\x7f]/.test(result) ? null : result;
}

function balanceValue(value: unknown): { value: number; negative: boolean; fractional: boolean; tooLarge: boolean } | null {
  if (typeof value === 'number') return Number.isFinite(value)
    ? { value, negative: value < 0, fractional: !Number.isInteger(value), tooLarge: value > MAX_QUANTITY } : null;
  if (typeof value !== 'string' || !/^-?(?:\d+(?:\.\d+)?|\.\d+)$/.test(value.trim())) return null;
  const raw = value.trim();
  const result = Number(raw);
  if (!Number.isFinite(result)) return null;
  const negative = raw.startsWith('-') && /[1-9]/.test(raw);
  const [integerPart, decimalPart = ''] = raw.replace(/^-/, '').split('.');
  const whole = integerPart.replace(/^0+/, '') || '0';
  const fractional = /[1-9]/.test(decimalPart);
  // Check text before Number() can round a tiny fraction or underflow to zero.
  const limit = String(MAX_QUANTITY);
  const tooLarge = !negative && (whole.length > limit.length || (whole.length === limit.length && (whole > limit || (whole === limit && fractional))));
  return { value: result, negative, fractional, tooLarge };
}

/** Read-only source screening, not an import or an Ecwid alignment approval. */
export function reviewStockSource(input: unknown): StockSourceReview {
  const snapshot = parseSnapshot(input);
  const skuOwners = new Map<string, number[]>();
  const storeOwners = new Map<string, number[]>();
  function owner(map: Map<string, number[]>, key: string | null, row: number) {
    if (key !== null) {
      const owners = map.get(key) ?? [];
      owners.push(row);
      map.set(key, owners);
    }
  }
  // A numeric SKU is never eligible, but it must also prevent a text SKU with
  // the same visible value from being treated as unambiguously unique.
  const skuKey = (value: unknown) => typeof value === 'number' && Number.isFinite(value) ? String(value) : normalizedText(value);
  function otherRows(owners: number[], current: number): string {
    const sample: number[] = [];
    for (const row of owners) {
      if (row !== current) sample.push(row);
      if (sample.length === 10) break;
    }
    const remaining = owners.length - 1 - sample.length;
    return `${sample.join(', ')}${remaining > 0 ? ` and ${remaining} more` : ''}`;
  }
  for (const row of snapshot.rows) {
    const sku = skuKey(row.sku);
    owner(skuOwners, placeholder(sku) ? null : sku, row.source_row);
    owner(storeOwners, storeId(row.store_id), row.source_row);
  }
  const candidateRows: StockCandidateRow[] = [];
  const rows = snapshot.rows.map((row): ReviewedStockRow => {
    const sku = normalizedText(row.sku);
    const unit = normalizedText(row.unit);
    const reasons: SourceIssue[] = [];
    let reviewRequired = false;
    let deferred = false;
    const issue = (code: string, message: string) => { reviewRequired = true; reasons.push({ code, message }); };
    const defer = (code: string, message: string) => { deferred = true; reasons.push({ code, message }); };
    if (normalizedText(row.status) !== 'ACTIVE') defer('STATUS_OUT_OF_SCOPE', 'Only explicitly Active rows are candidates; keep this item in the workbook.');
    if (normalizedText(row.item_type) !== 'SINGLE') defer('TYPE_OUT_OF_SCOPE', 'Only explicitly Single items are candidates; composite or unconfirmed types remain in the workbook.');
    if (unit !== 'PCS') defer('UNIT_OUT_OF_SCOPE', 'Only Pcs items are in the first phase; meter and other units remain in the workbook.');
    if (row.sku === null || (typeof row.sku === 'string' && placeholder(sku))) {
      defer('SKU_UNMAPPED', 'SKU is blank, NA or N/A. Keep in the workbook until mapped; Ecwid listing status has not been checked.');
    } else if (typeof row.sku !== 'string') {
      issue('SKU_NOT_TEXT', 'SKU must be stored as text so leading zeros and exact product identity can be verified.');
    } else if (sku.length > 200 || /[|\x00-\x1f\x7f]/.test(sku)) {
      issue('INVALID_SKU', 'SKU must be at most 200 characters, with no pipe or control characters.');
    }
    const sameSku = skuOwners.get(skuKey(row.sku)) ?? [];
    if (sameSku.length > 1) issue('DUPLICATE_SKU', `SKU also occurs on source rows ${otherRows(sameSku, row.source_row)}. Review every occurrence; balances are not combined.`);
    const id = storeId(row.store_id);
    if (id === null) issue('INVALID_STORE_ID', 'Store ID must be a nonblank identifier or positive whole number, with no placeholder, pipe or control characters.');
    const sameStore = id === null ? [] : storeOwners.get(id) ?? [];
    if (sameStore.length > 1) issue('DUPLICATE_STORE_ID', `Store ID also occurs on source rows ${otherRows(sameStore, row.source_row)}. Resolve the identity before migration.`);
    if (typeof row.name !== 'string' || !row.name.trim() || row.name.length > 1000 || /[\x00-\x1f\x7f]/.test(row.name)) issue('INVALID_NAME', 'An item name of 1–1,000 characters with no control characters is required.');
    const balance = balanceValue(row.balance);
    if (balance === null) issue('INVALID_BALANCE', 'Balance must be a finite numeric value, not blank, an error, or formatted text.');
    else if (balance.negative) issue('NEGATIVE_BALANCE', 'Negative physical stock must be reconciled before migration.');
    else if (balance.tooLarge) issue('BALANCE_TOO_LARGE', 'Balance exceeds the supported quantity limit of 2,147,483,647.');
    else if (unit === 'PCS' && balance.fractional) issue('FRACTIONAL_PCS_BALANCE', 'Pcs balance must be a whole number; do not round it during migration.');
    for (const audit of row.audit_issues) issue(audit.code, audit.message);
    const status = reviewRequired ? 'REVIEW' : deferred ? 'KEEP_WORKBOOK' : 'CANDIDATE';
    if (status === 'CANDIDATE') {
      reasons.push({ code: 'SOURCE_CANDIDATE_ONLY', message: 'Source checks passed for phase one. Ecwid mapping, outstanding orders and explicit cutover approval are still required; workbook remains authoritative.' });
      candidateRows.push({ sku, balance: balance!.value, name: (row.name as string).trim(),
        location: row.location === null ? '' : String(row.location).trim(), scan_code: sku,
        source_row: row.source_row, source_sheet: snapshot.source.sheet_name });
    }
    const { status: sourceStatus, sku: sourceSku, ...preserved } = row;
    return { ...preserved, source_sheet: snapshot.source.sheet_name, source_status: sourceStatus, source_sku: sourceSku, sku, status,
      current_authority: 'WORKBOOK', proposed_authority: status === 'CANDIDATE' ? 'APP_AFTER_APPROVAL' : 'WORKBOOK', reasons };
  });
  return {
    kind: 'SOURCE_REVIEW', dry_run: true, source: snapshot.source, generated_at: new Date().toISOString(),
    balance_meaning: 'PHYSICAL_ON_HAND', ecwid_checked: false, reservations_confirmed: false, ready_count: 0,
    counts: { total: rows.length, candidate: candidateRows.length, keep_workbook: rows.filter(row => row.status === 'KEEP_WORKBOOK').length,
      review: rows.filter(row => row.status === 'REVIEW').length },
    rows, candidate_rows: candidateRows,
  };
}
