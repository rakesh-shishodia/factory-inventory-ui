import { DomainError } from './domain';
import { previewImport } from './opening-import';

type Row = Record<string, unknown>;
export interface OpeningStageScope {
  sku: string;
  ecwid_product_id: string;
  ecwid_combination_id: string | null;
  ecwid_option_signature: string;
}
export interface OpeningStagePolicy { storeId: string; actor: string }
export interface OpeningStageResult {
  status: 'STAGED'; duplicate: boolean; operation_id: string; store_id: string;
  preview_hash: string; row_count: number; created_at: string;
  activated: false; reservations_loaded: false; ecwid_changed: false;
}
interface BatchRow {
  operation_id: string; fingerprint: string; store_id: string; preview_hash: string;
  row_count: number; created_at: string;
}

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const SHA256 = /^[a-f0-9]{64}$/;
const MAX_STAGE_ROWS = 200;
function object(value: unknown, name: string): Row {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new DomainError(400, 'INVALID_OPENING_STAGE', `${name} must be an object.`);
  }
  return value as Row;
}
function identity(row: OpeningStageScope): string {
  return JSON.stringify([row.sku, row.ecwid_product_id, row.ecwid_combination_id, row.ecwid_option_signature]);
}
async function hash(value: unknown): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(value)));
  return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
}
function result(row: BatchRow, duplicate: boolean): OpeningStageResult {
  return {
    status: 'STAGED', duplicate, operation_id: row.operation_id, store_id: row.store_id,
    preview_hash: row.preview_hash, row_count: row.row_count, created_at: row.created_at,
    activated: false, reservations_loaded: false, ecwid_changed: false
  };
}

/**
 * DB-only staging, not a cutover. The caller must authorize an administrator,
 * pin the configured store, and keep live inventory disabled. Never accepts a
 * caller's preview output as evidence: the original input is revalidated here.
 * Nonzero reservations need a future exact order-line opening importer.
 */
export async function stageOpeningImport(
  db: D1Database, value: unknown, policy: OpeningStagePolicy
): Promise<OpeningStageResult> {
  if (typeof policy.storeId !== 'string' || !/^[1-9]\d{0,19}$/.test(policy.storeId)) {
    throw new DomainError(503, 'OPENING_STORE_NOT_CONFIGURED', 'Configure the expected Ecwid store before staging opening stock.');
  }
  if (typeof policy.actor !== 'string' || !policy.actor.trim() || policy.actor.length > 320) {
    throw new DomainError(400, 'INVALID_OPENING_ACTOR', 'An authenticated administrator is required.');
  }
  const request = object(value, 'Opening staging request');
  const allowed = new Set(['operation_id', 'expected_hash', 'confirm_staging', 'input', 'scope']);
  if (Object.keys(request).some(key => !allowed.has(key)) || request.confirm_staging !== true ||
      typeof request.operation_id !== 'string' || !UUID.test(request.operation_id) ||
      typeof request.expected_hash !== 'string' || !SHA256.test(request.expected_hash)) {
    throw new DomainError(400, 'INVALID_OPENING_STAGE', 'Provide a UUID operation_id, reviewed expected_hash and explicit confirm_staging: true.');
  }
  const input = object(request.input, 'Original preview input');
  if (input.store_id !== policy.storeId) {
    throw new DomainError(400, 'OPENING_STORE_MISMATCH', 'The preview input must identify the configured store.');
  }
  const catalog = object(input.catalog, 'Complete read-only catalogue snapshot');
  if (catalog.store_id !== policy.storeId) {
    throw new DomainError(400, 'OPENING_STORE_MISMATCH', 'The catalogue must identify the configured store.');
  }
  if (typeof input.snapshot_source_hash !== 'string' || !SHA256.test(input.snapshot_source_hash)) {
    throw new DomainError(400, 'OPENING_SOURCE_HASH_REQUIRED', 'Provide the SHA-256 digest of the authoritative stock workbook.');
  }
  const preview = await previewImport(input);
  if (preview.source_hash !== request.expected_hash) {
    throw new DomainError(409, 'OPENING_PREVIEW_CHANGED', 'The opening-stock input changed after review. Generate and approve a new preview.');
  }
  if (preview.global_errors.length || preview.blocked_count || !preview.rows.length) {
    throw new DomainError(409, 'OPENING_PREVIEW_BLOCKED', 'Every selected row and all preview-wide checks must be READY before staging.');
  }
  if (preview.rows.length > MAX_STAGE_ROWS) {
    throw new DomainError(413, 'OPENING_STAGE_TOO_LARGE', `Stage at most ${MAX_STAGE_ROWS} explicitly reviewed stock targets per batch.`);
  }
  // Also reject reservations for omitted SKUs: a copied aggregate must never
  // quietly lose order commitments simply because a row was not selected.
  if (preview.rows.some(row => row.reserved !== 0) ||
      (Array.isArray(input.reservations) && input.reservations.some(row => Number(object(row, 'Reservation').quantity) !== 0))) {
    throw new DomainError(409, 'OPENING_RESERVATIONS_UNSUPPORTED', 'This staging engine cannot import outstanding orders. Nonzero reservations require exact reviewed order-line import first.');
  }
  if (typeof input.source_ref !== 'string' || input.source_ref.length > 1000 ||
      preview.rows.some(row => row.name.length > 500 || row.location.length > 500)) {
    throw new DomainError(400, 'INVALID_OPENING_STAGE', 'Source references, item names or locations exceed the staging limits.');
  }
  if (!Array.isArray(request.scope) || request.scope.length !== preview.rows.length) {
    throw new DomainError(400, 'OPENING_SCOPE_MISMATCH', 'Explicit scope must identify every selected stock target exactly once.');
  }
  const scope = request.scope.map(value => {
    const row = object(value, 'Scope target');
    if (Object.keys(row).sort().join(',') !== 'ecwid_combination_id,ecwid_option_signature,ecwid_product_id,sku' ||
        typeof row.sku !== 'string' || typeof row.ecwid_product_id !== 'string' ||
        (row.ecwid_combination_id !== null && typeof row.ecwid_combination_id !== 'string') ||
        typeof row.ecwid_option_signature !== 'string') {
      throw new DomainError(400, 'OPENING_SCOPE_MISMATCH', 'Each scope target needs its exact SKU, product, combination and canonical options.');
    }
    return identity({ sku: row.sku, ecwid_product_id: row.ecwid_product_id,
      ecwid_combination_id: row.ecwid_combination_id, ecwid_option_signature: row.ecwid_option_signature });
  }).sort();
  const expectedScope = preview.rows.map(row => identity({
    sku: row.sku, ecwid_product_id: row.ecwid_product_id!, ecwid_combination_id: row.ecwid_combination_id,
    ecwid_option_signature: row.ecwid_option_signature!
  })).sort();
  if (new Set(scope).size !== scope.length || JSON.stringify(scope) !== JSON.stringify(expectedScope)) {
    throw new DomainError(400, 'OPENING_SCOPE_MISMATCH', 'The approved stock-target identities do not match the recomputed preview.');
  }
  const operationId = request.operation_id.toLowerCase();
  const actor = policy.actor.trim();
  const fingerprint = await hash([1, policy.storeId, preview.source_hash, scope, actor]);
  const existingBatch = () => db.prepare(`SELECT operation_id,fingerprint,store_id,preview_hash,row_count,created_at
    FROM opening_import_batches WHERE operation_id=?`).bind(operationId).first<BatchRow>();
  const duplicateResult = (existing: BatchRow) => {
    if (existing.fingerprint !== fingerprint) {
      throw new DomainError(409, 'OPENING_OPERATION_REUSED', 'This operation_id already belongs to different opening-stock content or administrator.');
    }
    return result(existing, true);
  };
  const existing = await existingBatch();
  if (existing) return duplicateResult(existing);
  const createdAt = new Date().toISOString();
  const rowsJson = JSON.stringify(preview.rows.map(row => ({
    item_id: crypto.randomUUID(), opening_id: crypto.randomUUID(), issue_id: crypto.randomUUID(),
    sku: row.sku, name: row.name, scan_code: row.scan_code, location: row.location,
    ecwid_product_id: row.ecwid_product_id, ecwid_combination_id: row.ecwid_combination_id,
    ecwid_option_signature: row.ecwid_option_signature, on_hand: row.physical_on_hand,
    ecwid_quantity: row.ecwid_quantity, source_row: row.source_row ?? null, source_sheet: row.source_sheet ?? null
  })));
  if (new TextEncoder().encode(rowsJson).byteLength > 512_000) {
    throw new DomainError(413, 'OPENING_STAGE_TOO_LARGE', 'The staged stock-target payload is too large. Select fewer rows.');
  }
  const catalogueHash = await hash(catalog);
  try {
    // A fixed-size batch avoids a query per row. D1 rolls back every statement
    // when any uniqueness, audit or ledger guard fails, including a late failure.
    await db.batch([
      db.prepare(`INSERT INTO opening_import_batches
        (operation_id,fingerprint,store_id,preview_hash,source_hash,catalogue_hash,source_ref,scope_json,row_count,actor,created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?)`).bind(operationId, fingerprint, policy.storeId, preview.source_hash,
        input.snapshot_source_hash, catalogueHash, input.source_ref.trim(), JSON.stringify(scope), preview.rows.length, actor, createdAt),
      db.prepare(`INSERT INTO items
        (id,sku,name,scan_code,location,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,on_hand,last_ecwid_quantity,active)
        SELECT value->>'item_id',value->>'sku',value->>'name',value->>'scan_code',value->>'location',
          value->>'ecwid_product_id',value->>'ecwid_combination_id',value->>'ecwid_option_signature',0,value->>'ecwid_quantity',0
        FROM json_each(?)`).bind(rowsJson),
      db.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
        SELECT value->>'opening_id',value->>'item_id',value->>'on_hand',?,?,? FROM json_each(?)`)
        .bind(input.source_ref.trim(), actor, createdAt, rowsJson),
      db.prepare(`INSERT INTO opening_import_rows
        (operation_id,item_id,opening_balance_id,sku,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,
          on_hand,ecwid_quantity,source_row,source_sheet)
        SELECT ?,value->>'item_id',value->>'opening_id',value->>'sku',value->>'ecwid_product_id',value->>'ecwid_combination_id',
          value->>'ecwid_option_signature',value->>'on_hand',value->>'ecwid_quantity',value->>'source_row',value->>'source_sheet'
        FROM json_each(?)`).bind(operationId, rowsJson),
      db.prepare(`INSERT INTO sync_issues(id,item_id,kind,message,status,created_at)
        SELECT value->>'issue_id',value->>'item_id','OPENING_IMPORT_STAGED',?,'OPEN',? FROM json_each(?)`)
        .bind('Opening stock is staged and inactive. Fresh stock/order review and approved Ecwid alignment are required before activation.', createdAt, rowsJson)
    ]);
  } catch (error) {
    // Concurrent retries may pass the initial read together. The losing atomic
    // insert must re-read the committed winner and compare its content, not apply
    // more stock or silently accept a different operation with the same key.
    const committed = await existingBatch();
    if (committed) return duplicateResult(committed);
    if (error instanceof Error && error.message.includes('OPENING_DATABASE_STORE_MISMATCH')) {
      throw new DomainError(409, 'OPENING_DATABASE_STORE_MISMATCH', 'This database already contains opening stock for a different Ecwid store. Use a separate database.');
    }
    if (error instanceof Error && /UNIQUE constraint failed|AMBIGUOUS_ITEM_CODE|OPENING_BALANCE_ALREADY_STARTED/.test(error.message)) {
      throw new DomainError(409, 'OPENING_ITEM_CONFLICT', 'A selected SKU, scan code or Ecwid target already exists. No opening stock was staged.');
    }
    throw error;
  }
  return result({ operation_id: operationId, fingerprint, store_id: policy.storeId,
    preview_hash: preview.source_hash, row_count: preview.rows.length, created_at: createdAt }, false);
}
