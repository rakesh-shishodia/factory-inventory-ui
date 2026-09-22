import { DomainError, normalizeCode, PICKABLE_FULFILLMENT_STATUSES } from './domain';
import { canonicalVariationOptions, type EcwidOrder } from './ecwid';
import { previewImport } from './opening-import';
import { type OpeningStageScope } from './opening-apply';
import { buildWorkbookTargetStatements, validateWorkbookTargets, workbookTargetId } from './pilot-scope';
import { orderSnapshotHash } from './sync';

type Row = Record<string, unknown>;
export interface OpeningCutoverPolicy { storeId: string; actor: string; now?: string }
export interface OpeningLineConfirmation {
  order_id: string; ecwid_line_id: string; previously_picked_quantity: 0;
}
export interface OpeningOrdersSnapshot {
  kind: 'READONLY_ORDERS'; schema_version: 1; dry_run: true; complete: true; store_id: string;
  started_at: string; completed_at: string; creation_cutoff: number; orders_checked: number;
  pending_order_count: number; line_count: number; orders: EcwidOrder[];
}
export interface OpeningCutoverRequest {
  operation_id: string; expected_hash: string; confirm_staging: true; physical_counts_confirmed: true;
  freeze: { confirmed: true; started_at: string };
  input: Row; scope: OpeningStageScope[]; orders: OpeningOrdersSnapshot;
  line_confirmations: OpeningLineConfirmation[];
  workbook_scope: Array<{ sku: string; name: string; ecwid_product_id: string;
    ecwid_combination_id: string | null; ecwid_option_signature: string }>;
}
interface BatchRow {
  operation_id: string; fingerprint: string; store_id: string; review_hash: string;
  preview_hash: string; row_count: number; order_count: number; line_count: number;
  created_at: string; state: string;
  ecwid_change_verified?: number; ecwid_change_uncertain?: number;
}
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const SHA = /^[a-f0-9]{64}$/;
const ID = /^[1-9]\d{0,30}$/;
const MAX_QUANTITY = 2147483647;
const FRESH_MS = 15 * 60 * 1000;
function fail(code: string, message: string, status = 400): never { throw new DomainError(status, code, message); }
function object(value: unknown, name: string): Row {
  if (!value || typeof value !== 'object' || Array.isArray(value)) fail('INVALID_OPENING_CUTOVER', `${name} must be an object.`);
  return value as Row;
}
function array(value: unknown, name: string, maximum: number): Row[] {
  if (!Array.isArray(value) || value.length > maximum) fail('INVALID_OPENING_CUTOVER', `${name} must be a bounded array (maximum ${maximum}).`);
  return value.map(entry => object(entry, name));
}
function string(value: unknown, name: string, maximum = 200): string {
  if (typeof value !== 'string' || !value.trim() || value.length > maximum) fail('INVALID_OPENING_CUTOVER', `${name} must be nonblank and at most ${maximum} characters.`);
  return value;
}
function integer(value: unknown, name: string, minimum = 0, maximum = MAX_QUANTITY): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < minimum || value > maximum) fail('INVALID_OPENING_CUTOVER', `${name} is not a valid whole-number count.`);
  return value;
}
function utc(value: unknown, name: string): number {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value) || !Number.isFinite(Date.parse(value))) {
    fail('INVALID_OPENING_CUTOVER', `${name} must be a UTC timestamp.`);
  }
  return Date.parse(value);
}
async function hash(value: unknown): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(value)));
  return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
}
const identity = (row: OpeningStageScope) => JSON.stringify([row.sku, row.ecwid_product_id, row.ecwid_combination_id, row.ecwid_option_signature]);
const stockKey = (row: OpeningStageScope) => JSON.stringify([row.ecwid_product_id, row.ecwid_combination_id]);

function parseOrders(value: unknown, storeId: string): { snapshot: Row; orders: EcwidOrder[]; lineCount: number } {
  const snapshot = object(value, 'Orders snapshot');
  if (snapshot.kind !== 'READONLY_ORDERS' || snapshot.schema_version !== 1 || snapshot.dry_run !== true || snapshot.complete !== true || snapshot.store_id !== storeId) {
    fail('INVALID_OPENING_ORDERS', 'Provide a complete version-1 READONLY_ORDERS snapshot for the configured store.');
  }
  const started = utc(snapshot.started_at, 'Order snapshot started_at');
  const completed = utc(snapshot.completed_at, 'Order snapshot completed_at');
  if (completed < started || snapshot.creation_cutoff !== Math.floor(started / 1000)) fail('INVALID_OPENING_ORDERS', 'Order cutoff must match the beginning of the complete snapshot.');
  const values = array(snapshot.orders, 'Pending orders', 200);
  if (snapshot.pending_order_count !== values.length || integer(snapshot.orders_checked, 'orders_checked', 0, 50000) < values.length) {
    fail('INVALID_OPENING_ORDERS', 'Order snapshot counts do not agree with its contents.');
  }
  const seen = new Set<string>(); let lineCount = 0;
  const orders: EcwidOrder[] = values.map(raw => {
    const id = string(raw.id, 'Order ID');
    if (!ID.test(id) || seen.has(id)) fail('INVALID_OPENING_ORDERS', 'Order IDs must be unique positive identifiers.');
    seen.add(id);
    if (!['PAID', 'AWAITING_PAYMENT'].includes(String(raw.paymentStatus)) || !PICKABLE_FULFILLMENT_STATUSES.includes(String(raw.fulfillmentStatus))) {
      fail('OPENING_ORDER_STATUS_UNSUPPORTED', 'Every pending order must be Paid or Awaiting Payment and awaiting processing or processing. Review other statuses first.', 409);
    }
    if (utc(raw.updatedAt, 'Order updatedAt') > completed) fail('INVALID_OPENING_ORDERS', 'Order update cannot occur after snapshot completion.');
    if (raw.createdAt !== undefined && utc(raw.createdAt, 'Order createdAt') > started) fail('INVALID_OPENING_ORDERS', 'An order falls outside the creation cutoff.');
    const lines = array(raw.items, 'Order items', 500); const seenLines = new Set<string>();
    if (!lines.length || (lineCount += lines.length) > 2000) fail('INVALID_OPENING_ORDERS', 'Pending orders need nonempty lines, with at most 2,000 lines in total.');
    return { id, paymentStatus: String(raw.paymentStatus), fulfillmentStatus: String(raw.fulfillmentStatus), updatedAt: String(raw.updatedAt),
      ...(raw.createdAt !== undefined ? { createdAt: String(raw.createdAt) } : {}),
      items: lines.map(line => {
        const lineId = string(line.id, 'Order line ID');
        const productId = string(line.productId, 'Order line product ID');
        const sku = string(line.sku, 'Order line SKU');
        const name = string(line.name, 'Order line name', 500);
        const options = canonicalVariationOptions(line.selectedOptions);
        if (!ID.test(lineId) || !ID.test(productId) || seenLines.has(lineId) || sku !== normalizeCode(sku) || sku.includes('|')
          || (line.combinationId !== null && (typeof line.combinationId !== 'string' || !ID.test(line.combinationId)))
          || options === null || line.digital !== false || typeof line.trackQuantity !== 'boolean') {
          fail('INVALID_OPENING_ORDER_LINE', 'Each opening line needs an exact SKU/product/variation, supported option identity, unique line ID and explicit non-digital evidence.');
        }
        seenLines.add(lineId);
        return { id: lineId, productId, sku, name, quantity: integer(line.quantity, 'Order quantity', 1),
          combinationId: line.combinationId as string | null, selectedOptions: options, digital: false, trackQuantity: line.trackQuantity };
      }) };
  });
  if (snapshot.line_count !== lineCount) fail('INVALID_OPENING_ORDERS', 'Explicit snapshot line_count must match every pending order line.');
  return { snapshot, orders, lineCount };
}

async function prepare(value: unknown, policy: OpeningCutoverPolicy) {
  if (!/^[1-9]\d{0,19}$/.test(policy.storeId)) fail('OPENING_STORE_NOT_CONFIGURED', 'Configure the expected Ecwid store before staging.', 503);
  const actor = string(policy.actor, 'Authenticated administrator', 320).trim();
  const request = object(value, 'Opening cutover request');
  const allowed = new Set(['operation_id','expected_hash','confirm_staging','physical_counts_confirmed','freeze','input','scope','orders','line_confirmations','workbook_scope']);
  if (Object.keys(request).some(key => !allowed.has(key)) || typeof request.operation_id !== 'string' || !UUID.test(request.operation_id)
    || request.confirm_staging !== true || request.physical_counts_confirmed !== true) {
    fail('INVALID_OPENING_CUTOVER', 'Provide a UUID operation_id and explicit staging and physical-count confirmations.');
  }
  const freeze = object(request.freeze, 'Freeze confirmation');
  if (Object.keys(freeze).some(key => !['confirmed','started_at'].includes(key)) || freeze.confirmed !== true) fail('OPENING_FREEZE_REQUIRED', 'Confirm that pilot sales and stock movements are paused.');
  utc(freeze.started_at, 'Freeze started_at');
  const input = object(request.input, 'Original preview input');
  if (input.store_id !== policy.storeId) fail('OPENING_STORE_MISMATCH', 'Opening input must match the configured store.');
  const catalog = object(input.catalog, 'Complete catalogue snapshot');
  if (catalog.store_id !== policy.storeId) fail('OPENING_STORE_MISMATCH', 'Catalogue must match the configured store.');
  if (typeof input.snapshot_source_hash !== 'string' || !SHA.test(input.snapshot_source_hash)) fail('OPENING_SOURCE_HASH_REQUIRED', 'Provide the SHA-256 of the authoritative workbook.');
  if (typeof input.source_ref !== 'string' || !input.source_ref.trim() || input.source_ref.length > 1000) fail('INVALID_OPENING_CUTOVER', 'Provide a bounded workbook source reference.');
  const preview = await previewImport(input);
  if (preview.global_errors.length || preview.blocked_count || !preview.rows.length) fail('OPENING_PREVIEW_BLOCKED', 'All selected stock-limited items must pass the original import preview.', 409);
  if (preview.rows.length > 200 || preview.rows.some(row => row.name.length > 500 || row.location.length > 500)) fail('OPENING_CUTOVER_TOO_LARGE', 'Stage at most 200 bounded pilot stock targets.', 413);
  const scope = array(request.scope, 'Approved scope', 200).map(row => {
    if (Object.keys(row).sort().join(',') !== 'ecwid_combination_id,ecwid_option_signature,ecwid_product_id,sku'
      || typeof row.sku !== 'string' || typeof row.ecwid_product_id !== 'string' || typeof row.ecwid_option_signature !== 'string'
      || (row.ecwid_combination_id !== null && typeof row.ecwid_combination_id !== 'string')) fail('OPENING_SCOPE_MISMATCH', 'Scope needs exact stock-target identities.');
    return { sku: row.sku, ecwid_product_id: row.ecwid_product_id, ecwid_combination_id: row.ecwid_combination_id as string | null, ecwid_option_signature: row.ecwid_option_signature };
  });
  const expectedScope = preview.rows.map(row => ({ sku: row.sku, ecwid_product_id: row.ecwid_product_id!, ecwid_combination_id: row.ecwid_combination_id, ecwid_option_signature: row.ecwid_option_signature! }));
  if (JSON.stringify(scope.map(identity).sort()) !== JSON.stringify(expectedScope.map(identity).sort())) fail('OPENING_SCOPE_MISMATCH', 'Every approved target must match the recomputed preview exactly once.');
  const workbookScope = validateWorkbookTargets(request.workbook_scope);
  if (workbookScope.length > 2000) fail('OPENING_CUTOVER_TOO_LARGE', 'The workbook-managed registry is too large.', 413);
  if (workbookScope.some(row => scope.some(pilot => stockKey(pilot) === stockKey(row) || pilot.sku === row.sku))) {
    fail('OPENING_SCOPE_CONFLICT', 'A target cannot be both workbook-managed and app-managed.');
  }
  const catalogueTargets = array(catalog.stock_targets, 'Catalogue stock targets', 20000);
  for (const target of workbookScope) {
    const exact = catalogueTargets.filter(row => String(row.id) === target.ecwid_product_id
      && row.combinationId === target.ecwid_combination_id);
    const skuMatches = catalogueTargets.filter(row => typeof row.sku === 'string' && normalizeCode(row.sku) === target.sku
      && !(row.combinationId === null && row.hasVariations === true));
    if (exact.length !== 1 || skuMatches.length !== 1 || exact[0] !== skuMatches[0]
      || typeof exact[0].sku !== 'string' || normalizeCode(exact[0].sku) !== target.sku
      || JSON.stringify(canonicalVariationOptions(exact[0].variationOptions)) !== target.ecwid_option_signature) {
      fail('OPENING_WORKBOOK_CATALOGUE_MISMATCH', 'Every workbook-managed identity must match one unambiguous target in the complete catalogue. Units and stock policy remain outside the app.', 409);
    }
  }
  const { snapshot, orders, lineCount } = parseOrders(request.orders, policy.storeId);
  const confirmations = array(request.line_confirmations, 'Line confirmations', 2000).map(raw => {
    if (Object.keys(raw).sort().join(',') !== 'ecwid_line_id,order_id,previously_picked_quantity'
      || typeof raw.order_id !== 'string' || typeof raw.ecwid_line_id !== 'string' || raw.previously_picked_quantity !== 0) {
      fail('OPENING_PRIOR_PICKS_UNSUPPORTED', 'Explicitly confirm zero previous physical picks for each exact opening order line; historical picks are not supported.', 409);
    }
    return { order_id: raw.order_id, ecwid_line_id: raw.ecwid_line_id, previously_picked_quantity: 0 as const };
  });
  const lineKeys = orders.flatMap(order => order.items.map(line => `${order.id}:${line.id}`)).sort();
  const confirmationKeys = confirmations.map(line => `${line.order_id}:${line.ecwid_line_id}`).sort();
  if (JSON.stringify(lineKeys) !== JSON.stringify(confirmationKeys)) fail('OPENING_LINE_CONFIRMATIONS_MISMATCH', 'Confirm each complete opening order line exactly once.', 409);
  const pilotByIdentity = new Map(scope.map(row => [identity(row), row]));
  const workbookByIdentity = new Map(workbookScope.map(row => [identity(row), row]));
  const reservations = new Map(scope.map(row => [row.sku, 0]));
  const lines = orders.flatMap(order => order.items.map(line => {
    const target = { sku: line.sku, ecwid_product_id: line.productId, ecwid_combination_id: line.combinationId,
      ecwid_option_signature: JSON.stringify(canonicalVariationOptions(line.selectedOptions)) };
    const pilot = pilotByIdentity.get(identity(target)); const workbook = workbookByIdentity.get(identity(target));
    if (!pilot && !workbook) fail('OPENING_ORDER_LINE_UNMAPPED', 'An opening order line is not an exact approved pilot target or explicitly workbook-managed target. Reconcile all lines first.', 409);
    if (pilot) {
      const quantity = reservations.get(pilot.sku)! + line.quantity;
      if (quantity > MAX_QUANTITY) fail('INVALID_OPENING_ORDER_LINE', 'Combined order commitments exceed the supported range.');
      reservations.set(pilot.sku, quantity);
    }
    return { id: `${order.id}:${line.id}`, order_id: order.id, ecwid_line_id: line.id, sku: line.sku, name: line.name,
      ordered_qty: line.quantity, management_mode: pilot ? 'APP' : 'WORKBOOK', workbook_target_id: workbook ? workbookTargetId(workbook) : null };
  }));
  const summary = array(input.reservations, 'Reservation summaries', 200).map(row => ({ sku: row.sku, quantity: row.quantity }));
  if (summary.some(row => typeof row.sku !== 'string' || !reservations.has(row.sku) || row.quantity !== reservations.get(row.sku))
    || scope.some(target => (reservations.get(target.sku)! > 0 || summary.some(row => row.sku === target.sku))
      && summary.filter(row => row.sku === target.sku).length !== 1)
    || preview.rows.some(row => row.reserved !== reservations.get(row.sku))) {
    fail('OPENING_RESERVATIONS_MISMATCH', 'Reservation summaries must equal exact unpicked pilot order-line quantities, with no omitted or unrelated commitments.', 409);
  }
  const reviewHash = await hash([1, policy.storeId, preview.source_hash, scope.map(identity).sort(), snapshot, confirmations, workbookScope, freeze, true]);
  return { request, operationId: request.operation_id.toLowerCase(), actor, freeze, input, catalog, preview, scope,
    workbookScope, snapshot, orders, lines, lineCount, confirmations, reviewHash };
}

function assertFresh(prepared: Awaited<ReturnType<typeof prepare>>, policy: OpeningCutoverPolicy): void {
  const now = utc(policy.now ?? new Date().toISOString(), 'Current time');
  const frozen = utc(prepared.freeze.started_at, 'Freeze started_at');
  if (frozen > now || now - frozen > FRESH_MS) fail('OPENING_SNAPSHOT_STALE', 'The freeze confirmation expired. Obtain fresh stock and order snapshots.', 409);
  for (const snapshot of [prepared.catalog, prepared.snapshot]) {
    const started = utc(snapshot.started_at, 'Snapshot started_at');
    const completed = utc(snapshot.completed_at, 'Snapshot completed_at');
    if (started < frozen || completed < started || completed > now || now - started > FRESH_MS) {
      fail('OPENING_SNAPSHOT_STALE', 'Both complete snapshots must have been collected during the current confirmed stock freeze, within 15 minutes.', 409);
    }
  }
}

/** Pure, source-recomputed review. No database, network or remote writes. */
export async function previewOpeningCutover(value: unknown, policy: OpeningCutoverPolicy) {
  const prepared = await prepare(value, policy); assertFresh(prepared, policy);
  return { dry_run: true, review_hash: prepared.reviewHash, preview_hash: prepared.preview.source_hash,
    store_id: policy.storeId, row_count: prepared.preview.rows.length, order_count: prepared.orders.length,
    line_count: prepared.lineCount, workbook_line_count: prepared.lines.filter(line => line.management_mode === 'WORKBOOK').length,
    rows: prepared.preview.rows, reservations_loaded: false, activated: false, ecwid_changed: false };
}

function receipt(row: BatchRow, duplicate: boolean) {
  return { status: row.state, duplicate, operation_id: row.operation_id, store_id: row.store_id,
    review_hash: row.review_hash, preview_hash: row.preview_hash, row_count: row.row_count,
    order_count: row.order_count, line_count: row.line_count, created_at: row.created_at,
    reservations_loaded: true, activated: row.state === 'ACTIVE',
    ecwid_changed: row.ecwid_change_verified ? true : row.ecwid_change_uncertain ? null : false };
}

/** Administrator + disabled-live flags are enforced by the caller. This function never activates or calls Ecwid. */
export async function stageOpeningCutover(db: D1Database, value: unknown, policy: OpeningCutoverPolicy) {
  const p = await prepare(value, policy);
  if (typeof p.request.expected_hash !== 'string' || !SHA.test(p.request.expected_hash) || p.request.expected_hash !== p.reviewHash) {
    fail('OPENING_REVIEW_CHANGED', 'Approve the recomputed cutover review hash before staging.', 409);
  }
  const fingerprint = await hash([1, p.reviewHash, p.actor]);
  const find = () => db.prepare(`SELECT operation_id,fingerprint,store_id,review_hash,preview_hash,row_count,order_count,line_count,created_at,state,
    EXISTS(SELECT 1 FROM opening_cutover_rows r WHERE r.operation_id=b.operation_id
      AND r.alignment_status='VERIFIED' AND r.before_quantity<>r.after_quantity) AS ecwid_change_verified,
    EXISTS(SELECT 1 FROM opening_cutover_rows r WHERE r.operation_id=b.operation_id
      AND r.alignment_status IN ('PROCESSING','UNKNOWN')) AS ecwid_change_uncertain
    FROM opening_cutover_batches b WHERE operation_id=?`).bind(p.operationId).first<BatchRow>();
  const duplicate = (existing: BatchRow) => {
    if (existing.fingerprint !== fingerprint) fail('OPENING_OPERATION_REUSED', 'This operation ID belongs to different content or administrator.', 409);
    return receipt(existing, true);
  };
  const existing = await find(); if (existing) return duplicate(existing);
  assertFresh(p, policy);
  const timestamp = policy.now ?? new Date().toISOString();
  const rows = p.preview.rows.map(row => ({ ...row, item_id: crypto.randomUUID(), opening_id: crypto.randomUUID(), issue_id: crypto.randomUUID() }));
  const itemsBySku = new Map(rows.map(row => [row.sku, row.item_id]));
  const linesJson = JSON.stringify(p.lines.map(line => ({ ...line, item_id: itemsBySku.get(line.sku) ?? null })));
  const rowsJson = JSON.stringify(rows);
  const ordersJson = JSON.stringify(await Promise.all(p.orders.map(async order => ({ id: order.id,
    payment_status: order.paymentStatus, fulfillment_status: order.fulfillmentStatus, remote_updated_at: order.updatedAt,
    remote_lines_hash: await orderSnapshotHash(order), line_count: order.items.length }))));
  if ([rowsJson, linesJson, ordersJson].some(text => new TextEncoder().encode(text).byteLength > 1_500_000)) {
    fail('OPENING_CUTOVER_TOO_LARGE', 'Opening payload exceeds the bounded transaction size.', 413);
  }
  const batch: BatchRow = { operation_id: p.operationId, fingerprint, store_id: policy.storeId, review_hash: p.reviewHash,
    preview_hash: p.preview.source_hash, row_count: rows.length, order_count: p.orders.length, line_count: p.lineCount, created_at: timestamp, state: 'STAGED' };
  try {
    await db.batch([
      db.prepare(`INSERT INTO opening_cutover_batches(operation_id,fingerprint,store_id,review_hash,preview_hash,source_hash,orders_hash,catalogue_hash,
        source_ref,scope_json,orders_json,confirmations_json,workbook_scope_json,row_count,order_count,line_count,actor,frozen_at,created_at,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).bind(p.operationId, fingerprint, policy.storeId, p.reviewHash, p.preview.source_hash,
        p.input.snapshot_source_hash, await hash(p.snapshot), await hash(p.catalog), p.input.source_ref, JSON.stringify(p.scope), JSON.stringify(p.snapshot),
        JSON.stringify(p.confirmations), JSON.stringify(p.workbookScope), rows.length, p.orders.length, p.lineCount, p.actor, p.freeze.started_at, timestamp, timestamp),
      ...buildWorkbookTargetStatements(db, p.workbookScope, { reference: String(p.input.source_ref), actor: p.actor, timestamp }),
      db.prepare(`INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,on_hand,last_ecwid_quantity,active,inventory_mode)
        SELECT value->>'item_id',value->>'sku',value->>'name',value->>'scan_code',value->>'location',value->>'ecwid_product_id',
        value->>'ecwid_combination_id',value->>'ecwid_option_signature',0,value->>'ecwid_quantity',0,'STOCK_LIMITED' FROM json_each(?)`).bind(rowsJson),
      db.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
        SELECT value->>'opening_id',value->>'item_id',value->>'physical_on_hand',?,?,? FROM json_each(?)`).bind(p.input.source_ref,p.actor,timestamp,rowsJson),
      db.prepare(`INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,remote_lines_hash,updated_at,needs_review)
        SELECT value->>'id',value->>'payment_status',value->>'fulfillment_status',value->>'remote_updated_at',value->>'remote_lines_hash',?,0 FROM json_each(?)`).bind(timestamp,ordersJson),
      db.prepare(`INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty,picked_qty,management_mode,workbook_target_id)
        SELECT value->>'id',value->>'order_id',value->>'ecwid_line_id',value->>'item_id',value->>'sku',value->>'name',value->>'ordered_qty',0,
          value->>'management_mode',value->>'workbook_target_id' FROM json_each(?)`).bind(linesJson),
      db.prepare(`INSERT INTO opening_cutover_orders(operation_id,order_id,remote_lines_hash,line_count)
        SELECT ?,value->>'id',value->>'remote_lines_hash',value->>'line_count' FROM json_each(?)`).bind(p.operationId,ordersJson),
      db.prepare(`INSERT INTO opening_cutover_rows(operation_id,item_id,opening_balance_id,sku,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,
        physical,unpicked,target_quantity,expected_ecwid_quantity,source_row,source_sheet)
        SELECT ?,value->>'item_id',value->>'opening_id',value->>'sku',value->>'ecwid_product_id',value->>'ecwid_combination_id',value->>'ecwid_option_signature',
          value->>'physical_on_hand',value->>'reserved',value->>'desired_ecwid_quantity',value->>'ecwid_quantity',value->>'source_row',value->>'source_sheet' FROM json_each(?)`).bind(p.operationId,rowsJson),
      db.prepare(`INSERT INTO sync_issues(id,item_id,kind,message,status,created_at)
        SELECT value->>'issue_id',value->>'item_id','OPENING_CUTOVER_STAGED','Opening balances and order commitments are staged. Approved verified Ecwid alignment and activation are still required.','OPEN',? FROM json_each(?)`).bind(timestamp,rowsJson),
      // The update trigger verifies complete audit counts after all inserts, in the same transaction.
      db.prepare('UPDATE opening_cutover_batches SET updated_at=? WHERE operation_id=?').bind(timestamp,p.operationId),
    ]);
  } catch (error) {
    const committed = await find(); if (committed) return duplicate(committed);
    if (error instanceof Error && error.message.includes('OPENING_DATABASE_STORE_MISMATCH')) fail('OPENING_DATABASE_STORE_MISMATCH', 'Opening stock belongs to another store.',409);
    if (error instanceof Error && /UNIQUE constraint failed|AMBIGUOUS_ITEM_CODE|OPENING_BALANCE_ALREADY_STARTED|WORKBOOK_TARGET_CONFLICT/.test(error.message)) {
      fail('OPENING_CUTOVER_CONFLICT', 'An item, order, target or workbook registry entry conflicts with existing data. The entire staging transaction was rolled back.',409);
    }
    throw error;
  }
  return receipt(batch, false);
}
