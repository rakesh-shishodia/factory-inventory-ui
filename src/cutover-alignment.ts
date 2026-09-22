import { DomainError, TERMINAL_FULFILLMENT_STATUSES } from './domain';
import { canonicalVariationOptions, EcwidClient, EcwidError, type EcwidFetch, type EcwidOrder, type EcwidProduct } from './ecwid';
import type { OpeningOrdersSnapshot } from './opening-cutover';
import { validateWorkbookTargets, workbookTargetId, type WorkbookTarget } from './pilot-scope';
import { orderSnapshotHash } from './sync';
import { workbookLineMatches } from './workbook-identity';

export interface CutoverAlignmentRequest {
  operation_id: string;
  expected_hash: string;
  freeze: { confirmed: true; started_at: string };
}
export interface CutoverRecoveryRequest {
  operation_id: string;
  expected_hash: string;
  recovery_id: string;
  recovery_freeze: { confirmed: true; started_at: string };
}
export interface CutoverAlignmentPolicy {
  storeId: string; actor: string; token: string; mode: string;
  inventoryEnabled: string; liveSyncEnabled: string; orderSyncEnabled: string;
  /** Trusted test/orchestrator clock only; never accept this from an HTTP body. */
  now?: string;
}
export interface CutoverAlignmentOptions { fetcher?: EcwidFetch; clock?: () => string }
export interface CutoverRecoveryOptions extends CutoverAlignmentOptions {
  /** Trusted operator reporters; never populated from a request body. */
  onPreflight?: (receipt: Record<string, unknown>) => void;
  onRow?: (receipt: Record<string, unknown>) => void;
  beforeStockWrite?: () => Promise<void>;
  beforeActivation?: () => Promise<void>;
}
type BatchState = 'STAGED' | 'ALIGNING' | 'ALIGNED' | 'ACTIVE' | 'REVIEW';
type AlignmentStatus = 'PENDING' | 'PROCESSING' | 'VERIFIED' | 'UNKNOWN' | 'BLOCKED';
interface Batch {
  operation_id: string; store_id: string; review_hash: string; orders_hash: string;
  actor: string; frozen_at: string; created_at: string; state: BatchState;
  orders_json: string; workbook_scope_json: string; row_count: number; order_count: number; line_count: number;
}
interface StockRow {
  operation_id: string; item_id: string; sku: string; ecwid_product_id: string;
  ecwid_combination_id: string | null; ecwid_option_signature: string;
  physical: number; unpicked: number; target_quantity: number; expected_ecwid_quantity: number;
  alignment_status: AlignmentStatus; before_quantity: number | null; after_quantity: number | null;
}
interface OrderIdentity { id: string; payment: string; fulfillment: string; updated: string; hash: string; lines: number }
interface LineIdentity {
  id: string; order_id: string; ecwid_line_id: string; sku: string; quantity: number;
  item_id: string | null; mode: 'APP' | 'WORKBOOK'; workbook_target_id: string | null;
}
interface Context {
  batch: Batch; rows: StockRow[]; snapshot: OpeningOrdersSnapshot;
  orders: OrderIdentity[]; lines: LineIdentity[]; now: () => string;
  policy: CutoverAlignmentPolicy; fetcher?: EcwidFetch;
  workbook: WorkbookTarget[];
  leaseStartedAt: string; leaseMs: number;
}
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const SHA = /^[a-f0-9]{64}$/;
const FREEZE_MS = 15 * 60 * 1000;
const RECOVERY_FREEZE_MS = 30 * 60 * 1000;
const RECOVERY_WRITE_BUFFER_MS = 60 * 1000;
const RECOVERY_FINALIZATION_BUFFER_MS = 5 * 60 * 1000;
const MAX_ORDERS = 50_000;
function fail(code: string, message: string, status = 409): never { throw new DomainError(status, code, message); }
function utc(value: unknown): number {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value)
    || !Number.isFinite(Date.parse(value))) fail('CUTOVER_INVALID_TIME', 'Use a valid UTC freeze timestamp.', 400);
  return Date.parse(value);
}
async function sha(value: unknown): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(value)));
  return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
}
function remainingLease(context: Context): number {
  const elapsed = utc(context.now()) - utc(context.leaseStartedAt);
  if (elapsed < 0 || elapsed >= context.leaseMs) fail('CUTOVER_FREEZE_EXPIRED', 'The confirmed stock freeze expired. Stop and reconcile before any further alignment.');
  return context.leaseMs - elapsed;
}
function client(context: Context): EcwidClient {
  return new EcwidClient({ storeId: context.policy.storeId, token: context.policy.token }, context.fetcher,
    Math.max(1, Math.min(15_000, remainingLease(context))));
}
async function orderIdentity(order: EcwidOrder,workbookTargets:WorkbookTarget[]=[],allowSnapshotEvidence=false): Promise<OrderIdentity> {
  return { id: order.id, payment: order.paymentStatus, fulfillment: order.fulfillmentStatus,
    updated: order.updatedAt, hash: await orderSnapshotHash(order,{allowSnapshotEvidence,workbookTargets}), lines: order.items.length };
}
function stockKey(sku: string, product: string, combination: string | null, signature: string): string {
  return JSON.stringify([sku, product, combination, signature]);
}

async function loadContext(db: D1Database, value: unknown, policy: CutoverAlignmentPolicy,
  options: CutoverAlignmentOptions, rowRequest = false): Promise<Context> {
  if (policy.mode !== 'live' || policy.inventoryEnabled !== 'false' || policy.liveSyncEnabled !== 'false'
    || policy.orderSyncEnabled !== 'false') fail('CUTOVER_FLAGS_UNSAFE', 'Live mode and disabled inventory, order polling and outbound sync are required.');
  if (typeof policy.storeId !== 'string' || !/^[1-9]\d{0,19}$/.test(policy.storeId)
    || typeof policy.token !== 'string' || !policy.token.trim()
    || typeof policy.actor !== 'string' || !policy.actor.trim() || policy.actor.length > 320) {
    fail('CUTOVER_POLICY_INVALID', 'Configure the exact store, secret credential and authenticated administrator.', 400);
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) fail('CUTOVER_REQUEST_INVALID', 'A reviewed cutover request is required.', 400);
  const raw = value as Record<string, unknown>;
  const allowed = ['operation_id', 'expected_hash', 'freeze', ...(rowRequest ? ['item_id'] : [])];
  if (Object.keys(raw).some(key => !allowed.includes(key)) || typeof raw.operation_id !== 'string' || !UUID.test(raw.operation_id)
    || typeof raw.expected_hash !== 'string' || !SHA.test(raw.expected_hash)
    || !raw.freeze || typeof raw.freeze !== 'object' || Array.isArray(raw.freeze)) fail('CUTOVER_REQUEST_INVALID', 'Provide the exact operation ID, approved hash and freeze confirmation.', 400);
  const freeze = raw.freeze as Record<string, unknown>;
  if (Object.keys(freeze).sort().join(',') !== 'confirmed,started_at' || freeze.confirmed !== true) {
    fail('CUTOVER_FREEZE_REQUIRED', 'Explicitly confirm that pilot sales and physical movements remain paused.', 400);
  }
  utc(freeze.started_at);
  const batch = await db.prepare('SELECT * FROM opening_cutover_batches WHERE operation_id=?')
    .bind(raw.operation_id.toLowerCase()).first<Batch>();
  if (!batch) fail('CUTOVER_NOT_FOUND', 'The reviewed staged cutover does not exist.', 404);
  if (batch.store_id !== policy.storeId || batch.review_hash !== raw.expected_hash || batch.frozen_at !== freeze.started_at
    || batch.actor !== policy.actor.trim()) fail('CUTOVER_REVIEW_MISMATCH', 'The store, administrator, approved hash or freeze does not match the staged cutover.');
  const now = options.clock ?? (() => policy.now ?? new Date().toISOString());
  const rows = (await db.prepare('SELECT * FROM opening_cutover_rows WHERE operation_id=? ORDER BY item_id')
    .bind(batch.operation_id).all<StockRow>()).results;
  const snapshot: OpeningOrdersSnapshot = JSON.parse(batch.orders_json);
  if (await sha(snapshot) !== batch.orders_hash || snapshot.kind !== 'READONLY_ORDERS' || snapshot.store_id !== policy.storeId
    || snapshot.complete !== true || snapshot.orders.length !== batch.order_count || rows.length !== batch.row_count
    || !Number.isSafeInteger(snapshot.orders_checked) || snapshot.orders_checked < snapshot.orders.length || snapshot.orders_checked > MAX_ORDERS
    || snapshot.creation_cutoff !== Math.floor(utc(snapshot.started_at) / 1000)
    || utc(snapshot.started_at) < utc(batch.frozen_at) || utc(snapshot.completed_at) < utc(snapshot.started_at)
    || utc(snapshot.completed_at) > utc(now())) fail('CUTOVER_AUDIT_INVALID', 'The immutable staged cutover evidence is inconsistent.');
  const rowsByIdentity = new Map(rows.map(row => [stockKey(row.sku, row.ecwid_product_id, row.ecwid_combination_id, row.ecwid_option_signature), row]));
  const workbook = validateWorkbookTargets(JSON.parse(batch.workbook_scope_json));
  const orders = (await Promise.all(snapshot.orders.map(order=>orderIdentity(order,workbook,true)))).sort((a, b) => a.id.localeCompare(b.id));
  const lines = snapshot.orders.flatMap(order => order.items.map((line): LineIdentity => {
    const key = stockKey(line.sku, line.productId, line.combinationId, JSON.stringify(canonicalVariationOptions(line.selectedOptions)));
    const item = line.digital===false&&line.workbookOptionsEvidence===undefined?rowsByIdentity.get(key):undefined;
    const external=workbook.filter(target=>workbookLineMatches(line,target,true));
    const outside=external.length===1?external[0]:undefined;
    if (!item && !outside) fail('CUTOVER_AUDIT_INVALID', 'An opening order line no longer matches the immutable scope.');
    return { id: `${order.id}:${line.id}`, order_id: order.id, ecwid_line_id: line.id, sku: line.sku, quantity: line.quantity,
      item_id: item?.item_id ?? null, mode: item ? 'APP' : 'WORKBOOK', workbook_target_id: outside ? workbookTargetId(outside) : null };
  }));
  if (lines.length !== batch.line_count) fail('CUTOVER_AUDIT_INVALID', 'The staged order-line count does not match.');
  return { batch, rows, snapshot, orders, lines, now, policy, fetcher: options.fetcher,workbook,
    leaseStartedAt: batch.frozen_at, leaseMs: FREEZE_MS };
}

async function loadRecoveryContext(db: D1Database, value: unknown, policy: CutoverAlignmentPolicy,
  options: CutoverAlignmentOptions): Promise<{ context: Context; request: CutoverRecoveryRequest }> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) fail('CUTOVER_RECOVERY_REQUEST_INVALID', 'A reviewed recovery request is required.', 400);
  const raw = value as Record<string, unknown>;
  const allowed = ['operation_id', 'expected_hash', 'recovery_id', 'recovery_freeze'];
  if (Object.keys(raw).some(key => !allowed.includes(key)) || typeof raw.operation_id !== 'string' || !UUID.test(raw.operation_id)
    || typeof raw.expected_hash !== 'string' || !SHA.test(raw.expected_hash)
    || typeof raw.recovery_id !== 'string' || !UUID.test(raw.recovery_id)
    || !raw.recovery_freeze || typeof raw.recovery_freeze !== 'object' || Array.isArray(raw.recovery_freeze)) {
    fail('CUTOVER_RECOVERY_REQUEST_INVALID', 'Provide the exact operation, recovery ID, approved hash and new freeze confirmation.', 400);
  }
  const recoveryFreeze = raw.recovery_freeze as Record<string, unknown>;
  if (Object.keys(recoveryFreeze).sort().join(',') !== 'confirmed,started_at' || recoveryFreeze.confirmed !== true) {
    fail('CUTOVER_RECOVERY_FREEZE_REQUIRED', 'Explicitly confirm that store-wide Ecwid ordering and all pilot physical movements are paused.', 400);
  }
  const recoveryStarted = utc(recoveryFreeze.started_at);
  const stored = await db.prepare('SELECT frozen_at FROM opening_cutover_batches WHERE operation_id=?')
    .bind(raw.operation_id.toLowerCase()).first<{ frozen_at: string }>();
  if (!stored) fail('CUTOVER_NOT_FOUND', 'The reviewed staged cutover does not exist.', 404);
  if (recoveryStarted <= utc(stored.frozen_at)) fail('CUTOVER_RECOVERY_FREEZE_INVALID', 'The recovery freeze must be newer than the original cutover freeze.', 400);
  const context = await loadContext(db, { operation_id: raw.operation_id, expected_hash: raw.expected_hash,
    freeze: { confirmed: true, started_at: stored.frozen_at } }, policy, options);
  context.leaseStartedAt = String(recoveryFreeze.started_at);
  context.leaseMs = RECOVERY_FREEZE_MS;
  const request: CutoverRecoveryRequest = { operation_id: context.batch.operation_id, expected_hash: raw.expected_hash,
    recovery_id: raw.recovery_id.toLowerCase(), recovery_freeze: { confirmed: true, started_at: String(recoveryFreeze.started_at) } };
  return { context, request };
}

// Used in both read checks and the final write transaction. A conditional state
// transition deliberately violates the state constraint if any guard changes.
const DB_GUARD = `
 b.row_count=(SELECT COUNT(*) FROM opening_cutover_rows WHERE operation_id=b.operation_id)
 AND b.order_count=(SELECT COUNT(*) FROM opening_cutover_orders WHERE operation_id=b.operation_id)
 AND NOT EXISTS(SELECT 1 FROM opening_cutover_rows r LEFT JOIN item_stock i ON i.id=r.item_id
   LEFT JOIN opening_balances ob ON ob.id=r.opening_balance_id
   WHERE r.operation_id=b.operation_id AND (i.id IS NULL OR ob.id IS NULL OR i.active<>0
     OR i.inventory_mode<>'STOCK_LIMITED' OR i.on_hand<>r.physical OR i.reserved<>r.unpicked
     OR i.available<>r.target_quantity OR ob.item_id<>i.id OR ob.on_hand<>r.physical
     OR i.sku IS NOT r.sku COLLATE BINARY OR i.ecwid_product_id IS NOT r.ecwid_product_id
     OR i.ecwid_combination_id IS NOT r.ecwid_combination_id OR i.ecwid_option_signature IS NOT r.ecwid_option_signature
     OR i.last_ecwid_quantity IS NOT r.expected_ecwid_quantity
     OR EXISTS(SELECT 1 FROM movements WHERE item_id=i.id)
     OR EXISTS(SELECT 1 FROM outbox WHERE item_id=i.id)
     OR EXISTS(SELECT 1 FROM supplier_allocation_events WHERE item_id=i.id)
     OR NOT EXISTS(SELECT 1 FROM sync_issues WHERE item_id=i.id AND status='OPEN' AND kind='OPENING_CUTOVER_STAGED')))
 AND NOT EXISTS(SELECT 1 FROM sync_issues s WHERE s.status='OPEN' AND s.kind<>'OPENING_CUTOVER_STAGED' AND (
   s.item_id IN (SELECT item_id FROM opening_cutover_rows WHERE operation_id=b.operation_id)
   OR s.order_id IN (SELECT order_id FROM opening_cutover_orders WHERE operation_id=b.operation_id)))
 AND NOT EXISTS(SELECT 1 FROM json_each((SELECT orders_json FROM c)) e LEFT JOIN orders o ON o.id=e.value->>'id'
   LEFT JOIN opening_cutover_orders co ON co.order_id=o.id AND co.operation_id=b.operation_id
   WHERE o.id IS NULL OR co.order_id IS NULL OR o.needs_review<>0 OR o.payment_status IS NOT e.value->>'payment'
     OR o.fulfillment_status IS NOT e.value->>'fulfillment' OR o.remote_updated_at IS NOT e.value->>'updated'
     OR o.remote_lines_hash IS NOT e.value->>'hash' OR co.remote_lines_hash IS NOT e.value->>'hash'
     OR co.line_count<>e.value->>'lines' OR (SELECT COUNT(*) FROM order_lines WHERE order_id=o.id)<>e.value->>'lines')
 AND NOT EXISTS(SELECT 1 FROM json_each((SELECT lines_json FROM c)) e LEFT JOIN order_lines l ON l.id=e.value->>'id'
   WHERE l.id IS NULL OR l.order_id IS NOT e.value->>'order_id' OR l.ecwid_line_id IS NOT e.value->>'ecwid_line_id'
     OR l.sku IS NOT e.value->>'sku' COLLATE BINARY OR l.ordered_qty<>e.value->>'quantity' OR l.picked_qty<>0
     OR l.item_id IS NOT e.value->>'item_id' OR l.management_mode IS NOT e.value->>'mode'
     OR l.workbook_target_id IS NOT e.value->>'workbook_target_id')
 AND NOT EXISTS(SELECT 1 FROM sync_state WHERE key='orders_tracking_started' AND value<>b.frozen_at)
`;
const CONTEXT_SQL = 'WITH c AS (SELECT ? AS orders_json, ? AS lines_json)';
function guardBindings(context: Context): [string, string] { return [JSON.stringify(context.orders), JSON.stringify(context.lines)]; }
async function assertDatabase(db: D1Database, context: Context): Promise<void> {
  const good = await db.prepare(`${CONTEXT_SQL} SELECT (${DB_GUARD}) AS valid FROM opening_cutover_batches b WHERE operation_id=?`)
    .bind(...guardBindings(context), context.batch.operation_id).first<number>('valid');
  if (good !== 1) fail('CUTOVER_DATABASE_CHANGED', 'Staged physical stock, commitments, identities or review state changed. Reconcile before activation.');
}
async function review(db: D1Database, context: Context): Promise<void> {
  await db.prepare("UPDATE opening_cutover_batches SET state='REVIEW',updated_at=? WHERE operation_id=? AND state IN ('STAGED','ALIGNING','ALIGNED')")
    .bind(context.now(), context.batch.operation_id).run();
}
function pending(order: EcwidOrder): boolean {
  return !['CANCELLED', 'REFUNDED'].includes(order.paymentStatus)
    && !TERMINAL_FULFILLMENT_STATUSES.includes(order.fulfillmentStatus);
}

/** Complete unfiltered pagination prevents status changes shifting a filtered list.
 * At most 50,000 IDs plus 200 pending identities are retained, not historical orders.
 * Full scans may exceed a Free Worker's subrequest limit: invoke only from trusted
 * local orchestration or an explicitly budgeted runtime. No HTTP route exposes it.
 */
async function assertLiveOrders(context: Context): Promise<number> {
  const cutoff = Math.floor(utc(context.now()) / 1000);
  if (cutoff < context.snapshot.creation_cutoff) fail('CUTOVER_CLOCK_INVALID', 'The current cutoff predates the staged snapshot.');
  let count = 0; let expectedTotal: number | undefined;
  const ids = new Set<string>(); const open: OrderIdentity[] = [];
  for (;;) {
    const page = await client(context).listOrders({ offset: count, createdTo: cutoff });
    remainingLease(context);
    if (!Number.isSafeInteger(page.total) || page.total < 0 || page.total > MAX_ORDERS || page.total !== context.snapshot.orders_checked
      || (expectedTotal !== undefined && page.total !== expectedTotal) || page.offset !== count || page.count !== page.items.length
      || page.count > 100 || count + page.count > page.total || (!page.count && count < page.total)) {
      fail('CUTOVER_ORDERS_CHANGED', 'The complete order listing changed or its pagination is inconsistent.');
    }
    expectedTotal = page.total;
    for (const order of page.items) {
      if (ids.has(order.id) || !order.id.trim()) fail('CUTOVER_ORDERS_CHANGED', 'The order listing contains duplicate or missing identities.');
      ids.add(order.id);
      if (order.createdAt && Math.floor(utc(order.createdAt) / 1000) > context.snapshot.creation_cutoff) {
        fail('CUTOVER_ORDERS_CHANGED', 'An order was created after the reviewed snapshot.');
      }
      if (pending(order)) {
        if (open.length >= 200 || utc(order.updatedAt) > utc(context.now())) fail('CUTOVER_ORDERS_CHANGED', 'Pending order evidence is inconsistent.');
        open.push(await orderIdentity(order,context.workbook));
      }
    }
    count += page.count;
    if (count === page.total) break;
  }
  open.sort((a, b) => a.id.localeCompare(b.id));
  if (JSON.stringify(open) !== JSON.stringify(context.orders)) fail('CUTOVER_ORDERS_CHANGED', 'Pending order identities, statuses, timestamps or line quantities changed.');
  return cutoff;
}
async function assertNoNewOrders(context: Context): Promise<void> {
  const page = await client(context).listOrders({ offset: 0, createdTo: Math.floor(utc(context.now()) / 1000) });
  remainingLease(context);
  if (page.offset !== 0 || page.total !== context.snapshot.orders_checked || page.count !== page.items.length
    || page.count !== Math.min(100, page.total)) fail('CUTOVER_ORDERS_CHANGED', 'The order count changed during final verification.');
}
function assertTarget(row: StockRow, target: EcwidProduct, expectedQuantity: number): void {
  const signature = JSON.stringify(canonicalVariationOptions(target.variationOptions));
  const variation = row.ecwid_combination_id !== null;
  if (target.id !== row.ecwid_product_id || (target.combinationId ?? null) !== row.ecwid_combination_id
    || target.sku !== row.sku || signature !== row.ecwid_option_signature || target.enabled !== true
    || target.unlimited !== false || target.eligibilityVerified !== true || target.hasBundleRelationships !== false
    || target.hasExtraOptions !== false || target.hasVariations !== false || target.hasOptions !== variation
    || (variation ? signature === '[]' : signature !== '[]') || target.quantity !== expectedQuantity) {
    fail('CUTOVER_TARGET_CHANGED', 'An exact stock target, eligibility setting or reviewed quantity changed. No policy changes are permitted.');
  }
}
function rowStatusEvidence(rows: StockRow[]) {
  return rows.map(row => ({ item_id: row.item_id, alignment_status: row.alignment_status,
    before_quantity: row.before_quantity, after_quantity: row.after_quantity }));
}
async function assertRecoveryTargets(context: Context): Promise<string> {
  const evidence: Array<Record<string, unknown>> = [];
  for (const row of context.rows) {
    if (!['PENDING', 'VERIFIED'].includes(row.alignment_status)) {
      fail('CUTOVER_ROW_HELD', 'Recovery cannot reset or retry a previously attempted unresolved row.');
    }
    const expected = row.alignment_status === 'VERIFIED' ? row.target_quantity : row.expected_ecwid_quantity;
    const target = await client(context).getProductStock(row.ecwid_product_id, row.ecwid_combination_id);
    remainingLease(context);
    assertTarget(row, target, expected);
    evidence.push({ item_id: row.item_id, sku: row.sku, ecwid_product_id: row.ecwid_product_id,
      ecwid_combination_id: row.ecwid_combination_id, ecwid_option_signature: row.ecwid_option_signature,
      alignment_status: row.alignment_status, required_quantity: expected, live_quantity: target.quantity });
  }
  return sha(evidence);
}
async function receipt(db: D1Database, context: Context, duplicate = false) {
  const status = await db.prepare('SELECT state FROM opening_cutover_batches WHERE operation_id=?')
    .bind(context.batch.operation_id).first<BatchState>('state');
  const verified = await db.prepare("SELECT COUNT(*) AS count FROM opening_cutover_rows WHERE operation_id=? AND alignment_status='VERIFIED'")
    .bind(context.batch.operation_id).first<number>('count');
  return { operation_id: context.batch.operation_id, store_id: context.batch.store_id,
    review_hash: context.batch.review_hash, state: status, row_count: context.batch.row_count,
    verified_count: verified ?? 0, activated: status === 'ACTIVE', duplicate };
}

export async function beginCutoverAlignment(db: D1Database, value: unknown, policy: CutoverAlignmentPolicy,
  options: CutoverAlignmentOptions = {}) {
  const context = await loadContext(db, value, policy, options);
  if (context.batch.state === 'ALIGNING') return receipt(db, context, true);
  if (context.batch.state !== 'STAGED') fail('CUTOVER_STATE_INVALID', 'Only a newly staged cutover may begin alignment.');
  // Completed-operation receipts remain readable after the freeze expires.
  // Every new action checks the lease before network access or state changes.
  remainingLease(context);
  try {
    await assertDatabase(db, context);
    await assertLiveOrders(context);
    await assertNoNewOrders(context);
    remainingLease(context);
    const changed = await db.prepare(`${CONTEXT_SQL} UPDATE opening_cutover_batches AS b
      SET state=CASE WHEN (${DB_GUARD}) THEN 'ALIGNING' ELSE 'INVALID' END,updated_at=?
      WHERE operation_id=? AND state='STAGED' RETURNING operation_id`)
      .bind(...guardBindings(context), context.now(), context.batch.operation_id).first();
    if (!changed) {
      const current = await db.prepare('SELECT state FROM opening_cutover_batches WHERE operation_id=?')
        .bind(context.batch.operation_id).first<BatchState>('state');
      if (current === 'ALIGNING' || current === 'ALIGNED' || current === 'ACTIVE') return receipt(db, context, true);
      fail('CUTOVER_STATE_INVALID', 'The cutover state changed before alignment began.');
    }
    return receipt(db, context);
  } catch (error) {
    await review(db, context);
    if (error instanceof DomainError) throw error;
    fail('CUTOVER_REVIEW_REQUIRED', 'Fresh order or database verification failed. No stock write was attempted.');
  }
}

/** At most one PUT per invocation. Persisted PROCESSING/UNKNOWN/BLOCKED cannot be retried. */
export async function alignCutoverRow(db: D1Database, value: unknown, policy: CutoverAlignmentPolicy,
  options: CutoverAlignmentOptions = {}) {
  const context = await loadContext(db, value, policy, options, true);
  return alignLoadedCutoverRow(db, context, value);
}

async function alignLoadedCutoverRow(db: D1Database, context: Context, value: unknown, compact = false,
  beforeStockWrite?: () => Promise<void>) {
  const itemId = (value as Record<string, unknown>).item_id;
  if (typeof itemId !== 'string' || !UUID.test(itemId)) fail('CUTOVER_ITEM_INVALID', 'Use the exact staged item ID.', 400);
  const row = context.rows.find(row => row.item_id === itemId);
  if (!row) fail('CUTOVER_ITEM_INVALID', 'This item does not belong to the reviewed cutover.', 404);
  if (context.batch.state !== 'ALIGNING') fail('CUTOVER_STATE_INVALID', 'Alignment is not active or requires manual review.');
  if (row.alignment_status === 'VERIFIED') return compact
    ? { item_id: itemId, alignment_status: 'VERIFIED' as const, skipped: true }
    : { ...await receipt(db, context, true), item_id: itemId, alignment_status: 'VERIFIED' as const };
  if (row.alignment_status !== 'PENDING') fail('CUTOVER_ROW_HELD', 'This stock target was already attempted. Reconcile its journal before any further write.');
  if (context.rows.some(row => row.alignment_status === 'PROCESSING')) fail('CUTOVER_ROW_BUSY', 'Another stock target has an unresolved in-flight operation.');
  remainingLease(context);
  let claimed = false; let writeAttempted = false; let writeConfirmed = false;
  try {
    await assertDatabase(db, context);
    const target = await client(context).getProductStock(row.ecwid_product_id, row.ecwid_combination_id);
    remainingLease(context);
    assertTarget(row, target, row.expected_ecwid_quantity);
    const noOp = row.expected_ecwid_quantity === row.target_quantity;
    if (!noOp) await beforeStockWrite?.();
    if (compact && remainingLease(context) <= RECOVERY_WRITE_BUFFER_MS) {
      fail('CUTOVER_RECOVERY_FREEZE_EXPIRING', 'Recovery stopped before claiming another row. No stock write was attempted.');
    }
    const timestamp = context.now();
    const claim = await db.prepare(`${CONTEXT_SQL} UPDATE opening_cutover_rows SET alignment_status=?,before_quantity=?,
      attempted_at=?,after_quantity=?,verified_at=? WHERE operation_id=? AND item_id=? AND alignment_status='PENDING'
      AND EXISTS(SELECT 1 FROM opening_cutover_batches b WHERE b.operation_id=opening_cutover_rows.operation_id
        AND b.state='ALIGNING' AND (${DB_GUARD}))
      AND NOT EXISTS(SELECT 1 FROM opening_cutover_rows running WHERE running.operation_id=opening_cutover_rows.operation_id
        AND running.alignment_status='PROCESSING') RETURNING item_id`)
      .bind(...guardBindings(context), noOp ? 'VERIFIED' : 'PROCESSING', target.quantity, timestamp,
        noOp ? target.quantity : null, noOp ? timestamp : null, context.batch.operation_id, itemId).first();
    if (!claim) fail('CUTOVER_ROW_BUSY', 'The row, database or batch changed before its write could be claimed.');
    claimed = true;
    if (!noOp) {
      remainingLease(context);
      writeAttempted = true;
      await client(context).setStockQuantity(row.ecwid_product_id, row.target_quantity, row.ecwid_combination_id);
      writeConfirmed = true;
      const after = await client(context).getProductStock(row.ecwid_product_id, row.ecwid_combination_id);
      remainingLease(context);
      assertTarget(row, after, row.target_quantity);
      const verified = await db.prepare(`UPDATE opening_cutover_rows SET alignment_status='VERIFIED',after_quantity=?,verified_at=?
        WHERE operation_id=? AND item_id=? AND alignment_status='PROCESSING'
        AND EXISTS(SELECT 1 FROM opening_cutover_batches WHERE operation_id=? AND state='ALIGNING') RETURNING item_id`)
        .bind(after.quantity, context.now(), context.batch.operation_id, itemId, context.batch.operation_id).first();
      if (!verified) fail('CUTOVER_REVIEW_REQUIRED', 'The write was confirmed remotely but could not be verified in its journal.');
    }
    return { ...(compact ? {} : await receipt(db, context)), item_id: itemId, alignment_status: 'VERIFIED' as const, no_op: noOp };
  } catch (error) {
    if (compact && !claimed) throw error;
    if (error instanceof DomainError && error.code === 'CUTOVER_ROW_BUSY' && !claimed) throw error;
    const uncertain = writeAttempted && (writeConfirmed || !(error instanceof EcwidError && error.outcome !== 'UNKNOWN'));
    const status = uncertain ? 'UNKNOWN' : 'BLOCKED';
    const message = uncertain ? 'Stock write outcome or verification is uncertain. Never retry automatically.'
      : 'Stock preflight or the write was rejected. Review the frozen cutover before continuing.';
    // A concurrent halt or unavailable DB may leave PROCESSING. That is equally
    // fail-closed and must never be recovered to PENDING automatically.
    await db.prepare(`UPDATE opening_cutover_rows SET alignment_status=?,last_error=?
      WHERE operation_id=? AND item_id=? AND alignment_status=?
      AND EXISTS(SELECT 1 FROM opening_cutover_batches WHERE operation_id=? AND state='ALIGNING')`)
      .bind(status, message, context.batch.operation_id, itemId, claimed ? 'PROCESSING' : 'PENDING', context.batch.operation_id).run();
    await review(db, context);
    fail('CUTOVER_REVIEW_REQUIRED', message);
  }
}

/** No separate activation endpoint: final network verification and activation
 * share one call and one atomic DB transaction. The original freeze watermark
 * ensures an order created and shipped before polling starts is not skipped.
 */
export async function finishAndActivateCutover(db: D1Database, value: unknown, policy: CutoverAlignmentPolicy,
  options: CutoverAlignmentOptions = {}) {
  const context = await loadContext(db, value, policy, options);
  return finishLoadedCutover(db, context);
}

async function finishLoadedCutover(db: D1Database, context: Context, beforeActivation?: () => Promise<void>,
  preserveOnLeaseExpiry = false) {
  if (context.batch.state === 'ACTIVE') return receipt(db, context, true);
  if (context.batch.state !== 'ALIGNING') fail('CUTOVER_STATE_INVALID', 'Only a fully aligned in-progress cutover may be activated.');
  if (context.rows.some(row => row.alignment_status !== 'VERIFIED')) fail('CUTOVER_NOT_VERIFIED', 'Every stock target must be verified before activation.');
  remainingLease(context);
  let activationAttempted = false;
  try {
    await assertDatabase(db, context);
    // Check each exact target, then perform the complete order scan last so any
    // order changes during product read-back are detected before activation.
    for (const row of context.rows) {
      const target = await client(context).getProductStock(row.ecwid_product_id, row.ecwid_combination_id);
      remainingLease(context); assertTarget(row, target, row.target_quantity);
    }
    await assertLiveOrders(context);
    await assertNoNewOrders(context);
    remainingLease(context);
    await beforeActivation?.();
    remainingLease(context);
    const timestamp = context.now();
    activationAttempted = true;
    await db.batch([
      db.prepare(`${CONTEXT_SQL} UPDATE opening_cutover_batches AS b SET
        state=CASE WHEN b.state='ALIGNING' AND (${DB_GUARD}) THEN 'ALIGNED' ELSE 'INVALID' END,updated_at=?
        WHERE operation_id=?`).bind(...guardBindings(context), timestamp, context.batch.operation_id),
      db.prepare(`UPDATE items SET active=1,last_ecwid_quantity=(SELECT target_quantity FROM opening_cutover_rows WHERE item_id=items.id)
        WHERE id IN (SELECT item_id FROM opening_cutover_rows WHERE operation_id=?)`).bind(context.batch.operation_id),
      db.prepare(`UPDATE sync_issues SET status='RESOLVED',resolved_at=? WHERE status='OPEN' AND kind='OPENING_CUTOVER_STAGED'
        AND item_id IN (SELECT item_id FROM opening_cutover_rows WHERE operation_id=?)`).bind(timestamp, context.batch.operation_id),
      db.prepare(`INSERT INTO sync_state(key,value,updated_at) VALUES('orders_tracking_started',?,?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`).bind(context.batch.frozen_at, timestamp),
      db.prepare("UPDATE opening_cutover_batches SET state='ACTIVE',updated_at=? WHERE operation_id=? AND state='ALIGNED'")
        .bind(timestamp, context.batch.operation_id),
    ]);
    return receipt(db, context);
  } catch (error) {
    const code = error && typeof error === 'object' && 'code' in error ? String(error.code) : '';
    if (preserveOnLeaseExpiry && !activationAttempted
      && ['CUTOVER_FREEZE_EXPIRED', 'CUTOVER_RECOVERY_FREEZE_EXPIRING'].includes(code)) throw error;
    await review(db, context);
    if (error instanceof DomainError) throw error;
    fail('CUTOVER_REVIEW_REQUIRED', 'Final stock, order or atomic database verification failed. Inventory remains disabled.');
  }
}

/** One trusted recovery session: fresh full preflight, untouched PENDING rows only,
 * then the ordinary full read-back and atomic activation. The request and evidence
 * hash form the recovery receipt; existing per-row transitions remain the durable
 * no-retry journal. A stopped process requires a new explicit recovery freeze.
 */
export async function recoverAndActivateCutover(db: D1Database, value: unknown, policy: CutoverAlignmentPolicy,
  options: CutoverRecoveryOptions = {}) {
  const loaded = await loadRecoveryContext(db, value, policy, options);
  const { context, request } = loaded;
  if (context.batch.state !== 'ALIGNING') fail('CUTOVER_STATE_INVALID', 'Only the existing partially aligned batch may enter recovery.');
  const verified = context.rows.filter(row => row.alignment_status === 'VERIFIED').length;
  const pendingRows = context.rows.filter(row => row.alignment_status === 'PENDING');
  if (!verified || verified + pendingRows.length !== context.rows.length) {
    fail('CUTOVER_RECOVERY_ROWS_INVALID', 'Recovery requires preserved VERIFIED rows and only untouched PENDING rows.');
  }
  remainingLease(context);
  let liveTargetsHash: string; let cutoff: number;
  try {
    await assertDatabase(db, context);
    liveTargetsHash = await assertRecoveryTargets(context);
    cutoff = await assertLiveOrders(context);
    await assertNoNewOrders(context);
  } catch (error) {
    if (error instanceof DomainError) throw error;
    fail('CUTOVER_RECOVERY_PREFLIGHT_FAILED', 'Fresh target or order validation failed. No recovery stock write was attempted.');
  }
  if (remainingLease(context) <= RECOVERY_WRITE_BUFFER_MS) {
    fail('CUTOVER_RECOVERY_FREEZE_EXPIRING', 'Less than one minute remains in the recovery freeze. No recovery stock write was attempted.');
  }
  const evidence = { schema_version: 1, recovery_id: request.recovery_id, operation_id: context.batch.operation_id,
    review_hash: context.batch.review_hash, original_frozen_at: context.batch.frozen_at,
    recovery_frozen_at: request.recovery_freeze.started_at, row_count: context.rows.length,
    initial_verified_count: verified, pending_count: pendingRows.length,
    row_status_hash: await sha(rowStatusEvidence(context.rows)), live_targets_hash: liveTargetsHash,
    orders_hash: context.batch.orders_hash, orders_checked: context.snapshot.orders_checked, orders_cutoff: cutoff };
  const evidenceHash = await sha(evidence);
  options.onPreflight?.({ event: 'recovery_preflight_verified', ...evidence, evidence_hash: evidenceHash });
  let resumed = 0;
  for (const row of pendingRows) {
    if (remainingLease(context) <= RECOVERY_FINALIZATION_BUFFER_MS) {
      fail('CUTOVER_RECOVERY_FREEZE_EXPIRING', 'Recovery stopped with five minutes reserved for a fresh final verification. No uncertain write was retried.');
    }
    const result = await alignLoadedCutoverRow(db, context, { item_id: row.item_id }, true, options.beforeStockWrite);
    resumed += 1;
    options.onRow?.({ event: 'row_verified', recovery_id: request.recovery_id, evidence_hash: evidenceHash, ...result });
  }
  const refreshed = (await loadRecoveryContext(db, value, policy, options)).context;
  if (remainingLease(refreshed) <= RECOVERY_FINALIZATION_BUFFER_MS) {
    fail('CUTOVER_RECOVERY_FREEZE_EXPIRING', 'Recovery stopped before final verification. Start a new explicitly confirmed pause; verified rows will not be replayed.');
  }
  const activated = await finishLoadedCutover(db, refreshed, options.beforeActivation, true);
  return { ...activated, recovery_id: request.recovery_id, recovery_frozen_at: request.recovery_freeze.started_at,
    recovery_evidence_hash: evidenceHash, recovery_initial_verified_count: verified, recovery_resumed_count: resumed };
}
