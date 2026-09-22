import { EcwidClient, EcwidError, canonicalVariationOptions, parseWebhook, readBoundedJson, verifyWebhookSignature,
  type EcwidOrder, type EcwidWebhook, type EcwidFetch, type EcwidProduct } from './ecwid';
import { DomainError, PICKABLE_FULFILLMENT_STATUSES, TERMINAL_FULFILLMENT_STATUSES, type InventoryMode } from './domain';
import { opaqueWorkbookPolicy, orderOptionsHashMaterial, workbookLineMatches } from './workbook-identity';
import type { WorkbookTarget } from './pilot-scope';

export type SyncMessage = { kind: 'outbox' | 'webhook'; id: string };
export type SyncEnv = Pick<Env, 'DB' | 'SYNC_QUEUE' | 'ECWID_MODE' | 'LIVE_SYNC_ENABLED' | 'ECWID_STORE_ID'> & {
  ORDER_SYNC_ENABLED?: string;
  ECWID_TOKEN?: string;
  ECWID_CLIENT_SECRET?: string;
};

interface OutboxRow { id: string; item_id: string; ecwid_product_id: string; ecwid_combination_id: string | null; quantity_delta: number; status: string }
export interface ProductMapping { id: string; sku: string; ecwid_combination_id: string | null; ecwid_option_signature: string; inventory_mode: InventoryMode }
interface InboxRow { event_id: string; event_type: string; entity_id: string; payload: string }
const now = () => new Date().toISOString();

export function ecwidClient(env: SyncEnv, fetcher?: EcwidFetch): EcwidClient {
  return new EcwidClient({ storeId: env.ECWID_STORE_ID, token: env.ECWID_TOKEN ?? '' }, fetcher);
}

async function getState(db: D1Database, key: string): Promise<string | null> {
  return db.prepare('SELECT value FROM sync_state WHERE key=?').bind(key).first<string>('value');
}

function issueStatement(db: D1Database, id: string, kind: string, item: string | null, order: string | null, message: string) {
  return db.prepare(`INSERT INTO sync_issues(id,kind,item_id,order_id,message,status,created_at)
    VALUES (?,?,?,?,?,'OPEN',?) ON CONFLICT(id) DO NOTHING`).bind(id, kind, item, order, message, now());
}

export async function orderSnapshotHash(order: EcwidOrder, options: { allowSnapshotEvidence?: boolean; workbookTargets?: WorkbookTarget[] } = {}): Promise<string> {
  const canonical = (await Promise.all(order.items.map(async line => [line.id, line.productId, line.sku, line.quantity,
    line.combinationId, await orderOptionsHashMaterial(line,options.allowSnapshotEvidence===true,
      options.workbookTargets?.some(target=>opaqueWorkbookPolicy(target)&&workbookLineMatches(line,target,options.allowSnapshotEvidence===true))??false), line.digital])))
    .sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(canonical)));
  return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

/** Snapshot updates and line imports are one transaction; existing picked quantities never change. */
export async function upsertOrderSnapshot(db: D1Database, order: EcwidOrder,
  options: { skipUntrackedTerminal?: boolean } = {}): Promise<{ id: string; needs_review: boolean; skipped?: boolean }> {
  const existing = await db.prepare('SELECT id,remote_updated_at FROM orders WHERE id=?').bind(order.id)
    .first<{ id: string; remote_updated_at: string }>();
  const terminal = ['CANCELLED', 'REFUNDED', 'INCOMPLETE'].includes(order.paymentStatus) || TERMINAL_FULFILLMENT_STATUSES.includes(order.fulfillmentStatus);
  if (!existing && terminal && options.skipUntrackedTerminal) return { id: order.id, needs_review: false, skipped: true };
  if (existing && existing.remote_updated_at > order.updatedAt) {
    const current = await db.prepare('SELECT needs_review FROM orders WHERE id=?').bind(order.id).first<number>('needs_review');
    return { id: order.id, needs_review: current === 1, skipped: true };
  }
  const workbookTargets=(await db.prepare(`SELECT id,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,sku,name,sku_source,option_policy
    FROM workbook_managed_targets`).all<WorkbookTarget & {id:string}>()).results;
  const hash = await orderSnapshotHash(order,{workbookTargets});
  const timestamp = now();
  const lines = JSON.stringify(order.items.map((line) => {
    const options = canonicalVariationOptions(line.selectedOptions);
    const outside=workbookTargets.filter(target=>workbookLineMatches(line,target));
    return { id:line.id,productId:line.productId,sku:line.sku,name:line.name,quantity:line.quantity,combinationId:line.combinationId,
    // Ecwid documents order-line trackQuantity as the low-stock notification flag.
    // Catalog mapping validation checks actual quantity/unlimited settings instead.
      optionSignature: outside.length===1?outside[0].ecwid_option_signature:options === null ? null : JSON.stringify(options),
      supported: !line.digital && Boolean(line.sku) && options !== null
        && (line.combinationId ? options.length > 0 : options.length === 0) ? 1 : 0,
      workbookTargetId: outside.length===1?outside[0].id:null,
    };
  }));
  const unsupportedStatus = !['PAID', 'AWAITING_PAYMENT', 'CANCELLED', 'REFUNDED', 'INCOMPLETE'].includes(order.paymentStatus)
    || (!PICKABLE_FULFILLMENT_STATUSES.includes(order.fulfillmentStatus) && !TERMINAL_FULFILLMENT_STATUSES.includes(order.fulfillmentStatus));
  const statements = [
    db.prepare(`INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at,needs_review,remote_lines_hash)
      VALUES (?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET
        needs_review=CASE WHEN orders.needs_review=1 OR excluded.needs_review=1
          OR (orders.remote_lines_hash!='' AND orders.remote_lines_hash!=excluded.remote_lines_hash)
          OR (orders.remote_updated_at=excluded.remote_updated_at AND
            (orders.payment_status!=excluded.payment_status OR orders.fulfillment_status!=excluded.fulfillment_status))
          THEN 1 ELSE 0 END,
        payment_status=CASE WHEN orders.remote_updated_at<excluded.remote_updated_at THEN excluded.payment_status ELSE orders.payment_status END,
        fulfillment_status=CASE WHEN orders.remote_updated_at<excluded.remote_updated_at THEN excluded.fulfillment_status ELSE orders.fulfillment_status END,
        remote_updated_at=excluded.remote_updated_at,updated_at=excluded.updated_at,
        remote_lines_hash=CASE WHEN orders.remote_lines_hash='' THEN excluded.remote_lines_hash ELSE orders.remote_lines_hash END
      WHERE orders.remote_updated_at<=excluded.remote_updated_at`)
      .bind(order.id, order.paymentStatus, order.fulfillmentStatus, order.updatedAt, timestamp, unsupportedStatus ? 1 : 0, hash),
    db.prepare(`INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty,picked_qty,management_mode,workbook_target_id)
      SELECT ? || ':' || json_extract(j.value,'$.id'), ?, json_extract(j.value,'$.id'),
        CASE WHEN json_extract(j.value,'$.supported')=1 THEN i.id ELSE NULL END,
        json_extract(j.value,'$.sku'),json_extract(j.value,'$.name'),json_extract(j.value,'$.quantity'),0,
        CASE WHEN w.id IS NOT NULL THEN 'WORKBOOK' ELSE 'APP' END,w.id
      FROM json_each(?) j LEFT JOIN items i ON i.ecwid_product_id=json_extract(j.value,'$.productId')
        AND i.ecwid_combination_id IS json_extract(j.value,'$.combinationId')
        AND i.ecwid_option_signature=json_extract(j.value,'$.optionSignature')
        AND i.sku=json_extract(j.value,'$.sku') COLLATE NOCASE AND i.active=1
      LEFT JOIN workbook_managed_targets w ON w.id=json_extract(j.value,'$.workbookTargetId')
        AND w.ecwid_product_id=json_extract(j.value,'$.productId')
        AND w.ecwid_combination_id IS json_extract(j.value,'$.combinationId')
        AND w.sku=json_extract(j.value,'$.sku') COLLATE NOCASE
        AND NOT EXISTS(SELECT 1 FROM items conflict WHERE conflict.sku=json_extract(j.value,'$.sku') COLLATE NOCASE
          OR (conflict.ecwid_product_id=json_extract(j.value,'$.productId')
            AND conflict.ecwid_combination_id IS json_extract(j.value,'$.combinationId')))
      WHERE EXISTS(SELECT 1 FROM orders WHERE id=? AND remote_updated_at=? AND remote_lines_hash=?)
      ON CONFLICT(order_id,ecwid_line_id) DO NOTHING`)
      .bind(order.id, order.id, lines, order.id, order.updatedAt, hash),
    db.prepare(`UPDATE orders SET needs_review=1 WHERE id=? AND remote_updated_at=? AND (
      EXISTS(SELECT 1 FROM order_lines WHERE order_id=orders.id AND management_mode='APP' AND item_id IS NULL)
      OR (payment_status NOT IN ('PAID','AWAITING_PAYMENT') AND EXISTS(
        SELECT 1 FROM order_lines WHERE order_id=orders.id AND picked_qty>0))
      OR (fulfillment_status IN ('READY_FOR_PICKUP','SHIPPED','DELIVERED','OUT_FOR_DELIVERY','RETURNED','WILL_NOT_DELIVER') AND EXISTS(
        SELECT 1 FROM order_lines WHERE order_id=orders.id AND management_mode='APP' AND picked_qty<ordered_qty)))`)
      .bind(order.id, order.updatedAt),
    db.prepare(`INSERT INTO sync_issues(id,kind,item_id,order_id,message,status,created_at)
      SELECT 'order:' || o.id || ':' || coalesce(l.item_id,'unmapped'), 'ORDER_REVIEW', l.item_id, o.id,
        'Order requires review: changed lines, unsupported item/status, missing picks, or cancellation/refund after picking. Check Ecwid and physical stock before continuing.',
        'OPEN', ? FROM orders o LEFT JOIN order_lines l ON l.order_id=o.id
      WHERE o.id=? AND o.needs_review=1 GROUP BY l.item_id ON CONFLICT(id) DO NOTHING`)
      .bind(timestamp, order.id),
    // Edited or contradictory identities can implicate the exact target and/or
    // same-parent SKU. Quarantine those candidates, never every sibling.
    db.prepare(`INSERT INTO sync_issues(id,kind,item_id,order_id,message,status,created_at)
      SELECT 'order:' || ? || ':' || i.id,'ORDER_REVIEW',i.id,?,
        'An Ecwid order was edited or needs review. Reconcile its inventory effect before continuing.','OPEN',?
      FROM json_each(?) j JOIN items i ON i.ecwid_product_id=json_extract(j.value,'$.productId')
        AND (i.ecwid_combination_id IS json_extract(j.value,'$.combinationId')
          OR i.sku=json_extract(j.value,'$.sku') COLLATE NOCASE)
      WHERE EXISTS(SELECT 1 FROM orders WHERE id=? AND needs_review=1 AND remote_updated_at=?)
      GROUP BY i.id ON CONFLICT(id) DO NOTHING`).bind(order.id, order.id, timestamp, lines, order.id, order.updatedAt),
  ];
  await db.batch(statements);
  const needsReview = await db.prepare('SELECT needs_review FROM orders WHERE id=?').bind(order.id).first<number>('needs_review');
  return { id: order.id, needs_review: needsReview === 1 };
}

export async function refreshOrder(env: SyncEnv, id: string, fetcher?: EcwidFetch) {
  if (env.ECWID_MODE !== 'live') return { id, demo: true };
  if (env.ORDER_SYNC_ENABLED !== 'true') throw new DomainError(409, 'ORDER_SYNC_PAUSED', 'Order imports are paused until the reviewed opening cutover is complete.');
  return upsertOrderSnapshot(env.DB, await ecwidClient(env, fetcher).getOrder(id));
}

/** Queue is a delivery hint; D1 remains the durable source if send() fails. */
export async function enqueueSync(env: SyncEnv, message: SyncMessage): Promise<boolean> {
  try {
    await env.SYNC_QUEUE.send(message);
    return true;
  } catch {
    console.warn(JSON.stringify({ event: 'sync_queue_unavailable', kind: message.kind, id: message.id }));
    return false;
  }
}

export async function ingestWebhook(request: Request, env: SyncEnv): Promise<Response> {
  if (env.ECWID_MODE !== 'live' || !env.ECWID_CLIENT_SECRET) return Response.json({ error: 'Webhook integration is not configured.' }, { status: 503 });
  if (env.ORDER_SYNC_ENABLED !== 'true') return Response.json({ error: 'Order imports are paused until cutover is complete.', code: 'ORDER_SYNC_PAUSED' }, { status: 503 });
  let event: EcwidWebhook;
  let raw: unknown;
  try {
    raw = await readBoundedJson(request, 64_000);
    event = parseWebhook(raw);
  } catch {
    return Response.json({ error: 'Invalid webhook.' }, { status: 400 });
  }
  if (event.storeId !== env.ECWID_STORE_ID || !await verifyWebhookSignature(event,
    request.headers.get('X-Ecwid-Webhook-Signature'), env.ECWID_CLIENT_SECRET)) {
    return Response.json({ error: 'Invalid webhook signature or store.' }, { status: 401 });
  }
  const timestamp = now();
  await env.DB.prepare(`INSERT INTO webhook_events(event_id,event_type,entity_id,store_id,payload,status,attempts,received_at,updated_at)
    VALUES (?,?,?,?,?,'PENDING',0,?,?) ON CONFLICT(event_id) DO NOTHING`)
    .bind(event.eventId, event.eventType, event.entityId, event.storeId, JSON.stringify(raw), timestamp, timestamp).run();
  await enqueueSync(env, { kind: 'webhook', id: event.eventId });
  return Response.json({ received: true }, { status: 202 });
}

export async function claimOutbox(db: D1Database, id: string): Promise<OutboxRow | null> {
  return db.prepare(`UPDATE outbox SET status='PROCESSING',attempts=attempts+1,updated_at=?
    WHERE id=? AND status='PENDING'
      AND EXISTS(SELECT 1 FROM items i WHERE i.id=outbox.item_id AND i.inventory_mode='STOCK_LIMITED')
      AND NOT EXISTS(SELECT 1 FROM movements m WHERE m.id=outbox.id AND m.inventory_mode='SUPPLIER_BACKED_UNLIMITED')
      AND NOT EXISTS(SELECT 1 FROM sync_issues s WHERE s.item_id=outbox.item_id AND s.status='OPEN')
      AND NOT EXISTS(SELECT 1 FROM outbox other WHERE other.item_id=outbox.item_id AND other.id!=outbox.id
        AND (other.status IN ('PROCESSING','UNKNOWN','BLOCKED') OR (other.status='PENDING'
          AND other.rowid<outbox.rowid)))
    RETURNING id,item_id,ecwid_product_id,ecwid_combination_id,quantity_delta,status`).bind(now(), id).first<OutboxRow>();
}

async function failOutbox(db: D1Database, row: OutboxRow, status: 'UNKNOWN' | 'BLOCKED', message: string): Promise<void> {
  await db.batch([
    db.prepare(`UPDATE outbox SET status=?,last_error=?,updated_at=? WHERE id=? AND status='PROCESSING'`)
      .bind(status, message, now(), row.id),
    issueStatement(db, `outbox:${row.id}`, `STOCK_${status}`, row.item_id, null, message),
  ]);
}

async function enqueueNextOutbox(env: SyncEnv, itemId: string): Promise<void> {
  const next = await env.DB.prepare(`SELECT id FROM outbox WHERE item_id=? AND status='PENDING'
    ORDER BY rowid LIMIT 1`).bind(itemId).first<{ id: string }>();
  if (next) await enqueueSync(env, { kind: 'outbox', id: next.id });
}

export async function processOutbox(env: SyncEnv, id: string, fetcher?: EcwidFetch): Promise<void> {
  // Belt-and-braces protection for legacy or manually corrupted outbox rows.
  // This runs even with live sync disabled and never calls the remote API.
  const forbidden = await env.DB.prepare(`SELECT x.id,x.item_id FROM outbox x JOIN items i ON i.id=x.item_id
    WHERE x.id=? AND (i.inventory_mode='SUPPLIER_BACKED_UNLIMITED'
      OR EXISTS(SELECT 1 FROM movements m WHERE m.id=x.id AND m.inventory_mode='SUPPLIER_BACKED_UNLIMITED'))
      AND x.status!='APPLIED'`)
    .bind(id).first<{ id: string; item_id: string }>();
  if (forbidden) {
    await env.DB.batch([
      env.DB.prepare(`UPDATE outbox SET status='BLOCKED',last_error=?,updated_at=? WHERE id=? AND status!='APPLIED'`)
        .bind('Supplier-backed stock must never be sent to Ecwid.',now(),id),
      issueStatement(env.DB,`supplier-outbox:${id}`,'SUPPLIER_OUTBOX_BLOCKED',forbidden.item_id,null,
        'An invalid stock update was blocked. Supplier-backed items stay unlimited online; reconcile this queued entry.'),
    ]);
    return;
  }
  if (env.ECWID_MODE === 'live' && env.LIVE_SYNC_ENABLED !== 'true') return;
  if (!['demo', 'live'].includes(env.ECWID_MODE)) throw new Error('Unknown Ecwid mode.');
  const cooldown = await getState(env.DB, 'ecwid_rate_limit_until');
  if (cooldown && cooldown > now()) return;
  // Validate configuration before a claim: no request can have happened yet.
  const client = env.ECWID_MODE === 'live' ? ecwidClient(env, fetcher) : null;
  const row = await claimOutbox(env.DB, id);
  if (!row) return;
  if (!client) {
    await env.DB.batch([
      env.DB.prepare(`UPDATE items SET last_ecwid_quantity=coalesce(last_ecwid_quantity,0)+?
        WHERE id=? AND EXISTS(SELECT 1 FROM outbox WHERE id=? AND status='PROCESSING')`)
        .bind(row.quantity_delta, row.item_id, row.id),
      env.DB.prepare(`UPDATE outbox SET status='APPLIED',last_error=NULL,updated_at=? WHERE id=? AND status='PROCESSING'`)
        .bind(now(), row.id),
    ]);
    await enqueueNextOutbox(env, row.item_id);
    return;
  }
  let warning: string | undefined;
  try {
    ({ warning } = await client.adjustStock(row.ecwid_product_id, row.quantity_delta, row.ecwid_combination_id));
  } catch (error) {
    const failure = error instanceof EcwidError ? error : new EcwidError('Stock update outcome is uncertain; reconcile before retry.', 'UNKNOWN');
    if (failure.outcome === 'RETRYABLE') {
      await env.DB.batch([
        env.DB.prepare(`UPDATE outbox SET status='PENDING',last_error=?,updated_at=? WHERE id=? AND status='PROCESSING'`)
          .bind(failure.message, now(), row.id),
        env.DB.prepare(`INSERT INTO sync_state(key,value,updated_at) VALUES ('ecwid_rate_limit_until',?,?)
          ON CONFLICT(key) DO UPDATE SET value=MAX(sync_state.value,excluded.value),updated_at=excluded.updated_at`)
          .bind(new Date(Date.now() + failure.retryAfter * 1000).toISOString(), now()),
      ]);
    } else {
      await failOutbox(env.DB, row, failure.outcome === 'UNKNOWN' ? 'UNKNOWN' : 'BLOCKED', failure.message);
    }
    return;
  }
  // A successful stock write must never be retried just because a subsequent read failed.
  let quantity: number | null = null;
  try { quantity = (await client.getProductStock(row.ecwid_product_id, row.ecwid_combination_id)).quantity; } catch { /* mirror becomes visibly unknown */ }
  const statements = [
    env.DB.prepare(`UPDATE items SET last_ecwid_quantity=? WHERE id=?`).bind(quantity, row.item_id),
    env.DB.prepare(`UPDATE outbox SET status='APPLIED',last_error=?,updated_at=? WHERE id=? AND status='PROCESSING'`)
      .bind(warning ?? null, now(), row.id),
  ];
  if (warning || (quantity !== null && quantity < 0)) statements.push(issueStatement(env.DB,
    `outbox-warning:${row.id}`, 'NEGATIVE_ECWID_STOCK', row.item_id, null, warning ?? 'Ecwid reports negative stock.'));
  await env.DB.batch(statements);
  await enqueueNextOutbox(env, row.item_id);
}

async function markOrderDeleted(db: D1Database, id: string): Promise<void> {
  await db.batch([
    db.prepare('UPDATE orders SET needs_review=1,updated_at=? WHERE id=?').bind(now(), id),
    db.prepare(`INSERT INTO sync_issues(id,kind,item_id,order_id,message,status,created_at)
      SELECT 'deleted-order:' || order_id || ':' || coalesce(item_id,'unmapped'),'ORDER_DELETED',item_id,order_id,
        'Order was deleted in Ecwid. Check reservations and any picked items before continuing.','OPEN',?
      FROM order_lines WHERE order_id=? GROUP BY item_id ON CONFLICT(id) DO NOTHING`).bind(now(), id),
  ]);
}

export function targetMatchesMapping(target: EcwidProduct, item: ProductMapping): boolean {
  const options = canonicalVariationOptions(target.variationOptions);
  const stockPolicyMatches = item.inventory_mode === 'SUPPLIER_BACKED_UNLIMITED'
    ? target.unlimited : !target.unlimited && target.quantity !== null && target.quantity >= 0;
  return target.sku === item.sku && (target.combinationId ?? null) === item.ecwid_combination_id
    && target.enabled && stockPolicyMatches && target.eligibilityVerified === true
    && target.hasBundleRelationships === false && target.hasExtraOptions === false
    && options !== null && JSON.stringify(options) === item.ecwid_option_signature
    && (item.ecwid_combination_id !== null || (!target.hasOptions && !target.hasVariations));
}

async function markProductUnavailable(db: D1Database, productId: string, deleted: boolean): Promise<void> {
  await db.batch([
    db.prepare('UPDATE items SET last_ecwid_quantity=NULL WHERE ecwid_product_id=?').bind(productId),
    db.prepare(`INSERT INTO sync_issues(id,kind,item_id,message,status,created_at)
      SELECT ? || id,?,id,?,'OPEN',? FROM items WHERE ecwid_product_id=? ON CONFLICT(id) DO NOTHING`)
      .bind(deleted ? 'deleted-product:' : 'changed-product:', deleted ? 'PRODUCT_DELETED' : 'PRODUCT_REVIEW',
        deleted ? 'Mapped Ecwid product was deleted; all its mapped stock targets require review.'
          : 'Ecwid product could not be validated. Review every mapped stock target before continuing.', now(), productId),
  ]);
}

async function refreshMappedProduct(env: SyncEnv, productId: string, fetcher?: EcwidFetch): Promise<void> {
  const mapped = await env.DB.prepare(`SELECT id,sku,ecwid_combination_id,ecwid_option_signature,inventory_mode
    FROM items WHERE ecwid_product_id=?`).bind(productId).all<ProductMapping>();
  if (mapped.results.length === 0) return;
  // Fetch the parent once. A missing variation must never fall back to the parent
  // or another size, and a parent event applies to every mapped sibling.
  const targets = await ecwidClient(env, fetcher).getProductStockTargets(productId);
  const statements: D1PreparedStatement[] = [];
  for (const item of mapped.results) {
    const target = targets.find(candidate => (candidate.combinationId ?? null) === item.ecwid_combination_id);
    const matches = target !== undefined && targetMatchesMapping(target, item);
    statements.push(env.DB.prepare('UPDATE items SET last_ecwid_quantity=? WHERE id=?')
      .bind(matches && item.inventory_mode === 'STOCK_LIMITED' ? target.quantity : null, item.id));
    if (!matches) statements.push(issueStatement(env.DB, `changed-product:${item.id}`, 'PRODUCT_REVIEW', item.id, null,
      'Mapped stock target is missing, changed identity/options or finite/unlimited policy, bundled, disabled, or has invalid stock. Reconcile this SKU before continuing.'));
  }
  await env.DB.batch(statements);
}

export async function processWebhook(env: SyncEnv, id: string, fetcher?: EcwidFetch): Promise<void> {
  if (env.ECWID_MODE !== 'live' || env.ORDER_SYNC_ENABLED !== 'true') return;
  const row = await env.DB.prepare(`UPDATE webhook_events SET status='PROCESSING',attempts=attempts+1,updated_at=?
    WHERE event_id=? AND status='PENDING' RETURNING event_id,event_type,entity_id,payload`).bind(now(), id).first<InboxRow>();
  if (!row) return;
  try {
    if (row.event_type === 'order.created' || row.event_type === 'order.updated') {
      await refreshOrder(env, row.entity_id, fetcher);
    } else if (row.event_type === 'order.deleted') {
      await markOrderDeleted(env.DB, row.entity_id);
    } else if (row.event_type === 'product.updated' || row.event_type === 'product.deleted') {
      if (row.event_type === 'product.deleted') await markProductUnavailable(env.DB, row.entity_id, true);
      else await refreshMappedProduct(env, row.entity_id, fetcher);
    }
    await env.DB.prepare(`UPDATE webhook_events SET status='APPLIED',processed_at=?,updated_at=?,last_error=NULL WHERE event_id=?`)
      .bind(now(), now(), id).run();
  } catch (error) {
    if (error instanceof EcwidError && error.status === 404 && row.event_type.startsWith('order.')) {
      await markOrderDeleted(env.DB, row.entity_id);
    }
    const permanent = error instanceof EcwidError && error.outcome === 'REJECTED';
    if (permanent && row.event_type.startsWith('product.')) {
      await markProductUnavailable(env.DB, row.entity_id, error.status === 404);
    }
    await env.DB.prepare('UPDATE webhook_events SET status=?,last_error=?,updated_at=? WHERE event_id=?')
      .bind(permanent ? 'BLOCKED' : 'PENDING', error instanceof Error ? error.message : 'Webhook processing failed.', now(), id).run();
    if (permanent) await issueStatement(env.DB, `webhook:${id}`, 'WEBHOOK_BLOCKED', null, null,
      `Ecwid ${row.event_type} event could not be refreshed. Check integration access and reconcile the affected entity (${row.entity_id}).`).run();
    if (!permanent) throw error;
  }
}

export async function processSyncMessage(env: SyncEnv, message: SyncMessage, fetcher?: EcwidFetch): Promise<void> {
  if (message.kind === 'outbox') await processOutbox(env, message.id, fetcher);
  else if (message.kind === 'webhook') await processWebhook(env, message.id, fetcher);
  else throw new Error('Unknown sync message.');
}

/** A crash after sending stock is ambiguous. Read-only webhook refreshes can safely retry. */
export async function recoverStaleProcessing(db: D1Database, staleBefore = new Date(Date.now() - 10 * 60_000).toISOString()): Promise<void> {
  await db.batch([
    db.prepare(`INSERT INTO sync_issues(id,kind,item_id,message,status,created_at)
      SELECT 'outbox:' || id,'STOCK_UNKNOWN',item_id,'Worker stopped during Ecwid sync. Verify whether the adjustment happened before retrying.','OPEN',?
      FROM outbox WHERE status='PROCESSING' AND updated_at<? ON CONFLICT(id) DO NOTHING`).bind(now(), staleBefore),
    db.prepare(`UPDATE outbox SET status='UNKNOWN',last_error='Worker stopped during stock sync; manual reconciliation required.',updated_at=?
      WHERE status='PROCESSING' AND updated_at<?`).bind(now(), staleBefore),
    db.prepare(`UPDATE webhook_events SET status='PENDING',updated_at=? WHERE status='PROCESSING' AND updated_at<?`).bind(now(), staleBefore),
  ]);
}

export async function pumpSync(env: SyncEnv): Promise<{ queued: number }> {
  await recoverStaleProcessing(env.DB);
  // Keep scheduled polling + queue fan-out within the Free Worker invocation budget.
  const events = await env.DB.prepare(`SELECT event_id AS id FROM webhook_events WHERE status='PENDING' ORDER BY updated_at,event_id LIMIT 5`).all<{ id: string }>();
  const stock = await env.DB.prepare(`SELECT id FROM outbox WHERE status='PENDING'
    AND NOT EXISTS(SELECT 1 FROM sync_issues s WHERE s.item_id=outbox.item_id AND s.status='OPEN')
    AND NOT EXISTS(SELECT 1 FROM outbox other WHERE other.item_id=outbox.item_id AND other.id!=outbox.id
      AND (other.status IN ('PROCESSING','UNKNOWN','BLOCKED') OR (other.status='PENDING'
        AND other.rowid<outbox.rowid)))
    ORDER BY rowid LIMIT 5`).all<{ id: string }>();
  let queued = 0;
  for (const event of events.results) if (await enqueueSync(env, { kind: 'webhook', id: event.id })) queued++;
  for (const row of stock.results) if (await enqueueSync(env, { kind: 'outbox', id: row.id })) queued++;
  return { queued };
}

/** Full pagination with a persisted cursor, including old orders that are still awaiting payment. */
export async function pollHistoricalOrders(env: SyncEnv, fetcher?: EcwidFetch,
  options: { maximumOrders?: number; lease?: string } = {}): Promise<{ processed: number; complete: boolean; demo?: boolean }> {
  if (env.ECWID_MODE !== 'live') return { processed: 0, complete: true, demo: true };
  if (env.ORDER_SYNC_ENABLED !== 'true') throw new DomainError(409, 'ORDER_SYNC_PAUSED', 'Order imports are paused until the reviewed opening cutover is complete.');
  const client = ecwidClient(env, fetcher);
  await env.DB.prepare(`INSERT INTO sync_state(key,value,updated_at) VALUES ('orders_tracking_started',?,?)
    ON CONFLICT(key) DO NOTHING`).bind(now(), now()).run();
  const trackingStarted = await getState(env.DB, 'orders_tracking_started');
  const oldCursor = await getState(env.DB, 'orders_poll_cursor');
  const cursor: { offset: number; createdTo: number } = oldCursor ? JSON.parse(oldCursor)
    : { offset: 0, createdTo: Math.floor(Date.now() / 1000) };
  let processed = 0;
  let complete = false;
  const maximumOrders = options.maximumOrders ?? 3;
  if (!Number.isInteger(maximumOrders) || maximumOrders < 1 || maximumOrders > 3) throw new Error('Invalid historical polling budget.');
  for (let page = 0; page < 2 && processed < maximumOrders; page++) {
    const limit = maximumOrders - processed;
    const response = await client.listOrders({ ...cursor, limit });
    if (response.offset !== cursor.offset || response.count > limit || cursor.offset + response.count > response.total
      || (response.count === 0 && cursor.offset < response.total)) {
      throw new Error('Ecwid order pagination was inconsistent; cursor was retained.');
    }
    for (const order of response.items) {
      // Completed history predates the physical ledger. New completed orders without
      // recorded picks are discrepancies, including orders missed by webhooks.
      await upsertOrderSnapshot(env.DB, order, {
        skipUntrackedTerminal: Boolean(order.createdAt && trackingStarted && order.createdAt < trackingStarted
          && order.updatedAt < trackingStarted),
      });
      processed++;
    }
    cursor.offset += response.count;
    if (cursor.offset >= response.total) { complete = true; break; }
    await savePollState(env.DB, { orders_poll_cursor: JSON.stringify(cursor) }, options.lease);
  }
  if (complete) {
    await savePollState(env.DB, { orders_poll_cursor: '', orders_last_full_poll: now() }, options.lease);
  } else await savePollState(env.DB, { orders_poll_cursor: JSON.stringify(cursor) }, options.lease);
  return { processed, complete };
}

const RECENT_APPLY_LIMIT = 3;
const RECENT_MAX_ORDERS = 1_000;
const RECENT_OVERLAP_SECONDS = 120;
const RECENT_SETTLE_SECONDS = 5;
const POLL_LEASE_MS = 120_000;
interface RecentCursor {
  version: 1;
  updatedFrom: number;
  updatedTo: number;
  phase: 'APPLY' | 'VERIFY';
  offset: number;
  total: number | null;
  seen: Array<{ id: string; signature: string }>;
}

class RecentPollError extends Error {
  constructor(message: string, readonly restart = false, readonly overloaded = false) { super(message); }
}

/** One atomic state write, fenced against expired/replaced pollers. No stale worker can advance a watermark. */
async function savePollState(db: D1Database, values: Record<string, string>, lease?: string): Promise<void> {
  const result = await db.prepare(`INSERT INTO sync_state(key,value,updated_at)
    SELECT j.key,j.value,? FROM json_each(?) j WHERE (? IS NULL OR EXISTS(
      SELECT 1 FROM sync_state l WHERE l.key='orders_poll_lease' AND l.value=? AND json_extract(l.value,'$.expiresAt')>?))
    ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at RETURNING key`)
    .bind(now(), JSON.stringify(values), lease ?? null, lease ?? null, Date.now()).all<{ key: string }>();
  if (result.results.length !== Object.keys(values).length) throw new RecentPollError('Order poll lease expired; progress was retained.');
}

function parseRecentCursor(raw: string, baseline: number, watermark: number | null): RecentCursor {
  let value: unknown;
  try { value = JSON.parse(raw); } catch { throw new RecentPollError('Invalid recent-order checkpoint; administrator review required.'); }
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new RecentPollError('Invalid recent-order checkpoint.');
  const c = value as RecentCursor;
  if (c.version !== 1 || !['APPLY', 'VERIFY'].includes(c.phase) || !Number.isSafeInteger(c.updatedFrom)
    || !Number.isSafeInteger(c.updatedTo) || c.updatedFrom !== Math.max(baseline, (watermark ?? baseline) - RECENT_OVERLAP_SECONDS) || c.updatedFrom > c.updatedTo
    || c.updatedTo > Math.floor(Date.now() / 1000) || (watermark !== null && c.updatedTo < watermark)
    || !Number.isSafeInteger(c.offset) || c.offset < 0 || !Array.isArray(c.seen) || c.seen.length > RECENT_MAX_ORDERS
    || (c.total === null ? c.seen.length !== 0 : !Number.isSafeInteger(c.total) || c.total < c.seen.length || c.total > RECENT_MAX_ORDERS)
    || c.seen.some(entry => !entry || typeof entry.id !== 'string' || !entry.id || entry.id.length > 200
      || typeof entry.signature !== 'string' || !/^[a-f0-9]{64}$/.test(entry.signature))
    || new Set(c.seen.map(entry => entry.id)).size !== c.seen.length
    || (c.phase === 'APPLY' ? c.offset !== c.seen.length : c.total !== c.seen.length || c.offset > c.seen.length)) {
    throw new RecentPollError('Invalid recent-order checkpoint; administrator review required.');
  }
  return c;
}

async function recentOrderSignature(order: EcwidOrder): Promise<string> {
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify([
    order.id, order.createdAt, order.updatedAt, order.paymentStatus, order.fulfillmentStatus, await orderSnapshotHash(order),
  ])));
  return Array.from(new Uint8Array(bytes), byte => byte.toString(16).padStart(2, '0')).join('');
}

/** Recent updates are independent of order creation date and status. Historical scans are only a fallback.
 * At most 3 snapshot transactions + bounded state queries per call, leaving room for pumpSync on Free D1.
 * Ecwid offset pages are not snapshots: a second matching pass is required before advancing the watermark.
 */
export async function pollOrders(env: SyncEnv, fetcher?: EcwidFetch): Promise<{
  processed: number; complete: boolean; demo?: boolean; busy?: boolean; historical_processed?: number; historical_error?: boolean;
}> {
  if (env.ECWID_MODE !== 'live') return { processed: 0, complete: true, demo: true };
  if (env.ORDER_SYNC_ENABLED !== 'true') throw new DomainError(409, 'ORDER_SYNC_PAUSED', 'Order imports are paused until the reviewed opening cutover is complete.');
  const lease = JSON.stringify({ token: crypto.randomUUID(), expiresAt: Date.now() + POLL_LEASE_MS });
  const claimed = await env.DB.prepare(`INSERT INTO sync_state(key,value,updated_at) VALUES('orders_poll_lease',?,?)
    ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at
    WHERE json_extract(sync_state.value,'$.expiresAt')<=? RETURNING value`).bind(lease, now(), Date.now()).first<string>('value');
  if (claimed !== lease) return { processed: 0, complete: false, busy: true };
  let cursor: RecentCursor | undefined;
  let processed = 0;
  try {
    const rows = await env.DB.prepare(`SELECT key,value FROM sync_state WHERE key IN
      ('orders_tracking_started','orders_recent_watermark','orders_recent_poll_cursor')`).all<{ key: string; value: string }>();
    const state = Object.fromEntries(rows.results.map(row => [row.key, row.value]));
    const baseline = Math.floor(Date.parse(state.orders_tracking_started ?? '') / 1000);
    const watermark = state.orders_recent_watermark ? Math.floor(Date.parse(state.orders_recent_watermark) / 1000) : null;
    if (!Number.isSafeInteger(baseline) || baseline < 0 || baseline > Math.floor(Date.now() / 1000)
      || (watermark !== null && (!Number.isSafeInteger(watermark) || watermark < baseline || watermark > Math.floor(Date.now() / 1000)))) {
      throw new RecentPollError('A valid opening-cutover tracking baseline is required before recent order polling.');
    }
    cursor = state.orders_recent_poll_cursor ? parseRecentCursor(state.orders_recent_poll_cursor, baseline, watermark) : {
      version: 1, updatedFrom: Math.max(baseline, (watermark ?? baseline) - RECENT_OVERLAP_SECONDS),
      updatedTo: Math.floor(Date.now() / 1000) - RECENT_SETTLE_SECONDS, phase: 'APPLY', offset: 0, total: null, seen: [],
    };
    if (cursor.updatedTo < cursor.updatedFrom) {
      await savePollState(env.DB, { orders_recent_status: 'PENDING', orders_recent_error: '' }, lease);
      return { processed, complete: false };
    }
    await savePollState(env.DB, { orders_recent_poll_cursor: JSON.stringify(cursor), orders_recent_status: 'PENDING', orders_recent_error: '' }, lease);
    const client = ecwidClient(env, fetcher);
    // An APPLY page may immediately start VERIFY, but neither phase loops unboundedly.
    for (let pass = 0; pass < 2; pass++) {
      const phase = cursor.phase;
      const limit = phase === 'APPLY' ? RECENT_APPLY_LIMIT : 100;
      const page = await client.listOrders({ updatedFrom: cursor.updatedFrom, updatedTo: cursor.updatedTo, offset: cursor.offset, limit });
      if (page.total > RECENT_MAX_ORDERS) throw new RecentPollError('Recent-order window exceeds 1000 orders; administrator review required.', false, true);
      if (page.offset !== cursor.offset || page.count > limit || page.count !== page.items.length
        || cursor.offset + page.count > page.total || (!page.count && cursor.offset < page.total)
        || (cursor.total !== null && cursor.total !== page.total)) {
        throw new RecentPollError('Ecwid recent-order pages changed; the same window will restart.', true);
      }
      cursor.total = page.total;
      const entries = await Promise.all(page.items.map(async order => {
        const updated = Date.parse(order.updatedAt) / 1000;
        if (!order.id || order.id.length > 200 || updated < cursor!.updatedFrom || updated > cursor!.updatedTo || !Number.isFinite(updated)) {
          throw new RecentPollError('Ecwid returned an order outside the requested update window.', true);
        }
        return { id: order.id, signature: await recentOrderSignature(order) };
      }));
      if (new Set(entries.map(entry => entry.id)).size !== entries.length) throw new RecentPollError('Ecwid returned duplicate order IDs; the window will restart.', true);
      if (phase === 'APPLY') {
        if (entries.some(entry => cursor!.seen.some(seen => seen.id === entry.id))) throw new RecentPollError('Ecwid recent-order pages overlapped; the window will restart.', true);
        for (const order of page.items) { await upsertOrderSnapshot(env.DB, order); processed++; }
        cursor.seen.push(...entries);
        cursor.offset += page.count;
        if (cursor.offset === page.total) { cursor.phase = 'VERIFY'; cursor.offset = 0; }
        else break;
      } else {
        if (entries.some((entry, index) => entry.id !== cursor!.seen[cursor!.offset + index]?.id
          || entry.signature !== cursor!.seen[cursor!.offset + index]?.signature)) {
          throw new RecentPollError('Ecwid recent-order verification changed; the same window will restart.', true);
        }
        cursor.offset += page.count;
        if (cursor.offset === page.total) {
          await savePollState(env.DB, { orders_recent_poll_cursor: '', orders_recent_watermark: new Date(cursor.updatedTo * 1000).toISOString(),
            orders_recent_last_success: now(), orders_recent_status: 'CURRENT', orders_recent_error: '' }, lease);
          // Never spend a second snapshot budget when this invocation already applied recent orders.
          if (processed === 0) {
            try {
              const historical = await pollHistoricalOrders(env, fetcher, { maximumOrders: 2, lease });
              return { processed, complete: true, historical_processed: historical.processed };
            } catch {
              // Reconciliation failure does not invalidate the independently verified recent window.
              return { processed, complete: true, historical_error: true };
            }
          }
          return { processed, complete: true };
        }
        break;
      }
    }
    await savePollState(env.DB, { orders_recent_poll_cursor: JSON.stringify(cursor), orders_recent_status: 'PENDING', orders_recent_error: '' }, lease);
    return { processed, complete: false };
  } catch (error) {
    const values: Record<string, string> = { orders_recent_status: error instanceof RecentPollError && error.overloaded ? 'OVERLOADED' : 'ERROR',
      orders_recent_error: error instanceof RecentPollError ? error.message : 'Recent order polling failed; its watermark has not advanced. Retry or check integration access.' };
    if (cursor && error instanceof RecentPollError && error.restart) {
      values.orders_recent_poll_cursor = JSON.stringify({ ...cursor, phase: 'APPLY', offset: 0, total: null, seen: [] });
    }
    await savePollState(env.DB, values, lease);
    throw error;
  } finally {
    await env.DB.prepare("DELETE FROM sync_state WHERE key='orders_poll_lease' AND value=?").bind(lease).run();
  }
}
