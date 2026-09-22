import { authenticate, requireAdmin } from './auth';
import { DomainError, validateMovementInput } from './domain';
import { createMovement, dashboard, getItem, getOrder, listItems, listMovements, listOrders } from './inventory';
import { json, objectBody, readJson, requireSameOrigin } from './http';
import { previewImport } from './opening-import';
import { stageOpeningImport } from './opening-apply';
import { previewOpeningCutover, stageOpeningCutover } from './opening-cutover';
import { EcwidError } from './ecwid';
import { createAllocation, listAllocations, validateAllocationInput } from './supplier';
import { enqueueSync, ingestWebhook, pollOrders, processSyncMessage, pumpSync, refreshOrder, type SyncMessage } from './sync';

function itemCode(raw: string): { sku: string; combinationId: string | null } {
  const parts = raw.trim().split('|');
  if (parts.length > 3 || (parts[1]?.trim() && !/^[1-9]\d{0,30}$/.test(parts[1].trim()))) {
    throw new DomainError(400, 'INVALID_VARIATION_QR', 'Use the item’s unique SKU, or a QR containing its exact numeric Ecwid variation ID. Legacy variation labels cannot identify stock safely.');
  }
  if (!parts[0]?.trim() || parts[0].length > 200) throw new DomainError(400, 'INVALID_CODE', 'Enter or scan an item SKU.');
  return { sku: parts[0].trim(), combinationId: parts[1]?.trim() || null };
}

async function routes(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
  const url = new URL(request.url);
  const path = url.pathname;
  if (path !== '/api' && !path.startsWith('/api/')) return env.ASSETS.fetch(request);
  if (path === '/api/health' && request.method === 'GET') return json({ ok: true, version: '0.1.0' });
  if (path === '/api/webhooks/ecwid' && request.method === 'POST') return ingestWebhook(request, env);

  const identity = await authenticate(request, env);
  if (!['GET', 'HEAD'].includes(request.method)) requireSameOrigin(request);
  if (path === '/api/session' && request.method === 'GET') {
    return json({ ...identity, mode: env.ECWID_MODE, live_sync_enabled: env.LIVE_SYNC_ENABLED === 'true',
      order_sync_enabled: env.ECWID_MODE === 'demo' || env.ORDER_SYNC_ENABLED === 'true',
      inventory_enabled: env.ECWID_MODE === 'demo' || env.INVENTORY_ENABLED === 'true' });
  }
  if (path === '/api/dashboard' && request.method === 'GET') return json(await dashboard(env.DB));
  if (path === '/api/items/lookup' && request.method === 'GET') {
    const code = itemCode(url.searchParams.get('code') ?? '');
    const item = await getItem(env.DB, code.sku);
    if (code.combinationId && item.ecwid_combination_id !== code.combinationId) {
      throw new DomainError(409, 'VARIATION_QR_MISMATCH', 'The QR variation ID does not match this SKU. Check the box label before moving stock.');
    }
    return json({ item });
  }
  if (path === '/api/items' && request.method === 'GET') return json({ items: await listItems(env.DB, url.searchParams.get('search') ?? '') });
  if (path === '/api/orders' && request.method === 'GET') {
    const filters: Record<string, string> = { PAID: 'pickable', AWAITING_PAYMENT: 'awaiting_payment', ALL: 'all', REVIEW: 'review' };
    const status = url.searchParams.get('status') ?? 'PAID';
    if (!filters[status]) throw new DomainError(400, 'INVALID_STATUS', 'Choose Paid, Awaiting Payment, All, or Review.');
    return json({ orders: await listOrders(env.DB, filters[status]) });
  }
  const orderMatch = path.match(/^\/api\/orders\/([^/]+)(\/refresh)?$/);
  if (orderMatch && (request.method === 'GET' || (request.method === 'POST' && orderMatch[2]))) {
    let id: string;
    try { id = decodeURIComponent(orderMatch[1]); }
    catch { throw new DomainError(400, 'INVALID_ORDER_ID', 'Enter a valid Ecwid order ID.'); }
    if (!/^[A-Za-z0-9_-]{1,100}$/.test(id)) throw new DomainError(400, 'INVALID_ORDER_ID', 'Enter a valid Ecwid order ID.');
    if (env.ECWID_MODE === 'live' && (env.ORDER_SYNC_ENABLED === 'true' || request.method === 'POST')) await refreshOrder(env, id);
    return json({ order: await getOrder(env.DB, id) });
  }
  if (path === '/api/allocations' && request.method === 'GET') return json({ allocations: await listAllocations(env.DB) });
  if (path === '/api/allocations' && request.method === 'POST') {
    const input = validateAllocationInput(await readJson(request));
    // A lost response must resolve using the original event even if an order
    // was cancelled, or recording was disabled, after its allocation committed.
    const existing = await env.DB.prepare('SELECT id FROM supplier_allocation_events WHERE id=?').bind(input.operation_id).first();
    if (!existing) {
      if (env.ECWID_MODE === 'live' && env.INVENTORY_ENABLED !== 'true') {
        throw new DomainError(409, 'CUTOVER_REQUIRED', 'Opening stock alignment must be completed before recording live allocations.');
      }
      if (env.ECWID_MODE === 'live') await refreshOrder(env, input.order_id);
    }
    const result = await createAllocation(env.DB, input, identity.actor);
    return json(result, result.duplicate ? 200 : 201);
  }
  if (path === '/api/movements' && request.method === 'GET') return json({ movements: await listMovements(env.DB) });
  if (path === '/api/movements' && request.method === 'POST') {
    const input = validateMovementInput(await readJson(request));
    // Replays must still resolve after the original order changes state, or a
    // lost response could incorrectly look like an unsuccessful stock movement.
    const existing = await env.DB.prepare('SELECT id FROM movements WHERE id=?').bind(input.operation_id).first();
    if (!existing) {
      if (env.ECWID_MODE === 'live' && env.INVENTORY_ENABLED !== 'true') {
        throw new DomainError(409, 'CUTOVER_REQUIRED', 'Opening stock alignment must be completed before recording live movements.');
      }
      if (input.type === 'ECWID_PICK' && env.ECWID_MODE === 'live') await refreshOrder(env, input.order_id!);
    }
    const result = await createMovement(env.DB, input, identity.actor);
    if (result.sync_status === 'PENDING') ctx.waitUntil(enqueueSync(env, { kind: 'outbox', id: result.movement.id }));
    return json(result, result.duplicate ? 200 : 201);
  }
  if (path === '/api/sync' && request.method === 'GET') {
    const [outbox, issues, inboxCounts, syncState] = await env.DB.batch([
      env.DB.prepare(`SELECT o.*,i.sku,i.name FROM outbox o JOIN items i ON i.id=o.item_id
        ORDER BY o.created_at DESC LIMIT 100`),
      env.DB.prepare("SELECT * FROM sync_issues WHERE status='OPEN' ORDER BY created_at DESC LIMIT 100"),
      env.DB.prepare('SELECT status,COUNT(*) AS count FROM webhook_events GROUP BY status'),
      env.DB.prepare("SELECT key,value,updated_at FROM sync_state WHERE key IN ('orders_last_full_poll','orders_poll_cursor','orders_tracking_started')")
    ]);
    return json({ outbox: outbox.results, issues: issues.results, inbox_counts: inboxCounts.results, sync_state: syncState.results });
  }
  if (path === '/api/sync/pump' && request.method === 'POST') {
    requireAdmin(identity);
    await pumpSync(env);
    return json({ queued: true });
  }
  if (path === '/api/orders/poll' && request.method === 'POST') {
    requireAdmin(identity);
    await pollOrders(env);
    return json({ polled: true });
  }
  if (path === '/api/import/preview' && request.method === 'POST') {
    requireAdmin(identity);
    // Full read-only catalogue envelopes include every variation. Stock rows
    // still have their own limits; keep the entire request bounded as well.
    const input = objectBody(await readJson(request, 10_000_000));
    if (env.ECWID_STORE_ID) {
      if (input.store_id !== undefined && input.store_id !== env.ECWID_STORE_ID) {
        throw new DomainError(400, 'CATALOGUE_STORE_MISMATCH', 'The preview store must match the configured Ecwid store.');
      }
      input.store_id = env.ECWID_STORE_ID;
    }
    return json(await previewImport(input));
  }
  if (path === '/api/import/stage' && request.method === 'POST') {
    requireAdmin(identity);
    if (!env.ECWID_STORE_ID) throw new DomainError(409, 'IMPORT_STORE_REQUIRED', 'Configure the exact Ecwid store before staging opening stock.');
    if (env.INVENTORY_ENABLED !== 'false' || env.LIVE_SYNC_ENABLED !== 'false' || env.ORDER_SYNC_ENABLED !== 'false') {
      throw new DomainError(409, 'IMPORT_REQUIRES_DISABLED_LIVE_FLAGS', 'Disable inventory recording and live sync before staging opening stock.');
    }
    const result = await stageOpeningImport(env.DB, await readJson(request, 10_000_000), { storeId: env.ECWID_STORE_ID, actor: identity.actor });
    return json(result, result.duplicate ? 200 : 201);
  }
  if (['/api/cutover/preview', '/api/cutover/stage'].includes(path) && request.method === 'POST') {
    requireAdmin(identity);
    if (!env.ECWID_STORE_ID) throw new DomainError(409, 'IMPORT_STORE_REQUIRED', 'Configure the exact Ecwid store before preparing opening stock.');
    if (env.INVENTORY_ENABLED !== 'false' || env.LIVE_SYNC_ENABLED !== 'false' || env.ORDER_SYNC_ENABLED !== 'false') {
      throw new DomainError(409, 'CUTOVER_REQUIRES_DISABLED_FLAGS', 'Disable inventory recording, stock sync and order sync before preparing opening stock.');
    }
    const input = await readJson(request, 10_000_000);
    const policy = { storeId: env.ECWID_STORE_ID, actor: identity.actor };
    if (path.endsWith('/preview')) return json(await previewOpeningCutover(input, policy));
    const result = await stageOpeningCutover(env.DB, input, policy);
    return json(result, result.duplicate ? 200 : 201);
  }
  throw new DomainError(404, 'NOT_FOUND', 'This endpoint does not exist.');
}

export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    try { return await routes(request, env, ctx); }
    catch (error) {
      if (error instanceof DomainError) return json({ error: error.message, code: error.code }, error.status);
      if (error instanceof EcwidError) return json({ error: 'Ecwid could not confirm the latest store information. Please try again.', code: 'ECWID_UNAVAILABLE' }, 503);
      console.error(JSON.stringify({ event: 'request_failed', path: new URL(request.url).pathname,
        message: error instanceof Error ? error.message : 'Unknown error' }));
      return json({ error: 'The request could not be completed. If you submitted stock, retry the same operation.', code: 'INTERNAL_ERROR' }, 500);
    }
  },
  async queue(batch: MessageBatch<SyncMessage>, env: Env): Promise<void> {
    for (const message of batch.messages) {
      try {
        await processSyncMessage(env, message.body);
        message.ack();
      } catch {
        console.error(JSON.stringify({ event: 'sync_delivery_failed', id: message.id }));
        message.retry({ delaySeconds: 60 });
      }
    }
  },
  async scheduled(_controller: ScheduledController, env: Env): Promise<void> {
    // A polling failure must not prevent persisted stock work from being queued.
    const result = await Promise.allSettled([pollOrders(env), pumpSync(env)]);
    for (const task of result) if (task.status === 'rejected') {
      console.error(JSON.stringify({ event: 'scheduled_sync_failed', message: task.reason instanceof Error ? task.reason.message : 'Unknown error' }));
    }
  }
} satisfies ExportedHandler<Env, SyncMessage>;
