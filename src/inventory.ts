import { DomainError, movementFingerprint, normalizeCode, validateMovementInput } from './domain';
import type { Item, Movement, Order, OrderLine } from './domain';

const movementSelect = `SELECT m.*, i.sku, i.name, COALESCE(o.status, 'NOT_REQUIRED') AS sync_status
  FROM movements m JOIN items i ON i.id = m.item_id LEFT JOIN outbox o ON o.id = m.id`;
const itemSelect = `SELECT s.*, s.supplier_paid_demand AS paid_demand,
  s.supplier_awaiting_payment_demand AS awaiting_payment_demand FROM item_stock s`;
const pickableOrder = `o.payment_status = 'PAID' AND o.needs_review = 0
  AND o.fulfillment_status IN ('AWAITING_PROCESSING', 'PROCESSING')
  AND EXISTS (SELECT 1 FROM order_lines l WHERE l.order_id = o.id AND l.management_mode='APP' AND l.ordered_qty > l.picked_qty)`;

export async function getItem(db: D1Database, code: string): Promise<Item> {
  const item = await db.prepare(`${itemSelect} WHERE id = ? OR sku = ? OR scan_code = ? LIMIT 1`)
    .bind(code.trim(), normalizeCode(code), normalizeCode(code)).first<Item>();
  if (!item) throw new DomainError(404, 'ITEM_NOT_FOUND', 'No item matches that code.');
  return item;
}

export async function listItems(db: D1Database, search = ''): Promise<Item[]> {
  const query = `%${search.trim().replace(/[\\%_]/g, '\\$&')}%`;
  const result = await db.prepare(`${itemSelect}
    WHERE sku LIKE ? ESCAPE '\\' OR name LIKE ? ESCAPE '\\' OR scan_code LIKE ? ESCAPE '\\' OR location LIKE ? ESCAPE '\\'
    ORDER BY name, sku LIMIT 200`).bind(query, query, query, query).all<Item>();
  return result.results;
}

async function attachLines(db: D1Database, rows: Omit<Order, 'lines'>[]): Promise<Order[]> {
  if (!rows.length) return [];
  const result = await db.prepare(`SELECT l.*, l.ordered_qty - l.picked_qty AS remaining_qty,
      COALESCE(i.ecwid_combination_id,w.ecwid_combination_id) AS ecwid_combination_id,
      COALESCE(i.ecwid_option_signature,w.ecwid_option_signature) AS ecwid_option_signature, i.inventory_mode, i.on_hand,
      i.opening_verified,i.active AS item_active,
      COALESCE(a.allocated_qty,0) AS allocated_qty, i.free AS free_qty,
      CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED'
        THEN MAX(0,l.ordered_qty-l.picked_qty-COALESCE(a.allocated_qty,0)) ELSE 0 END AS unallocated_qty,
      CASE WHEN o.needs_review=0 AND o.payment_status='PAID'
        AND o.fulfillment_status IN ('AWAITING_PROCESSING','PROCESSING') AND i.active=1
        AND NOT EXISTS(SELECT 1 FROM sync_issues s WHERE s.status='OPEN' AND (s.item_id=i.id OR s.order_id=o.id))
        AND NOT EXISTS(SELECT 1 FROM outbox x WHERE x.item_id=i.id AND x.status IN ('UNKNOWN','BLOCKED'))
        THEN MAX(0,MIN(l.ordered_qty-l.picked_qty,i.on_hand,
          CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(a.allocated_qty,0)
            ELSE l.ordered_qty-l.picked_qty END)) ELSE 0 END AS pickable_qty,
      CASE WHEN i.active IS NOT 1
        OR EXISTS(SELECT 1 FROM sync_issues s WHERE s.status='OPEN' AND (s.item_id=i.id OR s.order_id=o.id))
        OR EXISTS(SELECT 1 FROM outbox x WHERE x.item_id=i.id AND x.status IN ('UNKNOWN','BLOCKED'))
        THEN 1 ELSE 0 END AS line_needs_review,
      'REVIEW' AS fulfillment_state
    FROM order_lines l JOIN orders o ON o.id=l.order_id
      LEFT JOIN item_stock i ON i.id=l.item_id
      LEFT JOIN workbook_managed_targets w ON w.id=l.workbook_target_id
      LEFT JOIN supplier_line_allocations a ON a.order_line_id=l.id
    WHERE l.order_id IN (${rows.map(() => '?').join(',')}) ORDER BY l.order_id, l.name, l.id`)
    .bind(...rows.map(row => row.id)).all<OrderLine & { line_needs_review: number }>();
  const lines = new Map<string, OrderLine[]>();
  const orders = new Map(rows.map(row => [row.id, row]));
  for (const line of result.results) {
    const order = orders.get(line.order_id)!;
    if (line.management_mode === 'WORKBOOK') line.fulfillment_state = 'WORKBOOK_MANAGED';
    else if (order.needs_review || line.line_needs_review || !line.item_id || line.inventory_mode === null) line.fulfillment_state = 'REVIEW';
    else if (line.remaining_qty === 0) line.fulfillment_state = 'PICKED';
    else if (!['PAID', 'AWAITING_PAYMENT'].includes(order.payment_status)
      || !['AWAITING_PROCESSING', 'PROCESSING'].includes(order.fulfillment_status)) line.fulfillment_state = 'CLOSED';
    else if (order.payment_status === 'AWAITING_PAYMENT') line.fulfillment_state = 'AWAITING_PAYMENT';
    else if (line.pickable_qty > 0) line.fulfillment_state = 'READY_TO_PICK';
    else if (line.inventory_mode === 'SUPPLIER_BACKED_UNLIMITED') {
      // Free stock is not assigned to every competing line. This state invites
      // an explicit allocation; it does not promise those units to the order.
      line.fulfillment_state = line.unallocated_qty > 0
        ? ((line.free_qty ?? 0) > 0 ? 'AWAITING_ASSIGNMENT' : 'AWAITING_SUPPLIER') : 'REVIEW';
    } else line.fulfillment_state = (line.on_hand ?? 0) > 0 ? 'REVIEW' : 'AWAITING_STOCK';
    const existing = lines.get(line.order_id) ?? [];
    existing.push(line);
    lines.set(line.order_id, existing);
  }
  return rows.map(row => ({ ...row, lines: lines.get(row.id) ?? [] }));
}

export async function listOrders(db: D1Database, filter = 'pickable'): Promise<Order[]> {
  if (!['pickable', 'all', 'awaiting_payment', 'review'].includes(filter)) {
    throw new DomainError(400, 'INVALID_ORDER_FILTER', 'Choose a supported order filter.');
  }
  const where = filter === 'pickable' ? pickableOrder
    : filter === 'awaiting_payment' ? "o.payment_status = 'AWAITING_PAYMENT'"
      : filter === 'review' ? 'o.needs_review = 1' : '1 = 1';
  const result = await db.prepare(`SELECT o.* FROM orders o WHERE ${where} ORDER BY o.updated_at DESC LIMIT 100`)
    .all<Omit<Order, 'lines'>>();
  return attachLines(db, result.results);
}

export async function getOrder(db: D1Database, id: string): Promise<Order> {
  const row = await db.prepare('SELECT * FROM orders WHERE id = ?').bind(id.trim()).first<Omit<Order, 'lines'>>();
  if (!row) throw new DomainError(404, 'ORDER_NOT_FOUND', 'That order has not been imported.');
  return (await attachLines(db, [row]))[0];
}

export async function listMovements(db: D1Database): Promise<Movement[]> {
  return (await db.prepare(`${movementSelect} ORDER BY m.created_at DESC, m.id DESC LIMIT 100`).all<Movement>()).results;
}

const guardErrors: Record<string, string> = {
  IDEMPOTENCY_CONFLICT: 'This operation ID was already used for a different action.',
  ITEM_UNAVAILABLE: 'This item is inactive or no longer available.',
  ITEM_NEEDS_REVIEW: 'This item has a stock sync issue that needs review before another movement.',
  STOCK_SYNC_PENDING: 'The previous stock change for this item is still syncing to Ecwid. Wait a moment, then fetch the item again.',
  ECWID_MAPPING_REQUIRED: 'This item needs a verified Ecwid product mapping.',
  ORDER_NOT_PICKABLE: 'Only paid, eligible orders without review flags can be picked. Check the selected item and order line.',
  PICK_QUANTITY_EXCEEDED: 'That quantity exceeds the quantity left to pick on this order line.',
  INSUFFICIENT_PHYSICAL_STOCK: 'There is not enough physical stock for this movement.',
  INSUFFICIENT_AVAILABLE_STOCK: 'There is not enough unreserved stock. Existing orders have reserved some of this item.',
  INSUFFICIENT_ALLOCATION: 'Assign received stock to this order line before picking that quantity.',
  INSUFFICIENT_LINE_ALLOCATION: 'Assign received stock to this order line before picking that quantity.',
  SUPPLIER_OPENING_REQUIRED: 'A verified physical opening count is required before supplier stock can move.',
  INVENTORY_MODE_MISMATCH: 'The inventory mode changed. Review the item before recording this movement.',
  WORKBOOK_LINE_HANDLED_EXTERNALLY: 'This order line remains workbook-managed. Do not record its stock in this app.',
};

export async function createMovement(db: D1Database, value: unknown, actor: string): Promise<{ movement: Movement; duplicate: boolean; sync_status: Movement['sync_status'] }> {
  const input = validateMovementInput(value);
  if (!actor?.trim()) throw new DomainError(401, 'ACTOR_REQUIRED', 'Sign in before recording stock movements.');
  const normalizedActor = actor.trim().toLowerCase();
  const fingerprint = await movementFingerprint(input, normalizedActor);
  const quantityDelta = input.type === 'RESTOCK' ? input.quantity : -input.quantity;
  const now = new Date().toISOString();
  // Existing UUIDs never enter the trigger. The database batch also gives the
  // caller a consistent result when two submissions arrive simultaneously.
  const insert = db.prepare(`INSERT INTO movements
    (id, fingerprint, type, item_id, quantity, quantity_delta, ecwid_quantity_delta, order_id, order_line_id, note, reason_code, actor, created_at,inventory_mode)
    SELECT ?, ?, ?, ?, ?, ?, CASE WHEN ?='ECWID_PICK' OR
      (SELECT inventory_mode FROM items WHERE id=?)='SUPPLIER_BACKED_UNLIMITED' THEN 0 ELSE ? END,
      ?, ?, ?, ?, ?, ?, COALESCE((SELECT inventory_mode FROM items WHERE id=?),'STOCK_LIMITED')
    WHERE NOT EXISTS (SELECT 1 FROM movements WHERE id = ?)`)
    .bind(input.operation_id, fingerprint, input.type, input.item_id, input.quantity, quantityDelta, input.type,input.item_id,quantityDelta,
      input.order_id ?? null, input.order_line_id ?? null, input.note ?? '', input.reason_code ?? '', normalizedActor, now,
      input.item_id,input.operation_id);
  let result: D1Result<Movement>[];
  try {
    result = await db.batch<Movement>([insert, db.prepare(`${movementSelect} WHERE m.id = ?`).bind(input.operation_id)]);
  } catch (error) {
    const message = error instanceof Error ? `${error.message} ${error.cause ?? ''}` : String(error);
    for (const [code, explanation] of Object.entries(guardErrors)) {
      if (message.includes(code)) throw new DomainError(409, code, explanation);
    }
    throw error;
  }
  const movement = result[1].results[0];
  if (!movement) throw new Error('Movement was not returned after recording.');
  if (movement.fingerprint !== fingerprint) {
    throw new DomainError(409, 'IDEMPOTENCY_CONFLICT', 'This operation ID was already used with different details.');
  }
  return { movement, duplicate: result[0].meta.changes === 0, sync_status: movement.sync_status };
}

export async function dashboard(db: D1Database) {
  const results = await db.batch<Record<string, unknown>>([
    db.prepare(`SELECT COUNT(*) AS items_count, COALESCE(SUM(on_hand), 0) AS physical_units,
      COALESCE(SUM(reserved), 0) AS reserved_units, COALESCE(SUM(available), 0) AS available_units,
      COALESCE(SUM(CASE WHEN inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN 1 ELSE 0 END),0) AS supplier_items,
      COALESCE(SUM(allocated),0) AS allocated_units,
      COALESCE(SUM(supplier_uncovered_demand),0) AS uncovered_supplier_demand FROM item_stock WHERE active = 1`),
    db.prepare(`SELECT COUNT(*) AS pickable_orders FROM orders o WHERE ${pickableOrder}`),
    db.prepare("SELECT COUNT(*) AS pending_sync FROM outbox WHERE status IN ('PENDING', 'PROCESSING')"),
    db.prepare("SELECT COUNT(*) AS attention_count FROM sync_issues WHERE status = 'OPEN'"),
    db.prepare("SELECT * FROM sync_issues WHERE status = 'OPEN' ORDER BY created_at DESC LIMIT 50"),
    db.prepare(`SELECT o.*, i.sku, i.name FROM outbox o JOIN items i ON i.id = o.item_id
      WHERE o.status <> 'APPLIED' ORDER BY o.created_at DESC LIMIT 50`),
  ]);
  return {
    ...results[0].results[0], ...results[1].results[0], ...results[2].results[0], ...results[3].results[0],
    issues: results[4].results,
    outbox: results[5].results,
  };
}
