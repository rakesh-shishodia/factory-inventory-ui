import { DomainError } from './domain';

export interface AllocationInput {
  operation_id: string;
  type: 'ALLOCATE' | 'RELEASE';
  item_id: string;
  order_id: string;
  order_line_id: string;
  quantity: number;
  note: string;
}

export interface AllocationEvent {
  id: string;
  fingerprint: string;
  type: 'ALLOCATE' | 'RELEASE' | 'PICK' | 'CANCEL_RELEASE';
  item_id: string;
  order_id: string;
  order_line_id: string;
  quantity: number;
  quantity_delta: number;
  movement_id: string | null;
  note: string;
  actor: string;
  created_at: string;
  sku: string;
  name: string;
}

const allocationSelect = `SELECT a.*,i.sku,i.name FROM supplier_allocation_events a
  JOIN items i ON i.id=a.item_id`;

export function validateAllocationInput(value: unknown): AllocationInput {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new DomainError(400, 'INVALID_INPUT', 'An allocation object is required.');
  }
  const input = value as Record<string, unknown>;
  if (typeof input.operation_id !== 'string'
    || !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(input.operation_id)) {
    throw new DomainError(400, 'INVALID_OPERATION_ID', 'A UUID operation ID is required.');
  }
  if (input.type !== 'ALLOCATE' && input.type !== 'RELEASE') {
    throw new DomainError(400, 'INVALID_ALLOCATION_TYPE', 'Choose allocation or release.');
  }
  for (const key of ['item_id', 'order_id', 'order_line_id']) {
    const field = input[key];
    if (typeof field !== 'string' || !field.trim() || field.length > 200) {
      throw new DomainError(400, 'ALLOCATION_TARGET_REQUIRED', 'An exact item, order and order line are required.');
    }
  }
  if (typeof input.quantity !== 'number' || !Number.isInteger(input.quantity) || input.quantity < 1 || input.quantity > 1_000_000) {
    throw new DomainError(400, 'INVALID_QUANTITY', 'Quantity must be a whole number between 1 and 1,000,000.');
  }
  if (input.note !== undefined && (typeof input.note !== 'string' || input.note.length > 1000)) {
    throw new DomainError(400, 'INVALID_NOTE', 'The note must contain at most 1,000 characters.');
  }
  return { operation_id: input.operation_id.toLowerCase(), type: input.type,
    item_id: (input.item_id as string).trim(), order_id: (input.order_id as string).trim(),
    order_line_id: (input.order_line_id as string).trim(), quantity: input.quantity,
    note: ((input.note as string | undefined) ?? '').trim() };
}

export async function allocationFingerprint(input: AllocationInput, actor: string): Promise<string> {
  const canonical = JSON.stringify(['supplier-allocation-v1', input.type, input.item_id, input.order_id,
    input.order_line_id, input.quantity, input.note, actor]);
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(canonical));
  return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
}

const guardErrors: Record<string, string> = {
  IDEMPOTENCY_CONFLICT: 'This operation ID was already used for a different action.',
  ITEM_UNAVAILABLE: 'This item is inactive or no longer available.',
  ITEM_NEEDS_REVIEW: 'This item requires review before stock can be assigned or released.',
  SUPPLIER_ITEM_REQUIRED: 'Only an explicitly mapped supplier-backed item can use allocations.',
  SUPPLIER_ITEM_UNAVAILABLE: 'A verified, active supplier-backed item is required for this allocation.',
  ALLOCATION_LINE_MISMATCH: 'Choose the exact order line mapped to this supplier item.',
  SUPPLIER_OPENING_REQUIRED: 'A verified physical opening count is required before assigning stock.',
  ECWID_MAPPING_REQUIRED: 'This item needs a verified Ecwid product mapping.',
  ORDER_NOT_ALLOCATABLE: 'The order and exact line must be open, mapped and free of review issues.',
  INSUFFICIENT_FREE_STOCK: 'There is not enough free physical stock to assign this quantity.',
  INSUFFICIENT_FREE_SUPPLIER_STOCK: 'There is not enough free physical stock to assign this quantity.',
  ALLOCATION_QUANTITY_EXCEEDED: 'This quantity exceeds the unassigned demand on this order line.',
  INSUFFICIENT_ALLOCATION: 'This quantity exceeds the stock currently assigned to this order line.',
  INSUFFICIENT_LINE_ALLOCATION: 'This quantity exceeds the stock currently assigned to this order line.',
};

/** One insert with SQL guards is the allocation transaction. Receipts and picking
 * use their own audited operation IDs; this service never calls Ecwid or queues stock. */
export async function createAllocation(db: D1Database, value: unknown, actor: string): Promise<{ allocation: AllocationEvent; duplicate: boolean }> {
  const input = validateAllocationInput(value);
  if (!actor?.trim()) throw new DomainError(401, 'ACTOR_REQUIRED', 'Sign in before assigning stock.');
  const normalizedActor = actor.trim().toLowerCase();
  const fingerprint = await allocationFingerprint(input, normalizedActor);
  const insert = db.prepare(`INSERT INTO supplier_allocation_events
    (id,fingerprint,type,item_id,order_id,order_line_id,quantity,quantity_delta,movement_id,note,actor,created_at)
    SELECT ?,?,?,?,?,?,?,?,NULL,?,?,? WHERE NOT EXISTS(SELECT 1 FROM supplier_allocation_events WHERE id=?)`)
    .bind(input.operation_id,fingerprint,input.type,input.item_id,input.order_id,input.order_line_id,input.quantity,
      input.type === 'ALLOCATE' ? input.quantity : -input.quantity,input.note,normalizedActor,new Date().toISOString(),input.operation_id);
  let results: D1Result<AllocationEvent>[];
  try {
    results = await db.batch<AllocationEvent>([insert,
      db.prepare(`${allocationSelect} WHERE a.id=?`).bind(input.operation_id)]);
  } catch (error) {
    const message = error instanceof Error ? `${error.message} ${error.cause ?? ''}` : String(error);
    for (const [code, explanation] of Object.entries(guardErrors)) {
      if (message.includes(code)) throw new DomainError(409,code,explanation);
    }
    throw error;
  }
  const allocation = results[1].results[0];
  if (!allocation) throw new Error('Allocation event was not returned after recording.');
  if (allocation.fingerprint !== fingerprint) {
    throw new DomainError(409,'IDEMPOTENCY_CONFLICT','This operation ID was already used with different details.');
  }
  return { allocation, duplicate: results[0].meta.changes === 0 };
}

export async function listAllocations(db: D1Database): Promise<AllocationEvent[]> {
  return (await db.prepare(`${allocationSelect} ORDER BY a.created_at DESC,a.id DESC LIMIT 100`).all<AllocationEvent>()).results;
}
