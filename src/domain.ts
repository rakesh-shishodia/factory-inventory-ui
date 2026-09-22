export type MovementType = 'ECWID_PICK' | 'EMAIL_SALE' | 'INTERNAL_USE' | 'RESTOCK';
export type SyncStatus = 'NOT_REQUIRED' | 'PENDING' | 'PROCESSING' | 'APPLIED' | 'UNKNOWN' | 'BLOCKED';
export type InventoryMode = 'STOCK_LIMITED' | 'SUPPLIER_BACKED_UNLIMITED';
export type LineManagementMode = 'APP' | 'WORKBOOK';
export type LineFulfillmentState = 'REVIEW' | 'PICKED' | 'CLOSED' | 'AWAITING_PAYMENT'
  | 'READY_TO_PICK' | 'AWAITING_ASSIGNMENT' | 'AWAITING_SUPPLIER' | 'AWAITING_STOCK' | 'WORKBOOK_MANAGED';

export const PICKABLE_FULFILLMENT_STATUSES: readonly string[] = ['AWAITING_PROCESSING', 'PROCESSING'];
// Store policy: self-pickup orders reach READY_FOR_PICKUP only after picking and
// packing. It closes picking like SHIPPED; it does not mean customer delivery.
export const TERMINAL_FULFILLMENT_STATUSES: readonly string[] = [
  'READY_FOR_PICKUP', 'SHIPPED', 'DELIVERED', 'OUT_FOR_DELIVERY', 'RETURNED', 'WILL_NOT_DELIVER',
];

export class DomainError extends Error {
  constructor(public readonly status: number, public readonly code: string, message: string) {
    super(message);
    this.name = 'DomainError';
  }
}

export interface MovementInput {
  operation_id: string;
  type: MovementType;
  item_id: string;
  quantity: number;
  order_id?: string;
  order_line_id?: string;
  note?: string;
}

export interface Item {
  id: string;
  sku: string;
  name: string;
  scan_code: string;
  location: string;
  ecwid_product_id: string | null;
  ecwid_combination_id: string | null;
  ecwid_option_signature: string;
  on_hand: number;
  reserved: number;
  available: number;
  last_ecwid_quantity: number | null;
  active: number;
  inventory_mode: InventoryMode;
  supplier_name: string;
  opening_verified: number;
  allocated: number;
  free: number;
  unallocated_demand: number;
  uncovered_demand: number;
  paid_demand: number;
  awaiting_payment_demand: number;
}

export interface OrderLine {
  id: string;
  order_id: string;
  ecwid_line_id: string;
  item_id: string | null;
  management_mode: LineManagementMode;
  workbook_target_id: string | null;
  ecwid_combination_id: string | null;
  ecwid_option_signature: string | null;
  sku: string;
  name: string;
  ordered_qty: number;
  picked_qty: number;
  remaining_qty: number;
  inventory_mode: InventoryMode | null;
  opening_verified: number | null;
  item_active: number | null;
  on_hand: number | null;
  allocated_qty: number;
  free_qty: number | null;
  unallocated_qty: number;
  pickable_qty: number;
  fulfillment_state: LineFulfillmentState;
}

export interface Order {
  id: string;
  payment_status: string;
  fulfillment_status: string;
  remote_updated_at: string;
  updated_at: string;
  needs_review: number;
  lines: OrderLine[];
}

export interface Movement {
  id: string;
  fingerprint: string;
  type: MovementType;
  item_id: string;
  sku: string;
  name: string;
  quantity: number;
  quantity_delta: number;
  ecwid_quantity_delta: number;
  order_id: string | null;
  order_line_id: string | null;
  note: string;
  actor: string;
  created_at: string;
  sync_status: SyncStatus;
  inventory_mode: InventoryMode;
}

export function normalizeCode(value: string): string {
  return value.trim().toUpperCase();
}

export function validateMovementInput(value: unknown): MovementInput {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new DomainError(400, 'INVALID_INPUT', 'A movement object is required.');
  }
  const input = value as Record<string, unknown>;
  if (typeof input.operation_id !== 'string' || !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(input.operation_id)) {
    throw new DomainError(400, 'INVALID_OPERATION_ID', 'A UUID operation ID is required.');
  }
  if (!['ECWID_PICK', 'EMAIL_SALE', 'INTERNAL_USE', 'RESTOCK'].includes(String(input.type))) {
    throw new DomainError(400, 'INVALID_MOVEMENT_TYPE', 'Choose a supported stock movement.');
  }
  if (typeof input.item_id !== 'string' || !input.item_id.trim() || input.item_id.length > 200) {
    throw new DomainError(400, 'INVALID_ITEM', 'An item ID is required.');
  }
  if (typeof input.quantity !== 'number' || !Number.isInteger(input.quantity) || input.quantity < 1 || input.quantity > 1000000) {
    throw new DomainError(400, 'INVALID_QUANTITY', 'Quantity must be a whole number between 1 and 1,000,000.');
  }
  if (input.note !== undefined && (typeof input.note !== 'string' || input.note.length > 1000)) {
    throw new DomainError(400, 'INVALID_NOTE', 'The note must contain at most 1,000 characters.');
  }
  if (input.type === 'ECWID_PICK') {
    for (const key of ['order_id', 'order_line_id']) {
      if (typeof input[key] !== 'string' || !(input[key] as string).trim() || (input[key] as string).length > 200) {
        throw new DomainError(400, 'ORDER_REQUIRED', 'An order and an order line are required for picking.');
      }
    }
  } else if (input.order_id !== undefined || input.order_line_id !== undefined) {
    throw new DomainError(400, 'UNEXPECTED_ORDER', 'Only an Ecwid order pick can have an order association.');
  }
  return {
    operation_id: input.operation_id.toLowerCase(),
    type: input.type as MovementType,
    item_id: input.item_id.trim(),
    quantity: input.quantity,
    ...(input.type === 'ECWID_PICK' ? { order_id: (input.order_id as string).trim(), order_line_id: (input.order_line_id as string).trim() } : {}),
    note: ((input.note as string | undefined) ?? '').trim(),
  };
}

export async function movementFingerprint(input: MovementInput, actor: string): Promise<string> {
  const payload = JSON.stringify([input.type, input.item_id, input.quantity, input.order_id ?? null, input.order_line_id ?? null, input.note ?? '', actor]);
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(payload));
  return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
}
