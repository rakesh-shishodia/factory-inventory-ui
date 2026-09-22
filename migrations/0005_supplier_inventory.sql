-- Forward-only supplier mode. Existing stock-limited balances, identities,
-- movements and outbox references are copied unchanged under deferred FKs.
PRAGMA defer_foreign_keys = ON;

ALTER TABLE items ADD COLUMN inventory_mode TEXT NOT NULL DEFAULT 'STOCK_LIMITED'
  CHECK (inventory_mode IN ('STOCK_LIMITED','SUPPLIER_BACKED_UNLIMITED'));
ALTER TABLE items ADD COLUMN supplier_name TEXT NOT NULL DEFAULT '';

DROP VIEW item_stock;
DROP TRIGGER movement_guard;
DROP TRIGGER movement_apply;
DROP TRIGGER opening_balance_guard;
DROP TRIGGER item_mapping_immutable;

CREATE TABLE movements_supplier_migration (
  id TEXT PRIMARY KEY,
  fingerprint TEXT NOT NULL,
  type TEXT NOT NULL CHECK (type IN ('ECWID_PICK','EMAIL_SALE','INTERNAL_USE','RESTOCK')),
  item_id TEXT NOT NULL REFERENCES items(id),
  quantity INTEGER NOT NULL CHECK (typeof(quantity)='integer' AND quantity BETWEEN 1 AND 1000000),
  quantity_delta INTEGER NOT NULL,
  ecwid_quantity_delta INTEGER NOT NULL,
  order_id TEXT REFERENCES orders(id),
  order_line_id TEXT REFERENCES order_lines(id),
  note TEXT NOT NULL DEFAULT '',
  actor TEXT NOT NULL,
  created_at TEXT NOT NULL,
  inventory_mode TEXT NOT NULL DEFAULT 'STOCK_LIMITED'
    CHECK (inventory_mode IN ('STOCK_LIMITED','SUPPLIER_BACKED_UNLIMITED')),
  CHECK (quantity_delta=CASE WHEN type='RESTOCK' THEN quantity ELSE -quantity END),
  CHECK (ecwid_quantity_delta=CASE WHEN inventory_mode='SUPPLIER_BACKED_UNLIMITED' OR type='ECWID_PICK'
    THEN 0 WHEN type='RESTOCK' THEN quantity ELSE -quantity END),
  CHECK ((type='ECWID_PICK' AND order_id IS NOT NULL AND order_line_id IS NOT NULL)
    OR (type<>'ECWID_PICK' AND order_id IS NULL AND order_line_id IS NULL))
);
INSERT INTO movements_supplier_migration
  (id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,order_id,order_line_id,note,actor,created_at)
SELECT id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,order_id,order_line_id,note,actor,created_at
FROM movements;
DROP TABLE movements;
ALTER TABLE movements_supplier_migration RENAME TO movements;
CREATE INDEX movements_item_created ON movements(item_id,created_at DESC);
CREATE INDEX movements_created ON movements(created_at DESC);

CREATE TABLE supplier_allocation_events (
  id TEXT PRIMARY KEY,
  fingerprint TEXT NOT NULL CHECK (length(fingerprint)>0),
  type TEXT NOT NULL CHECK (type IN ('ALLOCATE','RELEASE','PICK','CANCEL_RELEASE')),
  item_id TEXT NOT NULL REFERENCES items(id),
  order_id TEXT NOT NULL REFERENCES orders(id),
  order_line_id TEXT NOT NULL REFERENCES order_lines(id),
  -- A cancellation closes all accumulated assignments in one immutable event.
  -- Individual operator actions remain bounded; their total can exceed 1m.
  quantity INTEGER NOT NULL CHECK (typeof(quantity)='integer'
    AND quantity BETWEEN 1 AND CASE WHEN type='CANCEL_RELEASE' THEN 2147483647 ELSE 1000000 END),
  quantity_delta INTEGER NOT NULL,
  movement_id TEXT UNIQUE REFERENCES movements(id),
  note TEXT NOT NULL DEFAULT '',
  actor TEXT NOT NULL CHECK (length(trim(actor))>0),
  created_at TEXT NOT NULL,
  CHECK (quantity_delta=CASE WHEN type='ALLOCATE' THEN quantity ELSE -quantity END),
  CHECK ((type='PICK' AND movement_id IS NOT NULL) OR (type<>'PICK' AND movement_id IS NULL))
);
CREATE INDEX supplier_allocation_line ON supplier_allocation_events(order_line_id);
CREATE INDEX supplier_allocation_item ON supplier_allocation_events(item_id);
CREATE INDEX supplier_allocation_order ON supplier_allocation_events(order_id);

-- A projection, not another mutable balance. Held/review allocations remain
-- allocated even after a remote terminal status until explicitly reconciled.
CREATE VIEW supplier_line_allocations AS
SELECT order_line_id,order_id,item_id,SUM(quantity_delta) AS allocated_qty
FROM supplier_allocation_events GROUP BY order_line_id,order_id,item_id;

CREATE VIEW item_stock AS
WITH allocation AS (
  SELECT item_id,SUM(allocated_qty) AS allocated FROM supplier_line_allocations GROUP BY item_id
), demand AS (
  SELECT l.item_id,SUM(l.ordered_qty-l.picked_qty) AS reserved,
    SUM(CASE WHEN o.payment_status='PAID' THEN l.ordered_qty-l.picked_qty ELSE 0 END) AS paid,
    SUM(CASE WHEN o.payment_status='AWAITING_PAYMENT' THEN l.ordered_qty-l.picked_qty ELSE 0 END) AS unpaid,
    SUM(MAX(0,l.ordered_qty-l.picked_qty-COALESCE(a.allocated_qty,0))) AS unallocated
  FROM order_lines l JOIN orders o ON o.id=l.order_id
  LEFT JOIN supplier_line_allocations a ON a.order_line_id=l.id
  WHERE o.payment_status IN ('PAID','AWAITING_PAYMENT') OR o.needs_review=1
  GROUP BY l.item_id
), legacy_reserved AS (
  SELECT l.item_id,SUM(l.ordered_qty-l.picked_qty) AS reserved
  FROM order_lines l JOIN orders o ON o.id=l.order_id
  WHERE o.payment_status IN ('PAID','AWAITING_PAYMENT') GROUP BY l.item_id
)
SELECT i.*,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(a.allocated,0)
    ELSE COALESCE(r.reserved,0) END AS reserved,
  i.on_hand-CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(a.allocated,0)
    ELSE COALESCE(r.reserved,0) END AS available,
  COALESCE(a.allocated,0) AS allocated,i.on_hand-COALESCE(a.allocated,0) AS free,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(d.reserved,0) ELSE 0 END AS supplier_demand,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(d.paid,0) ELSE 0 END AS supplier_paid_demand,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(d.unpaid,0) ELSE 0 END AS supplier_awaiting_payment_demand,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(d.unallocated,0) ELSE 0 END AS supplier_unallocated_demand,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN MAX(0,COALESCE(d.unallocated,0)-(i.on_hand-COALESCE(a.allocated,0))) ELSE 0 END AS supplier_uncovered_demand,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN COALESCE(d.unallocated,0) ELSE 0 END AS unallocated_demand,
  CASE WHEN i.inventory_mode='SUPPLIER_BACKED_UNLIMITED' THEN MAX(0,COALESCE(d.unallocated,0)-(i.on_hand-COALESCE(a.allocated,0))) ELSE 0 END AS uncovered_demand,
  EXISTS(SELECT 1 FROM opening_balances b WHERE b.item_id=i.id) AS opening_verified
FROM items i LEFT JOIN allocation a ON a.item_id=i.id LEFT JOIN demand d ON d.item_id=i.id
LEFT JOIN legacy_reserved r ON r.item_id=i.id;

CREATE TRIGGER supplier_item_insert BEFORE INSERT ON items
WHEN NEW.inventory_mode='SUPPLIER_BACKED_UNLIMITED' BEGIN
  SELECT CASE WHEN NEW.active<>0 OR NEW.on_hand<>0
    THEN RAISE(ABORT,'SUPPLIER_OPENING_REQUIRED') END;
END;
CREATE TRIGGER item_mode_immutable BEFORE UPDATE OF inventory_mode ON items
WHEN OLD.inventory_mode IS NOT NEW.inventory_mode AND (OLD.active=1
  OR EXISTS(SELECT 1 FROM opening_balances WHERE item_id=OLD.id)
  OR EXISTS(SELECT 1 FROM movements WHERE item_id=OLD.id)
  OR EXISTS(SELECT 1 FROM order_lines WHERE item_id=OLD.id)
  OR EXISTS(SELECT 1 FROM outbox WHERE item_id=OLD.id)) BEGIN
  SELECT RAISE(ABORT,'INVENTORY_MODE_IMMUTABLE');
END;
CREATE TRIGGER supplier_activation_guard BEFORE UPDATE OF active,inventory_mode ON items
WHEN NEW.inventory_mode='SUPPLIER_BACKED_UNLIMITED' AND NEW.active=1 BEGIN
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM opening_balances WHERE item_id=NEW.id)
    THEN RAISE(ABORT,'SUPPLIER_OPENING_REQUIRED') END;
  SELECT CASE WHEN length(trim(NEW.supplier_name))=0 OR NEW.ecwid_product_id IS NULL OR length(trim(NEW.ecwid_product_id))=0
    THEN RAISE(ABORT,'SUPPLIER_MAPPING_REQUIRED') END;
END;
CREATE TRIGGER supplier_balance_guard BEFORE UPDATE OF on_hand ON items
WHEN NEW.inventory_mode='SUPPLIER_BACKED_UNLIMITED' BEGIN
  SELECT CASE WHEN NEW.on_hand<>COALESCE((SELECT on_hand FROM opening_balances WHERE item_id=NEW.id),0)
      +COALESCE((SELECT SUM(quantity_delta) FROM movements WHERE item_id=NEW.id),0)
    THEN RAISE(ABORT,'SUPPLIER_BALANCE_REQUIRES_LEDGER') END;
  SELECT CASE WHEN NEW.on_hand<COALESCE((SELECT SUM(allocated_qty) FROM supplier_line_allocations WHERE item_id=NEW.id),0)
    THEN RAISE(ABORT,'INSUFFICIENT_FREE_SUPPLIER_STOCK') END;
END;
CREATE TRIGGER item_mapping_immutable BEFORE UPDATE OF id,sku,ecwid_product_id,ecwid_combination_id,ecwid_option_signature ON items
WHEN (OLD.id IS NOT NEW.id OR OLD.sku IS NOT NEW.sku OR OLD.ecwid_product_id IS NOT NEW.ecwid_product_id
  OR OLD.ecwid_combination_id IS NOT NEW.ecwid_combination_id OR OLD.ecwid_option_signature IS NOT NEW.ecwid_option_signature)
  AND (EXISTS(SELECT 1 FROM opening_balances WHERE item_id=OLD.id)
    OR EXISTS(SELECT 1 FROM movements WHERE item_id=OLD.id)
    OR EXISTS(SELECT 1 FROM order_lines WHERE item_id=OLD.id)
    OR EXISTS(SELECT 1 FROM outbox WHERE item_id=OLD.id)) BEGIN
  SELECT RAISE(ABORT,'ITEM_MAPPING_IMMUTABLE');
END;
CREATE TRIGGER opening_balance_guard BEFORE INSERT ON opening_balances BEGIN
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM items WHERE id=NEW.item_id AND on_hand=0)
    OR EXISTS(SELECT 1 FROM movements WHERE item_id=NEW.item_id)
    THEN RAISE(ABORT,'OPENING_BALANCE_ALREADY_STARTED') END;
END;

CREATE TRIGGER supplier_allocation_guard BEFORE INSERT ON supplier_allocation_events BEGIN
  SELECT CASE WHEN NEW.type IN ('ALLOCATE','RELEASE') AND EXISTS(SELECT 1 FROM movements WHERE id=NEW.id)
    THEN RAISE(ABORT,'IDEMPOTENCY_CONFLICT') END;
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM items i JOIN opening_balances b ON b.item_id=i.id
    WHERE i.id=NEW.item_id AND i.inventory_mode='SUPPLIER_BACKED_UNLIMITED'
      AND (i.active=1 OR NEW.type='CANCEL_RELEASE')) THEN RAISE(ABORT,'SUPPLIER_ITEM_UNAVAILABLE') END;
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM order_lines l JOIN items i ON i.id=l.item_id
    WHERE l.id=NEW.order_line_id AND l.order_id=NEW.order_id AND l.item_id=NEW.item_id
      AND l.sku=i.sku COLLATE NOCASE) THEN RAISE(ABORT,'ALLOCATION_LINE_MISMATCH') END;
  SELECT CASE WHEN NEW.type<>'CANCEL_RELEASE' AND (
      EXISTS(SELECT 1 FROM sync_issues WHERE status='OPEN' AND (item_id=NEW.item_id OR order_id=NEW.order_id))
      OR EXISTS(SELECT 1 FROM outbox WHERE item_id=NEW.item_id AND status IN ('UNKNOWN','BLOCKED')))
    THEN RAISE(ABORT,'ITEM_NEEDS_REVIEW') END;
  SELECT CASE WHEN NEW.type IN ('ALLOCATE','RELEASE','PICK') AND NOT EXISTS(
    SELECT 1 FROM orders o WHERE o.id=NEW.order_id AND o.needs_review=0
      AND o.payment_status IN ('PAID','AWAITING_PAYMENT')
      AND (NEW.type<>'PICK' OR o.payment_status='PAID')
      AND o.fulfillment_status IN ('AWAITING_PROCESSING','PROCESSING')
      AND NOT EXISTS(SELECT 1 FROM order_lines WHERE order_id=o.id AND item_id IS NULL))
    THEN RAISE(ABORT,'ORDER_NOT_ALLOCATABLE') END;
  SELECT CASE WHEN NEW.type='ALLOCATE' AND NEW.quantity>(SELECT free FROM item_stock WHERE id=NEW.item_id)
    THEN RAISE(ABORT,'INSUFFICIENT_FREE_SUPPLIER_STOCK') END;
  SELECT CASE WHEN NEW.type='ALLOCATE' AND NEW.quantity>(
    SELECT l.ordered_qty-l.picked_qty-COALESCE(a.allocated_qty,0) FROM order_lines l
      LEFT JOIN supplier_line_allocations a ON a.order_line_id=l.id WHERE l.id=NEW.order_line_id)
    THEN RAISE(ABORT,'ALLOCATION_QUANTITY_EXCEEDED') END;
  SELECT CASE WHEN NEW.type<>'ALLOCATE' AND NEW.quantity>COALESCE((
    SELECT allocated_qty FROM supplier_line_allocations WHERE order_line_id=NEW.order_line_id),0)
    THEN RAISE(ABORT,'INSUFFICIENT_LINE_ALLOCATION') END;
  SELECT CASE WHEN NEW.type='PICK' AND NOT EXISTS(SELECT 1 FROM movements m WHERE m.id=NEW.movement_id
    AND m.type='ECWID_PICK' AND m.inventory_mode='SUPPLIER_BACKED_UNLIMITED' AND m.item_id=NEW.item_id
    AND m.order_id=NEW.order_id AND m.order_line_id=NEW.order_line_id AND m.quantity=NEW.quantity
    AND m.actor=NEW.actor AND m.fingerprint=NEW.fingerprint)
    THEN RAISE(ABORT,'ALLOCATION_PICK_REQUIRES_MOVEMENT') END;
  SELECT CASE WHEN NEW.type='CANCEL_RELEASE' AND NOT EXISTS(SELECT 1 FROM orders o WHERE o.id=NEW.order_id
    AND o.needs_review=0 AND o.payment_status IN ('CANCELLED','REFUNDED','INCOMPLETE')
    AND o.fulfillment_status IN ('AWAITING_PROCESSING','PROCESSING')
    AND NOT EXISTS(SELECT 1 FROM order_lines l WHERE l.order_id=o.id AND l.picked_qty>0))
    THEN RAISE(ABORT,'CANCELLATION_REQUIRES_REVIEW') END;
END;
CREATE TRIGGER supplier_allocation_no_update BEFORE UPDATE ON supplier_allocation_events BEGIN
  SELECT RAISE(ABORT,'ALLOCATION_IMMUTABLE');
END;
CREATE TRIGGER supplier_allocation_no_delete BEFORE DELETE ON supplier_allocation_events BEGIN
  SELECT RAISE(ABORT,'ALLOCATION_IMMUTABLE');
END;

CREATE TRIGGER supplier_order_line_identity_guard BEFORE UPDATE OF item_id,order_id,ecwid_line_id,sku,ordered_qty ON order_lines
WHEN (OLD.item_id IS NOT NEW.item_id OR OLD.order_id IS NOT NEW.order_id OR OLD.ecwid_line_id IS NOT NEW.ecwid_line_id
  OR OLD.sku IS NOT NEW.sku OR OLD.ordered_qty IS NOT NEW.ordered_qty)
  AND EXISTS(SELECT 1 FROM supplier_allocation_events WHERE order_line_id=OLD.id) BEGIN
  SELECT RAISE(ABORT,'ALLOCATED_ORDER_LINE_IMMUTABLE');
END;
CREATE TRIGGER supplier_order_line_pick_guard BEFORE UPDATE OF picked_qty ON order_lines
WHEN EXISTS(SELECT 1 FROM items WHERE id=NEW.item_id AND inventory_mode='SUPPLIER_BACKED_UNLIMITED') BEGIN
  SELECT CASE WHEN NEW.picked_qty<>COALESCE((SELECT SUM(quantity) FROM movements
    WHERE order_line_id=NEW.id AND type='ECWID_PICK'),0) THEN RAISE(ABORT,'SUPPLIER_PICK_REQUIRES_MOVEMENT') END;
END;

CREATE TRIGGER supplier_order_status AFTER UPDATE OF payment_status,fulfillment_status,needs_review ON orders
WHEN EXISTS(SELECT 1 FROM order_lines l JOIN items i ON i.id=l.item_id
  WHERE l.order_id=NEW.id AND i.inventory_mode='SUPPLIER_BACKED_UNLIMITED') BEGIN
  -- Any pick in a mixed order makes cancellation a reconciliation, not a release.
  UPDATE orders SET needs_review=1 WHERE id=NEW.id AND needs_review=0 AND (
    (payment_status NOT IN ('PAID','AWAITING_PAYMENT') AND EXISTS(SELECT 1 FROM order_lines WHERE order_id=NEW.id AND picked_qty>0))
    OR (fulfillment_status IN ('READY_FOR_PICKUP','SHIPPED','DELIVERED','OUT_FOR_DELIVERY','RETURNED','WILL_NOT_DELIVER')
      AND EXISTS(SELECT 1 FROM order_lines WHERE order_id=NEW.id AND picked_qty<ordered_qty)));
  INSERT INTO supplier_allocation_events
    (id,fingerprint,type,item_id,order_id,order_line_id,quantity,quantity_delta,actor,note,created_at)
  SELECT 'cancel:'||lower(hex(randomblob(16))),'cancel:'||NEW.id||':'||NEW.remote_updated_at,'CANCEL_RELEASE',
    a.item_id,a.order_id,a.order_line_id,a.allocated_qty,-a.allocated_qty,'ecwid-sync',
    'Unpicked order cancelled; allocation released without changing shelf stock.',NEW.updated_at
  FROM supplier_line_allocations a JOIN orders o ON o.id=a.order_id WHERE a.order_id=NEW.id AND a.allocated_qty>0
    AND o.needs_review=0 AND o.payment_status IN ('CANCELLED','REFUNDED','INCOMPLETE')
    AND o.fulfillment_status IN ('AWAITING_PROCESSING','PROCESSING')
    AND NOT EXISTS(SELECT 1 FROM order_lines WHERE order_id=NEW.id AND picked_qty>0);
END;

CREATE TRIGGER movement_guard BEFORE INSERT ON movements BEGIN
  SELECT CASE WHEN EXISTS(SELECT 1 FROM supplier_allocation_events WHERE id=NEW.id)
    THEN RAISE(ABORT,'IDEMPOTENCY_CONFLICT') END;
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM items WHERE id=NEW.item_id AND active=1)
    THEN RAISE(ABORT,'ITEM_UNAVAILABLE') END;
  SELECT CASE WHEN NEW.inventory_mode<>(SELECT inventory_mode FROM items WHERE id=NEW.item_id)
    THEN RAISE(ABORT,'INVENTORY_MODE_MISMATCH') END;
  SELECT CASE WHEN EXISTS(SELECT 1 FROM sync_issues WHERE item_id=NEW.item_id AND status='OPEN')
    OR EXISTS(SELECT 1 FROM outbox WHERE item_id=NEW.item_id AND status IN ('UNKNOWN','BLOCKED'))
    THEN RAISE(ABORT,'ITEM_NEEDS_REVIEW') END;
  SELECT CASE WHEN (NEW.type<>'ECWID_PICK' OR NEW.inventory_mode='SUPPLIER_BACKED_UNLIMITED') AND NOT EXISTS(
    SELECT 1 FROM items WHERE id=NEW.item_id AND length(trim(ecwid_product_id))>0)
    THEN RAISE(ABORT,'ECWID_MAPPING_REQUIRED') END;
  SELECT CASE WHEN NEW.inventory_mode='SUPPLIER_BACKED_UNLIMITED' AND NOT EXISTS(
    SELECT 1 FROM opening_balances WHERE item_id=NEW.item_id) THEN RAISE(ABORT,'SUPPLIER_OPENING_REQUIRED') END;
  SELECT CASE WHEN NEW.type='ECWID_PICK' AND NOT EXISTS(
    SELECT 1 FROM orders o JOIN order_lines l ON l.order_id=o.id
    WHERE o.id=NEW.order_id AND l.id=NEW.order_line_id AND l.item_id=NEW.item_id
      AND o.payment_status='PAID' AND o.needs_review=0
      AND o.fulfillment_status IN ('AWAITING_PROCESSING','PROCESSING')) THEN RAISE(ABORT,'ORDER_NOT_PICKABLE') END;
  SELECT CASE WHEN NEW.type='ECWID_PICK' AND NEW.quantity>(
    SELECT ordered_qty-picked_qty FROM order_lines WHERE id=NEW.order_line_id) THEN RAISE(ABORT,'PICK_QUANTITY_EXCEEDED') END;
  SELECT CASE WHEN NEW.type<>'RESTOCK' AND NEW.quantity>(SELECT on_hand FROM items WHERE id=NEW.item_id)
    THEN RAISE(ABORT,'INSUFFICIENT_PHYSICAL_STOCK') END;
  SELECT CASE WHEN NEW.type='ECWID_PICK' AND NEW.inventory_mode='SUPPLIER_BACKED_UNLIMITED'
    AND NEW.quantity>COALESCE((SELECT allocated_qty FROM supplier_line_allocations WHERE order_line_id=NEW.order_line_id),0)
    THEN RAISE(ABORT,'INSUFFICIENT_LINE_ALLOCATION') END;
  SELECT CASE WHEN NEW.type IN ('EMAIL_SALE','INTERNAL_USE') AND NEW.quantity>(SELECT available FROM item_stock WHERE id=NEW.item_id)
    THEN RAISE(ABORT,'INSUFFICIENT_AVAILABLE_STOCK') END;
END;

CREATE TRIGGER movement_apply AFTER INSERT ON movements BEGIN
  -- Consume the allocation before the shelf balance changes, in the same write.
  INSERT INTO supplier_allocation_events
    (id,fingerprint,type,item_id,order_id,order_line_id,quantity,quantity_delta,movement_id,note,actor,created_at)
  SELECT 'pick:'||NEW.id,NEW.fingerprint,'PICK',NEW.item_id,NEW.order_id,NEW.order_line_id,
    NEW.quantity,-NEW.quantity,NEW.id,NEW.note,NEW.actor,NEW.created_at
  WHERE NEW.type='ECWID_PICK' AND NEW.inventory_mode='SUPPLIER_BACKED_UNLIMITED';
  UPDATE items SET on_hand=on_hand+NEW.quantity_delta WHERE id=NEW.item_id;
  UPDATE order_lines SET picked_qty=picked_qty+NEW.quantity WHERE NEW.type='ECWID_PICK' AND id=NEW.order_line_id;
  INSERT INTO outbox(id,item_id,ecwid_product_id,ecwid_combination_id,quantity_delta,status,attempts,created_at,updated_at)
  SELECT NEW.id,i.id,i.ecwid_product_id,i.ecwid_combination_id,NEW.ecwid_quantity_delta,'PENDING',0,NEW.created_at,NEW.created_at
  FROM items i WHERE i.id=NEW.item_id AND NEW.inventory_mode='STOCK_LIMITED' AND NEW.ecwid_quantity_delta<>0;
END;
CREATE TRIGGER movement_no_update BEFORE UPDATE ON movements BEGIN
  SELECT RAISE(ABORT,'MOVEMENT_IMMUTABLE');
END;
CREATE TRIGGER movement_no_delete BEFORE DELETE ON movements BEGIN
  SELECT RAISE(ABORT,'MOVEMENT_IMMUTABLE');
END;
CREATE TRIGGER supplier_outbox_guard BEFORE INSERT ON outbox BEGIN
  SELECT CASE WHEN EXISTS(SELECT 1 FROM items WHERE id=NEW.item_id AND inventory_mode='SUPPLIER_BACKED_UNLIMITED')
    OR EXISTS(SELECT 1 FROM movements WHERE id=NEW.id AND inventory_mode='SUPPLIER_BACKED_UNLIMITED')
    THEN RAISE(ABORT,'SUPPLIER_ECWID_WRITE_FORBIDDEN') END;
END;

PRAGMA defer_foreign_keys = OFF;
