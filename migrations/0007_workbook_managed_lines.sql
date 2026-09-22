-- Only explicitly reviewed exact identities may stay in the workbook. An
-- unknown or conflicting line is never made safe merely by being outside pilot.
CREATE TABLE workbook_managed_targets (
  id TEXT PRIMARY KEY,
  ecwid_product_id TEXT NOT NULL CHECK(length(trim(ecwid_product_id))>0),
  ecwid_combination_id TEXT,
  ecwid_option_signature TEXT NOT NULL CHECK(json_valid(ecwid_option_signature)),
  sku TEXT NOT NULL COLLATE NOCASE UNIQUE CHECK(length(trim(sku))>0),
  name TEXT NOT NULL,
  review_reference TEXT NOT NULL CHECK(length(trim(review_reference))>0),
  reviewed_by TEXT NOT NULL CHECK(length(trim(reviewed_by))>0),
  reviewed_at TEXT NOT NULL
);
CREATE UNIQUE INDEX workbook_target_identity ON workbook_managed_targets(ecwid_product_id,COALESCE(ecwid_combination_id,''));
CREATE TRIGGER workbook_target_guard BEFORE INSERT ON workbook_managed_targets BEGIN
  SELECT CASE WHEN EXISTS(SELECT 1 FROM items i WHERE i.sku=NEW.sku COLLATE NOCASE
    OR (i.ecwid_product_id=NEW.ecwid_product_id AND i.ecwid_combination_id IS NEW.ecwid_combination_id))
    THEN RAISE(ABORT,'WORKBOOK_APP_IDENTITY_CONFLICT') END;
  SELECT CASE WHEN EXISTS(SELECT 1 FROM workbook_managed_targets w WHERE w.id=NEW.id AND (
    w.ecwid_product_id IS NOT NEW.ecwid_product_id OR w.ecwid_combination_id IS NOT NEW.ecwid_combination_id
    OR w.ecwid_option_signature IS NOT NEW.ecwid_option_signature OR w.sku IS NOT NEW.sku COLLATE NOCASE
    OR w.name IS NOT NEW.name OR w.review_reference IS NOT NEW.review_reference OR w.reviewed_by IS NOT NEW.reviewed_by))
    THEN RAISE(ABORT,'WORKBOOK_TARGET_CONFLICT') END;
END;
CREATE TRIGGER workbook_target_no_update BEFORE UPDATE ON workbook_managed_targets BEGIN
  SELECT RAISE(ABORT,'WORKBOOK_TARGET_IMMUTABLE');
END;
CREATE TRIGGER workbook_target_no_delete BEFORE DELETE ON workbook_managed_targets BEGIN
  SELECT RAISE(ABORT,'WORKBOOK_TARGET_IMMUTABLE');
END;
CREATE TRIGGER item_workbook_conflict_insert BEFORE INSERT ON items BEGIN
  SELECT CASE WHEN EXISTS(SELECT 1 FROM workbook_managed_targets w WHERE w.sku=NEW.sku COLLATE NOCASE
    OR (w.ecwid_product_id=NEW.ecwid_product_id AND w.ecwid_combination_id IS NEW.ecwid_combination_id))
    THEN RAISE(ABORT,'WORKBOOK_APP_IDENTITY_CONFLICT') END;
END;
CREATE TRIGGER item_workbook_conflict_update BEFORE UPDATE OF sku,ecwid_product_id,ecwid_combination_id ON items BEGIN
  SELECT CASE WHEN EXISTS(SELECT 1 FROM workbook_managed_targets w WHERE w.sku=NEW.sku COLLATE NOCASE
    OR (w.ecwid_product_id=NEW.ecwid_product_id AND w.ecwid_combination_id IS NEW.ecwid_combination_id))
    THEN RAISE(ABORT,'WORKBOOK_APP_IDENTITY_CONFLICT') END;
END;

ALTER TABLE order_lines ADD COLUMN management_mode TEXT NOT NULL DEFAULT 'APP' CHECK(management_mode IN ('APP','WORKBOOK'));
ALTER TABLE order_lines ADD COLUMN workbook_target_id TEXT REFERENCES workbook_managed_targets(id);
CREATE TRIGGER workbook_order_line_insert BEFORE INSERT ON order_lines BEGIN
  SELECT CASE WHEN (NEW.management_mode='WORKBOOK' AND (NEW.item_id IS NOT NULL OR NEW.picked_qty<>0
    OR NOT EXISTS(SELECT 1 FROM workbook_managed_targets w WHERE w.id=NEW.workbook_target_id AND w.sku=NEW.sku COLLATE NOCASE)))
    OR (NEW.management_mode='APP' AND NEW.workbook_target_id IS NOT NULL)
    THEN RAISE(ABORT,'INVALID_WORKBOOK_ORDER_LINE') END;
END;
CREATE TRIGGER workbook_order_line_update BEFORE UPDATE ON order_lines BEGIN
  SELECT CASE WHEN NEW.management_mode IS NOT OLD.management_mode OR NEW.workbook_target_id IS NOT OLD.workbook_target_id
    THEN RAISE(ABORT,'ORDER_LINE_MANAGEMENT_IMMUTABLE') END;
  SELECT CASE WHEN NEW.management_mode='WORKBOOK' AND (NEW.item_id IS NOT NULL OR NEW.picked_qty<>0
    OR NEW.id IS NOT OLD.id OR NEW.order_id IS NOT OLD.order_id OR NEW.ecwid_line_id IS NOT OLD.ecwid_line_id
    OR NEW.sku IS NOT OLD.sku OR NEW.ordered_qty IS NOT OLD.ordered_qty)
    THEN RAISE(ABORT,'WORKBOOK_LINE_HANDLED_EXTERNALLY') END;
END;
CREATE TRIGGER workbook_movement_guard BEFORE INSERT ON movements
WHEN EXISTS(SELECT 1 FROM order_lines WHERE id=NEW.order_line_id AND management_mode='WORKBOOK') BEGIN
  SELECT RAISE(ABORT,'WORKBOOK_LINE_HANDLED_EXTERNALLY');
END;

DROP TRIGGER supplier_allocation_guard;
CREATE TRIGGER supplier_allocation_guard BEFORE INSERT ON supplier_allocation_events BEGIN
  SELECT CASE WHEN NEW.type IN ('ALLOCATE','RELEASE') AND EXISTS(SELECT 1 FROM movements WHERE id=NEW.id)
    THEN RAISE(ABORT,'IDEMPOTENCY_CONFLICT') END;
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM items i JOIN opening_balances b ON b.item_id=i.id
    WHERE i.id=NEW.item_id AND i.inventory_mode='SUPPLIER_BACKED_UNLIMITED'
      AND (i.active=1 OR NEW.type='CANCEL_RELEASE')) THEN RAISE(ABORT,'SUPPLIER_ITEM_UNAVAILABLE') END;
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM order_lines l JOIN items i ON i.id=l.item_id
    WHERE l.id=NEW.order_line_id AND l.order_id=NEW.order_id AND l.item_id=NEW.item_id AND l.management_mode='APP'
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
      AND NOT EXISTS(SELECT 1 FROM order_lines WHERE order_id=o.id AND management_mode='APP' AND item_id IS NULL))
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

DROP TRIGGER supplier_order_status;
CREATE TRIGGER supplier_order_status AFTER UPDATE OF payment_status,fulfillment_status,needs_review ON orders
WHEN EXISTS(SELECT 1 FROM order_lines l JOIN items i ON i.id=l.item_id
  WHERE l.order_id=NEW.id AND i.inventory_mode='SUPPLIER_BACKED_UNLIMITED') BEGIN
  UPDATE orders SET needs_review=1 WHERE id=NEW.id AND needs_review=0 AND (
    (payment_status NOT IN ('PAID','AWAITING_PAYMENT') AND EXISTS(SELECT 1 FROM order_lines WHERE order_id=NEW.id AND picked_qty>0))
    OR (fulfillment_status IN ('READY_FOR_PICKUP','SHIPPED','DELIVERED','OUT_FOR_DELIVERY','RETURNED','WILL_NOT_DELIVER')
      AND EXISTS(SELECT 1 FROM order_lines WHERE order_id=NEW.id AND management_mode='APP' AND picked_qty<ordered_qty)));
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
