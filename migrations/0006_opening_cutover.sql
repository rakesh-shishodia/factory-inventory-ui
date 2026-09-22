-- An opening cutover stages physical balances and exact open orders together.
-- Remote alignment is a separately journalled, explicitly authorized operation.
CREATE TABLE opening_cutover_batches (
  operation_id TEXT PRIMARY KEY,
  fingerprint TEXT NOT NULL CHECK(length(fingerprint)=64),
  store_id TEXT NOT NULL CHECK(length(store_id)>0 AND store_id NOT GLOB '*[^0-9]*' AND store_id GLOB '[1-9]*'),
  review_hash TEXT NOT NULL CHECK(length(review_hash)=64),
  preview_hash TEXT NOT NULL CHECK(length(preview_hash)=64),
  source_hash TEXT NOT NULL CHECK(length(source_hash)=64),
  orders_hash TEXT NOT NULL CHECK(length(orders_hash)=64),
  catalogue_hash TEXT NOT NULL CHECK(length(catalogue_hash)=64),
  source_ref TEXT NOT NULL CHECK(length(trim(source_ref))>0),
  scope_json TEXT NOT NULL CHECK(json_valid(scope_json) AND json_type(scope_json)='array'),
  orders_json TEXT NOT NULL CHECK(json_valid(orders_json) AND json_type(orders_json)='object'),
  confirmations_json TEXT NOT NULL CHECK(json_valid(confirmations_json) AND json_type(confirmations_json)='array'),
  workbook_scope_json TEXT NOT NULL CHECK(json_valid(workbook_scope_json) AND json_type(workbook_scope_json)='array'),
  row_count INTEGER NOT NULL CHECK(typeof(row_count)='integer' AND row_count BETWEEN 1 AND 200 AND json_array_length(scope_json)=row_count),
  order_count INTEGER NOT NULL CHECK(typeof(order_count)='integer' AND order_count BETWEEN 0 AND 200),
  line_count INTEGER NOT NULL CHECK(typeof(line_count)='integer' AND line_count BETWEEN 0 AND 2000 AND json_array_length(confirmations_json)=line_count),
  actor TEXT NOT NULL CHECK(length(trim(actor))>0),
  frozen_at TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  state TEXT NOT NULL DEFAULT 'STAGED' CHECK(state IN ('STAGED','ALIGNING','ALIGNED','ACTIVE','REVIEW'))
);

CREATE TABLE opening_cutover_rows (
  operation_id TEXT NOT NULL REFERENCES opening_cutover_batches(operation_id),
  item_id TEXT NOT NULL UNIQUE REFERENCES items(id),
  opening_balance_id TEXT NOT NULL UNIQUE REFERENCES opening_balances(id),
  sku TEXT NOT NULL,
  ecwid_product_id TEXT NOT NULL,
  ecwid_combination_id TEXT,
  ecwid_option_signature TEXT NOT NULL,
  physical INTEGER NOT NULL CHECK(typeof(physical)='integer' AND physical BETWEEN 0 AND 2147483647),
  unpicked INTEGER NOT NULL CHECK(typeof(unpicked)='integer' AND unpicked BETWEEN 0 AND physical),
  target_quantity INTEGER NOT NULL CHECK(typeof(target_quantity)='integer' AND target_quantity=physical-unpicked),
  expected_ecwid_quantity INTEGER NOT NULL CHECK(typeof(expected_ecwid_quantity)='integer' AND expected_ecwid_quantity BETWEEN 0 AND 2147483647),
  source_row INTEGER CHECK(source_row IS NULL OR (typeof(source_row)='integer' AND source_row>0)),
  source_sheet TEXT,
  alignment_status TEXT NOT NULL DEFAULT 'PENDING' CHECK(alignment_status IN ('PENDING','PROCESSING','VERIFIED','UNKNOWN','BLOCKED')),
  before_quantity INTEGER CHECK(before_quantity IS NULL OR (typeof(before_quantity)='integer' AND before_quantity BETWEEN 0 AND 2147483647)),
  after_quantity INTEGER CHECK(after_quantity IS NULL OR (typeof(after_quantity)='integer' AND after_quantity BETWEEN 0 AND 2147483647)),
  last_error TEXT,
  attempted_at TEXT,
  verified_at TEXT,
  PRIMARY KEY(operation_id,item_id),
  CHECK(alignment_status<>'VERIFIED' OR (before_quantity IS NOT NULL AND after_quantity IS NOT NULL
    AND after_quantity=target_quantity AND verified_at IS NOT NULL)),
  CHECK(alignment_status<>'PROCESSING' OR (before_quantity IS NOT NULL AND attempted_at IS NOT NULL))
);

CREATE TABLE opening_cutover_orders (
  operation_id TEXT NOT NULL REFERENCES opening_cutover_batches(operation_id),
  order_id TEXT NOT NULL UNIQUE REFERENCES orders(id),
  remote_lines_hash TEXT NOT NULL CHECK(length(remote_lines_hash)=64),
  line_count INTEGER NOT NULL CHECK(typeof(line_count)='integer' AND line_count BETWEEN 1 AND 500),
  PRIMARY KEY(operation_id,order_id)
);

CREATE TRIGGER opening_cutover_store_guard BEFORE INSERT ON opening_cutover_batches BEGIN
  SELECT CASE WHEN NEW.state<>'STAGED' THEN RAISE(ABORT,'OPENING_CUTOVER_STATE_INVALID') END;
  SELECT CASE WHEN EXISTS(SELECT 1 FROM opening_cutover_batches WHERE store_id<>NEW.store_id)
    OR EXISTS(SELECT 1 FROM opening_import_batches WHERE store_id<>NEW.store_id)
    THEN RAISE(ABORT,'OPENING_DATABASE_STORE_MISMATCH') END;
END;
CREATE TRIGGER opening_import_cutover_store_guard BEFORE INSERT ON opening_import_batches BEGIN
  SELECT CASE WHEN EXISTS(SELECT 1 FROM opening_cutover_batches WHERE store_id<>NEW.store_id)
    THEN RAISE(ABORT,'OPENING_DATABASE_STORE_MISMATCH') END;
END;

CREATE TRIGGER opening_cutover_row_guard BEFORE INSERT ON opening_cutover_rows BEGIN
  SELECT CASE WHEN NEW.alignment_status<>'PENDING' OR NEW.before_quantity IS NOT NULL OR NEW.after_quantity IS NOT NULL
      OR NEW.attempted_at IS NOT NULL OR NEW.verified_at IS NOT NULL OR NEW.last_error IS NOT NULL
    THEN RAISE(ABORT,'OPENING_CUTOVER_STATE_INVALID') END;
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM items i JOIN opening_balances o ON o.item_id=i.id
    JOIN opening_cutover_batches b ON b.operation_id=NEW.operation_id
    WHERE i.id=NEW.item_id AND o.id=NEW.opening_balance_id AND b.state='STAGED'
      AND i.active=0 AND i.inventory_mode='STOCK_LIMITED' AND i.sku=NEW.sku COLLATE BINARY
      AND i.ecwid_product_id=NEW.ecwid_product_id AND i.ecwid_combination_id IS NEW.ecwid_combination_id
      AND i.ecwid_option_signature=NEW.ecwid_option_signature AND i.on_hand=NEW.physical AND o.on_hand=NEW.physical
      AND i.last_ecwid_quantity=NEW.expected_ecwid_quantity AND o.source_ref=b.source_ref
      AND o.actor=b.actor AND o.created_at=b.created_at)
    THEN RAISE(ABORT,'OPENING_CUTOVER_AUDIT_MISMATCH') END;
  SELECT CASE WHEN NEW.unpicked<>COALESCE((SELECT SUM(l.ordered_qty-l.picked_qty) FROM order_lines l
    JOIN orders o ON o.id=l.order_id WHERE l.item_id=NEW.item_id AND o.payment_status IN ('PAID','AWAITING_PAYMENT')),0)
    THEN RAISE(ABORT,'OPENING_CUTOVER_RESERVATIONS_MISMATCH') END;
END;

CREATE TRIGGER opening_cutover_order_guard BEFORE INSERT ON opening_cutover_orders BEGIN
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM orders o JOIN opening_cutover_batches b ON b.operation_id=NEW.operation_id
    WHERE o.id=NEW.order_id AND b.state='STAGED' AND o.remote_lines_hash=NEW.remote_lines_hash AND o.needs_review=0
      AND o.payment_status IN ('PAID','AWAITING_PAYMENT') AND o.fulfillment_status IN ('AWAITING_PROCESSING','PROCESSING')
      AND NEW.line_count=(SELECT COUNT(*) FROM order_lines WHERE order_id=o.id)
      AND NOT EXISTS(SELECT 1 FROM order_lines WHERE order_id=o.id AND picked_qty<>0))
    THEN RAISE(ABORT,'OPENING_CUTOVER_AUDIT_MISMATCH') END;
END;

CREATE TRIGGER opening_cutover_batch_guard BEFORE UPDATE ON opening_cutover_batches BEGIN
  SELECT CASE WHEN OLD.operation_id IS NOT NEW.operation_id OR OLD.fingerprint IS NOT NEW.fingerprint
    OR OLD.store_id IS NOT NEW.store_id OR OLD.review_hash IS NOT NEW.review_hash OR OLD.preview_hash IS NOT NEW.preview_hash
    OR OLD.source_hash IS NOT NEW.source_hash OR OLD.orders_hash IS NOT NEW.orders_hash OR OLD.catalogue_hash IS NOT NEW.catalogue_hash
    OR OLD.source_ref IS NOT NEW.source_ref OR OLD.scope_json IS NOT NEW.scope_json OR OLD.orders_json IS NOT NEW.orders_json
    OR OLD.confirmations_json IS NOT NEW.confirmations_json OR OLD.workbook_scope_json IS NOT NEW.workbook_scope_json
    OR OLD.row_count IS NOT NEW.row_count OR OLD.order_count IS NOT NEW.order_count OR OLD.line_count IS NOT NEW.line_count
    OR OLD.actor IS NOT NEW.actor OR OLD.frozen_at IS NOT NEW.frozen_at OR OLD.created_at IS NOT NEW.created_at
    THEN RAISE(ABORT,'OPENING_CUTOVER_IMMUTABLE') END;
  SELECT CASE WHEN NOT (OLD.state=NEW.state OR (OLD.state='STAGED' AND NEW.state IN ('ALIGNING','REVIEW'))
    OR (OLD.state='ALIGNING' AND NEW.state IN ('ALIGNED','REVIEW')) OR (OLD.state='ALIGNED' AND NEW.state IN ('ACTIVE','REVIEW')))
    THEN RAISE(ABORT,'OPENING_CUTOVER_STATE_INVALID') END;
  SELECT CASE WHEN NEW.row_count<>(SELECT COUNT(*) FROM opening_cutover_rows WHERE operation_id=NEW.operation_id)
    OR NEW.order_count<>(SELECT COUNT(*) FROM opening_cutover_orders WHERE operation_id=NEW.operation_id)
    OR NEW.line_count<>COALESCE((SELECT SUM(line_count) FROM opening_cutover_orders WHERE operation_id=NEW.operation_id),0)
    THEN RAISE(ABORT,'OPENING_CUTOVER_AUDIT_MISMATCH') END;
  SELECT CASE WHEN NEW.state IN ('ALIGNED','ACTIVE') AND EXISTS(SELECT 1 FROM opening_cutover_rows
    WHERE operation_id=NEW.operation_id AND alignment_status<>'VERIFIED')
    THEN RAISE(ABORT,'OPENING_CUTOVER_NOT_VERIFIED') END;
END;

CREATE TRIGGER opening_cutover_row_update BEFORE UPDATE ON opening_cutover_rows BEGIN
  SELECT CASE WHEN OLD.operation_id IS NOT NEW.operation_id OR OLD.item_id IS NOT NEW.item_id
    OR OLD.opening_balance_id IS NOT NEW.opening_balance_id OR OLD.sku IS NOT NEW.sku
    OR OLD.ecwid_product_id IS NOT NEW.ecwid_product_id OR OLD.ecwid_combination_id IS NOT NEW.ecwid_combination_id
    OR OLD.ecwid_option_signature IS NOT NEW.ecwid_option_signature OR OLD.physical IS NOT NEW.physical
    OR OLD.unpicked IS NOT NEW.unpicked OR OLD.target_quantity IS NOT NEW.target_quantity
    OR OLD.expected_ecwid_quantity IS NOT NEW.expected_ecwid_quantity OR OLD.source_row IS NOT NEW.source_row
    OR OLD.source_sheet IS NOT NEW.source_sheet THEN RAISE(ABORT,'OPENING_CUTOVER_IMMUTABLE') END;
  SELECT CASE WHEN OLD.alignment_status<>'PENDING' AND (OLD.before_quantity IS NOT NEW.before_quantity
      OR OLD.attempted_at IS NOT NEW.attempted_at)
    THEN RAISE(ABORT,'OPENING_CUTOVER_IMMUTABLE') END;
  -- Direct verification from PENDING is a read-only no-op, never a stock write.
  SELECT CASE WHEN OLD.alignment_status='PENDING' AND NEW.alignment_status='VERIFIED'
      AND (NEW.before_quantity IS NOT NEW.target_quantity OR NEW.after_quantity IS NOT NEW.target_quantity)
    THEN RAISE(ABORT,'OPENING_CUTOVER_STATE_INVALID') END;
  SELECT CASE WHEN NOT EXISTS(SELECT 1 FROM opening_cutover_batches WHERE operation_id=NEW.operation_id AND state='ALIGNING')
    OR NOT ((OLD.alignment_status='PENDING' AND NEW.alignment_status IN ('PROCESSING','VERIFIED','BLOCKED'))
      OR (OLD.alignment_status='PROCESSING' AND NEW.alignment_status IN ('VERIFIED','UNKNOWN','BLOCKED')))
    THEN RAISE(ABORT,'OPENING_CUTOVER_STATE_INVALID') END;
END;
CREATE TRIGGER opening_cutover_row_failure AFTER UPDATE OF alignment_status ON opening_cutover_rows
WHEN NEW.alignment_status IN ('UNKNOWN','BLOCKED') BEGIN
  UPDATE opening_cutover_batches SET state='REVIEW',updated_at=COALESCE(NEW.attempted_at,updated_at)
    WHERE operation_id=NEW.operation_id;
END;
CREATE TRIGGER opening_cutover_batch_no_delete BEFORE DELETE ON opening_cutover_batches BEGIN
  SELECT RAISE(ABORT,'OPENING_CUTOVER_IMMUTABLE');
END;
CREATE TRIGGER opening_cutover_row_no_delete BEFORE DELETE ON opening_cutover_rows BEGIN
  SELECT RAISE(ABORT,'OPENING_CUTOVER_IMMUTABLE');
END;
CREATE TRIGGER opening_cutover_order_no_update BEFORE UPDATE ON opening_cutover_orders BEGIN
  SELECT RAISE(ABORT,'OPENING_CUTOVER_IMMUTABLE');
END;
CREATE TRIGGER opening_cutover_order_no_delete BEFORE DELETE ON opening_cutover_orders BEGIN
  SELECT RAISE(ABORT,'OPENING_CUTOVER_IMMUTABLE');
END;
