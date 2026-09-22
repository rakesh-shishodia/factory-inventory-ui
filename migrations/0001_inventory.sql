PRAGMA foreign_keys = ON;

CREATE TABLE items (
  id TEXT PRIMARY KEY,
  sku TEXT NOT NULL COLLATE NOCASE UNIQUE CHECK (length(sku) > 0 AND sku COLLATE BINARY = upper(trim(sku))),
  name TEXT NOT NULL CHECK (length(trim(name)) > 0),
  scan_code TEXT NOT NULL COLLATE NOCASE UNIQUE CHECK (length(scan_code) > 0 AND scan_code COLLATE BINARY = upper(trim(scan_code))),
  location TEXT NOT NULL DEFAULT '',
  ecwid_product_id TEXT UNIQUE,
  on_hand INTEGER NOT NULL DEFAULT 0 CHECK (typeof(on_hand) = 'integer' AND on_hand BETWEEN 0 AND 2147483647),
  last_ecwid_quantity INTEGER CHECK (last_ecwid_quantity IS NULL OR typeof(last_ecwid_quantity) = 'integer'),
  active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1))
);

-- The picker accepts either a SKU or a QR code. Their namespaces must not point
-- to different items or a scan could silently return the wrong stock record.
CREATE TRIGGER item_codes_insert BEFORE INSERT ON items BEGIN
  SELECT CASE WHEN EXISTS (SELECT 1 FROM items WHERE sku = NEW.scan_code OR scan_code = NEW.sku)
    THEN RAISE(ABORT, 'AMBIGUOUS_ITEM_CODE') END;
END;
CREATE TRIGGER item_codes_update BEFORE UPDATE OF sku, scan_code ON items BEGIN
  SELECT CASE WHEN EXISTS (SELECT 1 FROM items WHERE id <> NEW.id AND (sku = NEW.scan_code OR scan_code = NEW.sku))
    THEN RAISE(ABORT, 'AMBIGUOUS_ITEM_CODE') END;
END;

CREATE TABLE orders (
  id TEXT PRIMARY KEY,
  payment_status TEXT NOT NULL,
  fulfillment_status TEXT NOT NULL DEFAULT 'AWAITING_PROCESSING',
  remote_updated_at TEXT NOT NULL,
  remote_lines_hash TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL,
  needs_review INTEGER NOT NULL DEFAULT 0 CHECK (needs_review IN (0, 1))
);

CREATE TABLE order_lines (
  id TEXT PRIMARY KEY,
  order_id TEXT NOT NULL REFERENCES orders(id),
  ecwid_line_id TEXT NOT NULL,
  item_id TEXT REFERENCES items(id),
  sku TEXT NOT NULL,
  name TEXT NOT NULL,
  ordered_qty INTEGER NOT NULL CHECK (typeof(ordered_qty) = 'integer' AND ordered_qty > 0),
  picked_qty INTEGER NOT NULL DEFAULT 0 CHECK (typeof(picked_qty) = 'integer' AND picked_qty BETWEEN 0 AND ordered_qty),
  UNIQUE (order_id, ecwid_line_id)
);
CREATE INDEX order_lines_item ON order_lines(item_id);
CREATE INDEX orders_status ON orders(payment_status, fulfillment_status, needs_review);

CREATE TABLE movements (
  id TEXT PRIMARY KEY,
  fingerprint TEXT NOT NULL,
  type TEXT NOT NULL CHECK (type IN ('ECWID_PICK', 'EMAIL_SALE', 'INTERNAL_USE', 'RESTOCK')),
  item_id TEXT NOT NULL REFERENCES items(id),
  quantity INTEGER NOT NULL CHECK (typeof(quantity) = 'integer' AND quantity BETWEEN 1 AND 1000000),
  quantity_delta INTEGER NOT NULL,
  ecwid_quantity_delta INTEGER NOT NULL,
  order_id TEXT REFERENCES orders(id),
  order_line_id TEXT REFERENCES order_lines(id),
  note TEXT NOT NULL DEFAULT '',
  actor TEXT NOT NULL,
  created_at TEXT NOT NULL,
  CHECK (quantity_delta = CASE WHEN type = 'RESTOCK' THEN quantity ELSE -quantity END),
  CHECK (ecwid_quantity_delta = CASE WHEN type = 'ECWID_PICK' THEN 0 WHEN type = 'RESTOCK' THEN quantity ELSE -quantity END),
  CHECK ((type = 'ECWID_PICK' AND order_id IS NOT NULL AND order_line_id IS NOT NULL)
      OR (type <> 'ECWID_PICK' AND order_id IS NULL AND order_line_id IS NULL))
);
CREATE INDEX movements_item_created ON movements(item_id, created_at DESC);
CREATE INDEX movements_created ON movements(created_at DESC);

CREATE TABLE opening_balances (
  id TEXT PRIMARY KEY,
  item_id TEXT NOT NULL UNIQUE REFERENCES items(id),
  on_hand INTEGER NOT NULL CHECK (typeof(on_hand) = 'integer' AND on_hand BETWEEN 0 AND 2147483647),
  source_ref TEXT NOT NULL CHECK (length(trim(source_ref)) > 0),
  actor TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TRIGGER opening_balance_guard BEFORE INSERT ON opening_balances BEGIN
  SELECT CASE WHEN NOT EXISTS (SELECT 1 FROM items WHERE id = NEW.item_id AND on_hand = 0)
    OR EXISTS (SELECT 1 FROM movements WHERE item_id = NEW.item_id)
    THEN RAISE(ABORT, 'OPENING_BALANCE_ALREADY_STARTED') END;
END;
CREATE TRIGGER opening_balance_apply AFTER INSERT ON opening_balances BEGIN
  UPDATE items SET on_hand = on_hand + NEW.on_hand WHERE id = NEW.item_id;
END;
CREATE TRIGGER opening_balance_no_update BEFORE UPDATE ON opening_balances BEGIN
  SELECT RAISE(ABORT, 'OPENING_BALANCE_IMMUTABLE');
END;
CREATE TRIGGER opening_balance_no_delete BEFORE DELETE ON opening_balances BEGIN
  SELECT RAISE(ABORT, 'OPENING_BALANCE_IMMUTABLE');
END;

CREATE TABLE outbox (
  id TEXT PRIMARY KEY REFERENCES movements(id),
  item_id TEXT NOT NULL REFERENCES items(id),
  ecwid_product_id TEXT NOT NULL,
  quantity_delta INTEGER NOT NULL CHECK (typeof(quantity_delta) = 'integer' AND quantity_delta <> 0),
  status TEXT NOT NULL DEFAULT 'PENDING' CHECK (status IN ('PENDING', 'PROCESSING', 'APPLIED', 'UNKNOWN', 'BLOCKED')),
  attempts INTEGER NOT NULL DEFAULT 0,
  last_error TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE INDEX outbox_status_created ON outbox(status, created_at);
CREATE INDEX outbox_item_status ON outbox(item_id, status);

CREATE TABLE webhook_events (
  event_id TEXT PRIMARY KEY,
  event_type TEXT NOT NULL,
  entity_id TEXT NOT NULL,
  store_id TEXT NOT NULL,
  payload TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'PENDING' CHECK (status IN ('PENDING', 'PROCESSING', 'APPLIED', 'BLOCKED')),
  attempts INTEGER NOT NULL DEFAULT 0,
  last_error TEXT,
  received_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  processed_at TEXT
);
CREATE INDEX webhook_events_status_received ON webhook_events(status, received_at);

CREATE TABLE sync_issues (
  id TEXT PRIMARY KEY,
  item_id TEXT REFERENCES items(id),
  order_id TEXT REFERENCES orders(id),
  kind TEXT NOT NULL,
  message TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'OPEN' CHECK (status IN ('OPEN', 'RESOLVED')),
  created_at TEXT NOT NULL,
  resolved_at TEXT
);
CREATE INDEX sync_issues_item_status ON sync_issues(item_id, status);

CREATE TABLE sync_state (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

-- Physical balance and reservation are different facts. Review orders still reserve
-- their last known lines until the discrepancy has been resolved by an operator.
CREATE VIEW item_stock AS
SELECT i.*,
       COALESCE(r.reserved, 0) AS reserved,
       i.on_hand - COALESCE(r.reserved, 0) AS available
FROM items i
LEFT JOIN (
  SELECT l.item_id, SUM(l.ordered_qty - l.picked_qty) AS reserved
  FROM order_lines l JOIN orders o ON o.id = l.order_id
  WHERE o.payment_status IN ('PAID', 'AWAITING_PAYMENT')
  GROUP BY l.item_id
) r ON r.item_id = i.id;

-- These guards execute in the same SQLite write as the balance update. A pair of
-- simultaneous requests cannot both spend the same stock or pick the same line.
CREATE TRIGGER movement_guard BEFORE INSERT ON movements BEGIN
  SELECT CASE WHEN NOT EXISTS (SELECT 1 FROM items WHERE id = NEW.item_id AND active = 1)
    THEN RAISE(ABORT, 'ITEM_UNAVAILABLE') END;
  SELECT CASE WHEN EXISTS (SELECT 1 FROM sync_issues WHERE item_id = NEW.item_id AND status = 'OPEN')
    OR EXISTS (SELECT 1 FROM outbox WHERE item_id = NEW.item_id AND status IN ('UNKNOWN', 'BLOCKED'))
    THEN RAISE(ABORT, 'ITEM_NEEDS_REVIEW') END;
  SELECT CASE WHEN NEW.type <> 'ECWID_PICK' AND NOT EXISTS (
    SELECT 1 FROM items WHERE id = NEW.item_id AND length(trim(ecwid_product_id)) > 0
  ) THEN RAISE(ABORT, 'ECWID_MAPPING_REQUIRED') END;
  SELECT CASE WHEN NEW.type = 'ECWID_PICK' AND NOT EXISTS (
    SELECT 1 FROM orders o JOIN order_lines l ON l.order_id = o.id
    WHERE o.id = NEW.order_id AND l.id = NEW.order_line_id AND l.item_id = NEW.item_id
      AND o.payment_status = 'PAID' AND o.needs_review = 0
      AND o.fulfillment_status IN ('AWAITING_PROCESSING', 'PROCESSING', 'READY_FOR_PICKUP')
  ) THEN RAISE(ABORT, 'ORDER_NOT_PICKABLE') END;
  SELECT CASE WHEN NEW.type = 'ECWID_PICK' AND NEW.quantity > (
    SELECT ordered_qty - picked_qty FROM order_lines WHERE id = NEW.order_line_id
  ) THEN RAISE(ABORT, 'PICK_QUANTITY_EXCEEDED') END;
  SELECT CASE WHEN NEW.type <> 'RESTOCK' AND NEW.quantity > (
    SELECT on_hand FROM items WHERE id = NEW.item_id
  ) THEN RAISE(ABORT, 'INSUFFICIENT_PHYSICAL_STOCK') END;
  SELECT CASE WHEN NEW.type IN ('EMAIL_SALE', 'INTERNAL_USE') AND NEW.quantity > (
    SELECT available FROM item_stock WHERE id = NEW.item_id
  ) THEN RAISE(ABORT, 'INSUFFICIENT_AVAILABLE_STOCK') END;
END;

CREATE TRIGGER movement_apply AFTER INSERT ON movements BEGIN
  UPDATE items SET on_hand = on_hand + NEW.quantity_delta WHERE id = NEW.item_id;
  UPDATE order_lines SET picked_qty = picked_qty + NEW.quantity
    WHERE NEW.type = 'ECWID_PICK' AND id = NEW.order_line_id;
  INSERT INTO outbox (id, item_id, ecwid_product_id, quantity_delta, status, attempts, created_at, updated_at)
    SELECT NEW.id, i.id, i.ecwid_product_id, NEW.ecwid_quantity_delta, 'PENDING', 0, NEW.created_at, NEW.created_at
    FROM items i WHERE i.id = NEW.item_id AND NEW.ecwid_quantity_delta <> 0;
END;

CREATE TRIGGER movement_no_update BEFORE UPDATE ON movements BEGIN
  SELECT RAISE(ABORT, 'MOVEMENT_IMMUTABLE');
END;
CREATE TRIGGER movement_no_delete BEFORE DELETE ON movements BEGIN
  SELECT RAISE(ABORT, 'MOVEMENT_IMMUTABLE');
END;
