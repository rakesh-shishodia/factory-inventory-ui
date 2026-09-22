-- D1 keeps foreign keys enabled. Rebuild only the parent table inside the
-- migration transaction, preserving IDs and every referencing audit row.
PRAGMA defer_foreign_keys = ON;

DROP VIEW item_stock;
DROP TRIGGER item_codes_insert;
DROP TRIGGER item_codes_update;
DROP TRIGGER opening_balance_guard;
DROP TRIGGER opening_balance_apply;
DROP TRIGGER movement_guard;
DROP TRIGGER movement_apply;

CREATE TABLE items_variation_migration (
  id TEXT PRIMARY KEY,
  sku TEXT NOT NULL COLLATE NOCASE UNIQUE CHECK (length(sku) > 0 AND sku COLLATE BINARY = upper(trim(sku))),
  name TEXT NOT NULL CHECK (length(trim(name)) > 0),
  scan_code TEXT NOT NULL COLLATE NOCASE UNIQUE CHECK (length(scan_code) > 0 AND scan_code COLLATE BINARY = upper(trim(scan_code))),
  location TEXT NOT NULL DEFAULT '',
  ecwid_product_id TEXT,
  ecwid_combination_id TEXT CHECK (ecwid_combination_id IS NULL OR
    (length(ecwid_combination_id)>0 AND ecwid_combination_id NOT GLOB '*[^0-9]*' AND ecwid_combination_id GLOB '[1-9]*')),
  ecwid_option_signature TEXT NOT NULL DEFAULT '[]' CHECK (json_valid(ecwid_option_signature)
    AND json_type(ecwid_option_signature)='array'),
  on_hand INTEGER NOT NULL DEFAULT 0 CHECK (typeof(on_hand) = 'integer' AND on_hand BETWEEN 0 AND 2147483647),
  last_ecwid_quantity INTEGER CHECK (last_ecwid_quantity IS NULL OR typeof(last_ecwid_quantity) = 'integer'),
  active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
  CHECK ((ecwid_combination_id IS NULL AND ecwid_option_signature='[]') OR
    (ecwid_combination_id IS NOT NULL AND ecwid_product_id IS NOT NULL
      AND length(trim(ecwid_product_id))>0 AND json_array_length(ecwid_option_signature)>0))
);

INSERT INTO items_variation_migration(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity,active)
SELECT id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity,active FROM items;
DROP TABLE items;
ALTER TABLE items_variation_migration RENAME TO items;

-- NULL denotes the independently stocked simple parent, never "any variation".
CREATE UNIQUE INDEX items_ecwid_target ON items(ecwid_product_id,coalesce(ecwid_combination_id,''))
  WHERE ecwid_product_id IS NOT NULL;

ALTER TABLE outbox ADD COLUMN ecwid_combination_id TEXT;

CREATE TRIGGER item_codes_insert BEFORE INSERT ON items BEGIN
  SELECT CASE WHEN EXISTS (SELECT 1 FROM items WHERE sku = NEW.scan_code OR scan_code = NEW.sku)
    THEN RAISE(ABORT, 'AMBIGUOUS_ITEM_CODE') END;
END;
CREATE TRIGGER item_codes_update BEFORE UPDATE OF sku, scan_code ON items BEGIN
  SELECT CASE WHEN EXISTS (SELECT 1 FROM items WHERE id <> NEW.id AND (sku = NEW.scan_code OR scan_code = NEW.sku))
    THEN RAISE(ABORT, 'AMBIGUOUS_ITEM_CODE') END;
END;

-- An approved mapping becomes part of the physical ledger's identity. A later
-- mapping correction needs explicit reconciliation, not a pending-write redirect.
CREATE TRIGGER item_mapping_immutable BEFORE UPDATE OF id,sku,ecwid_product_id,ecwid_combination_id,ecwid_option_signature ON items
WHEN (OLD.id IS NOT NEW.id OR OLD.sku IS NOT NEW.sku OR OLD.ecwid_product_id IS NOT NEW.ecwid_product_id
  OR OLD.ecwid_combination_id IS NOT NEW.ecwid_combination_id OR OLD.ecwid_option_signature IS NOT NEW.ecwid_option_signature)
  AND (EXISTS(SELECT 1 FROM opening_balances WHERE item_id=OLD.id)
    OR EXISTS(SELECT 1 FROM movements WHERE item_id=OLD.id)
    OR EXISTS(SELECT 1 FROM order_lines WHERE item_id=OLD.id)
    OR EXISTS(SELECT 1 FROM outbox WHERE item_id=OLD.id))
BEGIN
  SELECT RAISE(ABORT, 'ITEM_MAPPING_IMMUTABLE');
END;

CREATE TRIGGER outbox_target_immutable BEFORE UPDATE OF item_id,ecwid_product_id,ecwid_combination_id,quantity_delta ON outbox
WHEN OLD.item_id IS NOT NEW.item_id OR OLD.ecwid_product_id IS NOT NEW.ecwid_product_id
  OR OLD.ecwid_combination_id IS NOT NEW.ecwid_combination_id OR OLD.quantity_delta IS NOT NEW.quantity_delta
BEGIN
  SELECT RAISE(ABORT, 'OUTBOX_TARGET_IMMUTABLE');
END;

CREATE TRIGGER opening_balance_guard BEFORE INSERT ON opening_balances BEGIN
  SELECT CASE WHEN NOT EXISTS (SELECT 1 FROM items WHERE id = NEW.item_id AND on_hand = 0)
    OR EXISTS (SELECT 1 FROM movements WHERE item_id = NEW.item_id)
    THEN RAISE(ABORT, 'OPENING_BALANCE_ALREADY_STARTED') END;
END;
CREATE TRIGGER opening_balance_apply AFTER INSERT ON opening_balances BEGIN
  UPDATE items SET on_hand = on_hand + NEW.on_hand WHERE id = NEW.item_id;
END;

CREATE VIEW item_stock AS
SELECT i.*, COALESCE(r.reserved, 0) AS reserved, i.on_hand - COALESCE(r.reserved, 0) AS available
FROM items i LEFT JOIN (
  SELECT l.item_id, SUM(l.ordered_qty - l.picked_qty) AS reserved
  FROM order_lines l JOIN orders o ON o.id = l.order_id
  WHERE o.payment_status IN ('PAID', 'AWAITING_PAYMENT') GROUP BY l.item_id
) r ON r.item_id = i.id;

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
  INSERT INTO outbox (id,item_id,ecwid_product_id,ecwid_combination_id,quantity_delta,status,attempts,created_at,updated_at)
    SELECT NEW.id,i.id,i.ecwid_product_id,i.ecwid_combination_id,NEW.ecwid_quantity_delta,'PENDING',0,NEW.created_at,NEW.created_at
    FROM items i WHERE i.id=NEW.item_id AND NEW.ecwid_quantity_delta<>0;
END;

PRAGMA defer_foreign_keys = OFF;
