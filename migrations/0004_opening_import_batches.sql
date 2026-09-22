-- DB-only opening staging audit. Staged items are inactive and quarantined;
-- no reservation or Ecwid adjustment is implied by one of these records.
CREATE TABLE opening_import_batches (
  operation_id TEXT PRIMARY KEY,
  fingerprint TEXT NOT NULL CHECK (length(fingerprint)=64),
  store_id TEXT NOT NULL CHECK (length(store_id)>0 AND store_id NOT GLOB '*[^0-9]*' AND store_id GLOB '[1-9]*'),
  status TEXT NOT NULL DEFAULT 'STAGED' CHECK (status='STAGED'),
  preview_hash TEXT NOT NULL CHECK (length(preview_hash)=64),
  source_hash TEXT NOT NULL CHECK (length(source_hash)=64),
  catalogue_hash TEXT NOT NULL CHECK (length(catalogue_hash)=64),
  source_ref TEXT NOT NULL CHECK (length(trim(source_ref))>0),
  scope_json TEXT NOT NULL CHECK (json_valid(scope_json) AND json_type(scope_json)='array'),
  row_count INTEGER NOT NULL CHECK (typeof(row_count)='integer' AND row_count BETWEEN 1 AND 200
    AND json_array_length(scope_json)=row_count),
  actor TEXT NOT NULL CHECK (length(trim(actor))>0),
  created_at TEXT NOT NULL
);

CREATE TABLE opening_import_rows (
  operation_id TEXT NOT NULL REFERENCES opening_import_batches(operation_id),
  item_id TEXT NOT NULL UNIQUE REFERENCES items(id),
  opening_balance_id TEXT NOT NULL UNIQUE REFERENCES opening_balances(id),
  sku TEXT NOT NULL,
  ecwid_product_id TEXT NOT NULL,
  ecwid_combination_id TEXT,
  ecwid_option_signature TEXT NOT NULL,
  on_hand INTEGER NOT NULL CHECK (typeof(on_hand)='integer' AND on_hand BETWEEN 0 AND 2147483647),
  ecwid_quantity INTEGER NOT NULL CHECK (typeof(ecwid_quantity)='integer' AND ecwid_quantity BETWEEN 0 AND 2147483647),
  source_row INTEGER CHECK (source_row IS NULL OR (typeof(source_row)='integer' AND source_row>0)),
  source_sheet TEXT,
  PRIMARY KEY (operation_id,item_id)
);

-- Ecwid IDs are unique within a store, not across stores. Item records use the
-- application's one configured store, so configuration changes cannot turn a
-- database into a mixture of stock targets belonging to different shops.
CREATE TRIGGER opening_import_store_guard BEFORE INSERT ON opening_import_batches BEGIN
  SELECT CASE WHEN EXISTS (SELECT 1 FROM opening_import_batches WHERE store_id<>NEW.store_id)
    THEN RAISE(ABORT,'OPENING_DATABASE_STORE_MISMATCH') END;
END;

CREATE TRIGGER opening_import_row_guard BEFORE INSERT ON opening_import_rows BEGIN
  SELECT CASE WHEN NOT EXISTS (
    SELECT 1 FROM items i JOIN opening_balances o ON o.item_id=i.id
      JOIN opening_import_batches b ON b.operation_id=NEW.operation_id
    WHERE i.id=NEW.item_id AND o.id=NEW.opening_balance_id AND i.active=0
      AND i.sku=NEW.sku COLLATE BINARY AND i.ecwid_product_id=NEW.ecwid_product_id
      AND i.ecwid_combination_id IS NEW.ecwid_combination_id AND i.ecwid_option_signature=NEW.ecwid_option_signature
      AND i.on_hand=NEW.on_hand AND o.on_hand=NEW.on_hand AND i.last_ecwid_quantity=NEW.ecwid_quantity
      AND o.source_ref=b.source_ref AND o.actor=b.actor AND o.created_at=b.created_at
  ) THEN RAISE(ABORT,'OPENING_IMPORT_AUDIT_MISMATCH') END;
END;

CREATE TRIGGER opening_import_batch_no_update BEFORE UPDATE ON opening_import_batches BEGIN
  SELECT RAISE(ABORT,'OPENING_IMPORT_IMMUTABLE');
END;
CREATE TRIGGER opening_import_batch_no_delete BEFORE DELETE ON opening_import_batches BEGIN
  SELECT RAISE(ABORT,'OPENING_IMPORT_IMMUTABLE');
END;
CREATE TRIGGER opening_import_row_no_update BEFORE UPDATE ON opening_import_rows BEGIN
  SELECT RAISE(ABORT,'OPENING_IMPORT_IMMUTABLE');
END;
CREATE TRIGGER opening_import_row_no_delete BEFORE DELETE ON opening_import_rows BEGIN
  SELECT RAISE(ABORT,'OPENING_IMPORT_IMMUTABLE');
END;
