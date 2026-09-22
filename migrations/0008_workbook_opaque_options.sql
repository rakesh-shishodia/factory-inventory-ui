-- Reviewed workbook-only variants may inherit their parent's display SKU and
-- carry cutting instructions which this application must not store or fulfil.
-- Preserve immutable rows and order-line FK identities while removing only the
-- old globally-unique workbook SKU constraint. App SKU constraints are unchanged.
PRAGMA defer_foreign_keys = ON;
DROP TRIGGER workbook_target_guard;
DROP TRIGGER workbook_target_no_update;
DROP TRIGGER workbook_target_no_delete;
DROP TRIGGER item_workbook_conflict_insert;
DROP TRIGGER item_workbook_conflict_update;
DROP TRIGGER workbook_order_line_insert;

CREATE TABLE workbook_managed_targets_v2 (
  id TEXT PRIMARY KEY,
  ecwid_product_id TEXT NOT NULL CHECK(length(trim(ecwid_product_id))>0),
  ecwid_combination_id TEXT,
  ecwid_option_signature TEXT NOT NULL CHECK(json_valid(ecwid_option_signature)),
  sku TEXT NOT NULL COLLATE NOCASE CHECK(length(trim(sku))>0),
  name TEXT NOT NULL,
  review_reference TEXT NOT NULL CHECK(length(trim(review_reference))>0),
  reviewed_by TEXT NOT NULL CHECK(length(trim(reviewed_by))>0),
  reviewed_at TEXT NOT NULL,
  sku_source TEXT NOT NULL DEFAULT 'TARGET' CHECK(sku_source IN ('TARGET','PARENT_IF_VARIATION_BLANK')),
  option_policy TEXT NOT NULL DEFAULT 'EXACT' CHECK(option_policy IN ('EXACT','STOCK_SELECTION_PLUS_OPAQUE_EXTRAS')),
  CHECK((sku_source='TARGET' AND option_policy='EXACT') OR
    (sku_source IN ('TARGET','PARENT_IF_VARIATION_BLANK') AND option_policy='STOCK_SELECTION_PLUS_OPAQUE_EXTRAS'
      AND ecwid_combination_id IS NOT NULL AND length(ecwid_combination_id)>0
      AND json_type(ecwid_option_signature)='array' AND json_array_length(ecwid_option_signature)>0))
);
INSERT INTO workbook_managed_targets_v2(id,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,sku,name,review_reference,reviewed_by,reviewed_at)
SELECT id,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,sku,name,review_reference,reviewed_by,reviewed_at
FROM workbook_managed_targets;
DROP TABLE workbook_managed_targets;
ALTER TABLE workbook_managed_targets_v2 RENAME TO workbook_managed_targets;
CREATE UNIQUE INDEX workbook_target_identity ON workbook_managed_targets(ecwid_product_id,COALESCE(ecwid_combination_id,''));
CREATE INDEX workbook_target_sku ON workbook_managed_targets(sku);

CREATE TRIGGER workbook_target_guard BEFORE INSERT ON workbook_managed_targets BEGIN
  SELECT CASE WHEN EXISTS(SELECT 1 FROM items i WHERE i.sku=NEW.sku COLLATE NOCASE
    OR (i.ecwid_product_id=NEW.ecwid_product_id AND i.ecwid_combination_id IS NEW.ecwid_combination_id))
    THEN RAISE(ABORT,'WORKBOOK_APP_IDENTITY_CONFLICT') END;
  SELECT CASE WHEN EXISTS(SELECT 1 FROM workbook_managed_targets w WHERE w.id=NEW.id AND (
    w.ecwid_product_id IS NOT NEW.ecwid_product_id OR w.ecwid_combination_id IS NOT NEW.ecwid_combination_id
    OR w.ecwid_option_signature IS NOT NEW.ecwid_option_signature OR w.sku IS NOT NEW.sku COLLATE NOCASE
    OR w.name IS NOT NEW.name OR w.review_reference IS NOT NEW.review_reference OR w.reviewed_by IS NOT NEW.reviewed_by
    OR w.sku_source IS NOT NEW.sku_source OR w.option_policy IS NOT NEW.option_policy))
    THEN RAISE(ABORT,'WORKBOOK_TARGET_CONFLICT') END;
  -- A repeated display SKU is allowed only for explicitly reviewed independent
  -- blank-SKU variants of the SAME parent. No fuzzy or cross-parent fallback.
  SELECT CASE WHEN EXISTS(SELECT 1 FROM workbook_managed_targets w WHERE w.id<>NEW.id AND w.sku=NEW.sku COLLATE NOCASE
    AND NOT (w.sku_source='PARENT_IF_VARIATION_BLANK' AND NEW.sku_source='PARENT_IF_VARIATION_BLANK'
      AND w.option_policy='STOCK_SELECTION_PLUS_OPAQUE_EXTRAS' AND NEW.option_policy='STOCK_SELECTION_PLUS_OPAQUE_EXTRAS'
      AND w.ecwid_product_id=NEW.ecwid_product_id AND w.ecwid_combination_id IS NOT NEW.ecwid_combination_id))
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
CREATE TRIGGER workbook_order_line_insert BEFORE INSERT ON order_lines BEGIN
  SELECT CASE WHEN (NEW.management_mode='WORKBOOK' AND (NEW.item_id IS NOT NULL OR NEW.picked_qty<>0
    OR NOT EXISTS(SELECT 1 FROM workbook_managed_targets w WHERE w.id=NEW.workbook_target_id AND w.sku=NEW.sku COLLATE NOCASE)))
    OR (NEW.management_mode='APP' AND NEW.workbook_target_id IS NOT NULL)
    THEN RAISE(ABORT,'INVALID_WORKBOOK_ORDER_LINE') END;
END;
PRAGMA defer_foreign_keys = OFF;
