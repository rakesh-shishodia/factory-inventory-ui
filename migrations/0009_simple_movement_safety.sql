-- Preserve the worker's selected adjustment reason separately from the signed
-- stock movement, while keeping the established movement types and deltas.
ALTER TABLE movements ADD COLUMN reason_code TEXT NOT NULL DEFAULT ''
  CHECK(reason_code IN ('','ADJUST_UP','ADJUST_DOWN'));

CREATE TRIGGER movement_reason_guard BEFORE INSERT ON movements BEGIN
  SELECT CASE WHEN NEW.reason_code='ADJUST_UP'
      AND (NEW.type<>'RESTOCK' OR length(trim(NEW.note))=0)
    THEN RAISE(ABORT,'INVALID_ADJUSTMENT') END;
  SELECT CASE WHEN NEW.reason_code='ADJUST_DOWN'
      AND (NEW.type<>'INTERNAL_USE' OR length(trim(NEW.note))=0)
    THEN RAISE(ABORT,'INVALID_ADJUSTMENT') END;
END;

-- Keep one stock-limited Ecwid delta in flight per item. This closes the gap
-- between a fresh remote stock read and delivery of the previous local delta.
-- Exact idempotent retries do not insert and therefore do not enter this guard.
CREATE TRIGGER stock_limited_movement_serialization BEFORE INSERT ON movements
WHEN NEW.type<>'ECWID_PICK' AND NEW.inventory_mode='STOCK_LIMITED'
  AND EXISTS(SELECT 1 FROM outbox WHERE item_id=NEW.item_id AND status IN ('PENDING','PROCESSING'))
BEGIN
  SELECT RAISE(ABORT,'STOCK_SYNC_PENDING');
END;
