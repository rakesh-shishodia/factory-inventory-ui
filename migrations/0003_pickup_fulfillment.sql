-- Store policy: Ready for Pickup means picked and packed, equivalent to Shipped
-- for the picker, not a declaration that the customer has received the order.
DROP TRIGGER movement_guard;
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
      AND o.fulfillment_status IN ('AWAITING_PROCESSING', 'PROCESSING')
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

-- An existing pickup-status order without all physical picks is a discrepancy.
-- Preserve its quantities and reservations; quarantine instead of synthesizing
-- picks, releasing reserved units or sending another Ecwid stock deduction.
UPDATE orders SET needs_review = 1
WHERE fulfillment_status = 'READY_FOR_PICKUP'
  AND EXISTS (SELECT 1 FROM order_lines l WHERE l.order_id = orders.id AND l.picked_qty < l.ordered_qty);

INSERT INTO sync_issues(id,kind,item_id,order_id,message,status,created_at)
SELECT 'order:' || o.id || ':' || coalesce(l.item_id,'unmapped'), 'ORDER_REVIEW', l.item_id, o.id,
  'Ready for Pickup means picked and packed, but recorded picks are incomplete. Reconcile physical stock before continuing.',
  'OPEN', strftime('%Y-%m-%dT%H:%M:%fZ','now')
FROM orders o LEFT JOIN order_lines l ON l.order_id = o.id
WHERE o.fulfillment_status = 'READY_FOR_PICKUP'
  AND EXISTS (SELECT 1 FROM order_lines missing WHERE missing.order_id = o.id AND missing.picked_qty < missing.ordered_qty)
GROUP BY o.id, l.item_id ON CONFLICT(id) DO UPDATE SET status = 'OPEN';
