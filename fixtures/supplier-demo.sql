-- FABRICATED LOCAL DEMO ONLY. Never seed a real inventory database.
-- Known zero opening counts are explicit, not inferred from supplier demand.
-- Rerunning preserves receipts, allocations, picks, payments and deactivation.
WITH sample(id,sku,name,product,combination,signature,mode,supplier) AS (VALUES
  ('SUP-DEMO-ITEM-A','SUP-DEMO-PART-A','DEMO supplier part A','990001',NULL,'[]','SUPPLIER_BACKED_UNLIMITED','DEMO supplier'),
  ('SUP-DEMO-ITEM-B','SUP-DEMO-PART-B','DEMO supplier part B · 20 mm','990002','990003','[{"name":"Length","value":"20 mm"}]','SUPPLIER_BACKED_UNLIMITED','DEMO supplier'),
  ('SUP-DEMO-STOCK','SUP-DEMO-STOCK','DEMO stocked item','990004',NULL,'[]','STOCK_LIMITED','')
)
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,inventory_mode,supplier_name,active)
SELECT s.id,s.sku,s.name,s.sku,'DEMO shelf',s.product,s.combination,s.signature,s.mode,s.supplier,0 FROM sample s
WHERE NOT EXISTS(SELECT 1 FROM items old WHERE old.id=s.id);

WITH counts(id,quantity) AS (VALUES ('SUP-DEMO-ITEM-A',0),('SUP-DEMO-ITEM-B',0),('SUP-DEMO-STOCK',12))
INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
SELECT 'SUP-DEMO-OPENING:'||c.id,c.id,c.quantity,'FABRICATED SUPPLIER DEMO — explicit count, not real inventory',
  'demo@local','2026-09-22T00:00:00.000Z' FROM counts c
WHERE NOT EXISTS(SELECT 1 FROM opening_balances b WHERE b.item_id=c.id);

UPDATE items SET active=1 WHERE id IN ('SUP-DEMO-ITEM-A','SUP-DEMO-ITEM-B','SUP-DEMO-STOCK')
  AND NOT EXISTS(SELECT 1 FROM sync_state WHERE key='supplier_demo_seeded');

WITH sample_orders(id,payment) AS (VALUES ('SUP-DEMO-PAID','PAID'),('SUP-DEMO-AWAITING','AWAITING_PAYMENT'))
INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
SELECT s.id,s.payment,'AWAITING_PROCESSING','2026-09-22T00:00:00.000Z','2026-09-22T00:00:00.000Z'
FROM sample_orders s WHERE NOT EXISTS(SELECT 1 FROM orders old WHERE old.id=s.id);

WITH lines(order_id,line_id,item_id,quantity) AS (VALUES
  ('SUP-DEMO-PAID','1','SUP-DEMO-ITEM-A',20),('SUP-DEMO-PAID','2','SUP-DEMO-STOCK',2),
  ('SUP-DEMO-PAID','3','SUP-DEMO-ITEM-B',8),('SUP-DEMO-AWAITING','1','SUP-DEMO-ITEM-A',5)
)
INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
SELECT l.order_id||':'||l.line_id,l.order_id,l.line_id,l.item_id,i.sku,i.name,l.quantity
FROM lines l JOIN items i ON i.id=l.item_id
WHERE NOT EXISTS(SELECT 1 FROM order_lines old WHERE old.id=l.order_id||':'||l.line_id);

INSERT INTO sync_state(key,value,updated_at) VALUES('supplier_demo_seeded','fabricated-only','2026-09-22T00:00:00.000Z')
ON CONFLICT(key) DO NOTHING;
