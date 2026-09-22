-- LOCAL DEMO ONLY, after migration 0002. These numeric Ecwid IDs are fictitious.
-- Never apply this fixture to a live/staging inventory database or use its targets
-- with live sync. Existing simple demo items, movements and balances are untouched.
-- Every INSERT is additive: rerunning preserves picks, payments and quantities.
WITH variants(id,sku,name,combination,signature,available) AS (VALUES
  ('DEMO-VAR-BOLT-M6-20','DEMO-BOLT-M6-20','DEMO bolt M6 × 20 mm','910001','[{"name":"Diameter","value":"M6"},{"name":"Length","value":"20 mm"}]',26),
  ('DEMO-VAR-BOLT-M6-30','DEMO-BOLT-M6-30','DEMO bolt M6 × 30 mm','910002','[{"name":"Diameter","value":"M6"},{"name":"Length","value":"30 mm"}]',45)
)
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,on_hand,last_ecwid_quantity)
SELECT v.id,v.sku,v.name,v.sku,'DEMO · Variation shelf','900001',v.combination,v.signature,0,v.available
FROM variants v WHERE NOT EXISTS (SELECT 1 FROM items old WHERE old.id=v.id);

WITH balances(id,quantity) AS (VALUES ('DEMO-VAR-BOLT-M6-20',30),('DEMO-VAR-BOLT-M6-30',50))
INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
SELECT 'DEMO-VARIATION-OPENING:' || b.id,b.id,b.quantity,'LOCAL VARIATION DEMO — not real stock',
  'demo@local',strftime('%Y-%m-%dT%H:%M:%fZ','now')
FROM balances b WHERE NOT EXISTS (SELECT 1 FROM opening_balances old WHERE old.item_id=b.id);

WITH sample_orders(id,payment) AS (VALUES ('DEMO-VARIATION-PAID','PAID'),('DEMO-VARIATION-AWAITING','AWAITING_PAYMENT'))
INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
SELECT o.id,o.payment,'AWAITING_PROCESSING',strftime('%Y-%m-%dT%H:%M:%fZ','now'),strftime('%Y-%m-%dT%H:%M:%fZ','now')
FROM sample_orders o WHERE NOT EXISTS (SELECT 1 FROM orders old WHERE old.id=o.id);

WITH lines(order_id,line_id,item_id,quantity) AS (VALUES
  ('DEMO-VARIATION-PAID','1','DEMO-VAR-BOLT-M6-20',4),
  ('DEMO-VARIATION-PAID','2','DEMO-VAR-BOLT-M6-30',2),
  ('DEMO-VARIATION-AWAITING','1','DEMO-VAR-BOLT-M6-30',3)
)
INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
SELECT l.order_id || ':' || l.line_id,l.order_id,l.line_id,l.item_id,i.sku,i.name,l.quantity
FROM lines l JOIN items i ON i.id=l.item_id
WHERE NOT EXISTS (SELECT 1 FROM order_lines old WHERE old.id=l.order_id || ':' || l.line_id);
