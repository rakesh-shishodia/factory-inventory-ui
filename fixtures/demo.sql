-- Local sample data only. Re-running preserves existing picks and balances.
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
SELECT 'item-nut-m3','NUT-M3','M3 hex nut','NUT-M3','A1 · Bin 04','10001',0,188
WHERE NOT EXISTS (SELECT 1 FROM items WHERE id='item-nut-m3');
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
SELECT 'item-bolt-m3','BOLT-M3-12','M3 × 12 socket head bolt','BOLT-M3-12','A1 · Bin 06','10002',0,147
WHERE NOT EXISTS (SELECT 1 FROM items WHERE id='item-bolt-m3');
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
SELECT 'item-bearing','BEARING-608','608ZZ ball bearing','BEARING-608','A2 · Bin 02','10003',0,46
WHERE NOT EXISTS (SELECT 1 FROM items WHERE id='item-bearing');
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
SELECT 'item-shaft','SHAFT-8-300','8 mm smooth shaft · 300 mm','SHAFT-8-300','B1 · Shelf 01','10004',0,34
WHERE NOT EXISTS (SELECT 1 FROM items WHERE id='item-shaft');
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
SELECT 'item-fan','FAN-4010','40 mm cooling fan · 24 V','FAN-4010','B2 · Bin 03','10005',0,11
WHERE NOT EXISTS (SELECT 1 FROM items WHERE id='item-fan');
INSERT INTO items(id,sku,name,scan_code,location,ecwid_product_id,on_hand,last_ecwid_quantity)
SELECT 'item-coupler','COUPLER-5-8','Flexible shaft coupler · 5 × 8 mm','COUPLER-5-8','B1 · Bin 05','10006',0,24
WHERE NOT EXISTS (SELECT 1 FROM items WHERE id='item-coupler');

WITH balances(id,qty) AS (VALUES ('item-nut-m3',200),('item-bolt-m3',160),('item-bearing',48),('item-shaft',35),('item-fan',12),('item-coupler',24))
INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
SELECT 'demo-opening:' || b.id,b.id,b.qty,'LOCAL DEMO — not real stock','demo@local',strftime('%Y-%m-%dT%H:%M:%fZ','now')
FROM balances b WHERE NOT EXISTS (SELECT 1 FROM opening_balances o WHERE o.item_id=b.id);

INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
SELECT 'DEMO-1001','PAID','AWAITING_PROCESSING',strftime('%Y-%m-%dT%H:%M:%fZ','now'),strftime('%Y-%m-%dT%H:%M:%fZ','now')
WHERE NOT EXISTS (SELECT 1 FROM orders WHERE id='DEMO-1001');
INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
SELECT 'DEMO-1002','PAID','PROCESSING',strftime('%Y-%m-%dT%H:%M:%fZ','now'),strftime('%Y-%m-%dT%H:%M:%fZ','now')
WHERE NOT EXISTS (SELECT 1 FROM orders WHERE id='DEMO-1002');
INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
SELECT 'DEMO-1003','AWAITING_PAYMENT','AWAITING_PROCESSING',strftime('%Y-%m-%dT%H:%M:%fZ','now'),strftime('%Y-%m-%dT%H:%M:%fZ','now')
WHERE NOT EXISTS (SELECT 1 FROM orders WHERE id='DEMO-1003');

WITH lines(order_id,line_id,item_id,qty) AS (VALUES
  ('DEMO-1001','1','item-nut-m3',12),('DEMO-1001','2','item-bolt-m3',8),
  ('DEMO-1002','1','item-bearing',2),('DEMO-1002','2','item-shaft',1),
  ('DEMO-1003','1','item-bolt-m3',5),('DEMO-1003','2','item-fan',1))
INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
SELECT l.order_id || ':' || l.line_id,l.order_id,l.line_id,l.item_id,i.sku,i.name,l.qty
FROM lines l JOIN items i ON i.id=l.item_id
WHERE NOT EXISTS (SELECT 1 FROM order_lines old WHERE old.id=l.order_id || ':' || l.line_id);
