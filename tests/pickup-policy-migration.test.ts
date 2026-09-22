import Database from 'better-sqlite3';
import { readFileSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';
import { describe, expect, it } from 'vitest';

describe('pickup fulfillment policy forward migration', () => {
  it('preserves the ledger and reservations while quarantining existing incomplete pickup orders', async () => {
    const sqlite = new Database(':memory:');
    try {
      sqlite.pragma('foreign_keys = ON');
      const migrate = (file: string) => sqlite.transaction(() => sqlite.exec(
        readFileSync(new NodeURL(`../migrations/${file}`, import.meta.url), 'utf8')))();
      migrate('0001_inventory.sql');
      migrate('0002_variation_inventory.sql');
      for (const [id, picked] of [['partial', 1], ['packed', 3], ['processing', 0]] as const) {
        sqlite.prepare(`INSERT INTO items(id,sku,name,scan_code,ecwid_product_id)
          VALUES(?,?,?,?,?)`).run(id, id.toUpperCase(), id, id.toUpperCase(), `10${picked}`);
        sqlite.prepare(`INSERT INTO opening_balances(id,item_id,on_hand,source_ref,actor,created_at)
          VALUES(?,?,10,'approved-sheet','admin','2026-09-22T00:00:00.000Z')`).run(`opening-${id}`, id);
        sqlite.prepare(`INSERT INTO orders(id,payment_status,fulfillment_status,remote_updated_at,updated_at)
          VALUES(?,'PAID','PROCESSING','2026-09-22T00:00:00.000Z','2026-09-22T00:00:00.000Z')`).run(id);
        sqlite.prepare(`INSERT INTO order_lines(id,order_id,ecwid_line_id,item_id,sku,name,ordered_qty)
          VALUES(?,?,'1',?,?,?,3)`).run(`line-${id}`, id, id, id.toUpperCase(), id);
        if (picked) sqlite.prepare(`INSERT INTO movements
          (id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,order_id,order_line_id,actor,created_at)
          VALUES(?,?,'ECWID_PICK',?,?,-?,0,?,?,'picker','2026-09-22T00:00:00.000Z')`)
          .run(`pick-${id}`,`pick-${id}`,id,picked,picked,id,`line-${id}`);
      }
      sqlite.exec("UPDATE orders SET fulfillment_status='READY_FOR_PICKUP' WHERE id IN ('partial','packed')");
      const ledgerBefore = ['items', 'opening_balances', 'order_lines', 'movements', 'outbox']
        .map(table => sqlite.prepare(`SELECT * FROM ${table} ORDER BY rowid`).all());
      migrate('0003_pickup_fulfillment.sql');
      const ledgerAfter = ['items', 'opening_balances', 'order_lines', 'movements', 'outbox']
        .map(table => sqlite.prepare(`SELECT * FROM ${table} ORDER BY rowid`).all());
      expect(ledgerAfter).toEqual(ledgerBefore);
      expect(sqlite.pragma('foreign_key_check')).toEqual([]);
      expect(sqlite.prepare('SELECT id,needs_review,fulfillment_status FROM orders ORDER BY id').all()).toEqual([
        { id: 'packed', needs_review: 0, fulfillment_status: 'READY_FOR_PICKUP' },
        { id: 'partial', needs_review: 1, fulfillment_status: 'READY_FOR_PICKUP' },
        { id: 'processing', needs_review: 0, fulfillment_status: 'PROCESSING' },
      ]);
      expect(sqlite.prepare('SELECT item_id,order_id,status FROM sync_issues').all())
        .toEqual([{ item_id: 'partial', order_id: 'partial', status: 'OPEN' }]);
      expect(sqlite.prepare("SELECT * FROM item_stock WHERE id='partial'").get()).toMatchObject({ on_hand: 9, reserved: 2, available: 7 });
      expect(sqlite.prepare("SELECT * FROM item_stock WHERE id='packed'").get()).toMatchObject({ on_hand: 7, reserved: 0, available: 7 });
      expect(sqlite.prepare(`SELECT id FROM orders WHERE payment_status='PAID' AND needs_review=0
        AND fulfillment_status IN ('AWAITING_PROCESSING','PROCESSING')`).all()).toEqual([{id:'processing'}]);
      expect(() => sqlite.exec(`INSERT INTO movements
        (id,fingerprint,type,item_id,quantity,quantity_delta,ecwid_quantity_delta,order_id,order_line_id,actor,created_at)
        VALUES('blocked','blocked','ECWID_PICK','partial',1,-1,0,'partial','line-partial','picker','2026-09-22T00:00:00.000Z')`))
        .toThrow('ITEM_NEEDS_REVIEW');
    } finally { sqlite.close(); }
  });
});
