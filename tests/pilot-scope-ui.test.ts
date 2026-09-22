import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
const initialization = app.indexOf("$('#movement-form').addEventListener");
if (initialization < 0) throw new Error('The simple UI initialization boundary changed; update this harness.');
const definitions = app.slice(0, initialization);

function evaluate(script: string) {
  const context = createContext({ Intl, document: {}, localStorage: { getItem: () => null }, setTimeout, clearTimeout, AbortController });
  runInContext(definitions, context);
  runInContext(`
    const item={id:'tracked',inventory_mode:'STOCK_LIMITED',active:1,available:8};
    const appLine={id:'app-line',item_id:'tracked',management_mode:'APP',ordered_qty:3,picked_qty:0,pickable_qty:3};
    const workbookLine={id:'workbook-line',item_id:'tracked',management_mode:'WORKBOOK',ordered_qty:9,picked_qty:0,pickable_qty:9};
    const paid={id:'100',payment_status:'PAID',fulfillment_status:'PROCESSING',needs_review:0,lines:[appLine,workbookLine]};
  `, context);
  return runInContext(script, context);
}

describe('simple worker UI pilot scope', () => {
  it('never allows a workbook-managed order line to become pickable', () => {
    expect(evaluate('pickableQuantity(workbookLine)')).toBe(0);
    expect(evaluate('pickableQuantity(appLine)')).toBe(3);
  });

  it.each([
    ['AWAITING_PAYMENT', 'PROCESSING', 0],
    ['PAID', 'READY_FOR_PICKUP', 0],
    ['PAID', 'SHIPPED', 0],
    ['PAID', 'PROCESSING', 1],
  ])('does not offer an ineligible order (%s / %s / review %s)', (payment, fulfillment, needsReview) => {
    expect(evaluate(`canPickOrder({...paid,payment_status:${JSON.stringify(payment)},fulfillment_status:${JSON.stringify(fulfillment)},needs_review:${needsReview}})`)).toBe(false);
  });

  it('accepts only paid processing orders before resolving an exact item line', () => {
    expect(evaluate('canPickOrder(paid)')).toBe(true);
  });

  it('requires the fetched item to be active and its opening count to be known', () => {
    expect(evaluate('operationalItem(item)')).toBe(true);
    expect(evaluate("operationalItem({...item,active:0})")).toBe(false);
    expect(evaluate("operationalItem({...item,inventory_mode:'SUPPLIER_BACKED_UNLIMITED',opening_verified:0})")).toBe(false);
  });
});
