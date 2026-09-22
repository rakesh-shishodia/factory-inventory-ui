import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { beforeEach, describe, expect, it } from 'vitest';

// Exercise the actual shipped UI helpers without a browser or network. The
// initialization footer is excluded; helper and state definitions stay intact.
const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
const initialization = app.indexOf("$('#refresh-all').addEventListener");
if (initialization < 0) throw new Error('The UI initialization boundary changed; update this harness.');
const definitions = app.slice(0, initialization);
let saved: string | null;
let evaluate: (script: string) => unknown;

beforeEach(() => {
  saved = null;
  const context = createContext({ Intl, document: {}, localStorage: { getItem: () => saved } });
  runInContext(definitions, context);
  evaluate = script => runInContext(script, context);
  evaluate(`
    const paidOrder = {payment_status:'PAID',fulfillment_status:'PROCESSING',needs_review:0,lines:[{ordered_qty:20,picked_qty:0}]};
    const supplierLine = {inventory_mode:'SUPPLIER_BACKED_UNLIMITED',opening_verified:1,item_active:1,remaining_qty:20,allocated_qty:8,unallocated_qty:12,free_qty:0};
    const activeSupplier = {inventory_mode:'SUPPLIER_BACKED_UNLIMITED',opening_verified:1,active:1};
  `);
});

describe('supplier UI operation eligibility', () => {
  it('keeps paid picking and awaiting-payment assignment separate', () => {
    expect(evaluate('canPick(paidOrder)')).toBe(true);
    expect(evaluate("canPick({...paidOrder,payment_status:'AWAITING_PAYMENT'})")).toBe(false);
    expect(evaluate("canAllocate({...paidOrder,payment_status:'AWAITING_PAYMENT'})")).toBe(true);
  });

  it.each(['READY_FOR_PICKUP', 'SHIPPED', 'DELIVERED', 'CANCELLED'])('never opens picking or assignment for %s', status => {
    expect(evaluate(`canPick({...paidOrder,fulfillment_status:${JSON.stringify(status)}})`)).toBe(false);
    expect(evaluate(`canAllocate({...paidOrder,fulfillment_status:${JSON.stringify(status)}})`)).toBe(false);
  });

  it('preserves whole-order review holds', () => {
    expect(evaluate('canPick({...paidOrder,needs_review:1})')).toBe(false);
    expect(evaluate('canAllocate({...paidOrder,needs_review:1})')).toBe(false);
  });

  it('caps picks at assigned stock and respects backend eligibility', () => {
    expect(evaluate('pickableQuantity(supplierLine)')).toBe(8);
    expect(evaluate('pickableQuantity({...supplierLine,pickable_qty:4})')).toBe(4);
    expect(evaluate('pickableQuantity({...supplierLine,remaining_qty:3})')).toBe(3);
    expect(evaluate("pickableQuantity({...supplierLine,fulfillment_state:'REVIEW',pickable_qty:8})")).toBe(0);
    expect(evaluate('pickableQuantity({...supplierLine,pickable_qty:0})')).toBe(0);
  });

  it('requires both a verified supplier opening and active item', () => {
    expect(evaluate('operationalItem(activeSupplier)')).toBe(true);
    expect(evaluate('operationalItem({...activeSupplier,opening_verified:0})')).toBe(false);
    expect(evaluate('operationalItem({...activeSupplier,opening_verified:null})')).toBe(false);
    expect(evaluate('operationalItem({...activeSupplier,active:0})')).toBe(false);
    expect(evaluate("operationalItem({inventory_mode:'STOCK_LIMITED',active:1})")).toBe(true);
    expect(evaluate("operationalItem({inventory_mode:'STOCK_LIMITED',active:0})")).toBe(false);
  });

  it('never treats an unknown opening as a confirmed zero or pickable quantity', () => {
    expect(evaluate('unknownOpening({...activeSupplier,opening_verified:0,on_hand:0})')).toBe(true);
    expect(evaluate('knownValue(null)')).toBe('—');
    expect(evaluate('knownValue(undefined)')).toBe('—');
    expect(evaluate('knownValue(0)')).toBe('0');
    expect(evaluate('pickableQuantity({...supplierLine,opening_verified:0,pickable_qty:8})')).toBe(0);
    expect(evaluate('pickableQuantity({...supplierLine,item_active:0,pickable_qty:8})')).toBe(0);
    expect(evaluate('supplierLineEligible({...supplierLine,opening_verified:0})')).toBe(false);
  });

  it('distinguishes free received stock awaiting assignment from supplier shortage', () => {
    expect(evaluate('supplierLineStatus({...supplierLine,free_qty:25})')).toBe('Free stock available — assign explicitly');
    expect(evaluate('supplierLineStatus(supplierLine)')).toBe('Unassigned demand — awaiting supplier');
    expect(evaluate('supplierLineStatus({...supplierLine,opening_verified:0})')).toBe('Opening count required');
    expect(evaluate('supplierLineStatus({...supplierLine,item_active:0})')).toBe('Inactive — awaiting activation');
  });

  it('shows cancelled or closed lines as closed, not awaiting supplier, while retaining review holds', () => {
    expect(evaluate("supplierLineStatus({...supplierLine,fulfillment_state:'CLOSED',allocated_qty:0})")).toBe('Closed — no picking');
    expect(evaluate("supplierLineStatus({...supplierLine,fulfillment_state:'CLOSED'})")).toBe('Needs review — assignment retained');
    expect(evaluate("supplierLineStatus({...supplierLine,fulfillment_state:'REVIEW',item_active:0})")).toBe('Needs review');
    expect(evaluate("supplierLineEligible({...supplierLine,fulfillment_state:'CLOSED'})")).toBe(false);
    expect(evaluate("pickableQuantity({...supplierLine,fulfillment_state:'CLOSED',pickable_qty:8})")).toBe(0);
  });

  it('qualifies mixed-order picking guidance with the Paid requirement', () => {
    expect(app).toContain('On Paid orders, other eligible lines may be picked.');
    expect(app).not.toContain('Other eligible lines can be picked now');
  });
});

describe('saved operation and allocation confirmation', () => {
  const operation = { operation_id: '88888888-8888-4888-8888-888888888888', item_id: 'supplier-item',
    order_id: 'paid-order', order_line_id: 'line-1', quantity: 2, note: '', type: 'ALLOCATE' };

  it.each(['ALLOCATE', 'RELEASE', 'RESTOCK', 'ECWID_PICK'])('retains a saved %s with the original operation ID', type => {
    saved = JSON.stringify({ payload: { ...operation, type } });
    expect(evaluate('readPending().payload.operation_id')).toBe(operation.operation_id);
    expect(evaluate('readPending().payload.type')).toBe(type);
  });

  it.each(['PICK', 'CANCEL_RELEASE', 'constructor', '__proto__'])('does not accept %s as a user-submitted saved operation', type => {
    saved = JSON.stringify({ payload: { ...operation, type } });
    expect(evaluate('readPending()')).toBeNull();
  });

  it('requires exact order references on saved assignments', () => {
    saved = JSON.stringify({ payload: { ...operation, order_line_id: undefined } });
    expect(evaluate('readPending()')).toBeNull();
    saved = '{invalid JSON';
    expect(evaluate('readPending()')).toBeNull();
  });

  it('accepts only a matching allocation response, so ambiguous responses retain the saved retry', () => {
    evaluate(`const request=${JSON.stringify(operation)}; const confirmation={...request,id:request.operation_id};`);
    expect(evaluate('matchesAllocationConfirmation(confirmation,request)')).toBe(true);
    expect(evaluate('Boolean(matchesAllocationConfirmation(null,request))')).toBe(false);
    for (const changes of [{ id: 'wrong' }, { type: 'RELEASE' }, { item_id: 'wrong' }, { order_id: 'wrong' },
      { order_line_id: 'wrong' }, { quantity: 3 }, { note: 'changed' }]) {
      expect(evaluate(`matchesAllocationConfirmation({...confirmation,...${JSON.stringify(changes)}},request)`)).toBe(false);
    }
  });
});
