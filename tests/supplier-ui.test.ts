import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { beforeEach, describe, expect, it } from 'vitest';

const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
const initialization = app.indexOf("$('#movement-form').addEventListener");
if (initialization < 0) throw new Error('The simple UI initialization boundary changed; update this harness.');
const definitions = app.slice(0, initialization);
let saved: string | null;
let evaluate: (script: string) => unknown;

beforeEach(() => {
  saved = null;
  const context = createContext({
    Intl,
    document: {},
    localStorage: { getItem: () => saved },
    setTimeout,
    clearTimeout,
    AbortController,
  });
  runInContext(definitions, context);
  evaluate = script => runInContext(script, context);
});

describe('simple UI reason mapping', () => {
  it('uses existing immutable movement types for adjustments', () => {
    expect(evaluate("reasonConfig('ADJUST_UP').apiType")).toBe('RESTOCK');
    expect(evaluate("reasonConfig('ADJUST_DOWN').apiType")).toBe('INTERNAL_USE');
    expect(evaluate("reasonConfig('ADJUST_UP').reasonCode")).toBe('ADJUST_UP');
    expect(evaluate("reasonConfig('ADJUST_DOWN').reasonCode")).toBe('ADJUST_DOWN');
    expect(evaluate("reasonConfig('ADJUST_UP').noteRequired")).toBe(true);
    expect(evaluate("reasonConfig('ADJUST_DOWN').noteRequired")).toBe(true);
  });

  it('prefixes adjustment notes so their ledger meaning remains explicit', () => {
    expect(evaluate("adjustedNote('ADJUST_UP','Counted two extra')")).toBe('Adjustment up: Counted two extra');
    expect(evaluate("adjustedNote('ADJUST_DOWN','Damaged')")).toBe('Adjustment down: Damaged');
    expect(evaluate("adjustedNote('EMAIL_SALE','INV-42')")).toBe('INV-42');
  });
});

describe('supplier and saved-operation safeguards', () => {
  it('requires supplier opening verification and explicit allocation for picking', () => {
    expect(evaluate("operationalItem({inventory_mode:'SUPPLIER_BACKED_UNLIMITED',opening_verified:1,active:1})")).toBe(true);
    expect(evaluate("operationalItem({inventory_mode:'SUPPLIER_BACKED_UNLIMITED',opening_verified:0,active:1})")).toBe(false);
    expect(evaluate("pickableQuantity({inventory_mode:'SUPPLIER_BACKED_UNLIMITED',management_mode:'APP',remaining_qty:8,allocated_qty:3})")).toBe(3);
    expect(evaluate("pickableQuantity({inventory_mode:'SUPPLIER_BACKED_UNLIMITED',management_mode:'WORKBOOK',remaining_qty:8,allocated_qty:8})")).toBe(0);
  });

  it.each(['ALLOCATE', 'RELEASE', 'RESTOCK', 'ECWID_PICK'])('retains a saved %s retry with its original operation ID', type => {
    saved = JSON.stringify({ payload: { operation_id: '88888888-8888-4888-8888-888888888888', item_id: 'item', order_id: '9440', order_line_id: 'line-1', type } });
    expect(evaluate('readPending().payload.operation_id')).toBe('88888888-8888-4888-8888-888888888888');
  });

  it.each(['PICK', 'CANCEL_RELEASE', 'constructor', '__proto__'])('rejects unsupported saved operation %s', type => {
    saved = JSON.stringify({ payload: { operation_id: '88888888-8888-4888-8888-888888888888', item_id: 'item', type } });
    expect(evaluate('readPending()')).toBeNull();
  });

  it('requires an exact movement confirmation before clearing a saved retry', () => {
    evaluate(`const payload={operation_id:'88888888-8888-4888-8888-888888888888',type:'ECWID_PICK',item_id:'item',quantity:2,order_id:'9440',order_line_id:'line-1',note:''};
      const confirmation={...payload,id:payload.operation_id};
      const adjustment={operation_id:'99999999-9999-4999-8999-999999999999',type:'RESTOCK',reason_code:'ADJUST_UP',item_id:'item',quantity:2,note:'Adjustment up: Counted'};
      const adjustedConfirmation={...adjustment,id:adjustment.operation_id};`);
    expect(evaluate('matchesMovementConfirmation(confirmation,payload)')).toBe(true);
    expect(evaluate("matchesMovementConfirmation({...confirmation,order_line_id:'line-2'},payload)")).toBe(false);
    expect(evaluate("matchesMovementConfirmation({...confirmation,quantity:3},payload)")).toBe(false);
    expect(evaluate('matchesMovementConfirmation(adjustedConfirmation,adjustment)')).toBe(true);
    expect(evaluate("matchesMovementConfirmation({...adjustedConfirmation,reason_code:''},adjustment)")).toBe(false);
  });
});
