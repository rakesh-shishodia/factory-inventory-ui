import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

const app=readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const definitions=app.slice(0,app.indexOf("$('#refresh-all').addEventListener"));
function evaluate(script:string) {
  const context=createContext({Intl,document:{},localStorage:{getItem:()=>null}});
  runInContext(definitions,context);
  runInContext(`const external={management_mode:'WORKBOOK',ordered_qty:3,picked_qty:0,pickable_qty:99};
    const tracked={management_mode:'APP',ordered_qty:2,picked_qty:0};
    const mixed={payment_status:'PAID',fulfillment_status:'PROCESSING',needs_review:0,lines:[tracked,external]};`,context);
  return runInContext(script,context);
}
describe('workbook-managed mixed order UI',()=>{
  it('does not let workbook quantities enter app pick progress or eligibility',()=>{
    expect(evaluate('canPick(mixed)')).toBe(true);
    expect(evaluate('appOrderRemaining(mixed)')).toBe(2);
    expect(evaluate('orderRemaining(mixed)')).toBe(5);
    expect(evaluate('pickableQuantity(external)')).toBe(0);
    expect(evaluate('canPick({...mixed,lines:[external]})')).toBe(false);
    expect(evaluate('canPick({...mixed,lines:[{...tracked,picked_qty:2},external]})')).toBe(false);
  });
  it('continues to honor Paid and review restrictions',()=>{
    expect(evaluate("canPick({...mixed,payment_status:'AWAITING_PAYMENT'})")).toBe(false);
    expect(evaluate('canPick({...mixed,needs_review:1})')).toBe(false);
  });
  it('explains workbook handling and never represents app progress as whole-order completion',()=>{
    expect(app).toContain('Workbook-managed · handle outside this app');
    expect(app).toContain('App progress does not confirm that the whole order is complete.');
    expect(app).toContain('Workbook-managed lines still need an outside check; this app cannot confirm whole-order completion.');
  });
});
