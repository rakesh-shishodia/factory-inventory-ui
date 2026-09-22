import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

const html = readFileSync(new URL('../public/index.html', import.meta.url), 'utf8');
const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
const initialization = app.indexOf("$('#movement-form').addEventListener");
if (initialization < 0) throw new Error('The simple worker initialization boundary changed; update this contract test.');
const definitions = app.slice(0, initialization);

function evaluate(script: string) {
  const context = createContext({
    Intl,
    document: {},
    localStorage: { getItem: () => null },
    setTimeout,
    clearTimeout,
    AbortController,
  });
  runInContext(definitions, context);
  return runInContext(script, context);
}

describe('simple worker movement contract', () => {
  it('keeps the six worker reasons and treats adjustment quantity as a positive magnitude', () => {
    for (const [value, label] of [
      ['ECWID_PICK', 'Ecwid Order Pick (−)'],
      ['EMAIL_SALE', 'Email Sale (−)'],
      ['INTERNAL_USE', 'Internal Use (−)'],
      ['RESTOCK', 'Restock (+)'],
      ['ADJUST_UP', 'Adjust Up (+)'],
      ['ADJUST_DOWN', 'Adjust Down (−)'],
    ]) {
      expect(html).toContain(`<option value="${value}">${label}</option>`);
    }
    expect(evaluate("REASONS.ADJUST_UP.apiType")).toBe('RESTOCK');
    expect(evaluate("REASONS.ADJUST_DOWN.apiType")).toBe('INTERNAL_USE');
    expect(evaluate("REASONS.ADJUST_UP.reasonCode")).toBe('ADJUST_UP');
    expect(evaluate("REASONS.ADJUST_DOWN.reasonCode")).toBe('ADJUST_DOWN');
    expect(evaluate("REASONS.ADJUST_UP.noteRequired && REASONS.ADJUST_DOWN.noteRequired")).toBe(true);
    expect(evaluate("adjustedNote('ADJUST_UP','Cycle count')")).toBe('Adjustment up: Cycle count');
    expect(evaluate("adjustedNote('ADJUST_DOWN','Damaged')")).toBe('Adjustment down: Damaged');
    expect(evaluate('validQuantity(1)')).toBe(true);
    expect(evaluate('validQuantity(0)')).toBe(false);
    expect(evaluate('validQuantity(-1)')).toBe(false);
    expect(evaluate('validQuantity(1.5)')).toBe(false);
  });

  it('associates an Ecwid pick with the exact refreshed order line', () => {
    expect(app).toContain("api(`/api/orders/${encodeURIComponent(raw)}/refresh`, { method: 'POST' })");
    expect(app).toContain("String(line.item_id) === String(state.item.id)");
    expect(app).toContain("line.management_mode !== 'WORKBOOK'");
    expect(app).toContain('if (matches.length > 1)');
    expect(app).toContain('order_id: orderSelection.order.id');
    expect(app).toContain('order_line_id: orderSelection.line.id');
    expect(app).toContain('reason_code: config.reasonCode');
  });

  it('persists one operation ID and verifies the response before clearing a retry', () => {
    expect(app).toContain('operation_id: crypto.randomUUID()');
    expect(app).toContain('localStorage.setItem(PENDING_KEY');
    expect(app).toContain('matchesMovementConfirmation(confirmation, intent.payload)');
    expect(app).toContain('Use Retry safely; do not enter the movement again.');
    expect(app).toContain('already recorded. Nothing was recorded twice.');
  });
});
