import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
const definitions = app.slice(0, app.indexOf("$('#refresh-all').addEventListener"));
function evaluate(script: string) {
  const context = createContext({ Intl, document: {}, localStorage: { getItem: () => null } });
  runInContext(definitions, context);
  runInContext("state.session={mode:'live',inventory_enabled:false,order_sync_enabled:false};", context);
  return runInContext(script, context);
}

describe('pre-cutover workspace guidance', () => {
  it('does not describe a paused empty workspace as ready to pick', () => {
    expect(evaluate('pageHeading()[1]')).toBe('Preparing your inventory.');
    expect(evaluate('pageHeading()[2]')).toContain('do not mean zero factory stock');
    expect(evaluate("state.view='movement';pageHeading()[1]")).toBe('Preparing your inventory.');
  });
  it('distinguishes later maintenance from an unactivated opening import', () => {
    expect(evaluate("state.syncState=[{key:'orders_tracking_started',value:'2026-09-22T08:00:00Z'}];pageHeading()[1]")).toBe('Stock recording is paused.');
    expect(evaluate("state.session.mode='demo';openingImportPending()")).toBe(false);
    expect(evaluate("state.session.inventory_enabled=true;pageHeading()[1]")).toBe('Ready to pick.');
  });
  it('does not claim that an empty order list means all orders are picked', () => {
    expect(app).not.toContain('All caught up.');
    expect(app).toContain('This empty list does not confirm that all factory orders are picked.');
    expect(app).toContain("notImported ? '—'");
  });
});

describe('recent order-feed health', () => {
  function health(entries: Record<string, string>, enabled = true) {
    return evaluate(`state.syncCheckedAt='2026-09-22T10:00:00Z';state.session.order_sync_enabled=${enabled};
      state.syncState=${JSON.stringify(Object.entries({ orders_tracking_started: '2026-09-22T08:00:00Z', ...entries }).map(([key, value]) => ({ key, value })))};
      orderFeedHealth(Date.parse('2026-09-22T10:00:00Z'));`) as { text: string; label: string; note: string; attention: boolean };
  }
  it('does not mistake a historical full poll for recent freshness', () => {
    expect(health({ orders_last_full_poll: '2026-09-22T10:00:00Z' }).attention).toBe(true);
    expect(health({ orders_last_full_poll: '2026-09-22T10:00:00Z' }).text).toBe('First recent check pending');
  });
  it('uses the completed coverage watermark, not the last successful request walltime', () => {
    expect(health({ orders_recent_status: 'CURRENT', orders_recent_watermark: '2026-09-22T09:00:00Z', orders_recent_last_success: '2026-09-22T10:00:00Z' }).attention).toBe(true);
    expect(health({ orders_recent_status: 'CURRENT', orders_recent_watermark: '2026-09-22T09:59:00Z' }).attention).toBe(false);
  });
  it.each(['PENDING', 'ERROR', 'OVERLOADED', 'UNKNOWN'])('never presents %s as a clear feed', status => {
    expect(health({ orders_recent_status: status, orders_recent_watermark: '2026-09-22T09:59:00Z' }).attention).toBe(true);
  });
  it.each(['garbage', '2026-09-22T11:00:00Z'])('flags an invalid or future watermark %s', time => {
    expect(health({ orders_recent_status: 'CURRENT', orders_recent_watermark: time }).attention).toBe(true);
  });
  it('shows paused feed even when a prior completed window is fresh', () => {
    expect(health({ orders_recent_status: 'CURRENT', orders_recent_watermark: '2026-09-22T09:59:00Z' }, false).text).toBe('Order updates paused');
  });
});
