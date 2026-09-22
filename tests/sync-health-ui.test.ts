import { readFileSync } from 'node:fs';
import { URL } from 'node:url';
import { describe, expect, it } from 'vitest';

const html = readFileSync(new URL('../public/index.html', import.meta.url), 'utf8');
const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');

describe('reduced worker-facing UI', () => {
  it('ships one stock form instead of worker navigation and dashboards', () => {
    expect(html.match(/<form\b/g)).toHaveLength(1);
    expect(html).not.toContain('Main navigation');
    expect(html).not.toContain('Activity & sync');
    expect(html).not.toContain('Your inventory');
    expect(html).not.toContain('Order picking');
  });

  it('keeps the requested large, linear controls on the one screen', () => {
    for (const id of ['open-scanner', 'sku', 'fetch-item', 'quantity-minus', 'quantity', 'quantity-plus', 'reason', 'order-id', 'fetch-order', 'order-result', 'notes', 'submit-movement']) {
      expect(html).toContain(`id="${id}"`);
    }
    expect(html).toContain('Use front camera');
    expect(html).toContain('Close camera');
  });

  it('continues using the authenticated same-origin session', () => {
    expect(app).toContain("credentials: 'same-origin'");
    expect(app).toContain("api('/api/session')");
  });

  it('allows camera and SKU checks while paused but keeps all write controls disabled', () => {
    expect(app).toContain("$('#open-scanner').disabled = locked;");
    expect(app).toContain("$('#fetch-item').disabled = locked || !$('#sku').value.trim();");
    expect(app).toContain("$('#submit-movement').disabled = locked || recordingPaused");
    expect(app).toContain("$('#fetch-order').disabled = locked || recordingPaused");
  });

  it('refreshes one manually entered order instead of relying on a background order list', () => {
    expect(app).toContain("api(`/api/orders/${encodeURIComponent(raw)}/refresh`, { method: 'POST' })");
    expect(app).not.toContain("api('/api/orders?status=PAID')");
  });

  it('cleans up the scanner when the page hides or unloads', () => {
    expect(app).toContain("document.addEventListener('visibilitychange'");
    expect(app).toContain("window.addEventListener('pagehide'");
    expect(app).toContain('await scanner.stop()');
  });
});
