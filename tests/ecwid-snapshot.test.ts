import { describe, expect, it, vi } from 'vitest';
import { collectPages, fetchReadOnlySnapshot, getOnlyEcwidFetch, pendingOrderReview, renderOrdersReview } from '../scripts/ecwid-snapshot';
import type { EcwidOrder, EcwidPage } from '../src/ecwid';

const fields = {
  products: 'total,count,offset,items(id,sku,name,quantity,unlimited,enabled,options(name,type,choices(text)),combinations(id,sku,quantity,unlimited,options(name,value),compositeParents,compositeComponents),defaultCombinationId,compositeParents,compositeComponents)',
  orders: 'total,count,offset,items(id,paymentStatus,fulfillmentStatus,createTimestamp,updateTimestamp,items(id,productId,sku,name,quantity,combinationId,selectedOptions,digital,trackQuantity))',
};
function listUrl(kind: 'products' | 'orders' = 'products') {
  const url = new URL(`https://app.ecwid.com/api/v3/2442119/${kind}`);
  url.searchParams.set('responseFields', fields[kind]);
  url.searchParams.set('offset', '0');
  url.searchParams.set('limit', '100');
  if (kind === 'orders') url.searchParams.set('createdTo', '1700000000');
  return url;
}
function order(overrides: Partial<EcwidOrder> = {}): EcwidOrder {
  return {
    id: 'ORDER1', paymentStatus: 'PAID', fulfillmentStatus: 'AWAITING_PROCESSING', updatedAt: '2026-09-22T00:00:00Z',
    items: [{ id: 'LINE1', productId: '1001', combinationId: '501', sku: '0008', name: 'Bolt', quantity: 5,
      digital: false, trackQuantity: true, selectedOptions: [{ name: 'Length', value: '20 mm', type: 'CHOICE' }] }],
    ...overrides,
  };
}
function page(ids: string[], total = ids.length, offset = 0): EcwidPage<{ id: string }> {
  return { items: ids.map(id => ({ id })), count: ids.length, total, offset };
}

describe('snapshot transport boundary', () => {
  it.each(['products', 'orders'] as const)('permits only approved %s GET lists and pins redirect policy', async kind => {
    const fetcher = vi.fn().mockResolvedValue(Response.json({}));
    const url = listUrl(kind);
    await getOnlyEcwidFetch('2442119', fetcher)(url, { method: 'GET', redirect: 'follow', headers: { Authorization: 'Bearer test' } });
    expect(fetcher).toHaveBeenCalledExactlyOnceWith(url, expect.objectContaining({ method: 'GET', redirect: 'error' }));
  });

  it.each(['POST', 'PUT', 'DELETE', 'PATCH', 'HEAD', 'get'])('rejects method %s before invoking transport', async method => {
    const fetcher = vi.fn();
    await expect(getOnlyEcwidFetch('2442119', fetcher)(listUrl(), { method })).rejects.toThrow('restricted');
    expect(fetcher).not.toHaveBeenCalled();
  });

  it('rejects request bodies, including a Request whose method is overridden', async () => {
    const fetcher = vi.fn();
    const read = getOnlyEcwidFetch('2442119', fetcher);
    await expect(read(listUrl(), { body: '{}' })).rejects.toThrow('restricted');
    const request = new Request(listUrl(), { method: 'POST', body: '{}' });
    await expect(read(request, { method: 'GET' })).rejects.toThrow('restricted');
    expect(fetcher).not.toHaveBeenCalled();
  });

  it.each(['https://evil.example', 'http://app.ecwid.com', 'https://app.ecwid.com:8443'])('rejects alternate origin %s', async origin => {
    const url = listUrl();
    const fetcher = vi.fn();
    await expect(getOnlyEcwidFetch('2442119', fetcher)(origin + url.pathname + url.search)).rejects.toThrow('restricted');
    expect(fetcher).not.toHaveBeenCalled();
  });

  it.each(['/api/v3/2442119/products/1001', '/api/v3/2442119/products/1001/inventory', '/api/v3/2442119/profile', '/api/v3/999/products'])
    ('rejects unapproved resource %s', async path => {
      const url = listUrl(); url.pathname = path;
      const fetcher = vi.fn();
      await expect(getOnlyEcwidFetch('2442119', fetcher)(url)).rejects.toThrow('restricted');
      expect(fetcher).not.toHaveBeenCalled();
    });

  it.each(['token', 'email', 'paymentStatus', 'fulfillmentStatus', 'sortBy'])('rejects extra query parameter %s', async key => {
    const url = listUrl('orders'); url.searchParams.set(key, 'value');
    await expect(getOnlyEcwidFetch('2442119', vi.fn())(url)).rejects.toThrow('restricted');
  });

  it('rejects PII projections, duplicate parameters, missing cutoff and userinfo/fragment URLs', async () => {
    const urls = Array.from({ length: 6 }, () => listUrl('orders'));
    urls[0].searchParams.set('responseFields', 'items(email,billingPerson,shippingPerson)');
    urls[1].searchParams.append('responseFields', fields.orders);
    urls[2].searchParams.delete('createdTo');
    urls[3].username = 'secret';
    urls[4].hash = '#secret';
    urls[5].searchParams.set('offset', '-1');
    for (const url of urls) await expect(getOnlyEcwidFetch('2442119', vi.fn())(url)).rejects.toThrow('restricted');
  });
});

describe('complete, consistent listing pagination', () => {
  it('uses actual records read for offsets until complete without assuming full pages', async () => {
    const read = vi.fn().mockResolvedValueOnce(page(['a', 'b'], 3)).mockResolvedValueOnce(page(['c'], 3, 2));
    expect(await collectPages(read, 10)).toEqual([{ id: 'a' }, { id: 'b' }, { id: 'c' }]);
    expect(read.mock.calls).toEqual([[0], [2]]);
  });

  it('accepts a zero-result listing', async () => {
    expect(await collectPages(async () => page([]), 0)).toEqual([]);
  });

  it.each([
    { ...page(['a'], 2), total: -1 }, { ...page(['a'], 2), total: 1.5 },
    { ...page(['a'], 2), total: 101 }, { ...page(['a']), count: 2 },
    { ...page(['a']), offset: 1 }, page(['a', 'b'], 1),
    page(Array.from({ length: 101 }, (_, i) => String(i)), 101),
  ])('rejects inconsistent or excessive first-page metadata %#', async invalid => {
    await expect(collectPages(async () => invalid, 100)).rejects.toThrow('inconsistent');
  });

  it('rejects totals changing midway, duplicate IDs, or premature empty pages', async () => {
    for (const second of [page(['b'], 3, 1), page(['a'], 2, 1), page([], 2, 1)]) {
      const read = vi.fn().mockResolvedValueOnce(page(['a'], 2)).mockResolvedValueOnce(second);
      await expect(collectPages(read, 10)).rejects.toThrow();
    }
  });

  it.each(['', '   '])('rejects missing record identity %j', async id => {
    await expect(collectPages(async () => page([id]), 10)).rejects.toThrow('missing record');
  });
});

describe('order reconciliation review', () => {
  it('never treats ordered quantity as confirmed unpicked quantity', () => {
    const [paid, awaiting] = pendingOrderReview([order(), order({ id: 'ORDER2', paymentStatus: 'AWAITING_PAYMENT' })]);
    expect(paid.pickable_after_cutover).toBe(true);
    expect(awaiting.pickable_after_cutover).toBe(false);
    for (const record of [paid, awaiting]) expect(record.lines[0]).toMatchObject({ ordered_quantity: 5,
      previously_picked_quantity: null, unpicked_quantity: null, reconciliation_required: true,
      ecwid_product_id: '1001', ecwid_combination_id: '501' });
  });

  it('does not include terminal fulfillment or cancelled/refunded orders as ordinary pending picks', () => {
    const closed = ['READY_FOR_PICKUP', 'SHIPPED', 'DELIVERED', 'OUT_FOR_DELIVERY', 'RETURNED', 'WILL_NOT_DELIVER']
      .map(fulfillmentStatus => order({ fulfillmentStatus }));
    closed.push(order({ paymentStatus: 'CANCELLED' }), order({ paymentStatus: 'REFUNDED' }));
    expect(pendingOrderReview(closed)).toEqual([]);
  });

  it('excludes picked-and-packed pickup orders without relabelling their status as Delivered', () => {
    const pickup = order({ fulfillmentStatus: 'READY_FOR_PICKUP' });
    const processing = order({ id: 'PENDING', fulfillmentStatus: 'PROCESSING' });
    expect(pendingOrderReview([pickup, processing]).map(record => record.id)).toEqual(['PENDING']);
    expect(pickup.fulfillmentStatus).toBe('READY_FOR_PICKUP');
  });

  it('retains unknown fulfillment for review but never makes it pickable', () => {
    expect(pendingOrderReview([order({ fulfillmentStatus: 'CUSTOM_FULFILLMENT_STATUS_1' })])[0].pickable_after_cutover).toBe(false);
  });

  it.each(['PARTIALLY_REFUNDED', 'INCOMPLETE', 'CUSTOM_PAYMENT_STATUS_1', 'FUTURE_STATUS'])
    ('retains ambiguous %s payment as an unpickable reconciliation exception', paymentStatus => {
      const [reviewed] = pendingOrderReview([order({ paymentStatus })]);
      expect(reviewed).toMatchObject({ payment_status: paymentStatus, pickable_after_cutover: false,
        review_category: 'PAYMENT_STATUS_EXCEPTION' });
      expect(reviewed.lines[0]).toMatchObject({ unpicked_quantity: null, reconciliation_required: true });
    });

  it('removes raw customer text, uploaded-file URLs and unneeded option metadata', () => {
    const original = order();
    original.items[0].selectedOptions = [{ name: 'Inscription', value: 'customer@example.com', type: 'TEXT',
      files: [{ adminUrl: 'https://private.example?token=secret' }] }];
    const reviewed = pendingOrderReview([original]);
    expect(reviewed[0].lines[0]).toMatchObject({ selected_options: null, selected_options_supported: false });
    expect(JSON.stringify(reviewed)).not.toContain('customer@example.com');
    expect(JSON.stringify(reviewed)).not.toContain('private.example');
    original.items[0].selectedOptions = [{ name: 'Length', value: '20 mm', type: 'CHOICE',
      randomPrivateData: 'customer@example.com' }];
    expect(pendingOrderReview([original])[0].lines[0].selected_options).toEqual([{ name: 'Length', value: '20 mm' }]);
  });

  it('escapes all rendered identities and emphasizes unknown unpicked stock', () => {
    const injected = '<script>alert("unsafe")</script>&';
    const original = order({ id: injected });
    original.items[0].sku = injected;
    original.items[0].name = injected;
    const html = renderOrdersReview(pendingOrderReview([original]), injected);
    expect(html).not.toContain(injected);
    expect(html).toContain('&lt;script&gt;alert(&quot;unsafe&quot;)&lt;/script&gt;&amp;');
    expect(html).toContain('NOT confirmed unpicked reservations');
    expect(html).toContain('Needs confirmation');
    expect(html).toContain("default-src 'none'");
    expect(html).not.toContain('<script');
  });
});

describe('full snapshot assembly', () => {
  it('uses approved GET fields, discards PII, and returns review-only stock targets without secrets', async () => {
    const calls: URL[] = [];
    const fetcher = vi.fn(async (input, init?: RequestInit) => {
      const url = new URL(String(input)); calls.push(url);
      expect(init).toMatchObject({ method: 'GET', redirect: 'error' });
      if (url.pathname.endsWith('/products')) return Response.json({ total: 1, offset: 0, count: 1,
        items: [{ id: 1001, sku: 'BASE', name: 'Bolt', unlimited: false, enabled: true, quantity: 999,
          options: [{ name: 'Length', type: 'SELECT', choices: [{ text: '20 mm' }] }],
          combinations: [{ id: 501, sku: '0008', quantity: 7, unlimited: false, options: [{ name: 'Length', value: '20 mm' }] }] }] });
      return Response.json({ total: 1, offset: 0, count: 1, items: [{ id: 'ORDER', email: 'hidden@example.com',
        paymentStatus: 'PAID', fulfillmentStatus: 'PROCESSING', updateTimestamp: 1, createTimestamp: 1,
        items: [{ id: 1, productId: 1001, combinationId: 501, sku: '0008', name: 'Bolt', quantity: 3,
          digital: false, trackQuantity: true, selectedOptions: [{ name: 'Length', value: '20 mm', type: 'CHOICE', ignored: 'private' }] }] }] });
    });
    const snapshot = await fetchReadOnlySnapshot({ storeId: '2442119', token: 'secret_test_only' }, fetcher);
    expect(calls).toHaveLength(2);
    expect(calls[1].searchParams.has('createdTo')).toBe(true);
    expect(snapshot.catalogue).toMatchObject({ dry_run: true, complete: true, product_count: 1, stock_target_count: 1, reservations_confirmed: false });
    expect(snapshot.catalogue.stock_targets[0]).toMatchObject({ combinationId: '501', quantity: 7, sku: '0008' });
    expect(snapshot.orderReview).toMatchObject({ orders_checked: 1, pending_order_count: 1, reservations_confirmed: false });
    expect(snapshot.openingOrders).toMatchObject({ kind: 'READONLY_ORDERS', schema_version: 1, dry_run: true, complete: true,
      store_id: '2442119', orders_checked: 1, pending_order_count: 1, line_count: 1,
      orders: [{ id: 'ORDER', items: [{ id: '1', productId: '1001', combinationId: '501', digital: false,
        trackQuantity: true, selectedOptions: [{ name: 'Length', value: '20 mm' }] }] }] });
    expect(snapshot.openingOrders.creation_cutoff).toBe(Math.floor(Date.parse(snapshot.openingOrders.started_at) / 1000));
    expect(JSON.stringify(snapshot)).not.toContain('secret_test_only');
    expect(JSON.stringify(snapshot)).not.toContain('hidden@example.com');
    expect(JSON.stringify(snapshot)).not.toContain('private');
  });

  it('retains blocking digital evidence and redacts unsupported customer options in opening evidence', async () => {
    const fetcher = vi.fn(async input => new URL(String(input)).pathname.endsWith('/products')
      ? Response.json({ total: 0, count: 0, offset: 0, items: [] })
      : Response.json({ total: 2, count: 2, offset: 0, items: [
        { id: '12', paymentStatus: 'PAID', fulfillmentStatus: 'PROCESSING', updateTimestamp: 1,
          items: [{ id: 1, productId: 1, sku: 'ONE', name: 'Test', quantity: 1, digital: true, trackQuantity: false,
            selectedOptions: [{ name: 'Engraving', value: 'private@example.com', type: 'TEXT',
              files: [{ adminUrl: 'https://private.example/upload' }] }] }] },
        { id: '13', paymentStatus: 'PAID', fulfillmentStatus: 'READY_FOR_PICKUP', updateTimestamp: 1,
          items: [{ id: 2, productId: 2, sku: 'TWO', name: 'Packed test', quantity: 1 }] }
      ] }));
    const result = await fetchReadOnlySnapshot({ storeId: '2442119', token: 'test' }, fetcher);
    expect(result.openingOrders).toMatchObject({ orders_checked: 2, pending_order_count: 1, line_count: 1 });
    expect(result.openingOrders.orders[0].items[0]).toMatchObject({ digital: true, trackQuantity: false,
      selectedOptions: [{ type: 'REDACTED' }] });
    expect(JSON.stringify(result)).not.toContain('private@example.com');
    expect(JSON.stringify(result)).not.toContain('private.example');
  });

  it('fails closed rather than declaring completeness for fractional historical order quantities', async () => {
    const fetcher = vi.fn(async input => new URL(String(input)).pathname.endsWith('/products')
      ? Response.json({ total: 0, count: 0, offset: 0, items: [] })
      : Response.json({ total: 1, count: 1, offset: 0, items: [{ id: 'ORDER', paymentStatus: 'PAID', fulfillmentStatus: 'PROCESSING',
        updateTimestamp: 1, items: [{ id: 1, quantity: 0.5 }] }] }));
    await expect(fetchReadOnlySnapshot({ storeId: '2442119', token: 'test' }, fetcher)).rejects.toThrow('quantity');
  });

  it('rejects wire-level count inconsistencies before producing complete snapshots', async () => {
    const fetcher = vi.fn(async () => Response.json({ total: 0, count: 1, offset: 0, items: [] }));
    await expect(fetchReadOnlySnapshot({ storeId: '2442119', token: 'test' }, fetcher)).rejects.toThrow('page');
  });
});
