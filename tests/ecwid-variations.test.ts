import { describe, expect, it, vi } from 'vitest';
import { canonicalVariationOptions, EcwidClient, flattenProduct, parseOrder, productStockTargets } from '../src/ecwid';

function product(overrides: Record<string, unknown> = {}) {
  return {
    id: 1001, sku: '0008', name: 'Socket bolt', quantity: 999, unlimited: false, enabled: true,
    defaultCombinationId: 501, compositeParents: [], compositeComponents: [],
    options: [{ name: 'Length', type: 'SELECT', choices: [{ text: '20 mm' }, { text: '30 mm' }] }],
    combinations: [
      { id: 501, sku: '0008', quantity: 8, unlimited: false, options: [{ name: 'Length', value: '20 mm' }] },
      { id: 502, sku: '0009', quantity: 9, unlimited: false, options: [{ name: 'Length', value: '30 mm' }] },
    ],
    ...overrides,
  };
}

function order(selectedOptions: unknown) {
  return { id: 'ORDER', paymentStatus: 'PAID', fulfillmentStatus: 'AWAITING_PROCESSING', updateTimestamp: 1,
    items: [{ id: 1, productId: 1001, quantity: 1, combinationId: 501, selectedOptions }] };
}

describe('Ecwid variation stock normalization', () => {
  it('retains exact target IDs, leading-zero SKUs, and separate quantities under one parent', () => {
    const [parent, small, large] = flattenProduct(product());
    expect(parent).toMatchObject({ id: '1001', combinationId: null, hasVariations: true, quantity: 999 });
    expect(small).toMatchObject({ id: '1001', combinationId: '501', sku: '0008', quantity: 8,
      unlimited: false, enabled: true, hasExtraOptions: false, hasBundleRelationships: false, eligibilityVerified: true,
      variationOptions: [{ name: 'Length', value: '20 mm' }] });
    expect(large).toMatchObject({ id: '1001', combinationId: '502', sku: '0009', quantity: 9 });
    expect(productStockTargets(parent)).toEqual([small, large]);
    expect(productStockTargets(small)).toEqual([small]);
  });

  it('supports a simple item as a distinct base stock target', () => {
    const [simple] = flattenProduct(product({ options: [], combinations: [], defaultCombinationId: 0 }));
    expect(simple).toMatchObject({ combinationId: null, hasOptions: false, hasVariations: false,
      hasExtraOptions: false, eligibilityVerified: true });
    expect(productStockTargets(simple)).toEqual([simple]);
  });

  it('never inherits parent stock or SKU into an untracked variation', () => {
    const raw = product({ combinations: [{ id: 501, options: [{ name: 'Length', value: '20 mm' }] }] });
    const [, variation] = flattenProduct(raw);
    expect(variation).toMatchObject({ sku: '', quantity: null, unlimited: true, eligibilityVerified: false });
  });

  it.each([undefined, {}, null, 'Length'])('does not treat missing or malformed parent options %s as safe empty', options => {
    const [parent, variation] = flattenProduct(product({ options }));
    expect(parent.eligibilityVerified).toBe(false);
    expect(variation).toMatchObject({ eligibilityVerified: false, hasExtraOptions: true });
  });

  it.each([null, {}, '501'])('rejects malformed combinations %s', combinations => {
    expect(() => flattenProduct(product({ combinations }))).toThrow('variations');
  });

  it('does not make a parent with missing variation details an eligible base item', () => {
    const [parent] = flattenProduct(product({ combinations: undefined }));
    expect(parent).toMatchObject({ hasVariations: true, eligibilityVerified: false });
    expect(productStockTargets(parent)).toEqual([]);
  });

  it.each(['TEXTFIELD', 'CHECKBOX', 'FILES', 'unknown'])('blocks unsupported extra option type %s', type => {
    const raw = product();
    raw.options.push({ name: 'Add-on', type, choices: [{ text: 'Yes' }] });
    expect(flattenProduct(raw)[1]).toMatchObject({ hasExtraOptions: true, eligibilityVerified: false });
  });

  it('blocks an add-on not covered by the chosen variation even when it is a dropdown', () => {
    const raw = product();
    raw.options.push({ name: 'Add washer', type: 'SELECT', choices: [{ text: 'Yes' }, { text: 'No' }] });
    expect(flattenProduct(raw)[1].hasExtraOptions).toBe(true);
  });

  it('requires the variation option value to exist in the parent choices', () => {
    const raw = product();
    raw.combinations[0].options[0].value = '80 mm';
    expect(flattenProduct(raw)[1].hasExtraOptions).toBe(true);
  });

  it.each(['compositeParents', 'compositeComponents'])('blocks a nonempty %s relationship for every variation', key => {
    const rows = flattenProduct(product({ [key]: [123] }));
    expect(rows.every(row => row.hasBundleRelationships)).toBe(true);
  });

  it.each([{}, null, ''])('does not accept malformed optional bundle metadata %s', value => {
    expect(flattenProduct(product({ compositeComponents: value }))[1].hasBundleRelationships).toBe(true);
  });

  it('rejects duplicate variation IDs or duplicate variation option identities', () => {
    const raw = product();
    raw.combinations[1].id = 501;
    expect(() => flattenProduct(raw)).toThrow('Duplicate Ecwid variation ID');
    const other = product();
    other.combinations[1].options = other.combinations[0].options;
    expect(() => flattenProduct(other)).toThrow('Duplicate Ecwid variation option identity');
  });

  it('keeps negative stock visible for eligibility checks but rejects fractional stock', () => {
    const raw = product();
    raw.combinations[0].quantity = -2;
    expect(flattenProduct(raw)[1].quantity).toBe(-2);
    raw.combinations[0].quantity = 1.5;
    expect(() => flattenProduct(raw)).toThrow('variation quantity');
  });
});

describe('canonical option identity', () => {
  it('is deterministic without losing case or whitespace or adding price modifiers to stock identity', () => {
    expect(canonicalVariationOptions([{ name: 'Size', value: ' M8 ' }, { name: 'Finish', value: 'Steel', type: 'CHOICE',
      valuesArray: ['Steel'], files: null, selections: [{ selectionTitle: 'Steel', selectionModifier: 10, selectionModifierType: 'ABSOLUTE' }] }]))
      .toEqual([{ name: 'Finish', value: 'Steel' }, { name: 'Size', value: ' M8 ' }]);
    expect(canonicalVariationOptions([])).toEqual([]);
  });

  it.each([undefined, null, {}, [null], [{ name: 'Size', value: 8 }], [{ name: '', value: 'M8' }],
    [{ name: 'Size', value: 'M8' }, { name: 'Size', value: 'M9' }],
    [{ name: 'Size', value: 'M8', type: 'CHOICES' }], [{ name: 'Size', value: 'M8', type: 'TEXT' }],
    [{ name: 'Size', value: 'M8', valuesArray: ['M8', 'M9'] }], [{ name: 'Size', value: 'M8', valuesArray: ['M9'] }],
    [{ name: 'Size', value: 'M8', valuesArray: {} }], [{ name: 'Size', value: 'M8', files: ['file'] }],
    [{ name: 'Size', value: 'M8', selections: [{ selectionTitle: 'M9' }] }],
    [{ name: 'Size', value: 'M8', selections: { selectionTitle: 'M8' } }]])('fails closed for malformed or complex selected options %#', value => {
    expect(canonicalVariationOptions(value)).toBeNull();
  });

  it('preserves the order variation ID and its original selected options', () => {
    const selected = [{ name: 'Length', value: '20 mm', type: 'CHOICE', valuesArray: ['20 mm'] }];
    expect(parseOrder(order(selected)).items[0]).toMatchObject({ combinationId: '501', selectedOptions: selected });
  });

  it.each([null, {}, 'Length'])('does not convert malformed order option container %s into a simple order', value => {
    expect(() => parseOrder(order(value))).toThrow('selected options');
  });
});

describe('Ecwid variation transport', () => {
  it('lists parent products with full normalized variation and safety metadata', async () => {
    const fetcher = vi.fn().mockResolvedValue(Response.json({ total: 1, offset: 0, items: [product()] }));
    const result = await new EcwidClient({ storeId: '123', token: 'test' }, fetcher).listProducts();
    expect(result.count).toBe(1);
    expect(result.items[0].variations).toHaveLength(2);
    const url = new URL(String(fetcher.mock.calls[0][0]));
    expect(url.searchParams.get('responseFields')).toContain('options(name,type,choices(text))');
    expect(url.searchParams.get('responseFields')).toContain('combinations(id,sku,quantity,unlimited,options(name,value),compositeParents,compositeComponents)');
    expect(url.searchParams.get('responseFields')).not.toContain('imageUrl');
    expect(fetcher.mock.calls[0][1]).toMatchObject({ method: 'GET', redirect: 'error' });
  });

  it('refreshes one exact variation from a fresh parent snapshot without quantity fallback', async () => {
    const fetcher = vi.fn().mockResolvedValue(Response.json(product()));
    const client = new EcwidClient({ storeId: '123', token: 'test' }, fetcher);
    expect(await client.getProductStock('1001', '502')).toMatchObject({ combinationId: '502', quantity: 9 });
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(String(fetcher.mock.calls[0][0])).toContain('/products/1001?responseFields=');
  });

  it('returns all sibling stock targets from one read for a parent webhook', async () => {
    const fetcher = vi.fn().mockResolvedValue(Response.json(product()));
    expect((await new EcwidClient({ storeId: '123', token: 'test' }, fetcher).getProductStockTargets('1001'))
      .map(row => row.combinationId)).toEqual(['501', '502']);
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('rejects a missing variation and never silently returns parent stock', async () => {
    const client = new EcwidClient({ storeId: '123', token: 'test' }, vi.fn().mockResolvedValue(Response.json(product())));
    await expect(client.getProductStock('1001', '999')).rejects.toMatchObject({ outcome: 'REJECTED' });
  });

  it('rejects a response for a different parent', async () => {
    const client = new EcwidClient({ storeId: '123', token: 'test' }, vi.fn().mockResolvedValue(Response.json(product({ id: 999 }))));
    await expect(client.getProductStock('1001', '501')).rejects.toMatchObject({ outcome: 'REJECTED' });
  });

  it('writes a delta only to the exact combination inventory endpoint', async () => {
    const fetcher = vi.fn().mockResolvedValue(Response.json({ updateCount: 1 }));
    await new EcwidClient({ storeId: '123', token: 'test' }, fetcher).adjustStock('1001', -3, '501');
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(fetcher.mock.calls[0]).toEqual(['https://app.ecwid.com/api/v3/123/products/1001/combinations/501/inventory',
      expect.objectContaining({ method: 'PUT', body: '{"quantityDelta":-3}', redirect: 'error' })]);
  });

  it.each(['', '0', '-1', '../502', '501?token=oops'])('rejects malformed variation target %s before any request', async id => {
    const fetcher = vi.fn();
    const client = new EcwidClient({ storeId: '123', token: 'test' }, fetcher);
    await expect(client.adjustStock('1001', 1, id)).rejects.toMatchObject({ outcome: 'REJECTED' });
    await expect(client.getProductStock('1001', id)).rejects.toMatchObject({ outcome: 'REJECTED' });
    expect(fetcher).not.toHaveBeenCalled();
  });

  it.each([500, 502, 503])('never retries an uncertain variation write after HTTP %s', async status => {
    const fetcher = vi.fn().mockResolvedValue(new Response('response may contain secrets', { status }));
    await expect(new EcwidClient({ storeId: '123', token: 'test' }, fetcher).adjustStock('1001', 1, '501'))
      .rejects.toMatchObject({ outcome: 'UNKNOWN', message: `Ecwid returned HTTP ${status}.` });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
});
