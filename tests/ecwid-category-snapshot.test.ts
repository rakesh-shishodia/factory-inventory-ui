import { describe, expect, it, vi } from 'vitest';
import { fetchCategorySnapshot } from '../scripts/ecwid-category-snapshot';

const credentials = { storeId: '2442119', token: 'secret_test' };
const category = { id: 1, parentId: 0, name: 'Fasteners', enabled: true };
const product = { id: 2, sku: 'M6', name: 'Bolt', categoryIds: [1], enabled: true };
const response = (items: unknown[], total = items.length, offset = 0) => Response.json({ total, count: items.length, offset, items });

describe('read-only category membership snapshot', () => {
  it('reads only minimal GET projections and keeps parent IDs distinct from stock targets', async () => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValueOnce(response([category])).mockResolvedValueOnce(response([product]));
    const result = await fetchCategorySnapshot(credentials, fetcher);
    expect(result).toMatchObject({ complete: true, dry_run: true, store_id: '2442119', category_count: 1, product_count: 1 });
    expect(result.categories).toEqual([{ ...category, id: '1', parentId: '0' }]);
    expect(result.products).toEqual([{ ...product, id: '2', categoryIds: ['1'] }]);
    for (const [input, init] of fetcher.mock.calls) {
      const url = new URL(String(input));
      expect(url.origin).toBe('https://app.ecwid.com');
      expect(url.searchParams.get('responseFields')).not.toMatch(/quantity|orders|customer|description/);
      if (url.pathname.endsWith('/categories')) expect(url.searchParams.get('hidden_categories')).toBe('true');
      expect(init).toMatchObject({ method: 'GET', redirect: 'error' });
      expect(init?.body).toBeUndefined();
    }
  });
  it('rejects credentials before network access', async () => {
    const fetcher = vi.fn<typeof fetch>();
    await expect(fetchCategorySnapshot({ ...credentials, storeId: '../1' }, fetcher)).rejects.toThrow('Invalid local');
    expect(fetcher).not.toHaveBeenCalled();
  });
  it('accepts omitted parentId only as a root category', async () => {
    const { parentId: _parent, ...root } = category;
    const fetcher = vi.fn<typeof fetch>().mockResolvedValueOnce(response([root])).mockResolvedValueOnce(response([]));
    expect((await fetchCategorySnapshot(credentials, fetcher)).categories[0].parentId).toBe('0');
  });
  it.each([
    { ...product, categoryIds: undefined }, { ...product, categoryIds: [1, 1] },
    { ...product, categoryIds: [null] }, { ...product, id: 1.5 }, { ...product, enabled: 'true' }
  ])('rejects malformed product identity or category coverage', async invalid => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValueOnce(response([category])).mockResolvedValueOnce(response([invalid]));
    await expect(fetchCategorySnapshot(credentials, fetcher)).rejects.toThrow('invalid data');
  });
  it('rejects repeated page identities', async () => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(response([category, category]));
    await expect(fetchCategorySnapshot(credentials, fetcher)).rejects.toThrow('Duplicate');
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
  it('does not leak transport errors or retry an uncertain read', async () => {
    const fetcher = vi.fn<typeof fetch>().mockRejectedValue(new Error('secret_sensitive'));
    await expect(fetchCategorySnapshot(credentials, fetcher)).rejects.toThrow('No automatic retry');
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
  it('handles HTTP rejections without returning response bodies', async () => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(new Response('secret_sensitive', { status: 403 }));
    await expect(fetchCategorySnapshot(credentials, fetcher)).rejects.toThrow('Category snapshot categories HTTP 403.');
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
});
