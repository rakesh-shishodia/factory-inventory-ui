import { mkdir, realpath, stat, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { readBoundedJson, type EcwidFetch } from '../src/ecwid';
import { collectPages } from './ecwid-snapshot';
import { loadReadOnlyCredentials } from './ecwid-readonly';

type Category = { id: string; parentId: string; name: string; enabled: boolean };
type ProductCategories = { id: string; name: string; sku: string; categoryIds: string[]; enabled: boolean };
const fields = {
  categories: 'total,count,offset,items(id,parentId,name,enabled)',
  products: 'total,count,offset,items(id,name,sku,categoryIds,enabled)'
};
function numericId(value: unknown, zero = false): string {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < (zero ? 0 : 1)) throw new Error('Invalid category/product ID.');
  return String(value);
}
function object(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Invalid category listing.');
  return value as Record<string, unknown>;
}

/** Category membership only. No order/customer data, stock writes, or redirects. */
export async function fetchCategorySnapshot(credentials: { storeId: string; token: string }, fetcher: EcwidFetch = fetch) {
  if (!/^[1-9]\d{0,19}$/.test(credentials.storeId) || !/^secret_[A-Za-z0-9_-]+$/.test(credentials.token)) throw new Error('Invalid local Ecwid credentials.');
  const started = new Date().toISOString();
  async function page<T extends { id: string }>(resource: keyof typeof fields, offset: number, convert: (row: Record<string, unknown>) => T) {
    const url = new URL(`https://app.ecwid.com/api/v3/${credentials.storeId}/${resource}`);
    url.searchParams.set('limit', '100'); url.searchParams.set('offset', String(offset)); url.searchParams.set('responseFields', fields[resource]);
    if (resource === 'categories') url.searchParams.set('hidden_categories', 'true');
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 15_000);
    try {
      const response = await fetcher(url, { method: 'GET', redirect: 'error', signal: controller.signal,
        headers: { Authorization: `Bearer ${credentials.token}`, Accept: 'application/json' } });
      if (!response.ok) { await response.body?.cancel(); throw new Error(`Category snapshot ${resource} HTTP ${response.status}.`); }
      const body = object(await readBoundedJson(response, 2_000_000));
      if (!Array.isArray(body.items) || !Number.isSafeInteger(body.total) || !Number.isSafeInteger(body.count) || !Number.isSafeInteger(body.offset)) throw new Error('Invalid category listing.');
      return { total: body.total as number, count: body.count as number, offset: body.offset as number, items: body.items.map(row => convert(object(row))) };
    } catch (error) {
      if (error instanceof Error && /^Category snapshot (categories|products) HTTP \d+\.$/.test(error.message)) throw error;
      throw new Error(`Category snapshot ${resource} failed (network, timeout, redirect or invalid data). No automatic retry was made.`);
    } finally { clearTimeout(timer); }
  }
  const categories = await collectPages<Category>(offset => page('categories', offset, row => {
    if (typeof row.name !== 'string' || !row.name.trim() || typeof row.enabled !== 'boolean') throw new Error('Invalid category.');
    return { id: numericId(row.id), parentId: numericId(row.parentId ?? 0, true), name: row.name, enabled: row.enabled };
  }), 10_000);
  const products = await collectPages<ProductCategories>(offset => page('products', offset, row => {
    if (typeof row.name !== 'string' || typeof row.sku !== 'string' || typeof row.enabled !== 'boolean' || !Array.isArray(row.categoryIds)) throw new Error('Invalid product categories.');
    const categoryIds = row.categoryIds.map(id => numericId(id));
    if (new Set(categoryIds).size !== categoryIds.length) throw new Error('Duplicate category membership.');
    return { id: numericId(row.id), name: row.name, sku: row.sku, categoryIds, enabled: row.enabled };
  }), 10_000);
  return { kind: 'READONLY_CATEGORY_SCOPE', schema_version: 1, dry_run: true, complete: true,
    store_id: credentials.storeId, started_at: started, completed_at: new Date().toISOString(),
    hidden_categories_included: true, category_count: categories.length, product_count: products.length, categories, products };
}

async function main() {
  const { values, positionals } = parseArgs({ options: { out: { type: 'string' }, help: { type: 'boolean' } } });
  if (values.help) { console.log('Usage: npm run ecwid:categories -- --out import-data/new-directory\nRead-only category tree and product membership; no stock changes.'); return; }
  if (!values.out || positionals.length) throw new Error('Provide a new --out directory beneath import-data/.');
  const root = await realpath(resolve('import-data')), output = resolve(values.out), parent = await realpath(resolve(output, '..'));
  if (parent !== root && !parent.startsWith(root + '/')) throw new Error('Category snapshots must stay beneath import-data/.');
  try { await stat(output); throw new Error('Use a new snapshot directory.'); }
  catch (error) { if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error; }
  const report = await fetchCategorySnapshot(await loadReadOnlyCredentials());
  const serialized = JSON.stringify(report, null, 2) + '\n';
  if (Buffer.byteLength(serialized) > 10_000_000) throw new Error('Category snapshot exceeded size limit.');
  await mkdir(output, { mode: 0o700 });
  await writeFile(resolve(output, 'category-scope.json'), serialized, { mode: 0o600, flag: 'wx' });
  console.log(JSON.stringify({ store_id: report.store_id, categories: report.category_count, products: report.product_count, stock_writes: 0, output }, null, 2));
}
if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main().catch(error => { console.error(error instanceof Error ? error.message : 'Category snapshot failed.'); process.exitCode = 1; });
}
