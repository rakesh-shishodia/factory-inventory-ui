import { mkdir, realpath, stat, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { canonicalVariationOptions, EcwidClient, ECWID_ORDER_LIST_FIELDS, ECWID_PRODUCT_LIST_FIELDS, productStockTargets, type EcwidFetch, type EcwidPage, type EcwidOrder, type EcwidProduct } from '../src/ecwid';
import { loadReadOnlyCredentials } from './ecwid-readonly';
import { PICKABLE_FULFILLMENT_STATUSES, TERMINAL_FULFILLMENT_STATUSES } from '../src/domain';
import { sanitizedOpeningOrder } from '../src/workbook-identity';

/** Defense in depth: the snapshot cannot use the adapter's write methods. */
export function getOnlyEcwidFetch(storeId: string, fetcher: EcwidFetch = fetch): EcwidFetch {
  if (!/^[1-9]\d{0,19}$/.test(storeId)) throw new Error('Invalid store ID.');
  return async (input, init) => {
    const url = new URL(typeof input === 'string' || input instanceof URL ? input : input.url);
    const method = init?.method ?? (input instanceof Request ? input.method : 'GET');
    const orders = url.pathname === `/api/v3/${storeId}/orders`;
    const allowed = new Set(['limit', 'offset', 'responseFields', ...(orders ? ['createdTo'] : [])]);
    const projection = orders ? ECWID_ORDER_LIST_FIELDS : ECWID_PRODUCT_LIST_FIELDS;
    if (method !== 'GET' || init?.body != null || (input instanceof Request && input.body !== null)
      || url.origin !== 'https://app.ecwid.com' || url.username || url.password || url.hash
      || ![`/api/v3/${storeId}/products`, `/api/v3/${storeId}/orders`].includes(url.pathname)
      || [...url.searchParams.keys()].some(key => !allowed.has(key) || url.searchParams.getAll(key).length !== 1)
      || url.searchParams.get('responseFields') !== projection || url.searchParams.get('limit') !== '100'
      || !/^\d{1,5}$/.test(url.searchParams.get('offset') ?? '')
      || (orders && !/^[1-9]\d{0,12}$/.test(url.searchParams.get('createdTo') ?? ''))) {
      throw new Error('Snapshot requests are restricted to approved catalogue/order GET lists and response fields.');
    }
    return fetcher(input, { ...init, method: 'GET', redirect: 'error' });
  };
}

export async function collectPages<T extends { id: string }>(read: (offset: number) => Promise<EcwidPage<T>>, maximum: number): Promise<T[]> {
  if (!Number.isSafeInteger(maximum) || maximum < 0) throw new Error('Invalid listing size limit.');
  const result: T[] = [];
  const ids = new Set<string>();
  let expectedTotal: number | undefined;
  for (;;) {
    const page = await read(result.length);
    if (!page || !Array.isArray(page.items) || !Number.isSafeInteger(page.total) || page.total < 0 || page.total > maximum ||
        page.offset !== result.length || page.count !== page.items.length || page.count > 100 ||
        (expectedTotal !== undefined && page.total !== expectedTotal) || result.length + page.count > page.total) {
      throw new Error('The listing changed during pagination or returned inconsistent metadata. Create a new snapshot.');
    }
    expectedTotal = page.total;
    for (const item of page.items) {
      if (!item || typeof item.id !== 'string' || !item.id.trim() || ids.has(item.id)) {
        throw new Error('Duplicate or missing record ID during pagination. Create a new snapshot.');
      }
      ids.add(item.id); result.push(item);
    }
    if (result.length === page.total) return result;
    if (!page.count) throw new Error('Listing ended before every record was received.');
  }
}

export function pendingOrderReview(orders: EcwidOrder[]) {
  const activePayment = new Set(['PAID', 'AWAITING_PAYMENT']);
  const cancelledPayment = new Set(['CANCELLED', 'REFUNDED']);
  const terminalFulfillment = new Set(TERMINAL_FULFILLMENT_STATUSES);
  return orders.filter(order => !cancelledPayment.has(order.paymentStatus) && !terminalFulfillment.has(order.fulfillmentStatus))
    .map(order => ({
      id: order.id, payment_status: order.paymentStatus, fulfillment_status: order.fulfillmentStatus,
      review_category: activePayment.has(order.paymentStatus) ? 'PENDING_PICK_RECONCILIATION' : 'PAYMENT_STATUS_EXCEPTION',
      updated_at: order.updatedAt, pickable_after_cutover: order.paymentStatus === 'PAID' &&
        PICKABLE_FULFILLMENT_STATUSES.includes(order.fulfillmentStatus),
      lines: order.items.map(line => ({
        ecwid_line_id: line.id, ecwid_product_id: line.productId, ecwid_combination_id: line.combinationId,
        sku: line.sku, name: line.name, ordered_quantity: line.quantity,
        previously_picked_quantity: null, unpicked_quantity: null,
        // Free text/file selections can contain customer PII or private URLs. Only stock identity is retained.
        selected_options: canonicalVariationOptions(line.selectedOptions),
        selected_options_supported: canonicalVariationOptions(line.selectedOptions) !== null,
        reconciliation_required: true
      }))
    }));
}

export async function fetchReadOnlySnapshot(credentials: { storeId: string; token: string }, fetcher: EcwidFetch = fetch,
  progress: (message: string) => void = () => {}) {
  const started = new Date();
  const client = new EcwidClient(credentials, getOnlyEcwidFetch(credentials.storeId, fetcher));
  const products = await collectPages<EcwidProduct>(async offset => {
    const page = await client.listProducts(offset);
    progress(`Catalogue: ${offset + page.count}/${page.total} products read.`);
    return page;
  }, 10_000);
  const targets = products.flatMap(productStockTargets);
  if (targets.length > 20_000) throw new Error('Catalogue has more than 20,000 stock targets.');
  const identities = targets.map(target => `${target.id}:${target.combinationId ?? ''}`);
  if (new Set(identities).size !== identities.length) throw new Error('Catalogue contains repeated product/variation identities.');
  // Fixed creation cutoff, unfiltered pagination: status transitions do not shift a filtered open-order list.
  const orders = await collectPages<EcwidOrder>(async offset => {
    const page = await client.listOrders({ offset, createdTo: Math.floor(started.getTime() / 1000) });
    if (offset === 0 || (offset + page.count) % 1000 === 0 || offset + page.count === page.total) {
      progress(`Orders: ${offset + page.count}/${page.total} records checked.`);
    }
    return page;
  }, 50_000);
  const completedAt = new Date().toISOString();
  const catalogue = {
    kind: 'READONLY_CATALOGUE', schema_version: 1, dry_run: true, complete: true,
    store_id: credentials.storeId, started_at: started.toISOString(), completed_at: completedAt, generated_at: completedAt,
    product_count: products.length, stock_target_count: targets.length, products, stock_targets: targets,
    reservations_confirmed: false,
    note: 'All pages read, not a transactionally frozen stock snapshot. Product units and open-order picks still require review.'
  };
  const pending = pendingOrderReview(orders);
  const orderReview = {
    kind: 'READONLY_ORDER_REVIEW', schema_version: 1, dry_run: true, complete: true, store_id: credentials.storeId,
    started_at: started.toISOString(), completed_at: completedAt, orders_checked: orders.length,
    creation_cutoff: Math.floor(started.getTime() / 1000), pending_order_count: pending.length,
    reservations_confirmed: false, orders: pending,
    note: 'Ordered quantity is not unpicked quantity. Confirm prior physical picks and stale fulfillment statuses before any stock alignment. Orders created after the cutoff are excluded. No reservation quantities have been approved.'
  };
  const pendingIds = new Set(pending.map(order => order.id));
  // Separate normalized evidence for the opening importer. Do not drop digital
  // or stock-policy evidence, and never save customer-entered text/file URLs.
  const openingPending = await Promise.all(orders.filter(order => pendingIds.has(order.id)).map(order=>sanitizedOpeningOrder(order,targets)));
  const openingOrders = {
    kind: 'READONLY_ORDERS' as const, schema_version: 2 as const, dry_run: true as const, complete: true as const,
    store_id: credentials.storeId, started_at: started.toISOString(), completed_at: completedAt,
    creation_cutoff: Math.floor(started.getTime() / 1000), orders_checked: orders.length,
    pending_order_count: openingPending.length, line_count: openingPending.reduce((count, order) => count + order.items.length, 0),
    orders: openingPending
  };
  return { catalogue, orderReview, openingOrders };
}

export function renderOrdersReview(report: ReturnType<typeof pendingOrderReview>, storeId: string): string {
  const esc = (value: unknown) => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
  const rows = report.flatMap(order => order.lines.map(line => `<tr><td>${esc(order.id)}</td><td>${esc(order.payment_status)}<br>${esc(order.fulfillment_status)}</td><td>${esc(line.sku)}<br>${esc(line.name)}</td><td>${esc(line.ecwid_product_id)}<br>${esc(line.ecwid_combination_id ?? 'Base product')}</td><td>${esc(line.ordered_quantity)}</td><td>Needs confirmation</td></tr>`));
  return `<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"><title>Open-order reconciliation</title><style>body{font:15px/1.5 system-ui;background:#f5f7fa;color:#182333;margin:28px}h1{font-size:28px}.notice{padding:16px;background:#fff4df;border-left:4px solid #b7770c}.scroll{overflow:auto}table{border-collapse:collapse;width:100%;background:white;min-width:850px}td,th{padding:12px;text-align:left;vertical-align:top;border-bottom:1px solid #dce2e9}th{background:#21344b;color:white}</style><h1>Open-order reconciliation</h1><p>Store ${esc(storeId)}. ${report.length} orders with nonterminal fulfillment status, including ambiguous payment-status exceptions.</p><p class="notice">Read-only review. Ordered quantities below are NOT confirmed unpicked reservations. Confirm what has already left the shelf, and review old or incorrect order statuses. Awaiting Payment orders cannot be picked until Paid. Partial refunds, incomplete payments, and custom payment statuses need review and are not pickable. No stock was changed.</p><div class="scroll"><table><thead><tr><th>Order</th><th>Status</th><th>Item SKU and name</th><th>Product / variation</th><th>Ordered</th><th>Unpicked</th></tr></thead><tbody>${rows.join('')}</tbody></table></div></html>`;
}

async function main() {
  const { values, positionals } = parseArgs({ options: { out: { type: 'string' }, help: { type: 'boolean' } } });
  if (values.help) {
    console.log('Usage: npm run ecwid:snapshot -- --out import-data/ecwid-YYYYMMDD\nReads all catalogue and order list pages via GET only. Saves private review files, never imports or changes stock.');
    return;
  }
  if (!values.out || positionals.length) throw new Error('Provide a new --out directory beneath import-data/.');
  const privateRoot = await realpath(resolve('import-data'));
  const output = resolve(values.out);
  const parent = await realpath(resolve(output, '..'));
  if (parent !== privateRoot && !parent.startsWith(privateRoot + '/')) throw new Error('Snapshot files must stay beneath import-data/.');
  try { await stat(output); throw new Error('Output directory exists; use a new snapshot directory.'); }
  catch (error) { if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error; }
  const snapshot = await fetchReadOnlySnapshot(await loadReadOnlyCredentials(), fetch, message => console.error(message));
  const files = [
    ['catalog.json', JSON.stringify(snapshot.catalogue, null, 2) + '\n'],
    ['orders-review.json', JSON.stringify(snapshot.orderReview, null, 2) + '\n'],
    ['opening-orders.json', JSON.stringify(snapshot.openingOrders, null, 2) + '\n'],
    ['orders-review.html', renderOrdersReview(snapshot.orderReview.orders, snapshot.catalogue.store_id)]
  ];
  if (files.some(([, text]) => Buffer.byteLength(text) > 30_000_000)) throw new Error('A snapshot output exceeded the private review size limit.');
  await mkdir(output, { mode: 0o700 });
  for (const [name, text] of files) await writeFile(resolve(output, name), text, { flag: 'wx', mode: 0o600 });
  console.log(JSON.stringify({ store_id: snapshot.catalogue.store_id, products: snapshot.catalogue.product_count,
    stock_targets: snapshot.catalogue.stock_target_count, orders_checked: snapshot.orderReview.orders_checked,
    pending_orders_for_review: snapshot.orderReview.pending_order_count, stock_writes: 0, imported_items: 0, output }, null, 2));
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main().catch(error => { console.error(error instanceof Error ? error.message : 'Read-only snapshot failed.'); process.exitCode = 1; });
}
