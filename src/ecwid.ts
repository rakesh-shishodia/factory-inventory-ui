/** Only this adapter knows Ecwid's wire format. It never automatically retries a write. */
export type EcwidFailure = 'UNKNOWN' | 'REJECTED' | 'RETRYABLE';

export class EcwidError extends Error {
  constructor(message: string, public outcome: EcwidFailure, public status?: number, public retryAfter = 60) {
    super(message);
    this.name = 'EcwidError';
  }
}

export interface EcwidOrderLine {
  id: string;
  productId: string;
  sku: string;
  name: string;
  quantity: number;
  combinationId: string | null;
  selectedOptions: unknown[];
  digital: boolean;
  trackQuantity: boolean;
}

export interface EcwidOrder {
  id: string;
  createdAt?: string;
  paymentStatus: string;
  fulfillmentStatus: string;
  updatedAt: string;
  items: EcwidOrderLine[];
}

export interface EcwidProduct {
  id: string;
  sku: string;
  name: string;
  quantity: number | null;
  unlimited: boolean;
  enabled: boolean;
  hasOptions: boolean;
  hasVariations: boolean;
  /** id always identifies the parent; a combination is a separate stock target. */
  combinationId?: string | null;
  variationOptions?: EcwidVariationOption[];
  hasBundleRelationships?: boolean;
  hasExtraOptions?: boolean;
  eligibilityVerified?: boolean;
  variations?: EcwidProduct[];
}

export interface EcwidVariationOption { name: string; value: string }

export interface EcwidPage<T> { items: T[]; total: number; offset: number; count: number }
export type EcwidFetch = (input: string | URL | Request, init?: RequestInit) => Promise<Response>;

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Expected an Ecwid object.');
  return value as Record<string, unknown>;
}

function identifier(value: unknown, field: string): string {
  if ((typeof value !== 'string' && typeof value !== 'number') || !String(value).trim()) {
    throw new Error(`Missing Ecwid ${field}.`);
  }
  return String(value);
}

function wholeNumber(value: unknown, field: string, minimum = 0): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < minimum) {
    throw new Error(`Invalid Ecwid ${field}.`);
  }
  return value;
}

function stockIdentifier(value: unknown, field: string): string {
  const id = identifier(value, field);
  if (!/^[1-9]\d{0,30}$/.test(id) || (typeof value === 'number' && !Number.isSafeInteger(value))) {
    throw new Error(`Invalid Ecwid ${field}.`);
  }
  return id;
}

/** Exact, deterministic option identity; malformed or extra option types never become []. */
export function canonicalVariationOptions(value: unknown): EcwidVariationOption[] | null {
  if (!Array.isArray(value) || value.length > 100) return null;
  const names = new Set<string>();
  const result: EcwidVariationOption[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return null;
    const option = entry as Record<string, unknown>;
    if (typeof option.name !== 'string' || !option.name.trim() || option.name.length > 200
      || typeof option.value !== 'string' || !option.value.trim() || option.value.length > 2_000
      || names.has(option.name) || (option.type !== undefined && option.type !== 'CHOICE')) return null;
    if (option.valuesArray != null && (!Array.isArray(option.valuesArray)
      || option.valuesArray.length !== 1 || option.valuesArray[0] !== option.value)) return null;
    if (option.files != null && (!Array.isArray(option.files) || option.files.length !== 0)) return null;
    if (option.selections != null) {
      if (!Array.isArray(option.selections) || option.selections.length > 1) return null;
      for (const selection of option.selections) {
        if (!selection || typeof selection !== 'object' || Array.isArray(selection)
          || selection.selectionTitle !== option.value
          || (selection.selectionModifier !== undefined && (typeof selection.selectionModifier !== 'number'
            || !Number.isFinite(selection.selectionModifier)))
          || (selection.selectionModifierType !== undefined
            && !['ABSOLUTE', 'PERCENT'].includes(selection.selectionModifierType))) return null;
      }
    }
    names.add(option.name);
    result.push({ name: option.name, value: option.value });
  }
  return result.sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : a.value < b.value ? -1 : a.value > b.value ? 1 : 0);
}

export function parseOrder(value: unknown): EcwidOrder {
  const raw = record(value);
  if (!Array.isArray(raw.items) || raw.items.length > 500) throw new Error('Invalid Ecwid order items.');
  const updated = typeof raw.updateTimestamp === 'number' ? raw.updateTimestamp * 1000 : Date.parse(String(raw.updateDate));
  if (!Number.isFinite(updated)) throw new Error('Missing Ecwid order update timestamp.');
  const seen = new Set<string>();
  const items = raw.items.map((value): EcwidOrderLine => {
    const line = record(value);
    const id = identifier(line.id, 'line ID');
    if (seen.has(id)) throw new Error('Duplicate Ecwid order line ID.');
    seen.add(id);
    // Missing is Ecwid's legacy empty default; supplied malformed data must stop ingestion.
    if (line.selectedOptions !== undefined && !Array.isArray(line.selectedOptions)) {
      throw new Error('Invalid Ecwid selected options.');
    }
    const combinationId = line.combinationId == null || line.combinationId === 0 || line.combinationId === '0'
      ? null : stockIdentifier(line.combinationId, 'variation ID');
    return {
      id,
      productId: line.productId == null ? '' : String(line.productId),
      sku: typeof line.sku === 'string' ? line.sku.trim().toUpperCase() : '',
      name: typeof line.name === 'string' ? line.name : '',
      quantity: wholeNumber(line.quantity, 'order quantity', 1),
      combinationId,
      selectedOptions: Array.isArray(line.selectedOptions) ? line.selectedOptions : [],
      digital: line.digital === true,
      trackQuantity: line.trackQuantity === true,
    };
  });
  return {
    id: identifier(raw.id, 'order ID'),
    ...(typeof raw.createTimestamp === 'number' ? { createdAt: new Date(raw.createTimestamp * 1000).toISOString() } : {}),
    paymentStatus: identifier(raw.paymentStatus, 'payment status'),
    fulfillmentStatus: identifier(raw.fulfillmentStatus, 'fulfillment status'),
    updatedAt: new Date(updated).toISOString(),
    items,
  };
}

// Nested projections retain inventory evidence without variant image/price payloads.
const PRODUCT_FIELDS = 'id,sku,name,quantity,unlimited,enabled,options(name,type,choices(text)),combinations(id,sku,quantity,unlimited,options(name,value),compositeParents,compositeComponents),defaultCombinationId,compositeParents,compositeComponents';
export const ECWID_PRODUCT_LIST_FIELDS = `total,count,offset,items(${PRODUCT_FIELDS})`;
export const ECWID_ORDER_LIST_FIELDS = 'total,count,offset,items(id,paymentStatus,fulfillmentStatus,createTimestamp,updateTimestamp,items(id,productId,sku,name,quantity,combinationId,selectedOptions,digital,trackQuantity))';

function bundleRelationships(raw: Record<string, unknown>): boolean {
  // These optional API fields are absent for ordinary products on some stores.
  // Any nonempty relationship (including being a component) is out of this phase.
  return ['compositeParents', 'compositeComponents'].some(key => raw[key] !== undefined
    && (!Array.isArray(raw[key]) || raw[key].length > 0));
}

function parentChoices(raw: unknown): Map<string, Set<string>> | null {
  if (!Array.isArray(raw) || raw.length > 100) return null;
  const choices = new Map<string, Set<string>>();
  for (const value of raw) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
    const option = value as Record<string, unknown>;
    if (typeof option.name !== 'string' || !option.name.trim() || option.name.length > 200
      || choices.has(option.name) || typeof option.type !== 'string'
      || !['SELECT', 'RADIO', 'SIZE', 'SWATCHES'].includes(option.type)
      || !Array.isArray(option.choices) || option.choices.length === 0 || option.choices.length > 10_000) return null;
    const values = new Set<string>();
    for (const choice of option.choices) {
      if (!choice || typeof choice !== 'object' || Array.isArray(choice)
        || typeof choice.text !== 'string' || !choice.text.trim() || choice.text.length > 2_000
        || values.has(choice.text)) return null;
      values.add(choice.text);
    }
    choices.set(option.name, values);
  }
  return choices;
}

/** Normalize one full product response. No SKU or stock is inherited into variations. */
export function flattenProduct(value: unknown): EcwidProduct[] {
  const raw = record(value);
  if (raw.combinations !== undefined && (!Array.isArray(raw.combinations) || raw.combinations.length > 10_000)) {
    throw new Error('Invalid Ecwid product variations.');
  }
  const choices = parentChoices(raw.options);
  const combinations = Array.isArray(raw.combinations) ? raw.combinations : [];
  const hasVariations = combinations.length > 0 || Number(raw.defaultCombinationId) > 0;
  const hasBundles = bundleRelationships(raw);
  const parent: EcwidProduct = {
    id: stockIdentifier(raw.id, 'product ID'),
    sku: typeof raw.sku === 'string' ? raw.sku.trim().toUpperCase() : '',
    name: typeof raw.name === 'string' ? raw.name : '',
    quantity: raw.quantity == null ? null : wholeNumber(raw.quantity, 'product quantity', -Number.MAX_SAFE_INTEGER),
    unlimited: raw.unlimited !== false,
    enabled: raw.enabled === true,
    hasOptions: !Array.isArray(raw.options) || raw.options.length > 0,
    hasVariations,
    combinationId: null,
    variationOptions: [],
    hasBundleRelationships: hasBundles,
    hasExtraOptions: choices === null || (!hasVariations && choices.size > 0),
    eligibilityVerified: choices !== null && Array.isArray(raw.combinations)
      && typeof raw.enabled === 'boolean' && typeof raw.unlimited === 'boolean'
      && !(Number(raw.defaultCombinationId) > 0 && combinations.length === 0),
  };
  const ids = new Set<string>();
  const signatures = new Set<string>();
  const variations = combinations.map((value): EcwidProduct => {
    const combination = record(value);
    const combinationId = stockIdentifier(combination.id, 'variation ID');
    if (ids.has(combinationId)) throw new Error('Duplicate Ecwid variation ID.');
    ids.add(combinationId);
    const options = canonicalVariationOptions(combination.options);
    const signature = JSON.stringify(options);
    if (options && signatures.has(signature)) throw new Error('Duplicate Ecwid variation option identity.');
    if (options) signatures.add(signature);
    const hasExtraOptions = choices === null || options === null || options.length === 0
      || choices.size !== options.length || options.some(option => !choices.get(option.name)?.has(option.value));
    return {
      id: parent.id, combinationId,
      sku: typeof combination.sku === 'string' ? combination.sku.trim().toUpperCase() : '',
      name: parent.name,
      quantity: combination.quantity == null ? null : wholeNumber(combination.quantity, 'variation quantity', -Number.MAX_SAFE_INTEGER),
      unlimited: combination.unlimited !== false,
      enabled: parent.enabled,
      hasOptions: Boolean(options?.length), hasVariations: false,
      variationOptions: options ?? [],
      hasBundleRelationships: hasBundles || bundleRelationships(combination),
      hasExtraOptions,
      eligibilityVerified: parent.eligibilityVerified === true && options !== null
        && typeof combination.unlimited === 'boolean',
    };
  });
  parent.variations = variations;
  return [parent, ...variations];
}

/** Parent containers cannot be imported alongside their independently stocked children. */
export function productStockTargets(product: EcwidProduct): EcwidProduct[] {
  if (product.combinationId) return [product];
  return product.hasVariations ? product.variations ?? [] : [product];
}

function parseProduct(value: unknown): EcwidProduct {
  try {
    return flattenProduct(value)[0];
  } catch {
    // A successfully received but invalid catalogue snapshot is not a transient
    // transport failure. Product webhooks must quarantine every mapped target
    // until its stock identity/settings can be reconciled.
    throw new EcwidError('Ecwid product contains invalid stock or variation metadata.', 'REJECTED');
  }
}

/** Bounds both declared and chunked bodies, including incoming webhook requests. */
export async function readBoundedJson(response: Response | Request, maximum = 2_000_000): Promise<unknown> {
  const declared = Number(response.headers.get('Content-Length'));
  if (declared > maximum) throw new Error('JSON payload is too large.');
  if (!response.body) throw new Error('Missing JSON payload.');
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let size = 0;
  let text = '';
  try {
    for (;;) {
      const chunk = await reader.read();
      if (chunk.done) break;
      size += chunk.value.byteLength;
      if (size > maximum) {
        await reader.cancel();
        throw new Error('JSON payload is too large.');
      }
      text += decoder.decode(chunk.value, { stream: true });
    }
    return JSON.parse(text + decoder.decode());
  } finally {
    reader.releaseLock();
  }
}

export class EcwidClient {
  private base: string;
  constructor(private credentials: { storeId: string; token: string }, private fetcher: EcwidFetch = fetch, private timeoutMs = 15_000) {
    if (!/^\d+$/.test(credentials.storeId) || !credentials.token.trim()) {
      throw new EcwidError('Ecwid store ID and token are required.', 'REJECTED');
    }
    if (!Number.isInteger(timeoutMs) || timeoutMs < 1 || timeoutMs > 60_000) {
      throw new EcwidError('Ecwid timeout must be between 1 and 60000 milliseconds.', 'REJECTED');
    }
    this.base = `https://app.ecwid.com/api/v3/${credentials.storeId}`;
  }

  private async request(path: string, method = 'GET', body?: unknown): Promise<unknown> {
    const isWrite = method !== 'GET';
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await this.fetcher(`${this.base}${path}`, {
        method,
        headers: { Authorization: `Bearer ${this.credentials.token}`, 'Content-Type': 'application/json' },
        body: body === undefined ? undefined : JSON.stringify(body),
        signal: controller.signal,
        redirect: 'error',
      });
      if (response.status === 429) {
        // Ecwid documents rate-limited requests as ignored; these alone can safely retry.
        const after = Number(response.headers.get('Retry-After'));
        await response.body?.cancel();
        throw new EcwidError('Ecwid rate limit; request was not applied.', 'RETRYABLE', 429,
          Number.isFinite(after) && after > 0 ? Math.min(Math.ceil(after), 43_200) : 60);
      }
      if (!response.ok) {
        await response.body?.cancel();
        const definiteRejection = [400, 401, 402, 403, 404, 405, 409, 422].includes(response.status);
        throw new EcwidError(`Ecwid returned HTTP ${response.status}.`,
          definiteRejection ? 'REJECTED' : isWrite ? 'UNKNOWN' : 'RETRYABLE', response.status);
      }
      return await readBoundedJson(response);
    } catch (error) {
      if (error instanceof EcwidError) throw error;
      throw new EcwidError(isWrite
        ? 'Ecwid write outcome is uncertain. Verify it manually before any retry.'
        : 'Could not read a valid response from Ecwid.', isWrite ? 'UNKNOWN' : 'RETRYABLE');
    } finally {
      clearTimeout(timer);
    }
  }

  async getOrder(id: string): Promise<EcwidOrder> {
    return parseOrder(await this.request(`/orders/${encodeURIComponent(id)}`));
  }

  async listOrders(options: { offset?: number; createdTo?: number; updatedFrom?: number; updatedTo?: number } = {}): Promise<EcwidPage<EcwidOrder>> {
    const query = new URLSearchParams({ limit: '100', offset: String(options.offset ?? 0) });
    for (const key of ['createdTo', 'updatedFrom', 'updatedTo'] as const) {
      if (options[key] !== undefined) query.set(key, String(options[key]));
    }
    query.set('responseFields', ECWID_ORDER_LIST_FIELDS);
    return this.page(await this.request(`/orders?${query}`), parseOrder);
  }

  async listProducts(offset = 0): Promise<EcwidPage<EcwidProduct>> {
    const query = new URLSearchParams({ limit: '100', offset: String(offset),
      responseFields: ECWID_PRODUCT_LIST_FIELDS });
    return this.page(await this.request(`/products?${query}`), parseProduct);
  }

  async getProductStock(productId: string, combinationId?: string | null): Promise<EcwidProduct> {
    if (!/^[1-9]\d{0,30}$/.test(productId) || (combinationId != null && !/^[1-9]\d{0,30}$/.test(combinationId))) {
      throw new EcwidError('Invalid product or variation ID.', 'REJECTED');
    }
    // One parent snapshot provides variation stock and the option/bundle context atomically.
    const parent = parseProduct(await this.request(`/products/${productId}?responseFields=${PRODUCT_FIELDS}`));
    if (parent.id !== productId) throw new EcwidError('Ecwid returned a different product.', 'REJECTED');
    if (combinationId == null) return parent;
    const variation = parent.variations?.find(value => value.combinationId === combinationId);
    if (!variation) throw new EcwidError('Ecwid variation is missing from its parent product.', 'REJECTED');
    return variation;
  }

  async getProductStockTargets(productId: string): Promise<EcwidProduct[]> {
    return productStockTargets(await this.getProductStock(productId));
  }

  /**
   * One-time approved opening alignment, not the normal movement/outbox path.
   * The caller must freeze sales/movements, persist a write claim, and freshly
   * verify the exact finite stock target before calling. The request contains
   * only quantity, never an inventory policy or supplier-mode change.
   * A successful confirmation still requires a separate read-back by the caller.
   * Docs: /products/update-product and /products/product-variations/update-product-variation.
   */
  async setStockQuantity(productId: string, quantity: number, combinationId: string | null = null): Promise<{ warning?: string }> {
    if (typeof productId !== 'string' || !/^[1-9]\d{0,30}$/.test(productId)
      || (combinationId !== null && (typeof combinationId !== 'string' || !/^[1-9]\d{0,30}$/.test(combinationId)))
      || !Number.isSafeInteger(quantity) || quantity < 0 || quantity > 2_147_483_647) {
      throw new EcwidError('Invalid product ID, variation ID, or opening stock quantity.', 'REJECTED');
    }
    const target = combinationId === null ? `/products/${productId}` : `/products/${productId}/combinations/${combinationId}`;
    const value = await this.request(target, 'PUT', { quantity });
    let raw: Record<string, unknown>;
    try { raw = record(value); } catch { throw new EcwidError('Ecwid returned an invalid opening stock confirmation.', 'UNKNOWN'); }
    if (raw.updateCount === 0) throw new EcwidError('Ecwid did not update this stock target.', 'REJECTED');
    if (raw.updateCount !== 1) throw new EcwidError('Ecwid did not confirm whether the opening stock was applied.', 'UNKNOWN');
    return typeof raw.warning === 'string' ? { warning: raw.warning } : {};
  }

  async adjustStock(productId: string, delta: number, combinationId?: string | null): Promise<{ warning?: string }> {
    if (!/^[1-9]\d{0,30}$/.test(productId) || !Number.isSafeInteger(delta) || delta === 0
      || (combinationId != null && !/^[1-9]\d{0,30}$/.test(combinationId))) {
      throw new EcwidError('Invalid product ID, variation ID, or stock delta.', 'REJECTED');
    }
    const target = combinationId == null ? `/products/${productId}` : `/products/${productId}/combinations/${combinationId}`;
    const value = await this.request(`${target}/inventory`, 'PUT', { quantityDelta: delta });
    let raw: Record<string, unknown>;
    try { raw = record(value); } catch { throw new EcwidError('Ecwid returned an invalid stock update confirmation.', 'UNKNOWN'); }
    if (raw.updateCount === 0) throw new EcwidError('Ecwid did not update this product.', 'REJECTED');
    if (raw.updateCount !== 1) throw new EcwidError('Ecwid did not confirm whether the stock update was applied.', 'UNKNOWN');
    return typeof raw.warning === 'string' ? { warning: raw.warning } : {};
  }

  private page<T>(value: unknown, parse: (value: unknown) => T): EcwidPage<T> {
    const raw = record(value);
    if (!Array.isArray(raw.items) || raw.items.length > 100
      || (raw.count !== undefined && raw.count !== raw.items.length)) throw new Error('Invalid Ecwid page.');
    return { items: raw.items.map(parse), total: wholeNumber(raw.total, 'page total'),
      offset: wholeNumber(raw.offset, 'page offset'), count: raw.items.length };
  }
}

export interface EcwidWebhook {
  eventId: string;
  eventCreated: string;
  storeId: string;
  entityId: string;
  eventType: string;
}

export function parseWebhook(value: unknown): EcwidWebhook {
  const raw = record(value);
  const type = identifier(raw.eventType, 'webhook type');
  const data = raw.data && typeof raw.data === 'object' ? record(raw.data) : {};
  const eventCreated = identifier(raw.eventCreated, 'webhook timestamp');
  if (!/^\d+$/.test(eventCreated)) throw new Error('Invalid webhook timestamp.');
  return {
    eventId: identifier(raw.eventId, 'webhook ID'), eventCreated,
    storeId: identifier(raw.storeId, 'webhook store'), eventType: type,
    entityId: identifier(type.startsWith('order.') ? data.orderId : raw.entityId, 'webhook entity'),
  };
}

/** Ecwid signs timestamp.eventId, not the JSON body. Persist eventId to reject duplicates. */
export async function verifyWebhookSignature(event: EcwidWebhook, signature: string | null, clientSecret: string): Promise<boolean> {
  if (!signature || !clientSecret || !/^[A-Za-z0-9+/]{43}=$/.test(signature)) return false;
  try {
    const bytes = Uint8Array.from(atob(signature), (character) => character.charCodeAt(0));
    const encoder = new TextEncoder();
    const key = await crypto.subtle.importKey('raw', encoder.encode(clientSecret), { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']);
    return await crypto.subtle.verify('HMAC', key, bytes, encoder.encode(`${event.eventCreated}.${event.eventId}`));
  } catch {
    return false;
  }
}
