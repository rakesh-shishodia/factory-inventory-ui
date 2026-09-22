import { DomainError, normalizeCode } from './domain';
import { canonicalVariationOptions } from './ecwid';

type Row = Record<string, unknown>;
type PreviewLine = {
  row: number; raw_sku: string; sku: string; name: string; scan_code: string; location: string;
  source_row?: number; source_sheet?: string;
  physical_on_hand: number | null; reserved: number; desired_ecwid_quantity: number | null;
  ecwid_product_id: string | null; ecwid_quantity: number | null; difference: number | null;
  ecwid_combination_id: string | null; ecwid_option_signature: string | null;
  status: 'READY' | 'BLOCKED'; errors: string[];
};

function records(value: unknown, field: string, maximum = 10000): Row[] {
  if (!Array.isArray(value) || value.length > maximum || value.some(row => !row || typeof row !== 'object' || Array.isArray(row))) {
    throw new DomainError(400, 'INVALID_IMPORT', `${field} must be an array of at most ${maximum.toLocaleString('en-US')} objects.`);
  }
  return value as Row[];
}

function integer(value: unknown): number | null {
  if (typeof value !== 'number' && typeof value !== 'string') return null;
  const text = String(value).trim();
  if (!/^\d+$/.test(text)) return null;
  const result = Number(text);
  return Number.isSafeInteger(result) && result <= 2147483647 ? result : null;
}

function optionSignature(value: unknown): string | null {
  const options = canonicalVariationOptions(value);
  return options === null ? null : JSON.stringify(options);
}

const catalogIdentity = (product: Row) => JSON.stringify([String(product.id ?? ''), product.combinationId ?? null]);

function catalogInput(value: unknown, expectedStoreId: unknown): { rows: Row[]; snapshot?: Row } {
  if (expectedStoreId !== undefined && (typeof expectedStoreId !== 'string' || !/^[1-9]\d{0,19}$/.test(expectedStoreId))) {
    throw new DomainError(400, 'INVALID_STORE_ID', 'The expected store_id must be a positive numeric string.');
  }
  if (Array.isArray(value)) return { rows: records(value, 'catalog', 20000) };
  if (!value || typeof value !== 'object') throw new DomainError(400, 'INVALID_CATALOGUE_SNAPSHOT', 'Expected a catalogue array or complete READONLY_CATALOGUE snapshot.');
  const snapshot = value as Row;
  if (snapshot.kind !== 'READONLY_CATALOGUE' || snapshot.schema_version !== 1 || snapshot.dry_run !== true ||
      snapshot.complete !== true || snapshot.reservations_confirmed !== false ||
      typeof snapshot.store_id !== 'string' || !/^[1-9]\d{0,19}$/.test(snapshot.store_id)) {
    throw new DomainError(400, 'INVALID_CATALOGUE_SNAPSHOT', 'Expected a complete version-1 read-only catalogue with an identified store and unconfirmed reservations.');
  }
  if (expectedStoreId !== undefined && expectedStoreId !== snapshot.store_id) {
    throw new DomainError(400, 'CATALOGUE_STORE_MISMATCH', 'The catalogue store_id does not match the expected store.');
  }
  const products = records(snapshot.products, 'catalogue products');
  const targets = records(snapshot.stock_targets, 'catalogue stock_targets', 20000);
  if (snapshot.product_count !== products.length || snapshot.stock_target_count !== targets.length) {
    throw new DomainError(400, 'INVALID_CATALOGUE_SNAPSHOT', 'Catalogue counts do not match the complete snapshot contents.');
  }
  const date = (raw: unknown) => typeof raw === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(raw) ? Date.parse(raw) : NaN;
  const started = date(snapshot.started_at), completed = date(snapshot.completed_at);
  if (!Number.isFinite(started) || !Number.isFinite(completed) || completed < started) {
    throw new DomainError(400, 'INVALID_CATALOGUE_SNAPSHOT', 'Catalogue timestamps must identify an ordered UTC snapshot interval.');
  }
  return { rows: targets, snapshot: {
    kind: snapshot.kind, schema_version: 1, store_id: snapshot.store_id,
    started_at: snapshot.started_at, completed_at: snapshot.completed_at,
    product_count: snapshot.product_count, stock_target_count: snapshot.stock_target_count
  } };
}

/** RFC4180-style CSV reader. Balances remain strings until explicitly validated. */
export function parseStockCsv(csv: string): Row[] {
  if (csv.length > 2_000_000) throw new DomainError(413, 'IMPORT_TOO_LARGE', 'Split the stock export into smaller files.');
  csv = csv.replace(/^\uFEFF/, '');
  const lines: string[][] = [];
  let fields: string[] = [], field = '', quoted = false, closed = false;
  const pushField = () => { fields.push(field); field = ''; closed = false; };
  const pushRow = () => { pushField(); if (fields.some(part => part.trim())) lines.push(fields); fields = []; };
  for (let i = 0; i < csv.length; i++) {
    const char = csv[i];
    if (quoted) {
      if (char === '"' && csv[i + 1] === '"') { field += '"'; i++; }
      else if (char === '"') { quoted = false; closed = true; }
      else field += char;
    } else if (char === ',' ) pushField();
    else if (char === '\n' || char === '\r') { if (char === '\r' && csv[i + 1] === '\n') i++; pushRow(); }
    else if (char === '"' && !field && !closed) quoted = true;
    else if (char === '"' || closed) throw new DomainError(400, 'INVALID_CSV', 'The CSV contains invalid quoting. Export it again from the stock sheet.');
    else field += char;
  }
  if (quoted) throw new DomainError(400, 'INVALID_CSV', 'The CSV has an unfinished quoted value.');
  if (field || fields.length || closed) pushRow();
  const headers = lines.shift()?.map(value => value.trim().toLowerCase()) ?? [];
  if (!headers.includes('sku') || !headers.includes('balance') || new Set(headers).size !== headers.length) {
    throw new DomainError(400, 'INVALID_HEADERS', 'The stock CSV needs unique headers including SKU and Balance.');
  }
  return lines.map((values, i) => {
    if (values.length !== headers.length) throw new DomainError(400, 'INVALID_CSV', `Stock CSV row ${i + 2} has the wrong number of columns.`);
    return Object.fromEntries(headers.map((header, column) => [header, values[column]]));
  });
}

/** Join read-only source candidates to separately reviewed Ecwid and reservation data. */
export function preparePreviewInput(stock: string, format: 'CSV' | 'SOURCE_CANDIDATES', catalog: unknown, manifest: unknown): Row {
  if (!manifest || typeof manifest !== 'object' || Array.isArray(manifest)) {
    throw new DomainError(400, 'INVALID_IMPORT', 'Invalid reservation manifest.');
  }
  const metadata = manifest as Row;
  // Only reservation metadata belongs to the manifest. In particular, a `csv`
  // property must not take precedence over the stock file supplied by the caller.
  const input: Row = {
    source_ref: metadata.source_ref, balance_meaning: metadata.balance_meaning,
    ...(metadata.store_id !== undefined ? { store_id: metadata.store_id } : {}),
    reservations_confirmed: metadata.reservations_confirmed,
    reservations: metadata.reservations, catalog
  };
  if (format === 'CSV') return { ...input, rows: parseStockCsv(stock) };

  if (stock.length > 2_000_000) throw new DomainError(413, 'IMPORT_TOO_LARGE', 'Split the stock export into smaller files.');
  let source: unknown;
  try { source = JSON.parse(stock.replace(/^\uFEFF/, '')); }
  catch { throw new DomainError(400, 'INVALID_SOURCE_CANDIDATES', 'The source-candidate file must contain valid JSON.'); }
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    throw new DomainError(400, 'INVALID_SOURCE_CANDIDATES', 'The source-candidate file must contain an object.');
  }
  const candidate = source as Row;
  const permitted = new Set(['kind', 'dry_run', 'source_ref', 'source_hash', 'balance_meaning', 'ecwid_checked', 'reservations_confirmed', 'rows']);
  if (Object.keys(candidate).some(key => !permitted.has(key)) || candidate.kind !== 'SOURCE_CANDIDATES' ||
      candidate.dry_run !== true || candidate.ecwid_checked !== false || candidate.reservations_confirmed !== false ||
      candidate.balance_meaning !== 'PHYSICAL_ON_HAND') {
    throw new DomainError(400, 'INVALID_SOURCE_CANDIDATES', 'Expected dry-run SOURCE_CANDIDATES with physical balances, unchecked Ecwid data and unconfirmed reservations; catalog and reservations must come from separate files.');
  }
  if (typeof candidate.source_ref !== 'string' || !candidate.source_ref.trim() ||
      typeof candidate.source_hash !== 'string' || !/^[a-f0-9]{64}$/i.test(candidate.source_hash)) {
    throw new DomainError(400, 'INVALID_SOURCE_CANDIDATES', 'Source candidates need a nonblank source_ref and a SHA-256 source_hash.');
  }
  if (metadata.source_ref !== undefined && (typeof metadata.source_ref !== 'string' ||
      (metadata.source_ref.trim() && metadata.source_ref.trim() !== candidate.source_ref.trim()))) {
    throw new DomainError(400, 'SOURCE_REFERENCE_MISMATCH', 'The reservation manifest source_ref does not match the source-candidate file.');
  }
  if (metadata.balance_meaning !== undefined && metadata.balance_meaning !== candidate.balance_meaning) {
    throw new DomainError(400, 'BALANCE_MEANING_MISMATCH', 'The reservation manifest must treat source-candidate balances as PHYSICAL_ON_HAND.');
  }
  return {
    ...input, source_ref: candidate.source_ref.trim(), snapshot_source_hash: candidate.source_hash.toLowerCase(),
    balance_meaning: candidate.balance_meaning, rows: records(candidate.rows, 'rows')
  };
}

/** Pure preview. No database changes or Ecwid requests are made here. */
export async function previewImport(input: Row) {
  const rows = typeof input.csv === 'string' ? parseStockCsv(input.csv) : records(input.rows, 'rows');
  const catalogData = catalogInput(input.catalog, input.store_id);
  const catalog = catalogData.rows;
  const reservations = records(input.reservations ?? [], 'reservations');
  const globalErrors: string[] = [];
  if (!rows.length) globalErrors.push('The source stock sheet contains no rows.');
  if (input.reservations_confirmed !== true) globalErrors.push('Confirm the outstanding unpicked quantities of all open Ecwid orders before alignment.');
  if (input.balance_meaning !== 'PHYSICAL_ON_HAND') globalErrors.push('Confirm that Balance represents physical shelf stock before subtracting reservations.');
  if (typeof input.source_ref !== 'string' || !input.source_ref.trim()) globalErrors.push('Provide a source reference identifying the authoritative stock export.');
  if (input.snapshot_source_hash !== undefined && (typeof input.snapshot_source_hash !== 'string' || !/^[a-f0-9]{64}$/i.test(input.snapshot_source_hash))) {
    globalErrors.push('Snapshot source hash must be a SHA-256 hexadecimal digest.');
  }
  const reservedBySku = new Map<string, number>();
  for (const row of reservations) {
    const sku = typeof row.sku === 'string' ? normalizeCode(row.sku) : '';
    const qty = integer(row.quantity);
    if (!sku || qty === null) globalErrors.push('Every reservation needs a SKU and a nonnegative whole-number quantity.');
    else if (reservedBySku.has(sku)) globalErrors.push(`Duplicate reservation summary for ${sku}; aggregate each SKU once.`);
    else reservedBySku.set(sku, qty);
  }
  const catalogBySku = new Map<string, Row[]>();
  const catalogIds = new Map<string, number>();
  for (const product of catalog) {
    const sku = typeof product.sku === 'string' ? normalizeCode(product.sku) : '';
    // A parent with variations is a container, not another independently stocked
    // item. Its base SKU may intentionally equal one child's SKU.
    const container = product.combinationId === null && product.hasVariations === true;
    if (sku && !container) catalogBySku.set(sku, [...(catalogBySku.get(sku) ?? []), product]);
    const id = catalogIdentity(product);
    catalogIds.set(id, (catalogIds.get(id) ?? 0) + 1);
  }
  const sourceCount = new Map<string, number>();
  const allCodes = new Map<string, Set<number>>();
  rows.forEach((row, i) => {
    const sku = typeof row.sku === 'string' ? normalizeCode(row.sku) : '';
    sourceCount.set(sku, (sourceCount.get(sku) ?? 0) + 1);
    const scan = typeof row.scan_code === 'string' && row.scan_code.trim() ? normalizeCode(row.scan_code) : sku;
    for (const code of [sku, scan]) {
      const owners = allCodes.get(code) ?? new Set<number>(); owners.add(i); allCodes.set(code, owners);
    }
  });
  const result: PreviewLine[] = rows.map((row, i) => {
    const rawSku = typeof row.sku === 'string' ? row.sku : '';
    const sku = normalizeCode(rawSku);
    const errors: string[] = [];
    const physical = integer(row.balance);
    const reserved = reservedBySku.get(sku) ?? 0;
    const candidates = catalogBySku.get(sku) ?? [];
    const product = candidates.length === 1 ? candidates[0] : null;
    const scanCode = typeof row.scan_code === 'string' && row.scan_code.trim() ? normalizeCode(row.scan_code) : sku;
    const validSourceRow = typeof row.source_row === 'number' && Number.isSafeInteger(row.source_row) && row.source_row > 0;
    const validSourceSheet = typeof row.source_sheet === 'string' && row.source_sheet.length <= 200;
    if (row.source_row !== undefined && !validSourceRow) errors.push('Source row must be a positive whole-number row index, supplied as a number.');
    if (row.source_sheet !== undefined && !validSourceSheet) errors.push('Source sheet must be a string of at most 200 characters.');
    if (!sku || sku.length > 200 || sku.includes('|')) errors.push('SKU must be nonblank, at most 200 characters, and contain no pipe character.');
    if (!scanCode || scanCode.length > 200 || scanCode.includes('|')) errors.push('Scan code must be at most 200 characters and contain no pipe character.');
    if ((sourceCount.get(sku) ?? 0) > 1) errors.push('Duplicate source SKU, including differences in case or surrounding spaces.');
    if ((allCodes.get(scanCode)?.size ?? 0) > 1 || (allCodes.get(sku)?.size ?? 0) > 1) errors.push('This SKU or scan code also identifies another source row.');
    if (physical === null) errors.push('Balance must be a nonblank, nonnegative whole number.');
    if (row.single_unit_confirmed !== true && row.single_unit_confirmed !== 'true') {
      errors.push('Confirm single_unit_confirmed for this SKU: one Ecwid quantity unit must equal one physical piece, not a pack, bundle or measured length.');
    }
    if (!candidates.length) errors.push('No matching Ecwid product. Leave unmapped until reviewed.');
    if (candidates.length > 1) errors.push('Multiple Ecwid products match this SKU.');
    let productId: string | null = null, combinationId: string | null = null, optionsJson: string | null = null, ecwidQuantity: number | null = null;
    if (product) {
      productId = String(product.id ?? '');
      combinationId = typeof product.combinationId === 'string' ? product.combinationId : null;
      if (!/^[1-9]\d*$/.test(productId) || (catalogIds.get(catalogIdentity(product)) ?? 0) > 1) errors.push('Ecwid product/variation identity is missing, invalid, or duplicated in the catalog export.');
      if (product.combinationId !== null && (!combinationId || !/^[1-9]\d*$/.test(combinationId))) errors.push('Catalog must explicitly identify a positive combinationId for a variation, or null for a base product.');
      optionsJson = optionSignature(product.variationOptions);
      ecwidQuantity = integer(product.quantity);
      if (ecwidQuantity === null) errors.push('Ecwid quantity is missing, negative, or invalid; investigate before alignment.');
      if (product.unlimited !== false) errors.push('Ecwid product does not have confirmed stock tracking enabled.');
      if (product.enabled !== true) errors.push('The parent product must be explicitly enabled.');
      if (product.eligibilityVerified !== true || product.hasBundleRelationships !== false || product.hasExtraOptions !== false) errors.push('Live catalogue eligibility must confirm no bundle relationships or extra non-variation options.');
      if (typeof product.hasOptions !== 'boolean' || typeof product.hasVariations !== 'boolean' || optionsJson === null) errors.push('Catalog must explicitly provide valid option and variation metadata.');
      if (combinationId) {
        if (optionsJson === '[]' || product.hasOptions !== true || product.hasVariations !== false) errors.push('A variation needs its exact nonempty option selection and must be an independent leaf stock target.');
      } else if (product.hasOptions !== false || product.hasVariations !== false || optionsJson !== '[]' ||
          (Array.isArray(product.options) && product.options.length > 0) ||
          (Array.isArray(product.combinations) && product.combinations.length > 0) || Number(product.defaultCombinationId) > 0) {
        errors.push('Base product has options or variations; map an independently stocked variation instead.');
      }
      if (row.ecwid_product_id !== undefined && String(row.ecwid_product_id) !== productId) errors.push('Source mapping disagrees with the matched Ecwid product ID.');
      if (row.ecwid_combination_id !== undefined && row.ecwid_combination_id !== combinationId) errors.push('Source mapping disagrees with the matched Ecwid variation ID.');
      if (row.ecwid_option_signature !== undefined && row.ecwid_option_signature !== optionsJson) errors.push('Source mapping disagrees with the matched Ecwid option selection.');
    }
    const target = physical === null ? null : physical - reserved;
    if (target !== null && target < 0) errors.push('Open orders reserve more units than the physical balance.');
    return {
      row: i + 2, raw_sku: rawSku, sku, name: String(row.name || product?.name || sku), scan_code: scanCode,
      ...(validSourceRow ? { source_row: row.source_row as number } : {}),
      ...(validSourceSheet ? { source_sheet: row.source_sheet as string } : {}),
      location: String(row.location || ''), physical_on_hand: physical, reserved,
      desired_ecwid_quantity: target, ecwid_product_id: productId, ecwid_quantity: ecwidQuantity,
      ecwid_combination_id: combinationId, ecwid_option_signature: optionsJson,
      difference: target === null || ecwidQuantity === null ? null : target - ecwidQuantity,
      status: errors.length || globalErrors.length ? 'BLOCKED' : 'READY', errors
    };
  });
  const hashFields = [rows, input.catalog, reservations, input.reservations_confirmed, input.balance_meaning, input.source_ref];
  if (input.snapshot_source_hash !== undefined) hashFields.push(input.snapshot_source_hash);
  if (input.store_id !== undefined) hashFields.push(input.store_id);
  const hashInput = JSON.stringify(hashFields);
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(hashInput));
  const hash = Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
  return {
    dry_run: true, source_ref: input.source_ref ?? '', source_hash: hash, generated_at: new Date().toISOString(),
    ...(input.snapshot_source_hash !== undefined ? { snapshot_source_hash: input.snapshot_source_hash } : {}),
    ...(catalogData.snapshot ? { catalogue_snapshot: catalogData.snapshot } : {}),
    balance_meaning: input.balance_meaning ?? null, global_errors: [...new Set(globalErrors)],
    ready_count: result.filter(row => row.status === 'READY').length,
    blocked_count: result.filter(row => row.status === 'BLOCKED').length,
    rows: result,
    ecwid_only_skus: [...catalogBySku.keys()].filter(sku => !sourceCount.has(sku)),
    reservation_only_skus: [...reservedBySku.keys()].filter(sku => !sourceCount.has(sku)),
    next_step: 'Review this report, reconcile open orders, freeze stock activity, take backups, and create a fresh cutover plan. This preview has changed no stock.'
  };
}
