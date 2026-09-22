import { canonicalVariationOptions } from './ecwid';
import { DomainError, normalizeCode } from './domain';
import { opaqueWorkbookPolicy } from './workbook-identity';

/** Explicitly reviewed physical targets kept outside the app during the pilot. */
export interface WorkbookTarget {
  ecwid_product_id: string;
  ecwid_combination_id: string | null;
  ecwid_option_signature: string;
  sku: string;
  name: string;
  sku_source?: 'TARGET' | 'PARENT_IF_VARIATION_BLANK';
  option_policy?: 'EXACT' | 'STOCK_SELECTION_PLUS_OPAQUE_EXTRAS';
}

export interface WorkbookReview { reference: string; actor: string; timestamp?: string }

export function workbookTargetId(target: WorkbookTarget): string {
  return `workbook:${target.ecwid_product_id}:${target.ecwid_combination_id ?? 'simple'}`;
}

export function validateWorkbookTargets(value: unknown): WorkbookTarget[] {
  const invalid = () => new DomainError(400, 'INVALID_WORKBOOK_TARGETS', 'Workbook-managed targets need exact reviewed identities and explicit policies for inherited SKUs or extra options.');
  if (!Array.isArray(value) || value.length > 500) throw invalid();
  const ids = new Set<string>();
  const skus = new Map<string,WorkbookTarget>();
  return value.map(raw => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw invalid();
    const row = raw as Record<string, unknown>;
    if (Object.keys(row).some(key=>!['ecwid_product_id','ecwid_combination_id','ecwid_option_signature','sku','name','sku_source','option_policy'].includes(key))
      || typeof row.ecwid_product_id !== 'string' || !/^[1-9]\d{0,30}$/.test(row.ecwid_product_id)
      || (row.ecwid_combination_id !== null && (typeof row.ecwid_combination_id !== 'string' || !/^[1-9]\d{0,30}$/.test(row.ecwid_combination_id)))
      || typeof row.sku !== 'string' || !row.sku.trim() || row.sku.length > 200
      || typeof row.name !== 'string' || !row.name.trim() || row.name.length > 1000
      || typeof row.ecwid_option_signature !== 'string' || row.ecwid_option_signature.length > 20000) throw invalid();
    let parsed: unknown;
    try { parsed = JSON.parse(row.ecwid_option_signature); } catch { throw invalid(); }
    const options = canonicalVariationOptions(parsed);
    if (options === null || JSON.stringify(options) !== row.ecwid_option_signature) throw invalid();
    const custom=['TARGET','PARENT_IF_VARIATION_BLANK'].includes(String(row.sku_source))&&row.option_policy==='STOCK_SELECTION_PLUS_OPAQUE_EXTRAS';
    if(!custom&&((row.sku_source!==undefined&&row.sku_source!=='TARGET')||(row.option_policy!==undefined&&row.option_policy!=='EXACT')))throw invalid();
    if(custom&&(!row.ecwid_combination_id||options.length===0))throw invalid();
    const target: WorkbookTarget = { ecwid_product_id: row.ecwid_product_id,
      ecwid_combination_id: row.ecwid_combination_id, ecwid_option_signature: row.ecwid_option_signature,
      sku: normalizeCode(row.sku), name: row.name.trim(),
      ...(custom?{sku_source:row.sku_source as 'TARGET'|'PARENT_IF_VARIATION_BLANK',option_policy:'STOCK_SELECTION_PLUS_OPAQUE_EXTRAS' as const}:{}) };
    const id = workbookTargetId(target);
    const sameSku=skus.get(target.sku);
    if(ids.has(id)||(sameSku&&!(opaqueWorkbookPolicy(sameSku)&&opaqueWorkbookPolicy(target)
      &&sameSku.sku_source==='PARENT_IF_VARIATION_BLANK'&&target.sku_source==='PARENT_IF_VARIATION_BLANK'
      &&sameSku.ecwid_product_id===target.ecwid_product_id)))throw invalid();
    ids.add(id); skus.set(target.sku,target);
    return target;
  });
}

/** Include these statements in the SAME atomic batch as reviewed opening orders. */
export function buildWorkbookTargetStatements(db: D1Database, value: WorkbookTarget[], review: WorkbookReview): D1PreparedStatement[] {
  const targets = validateWorkbookTargets(value);
  if (!review.reference?.trim() || review.reference.length > 1000 || !review.actor?.trim() || review.actor.length > 320) {
    throw new DomainError(400, 'WORKBOOK_REVIEW_REQUIRED', 'An administrator and review reference are required for workbook-managed targets.');
  }
  const timestamp = review.timestamp ?? new Date().toISOString();
  if (!Number.isFinite(Date.parse(timestamp))) throw new DomainError(400, 'INVALID_REVIEW_TIMESTAMP', 'A valid review timestamp is required.');
  if (!targets.length) return [];
  const rows = targets.map(target => ({...target,id:workbookTargetId(target),sku_source:target.sku_source??'TARGET',option_policy:target.option_policy??'EXACT'}));
  return [db.prepare(`INSERT INTO workbook_managed_targets
    (id,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,sku,name,review_reference,reviewed_by,reviewed_at,sku_source,option_policy)
    SELECT json_extract(j.value,'$.id'),json_extract(j.value,'$.ecwid_product_id'),json_extract(j.value,'$.ecwid_combination_id'),
      json_extract(j.value,'$.ecwid_option_signature'),json_extract(j.value,'$.sku'),json_extract(j.value,'$.name'),?,?,?,
      json_extract(j.value,'$.sku_source'),json_extract(j.value,'$.option_policy')
    FROM json_each(?) j WHERE 1 ON CONFLICT(id) DO NOTHING`)
    .bind(review.reference.trim(), review.actor.trim().toLowerCase(), timestamp, JSON.stringify(rows))];
}

export async function registerWorkbookTargets(db: D1Database, targets: WorkbookTarget[], review: WorkbookReview): Promise<void> {
  const statements = buildWorkbookTargetStatements(db, targets, review);
  if (statements.length) await db.batch(statements);
}
