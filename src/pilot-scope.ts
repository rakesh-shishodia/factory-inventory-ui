import { canonicalVariationOptions } from './ecwid';
import { DomainError, normalizeCode } from './domain';

/** Explicitly reviewed physical targets kept outside the app during the pilot. */
export interface WorkbookTarget {
  ecwid_product_id: string;
  ecwid_combination_id: string | null;
  ecwid_option_signature: string;
  sku: string;
  name: string;
}

export interface WorkbookReview { reference: string; actor: string; timestamp?: string }

export function workbookTargetId(target: WorkbookTarget): string {
  return `workbook:${target.ecwid_product_id}:${target.ecwid_combination_id ?? 'simple'}`;
}

export function validateWorkbookTargets(value: unknown): WorkbookTarget[] {
  const invalid = () => new DomainError(400, 'INVALID_WORKBOOK_TARGETS', 'Workbook-managed targets need an explicit, unique SKU and exact reviewed Ecwid identity.');
  if (!Array.isArray(value) || value.length > 500) throw invalid();
  const ids = new Set<string>();
  const skus = new Set<string>();
  return value.map(raw => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw invalid();
    const row = raw as Record<string, unknown>;
    if (typeof row.ecwid_product_id !== 'string' || !/^[1-9]\d{0,30}$/.test(row.ecwid_product_id)
      || (row.ecwid_combination_id !== null && (typeof row.ecwid_combination_id !== 'string' || !/^[1-9]\d{0,30}$/.test(row.ecwid_combination_id)))
      || typeof row.sku !== 'string' || !row.sku.trim() || row.sku.length > 200
      || typeof row.name !== 'string' || !row.name.trim() || row.name.length > 1000
      || typeof row.ecwid_option_signature !== 'string' || row.ecwid_option_signature.length > 20000) throw invalid();
    let parsed: unknown;
    try { parsed = JSON.parse(row.ecwid_option_signature); } catch { throw invalid(); }
    const options = canonicalVariationOptions(parsed);
    if (options === null || JSON.stringify(options) !== row.ecwid_option_signature) throw invalid();
    const target: WorkbookTarget = { ecwid_product_id: row.ecwid_product_id,
      ecwid_combination_id: row.ecwid_combination_id, ecwid_option_signature: row.ecwid_option_signature,
      sku: normalizeCode(row.sku), name: row.name.trim() };
    const id = workbookTargetId(target);
    if (ids.has(id) || skus.has(target.sku)) throw invalid();
    ids.add(id); skus.add(target.sku);
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
  const rows = targets.map(target => ({...target,id:workbookTargetId(target)}));
  return [db.prepare(`INSERT INTO workbook_managed_targets
    (id,ecwid_product_id,ecwid_combination_id,ecwid_option_signature,sku,name,review_reference,reviewed_by,reviewed_at)
    SELECT json_extract(j.value,'$.id'),json_extract(j.value,'$.ecwid_product_id'),json_extract(j.value,'$.ecwid_combination_id'),
      json_extract(j.value,'$.ecwid_option_signature'),json_extract(j.value,'$.sku'),json_extract(j.value,'$.name'),?,?,?
    FROM json_each(?) j WHERE 1 ON CONFLICT(id) DO NOTHING`)
    .bind(review.reference.trim(), review.actor.trim().toLowerCase(), timestamp, JSON.stringify(rows))];
}

export async function registerWorkbookTargets(db: D1Database, targets: WorkbookTarget[], review: WorkbookReview): Promise<void> {
  const statements = buildWorkbookTargetStatements(db, targets, review);
  if (statements.length) await db.batch(statements);
}
