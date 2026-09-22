import { canonicalVariationOptions, type EcwidOrder, type EcwidOrderLine, type EcwidProduct, type EcwidVariationOption } from './ecwid';
import type { WorkbookTarget } from './pilot-scope';

const SHA256=/^[a-f0-9]{64}$/;
export function opaqueWorkbookPolicy(target: WorkbookTarget): boolean {
  return ['TARGET','PARENT_IF_VARIATION_BLANK'].includes(target.sku_source??'TARGET') && target.option_policy==='STOCK_SELECTION_PLUS_OPAQUE_EXTRAS';
}
export function validWorkbookEvidence(value: unknown): value is NonNullable<EcwidOrderLine['workbookOptionsEvidence']> {
  if(!value||typeof value!=='object'||Array.isArray(value))return false;
  const row=value as Record<string,unknown>;
  return Object.keys(row).sort().join(',')==='kind,sha256' && row.kind==='OPAQUE_OPTIONS_V1'
    && typeof row.sha256==='string' && SHA256.test(row.sha256);
}

/** Canonical JSON for fingerprints only. No customer values are returned or logged. */
function canonicalJson(value: unknown,depth=0): unknown {
  if(depth>30)throw new Error('Order options exceed the supported nesting depth.');
  if(value===null||typeof value==='string'||typeof value==='boolean')return value;
  if(typeof value==='number'&&Number.isFinite(value))return value;
  if(Array.isArray(value)){
    if(value.length>10000)throw new Error('Order options exceed the supported array size.');
    return value.map(entry=>canonicalJson(entry,depth+1));
  }
  if(value&&typeof value==='object'){
    const row=value as Record<string,unknown>;
    return Object.fromEntries(Object.keys(row).sort().map(key=>[key,canonicalJson(row[key],depth+1)]));
  }
  throw new Error('Order options must be ordinary JSON values.');
}
export async function opaqueOptionsHash(options: unknown): Promise<string> {
  if(!Array.isArray(options)||options.length>100)throw new Error('Invalid opaque order option evidence.');
  const payload=JSON.stringify(canonicalVariationOptions(options)??canonicalJson(options));
  if(new TextEncoder().encode(payload).byteLength>256000)throw new Error('Order option evidence is too large.');
  const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(payload));
  return Array.from(new Uint8Array(digest),byte=>byte.toString(16).padStart(2,'0')).join('');
}

/** Existing supported APP signatures stay unchanged. Redacted hashes are trusted
 * only in already-reviewed snapshot paths, never live upserts or API responses. */
export async function orderOptionsHashMaterial(line: EcwidOrderLine,allowSnapshotEvidence=false,forceOpaque=false):Promise<unknown> {
  if(line.workbookOptionsEvidence!==undefined){
    if(!allowSnapshotEvidence||!validWorkbookEvidence(line.workbookOptionsEvidence)
      ||canonicalVariationOptions(line.selectedOptions)===null)throw new Error('Untrusted redacted order option evidence.');
    return ['OPAQUE_OPTIONS_V1',line.workbookOptionsEvidence.sha256];
  }
  if(forceOpaque)return ['OPAQUE_OPTIONS_V1',await opaqueOptionsHash(line.selectedOptions)];
  return canonicalVariationOptions(line.selectedOptions)??['OPAQUE_OPTIONS_V1',await opaqueOptionsHash(line.selectedOptions)];
}

/** Keep only exact catalogue stock-option names. Extra cutting instructions are
 * outside the app, but missing/duplicate/contradictory stock selections are not. */
export function stockOptionSubset(options:unknown,expected:EcwidVariationOption[]):EcwidVariationOption[]|null {
  if(!Array.isArray(options)||options.length>100||!expected.length)return null;
  const names=new Set(expected.map(option=>option.name));
  const selected=options.filter(value=>value&&typeof value==='object'&&!Array.isArray(value)
    &&names.has((value as Record<string,unknown>).name as string));
  const canonical=canonicalVariationOptions(selected);
  return canonical!==null&&JSON.stringify(canonical)===JSON.stringify(expected)?canonical:null;
}

export function workbookLineMatches(line:EcwidOrderLine,target:WorkbookTarget,allowSnapshotEvidence=false):boolean {
  // Ecwid's digital flag means downloadable attachments are present. An exact
  // reviewed WORKBOOK target stays external even when it includes such files.
  if(typeof line.digital!=='boolean'||line.productId!==target.ecwid_product_id||line.combinationId!==target.ecwid_combination_id
    ||line.sku!==target.sku)return false;
  if(line.workbookOptionsEvidence!==undefined&&(!allowSnapshotEvidence||!validWorkbookEvidence(line.workbookOptionsEvidence)))return false;
  if(!opaqueWorkbookPolicy(target))return line.workbookOptionsEvidence===undefined
    &&JSON.stringify(canonicalVariationOptions(line.selectedOptions))===target.ecwid_option_signature;
  let expected:EcwidVariationOption[]|null;
  try{expected=canonicalVariationOptions(JSON.parse(target.ecwid_option_signature));}catch{return false;}
  if(expected===null||!expected.length||!line.combinationId)return false;
  const subset=stockOptionSubset(line.selectedOptions,expected);
  if(subset===null)return false;
  // A sanitized snapshot may only carry the stock subset, never arbitrary extras
  // alongside an opaque claim. Raw live responses carry their actual options.
  return line.workbookOptionsEvidence===undefined||JSON.stringify(canonicalVariationOptions(line.selectedOptions))===JSON.stringify(subset);
}

/** Produce version-2 opening evidence, not a mapping approval. Unsupported raw
 * fields exist only in memory while hashing; neither text nor file URLs survive. */
export async function sanitizedOpeningOrder(order:EcwidOrder,targets:EcwidProduct[]):Promise<EcwidOrder> {
  return {...order,items:await Promise.all(order.items.map(async line=>{
    if(line.workbookOptionsEvidence!==undefined)throw new Error('Snapshot input must be a live parsed Ecwid order, not redacted evidence.');
    const supported=canonicalVariationOptions(line.selectedOptions);
    const matches=targets.filter(target=>target.id===line.productId&&(target.combinationId??null)===line.combinationId);
    const expected=matches.length===1?canonicalVariationOptions(matches[0].variationOptions):null;
    const subset=expected?stockOptionSubset(line.selectedOptions,expected):null;
    if(supported!==null&&(!subset||(matches[0].hasExtraOptions!==true&&JSON.stringify(supported)===JSON.stringify(subset))))return {...line,selectedOptions:supported};
    if(!subset)return {...line,selectedOptions:[{name:'Unsupported selection',value:'Redacted',type:'REDACTED'}]};
    return {...line,selectedOptions:subset,workbookOptionsEvidence:{kind:'OPAQUE_OPTIONS_V1' as const,sha256:await opaqueOptionsHash(line.selectedOptions)}};
  }))};
}
