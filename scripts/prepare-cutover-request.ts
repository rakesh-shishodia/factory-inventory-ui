import { createHash, randomUUID } from 'node:crypto';
import { lstat, mkdir, open, realpath, stat, writeFile } from 'node:fs/promises';
import { basename, resolve, sep } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { canonicalVariationOptions } from '../src/ecwid';
import { normalizeCode } from '../src/domain';
import { preparePreviewInput } from '../src/opening-import';
import { previewOpeningCutover, type OpeningCutoverRequest } from '../src/opening-cutover';
import type { OpeningStageScope } from '../src/opening-apply';

type Row = Record<string, unknown>;
export interface CutoverSourceFiles { stock: string; catalogue: string; orders: string; scope: string }
export interface CutoverAssemblyOptions { actor: string; operationId?: string }
export interface ApprovedPilotScope {
  kind: 'APPROVED_PILOT_SCOPE'; schema_version: 1; store_id: string; review_reference: string;
  scope: OpeningStageScope[];
}
export interface CutoverConfirmations {
  kind: 'PILOT_CUTOVER_CONFIRMATIONS'; schema_version: 1; store_id: string; confirmed_by: string;
  source_hash: string; stock_sha256: string; catalogue_sha256: string; orders_sha256: string; scope_sha256: string;
  confirm_staging: true; physical_counts_confirmed: true;
  freeze: { confirmed: true; started_at: string };
  unit_confirmations: Array<OpeningStageScope & { single_unit_confirmed: true }>;
  line_confirmations: OpeningCutoverRequest['line_confirmations'];
  workbook_scope: OpeningCutoverRequest['workbook_scope'];
}
const MAX_FILE = { stock: 2_000_000, catalogue: 30_000_000, orders: 5_000_000, scope: 500_000, confirmations: 2_000_000 };
const MAX_REQUEST = 10_000_000;
const SHA = /^[a-f0-9]{64}$/;
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const identity = (target: OpeningStageScope) => JSON.stringify([target.sku, target.ecwid_product_id, target.ecwid_combination_id, target.ecwid_option_signature]);
function object(value: unknown, label: string): Row {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${label} must be an object.`);
  return value as Row;
}
function records(value: unknown, label: string, maximum: number): Row[] {
  if (!Array.isArray(value) || value.length > maximum) throw new Error(`${label} must contain at most ${maximum} rows.`);
  return value.map(entry => object(entry,label));
}
function text(value: unknown, label: string, maximum: number): string {
  if (typeof value !== 'string' || !value.trim() || value.length > maximum) throw new Error(`${label} is missing or too long.`);
  return value;
}
export function artifactSha256(value: string): string { return createHash('sha256').update(value,'utf8').digest('hex'); }
function parse(value: string, label: keyof typeof MAX_FILE): Row {
  if (Buffer.byteLength(value,'utf8') > MAX_FILE[label]) throw new Error(`${label} input exceeds its size limit.`);
  try { return object(JSON.parse(value.replace(/^\uFEFF/,'')),label); }
  catch { throw new Error(`${label} must contain a valid JSON object.`); }
}
function actor(value: string): string {
  if (typeof value !== 'string' || value.length > 320 || value !== value.trim().toLowerCase()
    || !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(value)) throw new Error('Use the exact lowercase authenticated administrator email.');
  return value;
}
function target(value: Row): OpeningStageScope {
  if (Object.keys(value).sort().join(',') !== 'ecwid_combination_id,ecwid_option_signature,ecwid_product_id,sku'
    || typeof value.sku !== 'string' || !value.sku || value.sku.length > 200 || value.sku !== normalizeCode(value.sku) || value.sku.includes('|')
    || typeof value.ecwid_product_id !== 'string' || !/^[1-9]\d{0,30}$/.test(value.ecwid_product_id)
    || (value.ecwid_combination_id !== null && (typeof value.ecwid_combination_id !== 'string' || !/^[1-9]\d{0,30}$/.test(value.ecwid_combination_id)))
    || typeof value.ecwid_option_signature !== 'string' || value.ecwid_option_signature.length > 20000) {
    throw new Error('Approved scope needs exact canonical SKU/product/variation/options identities, without extra fields.');
  }
  let options: unknown;
  try { options = JSON.parse(value.ecwid_option_signature); } catch { throw new Error('Approved scope options must be canonical JSON.'); }
  const canonical = canonicalVariationOptions(options);
  if (canonical === null || JSON.stringify(canonical) !== value.ecwid_option_signature) throw new Error('Approved scope options must be canonical supported selections.');
  return { sku: value.sku, ecwid_product_id: value.ecwid_product_id,
    ecwid_combination_id: value.ecwid_combination_id, ecwid_option_signature: value.ecwid_option_signature };
}
function sources(files: CutoverSourceFiles, options: CutoverAssemblyOptions) {
  const authenticatedActor = actor(options.actor);
  const stock = parse(files.stock,'stock'), catalogue = parse(files.catalogue,'catalogue');
  const orders = parse(files.orders,'orders'), approved = parse(files.scope,'scope');
  if (Object.keys(approved).some(key => !['kind','schema_version','store_id','review_reference','scope'].includes(key))
    || approved.kind !== 'APPROVED_PILOT_SCOPE' || approved.schema_version !== 1
    || typeof approved.store_id !== 'string' || !/^[1-9]\d{0,19}$/.test(approved.store_id)) {
    throw new Error('Supply an explicit version-1 APPROVED_PILOT_SCOPE for the expected store.');
  }
  text(approved.review_reference,'Scope review reference',1000);
  const scope = records(approved.scope,'Approved scope',200).map(target);
  if (!scope.length || new Set(scope.map(identity)).size !== scope.length || new Set(scope.map(row => row.sku)).size !== scope.length
    || new Set(scope.map(row => JSON.stringify([row.ecwid_product_id,row.ecwid_combination_id]))).size !== scope.length) {
    throw new Error('Approved scope must identify 1–200 unique independent stock targets.');
  }
  const input = preparePreviewInput(files.stock,'SOURCE_CANDIDATES',catalogue,{store_id:approved.store_id});
  const allSourceRows = records(input.rows,'Fresh source candidates',10000);
  const selectedRows = scope.map(stockTarget => {
    const matches = allSourceRows.filter(row => typeof row.sku === 'string' && normalizeCode(row.sku) === stockTarget.sku);
    if (matches.length !== 1) throw new Error(`Approved SKU ${stockTarget.sku} must have exactly one current source candidate; missing/duplicate rows cannot be substituted.`);
    const row = matches[0];
    // A fresh source file may carry exact mappings. Never overwrite a contradictory one.
    for (const field of ['ecwid_product_id','ecwid_combination_id','ecwid_option_signature'] as const) {
      if (row[field] !== undefined && row[field] !== stockTarget[field]) throw new Error(`Fresh source mapping conflicts with the approved identity for ${stockTarget.sku}.`);
    }
    return { ...row, ...stockTarget };
  });
  if (catalogue.store_id !== approved.store_id || orders.store_id !== approved.store_id) throw new Error('Catalogue, orders and approved scope must identify the same store.');
  const hashes = { source_hash: text(stock.source_hash,'Workbook SHA-256',64).toLowerCase(),
    stock_sha256: artifactSha256(files.stock), catalogue_sha256: artifactSha256(files.catalogue),
    orders_sha256: artifactSha256(files.orders), scope_sha256: artifactSha256(files.scope) };
  if (!SHA.test(hashes.source_hash)) throw new Error('The source workbook digest must be SHA-256.');
  const pending = records(orders.orders,'Opening orders',200);
  const lines = pending.flatMap(order => records(order.items,'Opening order lines',500).map(line => ({ order, line })));
  if (lines.length > 2000) throw new Error('At most 2,000 opening order lines are supported.');
  return { stock, catalogue, orders, approved, scope, input, selectedRows, hashes, lines, authenticatedActor };
}

/** A non-executable worksheet of questions. It contains no true/zero approvals. */
export function draftCutoverConfirmations(files: CutoverSourceFiles, options: CutoverAssemblyOptions) {
  const p = sources(files,options); const keys = new Set(p.scope.map(identity));
  return { kind: 'PILOT_CUTOVER_CONFIRMATION_DRAFT', schema_version: 1, dry_run: true, executable: false,
    instruction: 'Review each fact. To approve, change kind to PILOT_CUTOVER_CONFIRMATIONS, remove dry_run/executable/instruction/unmapped_lines, supply explicit confirmations and reviewed workbook_scope. Never approve a missing/unknown fact by guessing.',
    store_id: p.approved.store_id, confirmed_by: p.authenticatedActor, ...p.hashes,
    confirm_staging: false, physical_counts_confirmed: false, freeze: { confirmed: false, started_at: null },
    unit_confirmations: p.scope.map(row => ({ ...row, single_unit_confirmed: false })),
    line_confirmations: p.lines.map(({order,line}) => ({order_id:order.id,ecwid_line_id:line.id,previously_picked_quantity:null})),
    workbook_scope: [],
    unmapped_lines: p.lines.filter(({line}) => !keys.has(identity({sku:String(line.sku),ecwid_product_id:String(line.productId),
      ecwid_combination_id:typeof line.combinationId==='string'?line.combinationId:null,
      ecwid_option_signature:JSON.stringify(canonicalVariationOptions(line.selectedOptions))})))
      .map(({order,line}) => ({order_id:order.id,ecwid_line_id:line.id,sku:line.sku,ecwid_product_id:line.productId,
        ecwid_combination_id:line.combinationId,ordered_quantity:line.quantity,
        ecwid_option_signature:JSON.stringify(canonicalVariationOptions(line.selectedOptions)),status:'EXPLICIT_MAPPING_OR_RECONCILIATION_REQUIRED'})),
  };
}

/** Recompute the authoritative preview against the real clock. No DB/network or time overrides. */
export async function assembleCutoverRequest(files: CutoverSourceFiles, confirmationText: string, options: CutoverAssemblyOptions) {
  const p = sources(files,options), confirmations = parse(confirmationText,'confirmations');
  const allowed = new Set(['kind','schema_version','store_id','confirmed_by','source_hash','stock_sha256','catalogue_sha256','orders_sha256','scope_sha256',
    'confirm_staging','physical_counts_confirmed','freeze','unit_confirmations','line_confirmations','workbook_scope']);
  if (Object.keys(confirmations).some(key => !allowed.has(key)) || confirmations.kind !== 'PILOT_CUTOVER_CONFIRMATIONS'
    || confirmations.schema_version !== 1 || confirmations.store_id !== p.approved.store_id || confirmations.confirmed_by !== p.authenticatedActor) {
    throw new Error('Provide explicit version-1 cutover confirmations for this store and exact administrator. A draft is not an approval.');
  }
  for (const [name,digest] of Object.entries(p.hashes)) {
    if (confirmations[name] !== digest) throw new Error(`Confirmed ${name} does not match the supplied current artifact. Review the new evidence; do not reuse an old approval.`);
  }
  if (confirmations.confirm_staging !== true || confirmations.physical_counts_confirmed !== true) {
    throw new Error('Explicit current physical-count and staging confirmations are required.');
  }
  const unitConfirmations = records(confirmations.unit_confirmations,'Unit confirmations',200).map(row => {
    if (row.single_unit_confirmed !== true) throw new Error('Every approved target requires its explicit single-unit confirmation.');
    const { single_unit_confirmed: _confirmed, ...stockTarget } = row;
    return target(stockTarget);
  });
  if (JSON.stringify(unitConfirmations.map(identity).sort()) !== JSON.stringify(p.scope.map(identity).sort())) {
    throw new Error('Unit confirmations must match every approved target exactly once, without omitted or newly admitted SKUs.');
  }
  const lineConfirmations = records(confirmations.line_confirmations,'Prior-pick confirmations',2000);
  const expectedKeys = p.lines.map(({order,line}) => `${order.id}:${line.id}`).sort();
  const confirmedKeys = lineConfirmations.map(row => {
    if (Object.keys(row).sort().join(',') !== 'ecwid_line_id,order_id,previously_picked_quantity'
      || typeof row.order_id !== 'string' || typeof row.ecwid_line_id !== 'string' || row.previously_picked_quantity !== 0) {
      throw new Error('Every exact line needs an explicitly confirmed zero previous-pick quantity. Unknown or historical picks cannot be inferred.');
    }
    return `${row.order_id}:${row.ecwid_line_id}`;
  }).sort();
  if (JSON.stringify(confirmedKeys) !== JSON.stringify(expectedKeys)) throw new Error('Prior-pick confirmations must identify all opening order lines exactly once.');
  const reservations = new Map(p.scope.map(row => [identity(row),{sku:row.sku,quantity:0}]));
  for (const {line} of p.lines) {
    const signature = canonicalVariationOptions(line.selectedOptions);
    const key = identity({sku:String(line.sku),ecwid_product_id:String(line.productId),
      ecwid_combination_id:typeof line.combinationId==='string'?line.combinationId:null,ecwid_option_signature:JSON.stringify(signature)});
    const reservation = reservations.get(key);
    if (reservation) {
      if (typeof line.quantity !== 'number' || !Number.isSafeInteger(line.quantity) || line.quantity < 1
        || !Number.isSafeInteger(reservation.quantity+line.quantity) || reservation.quantity+line.quantity > 2147483647) {
        throw new Error('Confirmed opening quantities must be supported positive whole numbers.');
      }
      reservation.quantity += line.quantity;
    }
  }
  const operationId = options.operationId ?? randomUUID();
  if (!UUID.test(operationId)) throw new Error('operation-id must be a UUID retained for exact retries.');
  const request: unknown = { operation_id:operationId.toLowerCase(),expected_hash:'',confirm_staging:confirmations.confirm_staging,
    physical_counts_confirmed:confirmations.physical_counts_confirmed,freeze:confirmations.freeze,
    input:{...p.input,rows:p.selectedRows.map(row=>({...row,single_unit_confirmed:true})),reservations_confirmed:true,
      reservations:[...reservations.values()]},scope:p.scope,orders:p.orders,line_confirmations:lineConfirmations,workbook_scope:confirmations.workbook_scope };
  const review = await previewOpeningCutover(request,{storeId:String(p.approved.store_id),actor:p.authenticatedActor});
  // The authoritative validator above checks the complete untyped artifact.
  const validated = request as OpeningCutoverRequest;
  validated.expected_hash = review.review_hash;
  if (Buffer.byteLength(JSON.stringify(validated),'utf8') > MAX_REQUEST) throw new Error('The complete request exceeds the API 10 MB limit. Do not remove catalogue/order evidence to bypass it.');
  return {request:validated,review,manifest:{kind:'PREPARED_PILOT_CUTOVER',schema_version:1,dry_run:true,
    created_at:new Date().toISOString(),store_id:p.approved.store_id,actor:p.authenticatedActor,operation_id:validated.operation_id,
    review_hash:review.review_hash,scope_review_reference:p.approved.review_reference,...p.hashes,
    confirmations_sha256:artifactSha256(confirmationText),request_sha256:artifactSha256(JSON.stringify(validated,null,2)+'\n'),
    writes_performed:0,staged:false,activated:false}};
}

export async function readCutoverFile(path: string, maximum: number): Promise<string> {
  if (!(await stat(path)).isFile()) throw new Error('Cutover inputs must be regular files.');
  const file=await open(path,'r');
  try {
    if (!(await file.stat()).isFile()) throw new Error('Cutover inputs must be regular files.');
    const bytes=Buffer.alloc(maximum+1);let used=0;
    while(used<bytes.length){const result=await file.read(bytes,used,bytes.length-used,used);if(!result.bytesRead)break;used+=result.bytesRead;}
    if(used>maximum)throw new Error('A cutover input exceeds its size limit.');
    return new TextDecoder('utf-8',{fatal:true,ignoreBOM:true}).decode(bytes.subarray(0,used));
  } finally {await file.close();}
}

/** Private new directory only. Never overwrites existing evidence or writes public assets. */
export async function writeCutoverArtifacts(privateRootPath: string, outputPath: string, artifacts: Record<string,unknown>, publicRootPath?: string) {
  if((await lstat(privateRootPath)).isSymbolicLink())throw new Error('The private artifact root must not be a symlink.');
  const privateRoot=await realpath(privateRootPath),parent=await realpath(resolve(outputPath,'..'));
  if(parent!==privateRoot&&!parent.startsWith(privateRoot+sep))throw new Error('Cutover artifacts must stay beneath import-data/.');
  if(publicRootPath){const publicRoot=await realpath(publicRootPath);
    if(parent===publicRoot||parent.startsWith(publicRoot+sep)||privateRoot===publicRoot||privateRoot.startsWith(publicRoot+sep))throw new Error('Cutover artifacts must never be public assets.');}
  const directory=resolve(parent,basename(resolve(outputPath)));
  if(directory===privateRoot)throw new Error('Use a new subdirectory, not the private artifact root itself.');
  const files=Object.entries(artifacts).map(([name,value])=>{
    if(!/^[a-z][a-z0-9-]*\.json$/.test(name))throw new Error('Invalid artifact filename.');
    const body=JSON.stringify(value,null,2)+'\n';
    if(Buffer.byteLength(body,'utf8')>MAX_REQUEST)throw new Error('Prepared artifact exceeds the supported size limit.');
    return {name,body};
  });
  await mkdir(directory,{mode:0o700});
  for(const {name,body} of files)await writeFile(resolve(directory,name),body,{flag:'wx',mode:0o600});
  return directory;
}

export async function runCutoverPreparation(args: string[]): Promise<Row> {
  const {values,positionals}=parseArgs({args,options:{stock:{type:'string'},catalog:{type:'string'},orders:{type:'string'},scope:{type:'string'},
    confirmations:{type:'string'},actor:{type:'string'},out:{type:'string'},'operation-id':{type:'string'},draft:{type:'boolean'},help:{type:'boolean'}}});
  if(values.help)return{help:'node --import tsx scripts/prepare-cutover-request.ts --stock candidates.json --catalog catalog.json --orders opening-orders.json --scope approved-scope.json --actor admin@example.com --out import-data/NEW --confirmations confirmed.json [--operation-id UUID]\nUse --draft instead of --confirmations to create a non-executable false/null confirmation worksheet. No time override, credentials, network, DB, staging or Ecwid writes are supported.'};
  if(positionals.length||!values.stock||!values.catalog||!values.orders||!values.scope||!values.actor||!values.out
    ||(values.draft?Boolean(values.confirmations||values['operation-id']):!values.confirmations))throw new Error('Provide all source files, --actor, a new private --out, and exactly one of --draft or --confirmations.');
  const [stock,catalogue,orders,scope]=await Promise.all([readCutoverFile(values.stock,MAX_FILE.stock),readCutoverFile(values.catalog,MAX_FILE.catalogue),
    readCutoverFile(values.orders,MAX_FILE.orders),readCutoverFile(values.scope,MAX_FILE.scope)]);
  const files={stock,catalogue,orders,scope},options={actor:values.actor,operationId:values['operation-id']};
  const artifacts:Record<string,unknown>=values.draft?{'confirmations-draft.json':draftCutoverConfirmations(files,options)}:{};
  let summary:Row={dry_run:true,executable:false,writes_performed:0};
  if(!values.draft){const prepared=await assembleCutoverRequest(files,await readCutoverFile(values.confirmations!,MAX_FILE.confirmations),options);
    Object.assign(artifacts,{'cutover-request.json':prepared.request,'cutover-review.json':prepared.review,'artifact-manifest.json':prepared.manifest});
    summary={dry_run:true,prepared:true,staged:false,activated:false,writes_performed:0,operation_id:prepared.request.operation_id,
      review_hash:prepared.review.review_hash,row_count:prepared.review.row_count,order_count:prepared.review.order_count,line_count:prepared.review.line_count};}
  const output=await writeCutoverArtifacts(resolve('import-data'),resolve(values.out),artifacts,resolve('public'));
  return {...summary,output};
}

if(process.argv[1]&&import.meta.url===pathToFileURL(resolve(process.argv[1])).href){
  runCutoverPreparation(process.argv.slice(2)).then(result=>console.log(result.help??JSON.stringify(result,null,2)))
    .catch(error=>{console.error(error instanceof Error?error.message:'Cutover preparation failed.');process.exitCode=1;});
}
