import { mkdtemp, mkdir, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { artifactSha256, assembleCutoverRequest, draftCutoverConfirmations, readCutoverFile,
  runCutoverPreparation, writeCutoverArtifacts, type CutoverSourceFiles } from '../scripts/prepare-cutover-request';
import { previewOpeningCutover } from '../src/opening-cutover';

const actor='admin@example.com';
const timestamp='2026-09-22T12:01:00.000Z';
const encode=(value:unknown)=>JSON.stringify(value,null,2)+'\n';
let temporary:string[]=[];
function fixture() {
  const target={id:'123',sku:'001-BOLT',name:'Bolt',quantity:12,unlimited:false,enabled:true,hasOptions:false,hasVariations:false,
    combinationId:null,variationOptions:[],eligibilityVerified:true,hasExtraOptions:false,hasBundleRelationships:false};
  const scope={sku:'001-BOLT',ecwid_product_id:'123',ecwid_combination_id:null,ecwid_option_signature:'[]'};
  return {
    stock:{kind:'SOURCE_CANDIDATES',dry_run:true,source_ref:'Current workbook',source_hash:'a'.repeat(64),
      balance_meaning:'PHYSICAL_ON_HAND',ecwid_checked:false,reservations_confirmed:false,
      rows:[{sku:'001-BOLT',balance:'10',name:'Bolt',single_unit_confirmed:false,source_row:12,source_sheet:'Stock Sheet'},
        {sku:'NOT-APPROVED',balance:999,name:'Unapproved',single_unit_confirmed:false}]},
    catalogue:{kind:'READONLY_CATALOGUE',schema_version:1,dry_run:true,complete:true,store_id:'2442119',
      started_at:'2026-09-22T12:00:01.000Z',completed_at:'2026-09-22T12:00:20.000Z',product_count:1,stock_target_count:1,
      products:[target],stock_targets:[target],reservations_confirmed:false},
    orders:{kind:'READONLY_ORDERS',schema_version:1,dry_run:true,complete:true,store_id:'2442119',
      started_at:'2026-09-22T12:00:01.000Z',completed_at:'2026-09-22T12:00:20.000Z',creation_cutoff:Date.parse('2026-09-22T12:00:01.000Z')/1000,
      orders_checked:100,pending_order_count:1,line_count:1,orders:[{id:'12',paymentStatus:'PAID',fulfillmentStatus:'PROCESSING',
        updatedAt:'2026-09-22T11:00:00.000Z',items:[{id:'111',productId:'123',sku:'001-BOLT',name:'Bolt',quantity:3,
          combinationId:null as string|null,selectedOptions:[] as unknown[],digital:false,trackQuantity:true}]}]},
    scope:{kind:'APPROVED_PILOT_SCOPE',schema_version:1,store_id:'2442119',review_reference:'Explicit approved pilot identity list',scope:[scope]}
  };
}
function filesOf(value=fixture()):CutoverSourceFiles {return {stock:encode(value.stock),catalogue:encode(value.catalogue),orders:encode(value.orders),scope:encode(value.scope)};}
function confirmations(files:CutoverSourceFiles) {
  const draft=draftCutoverConfirmations(files,{actor});
  const {dry_run:_dry,executable:_executable,instruction:_instruction,unmapped_lines:_unmapped,...rest}=draft;
  return {...rest,kind:'PILOT_CUTOVER_CONFIRMATIONS',confirm_staging:true,physical_counts_confirmed:true,
    freeze:{confirmed:true,started_at:'2026-09-22T12:00:00.000Z'},
    unit_confirmations:draft.unit_confirmations.map(row=>({...row,single_unit_confirmed:true})),
    line_confirmations:draft.line_confirmations.map(row=>({...row,previously_picked_quantity:0}))};
}
async function temp() {const directory=await mkdtemp(resolve(tmpdir(),'cutover-assembler-'));temporary.push(directory);return directory;}
beforeEach(()=>{vi.useFakeTimers();vi.setSystemTime(timestamp);vi.stubGlobal('fetch',vi.fn(()=>{throw new Error('Assembler must not use network');}));});
afterEach(async()=>{vi.useRealTimers();vi.unstubAllGlobals();await Promise.all(temporary.map(path=>rm(path,{recursive:true,force:true})));temporary=[];});

describe('explicit fresh cutover request assembly',()=>{
  it('selects only approved targets and derives reservations only after exact zero-pick approvals',async()=>{
    const files=filesOf();const before={...files};const confirmed=encode(confirmations(files));
    const result=await assembleCutoverRequest(files,confirmed,{actor});
    expect(result.request.input.rows).toEqual([{sku:'001-BOLT',balance:'10',name:'Bolt',single_unit_confirmed:true,source_row:12,source_sheet:'Stock Sheet',
      ecwid_product_id:'123',ecwid_combination_id:null,ecwid_option_signature:'[]'}]);
    expect(result.request.input.reservations).toEqual([{sku:'001-BOLT',quantity:3}]);
    expect(result.request.input.snapshot_source_hash).toBe('a'.repeat(64));
    expect(result.review.rows[0]).toMatchObject({physical_on_hand:10,reserved:3,desired_ecwid_quantity:7});
    expect(result.request.expected_hash).toBe((await previewOpeningCutover(result.request,{storeId:'2442119',actor})).review_hash);
    expect(result.manifest).toMatchObject({dry_run:true,staged:false,activated:false,writes_performed:0,actor,
      request_sha256:artifactSha256(encode(result.request)),confirmations_sha256:artifactSha256(confirmed)});
    expect(files).toEqual(before);expect(fetch).not.toHaveBeenCalled();
  });
  it('preserves an explicit UUID across regenerated identical artifacts',async()=>{
    const files=filesOf();const confirmed=encode(confirmations(files));const options={actor,operationId:'12345678-1234-4234-8234-123456789012'};
    const first=await assembleCutoverRequest(files,confirmed,options),second=await assembleCutoverRequest(files,confirmed,options);
    expect(first).toEqual(second);
  });
  it('makes a non-executable draft with no physical, unit or previous-pick assumptions',async()=>{
    const files=filesOf();const draft=draftCutoverConfirmations(files,{actor});
    expect(draft).toMatchObject({kind:'PILOT_CUTOVER_CONFIRMATION_DRAFT',executable:false,confirm_staging:false,physical_counts_confirmed:false,
      freeze:{confirmed:false,started_at:null},unit_confirmations:[{single_unit_confirmed:false}],
      line_confirmations:[{previously_picked_quantity:null}],workbook_scope:[]});
    await expect(assembleCutoverRequest(files,encode(draft),{actor})).rejects.toThrow('draft');
  });
  it('lists unknown outside lines in a draft without silently adding workbook identities',async()=>{
    const source=fixture();source.orders.orders[0].items.push({...source.orders.orders[0].items[0],id:'112',sku:'UNKNOWN',productId:'999'});source.orders.line_count++;
    const files=filesOf(source);const draft=draftCutoverConfirmations(files,{actor});
    expect(draft.workbook_scope).toEqual([]);expect(draft.unmapped_lines).toMatchObject([{sku:'UNKNOWN',status:'EXPLICIT_MAPPING_OR_RECONCILIATION_REQUIRED'}]);
    await expect(assembleCutoverRequest(files,encode(confirmations(files)),{actor})).rejects.toMatchObject({code:'OPENING_ORDER_LINE_UNMAPPED'});
  });
  it('retains explicitly approved workbook lines without inventing app stock or reservations',async()=>{
    const source=fixture();const outside={...source.catalogue.stock_targets[0],id:'456',sku:'WORKBOOK',name:'Workbook item',unlimited:true};
    source.catalogue.stock_targets.push(outside);source.catalogue.products=source.catalogue.stock_targets;source.catalogue.product_count=source.catalogue.stock_target_count=2;
    source.orders.orders[0].items.push({...source.orders.orders[0].items[0],id:'112',sku:'WORKBOOK',productId:'456'});source.orders.line_count++;
    const files=filesOf(source);const confirmed=confirmations(files);
    Object.assign(confirmed,{workbook_scope:[{sku:'WORKBOOK',name:'Workbook item',ecwid_product_id:'456',ecwid_combination_id:null,ecwid_option_signature:'[]'}]});
    const result=await assembleCutoverRequest(files,encode(confirmed),{actor});
    expect(result.review.workbook_line_count).toBe(1);expect(result.request.input.reservations).toEqual([{sku:'001-BOLT',quantity:3}]);
  });
  it.each(['source_hash','stock_sha256','catalogue_sha256','orders_sha256','scope_sha256'])('binds current evidence hash %s',async key=>{
    const files=filesOf(),confirmed=confirmations(files);Object.assign(confirmed,{[key]:'b'.repeat(64)});
    await expect(assembleCutoverRequest(files,encode(confirmed),{actor})).rejects.toThrow(`Confirmed ${key}`);
  });
  it('does not reuse an old confirmation after balances change even if workbook hash was copied',async()=>{
    const original=filesOf(),confirmed=confirmations(original),source=fixture();source.stock.rows[0].balance='11';
    await expect(assembleCutoverRequest(filesOf(source),encode(confirmed),{actor})).rejects.toThrow('stock_sha256');
  });
  it.each(['physical','staging','unit','prior','freeze'])('rejects absent or false %s approval',async field=>{
    const files=filesOf(),confirmed=confirmations(files);
    if(field==='physical') confirmed.physical_counts_confirmed=false;
    if(field==='staging') confirmed.confirm_staging=false;
    if(field==='unit') confirmed.unit_confirmations[0].single_unit_confirmed=false;
    if(field==='prior') Object.assign(confirmed.line_confirmations[0],{previously_picked_quantity:null});
    if(field==='freeze') confirmed.freeze.confirmed=false;
    await expect(assembleCutoverRequest(files,encode(confirmed),{actor})).rejects.toBeInstanceOf(Error);
  });
  it.each(['omitted','extra','duplicate','wrong-product'])('requires exact per-target unit approvals: %s',async kind=>{
    const files=filesOf(),confirmed=confirmations(files);
    if(kind==='omitted') confirmed.unit_confirmations=[];
    if(kind==='extra') confirmed.unit_confirmations.push({...confirmed.unit_confirmations[0],sku:'OTHER'});
    if(kind==='duplicate') confirmed.unit_confirmations.push(confirmed.unit_confirmations[0]);
    if(kind==='wrong-product') confirmed.unit_confirmations[0].ecwid_product_id='999';
    await expect(assembleCutoverRequest(files,encode(confirmed),{actor})).rejects.toThrow('Unit confirmations');
  });
  it.each(['omitted','duplicate','nonzero','other-order'])('requires exact per-line zero-pick approvals: %s',async kind=>{
    const files=filesOf(),confirmed=confirmations(files);
    if(kind==='omitted') confirmed.line_confirmations=[];
    if(kind==='duplicate') confirmed.line_confirmations.push(confirmed.line_confirmations[0]);
    if(kind==='nonzero') confirmed.line_confirmations[0].previously_picked_quantity=1;
    if(kind==='other-order') confirmed.line_confirmations[0].order_id='99';
    await expect(assembleCutoverRequest(files,encode(confirmed),{actor})).rejects.toThrow(/Prior-pick|previous-pick/);
  });
  it('rejects missing/duplicated current source rows rather than falling back to old balances',()=>{
    const source=fixture();source.stock.rows=source.stock.rows.filter(row=>row.sku!=='001-BOLT');
    expect(()=>draftCutoverConfirmations(filesOf(source),{actor})).toThrow('exactly one current source');
    const duplicate=fixture();duplicate.stock.rows.push({...duplicate.stock.rows[0],sku:'001-bolt'});
    expect(()=>draftCutoverConfirmations(filesOf(duplicate),{actor})).toThrow('exactly one current source');
  });
  it('does not overwrite contradictory exact mappings carried by the fresh source',()=>{
    const source=fixture();Object.assign(source.stock.rows[0],{ecwid_product_id:'999'});
    expect(()=>draftCutoverConfirmations(filesOf(source),{actor})).toThrow('mapping conflicts');
  });
  it('uses real-clock authoritative freshness checks without a caller time override',async()=>{
    const files=filesOf(),confirmed=encode(confirmations(files));vi.setSystemTime('2026-09-22T12:16:00.000Z');
    await expect(assembleCutoverRequest(files,confirmed,{actor})).rejects.toMatchObject({code:'OPENING_SNAPSHOT_STALE'});
  });
  it.each(['not-an-email','Admin@example.com',' admin@example.com'])('rejects ambiguous administrator identity %s',value=>{
    expect(()=>draftCutoverConfirmations(filesOf(),{actor:value})).toThrow('administrator');
  });
  it('binds approval to the exact administrator and refuses extra confirmation fields',async()=>{
    const files=filesOf(),confirmed=confirmations(files);
    await expect(assembleCutoverRequest(files,encode({...confirmed,confirmed_by:'other@example.com'}),{actor})).rejects.toThrow('exact administrator');
    await expect(assembleCutoverRequest(files,encode({...confirmed,now:timestamp}),{actor})).rejects.toThrow('confirmations');
  });
  it('rejects overlarge input and malformed source JSON',()=>{
    expect(()=>draftCutoverConfirmations({...filesOf(),stock:' '.repeat(2_000_001)},{actor})).toThrow('size limit');
    expect(()=>draftCutoverConfirmations({...filesOf(),scope:'{'},{actor})).toThrow('valid JSON');
  });
});

describe('private exclusive artifacts and CLI boundaries',()=>{
  it('runs the actual draft CLI and emits only a non-executable private worksheet',async()=>{
    const root=await temp();await mkdir(resolve(root,'import-data'));await mkdir(resolve(root,'public'));
    const files=filesOf();
    for(const [name,body] of Object.entries(files))await writeFile(resolve(root,`${name}.json`),body);
    const command=[
      '--import',resolve('node_modules/tsx/dist/loader.mjs'),resolve('scripts/prepare-cutover-request.ts'),
      '--stock','stock.json','--catalog','catalogue.json','--orders','orders.json','--scope','scope.json',
      '--actor',actor,'--draft','--out','import-data/draft',
    ];
    const result=await promisify(execFile)(process.execPath,command,{cwd:root});
    expect(JSON.parse(result.stdout)).toMatchObject({dry_run:true,executable:false,writes_performed:0});
    const draft=JSON.parse(await readFile(resolve(root,'import-data/draft/confirmations-draft.json'),'utf8'));
    expect(draft).toMatchObject({physical_counts_confirmed:false,confirm_staging:false,workbook_scope:[]});
    await expect(stat(resolve(root,'import-data/draft/cutover-request.json'))).rejects.toMatchObject({code:'ENOENT'});
    await expect(promisify(execFile)(process.execPath,command,{cwd:root})).rejects.toThrow();
  });
  it('writes private exclusive artifacts and refuses to overwrite existing output',async()=>{
    const root=await temp(),privateRoot=resolve(root,'import-data'),publicRoot=resolve(root,'public');
    await mkdir(privateRoot);await mkdir(publicRoot);const output=resolve(privateRoot,'prepared');
    await writeCutoverArtifacts(privateRoot,output,{'request.json':{dry_run:true}},publicRoot);
    expect((await stat(output)).mode&0o777).toBe(0o700);expect((await stat(resolve(output,'request.json'))).mode&0o777).toBe(0o600);
    expect(JSON.parse(await readFile(resolve(output,'request.json'),'utf8'))).toEqual({dry_run:true});
    await expect(writeCutoverArtifacts(privateRoot,output,{'request.json':{changed:true}},publicRoot)).rejects.toThrow();
    expect(JSON.parse(await readFile(resolve(output,'request.json'),'utf8'))).toEqual({dry_run:true});
  });
  it('refuses public/outside destinations, symlink escape and path traversal filenames',async()=>{
    const root=await temp(),privateRoot=resolve(root,'import-data'),publicRoot=resolve(root,'public');
    await mkdir(privateRoot);await mkdir(publicRoot);await symlink(publicRoot,resolve(privateRoot,'escape'));
    for(const output of [resolve(publicRoot,'leak'),resolve(root,'outside'),resolve(privateRoot,'escape','leak')]){
      await expect(writeCutoverArtifacts(privateRoot,output,{'request.json':{}},publicRoot)).rejects.toThrow('beneath import-data');
    }
    await expect(writeCutoverArtifacts(privateRoot,resolve(privateRoot,'bad'),{'../leak.json':{}},publicRoot)).rejects.toThrow('filename');
    const linked=resolve(root,'linked');await symlink(privateRoot,linked);
    await expect(writeCutoverArtifacts(linked,resolve(linked,'new'),{'request.json':{}},publicRoot)).rejects.toThrow('symlink');
  });
  it('bounds UTF-8 input and rejects non-file/invalid UTF-8 content',async()=>{
    const root=await temp(),file=resolve(root,'input.json');await writeFile(file,'12345');
    await expect(readCutoverFile(file,4)).rejects.toThrow('size limit');expect(await readCutoverFile(file,5)).toBe('12345');
    await expect(readCutoverFile(root,5)).rejects.toThrow('regular files');
    await writeFile(file,Buffer.from([0xff]));await expect(readCutoverFile(file,5)).rejects.toThrow();
  });
  it('offers CLI help without sources and rejects time, credential or executable actions',async()=>{
    expect(await runCutoverPreparation(['--help'])).toMatchObject({help:expect.stringContaining('--draft')});
    for(const flag of ['--now','--token','--stage','--deploy'])await expect(runCutoverPreparation([flag,'anything'])).rejects.toThrow();
    await expect(runCutoverPreparation([])).rejects.toThrow('Provide all source');
    expect(fetch).not.toHaveBeenCalled();
  });
});
