import Database from 'better-sqlite3';
import { chmod, mkdtemp, readFile, readdir, stat, symlink } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath, URL } from 'node:url';
import { describe, expect, it } from 'vitest';
import { buildD1Bootstrap, MIGRATIONS_TABLE_SQL, prepareD1Bootstrap, type MigrationSource } from '../scripts/prepare-d1-bootstrap';
import { applyMigrations } from './d1';

const migrationDir=fileURLToPath(new URL('../migrations/',import.meta.url));
async function sources(): Promise<MigrationSource[]> {
  return Promise.all((await readdir(migrationDir)).map(async name=>({name,sql:await readFile(join(migrationDir,name))})));
}
function execute(db: Database.Database,sql:Uint8Array) {
  db.pragma('foreign_keys=ON');
  db.transaction(()=>db.exec(Buffer.from(sql).toString('utf8')))();
}
const schema=(db:Database.Database)=>db.prepare(`SELECT type,name,tbl_name,sql FROM sqlite_schema
  WHERE name NOT GLOB 'sqlite_*' AND name NOT IN ('d1_migrations','_cf_KV') ORDER BY type,name`).all();

describe('empty-only D1 bootstrap artifact',()=>{
  it('preserves all eight migrations byte-for-byte with deterministic names and the normal Wrangler journal',async()=>{
    const input=await sources();expect(input).toHaveLength(8);
    const first=buildD1Bootstrap(input);const second=buildD1Bootstrap([...input].reverse());
    expect(first.sql.equals(second.sql)).toBe(true);expect(first.sha256).toBe(second.sha256);
    let offset=0;
    for(const source of [...input].sort((a,b)=>a.name.localeCompare(b.name))) {
      const at=first.sql.indexOf(source.sql,offset);expect(at).toBeGreaterThanOrEqual(offset);offset=at+source.sql.length;
      expect(first.sql.subarray(at,offset)).toEqual(Buffer.from(source.sql));
    }
    const db=new Database(':memory:');const expected=new Database(':memory:');
    try {
      execute(db,first.sql);applyMigrations(expected);
      expect(schema(db)).toEqual(schema(expected));expect(db.pragma('foreign_key_check')).toEqual([]);
      const journal=db.prepare('SELECT id,name,applied_at FROM d1_migrations ORDER BY id').all() as {id:number;name:string;applied_at:string}[];
      expect(journal.map(row=>row.name)).toEqual(first.migrations.map(row=>row.name));
      expect(journal.map(row=>row.id)).toEqual([1,2,3,4,5,6,7,8]);expect(journal.every(row=>Number.isFinite(Date.parse(row.applied_at)))).toBe(true);
      expect(() => db.exec("INSERT INTO items(id,sku,name,scan_code,inventory_mode,active) VALUES('s','SUP','Supplier','SUP','SUPPLIER_BACKED_UNLIMITED',1)")).toThrow('SUPPLIER_OPENING_REQUIRED');
      db.exec("INSERT INTO items(id,sku,name,scan_code,ecwid_product_id) VALUES('a','APP','App','APP','100')");
      expect(()=>db.exec(`INSERT INTO workbook_managed_targets(id,sku,name,ecwid_product_id,ecwid_option_signature,review_reference,reviewed_by,reviewed_at)
        VALUES('w','APP','Conflict','200','[]','review','admin','now')`)).toThrow('WORKBOOK_APP_IDENTITY_CONFLICT');
    } finally {db.close();expected.close();}
  });
  it('allows only an empty normal journal and Cloudflare internal table',async()=>{
    const db=new Database(':memory:');
    try {
      db.exec(`${MIGRATIONS_TABLE_SQL} CREATE TABLE _cf_KV(key TEXT PRIMARY KEY,value BLOB) WITHOUT ROWID;`);
      execute(db,buildD1Bootstrap(await sources()).sql);
      expect(db.prepare('SELECT COUNT(*) AS n FROM d1_migrations').get()).toEqual({n:8});
      expect(db.prepare("SELECT COUNT(*) AS n FROM sqlite_schema WHERE name='_cf_KV'").get()).toEqual({n:1});
    } finally {db.close();}
  });
  it.each([
    'CREATE TABLE business(id TEXT PRIMARY KEY);',
    'CREATE VIEW existing_view AS SELECT 1;',
    `${MIGRATIONS_TABLE_SQL} INSERT INTO d1_migrations(name) VALUES('0001_inventory.sql');`,
    'CREATE TABLE d1_migrations(id INTEGER,name TEXT,applied_at TEXT);',
    'CREATE TABLE sqliteXbusiness(id TEXT);',
  ])('rejects nonempty or incompatible target without changing it: %s',async setup=>{
    const db=new Database(':memory:');
    try {
      db.exec(setup);const before=db.prepare('SELECT * FROM sqlite_schema ORDER BY name').all();
      expect(()=>execute(db,buildD1Bootstrap(awaitedSources).sql)).toThrow();
      expect(db.prepare('SELECT * FROM sqlite_schema ORDER BY name').all()).toEqual(before);
    } finally {db.close();}
  });
  it('rolls back all schema and journal writes when a later original statement fails',()=>{
    const db=new Database(':memory:');
    try {
      const artifact=buildD1Bootstrap([{name:'0001_example.sql',sql:Buffer.from('CREATE TABLE earlier(id TEXT);')},
        {name:'0002_failure.sql',sql:Buffer.from('SELECT * FROM missing_business_table;')}]);
      expect(()=>execute(db,artifact.sql)).toThrow('no such table');expect(db.prepare('SELECT name FROM sqlite_schema').all()).toEqual([]);
    } finally {db.close();}
  });
  it.each(([
    [],[{name:'../0001_bad.sql',sql:Buffer.from('SELECT 1;')}],
    [{name:'0001_bad.SQL',sql:Buffer.from('SELECT 1;')}],
    [{name:'0002_gap.sql',sql:Buffer.from('SELECT 1;')}],
    [{name:'0001_a.sql',sql:Buffer.from('SELECT 1;')},{name:'0001_b.sql',sql:Buffer.from('SELECT 1;')}],
    [{name:'0001_empty.sql',sql:Buffer.from(' ')}],
    [{name:'0001_nul.sql',sql:Buffer.from([0])}],
    [{name:'0001_invalid.sql',sql:Buffer.from([0xff])}],
  ] satisfies MigrationSource[][]).map(input=>({input})))('rejects malformed, incomplete or duplicate migration sources',({input})=>{
    expect(()=>buildD1Bootstrap(input)).toThrow();
  });
  it('exclusively creates a 0600 artifact inside the explicit private output directory',async()=>{
    const directory=await mkdtemp(join(tmpdir(),'inventory-bootstrap-'));
    await chmod(directory,0o700);const output=join(directory,'bootstrap.sql');
    const result=await prepareD1Bootstrap(migrationDir,output);
    expect(result).toMatchObject({output,remote_actions:0,database_changed:false});
    expect((await stat(output)).mode & 0o777).toBe(0o600);
    expect(await readFile(output)).toEqual(buildD1Bootstrap(await sources()).sql);
    await expect(prepareD1Bootstrap(migrationDir,output)).rejects.toThrow('EEXIST');
    expect((await readFile(output)).length).toBe(result.bytes);
  });
  it('rejects public output directories and symlink migration inputs',async()=>{
    const directory=await mkdtemp(join(tmpdir(),'inventory-bootstrap-invalid-'));
    await chmod(directory,0o755);
    await expect(prepareD1Bootstrap(migrationDir,join(directory,'bootstrap.sql'))).rejects.toThrow('private');
    await chmod(directory,0o700);
    const migrationFolder=await mkdtemp(join(tmpdir(),'inventory-bootstrap-symlink-'));
    await symlink(join(migrationDir,'0001_inventory.sql'),join(migrationFolder,'0001_inventory.sql'));
    await expect(prepareD1Bootstrap(migrationFolder,join(directory,'bootstrap.sql'))).rejects.toThrow('non-regular');
    await expect(prepareD1Bootstrap(migrationDir,'relative.sql')).rejects.toThrow('absolute');
  });
});

// Loaded once solely from checked-in migration files. No credentials, DB files,
// fixture data or remote endpoints participate in artifact tests.
const awaitedSources=await sources();
