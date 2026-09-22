import { createHash } from 'node:crypto';
import { lstat, open, readdir, readFile } from 'node:fs/promises';
import { dirname, isAbsolute, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';

/** Matches Wrangler 4's getCreateMigrationsTableQuery; never replace its journal. */
export const MIGRATIONS_TABLE_SQL = `CREATE TABLE IF NOT EXISTS "d1_migrations"(
		id         INTEGER PRIMARY KEY AUTOINCREMENT,
		name       TEXT UNIQUE,
		applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL
);`;

// File import owns the transaction. D1 rejects nested BEGIN/COMMIT wrappers.
// json_extract is evaluated only on failure: malformed JSON aborts the import
// without introducing a disposable table/trigger or modifying business data.
const HEADER = `-- EMPTY-DATABASE BOOTSTRAP ONLY. Import this entire file atomically.
-- A malformed-JSON error in these guards means STOP: the target is not empty.
SELECT CASE WHEN EXISTS (
  SELECT 1 FROM sqlite_schema WHERE name NOT GLOB 'sqlite_*'
    AND name NOT IN ('_cf_KV','d1_migrations')
) THEN json_extract('D1_BOOTSTRAP_REQUIRES_EMPTY_DATABASE','$') ELSE 1 END;
${MIGRATIONS_TABLE_SQL}
SELECT CASE WHEN EXISTS(SELECT 1 FROM "d1_migrations")
  THEN json_extract('D1_BOOTSTRAP_REQUIRES_EMPTY_JOURNAL','$') ELSE 1 END;
-- Refuse a different object masquerading as the standard Wrangler journal.
SELECT CASE WHEN (SELECT COUNT(*) FROM pragma_table_info('d1_migrations'))<>3
  OR NOT EXISTS(SELECT 1 FROM pragma_table_info('d1_migrations') WHERE name='id' AND type='INTEGER' AND pk=1)
  OR NOT EXISTS(SELECT 1 FROM pragma_table_info('d1_migrations') WHERE name='name' AND type='TEXT' AND pk=0)
  OR NOT EXISTS(SELECT 1 FROM pragma_table_info('d1_migrations') WHERE name='applied_at' AND type='TIMESTAMP'
    AND "notnull"=1 AND dflt_value='CURRENT_TIMESTAMP')
  OR NOT EXISTS(SELECT 1 FROM pragma_index_list('d1_migrations') p
    WHERE p."unique"=1 AND p.origin='u' AND (SELECT COUNT(*) FROM pragma_index_info(p.name))=1
      AND EXISTS(SELECT 1 FROM pragma_index_info(p.name) WHERE name='name'))
  THEN json_extract('D1_BOOTSTRAP_INVALID_JOURNAL_SCHEMA','$') ELSE 1 END;
`;
const MAX_FILE_BYTES = 2_000_000;
const MAX_TOTAL_BYTES = 10_000_000;
const digest = (data: Uint8Array) => createHash('sha256').update(data).digest('hex');
export interface MigrationSource { name: string; sql: Uint8Array }

/** Original migration bytes are concatenated verbatim, never parsed or rewritten. */
export function buildD1Bootstrap(input: MigrationSource[]) {
  if (!input.length || input.length > 1000) throw new Error('Provide between 1 and 1,000 ordered migration files.');
  const numbers = new Set<number>();
  const sources = input.map(source => {
    const match = /^(\d{4})_[a-z0-9][a-z0-9_-]*\.sql$/.exec(source.name);
    if (!match || numbers.has(Number(match[1]))) throw new Error('Malformed or duplicate migration name/number.');
    const number = Number(match[1]); numbers.add(number);
    const bytes = Buffer.from(source.sql);
    if (!bytes.length || bytes.length > MAX_FILE_BYTES || bytes.includes(0)) throw new Error('Migration SQL is empty, too large or contains a NUL byte.');
    if (!new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes).trim()) throw new Error('Migration SQL is empty.');
    return { name: source.name, number, bytes };
  }).sort((a,b) => a.number-b.number);
  if (sources.some((source,index) => source.number!==index+1)) throw new Error('Bootstrap requires the complete contiguous migration history starting at 0001.');
  const chunks = [Buffer.from(HEADER)];
  for (const source of sources) {
    chunks.push(Buffer.from(`\n-- Original migration: ${source.name}\n`), source.bytes,
      // This journal INSERT mirrors Wrangler's buildMigrationQuery.
      Buffer.from(`\nINSERT INTO "d1_migrations" (name)\nvalues ('${source.name}');\n`));
  }
  const sql = Buffer.concat(chunks);
  if (sql.length > MAX_TOTAL_BYTES) throw new Error('Bootstrap artifact exceeds the supported size.');
  return { sql, sha256: digest(sql), migrations: sources.map(source => ({ name: source.name, sha256: digest(source.bytes), bytes: source.bytes.length })) };
}

/** Creates one private LOCAL artifact. Does not connect to or execute any DB. */
export async function prepareD1Bootstrap(migrationsDir: string, output: string) {
  if (!isAbsolute(output) || !output.endsWith('.sql')) throw new Error('Provide an explicit absolute output .sql path.');
  const parent = await lstat(dirname(output));
  if (!parent.isDirectory() || parent.isSymbolicLink() || (parent.mode & 0o077)!==0) {
    throw new Error('The existing output directory must be private (0700) and must not be a symlink.');
  }
  const entries = await readdir(migrationsDir, { withFileTypes: true });
  const sources: MigrationSource[] = [];
  for (const entry of entries) {
    if (!entry.isFile() || !/^\d{4}_[a-z0-9][a-z0-9_-]*\.sql$/.test(entry.name)) throw new Error('Migration directory contains a malformed file or non-regular entry.');
    const path = join(migrationsDir, entry.name);
    const file = await lstat(path);
    if (!file.isFile() || file.isSymbolicLink() || file.size > MAX_FILE_BYTES) throw new Error('Migration files must be bounded regular files, not links.');
    sources.push({ name: entry.name, sql: await readFile(path) });
  }
  const artifact = buildD1Bootstrap(sources);
  const file = await open(output, 'wx', 0o600);
  try { await file.writeFile(artifact.sql); await file.sync(); } finally { await file.close(); }
  return { output, sha256: artifact.sha256, bytes: artifact.sql.length, migrations: artifact.migrations,
    remote_actions: 0, database_changed: false };
}

async function main() {
  const args = parseArgs({ options: { migrations: { type:'string' }, output: { type:'string' }, help: { type:'boolean' } }, strict:true, allowPositionals:false });
  if (args.values.help) {
    console.log('Usage: node --import tsx scripts/prepare-d1-bootstrap.ts --migrations ./migrations --output /absolute/private-directory/bootstrap.sql');
    return;
  }
  if (!args.values.migrations || !args.values.output) throw new Error('Both --migrations and --output are required; use --help.');
  console.log(JSON.stringify(await prepareD1Bootstrap(resolve(args.values.migrations),args.values.output),null,2));
}
if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main().catch(error => { console.error(error instanceof Error ? error.message : 'Bootstrap preparation failed.'); process.exitCode=1; });
}
