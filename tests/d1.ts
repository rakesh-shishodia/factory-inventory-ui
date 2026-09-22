import Database from 'better-sqlite3';
import { readFileSync, readdirSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';

/** Like D1, migrate under a transaction with foreign keys enabled throughout. */
export function applyMigrations(sqlite: Database.Database): void {
  sqlite.pragma('foreign_keys = ON');
  const directory = new NodeURL('../migrations/', import.meta.url);
  for (const file of readdirSync(directory).filter(name => /^\d+.*\.sql$/.test(name)).sort()) {
    sqlite.transaction(() => sqlite.exec(readFileSync(new NodeURL(file, directory), 'utf8')))();
  }
}

/** Real SQLite with the prepared statement and atomic batch subset used by the app. */
export function sqliteD1(sqlite: Database.Database): D1Database {
  class Statement {
    constructor(readonly sql: string, readonly values: unknown[] = []) {}
    bind(...values: unknown[]) { return new Statement(this.sql, values); }
    execute() {
      const statement = sqlite.prepare(this.sql);
      if (statement.reader) return { success: true, results: statement.all(...this.values), meta: { changes: 0 } };
      return { success: true, results: [], meta: { changes: statement.run(...this.values).changes } };
    }
    async all() { return this.execute(); }
    async run() { return this.execute(); }
    async first(column?: string) {
      const row = sqlite.prepare(this.sql).get(...this.values) as Record<string, unknown> | undefined;
      return row ? (column ? row[column] : row) : null;
    }
  }
  // Test boundary: unavailable platform-only methods are intentionally omitted.
  // Production code never uses this adapter or its type assertion.
  return {
    prepare: (sql: string) => new Statement(sql),
    batch: async (statements: Statement[]) => sqlite.transaction(() => statements.map(statement => statement.execute()))(),
  } as unknown as D1Database;
}
