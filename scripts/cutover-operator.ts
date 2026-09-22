import { execFile } from 'node:child_process';
import { constants } from 'node:fs';
import { open, readdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { parseEnv, promisify } from 'node:util';
import { DomainError } from '../src/domain';
import { readBoundedJson, type EcwidFetch } from '../src/ecwid';
import { previewOpeningCutover, stageOpeningCutover } from '../src/opening-cutover';
import { beginCutoverAlignment, alignCutoverRow, finishAndActivateCutover, recoverAndActivateCutover } from '../src/cutover-alignment';

type ObjectValue = Record<string, unknown>;
type Command = 'status' | 'preview' | 'stage' | 'begin' | 'row' | 'rows' | 'finish' | 'recover';
export interface OperatorConfig {
  schema_version: 1; account_id: string; worker_name: string; database_id: string; database_name: string;
  store_id: string; actor: string; expected_version_id: string; credentials_file: string; wrangler_config: string;
}
export interface OperatorCredentials { ecwidToken: string; cloudflareToken: string }
export interface OperatorDatabase { db: D1Database; dispose(): Promise<void> }
export interface OperatorDependencies {
  cloudflareFetch?: EcwidFetch; ecwidFetch?: EcwidFetch;
  connect?: (config: OperatorConfig, credentials: OperatorCredentials) => Promise<OperatorDatabase>;
  /** Trusted terminal reporter only; never receives request input or credentials. */
  onRow?: (receipt: ObjectValue) => void;
  onRecoveryPreflight?: (receipt: ObjectValue) => void;
}
export class OperatorError extends Error {
  constructor(readonly code: string) { super(code); this.name = 'OperatorError'; }
}
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const SHA = /^[a-f0-9]{64}$/;
const COMMANDS: Command[] = ['status', 'preview', 'stage', 'begin', 'row', 'rows', 'finish', 'recover'];
function stop(code: string): never { throw new OperatorError(code); }
function object(value: unknown): ObjectValue {
  if (!value || typeof value !== 'object' || Array.isArray(value)) stop('OPERATOR_INVALID_INPUT');
  return value as ObjectValue;
}
function exactKeys(value: ObjectValue, keys: string[]): void {
  if (Object.keys(value).sort().join(',') !== [...keys].sort().join(',')) stop('OPERATOR_INVALID_FIELDS');
}

/** Configuration is private operator authority, never an HTTP/request body. */
export function parseOperatorConfig(value: unknown): OperatorConfig {
  const input = object(value);
  exactKeys(input, ['schema_version', 'account_id', 'worker_name', 'database_id', 'database_name', 'store_id',
    'actor', 'expected_version_id', 'credentials_file', 'wrangler_config']);
  if (input.schema_version !== 1 || typeof input.account_id !== 'string' || !/^[a-f0-9]{32}$/.test(input.account_id)
    || typeof input.worker_name !== 'string' || !/^[a-z0-9][a-z0-9-]{0,62}$/.test(input.worker_name)
    || typeof input.database_id !== 'string' || !UUID.test(input.database_id)
    || typeof input.database_name !== 'string' || !/^[a-z0-9][a-z0-9_-]{0,63}$/.test(input.database_name)
    || typeof input.store_id !== 'string' || !/^[1-9]\d{0,19}$/.test(input.store_id)
    || typeof input.actor !== 'string' || !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(input.actor)
    || input.actor !== input.actor.toLowerCase() || input.actor.length > 320
    || typeof input.expected_version_id !== 'string' || !UUID.test(input.expected_version_id)
    || typeof input.credentials_file !== 'string' || !input.credentials_file.trim()
    || typeof input.wrangler_config !== 'string' || !input.wrangler_config.trim()) stop('OPERATOR_INVALID_CONFIG');
  return { schema_version: 1, account_id: input.account_id, worker_name: input.worker_name,
    database_id: input.database_id, database_name: input.database_name, store_id: input.store_id, actor: input.actor,
    expected_version_id: input.expected_version_id, credentials_file: input.credentials_file, wrangler_config: input.wrangler_config };
}

export function parseOperatorArgs(args: string[]) {
  const command = args[0] as Command;
  if (!COMMANDS.includes(command)) stop('OPERATOR_USAGE');
  const values = new Map<string, string>();
  for (let i = 1; i < args.length; i += 2) {
    const flag = args[i], value = args[i + 1];
    if (!['--config', '--request', '--item-id'].includes(flag) || !value || value.startsWith('--') || values.has(flag)) stop('OPERATOR_USAGE');
    values.set(flag, value);
  }
  if (!values.has('--config') || !values.has('--request') || (command === 'row') !== values.has('--item-id')) stop('OPERATOR_USAGE');
  const itemId = values.get('--item-id');
  if (itemId && !UUID.test(itemId)) stop('OPERATOR_USAGE');
  return { command, configPath: values.get('--config')!, requestPath: values.get('--request')!, itemId };
}

/** No symlinks, public files, unbounded allocation, or raw filesystem errors. */
export async function readPrivateFile(path: string, maximum: number): Promise<string> {
  try {
    const file = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW);
    try {
      const info = await file.stat();
      if (!info.isFile() || info.size > maximum || (process.platform !== 'win32' && (info.mode & 0o077))) stop('OPERATOR_PRIVATE_FILE_REQUIRED');
      const bytes = Buffer.alloc(maximum + 1); let used = 0;
      while (used < bytes.length) {
        const result = await file.read(bytes, used, bytes.length - used, used);
        if (!result.bytesRead) break;
        used += result.bytesRead;
      }
      if (used > maximum) stop('OPERATOR_INPUT_TOO_LARGE');
      return new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(bytes.subarray(0, used));
    } finally { await file.close(); }
  } catch (error) { if (error instanceof OperatorError) throw error; stop('OPERATOR_PRIVATE_FILE_REQUIRED'); }
}

export function parseEcwidCredentials(text: string, storeId: string): string {
  const env = parseEnv(text);
  if (Object.keys(env).some(key => !['ECWID_TOKEN', 'ECWID_STORE_ID'].includes(key))
    || env.ECWID_STORE_ID !== storeId || !/^secret_[A-Za-z0-9_-]+$/.test(env.ECWID_TOKEN ?? '')) stop('OPERATOR_CREDENTIALS_INVALID');
  return env.ECWID_TOKEN!;
}

/** Official CLI API: capture output in memory; never forward credential stdout/stderr. */
export async function loadWranglerBearer(configPath: string): Promise<string> {
  try {
    await verifyNoImplicitVars(configPath);
    const require = createRequire(import.meta.url);
    const cli = resolve(dirname(require.resolve('wrangler/package.json')), 'bin/wrangler.js');
    const result = await promisify(execFile)(process.execPath, [cli, 'auth', 'token', '--json', '--config', configPath], {
      timeout: 30_000, maxBuffer: 16_384, encoding: 'utf8',
      // `auth token` uses Wrangler's logger: disable its on-disk log explicitly.
      env: { ...process.env, CI: 'true', WRANGLER_SEND_METRICS: 'false', WRANGLER_LOG: 'log', WRANGLER_WRITE_LOGS: 'false' },
    });
    return parseWranglerBearerOutput(result.stdout);
  } catch { stop('OPERATOR_CLOUDFLARE_AUTH_REQUIRED'); }
}

/** Wrangler envFiles:[] still discovers sibling .dev.vars/.env files. */
export async function verifyNoImplicitVars(configPath: string): Promise<void> {
  try {
    const names = await readdir(dirname(configPath));
    if (names.some(name => /^(?:\.dev\.vars|\.env)(?:\.|$)/.test(name))) stop('OPERATOR_IMPLICIT_VARS_FORBIDDEN');
  } catch (error) { if (error instanceof OperatorError) throw error; stop('OPERATOR_CONFIG_DIRECTORY_UNAVAILABLE'); }
}

export function parseWranglerBearerOutput(text: string): string {
  try {
    if (text.length > 16_384) stop('OPERATOR_CLOUDFLARE_AUTH_REQUIRED');
    const auth = object(JSON.parse(text));
    if (!['oauth', 'api_token'].includes(String(auth.type)) || typeof auth.token !== 'string'
      || !/^[A-Za-z0-9_.-]{20,4096}$/.test(auth.token)) stop('OPERATOR_CLOUDFLARE_AUTH_REQUIRED');
    return auth.token;
  } catch { stop('OPERATOR_CLOUDFLARE_AUTH_REQUIRED'); }
}

/** A dedicated JSON config cannot accidentally start queues, cron or another binding. */
export function verifyProxyConfig(value: unknown, target: OperatorConfig): void {
  const config = object(value);
  exactKeys(config, ['name', 'account_id', 'compatibility_date', 'compatibility_flags', 'workers_dev', 'preview_urls', 'd1_databases']);
  if (config.name !== `${target.worker_name}-operator` || config.account_id !== target.account_id
    || typeof config.compatibility_date !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(config.compatibility_date)
    || JSON.stringify(config.compatibility_flags) !== '["nodejs_compat"]' || config.workers_dev !== false || config.preview_urls !== false
    || !Array.isArray(config.d1_databases) || config.d1_databases.length !== 1) stop('OPERATOR_PROXY_TARGET_MISMATCH');
  const database = object(config.d1_databases[0]);
  exactKeys(database, ['binding', 'database_name', 'database_id', 'remote']);
  if (database.binding !== 'DB' || database.database_name !== target.database_name || database.database_id !== target.database_id
    || database.remote !== true) stop('OPERATOR_PROXY_TARGET_MISMATCH');
}

async function cloudflare(config: OperatorConfig, credentials: OperatorCredentials, suffix: string,
  fetcher: EcwidFetch, body?: ObjectValue): Promise<unknown> {
  try {
    const response = await fetcher(`https://api.cloudflare.com/client/v4/accounts/${config.account_id}${suffix}`, {
      method: body ? 'POST' : 'GET', redirect: 'error', signal: AbortSignal.timeout(15_000),
      headers: { Authorization: `Bearer ${credentials.cloudflareToken}`, Accept: 'application/json', ...(body ? { 'Content-Type': 'application/json' } : {}) },
      ...(body ? { body: JSON.stringify(body) } : {}),
    });
    if (!response.ok) { await response.body?.cancel(); stop('OPERATOR_CLOUDFLARE_CHECK_FAILED'); }
    const result = object(await readBoundedJson(response, 1_000_000));
    if (result.success !== true || (Array.isArray(result.errors) && result.errors.length)) stop('OPERATOR_CLOUDFLARE_CHECK_FAILED');
    return result.result;
  } catch { stop('OPERATOR_CLOUDFLARE_CHECK_FAILED'); }
}

function deploymentVersion(value: unknown): string {
  const result = object(value);
  if (!Array.isArray(result.deployments) || !result.deployments.length) stop('OPERATOR_DEPLOYMENT_UNSAFE');
  // The API lists newest first; also verify timestamps so an unexpected order cannot select an older deployment.
  const deployments = result.deployments.map(object);
  const current = deployments[0];
  const created = Date.parse(String(current.created_on));
  if (!Number.isFinite(created) || deployments.some(entry => !Number.isFinite(Date.parse(String(entry.created_on)))
    || Date.parse(String(entry.created_on)) > created) || !Array.isArray(current.versions) || current.versions.length !== 1) stop('OPERATOR_DEPLOYMENT_UNSAFE');
  const version = object(current.versions[0]);
  if (version.percentage !== 100 || typeof version.version_id !== 'string' || !UUID.test(version.version_id)) stop('OPERATOR_DEPLOYMENT_UNSAFE');
  return version.version_id;
}

export async function verifyDeployment(config: OperatorConfig, credentials: OperatorCredentials, fetcher: EcwidFetch = fetch,
  recoveryReadOnly = false) {
  const base = `/workers/scripts/${config.worker_name}`;
  const versionId = deploymentVersion(await cloudflare(config, credentials, `${base}/deployments`, fetcher));
  const versionMatches = versionId === config.expected_version_id;
  if (!versionMatches && !recoveryReadOnly) stop('OPERATOR_DEPLOYMENT_CHANGED');
  const version = object(await cloudflare(config, credentials, `${base}/versions/${versionId}`, fetcher));
  const bindings = object(version.resources).bindings;
  if (version.id !== versionId || !Array.isArray(bindings)) stop('OPERATOR_DEPLOYMENT_UNSAFE');
  const named = new Map<string, ObjectValue>();
  for (const raw of bindings) {
    const binding = object(raw);
    if (typeof binding.name !== 'string' || named.has(binding.name)) stop('OPERATOR_DEPLOYMENT_UNSAFE');
    named.set(binding.name, binding);
  }
  const plain = (name: string) => { const binding = named.get(name); return binding?.type === 'plain_text' ? binding.text : undefined; };
  if (plain('ECWID_MODE') !== 'live' || plain('ECWID_STORE_ID') !== config.store_id) stop('OPERATOR_LIVE_FLAGS_UNSAFE');
  const flagsDisabled = ['INVENTORY_ENABLED', 'LIVE_SYNC_ENABLED', 'ORDER_SYNC_ENABLED'].every(name => plain(name) === 'false');
  if (!flagsDisabled && !recoveryReadOnly) stop('OPERATOR_LIVE_FLAGS_UNSAFE');
  if (named.get('DB')?.type !== 'd1' || named.get('DB')?.id !== config.database_id) stop('OPERATOR_DEPLOYED_DATABASE_MISMATCH');
  for (const list of ['STAFF_EMAILS', 'ADMIN_EMAILS']) {
    const emails = plain(list);
    if (typeof emails !== 'string' || !emails.split(',').map(email => email.trim().toLowerCase()).includes(config.actor)) stop('OPERATOR_ADMIN_NOT_ALLOWED');
  }
  if (typeof plain('ACCESS_TEAM_DOMAIN') !== 'string' || !/^https:\/\/[a-z0-9-]+\.cloudflareaccess\.com\/?$/i.test(String(plain('ACCESS_TEAM_DOMAIN')))
    || typeof plain('ACCESS_AUD') !== 'string' || !/^[a-f0-9]{64}$/.test(String(plain('ACCESS_AUD')))) stop('OPERATOR_ACCESS_NOT_CONFIGURED');
  const database = object(await cloudflare(config, credentials, `/d1/database/${config.database_id}`, fetcher));
  if (database.uuid !== config.database_id || database.name !== config.database_name) stop('OPERATOR_DEPLOYED_DATABASE_MISMATCH');
  if (deploymentVersion(await cloudflare(config, credentials, `${base}/deployments`, fetcher)) !== versionId) stop('OPERATOR_DEPLOYMENT_CHANGED');
  return { version_id: versionId, version_matches: versionMatches, live_flags_disabled: flagsDisabled,
    safe_for_cutover: versionMatches && flagsDisabled };
}

async function connectDatabase(config: OperatorConfig, credentials: OperatorCredentials): Promise<OperatorDatabase> {
  verifyProxyConfig(JSON.parse(await readPrivateFile(config.wrangler_config, 16_384)), config);
  await verifyNoImplicitVars(config.wrangler_config);
  const keys = ['CLOUDFLARE_API_TOKEN', 'CLOUDFLARE_ACCOUNT_ID', 'WRANGLER_LOG', 'WRANGLER_SEND_METRICS', 'WRANGLER_WRITE_LOGS'] as const;
  const previous = keys.map(key => process.env[key]);
  process.env.CLOUDFLARE_API_TOKEN = credentials.cloudflareToken;
  process.env.CLOUDFLARE_ACCOUNT_ID = config.account_id;
  process.env.WRANGLER_LOG = 'error'; process.env.WRANGLER_SEND_METRICS = 'false';
  process.env.WRANGLER_WRITE_LOGS = 'false';
  const restore = () => keys.forEach((key, i) => { if (previous[i] === undefined) delete process.env[key]; else process.env[key] = previous[i]; });
  try {
    const { getPlatformProxy } = await import('wrangler');
    const platform = await getPlatformProxy<{ DB: D1Database }>({ configPath: config.wrangler_config, envFiles: [], persist: false, remoteBindings: true });
    if (!platform.env.DB || typeof platform.env.DB.prepare !== 'function' || typeof platform.env.DB.batch !== 'function') {
      await platform.dispose(); stop('OPERATOR_DATABASE_UNAVAILABLE');
    }
    return { db: platform.env.DB, async dispose() { try { await platform.dispose(); } finally { restore(); } } };
  } catch { restore(); stop('OPERATOR_DATABASE_UNAVAILABLE'); }
}

/** Recheck live deployment immediately before every state mutation, including final atomic batch. */
function guardDatabase(db: D1Database, check: () => Promise<unknown>): D1Database {
  const originals = new WeakMap<D1PreparedStatement, D1PreparedStatement>();
  function prepared(statement: D1PreparedStatement, sql: string): D1PreparedStatement {
    const proxy = new Proxy(statement, { get(target, property) {
      if (property === 'bind') return (...values: unknown[]) => prepared(target.bind(...values), sql);
      const value = Reflect.get(target, property, target);
      if (typeof value !== 'function') return value;
      if (['first', 'all', 'run', 'raw'].includes(String(property))) return async (...args: unknown[]) => {
        // These are trusted service statements, never caller SQL. Recognize the
        // exact read-only CTE used by assertDatabase, not arbitrary WITH queries.
        const readOnly = /^\s*SELECT\b/i.test(sql)
          || /^WITH c AS \(SELECT \? AS orders_json, \? AS lines_json\) SELECT \(/.test(sql);
        if (!readOnly) await check();
        return Reflect.apply(value, target, args);
      };
      return value.bind(target);
    } });
    originals.set(proxy, statement);
    return proxy;
  }
  return new Proxy(db, { get(target, property) {
    if (property === 'prepare') return (sql: string) => prepared(target.prepare(sql), sql);
    if (property === 'batch') return async (statements: D1PreparedStatement[]) => {
      await check(); return target.batch(statements.map(statement => originals.get(statement) ?? statement));
    };
    if (property === 'exec') return () => stop('OPERATOR_RAW_SQL_FORBIDDEN');
    const value = Reflect.get(target, property, target);
    return typeof value === 'function' ? value.bind(target) : value;
  } });
}

function alignmentRequest(value: unknown, command: Command, itemId?: string): ObjectValue {
  const request = object(value);
  if (command === 'status') {
    exactKeys(request, ['operation_id']);
    if (typeof request.operation_id !== 'string' || !UUID.test(request.operation_id)) stop('OPERATOR_INVALID_REQUEST');
    return request;
  }
  if (command === 'preview' || command === 'stage') return request;
  if (command === 'recover') {
    exactKeys(request, ['operation_id', 'expected_hash', 'recovery_id', 'recovery_freeze']);
    if (typeof request.operation_id !== 'string' || !UUID.test(request.operation_id)
      || typeof request.expected_hash !== 'string' || !SHA.test(request.expected_hash)
      || typeof request.recovery_id !== 'string' || !UUID.test(request.recovery_id)) stop('OPERATOR_INVALID_REQUEST');
    const recoveryFreeze = object(request.recovery_freeze); exactKeys(recoveryFreeze, ['confirmed', 'started_at']);
    if (recoveryFreeze.confirmed !== true || typeof recoveryFreeze.started_at !== 'string') stop('OPERATOR_INVALID_REQUEST');
    return request;
  }
  exactKeys(request, ['operation_id', 'expected_hash', 'freeze']);
  if (typeof request.operation_id !== 'string' || !UUID.test(request.operation_id)
    || typeof request.expected_hash !== 'string' || !SHA.test(request.expected_hash)) stop('OPERATOR_INVALID_REQUEST');
  const freeze = object(request.freeze); exactKeys(freeze, ['confirmed', 'started_at']);
  if (freeze.confirmed !== true || typeof freeze.started_at !== 'string') stop('OPERATOR_INVALID_REQUEST');
  return command === 'row' ? { ...request, item_id: itemId } : request;
}

function sanitizedReceipt(value: unknown): ObjectValue {
  const receipt = object(value); const safe: ObjectValue = {};
  for (const key of ['status', 'state', 'duplicate', 'operation_id', 'store_id', 'review_hash', 'preview_hash', 'row_count',
    'order_count', 'line_count', 'workbook_line_count', 'created_at', 'verified_count', 'activated', 'reservations_loaded',
    'ecwid_changed', 'dry_run', 'item_id', 'alignment_status', 'no_op', 'recovery_id', 'recovery_frozen_at',
    'recovery_evidence_hash', 'recovery_initial_verified_count', 'recovery_resumed_count']) {
    if (receipt[key] === null || ['string', 'number', 'boolean'].includes(typeof receipt[key])) safe[key] = receipt[key];
  }
  return safe;
}

export async function runOperator(command: Command, rawConfig: unknown, credentials: OperatorCredentials,
  value: unknown, itemId?: string, dependencies: OperatorDependencies = {}): Promise<ObjectValue> {
  if (!COMMANDS.includes(command) || (command === 'row' ? !itemId || !UUID.test(itemId) : itemId !== undefined)) stop('OPERATOR_USAGE');
  const config = parseOperatorConfig(rawConfig);
  if (!/^secret_[A-Za-z0-9_-]+$/.test(credentials.ecwidToken) || !/^[A-Za-z0-9_.-]{20,4096}$/.test(credentials.cloudflareToken)) stop('OPERATOR_CREDENTIALS_INVALID');
  const request = alignmentRequest(value, command, itemId);
  const cfFetch = dependencies.cloudflareFetch ?? fetch;
  // Bindings/version metadata and D1 identity are immutable for this session.
  // Only cache those after full validation; active deployment is NEVER cached.
  const deployed = await verifyDeployment(config, credentials, cfFetch, command === 'status');
  const check = async () => {
    const current = deploymentVersion(await cloudflare(config, credentials, `/workers/scripts/${config.worker_name}/deployments`, cfFetch));
    if (current !== config.expected_version_id) stop('OPERATOR_DEPLOYMENT_CHANGED');
  };
  if (command === 'status') {
    const result = await cloudflare(config, credentials, `/d1/database/${config.database_id}/query`, cfFetch, {
      sql: `SELECT operation_id,store_id,review_hash,state,row_count,order_count,line_count,frozen_at,
        (SELECT COUNT(*) FROM opening_cutover_rows r WHERE r.operation_id=b.operation_id AND alignment_status='VERIFIED') AS verified_count,
        (SELECT COUNT(*) FROM opening_cutover_rows r WHERE r.operation_id=b.operation_id AND alignment_status IN ('PROCESSING','UNKNOWN','BLOCKED')) AS held_count
        FROM opening_cutover_batches b WHERE operation_id=? AND store_id=? AND actor=?`,
      params: [request.operation_id, config.store_id, config.actor],
    });
    if (!Array.isArray(result) || result.length !== 1) stop('OPERATOR_STATUS_INVALID');
    const statement = object(result[0]);
    if (statement.success !== true || !Array.isArray(statement.results) || statement.results.length > 1) stop('OPERATOR_STATUS_INVALID');
    const batch = statement.results.length ? object(statement.results[0]) : null;
    return { ...deployed, command, read_only: true, resubmission_authorized: false,
      found: batch !== null, ...(batch ? sanitizedReceipt(batch) : {}),
      ...(batch ? { held_count: batch.held_count, frozen_at: batch.frozen_at } : {}) };
  }
  const policy = { storeId: config.store_id, actor: config.actor, token: credentials.ecwidToken,
    mode: 'live', inventoryEnabled: 'false', liveSyncEnabled: 'false', orderSyncEnabled: 'false' };
  if (command === 'preview') return { ...deployed, command, ...sanitizedReceipt(await previewOpeningCutover(request, policy)) };
  const connection = await (dependencies.connect ?? connectDatabase)(config, credentials);
  try {
    const checkWrite = async () => {
      await check();
      // Cloudflare verification itself can take time. Never let its latency
      // extend the existing services' freeze lease immediately before a write.
      const freeze = object(command === 'recover' ? request.recovery_freeze : request.freeze);
      const started = typeof freeze.started_at === 'string' ? Date.parse(freeze.started_at) : NaN;
      const elapsed = Date.now() - started;
      const deadline = command === 'recover' ? 29 * 60 * 1000 : 15 * 60 * 1000;
      if (freeze.confirmed !== true || !Number.isFinite(elapsed) || elapsed < 0 || elapsed >= deadline) stop('CUTOVER_FREEZE_EXPIRED');
    };
    const restrictedEcwidFetch = (guardPut: boolean): EcwidFetch => async (input, init) => {
      const url = new URL(typeof input === 'string' || input instanceof URL ? input : input.url);
      if (url.origin !== 'https://app.ecwid.com' || !url.pathname.startsWith(`/api/v3/${config.store_id}/`)
        || !['GET', 'PUT'].includes(init?.method ?? '')) stop('OPERATOR_ECWID_REQUEST_FORBIDDEN');
      if (guardPut && init?.method === 'PUT') await checkWrite();
      if (init?.signal?.aborted) stop('OPERATOR_REQUEST_EXPIRED');
      return (dependencies.ecwidFetch ?? fetch)(input, init);
    };
    if (command === 'recover') {
      const result = await recoverAndActivateCutover(connection.db, request, policy, {
        fetcher: restrictedEcwidFetch(false), beforeStockWrite: checkWrite, beforeActivation: checkWrite,
        onPreflight: receipt => dependencies.onRecoveryPreflight?.(receipt),
        onRow: receipt => dependencies.onRow?.(receipt),
      });
      return { ...deployed, command, ...sanitizedReceipt(result) };
    }
    const db = guardDatabase(connection.db, checkWrite);
    const ecwidFetch = restrictedEcwidFetch(true);
    if (command === 'rows') {
      const freeze = object(request.freeze);
      const batch = await db.prepare(`SELECT state,row_count FROM opening_cutover_batches
        WHERE operation_id=? AND store_id=? AND actor=? AND review_hash=? AND frozen_at=?`)
        .bind(request.operation_id, config.store_id, config.actor, request.expected_hash, freeze.started_at)
        .first<{ state: string; row_count: number }>();
      if (!batch || batch.state !== 'ALIGNING') stop('CUTOVER_STATE_INVALID');
      const rows = (await db.prepare('SELECT item_id,alignment_status FROM opening_cutover_rows WHERE operation_id=? ORDER BY item_id LIMIT 201')
        .bind(request.operation_id).all<{ item_id: string; alignment_status: string }>()).results;
      if (!rows.length || rows.length > 200 || rows.length !== batch.row_count) stop('CUTOVER_AUDIT_INVALID');
      // Preflight the entire dispatch set before doing anything, including rows
      // later in the sequence. Never attempt PROCESSING/UNKNOWN/BLOCKED rows.
      if (rows.some(row => !UUID.test(row.item_id) || !['PENDING', 'VERIFIED'].includes(row.alignment_status))) stop('CUTOVER_ROW_HELD');
      const results: ObjectValue[] = [];
      for (const row of rows) {
        // A prior external halt must also stop a run consisting of receipt-only skips.
        const current = await db.prepare('SELECT state FROM opening_cutover_batches WHERE operation_id=?')
          .bind(request.operation_id).first<string>('state');
        if (current !== 'ALIGNING') stop('CUTOVER_STATE_INVALID');
        const result = row.alignment_status === 'VERIFIED'
          ? { item_id: row.item_id, alignment_status: 'VERIFIED', skipped: true }
          : sanitizedReceipt(await alignCutoverRow(db, { ...request, item_id: row.item_id }, policy, { fetcher: ecwidFetch }));
        if (result.alignment_status !== 'VERIFIED' || (result.state !== undefined && result.state !== 'ALIGNING')) stop('CUTOVER_REVIEW_REQUIRED');
        results.push(result); dependencies.onRow?.(result);
      }
      const finalState = await db.prepare('SELECT state FROM opening_cutover_batches WHERE operation_id=?')
        .bind(request.operation_id).first<string>('state');
      if (finalState !== 'ALIGNING') stop('CUTOVER_STATE_INVALID');
      return { ...deployed, command, operation_id: request.operation_id, state: 'ALIGNING', activated: false,
        row_count: rows.length, verified_count: results.length, rows: results };
    }
    const result = command === 'stage' ? await stageOpeningCutover(db, request, policy)
      : command === 'begin' ? await beginCutoverAlignment(db, request, policy, { fetcher: ecwidFetch })
        : command === 'row' ? await alignCutoverRow(db, request, policy, { fetcher: ecwidFetch })
          : await finishAndActivateCutover(db, request, policy, { fetcher: ecwidFetch });
    return { ...deployed, command, ...sanitizedReceipt(result) };
  } finally { await connection.dispose(); }
}

export function safeOperatorError(error: unknown): { error: string; message: string } {
  const code = error instanceof OperatorError || error instanceof DomainError ? error.code : 'OPERATOR_FAILED';
  return { error: /^[A-Z][A-Z0-9_]{0,80}$/.test(code) ? code : 'OPERATOR_FAILED',
    message: 'Stopped. No automatic retry was made. Inspect status and the durable journal before continuing; never resend an uncertain stock write.' };
}

export async function operatorMain(args: string[]): Promise<ObjectValue> {
  const parsed = parseOperatorArgs(args);
  const configPath = resolve(parsed.configPath);
  const config = parseOperatorConfig(JSON.parse(await readPrivateFile(configPath, 16_384)));
  config.credentials_file = resolve(dirname(configPath), config.credentials_file);
  config.wrangler_config = resolve(dirname(configPath), config.wrangler_config);
  verifyProxyConfig(JSON.parse(await readPrivateFile(config.wrangler_config, 16_384)), config);
  await verifyNoImplicitVars(config.wrangler_config);
  const ecwidToken = parseEcwidCredentials(await readPrivateFile(config.credentials_file, 16_384), config.store_id);
  const request = JSON.parse(await readPrivateFile(resolve(parsed.requestPath), 10_000_000));
  // Validate command/request before allowing OAuth refresh or any remote check.
  alignmentRequest(request, parsed.command, parsed.itemId);
  const cloudflareToken = await loadWranglerBearer(config.wrangler_config);
  return runOperator(parsed.command, config, { ecwidToken, cloudflareToken }, request, parsed.itemId,
    { onRecoveryPreflight: receipt => console.log(JSON.stringify(receipt)),
      onRow: receipt => console.log(JSON.stringify({ event: 'row_verified', ...receipt })) });
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  if (process.argv.length === 3 && process.argv[2] === '--help') {
    console.log('Usage: node --import tsx scripts/cutover-operator.ts <status|preview|stage|begin|row|rows|finish|recover> --config private-config.json --request private-request.json [--item-id UUID]\nNo credentials, clock overrides or automatic retries. recover is one trusted 30-minute session: it validates every target and all orders, resumes only PENDING rows, and performs final read-back/activation while preserving the original order watermark.');
  } else operatorMain(process.argv.slice(2)).then(result => console.log(JSON.stringify(result, null, 2))).catch(error => {
    console.error(JSON.stringify(safeOperatorError(error))); process.exitCode = 1;
  });
}
