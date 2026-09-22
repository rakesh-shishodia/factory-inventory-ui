import Database from 'better-sqlite3';
import { chmod, mkdir, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { DomainError } from '../src/domain';
import type { OpeningCutoverRequest } from '../src/opening-cutover';
import type { EcwidFetch } from '../src/ecwid';
import { applyMigrations, sqliteD1 } from './d1';
import { parseOperatorArgs, parseOperatorConfig, parseEcwidCredentials, parseWranglerBearerOutput, readPrivateFile, runOperator,
  safeOperatorError, verifyDeployment, verifyProxyConfig, verifyNoImplicitVars, type OperatorConfig, type OperatorDatabase } from '../scripts/cutover-operator';

const config: OperatorConfig = { schema_version: 1, account_id: 'a'.repeat(32), worker_name: 'inventory-test',
  database_id: '10000000-0000-4000-8000-000000000001', database_name: 'inventory-test-db', store_id: '12345',
  actor: 'admin@example.test', expected_version_id: '20000000-0000-4000-8000-000000000002',
  credentials_file: 'private.env', wrangler_config: 'operator.json' };
const credentials = { ecwidToken: 'secret_private_test_only', cloudflareToken: 'test_cloudflare_private_bearer' };
const at = '2026-09-22T12:01:00.000Z', frozen = '2026-09-22T12:00:00.000Z';
let sqlite: Database.Database, db: D1Database;
let deployedBindings: Record<string, unknown>[];
let cloudflareFetch: ReturnType<typeof vi.fn<EcwidFetch>>, ecwidFetch: ReturnType<typeof vi.fn<EcwidFetch>>;
let connect: ReturnType<typeof vi.fn<() => Promise<OperatorDatabase>>>, dispose: ReturnType<typeof vi.fn<() => Promise<void>>>;
let quantity: number;

function proxyConfig() { return { name: `${config.worker_name}-operator`, account_id: config.account_id,
  compatibility_date: '2026-09-22', compatibility_flags: ['nodejs_compat'], workers_dev: false, preview_urls: false,
  d1_databases: [{ binding: 'DB', database_name: config.database_name, database_id: config.database_id, remote: true }] }; }
function fixture(): OpeningCutoverRequest {
  const target = { id: '123', sku: 'BOLT-1', name: 'Bolt', quantity: 12, unlimited: false, hasOptions: false,
    hasVariations: false, combinationId: null, variationOptions: [], enabled: true, eligibilityVerified: true,
    hasExtraOptions: false, hasBundleRelationships: false };
  return { operation_id: '30000000-0000-4000-8000-000000000003', expected_hash: '', confirm_staging: true, physical_counts_confirmed: true,
    freeze: { confirmed: true, started_at: frozen },
    input: { store_id: config.store_id, source_ref: 'Reviewed workbook', snapshot_source_hash: 'a'.repeat(64),
      balance_meaning: 'PHYSICAL_ON_HAND', reservations_confirmed: true,
      rows: [{ sku: 'BOLT-1', balance: 10, name: 'Bolt', single_unit_confirmed: true }], reservations: [],
      catalog: { kind: 'READONLY_CATALOGUE', schema_version: 1, dry_run: true, complete: true, store_id: config.store_id,
        started_at: '2026-09-22T12:00:01.000Z', completed_at: '2026-09-22T12:00:30.000Z',
        product_count: 1, stock_target_count: 1, products: [target], stock_targets: [target], reservations_confirmed: false } },
    scope: [{ sku: 'BOLT-1', ecwid_product_id: '123', ecwid_combination_id: null, ecwid_option_signature: '[]' }],
    orders: { kind: 'READONLY_ORDERS', schema_version: 1, dry_run: true, complete: true, store_id: config.store_id,
      started_at: '2026-09-22T12:00:01.000Z', completed_at: '2026-09-22T12:00:30.000Z',
      creation_cutoff: Date.parse('2026-09-22T12:00:01.000Z') / 1000, orders_checked: 0, pending_order_count: 0, line_count: 0, orders: [] },
    line_confirmations: [], workbook_scope: [] };
}
function envelope(result: unknown) { return Response.json({ success: true, errors: [], result }); }
function deployments(version = config.expected_version_id) {
  return { deployments: [{ id: 'deployment-1', created_on: frozen, versions: [{ version_id: version, percentage: 100 }] }] };
}
async function cfTransport(input: string | URL | Request, init?: RequestInit): Promise<Response> {
  const url = new URL(typeof input === 'string' || input instanceof URL ? input : input.url);
  expect(url.origin).toBe('https://api.cloudflare.com');
  expect(url.pathname.startsWith(`/client/v4/accounts/${config.account_id}/`)).toBe(true);
  expect(init?.redirect).toBe('error'); expect(init?.signal).toBeDefined();
  expect(new Headers(init?.headers).get('Authorization')).toBe(`Bearer ${credentials.cloudflareToken}`);
  if (url.pathname.endsWith('/deployments')) return envelope(deployments());
  if (url.pathname.endsWith(`/versions/${config.expected_version_id}`)) return envelope({ id: config.expected_version_id, resources: { bindings: deployedBindings } });
  if (url.pathname.endsWith(`/database/${config.database_id}`)) return envelope({ uuid: config.database_id, name: config.database_name });
  if (url.pathname.endsWith('/query')) {
    expect(init?.method).toBe('POST');
    const body = JSON.parse(String(init?.body)); expect(body.sql.trim().startsWith('SELECT')).toBe(true);
    return envelope([{ success: true, results: sqlite.prepare(body.sql).all(...body.params) }]);
  }
  throw new Error('Unexpected Cloudflare request');
}
async function ecwidTransport(input: string | URL | Request, init?: RequestInit): Promise<Response> {
  const url = new URL(typeof input === 'string' || input instanceof URL ? input : input.url);
  expect(url.origin).toBe('https://app.ecwid.com');
  if (url.pathname.endsWith('/orders')) return Response.json({ total: 0, count: 0, offset: 0, items: [] });
  if (init?.method === 'PUT') {
    expect(sqlite.prepare('SELECT alignment_status FROM opening_cutover_rows').pluck().get()).toBe('PROCESSING');
    const body = JSON.parse(String(init.body)); expect(Object.keys(body)).toEqual(['quantity']); quantity = body.quantity;
    return Response.json({ updateCount: 1 });
  }
  return Response.json({ id: 123, sku: 'BOLT-1', name: 'Bolt', quantity, unlimited: false, enabled: true, options: [], combinations: [] });
}
const deps = () => ({ cloudflareFetch, ecwidFetch, connect });
async function staged() {
  const request = fixture();
  request.expected_hash = String((await runOperator('preview', config, credentials, request, undefined, deps())).review_hash);
  await runOperator('stage', config, credentials, request, undefined, deps());
  const alignment = { operation_id: request.operation_id, expected_hash: request.expected_hash, freeze: request.freeze };
  const itemId = sqlite.prepare('SELECT item_id FROM opening_cutover_rows').pluck().get() as string;
  return { request, alignment, itemId };
}
async function stagedMany() {
  const request = fixture();
  const catalog = request.input.catalog as { products: Record<string, unknown>[]; stock_targets: Record<string, unknown>[]; product_count: number; stock_target_count: number };
  const template = catalog.products[0];
  catalog.products = [1, 2, 3].map(i => ({ ...template, id: String(122 + i), sku: `BOLT-${i}` }));
  catalog.stock_targets = catalog.products; catalog.product_count = 3; catalog.stock_target_count = 3;
  request.input.rows = [1, 2, 3].map(i => ({ sku: `BOLT-${i}`, name: 'Bolt', balance: 10, single_unit_confirmed: true }));
  request.scope = [1, 2, 3].map(i => ({ sku: `BOLT-${i}`, ecwid_product_id: String(122 + i), ecwid_combination_id: null, ecwid_option_signature: '[]' }));
  request.expected_hash = String((await runOperator('preview', config, credentials, request, undefined, deps())).review_hash);
  await runOperator('stage', config, credentials, request, undefined, deps());
  const quantities = new Map([['123', 12], ['124', 12], ['125', 12]]);
  ecwidFetch.mockImplementation(async (input, init) => {
    const url = new URL(String(input));
    if (url.pathname.endsWith('/orders')) return Response.json({ total: 0, count: 0, offset: 0, items: [] });
    const id = url.pathname.split('/').at(-1)!;
    if (init?.method === 'PUT') {
      expect(sqlite.prepare("SELECT COUNT(*) FROM opening_cutover_rows WHERE alignment_status='PROCESSING'").pluck().get()).toBe(1);
      quantities.set(id, JSON.parse(String(init.body)).quantity); return Response.json({ updateCount: 1 });
    }
    return Response.json({ id: Number(id), sku: `BOLT-${Number(id) - 122}`, name: 'Bolt', quantity: quantities.get(id),
      unlimited: false, enabled: true, options: [], combinations: [] });
  });
  const alignment = { operation_id: request.operation_id, expected_hash: request.expected_hash, freeze: request.freeze };
  await runOperator('begin', config, credentials, alignment, undefined, deps());
  return alignment;
}
beforeEach(() => {
  vi.useFakeTimers({ toFake: ['Date'] }); vi.setSystemTime(at);
  vi.stubGlobal('fetch', vi.fn(() => { throw new Error('Real network forbidden'); }));
  sqlite = new Database(':memory:'); applyMigrations(sqlite); db = sqliteD1(sqlite); quantity = 12;
  deployedBindings = Object.entries({ ECWID_MODE: 'live', ECWID_STORE_ID: config.store_id, INVENTORY_ENABLED: 'false',
    LIVE_SYNC_ENABLED: 'false', ORDER_SYNC_ENABLED: 'false', STAFF_EMAILS: config.actor, ADMIN_EMAILS: config.actor,
    ACCESS_TEAM_DOMAIN: 'https://test.cloudflareaccess.com', ACCESS_AUD: 'b'.repeat(64) }).map(([name, text]) => ({ name, type: 'plain_text', text }));
  deployedBindings.push({ name: 'DB', type: 'd1', id: config.database_id });
  cloudflareFetch = vi.fn<EcwidFetch>(cfTransport); ecwidFetch = vi.fn<EcwidFetch>(ecwidTransport);
  dispose = vi.fn(async () => {}); connect = vi.fn(async () => ({ db, dispose }));
});
afterEach(() => { sqlite.close(); vi.useRealTimers(); vi.unstubAllGlobals(); });

describe('private command and target inputs', () => {
  it('accepts only explicit commands and known flags', () => {
    expect(parseOperatorArgs(['begin', '--config', 'c.json', '--request', 'r.json']).command).toBe('begin');
    expect(parseOperatorConfig(config)).toEqual(config); expect(() => verifyProxyConfig(proxyConfig(), config)).not.toThrow();
  });
  it.each([
    ['run-all'], ['begin', '--token', 'secret'], ['begin', '--now', at], ['begin', '--config', 'c', '--request', 'r', '--config', 'other'],
    ['row', '--config', 'c', '--request', 'r'], ['begin', '--config', 'c', '--request', 'r', '--item-id', config.database_id],
  ])('rejects unsupported argv %j', (...args) => expect(() => parseOperatorArgs(args)).toThrow());
  it.each(['token', 'now', 'fetcher', 'database_url'])('rejects injected config field %s', key => {
    expect(() => parseOperatorConfig({ ...config, [key]: 'override' })).toThrow();
  });
  it.each(['remote', 'database_id', 'database_name', 'binding'])('refuses ambiguous or local D1 %s', key => {
    const local = proxyConfig(); Object.assign(local.d1_databases[0], { [key]: key === 'remote' ? false : 'wrong' });
    expect(() => verifyProxyConfig(local, config)).toThrow();
  });
  it.each(['queues', 'triggers', 'assets', 'env', 'vars', 'main'])('refuses unrelated proxy config %s', key => {
    expect(() => verifyProxyConfig({ ...proxyConfig(), [key]: {} }, config)).toThrow();
  });
  it('pins Ecwid credentials to the intended store', () => {
    expect(parseEcwidCredentials(`ECWID_STORE_ID=${config.store_id}\nECWID_TOKEN=${credentials.ecwidToken}`, config.store_id)).toBe(credentials.ecwidToken);
    expect(() => parseEcwidCredentials('ECWID_STORE_ID=999\nECWID_TOKEN=secret_test', config.store_id)).toThrow();
    expect(() => parseEcwidCredentials(`ECWID_STORE_ID=${config.store_id}\nECWID_TOKEN=public_test`, config.store_id)).toThrow();
  });
  it('accepts captured OAuth/API bearer JSON but rejects global keys and malformed auth output', () => {
    for (const type of ['oauth', 'api_token']) expect(parseWranglerBearerOutput(JSON.stringify({ type, token: credentials.cloudflareToken }))).toBe(credentials.cloudflareToken);
    for (const output of ['not json', JSON.stringify({ type: 'api_key', key: 'private', email: 'admin@example.test' }), JSON.stringify({ type: 'oauth', token: 'short' })]) {
      expect(() => parseWranglerBearerOutput(output)).toThrow('OPERATOR_CLOUDFLARE_AUTH_REQUIRED');
    }
  });
  it('reads bounded private files and rejects symlinks and public permissions', async () => {
    const folder = await mkdtemp(join(tmpdir(), 'cutover-operator-test-'));
    try {
      const path = join(folder, 'private.json'); await writeFile(path, '{}', { mode: 0o600 });
      expect(await readPrivateFile(path, 10)).toBe('{}'); await expect(readPrivateFile(path, 1)).rejects.toThrow();
      await symlink(path, join(folder, 'link.json')); await expect(readPrivateFile(join(folder, 'link.json'), 10)).rejects.toThrow();
      await chmod(path, 0o644); await expect(readPrivateFile(path, 10)).rejects.toThrow();
      await mkdir(join(folder, 'directory')); await expect(readPrivateFile(join(folder, 'directory'), 10)).rejects.toThrow();
    } finally { await rm(folder, { recursive: true, force: true }); }
  });
  it('refuses implicit Wrangler variables beside the dedicated config', async () => {
    const folder = await mkdtemp(join(tmpdir(), 'cutover-operator-env-test-'));
    try {
      await expect(verifyNoImplicitVars(join(folder, 'operator.json'))).resolves.toBeUndefined();
      for (const name of ['.dev.vars', '.dev.vars.production', '.env', '.env.local']) {
        const file = join(folder, name); await writeFile(file, 'PRIVATE=test', { mode: 0o600 });
        await expect(verifyNoImplicitVars(join(folder, 'operator.json'))).rejects.toThrow('OPERATOR_IMPLICIT_VARS_FORBIDDEN');
        await rm(file);
      }
    } finally { await rm(folder, { recursive: true, force: true }); }
  });
});

describe('live deployed authority checks', () => {
  it('uses the active version and rechecks its deployment after verifying the exact D1', async () => {
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).resolves.toEqual({ version_id: config.expected_version_id,
      version_matches: true, live_flags_disabled: true, safe_for_cutover: true });
    expect(cloudflareFetch).toHaveBeenCalledTimes(4);
    expect(cloudflareFetch.mock.calls.every(([, init]) => init?.method === 'GET')).toBe(true);
  });
  it.each(['INVENTORY_ENABLED', 'LIVE_SYNC_ENABLED', 'ORDER_SYNC_ENABLED', 'ECWID_MODE', 'ECWID_STORE_ID'])('blocks deployed drift in %s before connecting', async name => {
    deployedBindings.find(binding => binding.name === name)!.text = 'wrong';
    await expect(runOperator('begin', config, credentials, { operation_id: fixture().operation_id, expected_hash: 'a'.repeat(64), freeze: fixture().freeze }, undefined, deps())).rejects.toThrow();
    expect(connect).not.toHaveBeenCalled(); expect(ecwidFetch).not.toHaveBeenCalled();
  });
  it.each(['STAFF_EMAILS', 'ADMIN_EMAILS', 'ACCESS_AUD', 'ACCESS_TEAM_DOMAIN'])('rejects missing administrator/access %s', async name => {
    deployedBindings = deployedBindings.filter(binding => binding.name !== name);
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow();
  });
  it('rejects changed D1 binding and does not guess the target', async () => {
    deployedBindings.find(binding => binding.name === 'DB')!.id = 'wrong';
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow('OPERATOR_DEPLOYED_DATABASE_MISMATCH');
  });
  it('rejects a different active version and split traffic', async () => {
    cloudflareFetch.mockResolvedValueOnce(envelope(deployments('40000000-0000-4000-8000-000000000004')));
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow('OPERATOR_DEPLOYMENT_CHANGED');
    const split = deployments(); split.deployments[0].versions[0].percentage = 50;
    cloudflareFetch.mockResolvedValueOnce(envelope(split));
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow('OPERATOR_DEPLOYMENT_UNSAFE');
  });
  it('blocks a deployment race during verification', async () => {
    cloudflareFetch.mockImplementation(async (input, init) => {
      if (cloudflareFetch.mock.calls.length === 4) return envelope(deployments('40000000-0000-4000-8000-000000000004'));
      return cfTransport(input, init);
    });
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow('OPERATOR_DEPLOYMENT_CHANGED');
  });
  it.each([401, 403, 429, 500])('fails once without retry or response leakage for HTTP %s', async status => {
    cloudflareFetch.mockResolvedValueOnce(new Response(credentials.cloudflareToken, { status }));
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow('OPERATOR_CLOUDFLARE_CHECK_FAILED');
    expect(cloudflareFetch).toHaveBeenCalledTimes(1);
  });
  it('sanitizes transport errors, malformed data and all CLI error output', async () => {
    cloudflareFetch.mockRejectedValueOnce(new Error(credentials.cloudflareToken));
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow('OPERATOR_CLOUDFLARE_CHECK_FAILED');
    for (const error of [new Error(credentials.ecwidToken), new DomainError(400, 'SAFE_CODE', credentials.cloudflareToken), credentials.cloudflareToken]) {
      expect(JSON.stringify(safeOperatorError(error))).not.toContain('private');
    }
    cloudflareFetch.mockResolvedValueOnce(new Response('not json'));
    await expect(verifyDeployment(config, credentials, cloudflareFetch)).rejects.toThrow('OPERATOR_CLOUDFLARE_CHECK_FAILED');
  });
});

describe('operator lifecycle uses existing services without retries', () => {
  it('previews without a database proxy or Ecwid call and does not print source rows', async () => {
    const result = await runOperator('preview', config, credentials, fixture(), undefined, deps());
    expect(result).toMatchObject({ dry_run: true, row_count: 1, activated: false });
    expect(result).not.toHaveProperty('rows'); expect(result).not.toHaveProperty('actor');
    expect(connect).not.toHaveBeenCalled(); expect(ecwidFetch).not.toHaveBeenCalled();
  });
  it('stages, aligns one exact row and atomically activates; repeated receipts make no further PUT', async () => {
    const { request, alignment, itemId } = await staged();
    expect(ecwidFetch).not.toHaveBeenCalled();
    await expect(runOperator('begin', config, credentials, alignment, undefined, deps())).resolves.toMatchObject({ state: 'ALIGNING' });
    await expect(runOperator('row', config, credentials, alignment, itemId, deps())).resolves.toMatchObject({ alignment_status: 'VERIFIED' });
    await expect(runOperator('row', config, credentials, alignment, itemId, deps())).resolves.toMatchObject({ duplicate: true });
    await expect(runOperator('finish', config, credentials, alignment, undefined, deps())).resolves.toMatchObject({ state: 'ACTIVE', activated: true });
    await expect(runOperator('finish', config, credentials, alignment, undefined, deps())).resolves.toMatchObject({ duplicate: true });
    expect(ecwidFetch.mock.calls.filter(([, init]) => init?.method === 'PUT')).toHaveLength(1);
    expect(sqlite.prepare('SELECT active,last_ecwid_quantity FROM items').get()).toEqual({ active: 1, last_ecwid_quantity: 10 });
    expect(sqlite.prepare("SELECT value FROM sync_state WHERE key='orders_tracking_started'").pluck().get()).toBe(frozen);
    const connects = connect.mock.calls.length;
    const status = await runOperator('status', config, credentials, { operation_id: request.operation_id }, undefined, deps());
    expect(status).toMatchObject({ state: 'ACTIVE', verified_count: 1, held_count: 0, found: true });
    expect(connect).toHaveBeenCalledTimes(connects); expect(dispose).toHaveBeenCalledTimes(connects);
  });
  it('stops on an uncertain PUT, retains REVIEW, and cannot resend it after restart', async () => {
    const { alignment, itemId } = await staged(); await runOperator('begin', config, credentials, alignment, undefined, deps());
    ecwidFetch.mockImplementation(async (input, init) => { if (init?.method === 'PUT') throw new Error('lost confirmation'); return ecwidTransport(input, init); });
    await expect(runOperator('row', config, credentials, alignment, itemId, deps())).rejects.toThrow();
    expect(sqlite.prepare('SELECT alignment_status FROM opening_cutover_rows').pluck().get()).toBe('UNKNOWN');
    expect(sqlite.prepare('SELECT state FROM opening_cutover_batches').pluck().get()).toBe('REVIEW');
    await expect(runOperator('row', config, credentials, alignment, itemId, deps())).rejects.toThrow();
    expect(ecwidFetch.mock.calls.filter(([, init]) => init?.method === 'PUT')).toHaveLength(1);
    expect(dispose).toHaveBeenCalledTimes(connect.mock.calls.length);
  });
  it('checks deployment again after preflight and stops the PUT when live flags change', async () => {
    const { alignment, itemId } = await staged(); await runOperator('begin', config, credentials, alignment, undefined, deps());
    ecwidFetch.mockImplementation(async (input, init) => {
      const response = await ecwidTransport(input, init);
      cloudflareFetch.mockResolvedValue(envelope(deployments('40000000-0000-4000-8000-000000000004'))); return response;
    });
    await expect(runOperator('row', config, credentials, alignment, itemId, deps())).rejects.toThrow();
    expect(ecwidFetch.mock.calls.filter(([, init]) => init?.method === 'PUT')).toHaveLength(0);
    expect(sqlite.prepare('SELECT alignment_status FROM opening_cutover_rows').pluck().get()).toBe('PENDING');
  });
  it('does not mutate the database or call Ecwid when fresh work has expired', async () => {
    const { alignment } = await staged(); vi.setSystemTime('2026-09-22T12:16:00.000Z');
    await expect(runOperator('begin', config, credentials, alignment, undefined, deps())).rejects.toMatchObject({ code: 'CUTOVER_FREEZE_EXPIRED' });
    expect(ecwidFetch).not.toHaveBeenCalled(); expect(sqlite.prepare('SELECT state FROM opening_cutover_batches').pluck().get()).toBe('STAGED');
  });
  it('does not activate if deployment verification itself runs past the freeze deadline', async () => {
    const { alignment, itemId } = await staged(); await runOperator('begin', config, credentials, alignment, undefined, deps());
    await runOperator('row', config, credentials, alignment, itemId, deps());
    let reads = 0;
    cloudflareFetch.mockImplementation(async (input, init) => {
      const response = await cfTransport(input, init); reads++;
      // Four initial checks, then one fresh active-deployment read immediately
      // before the atomic activation; immutable version/D1 metadata is cached.
      if (reads === 5) vi.setSystemTime('2026-09-22T12:16:00.000Z');
      return response;
    });
    await expect(runOperator('finish', config, credentials, alignment, undefined, deps())).rejects.toThrow();
    expect(sqlite.prepare('SELECT active FROM items').pluck().get()).toBe(0);
    expect(sqlite.prepare('SELECT state FROM opening_cutover_batches').pluck().get()).toBe('ALIGNING');
  });
  it('processes rows sequentially through one connection with only fresh deployment checks per mutation', async () => {
    const alignment = await stagedMany(); const opens = connect.mock.calls.length;
    cloudflareFetch.mockClear(); ecwidFetch.mockClear(); const progress = vi.fn();
    const result = await runOperator('rows', config, credentials, alignment, undefined, { ...deps(), onRow: progress });
    expect(result).toMatchObject({ state: 'ALIGNING', activated: false, verified_count: 3, row_count: 3 });
    expect(connect).toHaveBeenCalledTimes(opens + 1); expect(progress).toHaveBeenCalledTimes(3);
    expect(ecwidFetch.mock.calls.filter(([, init]) => init?.method === 'PUT')).toHaveLength(3);
    // Four initial metadata checks, plus claim / PUT / VERIFIED for each row.
    expect(cloudflareFetch).toHaveBeenCalledTimes(13);
    expect(cloudflareFetch.mock.calls.filter(([input]) => String(input).includes('/versions/'))).toHaveLength(1);
    expect(cloudflareFetch.mock.calls.filter(([input]) => String(input).endsWith(`/database/${config.database_id}`))).toHaveLength(1);
    cloudflareFetch.mockClear(); ecwidFetch.mockClear();
    const retry = await runOperator('rows', config, credentials, alignment, undefined, deps());
    expect(retry).toMatchObject({ verified_count: 3 });
    expect((retry.rows as { skipped?: boolean }[]).every(row => row.skipped)).toBe(true);
    expect(ecwidFetch).not.toHaveBeenCalled(); expect(cloudflareFetch).toHaveBeenCalledTimes(4);
  });
  it('stops the session at the first uncertain write and never dispatches later rows or resends it', async () => {
    const alignment = await stagedMany(); const previous = ecwidFetch.getMockImplementation()!; ecwidFetch.mockClear();
    ecwidFetch.mockImplementation(async (input, init) => { if (init?.method === 'PUT') throw new Error('lost response'); return previous(input, init); });
    await expect(runOperator('rows', config, credentials, alignment, undefined, deps())).rejects.toThrow();
    expect(sqlite.prepare("SELECT COUNT(*) FROM opening_cutover_rows WHERE alignment_status='UNKNOWN'").pluck().get()).toBe(1);
    expect(sqlite.prepare("SELECT COUNT(*) FROM opening_cutover_rows WHERE alignment_status='PENDING'").pluck().get()).toBe(2);
    await expect(runOperator('rows', config, credentials, alignment, undefined, deps())).rejects.toThrow('CUTOVER_STATE_INVALID');
    expect(ecwidFetch.mock.calls.filter(([, init]) => init?.method === 'PUT')).toHaveLength(1);
  });
  it.each(['PROCESSING', 'UNKNOWN', 'BLOCKED'])('never dispatches any row while a durable %s row exists', async status => {
    const alignment = await stagedMany();
    const id = sqlite.prepare('SELECT item_id FROM opening_cutover_rows ORDER BY item_id DESC LIMIT 1').pluck().get();
    sqlite.prepare("UPDATE opening_cutover_rows SET alignment_status='PROCESSING',before_quantity=12,attempted_at=? WHERE item_id=?").run(at, id);
    if (status !== 'PROCESSING') sqlite.prepare('UPDATE opening_cutover_rows SET alignment_status=? WHERE item_id=?').run(status, id);
    ecwidFetch.mockClear();
    await expect(runOperator('rows', config, credentials, alignment, undefined, deps())).rejects.toThrow();
    expect(ecwidFetch).not.toHaveBeenCalled();
  });
  it.each([1, 3])('stops a multi-row session if the batch enters REVIEW after receipt %s', async changedAfter => {
    const alignment = await stagedMany(); ecwidFetch.mockClear();
    const progress = vi.fn(() => { if (progress.mock.calls.length === changedAfter) sqlite.prepare("UPDATE opening_cutover_batches SET state='REVIEW'").run(); });
    await expect(runOperator('rows', config, credentials, alignment, undefined, { ...deps(), onRow: progress })).rejects.toThrow('CUTOVER_STATE_INVALID');
    expect(progress).toHaveBeenCalledTimes(changedAfter); expect(ecwidFetch.mock.calls.filter(([, init]) => init?.method === 'PUT')).toHaveLength(changedAfter);
  });
  it('keeps status read-only and available after flags/version drift, but never when identity changes', async () => {
    const { request } = await staged(); const opens = connect.mock.calls.length;
    const currentVersion = '40000000-0000-4000-8000-000000000004';
    deployedBindings.find(binding => binding.name === 'INVENTORY_ENABLED')!.text = 'true';
    cloudflareFetch.mockImplementation(async (input, init) => {
      const url = String(input);
      if (url.endsWith('/deployments')) return envelope(deployments(currentVersion));
      if (url.endsWith(`/versions/${currentVersion}`)) return envelope({ id: currentVersion, resources: { bindings: deployedBindings } });
      return cfTransport(input, init);
    });
    const result = await runOperator('status', config, credentials, { operation_id: request.operation_id }, undefined, deps());
    expect(result).toMatchObject({ found: true, state: 'STAGED', version_matches: false, live_flags_disabled: false,
      safe_for_cutover: false, read_only: true, resubmission_authorized: false });
    expect(connect).toHaveBeenCalledTimes(opens); expect(ecwidFetch).not.toHaveBeenCalled();
    deployedBindings.find(binding => binding.name === 'DB')!.id = 'wrong';
    await expect(runOperator('status', config, credentials, { operation_id: request.operation_id }, undefined, deps())).rejects.toThrow('OPERATOR_DEPLOYED_DATABASE_MISMATCH');
  });
  it('rolls back the entire guarded stage if the last transaction statement fails', async () => {
    const request = fixture(); request.expected_hash = String((await runOperator('preview', config, credentials, request, undefined, deps())).review_hash);
    sqlite.exec("CREATE TRIGGER reject_stage BEFORE INSERT ON sync_issues BEGIN SELECT RAISE(ABORT,'fail'); END");
    await expect(runOperator('stage', config, credentials, request, undefined, deps())).rejects.toThrow();
    expect(sqlite.prepare('SELECT COUNT(*) FROM items').pluck().get()).toBe(0);
    expect(sqlite.prepare('SELECT COUNT(*) FROM opening_cutover_batches').pluck().get()).toBe(0); expect(dispose).toHaveBeenCalledOnce();
  });
  it('runs guarded stage and activation atomically through a real local workerd D1 binding', async () => {
    const { Miniflare, convertV4MiniflareOptions } = await import('miniflare');
    const runtime = new Miniflare(convertV4MiniflareOptions({ modules: true, script: 'export default { fetch() { return new Response("test"); } };',
      compatibilityDate: '2026-09-22', compatibilityFlags: ['nodejs_compat'], d1Databases: ['DB'],
      cf: false }));
    try {
      const runtimeDb = await runtime.getD1Database('DB');
      // Recreate the exact fully migrated schema, preserving trigger bodies as one statement.
      const schema = sqlite.prepare('SELECT sql FROM sqlite_master WHERE sql IS NOT NULL ORDER BY rowid').all() as { sql: string }[];
      await runtimeDb.batch(schema.map(({ sql }) => runtimeDb.prepare(sql)));
      connect.mockImplementation(async () => ({ db: runtimeDb, dispose }));
      ecwidFetch.mockImplementation(async (input, init) => {
        const url = new URL(String(input));
        if (url.pathname.endsWith('/orders')) return Response.json({ total: 0, count: 0, offset: 0, items: [] });
        if (init?.method === 'PUT') {
          expect(await runtimeDb.prepare('SELECT alignment_status FROM opening_cutover_rows').first('alignment_status')).toBe('PROCESSING');
          quantity = JSON.parse(String(init.body)).quantity; return Response.json({ updateCount: 1 });
        }
        return Response.json({ id: 123, sku: 'BOLT-1', name: 'Bolt', quantity, unlimited: false, enabled: true, options: [], combinations: [] });
      });
      const request = fixture(); request.expected_hash = String((await runOperator('preview', config, credentials, request, undefined, deps())).review_hash);
      await runOperator('stage', config, credentials, request, undefined, deps());
      const alignment = { operation_id: request.operation_id, expected_hash: request.expected_hash, freeze: request.freeze };
      const itemId = (await runtimeDb.prepare('SELECT item_id FROM opening_cutover_rows').first<string>('item_id'))!;
      await runOperator('begin', config, credentials, alignment, undefined, deps());
      await runOperator('row', config, credentials, alignment, itemId, deps());
      await runtimeDb.prepare("CREATE TRIGGER reject_active BEFORE UPDATE OF state ON opening_cutover_batches WHEN NEW.state='ACTIVE' BEGIN SELECT RAISE(ABORT,'TEST_FAILURE'); END").run();
      await expect(runOperator('finish', config, credentials, alignment, undefined, deps())).rejects.toThrow();
      expect(await runtimeDb.prepare('SELECT active FROM items').first('active')).toBe(0);
      expect(await runtimeDb.prepare("SELECT status FROM sync_issues WHERE kind='OPENING_CUTOVER_STAGED'").first('status')).toBe('OPEN');
      expect(await runtimeDb.prepare("SELECT value FROM sync_state WHERE key='orders_tracking_started'").first()).toBeNull();
      expect(await runtimeDb.prepare('SELECT state FROM opening_cutover_batches').first('state')).toBe('REVIEW');
    } finally { await runtime.dispose(); }
  }, 30_000);
});
