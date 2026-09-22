import { beforeAll, describe, expect, it, vi } from 'vitest';
import { generateKeyPair, SignJWT } from 'jose';
import { authenticate, requireAdmin } from '../src/auth';
import { json, objectBody, readJson, requireSameOrigin } from '../src/http';

// Replace only network key discovery. jose still performs actual RS256 signature,
// issuer, audience, required-claim, and expiration verification in these tests.
const keySource = vi.hoisted(() => ({ publicKey: null as CryptoKey | null, urls: [] as string[] }));
vi.mock('jose', async importOriginal => {
  const actual = await importOriginal<typeof import('jose')>();
  return { ...actual, createRemoteJWKSet: (url: URL) => {
    keySource.urls.push(url.href);
    return async () => keySource.publicKey;
  } };
});

let privateKey: CryptoKey;
let otherPrivateKey: CryptoKey;
const live = { ECWID_MODE: 'live', ACCESS_TEAM_DOMAIN: 'https://inventory-team.cloudflareaccess.com',
  ACCESS_AUD: 'inventory-application-audience', ADMIN_EMAILS: ' owner@example.com , manager@example.com' };

beforeAll(async () => {
  const keys = await generateKeyPair('RS256');
  keySource.publicKey = keys.publicKey;
  privateKey = keys.privateKey;
  otherPrivateKey = (await generateKeyPair('RS256')).privateKey;
});

async function token(options: { email?: unknown; issuer?: string; audience?: string; expiration?: string | number;
  key?: CryptoKey; omitExpiration?: boolean } = {}): Promise<string> {
  let jwt = new SignJWT({ email: options.email ?? 'picker@example.com' })
    .setProtectedHeader({ alg: 'RS256', kid: 'test-key' }).setIssuedAt().setSubject('staff-identity')
    .setIssuer(options.issuer ?? live.ACCESS_TEAM_DOMAIN).setAudience(options.audience ?? live.ACCESS_AUD);
  if (!options.omitExpiration) jwt = jwt.setExpirationTime(options.expiration ?? '5m');
  return jwt.sign(options.key ?? privateKey);
}

function request(assertion?: string, url = 'https://inventory.example/api/session'): Request {
  const headers = new Headers();
  if (assertion) headers.set('cf-access-jwt-assertion', assertion);
  // An unverified identity header must not grant administrator access.
  headers.set('cf-access-authenticated-user-email', 'owner@example.com');
  return new Request(url, { headers });
}

describe('verified staff identity', () => {
  it('uses the signed email and grants admin only from the configured allowlist', async () => {
    expect(await authenticate(request(await token()), live)).toEqual({ actor: 'picker@example.com', role: 'picker' });
    expect(await authenticate(request(await token({ email: 'OWNER@EXAMPLE.COM' })), live))
      .toEqual({ actor: 'owner@example.com', role: 'admin' });
    expect(keySource.urls.at(-1)).toBe(`${live.ACCESS_TEAM_DOMAIN}/cdn-cgi/access/certs`);
  });

  it('rejects missing authentication and invalid signatures', async () => {
    await expect(authenticate(request(), live)).rejects.toMatchObject({ status: 401, code: 'SIGN_IN_REQUIRED' });
    await expect(authenticate(request(await token({ key: otherPrivateKey })), live))
      .rejects.toMatchObject({ status: 401, code: 'INVALID_SESSION' });
  });

  it.each([
    { audience: 'another-application' },
    { issuer: 'https://different-team.cloudflareaccess.com' },
    { expiration: Math.floor(Date.now() / 1000) - 60 },
    { omitExpiration: true },
    { email: 123 },
    { email: 'not-an-email' },
  ])('rejects untrusted token claims %j', async claims => {
    await expect(authenticate(request(await token(claims)), live))
      .rejects.toMatchObject({ status: 401, code: 'INVALID_SESSION' });
  });

  it('rejects HMAC tokens even when they claim an administrator identity', async () => {
    const forged = await new SignJWT({ email: 'owner@example.com' })
      .setProtectedHeader({ alg: 'HS256' }).setIssuedAt().setExpirationTime('5m').setSubject('staff-identity')
      .setIssuer(live.ACCESS_TEAM_DOMAIN).setAudience(live.ACCESS_AUD)
      .sign(new TextEncoder().encode('an-attacker-controlled-signing-key'));
    await expect(authenticate(request(forged), live)).rejects.toMatchObject({ code: 'INVALID_SESSION' });
  });

  it('fails closed on missing or unsafe Access configuration', async () => {
    for (const override of [{ ACCESS_AUD: '' }, { ACCESS_TEAM_DOMAIN: 'https://attacker.example' },
      { ACCESS_TEAM_DOMAIN: 'http://inventory-team.cloudflareaccess.com' }]) {
      await expect(authenticate(request(), { ...live, ...override }))
        .rejects.toMatchObject({ status: 503, code: 'ACCESS_NOT_CONFIGURED' });
    }
  });

  it('grants the demo identity only on loopback development URLs', async () => {
    const demo = { ...live, ECWID_MODE: 'demo' };
    for (const url of ['http://localhost:8787/api/session', 'http://127.0.0.1:8787/api/session', 'http://[::1]:8787/api/session']) {
      expect(await authenticate(request(undefined, url), demo)).toEqual({ actor: 'demo@local', role: 'admin' });
    }
    for (const url of ['https://inventory.example/api/session', 'https://localhost.attacker.example/api/session']) {
      await expect(authenticate(request(undefined, url), demo)).rejects.toMatchObject({ code: 'ACCESS_NOT_CONFIGURED' });
    }
    await expect(authenticate(request(undefined, 'http://localhost:8787/api/session'), live))
      .rejects.toMatchObject({ code: 'SIGN_IN_REQUIRED' });
  });

  it('enforces administrator-only actions after authentication', () => {
    expect(() => requireAdmin({ actor: 'picker@example.com', role: 'picker' })).toThrow('administrator');
    expect(() => requireAdmin({ actor: 'owner@example.com', role: 'admin' })).not.toThrow();
  });
});

describe('request boundaries', () => {
  it('rejects cross-origin browser submissions, including sibling subdomains', () => {
    const make = (headers: Record<string, string>) => new Request('https://inventory.example/api/movements', { method: 'POST', headers });
    expect(() => requireSameOrigin(make({ origin: 'https://inventory.example', 'sec-fetch-site': 'same-origin' }))).not.toThrow();
    const rejected: Array<Record<string, string>> = [{ origin: 'https://attacker.example' }, { 'sec-fetch-site': 'cross-site' },
      { origin: 'https://other.inventory.example', 'sec-fetch-site': 'same-site' }, { origin: 'null' }];
    for (const headers of rejected) {
      expect(() => requireSameOrigin(make(headers))).toThrow('inventory app');
    }
  });

  it('parses JSON but rejects malformed payloads and non-object operation bodies', async () => {
    const make = (body: string, contentType = 'application/json; charset=utf-8') =>
      new Request('https://inventory.example/api/movements', { method: 'POST', body, headers: { 'content-type': contentType } });
    expect(await readJson(make('{"quantity":2}'))).toEqual({ quantity: 2 });
    await expect(readJson(make('{broken'))).rejects.toMatchObject({ status: 400, code: 'INVALID_JSON' });
    await expect(readJson(make('{}', 'text/plain'))).rejects.toMatchObject({ status: 415, code: 'JSON_REQUIRED' });
    await expect(readJson(new Request('https://inventory.example/api/movements', { method: 'POST', headers: { 'content-type': 'application/json' } })))
      .rejects.toMatchObject({ status: 400, code: 'EMPTY_BODY' });
    for (const body of [null, [], 1, 'string']) expect(() => objectBody(body)).toThrow('JSON object');
  });

  it('bounds actual UTF-8 bytes even when Content-Length is absent or dishonest', async () => {
    const make = (declared?: string) => new Request('https://inventory.example/api/movements', { method: 'POST', body: '{"name":"ééé"}',
      headers: { 'content-type': 'application/json', ...(declared ? { 'content-length': declared } : {}) } });
    await expect(readJson(make(), 15)).rejects.toMatchObject({ status: 413, code: 'BODY_TOO_LARGE' });
    await expect(readJson(make('1'), 15)).rejects.toMatchObject({ status: 413, code: 'BODY_TOO_LARGE' });
    await expect(readJson(make('10000'), 100)).rejects.toMatchObject({ status: 413, code: 'BODY_TOO_LARGE' });
  });

  it('prevents caching authenticated inventory responses', () => {
    const response = json({ quantity: 2 });
    expect(response.headers.get('cache-control')).toBe('no-store');
    expect(response.headers.get('x-content-type-options')).toBe('nosniff');
  });
});
