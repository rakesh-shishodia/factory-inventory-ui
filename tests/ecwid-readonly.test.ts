import { describe, expect, it } from 'vitest';
import { readOnlyCredentials, checkReadOnlyAccess } from '../scripts/ecwid-readonly';
import type { EcwidFetch } from '../src/ecwid';

const credentials = { storeId: '2442119', token: 'secret_test_only' };
const metadata = () => Response.json({ total: 882, count: 1, offset: 0 });

describe('separate Ecwid read-only preparation', () => {
  it('reads the configured ID while keeping leading token characters intact', () => {
    expect(readOnlyCredentials('ECWID_STORE_ID=2442119\nECWID_TOKEN=secret_test_only')).toEqual(credentials);
  });
  it.each(['', 'public_example', 'secret_abc\nInjected', 'secret_abc with spaces'])('rejects missing/public/malformed credentials without echoing them', token => {
    expect(() => readOnlyCredentials(`ECWID_STORE_ID=2442119\nECWID_TOKEN="${token}"`)).toThrow();
  });
  it('rejects an oversized settings file', () => {
    expect(() => readOnlyCredentials(' '.repeat(16_385))).toThrow('too large');
  });
  it('uses only two GETs to fixed HTTPS endpoints with metadata-only responses', async () => {
    const calls: { url: string; init: RequestInit | undefined }[] = [];
    const fetcher: EcwidFetch = async (url, init) => {
      calls.push({ url: String(url), init });
      return metadata();
    };
    const result = await checkReadOnlyAccess(credentials, fetcher);
    expect(calls).toHaveLength(2);
    for (const [i, call] of calls.entries()) {
      const url = new URL(call.url);
      expect(url.origin).toBe('https://app.ecwid.com');
      expect(url.pathname).toBe(`/api/v3/2442119/${i === 0 ? 'products' : 'orders'}`);
      expect(url.searchParams.get('responseFields')).toBe('total,count,offset');
      expect(url.searchParams.get('limit')).toBe('1');
      expect(call.init?.method).toBe('GET');
      expect(call.init?.redirect).toBe('error');
      expect(call.init?.body).toBeUndefined();
    }
    expect(result.writes_performed).toBe(0);
    expect(JSON.stringify(result)).not.toContain(credentials.token);
  });
  it.each([401, 403, 429, 500])('does not retry or expose response bodies on HTTP %s', async status => {
    let calls = 0;
    await expect(checkReadOnlyAccess(credentials, async () => {
      calls++;
      return new Response('secret_body_do_not_print', { status });
    })).rejects.toThrow(`HTTP ${status}`);
    expect(calls).toBe(1);
  });
  it.each([null, [], { total: -1, count: 0, offset: 0 }, { total: 1, count: 2, offset: 0 },
    { total: '882', count: 1, offset: 0 }, { total: 1, count: 1, offset: 10 }])('rejects invalid metadata', async body => {
    await expect(checkReadOnlyAccess(credentials, async () => Response.json(body))).rejects.toThrow('invalid response');
  });
  it('bounds the response and suppresses transport error details', async () => {
    await expect(checkReadOnlyAccess(credentials, async () => new Response('x'.repeat(16_385)))).rejects.toThrow('invalid response');
    await expect(checkReadOnlyAccess(credentials, async () => { throw new Error(credentials.token); })).rejects.not.toThrow(credentials.token);
  });
  it('rejects an injected host/path before any network access', async () => {
    let called = false;
    await expect(checkReadOnlyAccess({ ...credentials, storeId: '2442119/../../evil' }, async () => {
      called = true; return metadata();
    })).rejects.toThrow('Invalid local');
    expect(called).toBe(false);
  });
});
