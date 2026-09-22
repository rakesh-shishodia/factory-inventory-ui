import { afterEach, describe, expect, it, vi } from 'vitest';
import { EcwidClient, EcwidError, type EcwidFetch } from '../src/ecwid';

const credentials = { storeId: '123', token: 'private-test-token' };
const maximumStock = 2_147_483_647;
afterEach(() => { vi.restoreAllMocks(); vi.useRealTimers(); });

describe('one-time Ecwid absolute opening alignment', () => {
  it.each([
    { combinationId: null, suffix: '/products/1001' },
    { combinationId: '501', suffix: '/products/1001/combinations/501' },
  ])('updates only the exact target quantity ($suffix)', async ({ combinationId, suffix }) => {
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(Response.json({ updateCount: 1 }));
    const client = new EcwidClient(credentials, fetcher);
    await expect(client.setStockQuantity('1001', 37, combinationId)).resolves.toEqual({});
    expect(fetcher).toHaveBeenCalledTimes(1);
    const [input, init] = fetcher.mock.calls[0];
    expect(input).toBe(`https://app.ecwid.com/api/v3/123${suffix}`);
    expect(String(input)).not.toContain(credentials.token);
    expect(new URL(String(input)).search).toBe('');
    expect(init).toMatchObject({ method: 'PUT', redirect: 'error', body: '{"quantity":37}' });
    expect(init?.headers).toEqual({ Authorization: `Bearer ${credentials.token}`, 'Content-Type': 'application/json' });
    expect(init?.signal).toBeInstanceOf(AbortSignal);
    expect(JSON.parse(String(init?.body))).toEqual({ quantity: 37 });
  });

  it.each([0, 1, maximumStock])('accepts a finite whole opening quantity of %s', async quantity => {
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(Response.json({ updateCount: 1 }));
    await new EcwidClient(credentials, fetcher).setStockQuantity('1001', quantity);
    expect(fetcher.mock.calls[0][1]?.body).toBe(JSON.stringify({ quantity }));
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it.each([-1, 0.5, maximumStock + 1, Number.MAX_SAFE_INTEGER, NaN, Infinity, -Infinity])(
    'rejects unsafe opening quantity %s before transport', async quantity => {
      const fetcher = vi.fn<EcwidFetch>();
      await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', quantity))
        .rejects.toMatchObject({ outcome: 'REJECTED' });
      expect(fetcher).not.toHaveBeenCalled();
    },
  );

  it.each(['', '0', '-1', '01', '../1002', '1001?token=leak', '1/2', '1.2', ' 1001', '1001 ', '1'.repeat(32)])(
    'rejects malformed parent and variation IDs %s before transport', async target => {
      const fetcher = vi.fn<EcwidFetch>();
      const client = new EcwidClient(credentials, fetcher);
      await expect(client.setStockQuantity(target, 10)).rejects.toMatchObject({ outcome: 'REJECTED' });
      await expect(client.setStockQuantity('1001', 10, target)).rejects.toMatchObject({ outcome: 'REJECTED' });
      expect(fetcher).not.toHaveBeenCalled();
    },
  );

  it('does not coerce untyped input into a stock target or quantity', async () => {
    const fetcher = vi.fn<EcwidFetch>();
    const client = new EcwidClient(credentials, fetcher);
    // @ts-expect-error Runtime inputs can be untyped; numeric IDs must not be coerced.
    await expect(client.setStockQuantity(1001, 10)).rejects.toMatchObject({ outcome: 'REJECTED' });
    // @ts-expect-error The runtime boundary must reject numeric variation IDs too.
    await expect(client.setStockQuantity('1001', 10, 501)).rejects.toMatchObject({ outcome: 'REJECTED' });
    // @ts-expect-error Runtime callers must not submit string quantities.
    await expect(client.setStockQuantity('1001', '10')).rejects.toMatchObject({ outcome: 'REJECTED' });
    expect(fetcher).not.toHaveBeenCalled();
  });

  it('treats updateCount zero as an explicit rejection', async () => {
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(Response.json({ updateCount: 0 }));
    await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10, '501'))
      .rejects.toMatchObject({ outcome: 'REJECTED', message: 'Ecwid did not update this stock target.' });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it.each([null, [], {}, 1, 'success', { updateCount: '1' }, { updateCount: true }, { updateCount: 2 }, { updateCount: -1 }])(
    'never assumes a malformed confirmation succeeded: %s', async payload => {
      const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(Response.json(payload));
      await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
        .rejects.toMatchObject({ outcome: 'UNKNOWN' });
      expect(fetcher).toHaveBeenCalledTimes(1);
    },
  );

  it.each([400, 401, 402, 403, 404, 405, 409, 422])('reports definite HTTP %s rejection without retry', async status => {
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(new Response('private server details', { status }));
    await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
      .rejects.toMatchObject({ outcome: 'REJECTED', status, message: `Ecwid returned HTTP ${status}.` });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it.each([301, 302, 307, 308, 408, 500, 502, 503, 504])('holds uncertain HTTP %s for read-back, without retry', async status => {
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(new Response('private server details', { status }));
    await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10, '501'))
      .rejects.toMatchObject({ outcome: 'UNKNOWN', status, message: `Ecwid returned HTTP ${status}.` });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it.each([
    { after: '120', expected: 120 },
    { after: '999999', expected: 43_200 },
    { after: '1.5', expected: 2 },
    { after: 'invalid', expected: 60 },
    { after: '0', expected: 60 },
    { after: '-5', expected: 60 },
  ])('reports ignored 429 request as retryable, but performs no retry ($after)', async ({ after, expected }) => {
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(new Response('', { status: 429, headers: { 'Retry-After': after } }));
    await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
      .rejects.toMatchObject({ outcome: 'RETRYABLE', status: 429, retryAfter: expected });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('redacts transport error details and does not replay an uncertain write', async () => {
    const fetcher = vi.fn<EcwidFetch>().mockRejectedValue(new Error(`request failed for ${credentials.token}`));
    const result = await new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10).catch(error => error);
    expect(result).toBeInstanceOf(EcwidError);
    expect(result).toMatchObject({ outcome: 'UNKNOWN' });
    expect(String(result)).not.toContain(credentials.token);
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('treats empty or non-JSON success responses as uncertain', async () => {
    for (const response of [new Response(null, { status: 204 }), new Response('not JSON')]) {
      const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(response);
      await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
        .rejects.toMatchObject({ outcome: 'UNKNOWN' });
      expect(fetcher).toHaveBeenCalledTimes(1);
    }
  });

  it.each([true, false])('bounds both declared and streamed confirmation payloads (declared=%s)', async declared => {
    const response = new Response('x'.repeat(2_000_001), declared ? { headers: { 'Content-Length': '2000001' } } : {});
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(response);
    await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
      .rejects.toMatchObject({ outcome: 'UNKNOWN' });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('holds a truncated response body as uncertain', async () => {
    const response = new Response(new ReadableStream({ start(controller) { controller.error(new Error('stream failed')); } }));
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(response);
    await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
      .rejects.toMatchObject({ outcome: 'UNKNOWN' });
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('aborts the default timeout after 15 seconds and never retries', async () => {
    vi.useFakeTimers();
    const fetcher = vi.fn<EcwidFetch>((_input, init) => new Promise((_resolve, reject) => {
      init?.signal?.addEventListener('abort', () => reject(new Error('timeout')), { once: true });
    }));
    const pending = expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
      .rejects.toMatchObject({ outcome: 'UNKNOWN' });
    await vi.advanceTimersByTimeAsync(15_000);
    await pending;
    expect(fetcher.mock.calls[0][1]?.signal?.aborted).toBe(true);
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('keeps the timeout active while consuming a response body', async () => {
    vi.useFakeTimers();
    const fetcher = vi.fn<EcwidFetch>(async (_input, init) => new Response(new ReadableStream({
      start(controller) {
        init?.signal?.addEventListener('abort', () => controller.error(new Error('timeout')), { once: true });
      },
    })));
    const pending = expect(new EcwidClient(credentials, fetcher, 100).setStockQuantity('1001', 10))
      .rejects.toMatchObject({ outcome: 'UNKNOWN' });
    await vi.advanceTimersByTimeAsync(100);
    await pending;
    expect(fetcher).toHaveBeenCalledTimes(1);
  });

  it('clears the deadline after a successful request', async () => {
    vi.useFakeTimers();
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(Response.json({ updateCount: 1 }));
    await new EcwidClient(credentials, fetcher, 100).setStockQuantity('1001', 10);
    await vi.advanceTimersByTimeAsync(100);
    expect(fetcher.mock.calls[0][1]?.signal?.aborted).toBe(false);
  });

  it.each([0, -1, 0.5, 60_001, Infinity, NaN])('rejects an unbounded or invalid timeout %s', timeout => {
    expect(() => new EcwidClient(credentials, vi.fn<EcwidFetch>(), timeout)).toThrow(EcwidError);
  });

  it('retains optional server warning after an explicit successful confirmation', async () => {
    const fetcher = vi.fn<EcwidFetch>().mockResolvedValue(Response.json({ updateCount: 1, warning: 'Review stock notification' }));
    await expect(new EcwidClient(credentials, fetcher).setStockQuantity('1001', 10))
      .resolves.toEqual({ warning: 'Review stock notification' });
  });
});
