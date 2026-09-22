import { DomainError } from './domain';

export async function readJson(request: Request, maxBytes = 64 * 1024): Promise<unknown> {
  if (!request.headers.get('content-type')?.toLowerCase().startsWith('application/json')) {
    throw new DomainError(415, 'JSON_REQUIRED', 'Send this request as application/json.');
  }
  if (Number(request.headers.get('content-length')) > maxBytes) {
    throw new DomainError(413, 'BODY_TOO_LARGE', 'The request is too large.');
  }
  if (!request.body) throw new DomainError(400, 'EMPTY_BODY', 'A JSON body is required.');
  const reader = request.body.getReader();
  const chunks: Uint8Array[] = [];
  let length = 0;
  try {
    while (true) {
      const part = await reader.read();
      if (part.done) break;
      length += part.value.byteLength;
      if (length > maxBytes) {
        await reader.cancel();
        throw new DomainError(413, 'BODY_TOO_LARGE', 'The request is too large.');
      }
      chunks.push(part.value);
    }
  } finally {
    reader.releaseLock();
  }
  const bytes = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
  try { return JSON.parse(new TextDecoder().decode(bytes)); }
  catch { throw new DomainError(400, 'INVALID_JSON', 'The request contains invalid JSON.'); }
}

export function objectBody(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new DomainError(400, 'INVALID_BODY', 'A JSON object is required.');
  }
  return value as Record<string, unknown>;
}

export function json(data: unknown, status = 200): Response {
  return Response.json(data, { status, headers: {
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'same-origin'
  } });
}

export function requireSameOrigin(request: Request): void {
  const origin = request.headers.get('origin');
  const fetchSite = request.headers.get('sec-fetch-site');
  if ((origin && origin !== new URL(request.url).origin) ||
      (fetchSite && !['same-origin', 'none'].includes(fetchSite))) {
    throw new DomainError(403, 'CROSS_ORIGIN', 'Open the inventory app to submit this request.');
  }
}
