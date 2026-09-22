import { readFile, stat } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseEnv } from 'node:util';
import { readBoundedJson, type EcwidFetch } from '../src/ecwid';

/** Separate from Worker configuration: never starts polling or imports any data. */
export function readOnlyCredentials(text: string): { storeId: string; token: string } {
  if (text.length > 16_384) throw new Error('The local Ecwid settings file is too large.');
  const env = parseEnv(text);
  const storeId = env.ECWID_STORE_ID?.trim() ?? '';
  const token = env.ECWID_TOKEN?.trim() ?? '';
  if (!/^[1-9]\d{0,19}$/.test(storeId)) throw new Error('Set a numeric ECWID_STORE_ID in .env.ecwid-readonly.');
  if (!token) throw new Error('API check not run: add the secret access token to ECWID_TOKEN in .env.ecwid-readonly locally, not in chat.');
  if (!/^secret_[A-Za-z0-9_-]+$/.test(token)) throw new Error('Use an Ecwid secret access token, not a public token.');
  return { storeId, token };
}

export async function checkReadOnlyAccess(credentials: { storeId: string; token: string }, fetcher: EcwidFetch = fetch) {
  // Validate even callers that bypass the settings-file reader.
  if (!/^[1-9]\d{0,19}$/.test(credentials.storeId) || !/^secret_[A-Za-z0-9_-]+$/.test(credentials.token)) {
    throw new Error('Invalid local Ecwid credentials.');
  }
  const checks: { resource: 'products' | 'orders'; total: number }[] = [];
  for (const resource of ['products', 'orders'] as const) {
    const url = new URL(`https://app.ecwid.com/api/v3/${credentials.storeId}/${resource}`);
    url.searchParams.set('limit', '1');
    // Metadata only: no customer names, addresses, email or payment details.
    url.searchParams.set('responseFields', 'total,count,offset');
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 15_000);
    try {
      const response = await fetcher(url, {
        method: 'GET', redirect: 'error', signal: controller.signal,
        headers: { Authorization: `Bearer ${credentials.token}`, Accept: 'application/json' }
      });
      if (!response.ok) {
        await response.body?.cancel();
        const scope = resource === 'products' ? 'read_catalog' : 'read_orders';
        throw new Error(`Ecwid ${resource} check returned HTTP ${response.status}. Check store ID, token and ${scope} permission. No automatic retry was made.`);
      }
      const body = await readBoundedJson(response, 16_384) as Record<string, unknown>;
      if (!body || typeof body !== 'object' || Array.isArray(body) ||
          !Number.isSafeInteger(body.total) || (body.total as number) < 0 ||
          !Number.isSafeInteger(body.count) || (body.count as number) < 0 || (body.count as number) > 1 ||
          body.offset !== 0 || (body.count as number) > (body.total as number)) {
        throw new Error('Unexpected Ecwid metadata response.');
      }
      checks.push({ resource, total: body.total as number });
    } catch (error) {
      // Never expose response bodies, request objects or transport errors containing credentials.
      if (error instanceof Error && /^Ecwid (products|orders) check returned HTTP \d+\./.test(error.message)) throw error;
      throw new Error(`Ecwid ${resource} check failed (network, timeout, redirect or invalid response). No automatic retry was made.`);
    } finally {
      clearTimeout(timer);
    }
  }
  return { store_id: credentials.storeId, connected: true, requests: 'GET_ONLY',
    checks, writes_performed: 0, imported_items: 0,
    note: 'Read access verified. This does not verify that the token lacks write permissions.' };
}

export async function loadReadOnlyCredentials() {
  const file = resolve('.env.ecwid-readonly');
  let text: string;
  try {
    const info = await stat(file);
    if (!info.isFile() || info.size > 16_384) throw new Error();
    if (process.platform !== 'win32' && (info.mode & 0o077)) {
      throw new Error('SET_PERMISSIONS');
    }
    text = await readFile(file, 'utf8');
  } catch (error) {
    if (error instanceof Error && error.message === 'SET_PERMISSIONS') {
      throw new Error('Make .env.ecwid-readonly private first: chmod 600 .env.ecwid-readonly');
    }
    throw new Error('Create a private .env.ecwid-readonly from .env.example. Do not put the token in command arguments or chat.');
  }
  return readOnlyCredentials(text);
}

async function main() {
  if (process.argv.includes('--help')) {
    console.log('Usage: npm run ecwid:check\nReads .env.ecwid-readonly and sends two GET metadata checks only. Never imports orders or changes stock.');
    return;
  }
  if (process.argv.length > 2) throw new Error('No credential arguments are accepted. Use .env.ecwid-readonly.');
  console.log(JSON.stringify(await checkReadOnlyAccess(await loadReadOnlyCredentials()), null, 2));
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main().catch(error => {
    console.error(error instanceof Error ? error.message : 'Ecwid read-only check failed.');
    process.exitCode = 1;
  });
}
