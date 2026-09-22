import { execFile } from 'node:child_process';
import { mkdir, realpath, stat, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath, URL } from 'node:url';
import { parseArgs, promisify } from 'node:util';
import { reviewStockSource } from '../src/source-review';
import { renderSourceReport } from './source-report';

const run = promisify(execFile);

/** Local preparation only: no imports from the Worker, no API clients or DB. */
async function main() {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      out: { type: 'string' },
      'source-ref': { type: 'string' },
      'source-modified-at': { type: 'string' },
      python: { type: 'string' },
      help: { type: 'boolean' }
    }
  });
  if (values.help) {
    console.log('Usage: npm run import:source -- stock.xlsm --out import-data/review-YYYYMMDD --source-ref <source URL> [--source-modified-at ISO-date] [--python /path/to/python3]');
    console.log('Writes a new local review folder only. Never changes the workbook, app stock or Ecwid. Python 3.11+ is required; no Python packages are needed.');
    return;
  }
  if (positionals.length !== 1 || !values.out || !values['source-ref']?.trim()) {
    throw new Error('Provide a workbook, a new --out directory, and --source-ref. Use --help for details.');
  }
  const workbook = await realpath(positionals[0]);
  const output = resolve(values.out);
  // Reports hold operational data. Keep them outside deployable public assets.
  const publicDir = await realpath(resolve('public'));
  const outputParent = await realpath(resolve(output, '..'));
  if (outputParent === publicDir || outputParent.startsWith(publicDir + '/')) {
    throw new Error('Private stock reports must not be placed under public/. Use import-data/.');
  }
  try {
    await stat(output);
    throw new Error('The output directory already exists. Use a new directory to preserve prior reviews.');
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
  }
  const args = [fileURLToPath(new URL('./extract_stock_workbook.py', import.meta.url)), workbook,
    '--source-ref', values['source-ref']];
  if (values['source-modified-at']) args.push('--source-modified-at', values['source-modified-at']);
  const { stdout } = await run(values.python || process.env.INVENTORY_PYTHON || 'python3', args,
    { timeout: 60_000, maxBuffer: 30_000_000, encoding: 'utf8' });
  const snapshot: unknown = JSON.parse(stdout);
  const report = reviewStockSource(snapshot);
  const candidates = {
    kind: 'SOURCE_CANDIDATES', dry_run: true,
    source_ref: report.source.source_ref, source_hash: report.source.sha256,
    balance_meaning: 'PHYSICAL_ON_HAND', ecwid_checked: false,
    reservations_confirmed: false, rows: report.candidate_rows
  };
  const files: [string, string][] = [
    ['source-snapshot.json', JSON.stringify(snapshot, null, 2) + '\n'],
    ['source-review.json', JSON.stringify(report, null, 2) + '\n'],
    ['candidates.json', JSON.stringify(candidates, null, 2) + '\n'],
    ['review.html', renderSourceReport(report)]
  ];
  await mkdir(output, { mode: 0o700 });
  for (const [name, content] of files) {
    await writeFile(resolve(output, name), content, { flag: 'wx', mode: 0o600 });
  }
  console.log(JSON.stringify({ report: resolve(output, 'review.html'), counts: report.counts,
    ecwid_checked: false, migrated_items: 0, source_hash: report.source.sha256 }, null, 2));
}

main().catch(error => {
  console.error(error instanceof Error ? error.message : 'Could not prepare the stock review.');
  process.exitCode = 1;
});
