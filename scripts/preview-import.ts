import { open } from 'node:fs/promises';
import { previewImport, preparePreviewInput } from '../src/opening-import';

async function readBounded(path: string, maximum: number): Promise<string> {
  const file = await open(path, 'r');
  try {
    const bytes = Buffer.alloc(maximum + 1);
    let used = 0;
    while (used < bytes.length) {
      const result = await file.read(bytes, used, bytes.length - used, used);
      if (!result.bytesRead) break;
      used += result.bytesRead;
    }
    if (used > maximum) throw new Error('An import input exceeds the supported size limit.');
    return new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(bytes.subarray(0, used));
  } finally { await file.close(); }
}

const [stockPath, catalogPath, manifestPath] = process.argv.slice(2);
if (!stockPath || !catalogPath || !manifestPath) {
  console.error('Usage: npm run import:preview -- <stock.csv|source-candidates.json> catalog.json reservations.json');
  console.error('catalog.json accepts a complete READONLY_CATALOGUE snapshot or a reviewed stock-target array.');
  console.error('reservations.json must include source_ref, balance_meaning="PHYSICAL_ON_HAND", reservations_confirmed, and reservations:[{sku,quantity}]. Include store_id for real snapshots.');
  process.exitCode = 1;
} else {
  try {
    const [stock, catalog, manifest] = await Promise.all([
      readBounded(stockPath, 2_000_000), readBounded(catalogPath, 30_000_000), readBounded(manifestPath, 2_000_000)
    ]);
    const jsonSource = /\.json$/i.test(stockPath) || /^[\s\uFEFF]*[\[{]/.test(stock);
    const report = await previewImport(preparePreviewInput(stock, jsonSource ? 'SOURCE_CANDIDATES' : 'CSV', JSON.parse(catalog), JSON.parse(manifest)));
    console.log(JSON.stringify(report, null, 2));
    if (report.blocked_count || report.global_errors.length) process.exitCode = 2;
  } catch (error) {
    console.error(error instanceof Error ? error.message : 'Could not read the import files.');
    process.exitCode = 1;
  }
}
