import type { StockSourceReview } from '../src/source-review';

/** Escape source-workbook values for both HTML text and quoted attributes. */
function escapeHtml(value: unknown): string {
  return String(value ?? '').replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]!);
}

function display(value: unknown): string {
  return value === null || value === undefined || value === '' ? '—' : escapeHtml(value);
}

function sourceLink(value: string): string {
  try {
    const url = new URL(value);
    if (url.protocol === 'https:' && !url.username && !url.password &&
        ['docs.google.com', 'drive.google.com'].includes(url.hostname)) {
      return `<a href="${escapeHtml(url.href)}" target="_blank" rel="noopener noreferrer">Open source in Google Drive ↗</a>`;
    }
  } catch { /* An unrecognized source reference is plain text, never a link. */ }
  return display(value);
}

const script = `
(() => {
  const search = document.getElementById('stock-search');
  const rows = Array.from(document.querySelectorAll('[data-stock-row]'));
  const filters = Array.from(document.querySelectorAll('[data-filter]'));
  const visibleCount = document.getElementById('visible-count');
  const empty = document.getElementById('empty-state');
  let activeFilter = 'REVIEW';
  const apply = () => {
    const query = search.value.trim().toLocaleLowerCase();
    let visible = 0;
    for (const row of rows) {
      const matchesStatus = activeFilter === 'ALL' || row.dataset.status === activeFilter;
      const matchesQuery = !query || row.dataset.search.toLocaleLowerCase().includes(query);
      row.hidden = !matchesStatus || !matchesQuery;
      if (!row.hidden) visible += 1;
    }
    for (const button of filters) {
      button.setAttribute('aria-pressed', String(button.dataset.filter === activeFilter));
    }
    visibleCount.textContent = String(visible) + (visible === 1 ? ' item shown' : ' items shown');
    empty.hidden = visible !== 0;
  };
  for (const button of filters) {
    button.addEventListener('click', () => { activeFilter = button.dataset.filter; apply(); });
  }
  search.addEventListener('input', apply);
  apply();
})();
`;

/** A self-contained, read-only review. Workbook values never enter executable JavaScript. */
export function renderSourceReport(report: StockSourceReview): string {
  const { source, counts } = report;
  const rows = report.rows.map(row => {
    const statusLabel = row.status === 'CANDIDATE' ? 'Candidate' :
      row.status === 'REVIEW' ? 'Needs review' : 'Keep in workbook';
    const statusClass = row.status === 'CANDIDATE' ? 'candidate' :
      row.status === 'REVIEW' ? 'review' : 'workbook';
    const searchText = [row.source_sku, row.sku, row.name, row.store_id, row.location].join(' ');
    const normalizedSku = row.sku && row.sku !== row.source_sku
      ? `<span class="subtle">App match: <code>${escapeHtml(row.sku)}</code></span>` : '';
    const reasons = row.reasons.length
      ? `<ul class="reasons">${row.reasons.map(reason => `<li>${escapeHtml(reason.message)}</li>`).join('')}</ul>`
      : '<span class="subtle">No source-data exceptions recorded.</span>';
    const handling = row.proposed_authority === 'APP_AFTER_APPROVAL'
      ? 'Proposed: app after matching and approval' : 'Continue using workbook';
    return `<tr data-stock-row data-status="${escapeHtml(row.status)}" data-search="${escapeHtml(searchText)}"${row.status !== 'REVIEW' ? ' hidden' : ''}>
      <td class="reference"><span class="source-row">${escapeHtml(source.sheet_name)} · row ${escapeHtml(row.source_row)}</span><span class="subtle">Store ID <code>${display(row.store_id)}</code></span></td>
      <td class="item"><strong>${display(row.name)}</strong><span class="subtle">${display(row.item_type)}${row.source_status ? ` · ${escapeHtml(row.source_status)}` : ''}</span><span class="subtle">Location: ${display(row.location)}</span></td>
      <td><code class="raw-sku">${display(row.source_sku)}</code>${normalizedSku}</td>
      <td class="quantity"><strong>${display(row.balance)}</strong><span class="subtle">${display(row.unit)}</span></td>
      <td class="handling"><span class="status ${statusClass}">${statusLabel}</span><span class="subtle current-authority">Workbook manages this item now.</span><span class="subtle">${handling}</span></td>
      <td class="notes">${reasons}</td>
    </tr>`;
  }).join('\n');

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src 'none'; connect-src 'none'; base-uri 'none'; form-action 'none'">
  <meta name="referrer" content="no-referrer">
  <title>Opening stock review</title>
  <style>
    :root{color-scheme:light;--ink:#173039;--muted:#53656c;--line:#d9e2e4;--teal:#08675d;--paper:#fff;--back:#f3f6f5;--amber:#83520c}
    *{box-sizing:border-box}body{margin:0;background:var(--back);color:var(--ink);font:15px/1.5 ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}main{max-width:1512px;margin:auto;padding:40px 28px 56px}h1{font-size:clamp(28px,4vw,40px);line-height:1.12;margin:12px 0}h2{font-size:19px;margin:0 0 10px}p{margin:8px 0}a{color:var(--teal);text-underline-offset:3px}button,input{font:inherit}button{cursor:pointer}code{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;font-size:.92em;overflow-wrap:anywhere}.eyebrow{color:var(--teal);font-size:12px;font-weight:750;letter-spacing:.12em;text-transform:uppercase}.intro{max-width:940px;color:var(--muted);font-size:16px}.guardrail{display:flex;gap:12px;align-items:flex-start;background:#e4f1ed;border:1px solid #c7ded5;border-radius:12px;padding:16px 18px;margin:24px 0}.guardrail-mark{font-size:20px;color:var(--teal)}.guardrail strong{display:block}.guardrail p{margin:2px 0 0;color:#32564e}.stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:24px 0}.stat{background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:18px}.stat .number{font-size:32px;line-height:1.2;font-weight:750;display:block;font-variant-numeric:tabular-nums}.stat .label{display:block;font-weight:650;margin-top:5px}.stat .hint{font-size:13px;color:var(--muted);display:block;margin-top:6px}.stat.review .number{color:var(--amber)}.stat.candidate .number{color:var(--teal)}.review-panel{background:var(--paper);border:1px solid var(--line);border-radius:14px;overflow:hidden}.toolbar{padding:20px;border-bottom:1px solid var(--line);display:flex;flex-wrap:wrap;gap:16px;align-items:flex-end;justify-content:space-between}.search-label{display:block;font-size:13px;font-weight:650;margin-bottom:6px}.search-wrap{flex:1 1 320px;max-width:440px}input[type=search]{width:100%;border:1px solid #9cafb5;border-radius:8px;background:#fff;padding:11px 12px;color:var(--ink)}input:focus,button:focus-visible,a:focus-visible{outline:3px solid #80bcb2;outline-offset:3px}.filters{display:flex;gap:7px;flex-wrap:wrap}.filters button{border:1px solid var(--line);background:#f8faf9;border-radius:8px;padding:9px 12px;color:var(--ink);font-size:13px;font-weight:600}.filters button[aria-pressed=true]{background:var(--ink);color:white;border-color:var(--ink)}.table-caption{display:flex;justify-content:space-between;gap:12px;padding:12px 20px;font-size:13px;color:var(--muted)}.table-scroll{overflow:auto;max-height:72vh}table{width:100%;min-width:1070px;border-collapse:collapse;text-align:left;font-size:13px}th{position:sticky;top:0;background:#edf2f1;z-index:1;padding:12px 14px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;vertical-align:bottom;border-bottom:1px solid var(--line)}td{padding:16px 14px;border-bottom:1px solid #e6eded;vertical-align:top}tbody tr:last-child td{border-bottom:0}tbody tr:hover{background:#f8fbfa}.reference{width:140px}.source-row{display:block;font-size:12px}.item{min-width:200px;max-width:280px}.item strong{font-size:14px;overflow-wrap:anywhere}.raw-sku{white-space:pre-wrap}.quantity{width:100px;font-variant-numeric:tabular-nums}.quantity strong{font-size:17px;white-space:nowrap}.handling{width:220px}.notes{min-width:230px;max-width:410px}.subtle{display:block;color:var(--muted);font-size:12px;margin-top:5px}.status{display:inline-block;border-radius:5px;padding:3px 7px;font-size:11px;font-weight:700}.status.candidate{background:#e1f0ea;color:#19594b}.status.review{background:#fff0d5;color:#79510f}.status.workbook{background:#eef0f3;color:#4b5867}.reasons{margin:0;padding-left:16px}.reasons li+li{margin-top:6px}.empty-state{padding:36px 20px;text-align:center;color:var(--muted)}[hidden]{display:none!important}.next{border-left:3px solid #76ada1;padding:3px 0 3px 18px;margin:26px 0}.next p{color:var(--muted);max-width:1000px}.provenance{margin-top:28px;padding-top:20px;border-top:1px solid var(--line);font-size:12px;color:var(--muted)}.provenance summary{cursor:pointer;font-size:13px;color:var(--ink);font-weight:650}.source-grid{display:grid;grid-template-columns:130px minmax(0,1fr);gap:8px 16px;max-width:1100px;margin:14px 0 0}.source-grid dt{font-weight:650}.source-grid dd{margin:0;overflow-wrap:anywhere}.noscript{padding:14px;background:#fff0d5;color:#79510f;border-radius:8px}
    @media(max-width:700px){main{padding:24px 14px 36px}.stats{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.stat{padding:14px}.stat .number{font-size:28px}.guardrail{padding:14px}.toolbar{padding:14px}.search-wrap{max-width:none;flex-basis:100%}.filters button{padding:9px 10px}.table-caption{padding:12px 14px}.table-caption .offline{display:none}.source-grid{grid-template-columns:1fr;gap:2px}.source-grid dd{margin-bottom:10px}.next{margin-top:22px}}
    @media print{body{background:white}main{max-width:none;padding:0}.toolbar,.offline{display:none}.table-scroll{max-height:none;overflow:visible}table{min-width:0;font-size:10px}th{position:static}td,th{padding:8px}.stats{gap:8px}.stat{padding:10px}.provenance{break-inside:avoid}.table-caption{padding-left:0}.review-panel{border:0}.item,.notes{min-width:0}.handling{width:auto}}
  </style>
</head>
<body>
<main>
  <header><span class="eyebrow">Factory inventory · source review</span><h1>Opening stock review</h1><p class="intro">A read-only check of the factory stock workbook, starting with simple items counted in whole units. Balance means physical stock held in the factory; it is not yet an Ecwid available-to-sell quantity.</p></header>
  <section class="guardrail" aria-label="Review status"><span class="guardrail-mark" aria-hidden="true">○</span><div><strong>No stock has been changed. Every item is still workbook-managed.</strong><p>A candidate has passed source-data checks only. Ecwid products, options and outstanding orders have not been checked. Moving an item to the app requires a later review and approval.</p></div></section>
  <section class="stats" aria-label="Source review totals">
    <div class="stat"><span class="number" data-count="total">${escapeHtml(counts.total)}</span><span class="label">Items checked</span><span class="hint">Rows in the source stock table</span></div>
    <div class="stat candidate"><span class="number" data-count="candidate">${escapeHtml(counts.candidate)}</span><span class="label">Candidates</span><span class="hint">Next: read-only Ecwid matching</span></div>
    <div class="stat review"><span class="number" data-count="review">${escapeHtml(counts.review)}</span><span class="label">Need review</span><span class="hint">Resolve source-data exceptions first</span></div>
    <div class="stat"><span class="number" data-count="keep-workbook">${escapeHtml(counts.keep_workbook)}</span><span class="label">Keep in workbook</span><span class="hint">Outside the first app rollout</span></div>
  </section>
  <section class="review-panel" aria-label="Stock items">
    <div class="toolbar"><div class="search-wrap"><label class="search-label" for="stock-search">Find an item</label><input type="search" id="stock-search" placeholder="SKU, item, Store ID or location" autocomplete="off"></div><div class="filters" role="group" aria-label="Filter items"><button type="button" data-filter="ALL" aria-pressed="false">All</button><button type="button" data-filter="CANDIDATE" aria-pressed="false">Candidates</button><button type="button" data-filter="REVIEW" aria-pressed="true">Needs review</button><button type="button" data-filter="KEEP_WORKBOOK" aria-pressed="false">Keep in workbook</button></div></div>
    <div class="table-caption"><span id="visible-count" role="status" aria-live="polite">${escapeHtml(counts.review)} items shown</span><span class="offline">Search and filters work offline · no data is sent</span></div>
    <noscript><p class="noscript">JavaScript is disabled. Showing only items needing review; search and other filters require JavaScript.</p></noscript>
    <div class="table-scroll" role="region" aria-label="Stock review table; scroll horizontally for all columns" tabindex="0"><table><thead><tr><th scope="col">Source reference</th><th scope="col">Item / location</th><th scope="col">Workbook SKU</th><th scope="col">Physical balance</th><th scope="col">Handling</th><th scope="col">Review notes</th></tr></thead><tbody>${rows}</tbody></table></div>
    <div id="empty-state" class="empty-state"${counts.review ? ' hidden' : ''}>No items match this view. Try another filter or search.</div>
  </section>
  <section class="next"><h2>What happens next</h2><p>Match candidate SKUs to Ecwid without changing stock, verify that each is a simple single-unit product, and account for outstanding unpicked orders. Then review the proposed opening balances before any one-time alignment. Items outside the rollout stay in the workbook; approved app-managed items must not also have their movements recorded there.</p><p>This review uses the workbook’s saved formula results, not a fresh Excel recalculation. Refresh the source and confirm physical quantities before cutover.</p></section>
  <footer class="provenance"><details><summary>Source and audit details</summary><dl class="source-grid"><dt>Workbook</dt><dd>${display(source.file_name)}</dd><dt>Source tab</dt><dd>${display(source.sheet_name)} · header row ${display(source.header_row)}</dd><dt>Source reference</dt><dd>${sourceLink(source.source_ref)}</dd><dt>Source modified</dt><dd>${display(source.source_modified_at)}</dd><dt>Snapshot extracted</dt><dd>${display(source.extracted_at)}</dd><dt>Review generated</dt><dd>${display(report.generated_at)}</dd><dt>Snapshot SHA-256</dt><dd><code>${display(source.sha256)}</code></dd><dt>Review boundary</dt><dd>Source checks only. Ecwid has not been checked; outstanding-order reservations have not been confirmed. This report cannot apply stock changes.</dd></dl></details></footer>
</main>
<script>${script}</script>
</body>
</html>`;
}
