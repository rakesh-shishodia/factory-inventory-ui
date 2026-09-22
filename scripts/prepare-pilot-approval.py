"""Produce a private, read-only approval proposal from reproducible pilot snapshots.

This is not an import format. No database, workbook or Ecwid mutation is supported.
Prior unit clarification is carried only for the exact reviewed snapshot and SKUs.
"""
import argparse
from collections import Counter
from copy import deepcopy
from datetime import datetime, timezone
import html
import importlib.util
import json
import os
from pathlib import Path
from urllib.parse import urlsplit
from zoneinfo import ZoneInfo

_spec = importlib.util.spec_from_file_location('pilot_review', Path(__file__).with_name('review-pilot-scope.py'))
pilot = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pilot)

CONTEXT_REVIEW_SHA = '4514021cd36c226441c79d5f62adacc82cb6a1c04acd993ddc16f22031c84e9a'
PRIOR_UNIT_SKUS = frozenset('''
LP-M5B-20 LP-M5B-25 LP-M5B-30 LP-M5B-35 LP-M5B-40 LP-M5B-45
LP-M5B-50 LP-M5B-55 LP-M5B-6 LP-M5B-60 LP-M5B-65
TND20-M3 TND20-M4 TND30-M3 TND40-M3 TND40-M4 TNLS40-M6
TNP20-M3 TNP20-M4 TNP20-M5 TNP20-M6 TNP30-M3 TNP30-M4
TNP30-M5 TNP30-M6 TNP30-M8 TNP45-M5 TNPS20-M3 TNPS20-M4
TNPS20-M5 TNS30-M3 TNS30-M4
'''.split())
SET_QUESTIONS = {
    '00691': 'Does 90 mean 90 complete riser plate sets, rather than 90 loose plates?',
    '00696': 'Does 10 mean 10 complete assembled carriages, rather than components?',
    'CUBEC3W-20-SET': 'Is this workbook row counted in complete connector sets, including all parts sold together?',
}
GROUPS = {
    'PRIOR_UNIT_CLARIFICATION': 'Unit already clarified',
    'GROUP_CONFIRMATION': 'Confirm one-to-one unit',
    'SET_ASSEMBLY_HOLD': 'Hold: set / assembly',
}
INPUT_KEYS = ('csv', 'source_review', 'live_catalog', 'live_scope', 'orders', 'pick_confirmation')


def reproduce_review(review_bytes, inputs):
    """Reject changed inputs or derived quantities instead of displaying stale joins."""
    report = json.loads(review_bytes)
    for key in INPUT_KEYS:
        data = inputs[key]
        recorded = report.get('inputs', {}).get(key, {})
        if recorded.get('sha256') != pilot.sha(data) or recorded.get('bytes') != len(data):
            raise ValueError(f'{key} differs from the pilot review snapshot.')
    scope = pilot.parse_scope_csv(inputs['csv'])
    source = pilot.parse_source(inputs['source_review'])
    live = pilot.catalogue.parse_live_catalogue(inputs['live_catalog'], 'catalog.json', '2442119')
    scope_check = pilot.verify_live_scope(json.loads(inputs['live_scope']), scope, '2442119')
    reservations = pilot.reservation_review(inputs['orders'], inputs['pick_confirmation'], '2442119')
    expected = pilot.create_review(scope, source, live['stock_targets'], reservations, scope_check)
    ignored = {'generated_at', 'inputs'}
    if ({k: v for k, v in report.items() if k not in ignored} !=
            {k: v for k, v in expected.items() if k not in ignored}):
        raise ValueError('Pilot review does not reproduce from the recorded source snapshots.')
    return report


def create_proposal(report, review_sha):
    if (report.get('kind') != 'PILOT_SCOPE_REVIEW' or report.get('schema_version') != 1 or
            report.get('dry_run') is not True or report.get('store_id') != '2442119' or
            report.get('current_authority') != 'WORKBOOK' or report.get('ready_count') != 0 or
            report.get('scope', {}).get('live_crosscheck', {}).get('verified') is not True or
            any(report.get(key) is not False for key in ('stock_writes_performed', 'database_imports_performed',
                                                        'single_unit_mappings_approved', 'cutover_approved'))):
        raise ValueError('An unapproved, verified, read-only pilot review is required.')
    rows, shortfalls, seen_skus, seen_targets = [], [], set(), set()
    for original in report['rows']:
        issues = {item['code'] for item in original['issues']}
        matched = original.get('technical_match') is True
        shortfall = 'INSUFFICIENT_PHYSICAL_BALANCE' in issues
        if not matched and not shortfall:
            continue
        row = deepcopy(original)
        identity = (row['ecwid_product_id'], row['ecwid_combination_id'])
        if not row['sku'] or row['sku'] in seen_skus or identity in seen_targets:
            raise ValueError('Approval rows must have unique SKU and live identities.')
        seen_skus.add(row['sku'])
        seen_targets.add(identity)
        sources = row['source_rows']
        physical, reserved = row['physical_balance'], row['reserved_ecwid_units']
        if (len(sources) != 1 or sources[0]['status'] != 'CANDIDATE' or
                not pilot.whole(physical) or not pilot.whole(reserved) or sources[0]['balance'] != physical or
                row['unresolved_order_lines'] or row['ready'] is not False or
                row['desired_ecwid_quantity'] is not None or row['single_unit_confirmed'] is not False or
                any(not pilot.whole(line['quantity']) or line['quantity'] == 0 for line in row['reservation_lines']) or
                sum(line['quantity'] for line in row['reservation_lines']) != reserved):
            raise ValueError('Approval row source, reservation or approval state is invalid.')
        row['opening_balance_approved'] = False
        row['unit_clarification_carried'] = False
        row['question'] = None
        if shortfall:
            if matched or physical >= reserved or row['provisional_quantity_if_one_to_one'] is not None:
                raise ValueError('Invalid stock shortfall.')
            row['shortfall_if_one_to_one'] = reserved - physical
            row['approval_group'] = 'STOCK_SHORTFALL_HOLD'
            shortfalls.append(row)
            continue
        if (issues != {'SINGLE_UNIT_CONFIRMATION_REQUIRED'} or physical < reserved or
                row['provisional_quantity_if_one_to_one'] != physical - reserved):
            raise ValueError('Technical match has unresolved checks or incorrect arithmetic.')
        row['approval_group'] = 'GROUP_CONFIRMATION'
        # This display-only context is not an approved importer mapping. A new
        # report hash intentionally requires a new review of the clarification.
        if review_sha == CONTEXT_REVIEW_SHA and row['sku'] in PRIOR_UNIT_SKUS and row['ecwid_combination_id']:
            row['approval_group'] = 'PRIOR_UNIT_CLARIFICATION'
            row['unit_clarification_carried'] = True
        if row['sku'] in SET_QUESTIONS:
            row['approval_group'] = 'SET_ASSEMBLY_HOLD'
            row['question'] = SET_QUESTIONS[row['sku']] if review_sha == CONTEXT_REVIEW_SHA else (
                'Is one workbook unit a complete set or assembly matching one Ecwid sale unit?')
        elif any(word in (row['name'] or '').lower() for word in (' set', 'bundle', 'assembled', ' kit', ' pack')):
            row['approval_group'] = 'SET_ASSEMBLY_HOLD'
            row['unit_clarification_carried'] = False
            row['question'] = 'Does the workbook count complete sale units, not loose parts or pack contents?'
        rows.append(row)
    counts = Counter(row['approval_group'] for row in rows)
    if len(rows) != report['counts']['technical_matches_awaiting_unit_confirmation']:
        raise ValueError('Technical-match count differs from the pilot review.')
    return {
        'kind': 'PILOT_APPROVAL_PROPOSAL', 'schema_version': 1, 'dry_run': True,
        'generated_at': datetime.now(timezone.utc).isoformat(), 'store_id': report['store_id'],
        'current_authority': 'WORKBOOK', 'batch_approved': False, 'cutover_approved': False,
        'ecwid_write_authorized': False, 'stock_writes_performed': False, 'database_imports_performed': False,
        'review_sha256': review_sha, 'source': deepcopy(report['source']),
        'inputs': deepcopy(report.get('inputs', {})), 'order_review': deepcopy(report['order_review']),
        'counts': {'matched_rows': len(rows), 'stock_shortfalls': len(shortfalls),
                   **{key: counts[key] for key in GROUPS}},
        'prior_unit_context': {
            'user_statement': 'The nuts and bolts which have variations are single-unit items; each variation has an independent SKU.',
            'context_review_sha256': CONTEXT_REVIEW_SHA,
            'applied_only_to_exact_snapshot': review_sha == CONTEXT_REVIEW_SHA,
            'is_stock_or_cutover_approval': False,
        },
        'rows': rows, 'shortfalls': shortfalls,
        'note': 'Read-only proposal, not an import. Proposed quantities assume one-to-one units. '
                'Approval is for a test-batch plan only; refresh stock and orders and seek separate approval before live alignment.',
    }


def timestamp(value):
    try:
        return datetime.fromisoformat(value.replace('Z', '+00:00')).astimezone(ZoneInfo('Asia/Kolkata')).strftime('%d %b %Y, %H:%M IST')
    except (ValueError, TypeError, AttributeError):
        return 'Not recorded'


def render_report(report):
    esc = lambda value: html.escape(str(value), quote=True)
    number = lambda value: 'Not available' if value is None else f'{value:,}'

    def item(row):
        options = ' · '.join(f"{option['name']}: {option['value']}" for option in row['variation_options'] or [])
        source = row['source_rows'][0]
        return (f'<strong class="sku">{esc(row["sku"])}</strong><span>{esc(row["name"])}</span>'
                f'<small>{esc(options)}</small><small>{esc(source["location"])} · '
                f'{esc(source["source_sheet"])} row {esc(source["source_row"])}</small>')

    def orders(row):
        return '<small>' + '<br>'.join(f'Order {esc(line["order_id"])}: {line["quantity"]}' for line in row['reservation_lines']) + '</small>'

    table_rows = []
    for row in report['rows']:
        group = row['approval_group']
        table_rows.append(f'''<tr data-group="{group}"><td class="item">{item(row)}</td>
<td data-label="Workbook balance">{number(row['physical_balance'])}</td>
<td data-label="Pending orders">{number(row['reserved_ecwid_units'])}{orders(row)}</td>
<td data-label="Proposed Ecwid*" class="proposed">{number(row['provisional_quantity_if_one_to_one'])}</td>
<td data-label="Ecwid snapshot">{number(row['ecwid_quantity'])}</td>
<td class="unit"><span class="badge {group}">{GROUPS[group]}</span><details><summary>Mapping details</summary>
<small>Product {esc(row['ecwid_product_id'])} · Variation {esc(row['ecwid_combination_id'] or 'none')}<br>
Workbook name: {esc(row['source_rows'][0]['name'])}<br>Workbook unit: {esc(row['source_rows'][0]['unit'])}<br>
{esc(' / '.join(row['category_paths']).replace(chr(92) + '/', '/'))}</small></details></td></tr>''')
    questions = ''.join(f'<li><strong>{esc(row["sku"])}</strong> — {esc(row["name"])}<p>{esc(row["question"])}</p></li>'
                        for row in report['rows'] if row['approval_group'] == 'SET_ASSEMBLY_HOLD')
    shortfalls = ''.join(f'''<tr><td class="item">{item(row)}</td><td data-label="Workbook balance">{number(row['physical_balance'])}</td>
<td data-label="Pending orders">{number(row['reserved_ecwid_units'])}{orders(row)}</td><td data-label="Shortfall if 1:1">{number(row['shortfall_if_one_to_one'])}</td>
<td class="unit">Recount / resolve units<small>{esc('; '.join(issue['message'] for issue in row['issues'] if issue['code'] not in {'INSUFFICIENT_PHYSICAL_BALANCE', 'SINGLE_UNIT_CONFIRMATION_REQUIRED'}))}</small></td></tr>'''
                         for row in report['shortfalls'])
    counts = report['counts']
    source_url = report['source'].get('source_ref', '')
    if urlsplit(source_url).scheme != 'https' or urlsplit(source_url).hostname != 'docs.google.com':
        source_url = ''
    source_link = f'<a href="{esc(source_url)}" target="_blank" rel="noopener noreferrer">Source workbook</a>' if source_url else 'Source workbook'
    provenance = esc(json.dumps({'review_sha256': report['review_sha256'], 'source': report['source'], 'inputs': report['inputs']}, ensure_ascii=False, indent=2))
    return f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="referrer" content="no-referrer"><title>Pilot approval list</title>
<style>
:root{{color-scheme:light;--ink:#183047;--muted:#566777;--border:#dce4e9;--paper:#fff;--teal:#126a61}}
*{{box-sizing:border-box}}body{{margin:0;background:#f5f7f9;color:var(--ink);font:15px/1.5 system-ui,sans-serif}}
main{{max-width:1460px;margin:0 auto;padding:28px 24px 64px}}h1{{font-size:30px;line-height:1.2;margin:6px 0 12px}}h2{{font-size:21px;margin:0 0 10px}}
p{{margin:8px 0 14px}}a{{color:#155f95}}.eyebrow{{font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--teal)}}
.intro{{max-width:930px}}.notice{{background:#e6f3ee;border-left:4px solid var(--teal);padding:12px 16px;margin:16px 0}}
.stats{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:20px 0}}.stat{{background:white;border:1px solid var(--border);border-radius:10px;padding:14px 18px}}
.stat strong{{display:block;font-size:28px;line-height:1.3}}.stat span{{color:var(--muted);font-size:13px}}
.checks{{background:white;border:1px solid var(--border);border-radius:10px;margin:12px 0;padding:14px 18px}}
summary{{cursor:pointer;font-weight:600}}summary:focus-visible,a:focus-visible,input:focus-visible,select:focus-visible{{outline:3px solid #719acc;outline-offset:3px}}
.checks[open]>summary{{margin-bottom:12px}}.checks li{{margin-bottom:12px}}.checks li p{{margin:3px 0}}
.instructions{{margin:22px 0}}.instructions li{{margin:6px 0}}.muted,small{{color:var(--muted)}}small{{display:block;font-size:12px;line-height:1.45;margin-top:5px}}
.toolbar{{display:flex;align-items:end;gap:12px;flex-wrap:wrap;margin:16px 0}}label{{display:grid;gap:5px;font-size:13px;font-weight:600}}
input,select{{font:inherit;background:white;border:1px solid #a6b5c1;border-radius:6px;padding:10px;min-height:42px}}input{{width:280px}}#count{{padding-bottom:10px;color:var(--muted)}}
.table-wrap{{background:white;border:1px solid var(--border);border-radius:10px;overflow:auto}}table{{width:100%;border-collapse:collapse;text-align:left;font-size:14px}}
th{{background:#eaf0f4;font-size:12px;color:#394f61;text-align:left}}th,td{{padding:14px 12px;vertical-align:top;border-bottom:1px solid var(--border)}}tbody tr:last-child td{{border-bottom:0}}
.item{{min-width:235px;max-width:360px}}.item>span{{display:block;margin-top:4px}}.sku{{font-size:14px;overflow-wrap:anywhere}}.proposed{{font-weight:700;color:var(--teal)}}
.badge{{font-size:11px;font-weight:650;display:inline-block;padding:4px 7px;border-radius:4px;background:#edf1f4;white-space:nowrap}}
.PRIOR_UNIT_CLARIFICATION{{background:#e6f3ee;color:#155d51}}.SET_ASSEMBLY_HOLD{{background:#fff0cf;color:#7c5010}}.unit details{{margin-top:8px}}.unit summary{{font-size:12px;color:var(--muted);font-weight:400}}
.footnote{{font-size:13px;color:var(--muted)}}pre{{white-space:pre-wrap;overflow-wrap:anywhere;font-size:11px}}[hidden]{{display:none!important}}
@media(max-width:850px){{main{{padding:22px 16px 40px}}.stats{{grid-template-columns:repeat(2,1fr)}}.stat{{padding:12px}}.toolbar label,input,select{{width:100%}}.toolbar{{gap:10px}}h1{{font-size:27px}}
.table-wrap{{border:0;background:transparent;overflow:visible}}table,tbody{{display:block}}thead{{display:none}}tbody tr{{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--border);border-radius:9px;background:white;margin:12px 0;padding:4px 12px 10px;gap:0 12px}}
td{{display:block;padding:10px 0;border-bottom:0}}td[data-label]:before{{content:attr(data-label);display:block;font-size:11px;font-weight:400;color:var(--muted);margin-bottom:3px}}
.item,.unit{{grid-column:1/-1;max-width:none;min-width:0}}.item{{border-bottom:1px solid var(--border)}}.checks{{padding:12px}}}}
@media print{{body{{background:white}}main{{padding:0}}.toolbar,.unit details{{display:none}}.table-wrap{{overflow:visible}}tr{{break-inside:avoid}}}}
</style></head><body><main>
<div class="eyebrow">Factory inventory · Store {esc(report['store_id'])} · Review only</div>
<h1>Approve the first pilot batch</h1>
<p class="intro">{counts['matched_rows']} SKU matches from your chosen categories. Review the saved balances below; unresolved products stay workbook-managed. This is a proposal, not a live stock update.</p>
<div class="notice">No stock imported. No workbook or Ecwid stock changed. Your workbook remains the stock authority.</div>
<div class="stats"><div class="stat"><strong>{counts['PRIOR_UNIT_CLARIFICATION']}</strong><span>Nut / screw variations — units already clarified</span></div>
<div class="stat"><strong>{counts['GROUP_CONFIRMATION']}</strong><span>Other matches — confirm one-to-one units</span></div>
<div class="stat"><strong>{counts['SET_ASSEMBLY_HOLD']}</strong><span>Sets / assemblies — held for clarification</span></div>
<div class="stat"><strong>{counts['stock_shortfalls']}</strong><span>Additional stock discrepancies — held for recount</span></div></div>
<section class="instructions"><h2>What I need from you</h2><ol>
<li>Review the {counts['PRIOR_UNIT_CLARIFICATION'] + counts['GROUP_CONFIRMATION']} non-held matches. Confirm the proposed opening balances for the test-batch plan, or give exceptions. For the {counts['GROUP_CONFIRMATION']} other items, also confirm that one workbook unit equals one Ecwid sale unit.</li>
<li>Answer the set / assembly questions, or leave those SKUs out for now.</li>
<li>Recount the six discrepant SKUs when convenient; they can stay excluded from the first batch.</li></ol>
<p class="footnote">Reply in our chat. This page does not save approvals. These decisions do not authorise a live Ecwid update.</p></section>
<details class="checks" id="unit-questions"><summary>{counts['SET_ASSEMBLY_HOLD']} set / assembly questions</summary><ul>{questions}</ul>
<p class="footnote">A complete prepacked set can be one sale unit. Loose components need a separate mapping; no conversion has been assumed.</p></details>
<details class="checks" id="stock-discrepancies"><summary>{counts['stock_shortfalls']} stock discrepancies — outside the proposed batch</summary>
<p>These saved balances are smaller than the quantities in the confirmed unpicked orders. Please provide the actual shelf count and whether the order items are still on the shelf. Do not infer a return or overwrite the balance to make it fit.</p>
<div class="table-wrap"><table><thead><tr><th>SKU / source</th><th>Workbook balance</th><th>Pending units</th><th>Shortfall if 1:1</th><th>Other checks</th></tr></thead><tbody>{shortfalls}</tbody></table></div>
<p class="footnote">Set quantities need unit clarification too. Five of these six targets also need Ecwid stock-tracking review. No negative or zero-clamped write quantity has been proposed.</p></details>
<section id="matches" style="margin-top:28px"><h2>The {counts['matched_rows']} matched SKUs</h2>
<p class="footnote">* Proposed Ecwid = workbook physical balance − unpicked order units, only if the units match one-to-one. Held items show comparison arithmetic, not approved quantities. Ecwid snapshot values are not the stock authority.</p>
<div class="toolbar"><label for="search">Find a SKU, item, location or order<input type="search" id="search" placeholder="e.g. LP-M5B-20 or 9437"></label>
<label for="group">Review group<select id="group"><option value="">All {counts['matched_rows']} matches</option><option value="NON_HELD">{counts['PRIOR_UNIT_CLARIFICATION'] + counts['GROUP_CONFIRMATION']} non-held matches</option>
{''.join(f'<option value="{key}">{counts[key]} — {label}</option>' for key, label in GROUPS.items())}</select></label><output id="count" aria-live="polite"></output></div>
<div class="table-wrap"><table id="match-table"><thead><tr><th>SKU / item / location</th><th>Workbook balance</th><th>Pending order units</th><th>Proposed Ecwid*</th><th>Ecwid snapshot</th><th>Unit review</th></tr></thead><tbody>{''.join(table_rows)}</tbody></table></div>
<p id="empty" hidden>No matches. Clear the search or change the review group.</p></section>
<p class="footnote" style="margin-top:24px">{source_link} · Saved workbook modified {esc(timestamp(report['source'].get('source_modified_at')))}.<br>
Order snapshot: {esc(timestamp(report['order_review'].get('order_snapshot_completed_at')))} · {report['order_review'].get('pending_order_count', 'Unknown')} paid, unpicked orders.
Ready for Pickup means picked and packed, equivalent to Shipped; those orders are excluded. Refresh stock and orders before any cutover.</p>
<details class="checks"><summary>Source files and verification hashes</summary><pre>{provenance}</pre></details>
</main><script>
const search=document.querySelector('#search'),group=document.querySelector('#group'),rows=[...document.querySelectorAll('#match-table tbody tr')];
function filter(){{let count=0;for(const row of rows){{const selected=!group.value||(group.value==='NON_HELD'?row.dataset.group!=='SET_ASSEMBLY_HOLD':row.dataset.group===group.value);row.hidden=!(selected&&row.textContent.toLowerCase().includes(search.value.trim().toLowerCase()));if(!row.hidden)count++}}document.querySelector('#count').textContent=count+' of '+rows.length+' matches';document.querySelector('#empty').hidden=count!==0}}
search.addEventListener('input',filter);group.addEventListener('change',filter);filter();
</script></body></html>'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--review', required=True)
    for key in INPUT_KEYS:
        parser.add_argument('--' + key.replace('_', '-'), required=True)
    parser.add_argument('--out', required=True, help='New private directory underneath import-data')
    args = parser.parse_args()
    review_bytes = pilot.catalogue.bounded_bytes(args.review)
    inputs = {key: pilot.catalogue.bounded_bytes(getattr(args, key)) for key in INPUT_KEYS}
    report = reproduce_review(review_bytes, inputs)
    proposal = create_proposal(report, pilot.sha(review_bytes))
    output = Path(args.out).resolve()
    private_root = Path('import-data').resolve()
    if output == private_root or private_root not in output.parents or output.exists():
        raise ValueError('Output must be a new directory under private import-data.')
    page = render_report(proposal)
    output.mkdir(mode=0o700)
    for name, content in [('approval-list.json', json.dumps(proposal, ensure_ascii=False, indent=2) + '\n'), ('approval.html', page)]:
        descriptor = os.open(output / name, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        with os.fdopen(descriptor, 'w', encoding='utf-8') as stream:
            stream.write(content)
    print(json.dumps({'counts': proposal['counts'], 'batch_approved': False, 'stock_writes_performed': False, 'output': str(output)}, indent=2))


if __name__ == '__main__':
    main()
