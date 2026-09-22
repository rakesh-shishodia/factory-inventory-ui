"""Read-only pilot category / stock coverage review. Never creates an import.

CSV category paths select parent products. Live stock targets retain independent
variation IDs. All workbook rows participate, including excluded/problem rows.
An explicit unit confirmation and fresh cutover checks remain separate approvals.
"""
import argparse
from collections import Counter, defaultdict
import csv
from datetime import datetime, timezone
import hashlib
import html
import importlib.util
from io import StringIO
import json
import os
from pathlib import Path
import re

_spec = importlib.util.spec_from_file_location('catalogue_review', Path(__file__).with_name('review-ecwid-catalogue.py'))
catalogue = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(catalogue)
ROOT_IDS = ['40854040', '40854051', '12115348', '48083853', '40865024', '159132005', '159154502']
FASTENERS_NAME = 'Fasteners & Spacers'
TERMINAL = {'READY_FOR_PICKUP', 'SHIPPED', 'DELIVERED', 'OUT_FOR_DELIVERY', 'RETURNED', 'WILL_NOT_DELIVER'}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def norm(value):
    return value.strip().upper() if isinstance(value, str) else ''


def whole(value):
    return type(value) is int and 0 <= value <= 2_147_483_647


def metadata(data, filename):
    return {'file_name': Path(filename).name, 'sha256': sha(data), 'bytes': len(data)}


def parse_scope_csv(data):
    if len(data) > 30_000_000:
        raise ValueError('CSV is too large.')
    reader = csv.DictReader(StringIO(data.decode('utf-8-sig'), newline=''), strict=True)
    fields = reader.fieldnames
    required = {'type', 'category_internal_id', 'category_path', 'product_internal_id', 'product_category_1'}
    if not fields or len(fields) != len(set(fields)) or not required.issubset(fields):
        raise ValueError('CSV category columns are missing or duplicate.')
    category_fields = [key for key in fields if re.fullmatch(r'product_category_\d+', key)]
    categories, products = {}, {}
    for ordinal, row in enumerate(reader, 2):
        if ordinal > 100_001 or None in row or any(value is None for value in row.values()):
            raise ValueError('Invalid CSV record width or row count.')
        if row['type'] == 'category':
            identity, path = row['category_internal_id'].strip(), row['category_path'].strip()
            if not re.fullmatch(r'[1-9]\d{0,19}', identity) or not path or identity in categories:
                raise ValueError('CSV category identity is missing or duplicated.')
            categories[identity] = {'id': identity, 'path': path}
        elif row['type'] == 'product':
            identity = row['product_internal_id'].strip()
            if not re.fullmatch(r'[1-9]\d{0,19}', identity) or identity in products:
                raise ValueError('CSV parent identity is missing or duplicated.')
            products[identity] = {'id': identity, 'name': row.get('product_name', ''),
                                  'sku': row.get('product_sku', ''), 'csv_record': ordinal,
                                  'category_paths': [row[key].strip() for key in category_fields if row[key].strip()]}
    paths = [value['path'] for value in categories.values()]
    if len(paths) != len(set(paths)):
        raise ValueError('CSV category paths are ambiguous.')
    fasteners = [key for key, value in categories.items() if value['path'] == FASTENERS_NAME]
    if len(fasteners) != 1 or any(key not in categories for key in ROOT_IDS):
        raise ValueError('Requested pilot categories were not all found exactly once.')
    roots = ROOT_IDS + fasteners
    root_paths = {key: categories[key]['path'] for key in roots}
    def under(path, ancestor):
        return path == ancestor or path.startswith(ancestor + ' / ')
    groups = {key: {identity for identity, product in products.items()
                    if any(under(path, ancestor) for path in product['category_paths'])}
              for key, ancestor in root_paths.items()}
    selected = set().union(*groups.values())
    category_ids = {key for key, value in categories.items() if any(under(value['path'], path) for path in root_paths.values())}
    return {'categories': categories, 'products': products, 'root_ids': roots,
            'fasteners_id': fasteners[0], 'groups': groups, 'selected_ids': selected, 'category_ids': category_ids}


def parse_source(data):
    payload = json.loads(data)
    if (payload.get('kind') != 'SOURCE_REVIEW' or payload.get('dry_run') is not True or
            payload.get('balance_meaning') != 'PHYSICAL_ON_HAND' or not isinstance(payload.get('rows'), list) or
            not 0 < len(payload['rows']) <= 10_000 or payload.get('counts', {}).get('total') != len(payload['rows'])):
        raise ValueError('A complete full-source physical-stock review is required, not candidates only.')
    if not re.fullmatch('[a-f0-9]{64}', payload.get('source', {}).get('sha256', '')):
        raise ValueError('Source workbook hash is missing.')
    seen = set()
    for row in payload['rows']:
        key = (row.get('source_sheet'), row.get('source_row'))
        if (not isinstance(row.get('source_sheet'), str) or type(row.get('source_row')) is not int or
                not isinstance(row.get('sku'), str) or key in seen):
            raise ValueError('Source row identity is missing or duplicated.')
        seen.add(key)
    return payload


def verify_live_scope(payload, scope, store_id):
    if (payload.get('kind') != 'READONLY_CATEGORY_SCOPE' or payload.get('schema_version') != 1 or
            payload.get('complete') is not True or payload.get('dry_run') is not True or payload.get('store_id') != store_id or
            payload.get('hidden_categories_included') is not True):
        raise ValueError('Invalid live category snapshot envelope.')
    cats, products = payload.get('categories'), payload.get('products')
    if (not isinstance(cats, list) or not isinstance(products, list) or
            payload.get('category_count') != len(cats) or payload.get('product_count') != len(products)):
        raise ValueError('Incomplete live category snapshot.')
    by_id = {row['id']: row for row in cats}
    by_product = {row['id']: row for row in products}
    if len(by_id) != len(cats) or len(by_product) != len(products) or any(key not in by_id for key in scope['root_ids']):
        raise ValueError('Live category or product identity is missing or duplicated.')
    def descendants(identity, roots):
        seen = set()
        while identity != '0':
            if identity in seen or identity not in by_id:
                raise ValueError('Live category hierarchy is cyclic or incomplete.')
            seen.add(identity)
            if identity in roots:
                return True
            identity = by_id[identity]['parentId']
        return False
    selected_cats = {identity for identity in by_id if descendants(identity, set(scope['root_ids']))}
    fastener_cats = {identity for identity in by_id if descendants(identity, {scope['fasteners_id']})}
    selected_products = {row['id'] for row in products if set(row['categoryIds']) & selected_cats}
    fastener_products = {row['id'] for row in products if set(row['categoryIds']) & fastener_cats}
    differences = {'only_live_products': sorted(selected_products - scope['selected_ids']),
                   'only_csv_products': sorted(scope['selected_ids'] - selected_products),
                   'only_live_categories': sorted(selected_cats - scope['category_ids']),
                   'only_csv_categories': sorted(scope['category_ids'] - selected_cats)}
    return {'verified': not any(differences.values()), 'live_parent_count': len(selected_products),
            'live_fastener_parent_count': len(fastener_products), 'differences': differences,
            'started_at': payload.get('started_at'), 'completed_at': payload.get('completed_at')}


def reservation_review(order_bytes, confirmation_bytes, store_id):
    orders, confirmation = json.loads(order_bytes), json.loads(confirmation_bytes)
    if (orders.get('kind') != 'READONLY_ORDER_REVIEW' or orders.get('complete') is not True or
            orders.get('dry_run') is not True or orders.get('store_id') != store_id or
            orders.get('pending_order_count') != len(orders.get('orders', [])) or
            confirmation.get('kind') != 'SNAPSHOT_PICK_CONFIRMATION' or confirmation.get('store_id') != store_id or
            confirmation.get('order_snapshot', {}).get('sha256') != sha(order_bytes) or
            confirmation.get('outstanding_order_scope_confirmed') is not True):
        raise ValueError('Orders must match the exact snapshot-scoped physical confirmation.')
    physical = confirmation.get('physical_pick_confirmation', {})
    if physical.get('all_items_still_on_shelves') is not True or physical.get('previously_picked_quantity_per_line') != 0:
        raise ValueError('Physical zero-picked confirmation is required.')
    confirmed = {row['order_id']: row for row in confirmation.get('confirmed_shipment_orders', [])}
    pending = [row for row in orders['orders'] if row['fulfillment_status'] not in TERMINAL]
    if {row['id'] for row in pending} != set(confirmed) or len(pending) != len(confirmed):
        raise ValueError('Pending order scope differs from the physical confirmation.')
    lines = []
    for order in pending:
        approved = confirmed[order['id']]
        if (order['payment_status'] != 'PAID' or approved.get('unchanged') is not True or
                approved.get('line_count') != len(order['lines'])):
            raise ValueError('Order changed or lacks a supported physical confirmation.')
        for line in order['lines']:
            if not whole(line.get('ordered_quantity')) or line['ordered_quantity'] == 0:
                raise ValueError('Order quantity is invalid.')
            lines.append({**line, 'order_id': order['id']})
    if physical.get('confirmed_line_count') != len(lines):
        raise ValueError('Physical line confirmation count differs.')
    return {'lines': lines, 'pending_order_count': len(pending),
            'excluded_picked_packed_order_count': len(orders['orders']) - len(pending),
            'order_snapshot_completed_at': orders.get('completed_at')}


def create_review(scope, source, targets, reservations=None, scope_check=None):
    source_by_sku, target_by_sku = defaultdict(list), defaultdict(list)
    for row in source['rows']:
        if norm(row['sku']) not in {'', 'NA', 'N/A'}:
            source_by_sku[norm(row['sku'])].append(row)
    identities = Counter((target.get('id'), target.get('combinationId')) for target in targets)
    for target in targets:
        if norm(target.get('sku')):
            target_by_sku[norm(target['sku'])].append(target)
    output, missing_parents = [], []
    selected_targets = [target for target in targets if target.get('id') in scope['selected_ids']]
    for identity in sorted(scope['selected_ids'] - {target.get('id') for target in selected_targets}):
        missing_parents.append({'product_id': identity, 'name': scope['products'][identity]['name'], 'issue': 'Parent has no live independent stock target.'})
    for target in selected_targets:
        sku = norm(target.get('sku'))
        sources = source_by_sku[sku] if sku else []
        issues = []
        def issue(code, message):
            issues.append({'code': code, 'message': message})
        if not sku:
            issue('MISSING_SKU', 'Live independent target has no SKU.')
        elif len(target_by_sku[sku]) != 1:
            issue('DUPLICATE_LIVE_SKU', 'More than one live independent stock target uses this SKU.')
        if identities[(target.get('id'), target.get('combinationId'))] != 1:
            issue('DUPLICATE_LIVE_IDENTITY', 'Live product/variation identity is duplicated.')
        if not sources:
            issue('NO_SOURCE_ROW', 'No matching SKU in the full stock workbook. Missing does not mean zero.')
        elif len(sources) > 1:
            issue('DUPLICATE_SOURCE_SKU', 'Multiple workbook rows share this SKU. No balances were combined.')
        elif sources[0].get('status') != 'CANDIDATE':
            issue('SOURCE_REVIEW_REQUIRED', 'Workbook row is not a phase-one candidate: ' + ', '.join(reason['message'] for reason in sources[0].get('reasons', [])))
        if len(sources) == 1 and not whole(sources[0].get('balance')):
            issue('INVALID_SOURCE_BALANCE', 'Workbook balance is missing, negative, fractional or invalid.')
        for reason in catalogue.live_target_reasons(target):
            code = 'LIVE_ELIGIBILITY'
            if 'tracking' in reason or 'quantity' in reason:
                code = 'TRACKING_OR_QUANTITY'
            elif 'enabled' in reason:
                code = 'PRODUCT_DISABLED'
            elif 'bundles' in reason:
                code = 'BUNDLE_OR_EXTRA_OPTIONS'
            issue(code, reason)
        if scope_check is None or not scope_check['verified']:
            issue('CATEGORY_SCOPE_UNVERIFIED', 'Live category membership has not matched the CSV scope.')
        reserve, matched, unknown = None, [], []
        if reservations is not None:
            reserve = 0
            for line in reservations['lines']:
                identity_match = (line.get('ecwid_product_id') == target.get('id') and
                                  line.get('ecwid_combination_id') == target.get('combinationId'))
                # Unknown/deleted combination IDs affect the entire mapped
                # parent, not an invented zero reservation for every sibling.
                unresolved_parent = (line.get('ecwid_product_id') == target.get('id') and
                                     identities[(line.get('ecwid_product_id'), line.get('ecwid_combination_id'))] != 1)
                if not identity_match and norm(line.get('sku')) != sku and not unresolved_parent:
                    continue
                exact = (identity_match and norm(line.get('sku')) == sku and
                         line.get('selected_options_supported') is True and
                         catalogue.canonical_options(line.get('selected_options')) is not None and
                         catalogue.canonical_options(line.get('selected_options')) == catalogue.canonical_options(target.get('variationOptions')))
                if exact:
                    reserve += line['ordered_quantity']
                    matched.append({'order_id': line['order_id'], 'ecwid_line_id': line['ecwid_line_id'], 'quantity': line['ordered_quantity']})
                else:
                    unknown.append({'order_id': line['order_id'], 'ecwid_line_id': line['ecwid_line_id']})
            if unknown:
                reserve = None
                issue('ORDER_IDENTITY_REVIEW', 'Pending order identity, SKU or options disagree. No reservation quantity was inferred.')
        else:
            issue('RESERVATION_REVIEW_REQUIRED', 'Exact snapshot-scoped physical order confirmation is missing.')
        physical = sources[0].get('balance') if len(sources) == 1 else None
        if reserve is not None and whole(physical) and physical < reserve:
            issue('INSUFFICIENT_PHYSICAL_BALANCE', 'Physical balance is lower than pending order units. Recount or resolve unit mapping.')
        technical_match = not issues
        provisional = physical - reserve if technical_match and reserve is not None else None
        issue('SINGLE_UNIT_CONFIRMATION_REQUIRED', 'Confirm one workbook piece equals one Ecwid sale unit. Category selection is not a unit-conversion approval.')
        output.append({'sku': sku, 'name': target.get('name'), 'ecwid_product_id': target.get('id'),
                       'ecwid_combination_id': target.get('combinationId'), 'variation_options': target.get('variationOptions'),
                       'category_paths': scope['products'][target['id']]['category_paths'],
                       'ecwid_quantity': target.get('quantity'), 'source_rows': [{key: row.get(key) for key in
                       ('source_row', 'source_sheet', 'sku', 'name', 'unit', 'location', 'balance', 'status')} for row in sources],
                       'physical_balance': physical, 'reserved_ecwid_units': reserve, 'reservation_lines': matched,
                       'unresolved_order_lines': unknown, 'technical_match': technical_match,
                       'provisional_quantity_if_one_to_one': provisional, 'desired_ecwid_quantity': None,
                       'single_unit_confirmed': False, 'ready': False, 'status': 'UNIT_CONFIRMATION' if technical_match else 'REVIEW',
                       'issues': issues})
    output.sort(key=lambda row: (row['sku'], row['ecwid_product_id'], row['ecwid_combination_id'] or ''))
    counts = Counter(code for row in output for code in {issue['code'] for issue in row['issues']})
    return {'kind': 'PILOT_SCOPE_REVIEW', 'schema_version': 1, 'dry_run': True, 'ready_count': 0,
            'generated_at': datetime.now(timezone.utc).isoformat(), 'store_id': '2442119',
            'current_authority': 'WORKBOOK', 'stock_writes_performed': False, 'database_imports_performed': False,
            'single_unit_mappings_approved': False, 'cutover_approved': False,
            'scope': {'root_categories': [{'id': key, 'path': scope['categories'][key]['path'], 'parent_count': len(scope['groups'][key])} for key in scope['root_ids']],
                      'category_count': len(scope['category_ids']), 'parent_product_count': len(scope['selected_ids']),
                      'fastener_parent_count': len(scope['groups'][scope['fasteners_id']]), 'fastener_expected_parent_count': 86,
                      'stock_target_count': len(output), 'variation_target_count': sum(row['ecwid_combination_id'] is not None for row in output),
                      'simple_target_count': sum(row['ecwid_combination_id'] is None for row in output), 'live_crosscheck': scope_check},
            'counts': {'technical_matches_awaiting_unit_confirmation': sum(row['technical_match'] for row in output),
                       'review_required': sum(not row['technical_match'] for row in output), 'issue_rows': dict(sorted(counts.items()))},
            'order_review': {key: value for key, value in (reservations or {}).items() if key != 'lines'},
            'source': source['source'], 'missing_live_parents': missing_parents, 'rows': output,
            'note': 'Review only. Provisional quantities assume one physical piece per Ecwid unit, are not approved import/write quantities, and require fresh stock/order checks at cutover. Ready for Pickup means picked and packed, equivalent to Shipped.'}


def render_report(report):
    esc = lambda value: html.escape(str(value), quote=True)
    display = lambda value: 'Not available' if value is None else esc(value)
    rows = []
    for row in report['rows']:
        source = '<br>'.join(f"{esc(s['source_sheet'])} row {esc(s['source_row'])}: {esc(s['balance'])} {esc(s['unit'])}, {esc(s['location'])}" for s in row['source_rows']) or 'No workbook SKU match'
        options = ', '.join(f"{x['name']}: {x['value']}" for x in row['variation_options'] or [])
        rows.append(f'<tr data-status="{row["status"]}"><td>{esc(row["sku"] or "Missing SKU")}<small>{esc(row["name"])}<br>{esc(options)}</small></td><td>{esc(row["ecwid_product_id"])}<small>Variation: {esc(row["ecwid_combination_id"] or "None")}</small></td><td>{source}</td><td>{display(row["ecwid_quantity"])}</td><td>{display(row["reserved_ecwid_units"])}</td><td>{display(row["provisional_quantity_if_one_to_one"])}</td><td>{"<br>".join(esc(x["message"]) for x in row["issues"])}</td></tr>')
    scope = report['scope']
    groups = ''.join(f'<li>{esc(row["path"].replace(chr(92)+"/", "/"))} ({esc(row["id"])}): {row["parent_count"]} products</li>' for row in scope['root_categories'])
    expected = 'matching the expected 86' if scope['fastener_parent_count'] == 86 else 'different from the expected 86; confirm scope'
    category_verified = bool(scope.get('live_crosscheck') and scope['live_crosscheck']['verified'])
    live_scope_note = 'Live category membership matches the exported catalogue, including hidden categories.' if category_verified else 'Live category membership is unverified or differs. Resolve category scope before proceeding.'
    return f'''<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Pilot stock review</title>
<style>body{{font:15px system-ui,sans-serif;color:#182638;margin:28px auto;max-width:1600px;padding:0 20px}}h1{{font-size:28px}}p{{max-width:1000px;line-height:1.5}}.notice{{background:#fff4d7;padding:16px;border-left:4px solid #ba8500}}small{{display:block;color:#526171;margin-top:5px}}table{{border-collapse:collapse;width:100%;font-size:13px}}th{{background:#253d59;color:white;text-align:left;position:sticky;top:0}}th,td{{padding:12px;vertical-align:top;border-bottom:1px solid #d8e0e7}}tr:nth-child(even){{background:#f6f8fa}}input,select{{font:inherit;padding:10px;margin:16px 10px 16px 0}}td:first-child{{min-width:180px}}td:last-child{{min-width:310px}}.table{{overflow:auto;max-height:70vh}}details{{margin:14px 0}}</style>
<h1>Pilot stock review</h1><p>{scope['parent_product_count']} products, {scope['stock_target_count']} independent stock targets ({scope['variation_target_count']} variations and {scope['simple_target_count']} simple items). Fasteners &amp; Spacers: {scope['fastener_parent_count']} products, {expected}.</p><p>{live_scope_note}</p>
<p class="notice">No stock has been imported or changed in Ecwid. {report['counts']['technical_matches_awaiting_unit_confirmation']} targets pass the current technical and source checks but still need one-piece unit confirmation. Other targets need the checks listed below. Blank or missing balances are never treated as zero.</p>
<p>Physical balances come from the full stock workbook. Pending units use the exact confirmed five-order snapshot. Ready for Pickup means picked and packed, equivalent to Shipped. The provisional column is physical balance minus reserved order units <strong>only if the units are one-to-one</strong>. It is not an approved write quantity. Refresh both stock and orders before cutover.</p>
<details><summary>Selected categories and product counts (overlaps counted once in the total)</summary><ul>{groups}</ul></details>
<label>Find <input id="search" placeholder="SKU, product, location or issue"></label><label>Status <select id="status"><option value="">All</option><option value="UNIT_CONFIRMATION">Only unit confirmation remains</option><option value="REVIEW">Other review needed</option></select></label><span id="count"></span>
<div class="table"><table><thead><tr><th>SKU and item</th><th>Ecwid identity</th><th>Workbook source</th><th>Current Ecwid quantity</th><th>Reserved order units</th><th>Provisional if one-to-one</th><th>Required checks</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>
<script>const search=document.querySelector('#search'),status=document.querySelector('#status'),rows=[...document.querySelectorAll('tbody tr')];function filter(){{let n=0;for(const row of rows){{row.hidden=!(row.textContent.toLowerCase().includes(search.value.toLowerCase())&&(!status.value||row.dataset.status===status.value));if(!row.hidden)n++}}document.querySelector('#count').textContent=n+' of '+rows.length+' targets'}}search.addEventListener('input',filter);status.addEventListener('change',filter);filter();</script></html>'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('csv')
    parser.add_argument('--source-review', required=True)
    parser.add_argument('--live-catalog', required=True)
    parser.add_argument('--live-scope')
    parser.add_argument('--orders')
    parser.add_argument('--pick-confirmation')
    parser.add_argument('--out', required=True)
    args = parser.parse_args()
    if bool(args.orders) != bool(args.pick_confirmation):
        parser.error('--orders and --pick-confirmation must be supplied together.')
    inputs = {key: (path, catalogue.bounded_bytes(path)) for key, path in vars(args).items() if key != 'out' and path}
    scope = parse_scope_csv(inputs['csv'][1])
    source = parse_source(inputs['source_review'][1])
    live = catalogue.parse_live_catalogue(inputs['live_catalog'][1], args.live_catalog, '2442119')
    scope_check = verify_live_scope(json.loads(inputs['live_scope'][1]), scope, '2442119') if args.live_scope else None
    reservations = reservation_review(inputs['orders'][1], inputs['pick_confirmation'][1], '2442119') if args.orders else None
    report = create_review(scope, source, live['stock_targets'], reservations, scope_check)
    report['inputs'] = {key: metadata(data, path) for key, (path, data) in inputs.items()}
    output = Path(args.out).resolve()
    private_root = Path('import-data').resolve()
    if output == private_root or private_root not in output.parents or output.exists():
        raise ValueError('Output must be a new directory under private import-data.')
    output.mkdir(mode=0o700)
    for name, content in [('pilot-review.json', json.dumps(report, ensure_ascii=False, indent=2) + '\n'), ('review.html', render_report(report))]:
        descriptor = os.open(output / name, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        with os.fdopen(descriptor, 'w', encoding='utf-8') as stream:
            stream.write(content)
    print(json.dumps({'scope': report['scope'], 'counts': report['counts'], 'ready_count': 0, 'output': str(output)}, indent=2))


if __name__ == '__main__':
    main()
