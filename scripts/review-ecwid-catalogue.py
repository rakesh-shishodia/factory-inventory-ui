"""Offline, read-only Ecwid CSV comparison. Never creates an inventory import.

Each source is read once into bounded immutable bytes. Product IDs join child
records; normalized SKU strings only match identities, never aggregate stock.
"""
import argparse
from collections import Counter, defaultdict
import csv
from datetime import datetime, timezone
from decimal import Decimal
import hashlib
import html
from io import StringIO
import json
import os
from pathlib import Path
import re
import sys

MAX_FILE = 30_000_000
MAX_ROWS = 100_000
MAX_QUANTITY = 2_147_483_647
FIELDS = ('type', 'product_internal_id', 'product_sku', 'product_name',
          'product_is_inventory_tracked', 'product_quantity', 'product_is_available',
          'product_option_name', 'product_option_type', 'product_variation_sku')
TYPES = {'product', 'product_option', 'product_variation', 'product_file', 'category'}
LABELS = {'SIMPLE_MATCH': 'Simple matches to verify',
          'OPTIONS_VARIATIONS': 'Options or variations',
          'VARIATION_MATCH': 'Independent variations to verify', 'MISSING': 'Not found',
          'AMBIGUOUS_SKU': 'Ambiguous SKU', 'REVIEW': 'Other review'}


def digest(data):
    return hashlib.sha256(data).hexdigest()


def bounded_bytes(path, maximum=MAX_FILE):
    with Path(path).open('rb') as stream:
        data = stream.read(maximum + 1)
    if len(data) > maximum:
        raise ValueError('Source exceeds the supported file size.')
    return data


def normalized(value):
    return value.strip().upper()


def boolean(value):
    return {'true': True, 'false': False}.get(value.strip().lower())


def quantity(value):
    value = value.strip()
    if len(value) > 40 or not re.fullmatch(r'[+-]?\d+(?:\.\d+)?', value):
        return None
    number = Decimal(value)
    if number < 0 or number > MAX_QUANTITY or number != number.to_integral_value():
        return None
    return int(number)


def parse_catalogue(data, filename):
    if len(data) > MAX_FILE:
        raise ValueError('Catalogue exceeds the supported file size.')
    try:
        text = data.decode('utf-8-sig')
    except UnicodeDecodeError as exc:
        raise ValueError('Catalogue must be UTF-8 CSV.') from exc
    if '\x00' in text:
        raise ValueError('Catalogue contains a NUL character.')
    previous_limit = csv.field_size_limit(1_000_000)
    products, children, counts = [], [], Counter()
    try:
        reader = csv.reader(StringIO(text, newline=''), strict=True)
        header = next(reader, None)
        if not header or len(header) > 1000 or len(set(header)) != len(header):
            raise ValueError('Catalogue header is missing, duplicate or too wide.')
        if not set(FIELDS).issubset(header):
            raise ValueError('Catalogue is missing required columns: ' + ', '.join(sorted(set(FIELDS) - set(header))))
        indices = {name: header.index(name) for name in FIELDS}
        option_indices = [(i, name[len('product_variation_option_'):]) for i, name in enumerate(header)
                          if name.startswith('product_variation_option_')]
        for ordinal, values in enumerate(reader, 2):
            if ordinal > MAX_ROWS + 1:
                raise ValueError('Catalogue has too many records.')
            if len(values) != len(header):
                raise ValueError(f'CSV record {ordinal} has a different number of columns.')
            row = {name: values[index] for name, index in indices.items()}
            kind = row['type'].strip()
            if kind not in TYPES:
                raise ValueError(f'CSV record {ordinal} has an unsupported record type.')
            counts[kind] += 1
            if kind == 'category':
                continue
            product_id = row['product_internal_id'].strip()
            if not re.fullmatch(r'[1-9]\d{0,19}', product_id):
                raise ValueError(f'CSV record {ordinal} has a missing or invalid product ID.')
            record = {'csv_record': ordinal, 'csv_end_line': reader.line_num,
                      'product_id': product_id, 'source_sku': row['product_sku'],
                      'sku': normalized(row['product_sku']), 'kind': kind}
            if kind == 'product':
                record.update(name=row['product_name'],
                              inventory_tracked=boolean(row['product_is_inventory_tracked']),
                              available=boolean(row['product_is_available']),
                              source_inventory_tracked=row['product_is_inventory_tracked'],
                              source_available=row['product_is_available'],
                              source_quantity=row['product_quantity'],
                              quantity=quantity(row['product_quantity']),
                              option_records=[], variation_records=[], file_records=[], issues=[])
                products.append(record)
            elif kind == 'product_variation':
                options = [{'name': name[1:-1] if name.startswith('{') and name.endswith('}') else name,
                            'value': values[index]} for index, name in option_indices if values[index]]
                record.update(variation_sku=normalized(row['product_variation_sku']),
                              source_variation_sku=row['product_variation_sku'],
                              inventory_tracked=boolean(row['product_is_inventory_tracked']),
                              source_quantity=row['product_quantity'],
                              quantity=quantity(row['product_quantity']), variation_options=options)
                children.append(record)
            else:
                record.update(option_name=row['product_option_name'], option_type=row['product_option_type'])
                children.append(record)
    except csv.Error as exc:
        raise ValueError('Malformed CSV: ' + str(exc)) from exc
    finally:
        csv.field_size_limit(previous_limit)
    if not products:
        raise ValueError('Catalogue has no product records.')
    by_id = defaultdict(list)
    for product in products:
        by_id[product['product_id']].append(product)
    for child in children:
        parents = by_id.get(child['product_id'], [])
        if not parents:
            raise ValueError(f"CSV record {child['csv_record']} is an orphan {child['kind']} record.")
        key = {'product_option': 'option_records', 'product_variation': 'variation_records', 'product_file': 'file_records'}[child['kind']]
        for parent in parents:
            parent[key].append(child)
            if child['sku'] != parent['sku']:
                parent['issues'].append(f"Child record {child['csv_record']} has a different parent SKU.")
    for product in products:
        if len(by_id[product['product_id']]) > 1:
            product['issues'].append('Product ID is duplicated in the export.')
        product['issues'] = list(dict.fromkeys(product['issues']))
    return {'source': {'file_name': Path(filename).name, 'sha256': digest(data),
                       'bytes': len(data), 'format': 'ECWID_CSV'},
            'record_counts': dict(sorted(counts.items())), 'products': products,
            'variations': [row for row in children if row['kind'] == 'product_variation']}


def parse_candidates(data, filename):
    if len(data) > 8_000_000:
        raise ValueError('Candidate file is too large.')
    try:
        payload = json.loads(data)
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError('Candidate file is not valid JSON.') from exc
    if not isinstance(payload, dict) or payload.get('kind') != 'SOURCE_CANDIDATES' or payload.get('dry_run') is not True:
        raise ValueError('Expected a staged SOURCE_CANDIDATES file.')
    if (payload.get('balance_meaning') != 'PHYSICAL_ON_HAND' or
            payload.get('ecwid_checked') is not False or payload.get('reservations_confirmed') is not False):
        raise ValueError('Candidate file must contain unconfirmed physical balances.')
    if not isinstance(payload.get('source_hash'), str) or not re.fullmatch('[a-f0-9]{64}', payload['source_hash']):
        raise ValueError('Candidate file is missing its workbook hash.')
    if not isinstance(payload.get('source_ref'), str) or not payload['source_ref'].strip() or len(payload['source_ref']) > 2000:
        raise ValueError('Candidate file is missing its source reference.')
    rows = payload.get('rows')
    if not isinstance(rows, list) or not 0 < len(rows) <= 10_000:
        raise ValueError('Candidate rows must be a nonempty bounded array.')
    seen, seen_source = set(), set()
    for row in rows:
        if not isinstance(row, dict) or not isinstance(row.get('sku'), str) or not normalized(row['sku']):
            raise ValueError('Candidate SKU must be nonempty text.')
        sku = normalized(row['sku'])
        if sku in seen:
            raise ValueError('Candidate SKU is duplicated; review the workbook first.')
        seen.add(sku)
        if type(row.get('balance')) is not int or not 0 <= row['balance'] <= MAX_QUANTITY:
            raise ValueError('Candidate balance must be a nonnegative whole number.')
        if type(row.get('source_row')) is not int or not 1 <= row['source_row'] <= 1_048_576:
            raise ValueError('Candidate source row is invalid.')
        if not isinstance(row.get('source_sheet'), str) or not row['source_sheet'].strip():
            raise ValueError('Candidate source sheet is missing.')
        source_key = (row['source_sheet'], row['source_row'])
        if source_key in seen_source:
            raise ValueError('Candidate source row is duplicated.')
        seen_source.add(source_key)
        if any(not isinstance(row.get(key), str) for key in ('name', 'location', 'scan_code')):
            raise ValueError('Candidate name, location and scan code must be text.')
    return {'source': {'file_name': Path(filename).name, 'sha256': digest(data),
                       'workbook_sha256': payload['source_hash'], 'source_ref': payload['source_ref']},
            'rows': rows}


def product_summary(product):
    keys = ('product_id', 'sku', 'source_sku', 'name', 'csv_record', 'csv_end_line',
            'inventory_tracked', 'available', 'source_quantity', 'quantity')
    return {**{key: product[key] for key in keys},
            'option_record_count': len(product['option_records']),
            'variation_record_count': len(product['variation_records']),
            'file_record_count': len(product['file_records']),
            'issues': product['issues']}


def canonical_options(value):
    if not isinstance(value, list) or len(value) > 100:
        return None
    names, result = set(), []
    for option in value:
        if (not isinstance(option, dict) or not isinstance(option.get('name'), str) or
                not option['name'].strip() or len(option['name']) > 200 or
                not isinstance(option.get('value'), str) or not option['value'].strip() or
                len(option['value']) > 2000 or option['name'] in names):
            return None
        names.add(option['name'])
        result.append({'name': option['name'], 'value': option['value']})
    return sorted(result, key=lambda option: option['name'])


def parse_live_catalogue(data, filename, store_id):
    if len(data) > MAX_FILE:
        raise ValueError('Live catalogue exceeds the supported file size.')
    try:
        payload = json.loads(data)
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError('Live catalogue is not valid JSON.') from exc
    if (not isinstance(payload, dict) or payload.get('kind') != 'READONLY_CATALOGUE' or
            payload.get('schema_version') != 1 or payload.get('dry_run') is not True or
            payload.get('complete') is not True or payload.get('store_id') != store_id):
        raise ValueError('Expected a complete read-only catalogue snapshot for this store.')
    targets = payload.get('stock_targets')
    if (not isinstance(targets, list) or len(targets) > 20_000 or
            type(payload.get('stock_target_count')) is not int or
            payload.get('stock_target_count') != len(targets) or any(not isinstance(row, dict) for row in targets)):
        raise ValueError('Live stock targets are incomplete or invalid.')
    for key in ('started_at', 'completed_at'):
        value = payload.get(key)
        if not isinstance(value, str) or len(value) > 40:
            raise ValueError('Live snapshot timestamps are missing or invalid.')
        try:
            datetime.fromisoformat(value.replace('Z', '+00:00'))
        except ValueError as exc:
            raise ValueError('Live snapshot timestamps are missing or invalid.') from exc
    return {'source': {'file_name': Path(filename).name, 'sha256': digest(data), 'format': 'ECWID_API_SNAPSHOT',
                       'started_at': payload['started_at'], 'completed_at': payload['completed_at']},
            'stock_targets': targets}


def live_target_reasons(target):
    reasons = []
    if not isinstance(target.get('id'), str) or not re.fullmatch(r'[1-9]\d{0,19}', target['id']):
        reasons.append('Live product ID is missing or invalid.')
    combo = target.get('combinationId')
    if 'combinationId' not in target or (combo is not None and (not isinstance(combo, str) or not re.fullmatch(r'[1-9]\d{0,19}', combo))):
        reasons.append('Live variation identity is missing or invalid.')
    if target.get('unlimited') is not False:
        reasons.append('Live target does not have independent inventory tracking confirmed.')
    qty = target.get('quantity')
    if type(qty) is not int or not 0 <= qty <= MAX_QUANTITY:
        reasons.append('Live target quantity is missing, negative, fractional or invalid.')
    if target.get('enabled') is not True:
        reasons.append('Live parent product is not confirmed enabled.')
    if (target.get('eligibilityVerified') is not True or target.get('hasBundleRelationships') is not False or
            target.get('hasExtraOptions') is not False):
        reasons.append('Live eligibility has not confirmed the absence of bundles and extra options.')
    options = canonical_options(target.get('variationOptions'))
    if combo is not None:
        if not options or target.get('hasOptions') is not True or target.get('hasVariations') is not False:
            reasons.append('Live variation does not have a verified exact option selection.')
    elif options != [] or target.get('hasOptions') is not False or target.get('hasVariations') is not False:
        reasons.append('Live base product has options or variations; select an independent variation.')
    return reasons


def compare(catalogue, candidates, store_id, live=None):
    if not re.fullmatch(r'[1-9]\d{0,19}', store_id):
        raise ValueError('Store ID must be a positive numeric identifier.')
    products_by_sku, variations_by_sku, parents_by_id = defaultdict(list), defaultdict(list), defaultdict(list)
    for product in catalogue['products']:
        parents_by_id[product['product_id']].append(product)
        if product['sku']:
            products_by_sku[product['sku']].append(product)
    for variation in catalogue['variations']:
        if variation['variation_sku']:
            variations_by_sku[variation['variation_sku']].append(variation)
    live_by_sku, live_ids = defaultdict(list), Counter()
    for target in live['stock_targets'] if live else []:
        if isinstance(target.get('sku'), str) and normalized(target['sku']):
            live_by_sku[normalized(target['sku'])].append(target)
        live_ids[(str(target.get('id')), str(target.get('combinationId')))] += 1
    rows = []
    for candidate in candidates['rows']:
        sku = normalized(candidate['sku'])
        products, variations = products_by_sku[sku], variations_by_sku[sku]
        stocked_products = [product for product in products if not product['variation_records']]
        reasons = []
        if len(stocked_products) + len(variations) > 1:
            status = 'AMBIGUOUS_SKU'
            reasons.append(f'SKU appears on {len(products)} product record(s) and {len(variations)} variation record(s). No identity was selected.')
        elif variations:
            status = 'VARIATION_MATCH'
            variation = variations[0]
            parents = parents_by_id[variation['product_id']]
            if len(parents) != 1:
                reasons.append('Variation parent identity is duplicated in the export.')
            else:
                parent = parents[0]
                reasons.extend(parent['issues'])
                if parent['available'] is not True:
                    reasons.append('Variation parent is disabled or availability is unconfirmed.')
                if parent['file_records']:
                    reasons.append('Variation parent has downloadable files; physical eligibility needs review.')
            if variation['inventory_tracked'] is not True:
                reasons.append('Variation does not have independent inventory tracking confirmed in this export.')
            if variation['quantity'] is None:
                reasons.append('Variation quantity is missing, negative, fractional or invalid.')
            if not canonical_options(variation['variation_options']):
                reasons.append('Variation option selection is missing or invalid in the export.')
            if reasons:
                status = 'REVIEW'
            else:
                reasons.append('Unique independently tracked variation. API variation ID, eligibility and one physical piece per Ecwid unit must be confirmed before approval.')
        elif not products:
            status = 'MISSING'
            reasons.append('No product or variation SKU matches in this export. Do not create an automatic mapping.')
        else:
            product = products[0]
            if product['option_records'] or product['variation_records']:
                status = 'OPTIONS_VARIATIONS'
                reasons.append('This parent has options or variations. Map an independently stocked variation SKU rather than treating its parent as a simple stock item.')
            else:
                status = 'SIMPLE_MATCH'
            reasons.extend(product['issues'])
            if product['inventory_tracked'] is not True:
                reasons.append('Inventory tracking is disabled.' if product['inventory_tracked'] is False else 'Inventory tracking flag is missing or invalid.')
            if product['available'] is not True:
                reasons.append('Product is disabled in the export.' if product['available'] is False else 'Product availability flag is missing or invalid.')
            if product['quantity'] is None:
                reasons.append('Export quantity is missing, invalid, negative, fractional or outside the supported range.')
            if product['file_records']:
                reasons.append('Product has downloadable-file records. Confirm physical-item eligibility.')
            if not product['name'].strip():
                reasons.append('Product name is missing.')
            if status == 'SIMPLE_MATCH' and reasons:
                status = 'REVIEW'
            if not reasons:
                reasons.append('Unique enabled, tracked product with whole-unit quantity and no option or variation records in this export. Confirm the physical identity and live product before approval.')
        live_matches = live_by_sku[sku] if live else []
        if live:
            if not live_matches:
                status, reasons = 'MISSING', ['No independent stock target matches this SKU in the complete live snapshot.']
            elif len(live_matches) > 1:
                status, reasons = 'AMBIGUOUS_SKU', ['Multiple independent live stock targets share this SKU; no mapping was selected.']
            else:
                target = live_matches[0]
                reasons = live_target_reasons(target)
                if live_ids[(str(target.get('id')), str(target.get('combinationId')))] > 1:
                    reasons.append('Live product/variation identity is duplicated in the snapshot.')
                # A same-SKU identity moved since export is material, even when
                # today's target is unique. Keep both sources for human review.
                exported_ids = {row['product_id'] for row in (variations or stocked_products)}
                if exported_ids and target.get('id') not in exported_ids:
                    reasons.append('Live parent identity disagrees with the CSV export.')
                if (len(variations) == 1 and canonical_options(variations[0]['variation_options']) is not None and
                        canonical_options(variations[0]['variation_options']) != canonical_options(target.get('variationOptions'))):
                    reasons.append('Live variation selection disagrees with the CSV export.')
                status = 'REVIEW' if reasons else ('VARIATION_MATCH' if target.get('combinationId') else 'SIMPLE_MATCH')
                if not reasons:
                    reasons.append('Live inventory target and options verified. Confirm one physical piece per Ecwid unit and reconcile all outstanding orders before approval.')
        rows.append({'sku': sku, 'source_sku': candidate['sku'], 'name': candidate['name'],
                     'location': candidate['location'], 'source_sheet': candidate['source_sheet'],
                     'source_row': candidate['source_row'], 'physical_balance': candidate['balance'],
                     'status': status, 'current_authority': 'WORKBOOK', 'ready': False,
                     'reasons': reasons, 'product_matches': [product_summary(product) for product in products],
                     'variation_matches': [{key: variation[key] for key in ('product_id', 'variation_sku', 'source_variation_sku', 'csv_record', 'csv_end_line', 'inventory_tracked', 'source_quantity', 'quantity', 'variation_options')}
                                           for variation in variations],
                     'live_matches': live_matches,
                     'single_unit_confirmed': False})
    counts = Counter(row['status'] for row in rows)
    return {'kind': 'CATALOGUE_REVIEW', 'dry_run': True,
            'generated_at': datetime.now(timezone.utc).isoformat(),
            'store_id': store_id, 'store_id_verified': live is not None,
            'catalogue_source': catalogue['source'], 'stock_source': candidates['source'],
            'record_counts': catalogue['record_counts'], 'candidate_count': len(rows),
            'counts': {key: counts[key] for key in LABELS}, 'ready_count': 0,
            'live_api_verified': live is not None, 'reservations_confirmed': False,
            'live_catalogue_source': live['source'] if live else None,
            'bundle_relationships_verified': False, 'balance_meaning': 'PHYSICAL_ON_HAND',
            'limitations': [('The live snapshot was fetched for the stated store; CSV identity remains separately recorded.' if live else 'The store ID is user-supplied and has not been verified against this CSV.'),
                            'A simple match is preliminary, not approval to import or change stock.',
                            'The export does not establish bundle or component relationships.',
                            'Open-order reservations and per-SKU physical unit mapping remain unconfirmed. A unique SKU does not establish that a store unit is one piece.',
                            'Physical workbook balances and exported Ecwid quantities have different meanings. No stock target or adjustment is calculated.',
                            'The workbook remains the authority for every item until its cutover is approved.'],
            'rows': rows}


def render_report(report):
    esc = lambda value: html.escape(str(value), quote=True)
    table_rows = []
    for row in report['rows']:
        matches = row['product_matches']
        evidence = [f"{p['name']} (ID {p['product_id']}, CSV record {p['csv_record']})" for p in matches]
        evidence += [f"Variation {v['variation_sku']} (product {v['product_id']}, CSV record {v['csv_record']})" for v in row['variation_matches']]
        quantities = [f"{p['source_quantity'] or 'Missing'} (tracked: {p['inventory_tracked']}, enabled: {p['available']})" for p in matches]
        quantities += [f"Variation: {v['source_quantity'] or 'Missing'} (tracked: {v['inventory_tracked']})" for v in row['variation_matches']]
        for target in row.get('live_matches', []):
            evidence.append(f"API product {target.get('id')}, variation {target.get('combinationId') or 'none'}")
            options = canonical_options(target.get('variationOptions'))
            if options:
                evidence.append(', '.join(f"{option['name']}: {option['value']}" for option in options))
            quantities.append(f"API: {target.get('quantity')} (tracked: {target.get('unlimited') is False}, enabled: {target.get('enabled')})")
        search = ' '.join([row['sku'], row['name'], row['location'], *evidence, *row['reasons']]).casefold()
        table_rows.append(f'<tr data-status="{esc(row["status"])}" data-search="{esc(search)}">'
                          f'<td><strong>{esc(row["sku"])}</strong><br>{esc(row["name"])}<small>{esc(row["location"])}</small></td>'
                          f'<td>{esc(row["source_sheet"])} row {row["source_row"]}</td>'
                          f'<td class="num">{row["physical_balance"]}</td>'
                          f'<td>{"<br>".join(map(esc, evidence)) or "No match"}</td>'
                          f'<td>{"<br>".join(map(esc, quantities)) or "Not applicable"}</td>'
                          f'<td><strong>{esc(LABELS[row["status"]])}</strong><ul>{"".join("<li>"+esc(reason)+"</li>" for reason in row["reasons"])}</ul></td></tr>')
    filters = ''.join(f'<option value="{key}">{esc(label)} ({report["counts"][key]})</option>' for key, label in LABELS.items())
    summary = ''.join(f'<div><strong>{report["counts"][key]}</strong><span>{esc(label)}</span></div>' for key, label in LABELS.items())
    live_source = report.get('live_catalogue_source')
    live_details = (f"<p>Live catalogue: {esc(live_source['file_name'])}<br><code>SHA-256 {esc(live_source['sha256'])}</code><br>"
                    f"Fetched {esc(live_source['started_at'])} to {esc(live_source['completed_at'])}</p>") if live_source else ''
    verification = 'has a complete read-only API snapshot' if report['live_api_verified'] else 'is not yet API-verified'
    return '''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; base-uri 'none'; form-action 'none'; connect-src 'none'">
<title>Ecwid catalogue review</title><style>
*{box-sizing:border-box}body{font:15px/1.5 system-ui,sans-serif;color:#182333;background:#f5f7fa;margin:0}main{max-width:1500px;margin:auto;padding:28px}h1{font-size:28px;margin:0 0 8px}h2{font-size:18px}p{margin:8px 0}.notice{border-left:4px solid #b7770c;background:#fff4df;padding:14px 18px;margin:20px 0}.summary{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin:22px 0}.summary div{background:white;border:1px solid #dce2e9;border-radius:8px;padding:12px}.summary strong{display:block;font-size:26px}.summary span{font-size:13px}.controls{display:flex;flex-wrap:wrap;gap:16px;align-items:end;margin:20px 0}label{display:flex;flex-direction:column;gap:6px}input,select{font:inherit;padding:9px;border:1px solid #9caabc;border-radius:5px}input{width:320px;max-width:100%}.table-wrap{overflow:auto;border:1px solid #dce2e9;border-radius:8px;background:white}table{border-collapse:collapse;width:100%;min-width:1200px}th,td{text-align:left;vertical-align:top;padding:13px;border-bottom:1px solid #e4e8ef}th{background:#21344b;color:white;position:sticky;top:0}td:first-child{min-width:230px}td:last-child{min-width:300px}small{display:block;color:#58667a}.num{text-align:right}ul{margin:7px 0;padding-left:18px}details{margin:22px 0;overflow-wrap:anywhere}code{font-size:12px}#empty{padding:24px}tr[hidden],#empty[hidden]{display:none}@media(max-width:900px){main{padding:18px}.summary{grid-template-columns:repeat(3,1fr)}}@media(max-width:500px){.summary{grid-template-columns:repeat(2,1fr)}input{width:100%}label{width:100%}}
</style></head><body><main><h1>Ecwid catalogue review</h1>''' + f'''
<p>{report['candidate_count']} workbook candidates compared with {report['record_counts'].get('product', 0)} exported products. Store {esc(report['store_id'])} {verification}.</p>
<div class="notice"><strong>No stock changes. No items ready for import.</strong><p>All items still use the workbook as their stock authority. Simple products and independent variations need physical-identity approval, one-piece-per-store-unit confirmation and open-order reservation reconciliation. CSV-only matches also need API eligibility checks.</p><p>The workbook balance is physical stock. Ecwid quantities are shown for reference only, not as a target or a directly comparable balance.</p></div>
<section class="summary" aria-label="Match counts">{summary}</section>
<div class="controls"><label>Search SKU, name, location or reason<input id="search" type="search" placeholder="For example, 01114-1"></label><label>Match category<select id="filter"><option value="ALL">All candidates ({report['candidate_count']})</option>{filters}</select></label><span id="shown" role="status"></span></div>
<div class="table-wrap"><table><thead><tr><th>Workbook item</th><th>Workbook reference</th><th>Physical balance</th><th>Catalogue identity</th><th>Export quantity</th><th>Review result</th></tr></thead><tbody>{''.join(table_rows)}</tbody></table><p id="empty" hidden>No candidates match these filters.</p></div>
<details><summary>Source files and review limits</summary><p>Catalogue: {esc(report['catalogue_source']['file_name'])}<br><code>SHA-256 {esc(report['catalogue_source']['sha256'])}</code></p>{live_details}<p>Workbook snapshot hash:<br><code>{esc(report['stock_source']['workbook_sha256'])}</code></p><p>Candidate file: {esc(report['stock_source']['file_name'])}<br><code>SHA-256 {esc(report['stock_source']['sha256'])}</code></p><p>Source reference: {esc(report['stock_source']['source_ref'])}</p><p>CSV record numbers include the header as record 1. Quoted multiline text can make physical line numbers differ. Matching trims and uppercases SKU text but preserves leading zeros.</p><ul>{''.join('<li>'+esc(note)+'</li>' for note in report['limitations'])}</ul></details>
''' + '''</main><script>
const search=document.getElementById('search'),filter=document.getElementById('filter'),rows=[...document.querySelectorAll('tbody tr')];
function update(){const query=search.value.trim().toLowerCase();let count=0;for(const row of rows){const visible=(filter.value==='ALL'||row.dataset.status===filter.value)&&row.dataset.search.toLowerCase().includes(query);row.hidden=!visible;if(visible)count++;}document.getElementById('shown').textContent=count+' candidates shown';document.getElementById('empty').hidden=count!==0;}
search.addEventListener('input',update);filter.addEventListener('change',update);update();
</script></body></html>'''


def write_report(report, output):
    output = Path(output).absolute()
    public = Path(__file__).resolve().parents[1] / 'public'
    resolved = output.resolve()
    if resolved == public or public in resolved.parents:
        raise ValueError('Private review reports cannot be written into public assets.')
    if not output.parent.is_dir():
        raise ValueError('The output parent directory must already exist.')
    output.mkdir(mode=0o700)  # Existing destinations are never overwritten.
    for name, content in [('catalogue-review.json', json.dumps(report, ensure_ascii=False, indent=2) + '\n'),
                          ('review.html', render_report(report))]:
        descriptor = os.open(output / name, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(descriptor, 'w', encoding='utf-8') as stream:
            stream.write(content)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('catalogue', type=Path)
    parser.add_argument('--candidates', type=Path, required=True)
    parser.add_argument('--store-id', required=True)
    parser.add_argument('--live-catalog', type=Path, help='Optional complete READONLY_CATALOGUE snapshot for current API identities and eligibility.')
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        catalogue = parse_catalogue(bounded_bytes(args.catalogue), args.catalogue.name)
        candidates = parse_candidates(bounded_bytes(args.candidates, 8_000_000), args.candidates.name)
        live = parse_live_catalogue(bounded_bytes(args.live_catalog), args.live_catalog.name, args.store_id) if args.live_catalog else None
        report = compare(catalogue, candidates, args.store_id, live)
        write_report(report, args.out)
        print(json.dumps({'kind': report['kind'], 'counts': report['counts'], 'ready_count': 0,
                          'record_counts': report['record_counts'], 'output': str(args.out.resolve())}))
        return 0
    except (ValueError, OSError) as exc:
        print('Catalogue review failed: ' + str(exc), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
