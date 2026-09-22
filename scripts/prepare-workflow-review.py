"""Record explicit business clarification in a private design/status review.

Reads frozen proposal/catalogue/orders and a user-authored clarification record.
Writes report files only; never imports inventory, edits sources or calls an API.
"""
import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import html
import json
import os
from pathlib import Path


def sha(data):
    return hashlib.sha256(data).hexdigest()


def read(path):
    with Path(path).open('rb') as stream:
        data = stream.read(30_000_001)
    if len(data) > 30_000_000:
        raise ValueError('Input exceeds supported size.')
    return data


def create_report(clarification_bytes, proposal_bytes, catalogue_bytes, orders_bytes):
    clarification, proposal, catalogue, orders = [json.loads(data) for data in
        (clarification_bytes, proposal_bytes, catalogue_bytes, orders_bytes)]
    for key, data in [('approval_proposal_sha256', proposal_bytes), ('catalogue_sha256', catalogue_bytes), ('orders_sha256', orders_bytes)]:
        if clarification.get(key) != sha(data):
            raise ValueError(f'{key} does not match the clarified snapshot.')
    if (clarification.get('kind') != 'PILOT_BUSINESS_CLARIFICATION' or clarification.get('schema_version') != 1 or
            any(value.get('store_id') != '2442119' for value in (clarification, proposal, catalogue, orders)) or
            proposal.get('kind') != 'PILOT_APPROVAL_PROPOSAL' or proposal.get('dry_run') is not True or
            catalogue.get('kind') != 'READONLY_CATALOGUE' or catalogue.get('complete') is not True or
            orders.get('kind') != 'READONLY_ORDER_REVIEW' or orders.get('complete') is not True or
            clarification.get('review_sha256') != proposal.get('review_sha256') or
            clarification.get('catalogue_sha256') != proposal.get('inputs', {}).get('live_catalog', {}).get('sha256') or
            clarification.get('orders_sha256') != proposal.get('inputs', {}).get('orders', {}).get('sha256') or
            clarification.get('supersedes_all_on_shelves_confirmation_sha256') != proposal.get('inputs', {}).get('pick_confirmation', {}).get('sha256')):
        raise ValueError('Clarification provenance does not match the complete store snapshots.')
    for key in ('opening_balances_approved', 'database_import_approved', 'cutover_approved',
                'ecwid_stock_write_authorized', 'ecwid_policy_change_authorized'):
        if clarification.get(key) is not False:
            raise ValueError('This review cannot grant stock, import or policy-change approval.')
    if any(proposal.get(key) is not False for key in ('batch_approved', 'cutover_approved', 'ecwid_write_authorized',
                                                     'stock_writes_performed', 'database_imports_performed')):
        raise ValueError('Expected an unapproved read-only proposal.')
    units = clarification.get('unit_confirmation', {})
    supplier = clarification.get('supplier_workflow', {})
    if (units.get('confirmed') is not True or type(units.get('ratio')) is not int or units['ratio'] != 1 or
            supplier.get('goods_always_arrive_at_factory') is not True or
            supplier.get('current_pending_supplier_lines_awaiting_receipt') is not True or
            supplier.get('workbook_balances_are_complete') is not False or
            supplier.get('opening_physical_balances') is not None or
            supplier.get('identification') != 'EXPLICIT_SKU_MAPPING' or
            supplier.get('attribute_identification_deferred') is not True or
            supplier.get('desired_mode') != 'SUPPLIER_BACKED_UNLIMITED'):
        raise ValueError('Unsupported business clarification.')
    for section, rows, count in ((units, proposal['rows'], 76), (supplier, proposal['shortfalls'], 6)):
        skus = section.get('skus', [])
        if (len(skus) != count or len(set(skus)) != count or len(rows) != count or
                set(skus) != {row['sku'] for row in rows}):
            raise ValueError('Clarification must cover the exact explicit SKU scope without duplicates.')
    if set(units['skus']) & set(supplier['skus']):
        raise ValueError('Pilot stock and supplier scopes overlap.')

    pilot_items, supplier_items, identities = [], [], set()
    for row in proposal['rows'] + proposal['shortfalls']:
        identity = (row['ecwid_product_id'], row['ecwid_combination_id'])
        if identity in identities:
            raise ValueError('Duplicate stock target in proposal.')
        identities.add(identity)
        matches = [target for target in catalogue['stock_targets'] if
                   (target['id'], target.get('combinationId')) == identity]
        if len(matches) != 1:
            raise ValueError('Live identity is missing or duplicated.')
        target = matches[0]
        if (target['sku'] != row['sku'] or target.get('variationOptions') != row['variation_options'] or
                target.get('quantity') != row['ecwid_quantity']):
            raise ValueError('Live identity, options or snapshot quantity differs.')
        item = {key: deepcopy(row[key]) for key in ('sku', 'name', 'ecwid_product_id', 'ecwid_combination_id', 'variation_options')}
        item.update({'active': False, 'opening_balance_approved': False, 'ecwid_stock_write_authorized': False})
        if row['sku'] in units['skus']:
            item.update({'unit_confirmed_by_user': True, 'unit_ratio': 1,
                         'snapshot_workbook_balance': row['physical_balance'],
                         'snapshot_unpicked_units': row['reserved_ecwid_units'],
                         'snapshot_proposed_ecwid_quantity': row['provisional_quantity_if_one_to_one'],
                         'source_rows': deepcopy(row['source_rows'])})
            pilot_items.append(item)
            continue
        confirmed_lines = []
        for line in row['reservation_lines']:
            matching_orders = [order for order in orders['orders'] if order['id'] == line['order_id']]
            if len(matching_orders) != 1:
                raise ValueError('Supplier order is missing or duplicated.')
            order = matching_orders[0]
            matching_lines = [order_line for order_line in order['lines'] if order_line['ecwid_line_id'] == line['ecwid_line_id']]
            if order['payment_status'] != 'PAID' or order['fulfillment_status'] != 'PROCESSING' or len(matching_lines) != 1:
                raise ValueError('Supplier order status or line scope changed.')
            order_line = matching_lines[0]
            if (order_line['sku'] != row['sku'] or order_line['ecwid_product_id'] != identity[0] or
                    order_line.get('ecwid_combination_id') != identity[1] or
                    order_line.get('selected_options_supported') is not True or
                    order_line.get('selected_options') != row['variation_options'] or
                    type(order_line.get('ordered_quantity')) is not int or order_line['ordered_quantity'] <= 0 or
                    order_line['ordered_quantity'] != line['quantity']):
                raise ValueError('Supplier order line identity, options or quantity differs.')
            confirmed_lines.append({**line, 'receipt_status': 'AWAITING_RECEIPT', 'physical_allocation_confirmed': False})
        if not confirmed_lines or sum(line['quantity'] for line in confirmed_lines) != row['reserved_ecwid_units']:
            raise ValueError('Supplier demand does not match exact order lines.')
        item.update({'desired_mode': supplier['desired_mode'], 'classification_source': 'USER_EXPLICIT_SKU_LIST',
                     'historical_workbook_rows': deepcopy(row['source_rows']), 'opening_physical_balance': None,
                     'physical_allocation': None, 'order_lines': confirmed_lines,
                     'snapshot_pending_receipt_units': row['reserved_ecwid_units'],
                     'snapshot_ecwid_unlimited': target['unlimited'], 'snapshot_ecwid_quantity': target.get('quantity'),
                     'ecwid_policy_review_required': target['unlimited'] is not True,
                     'proposed_ecwid_write_quantity': None})
        supplier_items.append(item)
    return {'kind': 'SUPPLIER_WORKFLOW_DESIGN_REVIEW', 'schema_version': 1, 'dry_run': True, 'store_id': '2442119',
            'generated_at': datetime.now(timezone.utc).isoformat(), 'runtime_implemented': False,
            'database_imports_performed': False, 'stock_writes_performed': False, 'cutover_approved': False,
            'clarification': clarification, 'clarification_sha256': sha(clarification_bytes),
            'pilot_items': pilot_items, 'supplier_items': supplier_items,
            'counts': {'unit_confirmed': len(pilot_items), 'supplier_items_awaiting_receipt': len(supplier_items),
                       'ecwid_policy_exceptions': sum(item['ecwid_policy_review_required'] for item in supplier_items)},
            'note': 'The old all-on-shelf statement is superseded for these exact supplier lines. '
                    'Unknown opening physical balances remain unknown. This is a design review, not an import payload.'}


def render_report(report):
    esc = lambda value: html.escape(str(value), quote=True)
    supplier_rows = []
    for item in report['supplier_items']:
        lines = ', '.join(f"{line['order_id']}: {line['quantity']}" for line in item['order_lines'])
        policy = ('Unlimited' if item['snapshot_ecwid_unlimited'] else
                  f"Finite ({item['snapshot_ecwid_quantity']:,}) — recheck before activation")
        supplier_rows.append(f'<tr><td><strong>{esc(item["sku"])}</strong><small>{esc(item["name"])}</small></td>'
                             f'<td>{esc(lines)}</td><td>Awaiting receipt</td><td>Not yet confirmed</td><td>{esc(policy)}</td></tr>')
    exceptions = [item for item in report['supplier_items'] if item['ecwid_policy_review_required']]
    exception_notice = ('<p class="warning"><strong>Settings to recheck:</strong> ' +
                        ', '.join(esc(item['sku']) for item in exceptions) +
                        ' had finite/tracked stock in the saved catalogue. Do not automatically switch these targets '
                        'or use their Ecwid quantities as physical stock.</p>') if exceptions else ''
    pilot_rows = []
    for item in report['pilot_items']:
        source = item['source_rows'][0]
        pilot_rows.append(f'<tr><td><strong>{esc(item["sku"])}</strong><small>{esc(item["name"])}</small></td>'
                          f'<td>Confirmed 1:1</td><td>{item["snapshot_workbook_balance"]:,}</td>'
                          f'<td>{item["snapshot_unpicked_units"]:,}</td><td>{item["snapshot_proposed_ecwid_quantity"]:,}</td>'
                          f'<td>{esc(source["location"])}<small>{esc(source["source_sheet"])} row {source["source_row"]}</small></td></tr>')
    return f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="referrer" content="no-referrer"><title>Supplier workflow and pilot decisions</title><style>
*{{box-sizing:border-box}}body{{font:15px/1.55 system-ui,sans-serif;margin:0;color:#173149;background:#f5f7f9}}main{{max-width:1180px;margin:auto;padding:28px 20px 60px}}
h1{{font-size:28px;line-height:1.25}}h2{{font-size:21px;margin-top:28px}}h3{{font-size:17px}}p{{max-width:920px}}a{{color:#135b91}}.notice{{padding:14px 18px;background:#e6f3ee;border-left:4px solid #247b69}}
.warning{{padding:14px 18px;background:#fff2d8;border-left:4px solid #a97415}}section,details{{background:white;border:1px solid #dbe3e9;border-radius:9px;padding:18px;margin:16px 0}}
.steps{{padding-left:22px}}.steps li{{margin:12px 0}}.steps strong{{display:block}}small{{display:block;font-size:12px;color:#5a6a77}}summary{{cursor:pointer;font-weight:600}}
.table{{overflow:auto}}table{{border-collapse:collapse;width:100%;font-size:13px}}th,td{{padding:12px;text-align:left;border-bottom:1px solid #dbe3e9;vertical-align:top}}th{{background:#edf2f5}}td:first-child{{min-width:195px}}input{{width:min(100%,360px);padding:10px;font:inherit;margin:10px 0;border:1px solid #a5b7c5;border-radius:6px}}.muted{{font-size:13px;color:#536776}}[hidden]{{display:none!important}}
@media(max-width:600px){{main{{padding:20px 14px}}h1{{font-size:25px}}section,details{{padding:13px}}}}
</style></head><body><main>
<small>Factory inventory · Store 2442119 · Updated business decisions</small><h1>Supplier workflow and pilot decisions</h1>
<p class="notice"><strong>All 76 pilot units are confirmed.</strong> One workbook unit equals one Ecwid sale unit, including items named “Set”. Opening quantities still need approval.</p>
<p>The six previously flagged stock discrepancies are supplier-sourced items with incomplete workbook records. Their current order goods are awaiting MISUMI receipt, not confirmed shelf stock.</p>
<p class="muted">Design only. No runtime workflow enabled, no inventory imported and no Ecwid stock or settings changed. Attribute-based identification is deferred; use an explicit SKU list.</p>
<section><h2 style="margin-top:0">Proposed staff workflow</h2><ol class="steps">
<li><strong>Order needs supplier goods</strong>Show demand, free shelf stock and allocations separately. Received but unassigned goods are “awaiting assignment”, not “awaiting supplier”. “Awaiting supply” does not mean “buy again”; goods may already be on order.</li>
<li><strong>Receive at the factory</strong>Scan and record the entire delivery, including exact-order purchases and extras. Nothing is marked picked just because it arrived.</li>
<li><strong>Assign received or free stock</strong>Staff chooses the exact order lines. Unassigned extras remain free. No automatic allocation priority is introduced.</li>
<li><strong>Pick a Paid order</strong>Pick only physically available, allocated units. Deduct from factory stock; keep supplier-backed Ecwid stock unlimited without quantity writes.</li>
<li><strong>Complete every order line</strong>Available lines in mixed orders may be picked while supplier lines wait. Fully picked does not automatically mean packed, shipped or Ready for Pickup.</li></ol>
<p class="muted">Unmapped items, changed order identities and contradictory statuses still require review. Awaiting Payment remains unpickable. Cancellation does not create a physical return.</p></section>
<h2>Six explicitly identified supplier items</h2><p>Order quantities below come from the saved 22 September snapshot. Your latest statement updates their receipt status. Their actual opening shelf counts remain unknown.</p>
<div class="table"><table><thead><tr><th>SKU / item</th><th>Order: units awaiting receipt</th><th>Receipt status</th><th>Opening shelf count</th><th>Saved Ecwid policy</th></tr></thead><tbody>{''.join(supplier_rows)}</tbody></table></div>
{exception_notice}
<h2>Next implementation steps</h2><ol><li>Add supplier mode, allocation records and atomic stock guards.</li><li>Build receiving, allocation and mixed-order picking screens.</li><li>Test with fictitious orders in an isolated database before importing real orders.</li><li>Verify opening counts and refresh orders/catalogue before a separately approved activation.</li></ol>
<details><summary>76 pilot SKUs — confirmed units, unchanged snapshot quantities</summary><p class="muted">These are historical proposal values, not a fresh shelf count or approved Ecwid write. Original source rows remain unchanged.</p>
<label for="search">Find SKU, name or location</label><br><input id="search" type="search" placeholder="e.g. 00691"><output id="count" aria-live="polite"></output>
<div class="table"><table id="pilot"><thead><tr><th>SKU / item</th><th>Unit mapping</th><th>Workbook balance</th><th>Unpicked units</th><th>Proposed Ecwid</th><th>Workbook source</th></tr></thead><tbody>{''.join(pilot_rows)}</tbody></table></div></details>
<p class="muted">This replaces the old unit questions and all-on-shelf assumption for the six supplier lines. Original review files are retained as historical evidence. No new stock, receipt or live-write approval is implied.</p>
</main><script>const search=document.querySelector('#search'),rows=[...document.querySelectorAll('#pilot tbody tr')];function filter(){{let count=0;for(const row of rows){{row.hidden=!row.textContent.toLowerCase().includes(search.value.trim().toLowerCase());if(!row.hidden)count++}}document.querySelector('#count').textContent=' '+count+' of '+rows.length+' SKUs'}}search.addEventListener('input',filter);filter();</script></body></html>\n'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for key in ('clarification', 'proposal', 'catalogue', 'orders', 'out'):
        parser.add_argument('--' + key, required=True)
    args = parser.parse_args()
    report = create_report(*(read(getattr(args, key)) for key in ('clarification', 'proposal', 'catalogue', 'orders')))
    output, root = Path(args.out).resolve(), Path('import-data').resolve()
    if output == root or root not in output.parents or output.exists():
        raise ValueError('Output must be a new private directory below import-data.')
    page = render_report(report)
    output.mkdir(mode=0o700)
    for name, content in [('workflow-review.json', json.dumps(report, ensure_ascii=False, indent=2) + '\n'), ('workflow.html', page)]:
        descriptor = os.open(output / name, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(descriptor, 'w', encoding='utf-8') as stream:
            stream.write(content)
    print(json.dumps({'counts': report['counts'], 'runtime_implemented': False, 'stock_writes_performed': False, 'output': str(output)}, indent=2))


if __name__ == '__main__':
    main()
