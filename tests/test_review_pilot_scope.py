"""Pilot scope reports preserve identity, missing values and approval boundaries."""
import csv
import importlib.util
from io import StringIO
import json
from pathlib import Path
import unittest

SCRIPT = Path(__file__).resolve().parents[1] / 'scripts' / 'review-pilot-scope.py'
spec = importlib.util.spec_from_file_location('pilot_review', SCRIPT)
review = importlib.util.module_from_spec(spec)
spec.loader.exec_module(review)


def scope():
    fields = ['type', 'category_internal_id', 'category_path', 'product_internal_id', 'product_category_1', 'product_name']
    text = StringIO(newline='')
    writer = csv.DictWriter(text, fieldnames=fields)
    writer.writeheader()
    for identity in review.ROOT_IDS:
        writer.writerow({'type': 'category', 'category_internal_id': identity, 'category_path': 'Category ' + identity})
    writer.writerow({'type': 'category', 'category_internal_id': '7135088', 'category_path': review.FASTENERS_NAME})
    writer.writerow({'type': 'category', 'category_internal_id': '3', 'category_path': review.FASTENERS_NAME + ' / Nuts'})
    writer.writerow({'type': 'product', 'product_internal_id': '1', 'product_category_1': review.FASTENERS_NAME + ' / Nuts', 'product_name': 'Nut'})
    writer.writerow({'type': 'product', 'product_internal_id': '2', 'product_category_1': 'Not selected', 'product_name': 'Other'})
    return review.parse_scope_csv(text.getvalue().encode())


def source_row(**changes):
    return {'source_sheet': 'Stock Sheet', 'source_row': 4, 'sku': 'NUT', 'balance': 8, 'unit': 'Pcs',
            'location': 'Shelf A', 'status': 'CANDIDATE', 'name': 'Nut', **changes}


def source(*rows):
    return {'kind': 'SOURCE_REVIEW', 'dry_run': True, 'balance_meaning': 'PHYSICAL_ON_HAND',
            'counts': {'total': len(rows)}, 'source': {'sha256': 'a' * 64}, 'rows': list(rows)}


def target(**changes):
    return {'id': '1', 'combinationId': '23', 'sku': 'NUT', 'name': 'Nut', 'quantity': 7,
            'unlimited': False, 'enabled': True, 'hasOptions': True, 'hasVariations': False,
            'variationOptions': [{'name': 'Size', 'value': 'M3'}], 'eligibilityVerified': True,
            'hasExtraOptions': False, 'hasBundleRelationships': False, **changes}


def line(**changes):
    return {'order_id': '12', 'ecwid_line_id': '200', 'ecwid_product_id': '1', 'ecwid_combination_id': '23',
            'sku': 'NUT', 'ordered_quantity': 2, 'selected_options': [{'name': 'Size', 'value': 'M3'}],
            'selected_options_supported': True, **changes}


def run(rows=None, targets=None, lines=None, checked=True):
    return review.create_review(scope(), source(*(rows if rows is not None else [source_row()])),
                                targets if targets is not None else [target()],
                                {'lines': lines if lines is not None else []}, {'verified': checked})


class PilotReviewTests(unittest.TestCase):
    def test_parent_selection_includes_subcategories_not_unrelated(self):
        selected = scope()
        self.assertEqual(selected['selected_ids'], {'1'})
        self.assertEqual(selected['groups']['7135088'], {'1'})
        self.assertIn('3', selected['category_ids'])

    def test_variant_targets_count_separately_without_inventing_readiness(self):
        report = run(targets=[target(), target(combinationId='24', sku='NUT2')])
        self.assertEqual(report['scope']['parent_product_count'], 1)
        self.assertEqual(report['scope']['stock_target_count'], 2)
        self.assertEqual(report['scope']['variation_target_count'], 2)
        self.assertEqual(report['ready_count'], 0)
        self.assertFalse(report['rows'][0]['single_unit_confirmed'])
        self.assertIsNone(report['rows'][0]['desired_ecwid_quantity'])

    def test_missing_source_never_becomes_zero(self):
        row = run(rows=[])['rows'][0]
        self.assertIsNone(row['physical_balance'])
        self.assertIsNone(row['provisional_quantity_if_one_to_one'])
        self.assertIn('NO_SOURCE_ROW', [x['code'] for x in row['issues']])

    def test_real_zero_survives(self):
        row = run(rows=[source_row(balance=0)])['rows'][0]
        self.assertEqual(row['physical_balance'], 0)
        self.assertEqual(row['provisional_quantity_if_one_to_one'], 0)
        self.assertTrue(row['technical_match'])

    def test_full_source_duplicates_are_not_combined(self):
        row = run(rows=[source_row(), source_row(source_row=9, balance=4, status='REVIEW')])['rows'][0]
        self.assertEqual(len(row['source_rows']), 2)
        self.assertIsNone(row['physical_balance'])
        self.assertIn('DUPLICATE_SOURCE_SKU', [x['code'] for x in row['issues']])

    def test_candidate_wrapper_cannot_hide_other_rows(self):
        payload = source(source_row())
        self.assertEqual(len(review.parse_source(json.dumps(payload).encode())['rows']), 1)
        for changes in ({'kind': 'SOURCE_CANDIDATES'}, {'counts': {'total': 2}}, {'rows': [source_row(), source_row()]}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                review.parse_source(json.dumps({**payload, **changes}).encode())

    def test_order_reservation_requires_exact_id_sku_and_options(self):
        row = run(lines=[line()])['rows'][0]
        self.assertEqual(row['reserved_ecwid_units'], 2)
        self.assertEqual(row['provisional_quantity_if_one_to_one'], 6)
        for changes in ({'ecwid_combination_id': '24'}, {'selected_options_supported': False},
                        {'selected_options': [{'name': 'Size', 'value': 'M4'}]}, {'sku': 'WRONG'}):
            with self.subTest(changes=changes):
                bad = run(lines=[line(**changes)])['rows'][0]
                self.assertIsNone(bad['reserved_ecwid_units'])
                self.assertIsNone(bad['provisional_quantity_if_one_to_one'])

    def test_duplicate_live_sku_outside_scope_still_blocks_mapping(self):
        row = run(targets=[target(), target(id='2')])['rows'][0]
        self.assertFalse(row['technical_match'])
        self.assertIn('DUPLICATE_LIVE_SKU', [x['code'] for x in row['issues']])

    def test_tracking_disabled_and_source_shortage_block_plan(self):
        row = run(targets=[target(unlimited=True, quantity=None)], lines=[line(ordered_quantity=10)])['rows'][0]
        codes = [x['code'] for x in row['issues']]
        self.assertIn('TRACKING_OR_QUANTITY', codes)
        self.assertIn('INSUFFICIENT_PHYSICAL_BALANCE', codes)
        self.assertIsNone(row['provisional_quantity_if_one_to_one'])

    def test_live_category_crosscheck_detects_membership_drift(self):
        selected = scope()
        cats = [{'id': identity, 'name': value['path'], 'parentId': '7135088' if identity == '3' else '0'}
                for identity, value in selected['categories'].items()]
        live = {'kind': 'READONLY_CATEGORY_SCOPE', 'schema_version': 1, 'complete': True, 'dry_run': True, 'hidden_categories_included': True,
                'store_id': '2442119', 'category_count': len(cats), 'product_count': 2, 'categories': cats,
                'products': [{'id': '1', 'categoryIds': ['3']}, {'id': '2', 'categoryIds': []}]}
        self.assertTrue(review.verify_live_scope(live, selected, '2442119')['verified'])
        live['products'][0]['categoryIds'] = []
        result = review.verify_live_scope(live, selected, '2442119')
        self.assertFalse(result['verified'])
        self.assertEqual(result['differences']['only_csv_products'], ['1'])
        live['hidden_categories_included'] = False
        with self.assertRaises(ValueError):
            review.verify_live_scope(live, selected, '2442119')

    def test_unknown_order_combination_blocks_parent_even_when_sku_changed(self):
        row = run(lines=[line(ecwid_combination_id='999', sku='OLD-SKU')])['rows'][0]
        self.assertIsNone(row['reserved_ecwid_units'])
        self.assertIn('ORDER_IDENTITY_REVIEW', [x['code'] for x in row['issues']])

    def test_snapshot_confirmation_hash_and_scope_required(self):
        orders = {'kind': 'READONLY_ORDER_REVIEW', 'dry_run': True, 'complete': True, 'store_id': '2442119', 'pending_order_count': 1,
                  'orders': [{'id': '12', 'payment_status': 'PAID', 'fulfillment_status': 'PROCESSING', 'lines': [line()]}]}
        data = json.dumps(orders).encode()
        confirmation = {'kind': 'SNAPSHOT_PICK_CONFIRMATION', 'store_id': '2442119', 'order_snapshot': {'sha256': review.sha(data)},
                        'outstanding_order_scope_confirmed': True,
                        'confirmed_shipment_orders': [{'order_id': '12', 'line_count': 1, 'unchanged': True}],
                        'physical_pick_confirmation': {'all_items_still_on_shelves': True, 'previously_picked_quantity_per_line': 0, 'confirmed_line_count': 1}}
        self.assertEqual(len(review.reservation_review(data, json.dumps(confirmation).encode(), '2442119')['lines']), 1)
        confirmation['order_snapshot']['sha256'] = 'x' * 64
        with self.assertRaises(ValueError):
            review.reservation_review(data, json.dumps(confirmation).encode(), '2442119')

    def test_html_escapes_source_and_catalogue_text(self):
        attack = '<img src=x onerror=alert(1)>'
        page = review.render_report(run(targets=[target(name=attack)], rows=[source_row(location=attack)]))
        self.assertNotIn(attack, page)
        self.assertIn('&lt;img', page)
        self.assertIn('No stock has been imported', page)


if __name__ == '__main__':
    unittest.main()
