"""Offline catalogue identity checks, never inventory mutations."""
import csv
import importlib.util
from io import StringIO
import json
from pathlib import Path
import tempfile
import unittest

SCRIPT = Path(__file__).resolve().parents[1] / 'scripts' / 'review-ecwid-catalogue.py'
spec = importlib.util.spec_from_file_location('catalogue_review', SCRIPT)
review = importlib.util.module_from_spec(spec)
spec.loader.exec_module(review)


def csv_bytes(*records, fields=(*review.FIELDS, 'product_variation_option_Size')):
    text = StringIO(newline='')
    writer = csv.DictWriter(text, fieldnames=fields)
    writer.writeheader()
    for record in records:
        writer.writerow(record)
    return text.getvalue().encode('utf-8')


def product(sku='00123', product_id='1', **changes):
    return {'type': 'product', 'product_internal_id': product_id, 'product_sku': sku,
            'product_name': 'Bolt', 'product_is_inventory_tracked': 'true',
            'product_quantity': '5', 'product_is_available': 'true', **changes}


def child(kind, sku='00123', product_id='1', **changes):
    defaults = {'product_is_inventory_tracked': 'true', 'product_quantity': '5', 'product_variation_option_Size': 'M3'} if kind == 'product_variation' else {}
    return {'type': kind, 'product_internal_id': product_id, 'product_sku': sku, **defaults, **changes}


def candidate_bytes(sku='00123', **changes):
    return json.dumps({'kind': 'SOURCE_CANDIDATES', 'dry_run': True,
                       'source_ref': 'https://example.test/stock', 'source_hash': 'a' * 64,
                       'balance_meaning': 'PHYSICAL_ON_HAND', 'ecwid_checked': False,
                       'reservations_confirmed': False,
                       'rows': [{'sku': sku, 'balance': 8, 'name': 'Factory bolt',
                                 'location': 'Shelf A', 'scan_code': sku,
                                 'source_row': 4, 'source_sheet': 'Stock Sheet', **changes}]}).encode()


def run_review(*records, sku='00123', candidate_changes=None):
    catalogue = review.parse_catalogue(csv_bytes(*records), 'catalogue.csv')
    candidates = review.parse_candidates(candidate_bytes(sku, **(candidate_changes or {})), 'candidates.json')
    return review.compare(catalogue, candidates, '2442119')


def live_target(**changes):
    return {'id': '1', 'combinationId': '23', 'sku': '00123', 'name': 'Bolt', 'quantity': 3,
            'unlimited': False, 'enabled': True, 'hasOptions': True, 'hasVariations': False,
            'variationOptions': [{'name': 'Size', 'value': 'M3'}], 'eligibilityVerified': True,
            'hasExtraOptions': False, 'hasBundleRelationships': False, **changes}


def live_bytes(targets=None, **changes):
    targets = [live_target()] if targets is None else targets
    return json.dumps({'kind': 'READONLY_CATALOGUE', 'schema_version': 1, 'dry_run': True,
                       'store_id': '2442119', 'complete': True,
                       'started_at': '2026-09-22T08:00:00Z', 'completed_at': '2026-09-22T08:00:10Z',
                       'stock_target_count': len(targets), 'stock_targets': targets, **changes}).encode()


def compare_live(data):
    catalogue = review.parse_catalogue(csv_bytes(product('PARENT'), child('product_variation', 'PARENT', product_variation_sku='00123')), 'catalogue.csv')
    candidates = review.parse_candidates(candidate_bytes(), 'candidates.json')
    live = review.parse_live_catalogue(data, 'catalog.json', '2442119')
    return review.compare(catalogue, candidates, '2442119', live)


class CatalogueTests(unittest.TestCase):
    def test_live_variant_snapshot_retains_both_source_hashes_and_no_ready(self):
        data = live_bytes()
        report = compare_live(data)
        self.assertEqual(report['rows'][0]['status'], 'VARIATION_MATCH')
        self.assertEqual(report['rows'][0]['live_matches'][0]['combinationId'], '23')
        self.assertEqual(report['live_catalogue_source']['sha256'], review.digest(data))
        self.assertEqual(report['catalogue_source']['format'], 'ECWID_CSV')
        self.assertEqual(report['ready_count'], 0)
        self.assertFalse(report['reservations_confirmed'])
        self.assertTrue(report['live_api_verified'])
        self.assertIn('API product 1, variation 23', review.render_report(report))

    def test_live_snapshot_envelope_fails_closed(self):
        for change in ({'kind': 'READY'}, {'complete': False}, {'dry_run': False}, {'schema_version': 2},
                       {'store_id': '999'}, {'stock_target_count': 100}, {'started_at': 'bad'}, {'stock_targets': ['bad']}):
            with self.subTest(change=change), self.assertRaises(ValueError):
                review.parse_live_catalogue(live_bytes(**change), 'catalog.json', '2442119')

    def test_live_variation_eligibility_fails_closed(self):
        for change in ({'quantity': None}, {'quantity': True}, {'quantity': -1}, {'quantity': 1.5}, {'unlimited': True},
                       {'enabled': False}, {'eligibilityVerified': None}, {'hasBundleRelationships': True},
                       {'hasExtraOptions': True}, {'combinationId': '0'}, {'variationOptions': []},
                       {'variationOptions': [{'name': 'Size', 'value': 'M3'}, {'name': 'Size', 'value': 'M4'}]}):
            with self.subTest(change=change):
                self.assertEqual(compare_live(live_bytes([live_target(**change)]))['rows'][0]['status'], 'REVIEW')

    def test_live_duplicates_and_changed_identity_are_blocked(self):
        report = compare_live(live_bytes([live_target(), live_target(combinationId='24')]))
        self.assertEqual(report['rows'][0]['status'], 'AMBIGUOUS_SKU')
        report = compare_live(live_bytes([live_target(), live_target(sku='OTHER')]))
        self.assertEqual(report['rows'][0]['status'], 'REVIEW')
        report = compare_live(live_bytes([live_target(id='2')]))
        self.assertEqual(report['rows'][0]['status'], 'REVIEW')
        report = compare_live(live_bytes([]))
        self.assertEqual(report['rows'][0]['status'], 'MISSING')

    def test_live_untrusted_option_text_is_escaped(self):
        attack = '<img src=x onerror=alert(1)>'
        report = compare_live(live_bytes([live_target(variationOptions=[{'name': 'Size', 'value': attack}])]))
        page = review.render_report(report)
        self.assertNotIn(attack, page)
        self.assertIn('&lt;img', page)

    def test_simple_match_keeps_zero_and_never_ready(self):
        report = run_review(product())
        row = report['rows'][0]
        self.assertEqual(row['sku'], '00123')
        self.assertEqual(row['status'], 'SIMPLE_MATCH')
        self.assertEqual(row['physical_balance'], 8)
        self.assertEqual(row['product_matches'][0]['quantity'], 5)
        self.assertEqual(report['ready_count'], 0)
        self.assertFalse(row['ready'])
        self.assertFalse(report['reservations_confirmed'])
        self.assertFalse(report['store_id_verified'])
        self.assertFalse(report['bundle_relationships_verified'])
        self.assertNotIn('ecwid_target_quantity', row)

    def test_normalize_spaces_case_not_leading_zeros(self):
        self.assertEqual(run_review(product(' ab001 '), sku='AB001')['rows'][0]['status'], 'SIMPLE_MATCH')
        self.assertEqual(run_review(product('123'))['rows'][0]['status'], 'MISSING')

    def test_quoted_commas_newline_utf8(self):
        name = 'Nut, bolt\nMétrique "small"'
        catalogue = review.parse_catalogue(csv_bytes(product(product_name=name)), 'catalogue.csv')
        row = catalogue['products'][0]
        self.assertEqual(row['name'], name)
        self.assertEqual(row['csv_record'], 2)
        self.assertEqual(row['csv_end_line'], 3)

    def test_duplicate_product_sku_never_first_match(self):
        row = run_review(product(), product('00123', '2'))['rows'][0]
        self.assertEqual(row['status'], 'AMBIGUOUS_SKU')
        self.assertEqual(len(row['product_matches']), 2)

    def test_duplicate_id_blocks_otherwise_unique_sku(self):
        row = run_review(product(), product('999', '1'))['rows'][0]
        self.assertEqual(row['status'], 'REVIEW')
        self.assertTrue(any('ID is duplicated' in reason for reason in row['reasons']))

    def test_option_child_join_works_before_product(self):
        row = run_review(child('product_option', product_option_name='Colour'), product())['rows'][0]
        self.assertEqual(row['status'], 'OPTIONS_VARIATIONS')
        self.assertEqual(row['product_matches'][0]['option_record_count'], 1)

    def test_variant_makes_parent_non_simple(self):
        row = run_review(product(), child('product_variation', product_variation_sku='00123-RED'))['rows'][0]
        self.assertEqual(row['status'], 'OPTIONS_VARIATIONS')

    def test_same_parent_variant_sku_collision(self):
        row = run_review(product(), child('product_variation', product_variation_sku='00123'))['rows'][0]
        self.assertEqual(row['status'], 'VARIATION_MATCH')

    def test_other_parent_variant_sku_collision(self):
        row = run_review(product(), product('PARENT', '2'), child('product_variation', 'PARENT', '2', product_variation_sku='00123'))['rows'][0]
        self.assertEqual(row['status'], 'AMBIGUOUS_SKU')

    def test_unique_variation_only(self):
        row = run_review(product('PARENT'), child('product_variation', 'PARENT', product_variation_sku='00123'))['rows'][0]
        self.assertEqual(row['status'], 'VARIATION_MATCH')
        self.assertFalse(row['ready'])
        self.assertFalse(row['single_unit_confirmed'])
        self.assertEqual(row['variation_matches'][0]['quantity'], 5)

    def test_variation_tracking_quantity_and_option_selection_required(self):
        for change in ({'product_is_inventory_tracked': 'false'}, {'product_is_inventory_tracked': ''},
                       {'product_quantity': ''}, {'product_quantity': '1.5'}, {'product_variation_option_Size': ''}):
            with self.subTest(change=change):
                row = run_review(product('PARENT'), child('product_variation', 'PARENT', product_variation_sku='00123', **change))['rows'][0]
                self.assertEqual(row['status'], 'REVIEW')

    def test_variation_uses_parent_enabled_but_not_parent_quantity_or_tracking(self):
        row = run_review(product('PARENT', product_quantity='100', product_is_inventory_tracked='true'),
                         child('product_variation', 'PARENT', product_variation_sku='00123', product_quantity='', product_is_inventory_tracked=''))['rows'][0]
        self.assertEqual(row['status'], 'REVIEW')
        disabled = run_review(product('PARENT', product_is_available='false'), child('product_variation', 'PARENT', product_variation_sku='00123'))
        self.assertEqual(disabled['rows'][0]['status'], 'REVIEW')

    def test_duplicate_variation_only_is_ambiguous(self):
        row = run_review(product('PARENT'), child('product_variation', 'PARENT', product_variation_sku='00123'),
                         child('product_variation', 'PARENT', product_variation_sku='00123'))['rows'][0]
        self.assertEqual(row['status'], 'AMBIGUOUS_SKU')

    def test_file_records_require_review(self):
        self.assertEqual(run_review(product(), child('product_file'))['rows'][0]['status'], 'REVIEW')

    def test_child_parent_sku_mismatch_blocks(self):
        row = run_review(product(), child('product_file', 'OTHER'))['rows'][0]
        self.assertTrue(any('different parent SKU' in reason for reason in row['reasons']))

    def test_orphan_child_fails(self):
        for kind in ('product_option', 'product_variation', 'product_file'):
            with self.subTest(kind=kind), self.assertRaisesRegex(ValueError, 'orphan'):
                review.parse_catalogue(csv_bytes(product(), child(kind, product_id='99')), 'test.csv')

    def test_blank_invalid_and_false_booleans_block(self):
        for field in ('product_is_inventory_tracked', 'product_is_available'):
            for value in ('', 'yes', '0', 'false'):
                with self.subTest(field=field, value=value):
                    self.assertEqual(run_review(product(**{field: value}))['rows'][0]['status'], 'REVIEW')

    def test_invalid_stock_blocks(self):
        for value in ('', '-1', '1.5', '1.00000000000000000001', 'NaN', 'Infinity', '1e3', '2147483648'):
            with self.subTest(value=value):
                self.assertEqual(run_review(product(product_quantity=value))['rows'][0]['status'], 'REVIEW')
        for value in ('0', '5.00', '2147483647'):
            with self.subTest(value=value):
                self.assertEqual(run_review(product(product_quantity=value))['rows'][0]['status'], 'SIMPLE_MATCH')

    def test_unknown_type_missing_header_uneven_rows_rejected(self):
        with self.assertRaisesRegex(ValueError, 'unsupported record type'):
            review.parse_catalogue(csv_bytes(product(type='bundle')), 'test.csv')
        with self.assertRaisesRegex(ValueError, 'missing required'):
            review.parse_catalogue(b'type,product_sku\nproduct,123\n', 'test.csv')
        with self.assertRaisesRegex(ValueError, 'different number'):
            review.parse_catalogue(csv_bytes(product()) + b'product,123\n', 'test.csv')

    def test_duplicate_header_malformed_csv_and_ids_rejected(self):
        with self.assertRaisesRegex(ValueError, 'duplicate'):
            review.parse_catalogue(b'type,type\nproduct,product\n', 'test.csv')
        with self.assertRaisesRegex(ValueError, 'Malformed CSV'):
            review.parse_catalogue(csv_bytes(product()) + b'"unterminated', 'test.csv')
        for value in ('', '0', '1.0', '-1', 'abc'):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, 'product ID'):
                review.parse_catalogue(csv_bytes(product(product_id=value)), 'test.csv')

    def test_categories_do_not_need_product_id(self):
        report = run_review(product(), {'type': 'category'})
        self.assertEqual(report['record_counts']['category'], 1)

    def test_source_hashes_from_exact_bytes(self):
        report = run_review(product())
        self.assertEqual(report['catalogue_source']['sha256'], review.digest(csv_bytes(product())))
        self.assertEqual(report['stock_source']['sha256'], review.digest(candidate_bytes()))
        self.assertEqual(report['stock_source']['workbook_sha256'], 'a' * 64)

    def test_candidate_schema_flags_and_quantity_rejected(self):
        for key, value in [('kind', 'READY'), ('dry_run', False), ('reservations_confirmed', True),
                           ('source_hash', 'bad'), ('rows', []), ('balance_meaning', 'AVAILABLE')]:
            payload = json.loads(candidate_bytes())
            payload[key] = value
            with self.subTest(key=key), self.assertRaises(ValueError):
                review.parse_candidates(json.dumps(payload).encode(), 'test.json')
        for value in (True, -1, 1.2, '5'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                review.parse_candidates(candidate_bytes(balance=value), 'test.json')

    def test_duplicate_candidate_fails(self):
        payload = json.loads(candidate_bytes())
        payload['rows'].append(dict(payload['rows'][0], source_row=5))
        with self.assertRaisesRegex(ValueError, 'SKU is duplicated'):
            review.parse_candidates(json.dumps(payload).encode(), 'test.json')

    def test_html_escapes_all_source_text(self):
        attack = '<script>alert("x")</script><img src=x onerror=alert(1)>'
        report = run_review(product(product_name=attack), candidate_changes={'name': attack, 'location': attack})
        page = review.render_report(report)
        self.assertNotIn(attack, page)
        self.assertIn('&lt;script&gt;', page)
        self.assertIn("connect-src 'none'", page)
        self.assertIn('No items ready for import', page)

    def test_output_new_private_directory_only(self):
        report = run_review(product())
        with tempfile.TemporaryDirectory() as root:
            output = Path(root) / 'review'
            review.write_report(report, output)
            self.assertEqual(output.stat().st_mode & 0o777, 0o700)
            for path in output.iterdir():
                self.assertEqual(path.stat().st_mode & 0o777, 0o600)
            with self.assertRaises(FileExistsError):
                review.write_report(report, output)

    def test_public_and_symlink_public_output_rejected(self):
        public = SCRIPT.parents[1] / 'public'
        report = run_review(product())
        with self.assertRaisesRegex(ValueError, 'public assets'):
            review.write_report(report, public / 'private-review')
        with tempfile.TemporaryDirectory() as root:
            alias = Path(root) / 'alias'
            alias.symlink_to(public, target_is_directory=True)
            with self.assertRaisesRegex(ValueError, 'public assets'):
                review.write_report(report, alias / 'private-review')


if __name__ == '__main__':
    unittest.main()
