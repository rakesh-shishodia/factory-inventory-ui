"""Business clarification reports cannot become inventory or policy approvals."""
from copy import deepcopy
import importlib.util
import json
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location('workflow_review', Path(__file__).resolve().parents[1] / 'scripts' / 'prepare-workflow-review.py')
review = importlib.util.module_from_spec(spec)
spec.loader.exec_module(review)


def encoded(value):
    return json.dumps(value).encode()


def fixture():
    rows = []
    targets = []
    order_lines = []
    for index in range(82):
        sku = f'SKU-{index:03}'
        supplier = index >= 76
        row = {'sku': sku, 'name': 'Test item', 'ecwid_product_id': str(index + 1),
               'ecwid_combination_id': None, 'variation_options': [], 'ecwid_quantity': None if supplier else 12,
               'physical_balance': 2 if supplier else 10, 'reserved_ecwid_units': 4 if supplier else 0,
               'provisional_quantity_if_one_to_one': None if supplier else 10,
               'source_rows': [{'source_row': index + 4, 'source_sheet': 'Stock Sheet', 'location': 'Shelf A', 'balance': 2 if supplier else 10}],
               'reservation_lines': [{'order_id': '99', 'ecwid_line_id': str(index), 'quantity': 4}] if supplier else []}
        rows.append(row)
        targets.append({'id': str(index + 1), 'combinationId': None, 'sku': sku, 'variationOptions': [],
                        'quantity': row['ecwid_quantity'], 'unlimited': supplier})
        if supplier:
            order_lines.append({'ecwid_line_id': str(index), 'ecwid_product_id': str(index + 1),
                                'ecwid_combination_id': None, 'sku': sku, 'ordered_quantity': 4,
                                'selected_options': [], 'selected_options_supported': True})
    catalogue = {'kind': 'READONLY_CATALOGUE', 'complete': True, 'store_id': '2442119', 'stock_targets': targets}
    orders = {'kind': 'READONLY_ORDER_REVIEW', 'complete': True, 'store_id': '2442119',
              'orders': [{'id': '99', 'payment_status': 'PAID', 'fulfillment_status': 'PROCESSING', 'lines': order_lines}]}
    proposal = {'kind': 'PILOT_APPROVAL_PROPOSAL', 'dry_run': True, 'store_id': '2442119', 'review_sha256': 'a' * 64,
                'batch_approved': False, 'cutover_approved': False, 'ecwid_write_authorized': False,
                'stock_writes_performed': False, 'database_imports_performed': False, 'rows': rows[:76], 'shortfalls': rows[76:]}
    clarification = {'kind': 'PILOT_BUSINESS_CLARIFICATION', 'schema_version': 1, 'store_id': '2442119',
                     'review_sha256': 'a' * 64, 'supersedes_all_on_shelves_confirmation_sha256': 'b' * 64,
                     'opening_balances_approved': False, 'database_import_approved': False, 'cutover_approved': False,
                     'ecwid_stock_write_authorized': False, 'ecwid_policy_change_authorized': False,
                     'unit_confirmation': {'confirmed': True, 'ratio': 1, 'skus': [item['sku'] for item in rows[:76]]},
                     'supplier_workflow': {'goods_always_arrive_at_factory': True,
                         'current_pending_supplier_lines_awaiting_receipt': True, 'workbook_balances_are_complete': False,
                         'opening_physical_balances': None, 'identification': 'EXPLICIT_SKU_MAPPING',
                         'attribute_identification_deferred': True, 'desired_mode': 'SUPPLIER_BACKED_UNLIMITED',
                         'skus': [item['sku'] for item in rows[76:]]}}
    return clarification, proposal, catalogue, orders


def inputs(fixture_values):
    clarification, proposal, catalogue, orders = deepcopy(fixture_values)
    cat_bytes, order_bytes = encoded(catalogue), encoded(orders)
    proposal['inputs'] = {'live_catalog': {'sha256': review.sha(cat_bytes)}, 'orders': {'sha256': review.sha(order_bytes)},
                          'pick_confirmation': {'sha256': clarification['supersedes_all_on_shelves_confirmation_sha256']}}
    proposal_bytes = encoded(proposal)
    clarification.update({'approval_proposal_sha256': review.sha(proposal_bytes), 'catalogue_sha256': review.sha(cat_bytes),
                          'orders_sha256': review.sha(order_bytes)})
    return encoded(clarification), proposal_bytes, cat_bytes, order_bytes


class WorkflowReviewTests(unittest.TestCase):
    def test_units_confirmed_without_stock_or_import_approval(self):
        report = review.create_report(*inputs(fixture()))
        self.assertEqual(report['counts']['unit_confirmed'], 76)
        self.assertTrue(all(item['unit_confirmed_by_user'] for item in report['pilot_items']))
        self.assertTrue(all(not item['active'] and not item['opening_balance_approved'] for item in report['pilot_items']))
        for key in ('runtime_implemented', 'database_imports_performed', 'stock_writes_performed', 'cutover_approved'):
            self.assertFalse(report[key])

    def test_supplier_physical_unknown_not_zero_or_workbook_value(self):
        report = review.create_report(*inputs(fixture()))
        for item in report['supplier_items']:
            self.assertIsNone(item['opening_physical_balance'])
            self.assertIsNone(item['physical_allocation'])
            self.assertIsNone(item['proposed_ecwid_write_quantity'])
            self.assertEqual(item['historical_workbook_rows'][0]['balance'], 2)
            self.assertEqual(item['snapshot_pending_receipt_units'], 4)
            self.assertEqual(item['order_lines'][0]['receipt_status'], 'AWAITING_RECEIPT')
            self.assertFalse(item['order_lines'][0]['physical_allocation_confirmed'])

    def test_finite_exception_retained_without_policy_mutation(self):
        values = fixture()
        values[1]['shortfalls'][0]['ecwid_quantity'] = 4797
        values[2]['stock_targets'][76].update({'unlimited': False, 'quantity': 4797})
        report = review.create_report(*inputs(values))
        item = report['supplier_items'][0]
        self.assertEqual(report['counts']['ecwid_policy_exceptions'], 1)
        self.assertFalse(item['snapshot_ecwid_unlimited'])
        self.assertEqual(item['snapshot_ecwid_quantity'], 4797)
        self.assertFalse(report['clarification']['ecwid_policy_change_authorized'])
        self.assertIn('Settings to recheck', review.render_report(report))

    def test_no_exception_notice_when_all_snapshot_targets_unlimited(self):
        self.assertNotIn('Settings to recheck', review.render_report(review.create_report(*inputs(fixture()))))

    def test_changed_source_bytes_rejected(self):
        original = inputs(fixture())
        for index in (1, 2, 3):
            changed = list(original)
            changed[index] += b' '
            with self.subTest(index=index), self.assertRaisesRegex(ValueError, 'snapshot'):
                review.create_report(*changed)

    def test_duplicate_missing_or_extra_unit_confirmation_rejected(self):
        for operation in ('duplicate', 'missing', 'extra'):
            values = fixture()
            skus = values[0]['unit_confirmation']['skus']
            if operation == 'duplicate':
                skus[-1] = skus[0]
            elif operation == 'missing':
                skus.pop()
            else:
                skus.append('EXTRA')
            with self.subTest(operation=operation), self.assertRaises(ValueError):
                review.create_report(*inputs(values))

    def test_attributes_or_opening_balance_inference_rejected(self):
        for change in ({'identification': 'ATTRIBUTE'}, {'attribute_identification_deferred': False},
                       {'opening_physical_balances': 0}, {'workbook_balances_are_complete': True}):
            values = fixture()
            values[0]['supplier_workflow'].update(change)
            with self.subTest(change=change), self.assertRaises(ValueError):
                review.create_report(*inputs(values))

    def test_live_authority_not_accepted(self):
        for key in ('opening_balances_approved', 'database_import_approved', 'cutover_approved',
                    'ecwid_stock_write_authorized', 'ecwid_policy_change_authorized'):
            values = fixture()
            values[0][key] = True
            with self.subTest(key=key), self.assertRaises(ValueError):
                review.create_report(*inputs(values))

    def test_wrong_order_identity_options_quantity_or_payment_rejected(self):
        for change in ({'ordered_quantity': 5}, {'ecwid_product_id': '1000'}, {'sku': 'WRONG'},
                       {'selected_options': [{'name': 'Size', 'value': 'Other'}]}, {'selected_options_supported': False}):
            values = fixture()
            values[3]['orders'][0]['lines'][0].update(change)
            with self.subTest(change=change), self.assertRaises(ValueError):
                review.create_report(*inputs(values))
        values = fixture()
        values[3]['orders'][0]['payment_status'] = 'AWAITING_PAYMENT'
        with self.assertRaises(ValueError):
            review.create_report(*inputs(values))

    def test_prior_confirmation_and_store_are_pinned(self):
        args = list(inputs(fixture()))
        clarification = json.loads(args[0])
        clarification['supersedes_all_on_shelves_confirmation_sha256'] = 'c' * 64
        args[0] = encoded(clarification)
        with self.assertRaises(ValueError):
            review.create_report(*args)
        values = fixture()
        values[0]['store_id'] = '999'
        with self.assertRaises(ValueError):
            review.create_report(*inputs(values))

    def test_html_safety_and_design_only_wording(self):
        report = review.create_report(*inputs(fixture()))
        attack = '<img src=x onerror=alert(1)>'
        report['supplier_items'][0]['name'] = attack
        page = review.render_report(report)
        self.assertNotIn(attack, page)
        self.assertIn('&lt;img', page)
        self.assertIn('Design only', page)
        self.assertIn('awaiting assignment', page)
        self.assertIn('Opening quantities still need approval', page)
        self.assertIn('Attribute-based identification is deferred', page)


if __name__ == '__main__':
    unittest.main()
