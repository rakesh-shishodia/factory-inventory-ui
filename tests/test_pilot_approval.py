"""Approval proposals preserve snapshot provenance, source values and user scope."""
from copy import deepcopy
import importlib.util
import json
from pathlib import Path
import unittest
from unittest.mock import patch

SCRIPT = Path(__file__).resolve().parents[1] / 'scripts' / 'prepare-pilot-approval.py'
spec = importlib.util.spec_from_file_location('approval', SCRIPT)
approval = importlib.util.module_from_spec(spec)
spec.loader.exec_module(approval)


def row(sku='LP-M5B-20', **changes):
    return {
        'sku': sku, 'name': 'M5 Screw', 'ecwid_product_id': '123', 'ecwid_combination_id': '456',
        'variation_options': [{'name': 'Length', 'value': '20mm'}], 'category_paths': ['Fasteners / Screws'],
        'ecwid_quantity': 100, 'source_rows': [{'source_row': 20, 'source_sheet': 'Stock Sheet',
            'sku': sku, 'name': 'M5 Screw', 'unit': 'Pcs', 'location': 'A1', 'balance': 20, 'status': 'CANDIDATE'}],
        'physical_balance': 20, 'reserved_ecwid_units': 5,
        'reservation_lines': [{'order_id': '12', 'ecwid_line_id': '13', 'quantity': 5}],
        'unresolved_order_lines': [], 'technical_match': True,
        'provisional_quantity_if_one_to_one': 15, 'desired_ecwid_quantity': None,
        'single_unit_confirmed': False, 'ready': False, 'status': 'UNIT_CONFIRMATION',
        'issues': [{'code': 'SINGLE_UNIT_CONFIRMATION_REQUIRED', 'message': 'Confirm units.'}], **changes,
    }


def report(*rows):
    return {
        'kind': 'PILOT_SCOPE_REVIEW', 'schema_version': 1, 'dry_run': True, 'store_id': '2442119',
        'current_authority': 'WORKBOOK', 'ready_count': 0, 'stock_writes_performed': False,
        'database_imports_performed': False, 'single_unit_mappings_approved': False, 'cutover_approved': False,
        'scope': {'live_crosscheck': {'verified': True}},
        'counts': {'technical_matches_awaiting_unit_confirmation': sum(item['technical_match'] for item in rows)},
        'source': {'sha256': 'a' * 64, 'source_ref': 'https://docs.google.com/spreadsheets/d/test/edit',
                   'source_modified_at': '2026-09-22T05:31:10.702Z'},
        'order_review': {'pending_order_count': 1, 'order_snapshot_completed_at': '2026-09-22T07:07:03.784Z'},
        'rows': list(rows),
    }


def proposal(*rows, digest=approval.CONTEXT_REVIEW_SHA):
    return approval.create_proposal(report(*(rows or [row()])), digest)


class PilotApprovalTests(unittest.TestCase):
    def test_prior_clarification_is_display_context_not_write_approval(self):
        original = report(row())
        before = deepcopy(original)
        result = approval.create_proposal(original, approval.CONTEXT_REVIEW_SHA)
        self.assertEqual(original, before)
        self.assertEqual(result['kind'], 'PILOT_APPROVAL_PROPOSAL')
        self.assertTrue(result['rows'][0]['unit_clarification_carried'])
        self.assertFalse(result['rows'][0]['single_unit_confirmed'])
        self.assertFalse(result['rows'][0]['opening_balance_approved'])
        self.assertIsNone(result['rows'][0]['desired_ecwid_quantity'])
        for key in ('batch_approved', 'cutover_approved', 'ecwid_write_authorized', 'database_imports_performed', 'stock_writes_performed'):
            self.assertFalse(result[key])

    def test_changed_snapshot_does_not_carry_old_clarification(self):
        result = proposal(digest='b' * 64)
        self.assertFalse(result['rows'][0]['unit_clarification_carried'])
        self.assertEqual(result['rows'][0]['approval_group'], 'GROUP_CONFIRMATION')

    def test_name_or_sku_prefix_does_not_auto_confirm(self):
        for item in (row('LP-M5B-999'), row('TNP40-M5'), row('01517', name='Makerlink Tee Nut'),
                     row(ecwid_combination_id=None)):
            with self.subTest(sku=item['sku']):
                self.assertFalse(proposal(item)['rows'][0]['unit_clarification_carried'])

    def test_three_ambiguous_sets_and_new_pack_names_stay_held(self):
        for sku in (*approval.SET_QUESTIONS, 'UNKNOWN-PACK'):
            with self.subTest(sku=sku):
                result = proposal(row(sku, name='Component pack'))['rows'][0]
                self.assertEqual(result['approval_group'], 'SET_ASSEMBLY_HOLD')
                self.assertFalse(result['unit_clarification_carried'])
                self.assertIsNotNone(result['question'])

    def test_true_zero_is_preserved_not_missing(self):
        item = row(physical_balance=0, reserved_ecwid_units=0, reservation_lines=[], provisional_quantity_if_one_to_one=0)
        item['source_rows'][0]['balance'] = 0
        self.assertEqual(proposal(item)['rows'][0]['provisional_quantity_if_one_to_one'], 0)

    def test_invalid_and_unknown_balances_are_rejected(self):
        for value in (None, -1, 1.5, True):
            with self.subTest(value=value), self.assertRaises(ValueError):
                proposal(row(physical_balance=value))

    def test_shortfalls_have_no_negative_or_clamped_proposal(self):
        item = row(reserved_ecwid_units=25, reservation_lines=[{'order_id': '12', 'ecwid_line_id': '13', 'quantity': 25}],
                   technical_match=False, provisional_quantity_if_one_to_one=None,
                   issues=[{'code': 'INSUFFICIENT_PHYSICAL_BALANCE', 'message': 'Recount.'}])
        result = proposal(item)
        self.assertEqual(result['rows'], [])
        self.assertEqual(result['shortfalls'][0]['shortfall_if_one_to_one'], 5)
        self.assertIsNone(result['shortfalls'][0]['provisional_quantity_if_one_to_one'])
        self.assertIsNone(result['shortfalls'][0]['desired_ecwid_quantity'])
        for value in (-5, 0):
            with self.subTest(value=value), self.assertRaises(ValueError):
                proposal({**item, 'provisional_quantity_if_one_to_one': value})

    def test_changed_quantity_or_reservation_breaks_validation(self):
        for changes in ({'provisional_quantity_if_one_to_one': 14}, {'reserved_ecwid_units': 6},
                        {'reservation_lines': []}, {'unresolved_order_lines': [{'order_id': '14'}]},
                        {'desired_ecwid_quantity': 15}, {'ready': True}, {'single_unit_confirmed': True}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                proposal(row(**changes))

    def test_duplicate_source_sku_or_target_rejected(self):
        for items in ((row(), row()), (row(), row('NEW')), (row(source_rows=[]),),
                      (row(source_rows=[row()['source_rows'][0]] * 2),)):
            with self.assertRaises(ValueError):
                proposal(*items)

    def test_unverified_or_approved_envelope_rejected(self):
        for changes in ({'scope': {'live_crosscheck': {'verified': False}}}, {'store_id': '999'},
                        {'batch_approved': True, 'cutover_approved': True}, {'database_imports_performed': True}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                approval.create_proposal({**report(row()), **changes}, approval.CONTEXT_REVIEW_SHA)

    def test_extra_eligibility_issue_is_not_approval_ready(self):
        item = row()
        item['issues'].append({'code': 'TRACKING_OR_QUANTITY', 'message': 'Unknown quantity.'})
        with self.assertRaises(ValueError):
            proposal(item)

    def test_render_escapes_data_and_rejects_unsafe_source_link(self):
        attack = '<img src=x onerror=alert(1)>'
        result = proposal(row(name=attack))
        result['source']['source_ref'] = 'javascript:alert(1)'
        page = approval.render_report(result)
        self.assertNotIn(attack, page)
        self.assertIn('&lt;img', page)
        self.assertNotIn('href="javascript:', page)
        self.assertIn('No stock imported', page)
        self.assertIn('does not save approvals', page)
        self.assertIn('11:01 IST', page)
        self.assertIn('12:37 IST', page)

    def test_reproduction_requires_all_exact_input_hashes_and_lengths(self):
        inputs = {key: b'{}' for key in approval.INPUT_KEYS}
        base = report(row())
        base['inputs'] = {key: {'sha256': approval.pilot.sha(data), 'bytes': len(data)} for key, data in inputs.items()}
        for key in approval.INPUT_KEYS:
            changed = {**inputs, key: b'{ }'}
            with self.subTest(key=key), self.assertRaisesRegex(ValueError, 'differs'):
                approval.reproduce_review(json.dumps(base).encode(), changed)

    def test_reproduction_detects_tampered_derived_report(self):
        inputs = {key: b'{}' for key in approval.INPUT_KEYS}
        expected = report(row())
        base = deepcopy(expected)
        base['inputs'] = {key: {'sha256': approval.pilot.sha(data), 'bytes': len(data)} for key, data in inputs.items()}
        with patch.object(approval.pilot, 'parse_scope_csv', return_value={}), \
                patch.object(approval.pilot, 'parse_source', return_value={}), \
                patch.object(approval.pilot.catalogue, 'parse_live_catalogue', return_value={'stock_targets': []}), \
                patch.object(approval.pilot, 'verify_live_scope', return_value={}), \
                patch.object(approval.pilot, 'reservation_review', return_value={}), \
                patch.object(approval.pilot, 'create_review', return_value=expected):
            self.assertEqual(approval.reproduce_review(json.dumps(base).encode(), inputs), base)
            base['rows'][0]['physical_balance'] = 999
            with self.assertRaisesRegex(ValueError, 'does not reproduce'):
                approval.reproduce_review(json.dumps(base).encode(), inputs)


if __name__ == '__main__':
    unittest.main()
