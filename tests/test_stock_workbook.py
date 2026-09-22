import hashlib
import importlib.util
import io
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

spec = importlib.util.spec_from_file_location('stock_reader', Path(__file__).parent.parent / 'scripts' / 'extract_stock_workbook.py')
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
S = reader.NS['s']


def cell_xml(address, value, formula=None, attrs=None, cell_type=None):
    c = ET.Element('{' + S + '}c', {'r': address})
    if formula is not None:
        ET.SubElement(c, '{' + S + '}f', attrs or {}).text = formula
    if cell_type:
        c.set('t', cell_type)
        if value is not None:
            ET.SubElement(c, '{' + S + '}v').text = str(value)
    elif isinstance(value, str):
        c.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(c, '{' + S + '}is'), '{' + S + '}t').text = value
    elif value is not None:
        ET.SubElement(c, '{' + S + '}v').text = str(value)
    return c


def sheet_xml(rows):
    root = ET.Element('{' + S + '}worksheet')
    data = ET.SubElement(root, '{' + S + '}sheetData')
    for n, values in sorted(rows.items()):
        row = ET.SubElement(data, '{' + S + '}row', {'r': str(n)})
        for col, value in values.items():
            args = value if isinstance(value, tuple) else (value,)
            row.append(cell_xml(f'{col}{n}', *args))
    return ET.tostring(root)


def stock_row(n, sku='00123', sid='T1'):
    return {'B': 'Active', 'C': 'Single', 'G': sid, 'H': 'Bolt', 'I': sku,
            'K': 'Pcs', 'M': 'A1', 'N': 2, 'O': 5,
            'P': (3, f"SUMIF('Material Register'!$F:$F,'Stock Sheet'!G{n},'Material Register'!$L:$L)+SUMIF('BOM Master'!$D:$D,'Stock Sheet'!H:H,'BOM Master'!$G:$G)"),
            'Q': (1, f"SUMIF('Material Register'!$F:$F,'Stock Sheet'!G{n},'Material Register'!$K:$K)+SUMIF('BOM Master'!$D:$D,'Stock Sheet'!H:H,'BOM Master'!$H:$H)"),
            'R': (7, f'(P{n}+O{n}-Q{n})')}


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / 'source.xlsm'
        self.stock = {3: dict(reader.HEADERS), 4: stock_row(4)}
        self.material = {5: {'F': 'Store ID', 'K': 'Outward Qty', 'L': 'Inward Qty'}, 6: {'F': 'T1', 'K': 1, 'L': 3}}
        self.bom = {3: {'D': 'Item Name', 'G': 'In', 'H': 'Out'}}

    def tearDown(self):
        self.temp.cleanup()

    def save(self, extra=None):
        with ZipFile(self.path, 'w') as z:
            z.writestr('xl/workbook.xml', f'<workbook xmlns="{S}" xmlns:r="{reader.REL_NS}"><sheets><sheet name="Stock Sheet" sheetId="2" r:id="stock"/><sheet name="Material Register" sheetId="8" r:id="material"/><sheet name="BOM Master" sheetId="19" r:id="bom"/></sheets></workbook>')
            z.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="stock" Target="worksheets/stock.xml"/><Relationship Id="material" Target="worksheets/mat.xml"/><Relationship Id="bom" Target="/xl/worksheets/bom.xml"/></Relationships>')
            z.writestr('xl/worksheets/stock.xml', sheet_xml(self.stock))
            z.writestr('xl/worksheets/mat.xml', sheet_xml(self.material))
            z.writestr('xl/worksheets/bom.xml', sheet_xml(self.bom))
            z.writestr('xl/vbaProject.bin', b'not-executable-in-this-reader')
            for key, value in (extra or {}).items():
                z.writestr(key, value)

    def extract(self):
        return reader.extract(self.path, 'https://docs.google.com/spreadsheets/d/example/edit')

    def codes(self):
        return [i['code'] for i in self.extract()['rows'][0]['audit_issues']]

    def test_read_only_provenance_and_leading_zero_sku(self):
        self.save()
        before = self.path.read_bytes()
        snapshot = self.extract()
        self.assertEqual(snapshot['rows'][0]['sku'], '00123')
        self.assertEqual(snapshot['rows'][0]['source_row'], 4)
        self.assertEqual(snapshot['rows'][0]['balance'], 7)
        self.assertEqual(snapshot['rows'][0]['audit_issues'], [])
        self.assertEqual(snapshot['source']['sha256'], hashlib.sha256(before).hexdigest())
        self.assertEqual(self.path.read_bytes(), before)

    def test_shared_balance_formulas_resolve_relative_rows(self):
        self.stock[4]['R'] = (7, '(P4+O4-Q4)', {'t': 'shared', 'si': '1', 'ref': 'R4:R5'})
        self.stock[5] = stock_row(5, '00124')
        self.stock[5]['R'] = (7, '', {'t': 'shared', 'si': '1'})
        self.save()
        rows = self.extract()['rows']
        self.assertEqual(rows[1]['formulas']['balance'], '=(P5+O5-Q5)')
        self.assertFalse(rows[1]['audit_issues'])

    def test_unknown_shared_formula_stays_for_review(self):
        self.stock[4]['R'] = (7, 'SUM(O4:P4)-Q4', {'t': 'shared', 'si': '1', 'ref': 'R4:R5'})
        self.save()
        self.assertIn('UNSUPPORTED_FORMULA', self.codes())

    def test_missing_formula_cache_not_zero(self):
        self.stock[4]['R'] = (None, '(P4+O4-Q4)')
        self.save()
        self.assertIsNone(self.extract()['rows'][0]['balance'])
        self.assertIn('MISSING_FORMULA_RESULT', self.codes())

    def test_numeric_sku_is_not_reformatted_as_text(self):
        self.stock[4]['I'] = 123
        self.save()
        self.assertEqual(self.extract()['rows'][0]['sku'], 123)

    def test_shared_strings_preserve_leading_zeros(self):
        self.stock[4]['I'] = ('0', None, None, 's')
        self.save({'xl/sharedStrings.xml': f'<sst xmlns="{S}"><si><t>00123</t></si></sst>'})
        self.assertEqual(self.extract()['rows'][0]['sku'], '00123')

    def test_manual_balance_and_arithmetic_not_silently_repaired(self):
        self.stock[4]['R'] = 8
        self.save()
        self.assertIn('MANUAL_STOCK_VALUE', self.codes())
        self.assertIn('BALANCE_ARITHMETIC', self.codes())
        self.assertEqual(self.extract()['rows'][0]['balance'], 8)

    def test_shifted_movement_formula(self):
        self.stock[4]['P'] = (3, "SUMIF('Material Register'!F:F,F4,'Material Register'!L:L)")
        self.save()
        self.assertIn('UNEXPECTED_STOCK_FORMULA', self.codes())

    def test_quoted_sheet_name_spaces_remain_significant(self):
        self.stock[4]['P'] = (3, self.stock[4]['P'][1].replace('Material Register', 'MaterialRegister'))
        self.save()
        self.assertIn('UNEXPECTED_STOCK_FORMULA', self.codes())

    def test_formatting_whitespace_still_normalizes(self):
        self.stock[4]['R'] = (7, '( P4 + O4 - Q4 )')
        self.save()
        self.assertFalse(self.codes())

    def test_unassigned_movement_prevents_false_candidacy(self):
        for key in (None, '', (None, 'Other!A1'), ('#REF!', 'Other!A1', None, 'e')):
            with self.subTest(key=key):
                self.material[7] = {'F': key, 'L': 3}
                self.save()
                with self.assertRaisesRegex(ValueError, 'Unassigned register/BOM movement'):
                    self.extract()

    def test_hash_and_parse_share_one_byte_snapshot(self):
        from unittest.mock import patch
        self.save()
        with patch.object(reader, 'ZipFile', wraps=ZipFile) as zipped:
            self.extract()
            self.assertIsInstance(zipped.call_args.args[0], io.BytesIO)

    def test_stale_saved_movement_total_is_held(self):
        self.material[6]['L'] = 5
        self.save()
        self.assertIn('MOVEMENT_TOTAL_MISMATCH', self.codes())

    def test_text_register_quantity_is_not_silently_ignored(self):
        self.material[7] = {'F': 'T1', 'L': '10 SET'}
        self.save()
        self.assertIn('MOVEMENT_INPUT_ERROR', self.codes())

    def test_bom_cached_totals_are_included_but_not_recalculated(self):
        self.bom[4] = {'D': 'Bolt', 'G': (2, 'some_native_formula()'), 'H': 0}
        self.stock[4]['P'] = (5, self.stock[4]['P'][1])
        self.stock[4]['R'] = (9, '(P4+O4-Q4)')
        self.save()
        self.assertEqual(self.codes(), [])

    def test_missing_bom_cached_result_is_held(self):
        self.bom[4] = {'D': 'Bolt', 'G': (None, 'some_native_formula()'), 'H': 0}
        self.save()
        self.assertIn('MOVEMENT_INPUT_ERROR', self.codes())

    def test_template_empty_rows_skipped_orphan_balance_preserved(self):
        self.stock[5] = {'O': None, 'P': (0, 'SUM(0)'), 'Q': (0, 'SUM(0)'), 'R': (0, '(P5+O5-Q5)')}
        self.stock[6] = {'R': 5}
        self.save()
        self.assertEqual([r['source_row'] for r in self.extract()['rows']], [4, 6])

    def test_changed_headers_fail_closed(self):
        self.stock[3]['R'] = 'Available to sell'
        self.save()
        with self.assertRaisesRegex(ValueError, 'layout needs review'):
            self.extract()

    def test_changed_bom_headers_fail_closed(self):
        self.bom[3]['G'] = 'Out'
        self.save()
        with self.assertRaisesRegex(ValueError, 'BOM Master headers'):
            self.extract()

    def test_formula_error_preserved(self):
        self.stock[4]['R'] = ('#REF!', '(P4+O4-Q4)', None, 'e')
        self.save()
        self.assertIn('CELL_ERROR', self.codes())

    def test_entities_rejected(self):
        self.save({'xl/sharedStrings.xml': '<!DOCTYPE foo [<!ENTITY x "data">]><sst/>'})
        with self.assertRaisesRegex(ValueError, 'entities'):
            self.extract()

    def test_duplicate_zip_paths_rejected(self):
        self.save()
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            with ZipFile(self.path, 'a') as z:
                z.writestr('xl/workbook.xml', '<workbook/>')
        with self.assertRaisesRegex(ValueError, 'ZIP directory'):
            self.extract()


if __name__ == '__main__':
    unittest.main()
