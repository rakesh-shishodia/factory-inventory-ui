"""Read the established Stock Sheet layout without Excel, macros or any writes.

Standard-library-only OOXML extraction. Saved formula results are preserved and
cross-checked, not recalculated. Unknown layouts/formulas fail closed for review.
"""
import argparse
from collections import defaultdict
from datetime import datetime, timezone
import hashlib
from io import BytesIO
import json
import math
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MAX_FILE = 30_000_000
MAX_PART = 40_000_000
MAX_EXPANDED = 150_000_000
MAX_ROW = 100_000
MAX_CELLS = 1_000_000
HEADERS = {'B': 'Status', 'C': 'Type', 'G': 'Store ID', 'I': 'SKU',
           'K': 'Unit', 'M': 'Stock Location', 'N': 'Min LVL',
           'O': 'Opening Balance', 'P': 'In', 'Q': 'Out', 'R': 'Balance'}
FIELDS = {'status': 'B', 'item_type': 'C', 'store_id': 'G', 'name': 'H',
          'sku': 'I', 'unit': 'K', 'location': 'M', 'minimum': 'N',
          'opening': 'O', 'inbound': 'P', 'outbound': 'Q', 'balance': 'R'}


def scalar_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def numeric_or_blank(value):
    if value is None or value == '':
        return 0
    return value if scalar_number(value) else None


def canonical_formula(value):
    # Spaces in quoted sheet names/literals are significant; only remove
    # formatting whitespace and absolute-reference markers outside them.
    return re.sub(r"'(?:[^']|'')*'|\"(?:[^\"]|\"\")*\"|\s+|\$",
                  lambda m: m[0] if m[0].startswith(("'", '"')) else '', value or '').upper()


class WorkbookReader:
    def __init__(self, archive):
        self.archive = archive
        entries = archive.infolist()
        if len(entries) > 5000 or len({e.filename for e in entries}) != len(entries):
            raise ValueError('Unsupported workbook ZIP directory.')
        if sum(e.file_size for e in entries) > MAX_EXPANDED or any(e.file_size > MAX_PART for e in entries):
            raise ValueError('Workbook expands beyond the supported read limit.')
        self.shared = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            self.shared = [self.string_text(s) for s in self.xml('xl/sharedStrings.xml').findall('s:si', NS)]
        self.sheets = {}
        rels = {r.get('Id'): r for r in self.xml('xl/_rels/workbook.xml.rels')}
        for sheet in self.xml('xl/workbook.xml').findall('s:sheets/s:sheet', NS):
            rel = rels.get(sheet.get('{' + REL_NS + '}id'))
            if rel is None or rel.get('TargetMode') == 'External':
                raise ValueError('Workbook contains an unsupported worksheet relationship.')
            target = rel.get('Target', '')
            path = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('xl', target))
            if not path.startswith('xl/') or path not in archive.namelist():
                raise ValueError('Invalid worksheet path.')
            name = sheet.get('name')
            if name in self.sheets:
                raise ValueError('Duplicate worksheet name.')
            self.sheets[name] = path

    def xml(self, path):
        content = self.archive.read(path)
        if b'<!DOCTYPE' in content.upper() or b'<!ENTITY' in content.upper():
            raise ValueError('XML document declarations/entities are not supported.')
        return ET.fromstring(content)

    @staticmethod
    def string_text(element):
        # Ignore phonetic annotations; retain rich-text runs in their source order.
        return ''.join(n.text or '' for n in element.findall('s:t', NS) + element.findall('s:r/s:t', NS))

    def sheet(self, name, columns):
        if name not in self.sheets:
            raise ValueError(f'Missing required worksheet: {name}.')
        root = self.xml(self.sheets[name])
        result = {}
        shared_formulas = {}
        count = 0
        for row in root.findall('s:sheetData/s:row', NS):
            row_id = int(row.get('r', '0'))
            if row_id < 1 or row_id > MAX_ROW or row_id in result:
                raise ValueError(f'Unsupported or duplicate row in {name}.')
            cells = {}
            for cell in row.findall('s:c', NS):
                count += 1
                if count > MAX_CELLS:
                    raise ValueError(f'Too many cells in {name}.')
                coord = re.fullmatch(r'([A-Z]+)([1-9][0-9]*)', cell.get('r', ''))
                if coord is None or int(coord[2]) != row_id:
                    raise ValueError(f'Invalid cell address in {name}.')
                col = coord[1]
                if col not in columns:
                    continue
                if col in cells:
                    raise ValueError(f'Duplicate cell in {name}: {col}{row_id}.')
                cell_type = cell.get('t', 'n')
                formula = cell.find('s:f', NS)
                if formula is not None and formula.get('t') == 'shared' and formula.text:
                    index = formula.get('si')
                    if index in shared_formulas:
                        raise ValueError('Duplicate shared formula master.')
                    shared_formulas[index] = (col, row_id, formula.text, formula.get('ref', ''))
                raw = cell.findtext('s:v', default=None, namespaces=NS)
                value = None
                if cell_type == 'inlineStr':
                    inline = cell.find('s:is', NS)
                    value = self.string_text(inline) if inline is not None else ''
                elif raw is not None:
                    if cell_type == 's':
                        index = int(raw)
                        if index < 0 or index >= len(self.shared):
                            raise ValueError('Invalid shared string index.')
                        value = self.shared[index]
                    elif cell_type == 'b':
                        value = raw == '1'
                    elif cell_type in ('str', 'e', 'd'):
                        value = raw
                    elif cell_type == 'n':
                        value = float(raw)
                        if not math.isfinite(value):
                            raise ValueError('Non-finite numeric cell.')
                        if value.is_integer():
                            value = int(value)
                    else:
                        raise ValueError(f'Unsupported cell type: {cell_type}.')
                cells[col] = {'value': value, 'type': cell_type,
                              'formula': '=' + (formula.text or '') if formula is not None else None,
                              'formula_type': formula.get('t', 'normal') if formula is not None else None,
                              'shared_index': formula.get('si') if formula is not None else None}
            result[row_id] = cells
        # Resolve only the exact known relative balance formula. Never guess at
        # shared formulas involving functions, other sheets or absolute refs.
        for row_id, cells in result.items():
            for col, cell in cells.items():
                if cell['formula_type'] != 'shared':
                    continue
                master = shared_formulas.get(cell['shared_index'])
                if master is None:
                    continue
                master_col, master_row, formula, scope = master
                bounds = re.fullmatch(r'R([1-9][0-9]*):R([1-9][0-9]*)', scope)
                exact = re.sub(r'\s+', '', formula).upper() == f'(P{master_row}+O{master_row}-Q{master_row})'
                if col == master_col == 'R' and exact and bounds and int(bounds[1]) <= row_id <= int(bounds[2]):
                    cell['formula'] = f'=(P{row_id}+O{row_id}-Q{row_id})'
                    cell['formula_type'] = 'normal'
        return result


def value(cells, col):
    return cells.get(col, {}).get('value')


def sum_register(rows, start, key_col, inward_col, outward_col):
    inward, outward = defaultdict(float), defaultdict(float)
    invalid = defaultdict(list)
    for row_id, cells in rows.items():
        if row_id < start:
            continue
        key_cell = cells.get(key_col, {})
        key = value(cells, key_col)
        valid_key = (isinstance(key, str) and bool(key.strip()) or scalar_number(key)) and key_cell.get('type') != 'e'
        if not valid_key:
            movement_bearing = any(
                value(cells, col) not in (None, '', 0) or cells.get(col, {}).get('type') == 'e' or
                (cells.get(col, {}).get('formula') and not scalar_number(value(cells, col)))
                for col in (inward_col, outward_col))
            if movement_bearing:
                raise ValueError(f'Unassigned register/BOM movement at {key_col}{row_id}: missing or invalid item key. Review before screening stock.')
            continue
        # Excel SUMIF is case-insensitive; surrounding spaces remain significant.
        key = str(key).upper()
        for column, totals in [(inward_col, inward), (outward_col, outward)]:
            cell = cells.get(column, {})
            amount = value(cells, column)
            if cell.get('type') == 'e' or (cell.get('formula') and not scalar_number(amount)):
                invalid[key].append(f'{column}{row_id}')
            elif scalar_number(amount):
                totals[key] += amount
            elif amount not in (None, ''):
                invalid[key].append(f'{column}{row_id}')
    return inward, outward, invalid


def extract(path, source_ref, source_modified_at=None):
    path = Path(path)
    if path.suffix.lower() not in ('.xlsx', '.xlsm') or path.stat().st_size > MAX_FILE:
        raise ValueError('Provide an .xlsx/.xlsm file no larger than 30 MB.')
    with path.open('rb') as stream:
        content = stream.read(MAX_FILE + 1)
    if len(content) > MAX_FILE:
        raise ValueError('Workbook exceeds the supported read limit.')
    digest = hashlib.sha256(content).hexdigest()
    with ZipFile(BytesIO(content)) as archive:
        reader = WorkbookReader(archive)
        stock = reader.sheet('Stock Sheet', set(FIELDS.values()))
        material = reader.sheet('Material Register', {'F', 'K', 'L'})
        bom = reader.sheet('BOM Master', {'D', 'G', 'H'})
    for col, label in HEADERS.items():
        if str(value(stock.get(3, {}), col) or '').strip().casefold() != label.casefold():
            raise ValueError(f'Expected {label!r} at Stock Sheet!{col}3; layout needs review.')
    if str(value(stock[3], 'H') or '').strip().casefold() not in ('', 'item name', 'name'):
        raise ValueError('Unexpected item-name column at Stock Sheet!H3.')
    if any(str(value(material.get(5, {}), c) or '').strip() != h for c, h in {'F': 'Store ID', 'K': 'Outward Qty', 'L': 'Inward Qty'}.items()):
        raise ValueError('Material Register headers changed; movement checks need review.')
    if any(str(value(bom.get(3, {}), c) or '').strip() != h for c, h in {'D': 'Item Name', 'G': 'In', 'H': 'Out'}.items()):
        raise ValueError('BOM Master headers changed; movement checks need review.')
    mat_in, mat_out, mat_errors = sum_register(material, 6, 'F', 'L', 'K')
    bom_in, bom_out, bom_errors = sum_register(bom, 4, 'D', 'G', 'H')
    records = []
    for row_id, cells in stock.items():
        if row_id <= 3:
            continue
        identities = [value(cells, col) for col in ('G', 'H', 'I')]
        nonzero = any(value(cells, c) not in (None, '', 0) for c in ('O', 'P', 'Q', 'R'))
        if not any(x not in (None, '') for x in identities) and not nonzero:
            continue  # Empty formatting/formula template rows, not stock records.
        record = {field: value(cells, col) for field, col in FIELDS.items()}
        record['source_row'] = row_id
        record['formulas'] = {field: cells.get(FIELDS[field], {}).get('formula') for field in ('opening', 'inbound', 'outbound', 'balance')}
        issues = []

        def issue(code, message):
            issues.append({'code': code, 'message': message})

        for field, col in FIELDS.items():
            cell = cells.get(col, {})
            if cell.get('type') == 'e':
                issue('CELL_ERROR', f'{col}{row_id} contains {cell.get("value")}.')
            if cell.get('formula') and cell.get('value') is None:
                issue('MISSING_FORMULA_RESULT', f'{col}{row_id} has no saved formula result.')
            if cell.get('formula_type') not in (None, 'normal'):
                issue('UNSUPPORTED_FORMULA', f'{col}{row_id} uses an unsupported formula representation.')
            if field in ('sku', 'store_id', 'name', 'unit', 'status', 'item_type') and cell.get('formula'):
                issue('DERIVED_IDENTITY', f'{col}{row_id} is formula-derived; confirm its value before mapping.')
        expected = {
            'balance': f'=(P{row_id}+O{row_id}-Q{row_id})',
            'inbound': f"=SUMIF('Material Register'!$F:$F,'Stock Sheet'!G{row_id},'Material Register'!$L:$L)+SUMIF('BOM Master'!$D:$D,'Stock Sheet'!H:H,'BOM Master'!$G:$G)",
            'outbound': f"=SUMIF('Material Register'!$F:$F,'Stock Sheet'!G{row_id},'Material Register'!$K:$K)+SUMIF('BOM Master'!$D:$D,'Stock Sheet'!H:H,'BOM Master'!$H:$H)",
        }
        for field, formula in expected.items():
            actual = record['formulas'][field]
            if not actual:
                issue('MANUAL_STOCK_VALUE', f'{FIELDS[field]}{row_id} is manually entered instead of the established {field} formula.')
            elif canonical_formula(actual) != canonical_formula(formula):
                issue('UNEXPECTED_STOCK_FORMULA', f'{FIELDS[field]}{row_id} differs from the expected row references; inspect the original formula.')
        parts = [numeric_or_blank(record[f]) for f in ('opening', 'inbound', 'outbound')]
        if all(v is not None for v in parts) and scalar_number(record['balance']):
            if abs(parts[0] + parts[1] - parts[2] - record['balance']) > 1e-7:
                issue('BALANCE_ARITHMETIC', f'R{row_id} does not equal Opening Balance + In − Out.')
        else:
            issue('UNVERIFIED_BALANCE_INPUT', f'Cannot verify the numeric inputs to R{row_id}.')
        sid, name = str(record['store_id'] or '').upper(), str(record['name'] or '').upper()
        if sid in mat_errors or name in bom_errors:
            refs = [f'Material Register!{c}' for c in mat_errors.get(sid, [])[:8]] + [f'BOM Master!{c}' for c in bom_errors.get(name, [])[:8]]
            issue('MOVEMENT_INPUT_ERROR', 'Invalid quantity or missing saved result at ' + ', '.join(refs) + '.')
        if any(ch in sid + name for ch in '*?~'):
            issue('WILDCARD_MATCH', 'SUMIF key contains wildcard characters; matching must be reviewed.')
        for field, total in [('inbound', mat_in[sid] + bom_in[name]), ('outbound', mat_out[sid] + bom_out[name])]:
            if scalar_number(record[field]) and abs(record[field] - total) > 1e-7:
                issue('MOVEMENT_TOTAL_MISMATCH', f'{FIELDS[field]}{row_id} does not match the saved register/BOM totals for this item.')
        record['audit_issues'] = issues
        records.append(record)
    if not records or len(records) > 10000:
        raise ValueError('Expected between 1 and 10,000 source stock rows.')
    source = {'file_name': path.name, 'sha256': digest, 'sheet_name': 'Stock Sheet', 'header_row': 3,
              'source_ref': source_ref, 'extracted_at': datetime.now(timezone.utc).isoformat()}
    if source_modified_at:
        source['source_modified_at'] = source_modified_at
    return {'schema_version': 1, 'source': source, 'rows': records}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook')
    parser.add_argument('--source-ref', required=True)
    parser.add_argument('--source-modified-at')
    args = parser.parse_args()
    try:
        result = extract(args.workbook, args.source_ref, args.source_modified_at)
        json.dump(result, sys.stdout, ensure_ascii=False, allow_nan=False)
        print()
    except (ValueError, OSError, KeyError, BadZipFile, ET.ParseError) as error:
        print(f'Stock workbook extraction failed: {error}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
