"""Customer validation and grouped allocation queries, independent of server state.

Only dataset.summary is upload metadata. Customer records stay in memory and are
emitted exclusively by explicitly requested detail/issue queries.
"""
import copy
import io
import math
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass


MAX_ZIP_ENTRIES = 4096
MAX_EXPANDED_BYTES = 256 * 1024 * 1024
MAX_CONTENT_BYTES = 64 * 1024 * 1024
REQUIRED = {
    'county': ('county',),
    'district_type': ('county', 'marketType'),
    'county_business': ('county', 'businessArea'),
    'market_type': ('marketType',),
}
HEADERS = {
    'orderCycle': ('\u8ba2\u8d27\u5468\u671f\u7c7b\u578b', '\u8ba2\u8d27\u5468\u671f'),
    'orderWeek': ('\u5468\u6b21', '\u8ba2\u8d27\u5468\u6b21'),
    'customerCode': ('\u5ba2\u6237\u7f16\u7801', '\u5ba2\u6237\u4ee3\u7801'),
    'customerName': ('\u5ba2\u6237\u540d\u79f0', '\u5ba2\u6237\u59d3\u540d'),
    'county': ('\u8425\u9500\u90e8', '\u6240\u5c5e\u8425\u9500\u90e8', '\u533a\u53bf', '\u53bf\u533a'),
    'tier': ('\u6863\u4f4d', '\u5ba2\u6237\u6863\u4f4d', '\u5ba2\u6237\u6863\u6b21'),
    'marketType': ('\u5e02\u573a\u7c7b\u578b(\u6269\u5c55)',),
    'businessArea': ('\u5546\u5708\u7c7b\u578b(\u6269\u5c55)',),
}
COUNTY_NAMES = (
    '\u5b9c\u5dde\u5e02\u8425\u9500\u90e8', '\u90fd\u5b89\u53bf\u8425\u9500\u90e8',
    '\u57ce\u533a\u8425\u9500\u90e8', '\u73af\u6c5f\u8425\u9500\u90e8',
    '\u5357\u4e39\u53bf\u8425\u9500\u90e8', '\u7f57\u57ce\u8425\u9500\u90e8',
    '\u5927\u5316\u53bf\u8425\u9500\u90e8', '\u5df4\u9a6c\u8425\u9500\u90e8',
    '\u4e1c\u5170\u8425\u9500\u90e8', '\u51e4\u5c71\u53bf\u8425\u9500\u90e8',
    '\u5929\u5ce8\u53bf\u8425\u9500\u90e8',
)
COUNTY_ROOTS = ('\u5b9c\u5dde', '\u90fd\u5b89', '\u57ce\u533a', '\u73af\u6c5f',
                '\u5357\u4e39', '\u7f57\u57ce', '\u5927\u5316', '\u5df4\u9a6c',
                '\u4e1c\u5170', '\u51e4\u5c71', '\u5929\u5ce8')
COUNTY_ALIASES = {}
for _root, _name in zip(COUNTY_ROOTS, COUNTY_NAMES):
    for _alias in (_root, _root + '\u8425\u9500\u90e8', _name):
        COUNTY_ALIASES[_alias] = _name
for _area, _name in zip(('\u5b9c\u5dde\u5e02', '\u90fd\u5b89\u53bf', '\u91d1\u57ce\u6c5f\u533a',
                         '\u73af\u6c5f\u53bf', '\u5357\u4e39\u53bf', '\u7f57\u57ce\u53bf',
                         '\u5927\u5316\u53bf', '\u5df4\u9a6c\u53bf', '\u4e1c\u5170\u53bf',
                         '\u51e4\u5c71\u53bf', '\u5929\u5ce8\u53bf'), COUNTY_NAMES):
    COUNTY_ALIASES[_area] = _name
    COUNTY_ALIASES[_area + '\u8425\u9500\u90e8'] = _name
for _alias in ('\u91d1\u57ce\u6c5f', '\u91d1\u57ce\u6c5f\u533a', '\u91d1\u57ce\u6c5f\u8425\u9500\u90e8'):
    COUNTY_ALIASES[_alias] = COUNTY_NAMES[2]
MARKET_ALIASES = {v: v for v in ('\u57ce\u5e02', '\u53bf\u57ce', '\u4e61\u9547', '\u519c\u6751')}
MARKET_ALIASES.update({'\u5e02\u533a': '\u57ce\u5e02', '\u57ce\u533a': '\u57ce\u5e02', '\u9547\u533a': '\u4e61\u9547'})
DIMENSIONS = ('county', 'tier', 'marketType', 'businessArea')


def _text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalized(value):
    return re.sub(r'\s+', '', _text(value)).replace('\uff08', '(').replace('\uff09', ')')


def _dimension(key, value):
    value = _normalized(value)
    if key == 'county':
        return COUNTY_ALIASES.get(value)
    if key == 'marketType':
        return MARKET_ALIASES.get(value)
    if key == 'tier':
        match = re.fullmatch(r'(\d{1,2})(?:\.0+)?(?:\u6863)?', value)
        return int(match[1]) if match and 1 <= int(match[1]) <= 30 else None
    return value or None


def _cycle(record):
    value = _normalized(record.get('orderCycle'))
    if value in ('\u4e00\u5468\u4e00\u6b21', '\u4e00\u5468\u4e00\u8ba2', '\u6bcf\u5468\u4e00\u8ba2'):
        return 'weekly', None, ''
    if value in ('\u6bcf\u6708\u4e00\u8ba2', '\u6bcf\u6708\u4e00\u6b21', '\u4e00\u6708\u4e00\u6b21', '\u56db\u5468\u4e00\u6b21'):
        return 'monthly', None, ''
    if value in ('\u4e24\u5468\u4e00\u6b21', '\u4e24\u5468\u4e00\u8ba2'):
        week = {'\u7b2c\u4e00\u5468': 'odd', '\u7b2c\u4e8c\u5468': 'even'}.get(_normalized(record.get('orderWeek')))
        return 'biweekly', week, '' if week else '\u4e24\u5468\u4e00\u8ba2\u7684\u5468\u6b21\u7f3a\u5931\u6216\u65e0\u6548\uff0c\u63090\u6761\u8ba1\u7b97\uff0c\u8bf7\u590d\u6838'
    return None, None, '\u8ba2\u8d27\u5468\u671f\u7f3a\u5931\u6216\u65e0\u6cd5\u8bc6\u522b\uff0c\u63090\u6761\u8ba1\u7b97\uff0c\u8bf7\u590d\u6838'


def _check_archive(content):
    if not isinstance(content, bytes) or len(content) > MAX_CONTENT_BYTES:
        raise ValueError('\u5ba2\u6237\u5de5\u4f5c\u7c3f\u65e0\u6548\u6216\u8d85\u8fc7\u4e0a\u4f20\u5927\u5c0f\u9650\u5236\u3002')
    if content.startswith(b'\xd0\xcf\x11\xe0'):
        return
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            entries = archive.infolist()
            if (len(entries) > MAX_ZIP_ENTRIES
                    or sum(e.file_size for e in entries) > MAX_EXPANDED_BYTES
                    or len({e.filename for e in entries}) != len(entries)
                    or any(e.flag_bits & 1 for e in entries)):
                raise ValueError('\u5ba2\u6237\u5de5\u4f5c\u7c3f\u89e3\u538b\u5927\u5c0f\u6216\u6587\u4ef6\u6570\u91cf\u8d85\u8fc7\u5b89\u5168\u9650\u5236\u3002')
    except (zipfile.BadZipFile, OSError):
        raise ValueError('\u5ba2\u6237\u6587\u4ef6\u4e0d\u662f\u53ef\u8bfb\u53d6\u7684 Excel \u5de5\u4f5c\u7c3f\u3002') from None


def _sheets(book):
    if callable(book.sheets):
        for sheet in book.sheets():
            yield ((r + 1, {c + 1: sheet.cell_value(r, c) for c in range(sheet.ncols)})
                   for r in range(sheet.nrows))
    else:
        for sheet in book.sheets:
            yield ((int(row.attrib.get('r', index)),
                    {col: book.cell_text(cell) for col, cell in book.row_cells(row).items()})
                   for index, row in enumerate(book.rows(sheet), 1))


@dataclass(frozen=True)
class CustomerDataset:
    template_key: str
    customers: tuple
    invalid: tuple
    duplicates: tuple
    summary: dict


def parse_customers(content, template_key, book_factory):
    """Read XLSX through book_factory(content), or XLS through xlrd.

    Raises ValueError for unsupported templates, unsafe workbooks, absent or
    ambiguous required headers. summary contains counts only plus headerRow.
    """
    if template_key not in REQUIRED:
        raise ValueError('\u4e0d\u652f\u6301\u7684\u5ba2\u6237\u6295\u653e\u6a21\u677f\u3002')
    _check_archive(content)
    required = ('customerCode', 'customerName', 'tier', *REQUIRED[template_key])
    book = None
    try:
        if content.startswith(b'\xd0\xcf\x11\xe0'):
            import xlrd
            book = xlrd.open_workbook(file_contents=content)
        else:
            book = book_factory(content)
        found = None
        best = set()
        lookup = {_normalized(alias): key for key, aliases in HEADERS.items() for alias in aliases}
        for sheet_rows in _sheets(book):
            for number, cells in sheet_rows:
                columns = {}
                repeated = set()
                for col, value in cells.items():
                    key = lookup.get(_normalized(value))
                    if key:
                        if key in columns:
                            repeated.add(key)
                        columns[key] = col
                if len(columns) > len(best):
                    best = set(columns)
                if set(required) <= columns.keys():
                    if repeated:
                        raise ValueError('\u5ba2\u6237\u5de5\u4f5c\u7c3f\u5b58\u5728\u91cd\u590d\u5b57\u6bb5\u5217\uff0c\u8bf7\u4fdd\u7559\u552f\u4e00\u8868\u5934\u3002')
                    found = (number, columns, sheet_rows)
                    break
            if found:
                break
        if not found:
            missing = [HEADERS[k][0] for k in required if k not in best]
            raise ValueError('\u7f3a\u5c11\u5ba2\u6237\u5b57\u6bb5\u5217\uff1a' + '\u3001'.join(missing))
        header_row, columns, rows = found
        by_code = defaultdict(list)
        source_count = 0
        footer_count = 0
        for number, cells in rows:
            if not any(_text(v) for v in cells.values()):
                continue
            # Only explicit metadata outside all recognized customer columns is
            # exempt from the nonempty-row validation rule.
            if (not any(_text(cells.get(col)) for col in columns.values())
                    and all(re.fullmatch(r'(?:\u5236\u8868\u4eba|\u5236\u8868\u65e5\u671f)\s*[:\uff1a]\s*.+', _text(value))
                            for value in cells.values() if _text(value))):
                footer_count += 1
                continue
            source_count += 1
            record = {k: _text(cells.get(col)) for k, col in columns.items()}
            record['sourceRow'] = number
            record.setdefault('customerName', '')
            record.setdefault('customerCode', '')
            record['orderCycle'] = record.get('orderCycle', '')
            record['orderWeek'] = record.get('orderWeek', '')
            record['_cycle'], record['_week'], record['_cycleReason'] = _cycle(record)
            errors = []
            raw = {}
            for key in DIMENSIONS:
                raw[key] = _text(record.get(key))
                record[key] = _dimension(key, record.get(key))
                if key in ('tier', *REQUIRED[template_key]) and record[key] is None:
                    errors.append(LABELS[key] + '\u7f3a\u5931\u6216\u65e0\u6cd5\u8bc6\u522b')
            if not record['customerCode']:
                errors.append('\u7f3a\u5c11\u5ba2\u6237\u7f16\u7801')
            if not record['customerName']:
                errors.append('\u7f3a\u5c11\u5ba2\u6237\u540d\u79f0')
            record['_raw'] = raw
            record['_reason'] = '\uff1b'.join(errors)
            identity = ('code', record['customerCode']) if record['customerCode'] else ('row', number)
            by_code[identity].append(record)
        customers, invalid, duplicates = [], [], []
        for records in by_code.values():
            signatures = {tuple(r[k] if r[k] is not None else _normalized(r['_raw'][k]) for k in DIMENSIONS)
                          + (r['_cycle'], r['_week'], r['_cycleReason'],
                             _normalized(r['orderCycle']) if r['_cycleReason'] else '',
                             _normalized(r['orderWeek']) if r['_cycleReason'] else '')
                          for r in records}
            if len(signatures) > 1:
                invalid.append(tuple(dict(r, _reason='\u540c\u4e00\u5ba2\u6237\u7f16\u7801\u7684\u6d4b\u7b97\u5b57\u6bb5\u51b2\u7a81\uff0c\u5df2\u6392\u9664') for r in records))
            elif any(r['_reason'] for r in records):
                invalid.append(tuple(records))
            else:
                customers.append(records[0])
                duplicates.extend(dict(r, _reason='\u91cd\u590d\u5ba2\u6237\u884c\u5df2\u5408\u5e76\uff0c\u4ec5\u8ba1\u7b97\u4e00\u6b21', _warning=True) for r in records[1:])
        return CustomerDataset(template_key, tuple(customers), tuple(invalid), tuple(duplicates), {
            'headerRow': header_row, 'sourceRowCount': source_count, 'footerRowCount': footer_count,
            'customerCount': len(by_code), 'validCustomerCount': len(customers),
            'invalidCustomerCount': len(invalid), 'duplicateRowCount': len(duplicates),
            'issueCount': sum(map(len, invalid)) + len(duplicates),
        })
    except ValueError:
        raise
    except Exception:
        # Reader exceptions may include cell contents; do not surface those.
        raise ValueError('\u5ba2\u6237\u5de5\u4f5c\u7c3f\u8bfb\u53d6\u5931\u8d25\uff0c\u8bf7\u68c0\u67e5\u6587\u4ef6\u683c\u5f0f\u3002') from None
    finally:
        if book is not None:
            if hasattr(book, 'zf'):
                book.zf.close()
            if hasattr(book, 'release_resources'):
                book.release_resources()


LABELS = {
    'orderCycle': '\u8ba2\u8d27\u5468\u671f', 'orderWeek': '\u5468\u6b21',
    'allocationStatus': '\u6295\u653e\u72b6\u6001',
    'productCode': '\u5546\u54c1\u7f16\u7801', 'productName': '\u5546\u54c1\u540d\u79f0',
    'matchedCustomers': '\u5339\u914d\u5ba2\u6237\u6570', 'nonzeroCustomers': '\u975e\u96f6\u6295\u653e\u5ba2\u6237\u6570',
    'unmatchedCustomers': '\u672a\u5339\u914d\u5ba2\u6237\u6570',
    'quantityTiao': '\u6d4b\u7b97\u6295\u653e\u91cf(\u6761)', 'quantityXiang': '\u6d4b\u7b97\u6295\u653e\u91cf(\u7bb1)',
    'customerCode': '\u5ba2\u6237\u7f16\u7801', 'customerName': '\u5ba2\u6237\u540d\u79f0',
    'county': '\u8425\u9500\u90e8', 'tier': '\u6863\u4f4d',
    'marketType': '\u5e02\u573a\u7c7b\u578b(\u6269\u5c55)', 'businessArea': '\u5339\u914d\u5546\u5708',
    'sourceRow': '\u5ba2\u6237\u539f\u8868\u884c\u53f7', 'reason': '\u539f\u56e0',
}
SUMMARY_KEYS = ('productCode', 'productName', 'matchedCustomers', 'nonzeroCustomers',
                'unmatchedCustomers', 'quantityTiao', 'quantityXiang')
DETAIL_KEYS = ('productCode', 'productName', 'customerCode', 'customerName', 'county', 'tier', 'marketType')


def _filters(filters):
    if not isinstance(filters, dict) or set(filters) - {'productRows', 'counties', 'markets', 'minTier', 'maxTier', 'weekParity', 'monthlyWeek'}:
        raise ValueError('\u67e5\u8be2\u7b5b\u9009\u6761\u4ef6\u683c\u5f0f\u65e0\u6548\u3002')
    result = {}
    # Legacy in-process callers may omit the schedule; the HTTP service requires it.
    result['weekParity'] = filters.get('weekParity')
    result['monthlyWeek'] = filters.get('monthlyWeek')
    if ('weekParity' in filters or 'monthlyWeek' in filters) and (
            result['weekParity'] not in ('odd', 'even') or type(result['monthlyWeek']) is not bool):
        raise ValueError('\u8bf7\u9009\u62e9\u5355\u5468\u6216\u53cc\u5468\uff0c\u5e76\u786e\u8ba4\u662f\u5426\u4e3a\u6bcf\u6708\u4e00\u8ba2\u5468\u6b21\u3002')
    for key, kind in (('productRows', int), ('counties', str), ('markets', str)):
        values = filters.get(key)
        if values is not None and (not isinstance(values, list) or any(type(v) is not kind for v in values)):
            raise ValueError('\u7b5b\u9009\u5217\u8868\u683c\u5f0f\u65e0\u6548\uff1a' + {'productRows': '\u54c1\u89c4', 'counties': '\u8425\u9500\u90e8', 'markets': '\u5e02\u573a\u7c7b\u578b'}[key])
        result[key] = None if values is None else frozenset(values)
    for key, default in (('minTier', 1), ('maxTier', 30)):
        value = filters.get(key, default)
        if type(value) is not int or not 1 <= value <= 30:
            raise ValueError('\u8d77\u6b62\u6863\u4f4d\u5fc5\u987b\u4e3a 1 \u81f3 30 \u7684\u6574\u6570\u3002')
        result[key] = value
    if result['minTier'] > result['maxTier']:
        raise ValueError('\u8d77\u59cb\u6863\u4f4d\u4e0d\u80fd\u5927\u4e8e\u7ed3\u675f\u6863\u4f4d\u3002')
    return result


def _selected(record, filters, possible=False):
    for key, field in (('counties', 'county'), ('markets', 'marketType')):
        selected = filters[key]
        value = record[field]
        if selected is not None:
            if not selected or (not (possible and value is None) and value not in selected):
                return False
    tier = record['tier']
    return (possible and tier is None) or (tier is not None and filters['minTier'] <= tier <= filters['maxTier'])


def _cycle_eligible(record, filters):
    return (not record['_cycleReason']
            and not (record['_cycle'] == 'biweekly' and record['_week'] != filters['weekParity'])
            and not (record['_cycle'] == 'monthly' and not filters['monthlyWeek']))


class AllocationQuery:
    """Snapshot reviewed assignments; aggregate by full customer dimensions.

    choices must use exactly the confirmed product source-row strings. A customer
    may appear in both matchedCustomerCount and unmatchedCustomerCount when only
    some selected products match. Neither count counts product/customer pairs.
    """
    def __init__(self, dataset, review, choices, template_key):
        if template_key not in REQUIRED or dataset.template_key != template_key:
            raise ValueError('\u6a21\u677f\u5df2\u53d8\u5316\uff0c\u8bf7\u91cd\u65b0\u4e0a\u4f20\u5ba2\u6237\u8868\u3002')
        self.dataset = dataset
        self.template_key = template_key
        self.required = REQUIRED[template_key]
        review_rows = copy.deepcopy(review.get('reviewRows', []))
        by_source = defaultdict(list)
        for row in review_rows:
            if type(row.get('sourceRow')) is not int:
                raise ValueError('\u54c1\u89c4\u539f\u8868\u884c\u53f7\u65e0\u6548\uff0c\u8bf7\u91cd\u65b0\u89e3\u6790\u7b56\u7565\u3002')
            by_source[row['sourceRow']].append(row)
        products = review.get('exportProducts', []) or review.get('rules', [])
        expected = {str(p['sourceRow']) for p in products} if products else {str(s) for s in by_source}
        if (not isinstance(choices, dict) or set(choices) != expected
                or any(type(k) is not str or v not in ('keep', 'omit') for k, v in choices.items())):
            raise ValueError('\u54c1\u89c4\u786e\u8ba4\u6e05\u5355\u4e0d\u5b8c\u6574\u6216\u5df2\u53d8\u5316\uff0c\u8bf7\u91cd\u65b0\u786e\u8ba4\u3002')
        catalog = {p['sourceRow']: p for p in products}
        self.products = []
        self.index = {}
        codes = defaultdict(list)
        for source in sorted(int(s) for s in choices):
            rows = by_source.get(source, [])
            first = rows[0] if rows else catalog.get(source, {})
            product = {'sourceRow': source, 'productCode': _text(first.get('productCode', first.get('code'))),
                       'productName': _text(first.get('productName', first.get('name'))),
                       'allocationStatus': '\u5df2\u6392\u9664\u6295\u653e' if choices[str(source)] == 'omit' else '\u53c2\u4e0e\u6295\u653e'}
            self.products.append(product)
            codes[product['productCode']].append(source)
            index = defaultdict(list)
            for row in rows:
                key = tuple(_dimension(k, row.get(k)) for k in self.required)
                if None not in key:
                    index[key].append(row.get('assignments', {}))
            self.index[source] = index
        self.conflicting = {s for code, sources in codes.items() for s in sources if not code or len(sources) > 1}
        self.groups = defaultdict(list)
        group_dimensions = DIMENSIONS if 'businessArea' in self.required else DIMENSIONS[:-1]
        group_dimensions += ('_cycle', '_week', '_cycleReason')
        for record in dataset.customers:
            self.groups[tuple(record[k] for k in group_dimensions)].append(record)
        self.groups = tuple(tuple(records) for records in self.groups.values())

    def filter_options(self):
        records = [*self.dataset.customers, *(r for group in self.dataset.invalid for r in group)]
        return {'products': copy.deepcopy(self.products),
                'counties': sorted({r['county'] for r in records if r['county']}),
                'markets': sorted({r['marketType'] for r in records if r['marketType']})}

    def _match(self, product, record, filters=None):
        if filters and filters['weekParity'] is not None:
            if record['_cycleReason']:
                return 0, record['_cycleReason']
            if not _cycle_eligible(record, filters):
                return 0, None
        if product['allocationStatus'] == '\u5df2\u6392\u9664\u6295\u653e':
            return 0, None
        source = product['sourceRow']
        if source in self.conflicting:
            return None, '\u5546\u54c1\u7f16\u7801\u7f3a\u5931\u6216\u5b58\u5728\u591a\u6761\u54c1\u89c4\u7b56\u7565\u51b2\u7a81'
        key = tuple(record[k] for k in self.required)
        candidates = self.index[source].get(key, ())
        if not candidates:
            return None, '\u5ba2\u6237\u7ef4\u5ea6\u5728\u5f53\u524d\u6a21\u677f\u4e2d\u65e0\u5339\u914d\u9879'
        if len(candidates) != 1:
            return None, '\u5ba2\u6237\u7ef4\u5ea6\u5339\u914d\u5230\u591a\u6761\u6a21\u677f\u8bb0\u5f55\uff0c\u65e0\u6cd5\u552f\u4e00\u786e\u5b9a'
        assignments = candidates[0]
        tier = record['tier']
        value = assignments.get(tier, assignments.get(str(tier)))
        if type(value) not in (int, float) or not math.isfinite(value) or value < 0:
            return None, '\u5df2\u5e94\u7528\u7b56\u7565\u7f3a\u5c11\u8be5\u6863\u4f4d\u6570\u91cf\u6216\u6570\u91cf\u65e0\u6548'
        return value, None

    def plan_query(self, filters, view):
        if view not in ('summary', 'detail', 'issues'):
            raise ValueError('\u4e0d\u652f\u6301\u7684\u5ba2\u6237\u6295\u653e\u67e5\u8be2\u89c6\u56fe\u3002')
        filters = _filters(filters)
        products = [p for p in self.products if filters['productRows'] is None or p['sourceRow'] in filters['productRows']]
        groups = [g for g in self.groups if _selected(g[0], filters)]
        invalid = [selected for g in self.dataset.invalid
                   if (selected := tuple(r for r in g if _selected(r, filters, True)))]
        invalid_rows = tuple(r for group in invalid for r in group)
        duplicates = tuple(r for r in self.dataset.duplicates if _selected(r, filters))
        # Keep one count per product; even invalid cross-products are generated lazily.
        sources, summaries = [], []
        matched_groups, unmatched_groups = set(), set()
        detail_count = 0
        quantity = 0
        for product in products:
            matched = nonzero = 0
            unmatched = len(invalid)
            issue_rows = len(invalid_rows) + len(duplicates)
            subtotal = 0
            for group_no, group in enumerate(groups):
                value, reason = self._match(product, group[0], filters)
                if value is not None:
                    matched_groups.add(group_no)
                    matched += len(group)
                    nonzero += len(group) if value else 0
                    subtotal += value * len(group)
                else:
                    unmatched_groups.add(group_no)
                    unmatched += len(group)
                if reason is not None:
                    issue_rows += len(group)
            sources.append((product, matched if view == 'detail' else issue_rows))
            if matched or unmatched:
                summaries.append(dict(product, matchedCustomers=matched, nonzeroCustomers=nonzero,
                    unmatchedCustomers=unmatched, quantityTiao=subtotal, quantityXiang=round(subtotal / 250, 3)))
            detail_count += matched
            quantity += subtotal
        totals = {'productCount': len(products),
                  'customerCount': sum(map(len, groups)) + len(invalid),
                  'matchedCustomerCount': sum(len(groups[i]) for i in matched_groups),
                  'unmatchedCustomerCount': (sum(len(groups[i]) for i in unmatched_groups) + len(invalid)) if products else 0,
                  'quantityTiao': quantity, 'quantityXiang': round(quantity / 250, 3),
                  'incomplete': bool(products and (invalid or unmatched_groups)), 'detailRowCount': detail_count}
        if filters['weekParity'] is not None:
            totals['eligibleCustomerCount'] = sum(len(g) for g in groups if _cycle_eligible(g[0], filters))
            totals['cycleReviewCustomerCount'] = sum(len(g) for g in groups if g[0]['_cycleReason']) if products else 0
        return QueryPlan(view, self.template_key, totals, tuple(summaries), tuple(sources),
                         self, tuple(groups), invalid_rows, duplicates, filters)


class QueryPlan:
    def __init__(self, view, template_key, totals, summaries, sources, engine, groups, invalid, duplicates, filters):
        self.view = view
        self.totals = totals
        self._summaries = summaries
        self._sources = sources
        self._engine = engine
        self._groups = groups
        self._invalid = invalid
        self._duplicates = duplicates
        self._filters = filters
        keys = list(SUMMARY_KEYS if view == 'summary' else DETAIL_KEYS)
        keys.insert(2, 'allocationStatus')
        if view != 'summary':
            keys.extend(('orderCycle', 'orderWeek'))
            if template_key == 'county_business':
                keys.append('businessArea')
            keys.extend(('sourceRow', 'reason') if view == 'issues' else ('quantityTiao', 'quantityXiang'))
        self.columns = [{'key': k, 'label': LABELS[k]} for k in keys]
        self._keys = tuple(keys)
        self.row_count = len(summaries) if view == 'summary' else sum(count for _, count in sources)

    def _parts(self, product):
        for group in self._groups:
            value, reason = self._engine._match(product, group[0], self._filters)
            if (self.view == 'detail' and value is not None) or (self.view == 'issues' and reason is not None):
                yield group, value, reason
        if self.view == 'issues':
            for records in (self._invalid, self._duplicates):
                if records:
                    yield records, None, None

    def _rows(self, start, stop):
        if self.view == 'summary':
            for row in self._summaries[start:stop]:
                yield {k: row[k] for k in self._keys}
            return
        offset = 0
        for product, count in self._sources:
            if offset >= stop:
                return
            if offset + count <= start:
                offset += count
                continue
            for customers, value, reason in self._parts(product):
                end = offset + len(customers)
                if end > start and offset < stop:
                    for index in range(max(0, start - offset), min(len(customers), stop - offset)):
                        record = customers[index]
                        row = {k: record.get(k) for k in self._keys}
                        row.update(productCode=product['productCode'], productName=product['productName'])
                        row['allocationStatus'] = product['allocationStatus']
                        if self.view == 'detail':
                            row.update(quantityTiao=value, quantityXiang=round(value / 250, 3))
                        else:
                            for key in DIMENSIONS:
                                if key in row and row[key] is None:
                                    row[key] = record['_raw'][key]
                            row['reason'] = reason or record['_reason']
                        yield row
                offset = end
                if offset >= stop:
                    return

    def iter_rows(self):
        return self._rows(0, self.row_count)

    def page(self, page, page_size):
        if type(page) is not int or type(page_size) is not int or page < 1 or not 1 <= page_size <= 200:
            raise ValueError('\u9875\u7801\u5fc5\u987b\u4e3a\u6b63\u6574\u6570\uff0c\u6bcf\u9875\u6570\u91cf\u5fc5\u987b\u4e3a 1 \u81f3 200 \u7684\u6574\u6570\u3002')
        start = (page - 1) * page_size
        return {'rows': list(self._rows(start, min(start + page_size, self.row_count))),
                'totalRows': self.row_count, 'page': page, 'pageSize': page_size}
