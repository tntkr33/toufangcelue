import io
import json
import time
import unittest
import zipfile
from xml.etree import ElementTree as ET

from server import XlsxBook

try:
    from customer_allocation import AllocationQuery, parse_customers
except ImportError:
    AllocationQuery = parse_customers = None


HEADERS = ['\u5ba2\u6237\u7f16\u7801', '\u5ba2\u6237\u540d\u79f0', '\u8425\u9500\u90e8', '\u6863\u4f4d',
           '\u5e02\u573a\u7c7b\u578b(\u6269\u5c55)', '\u5546\u5708\u7c7b\u578b(\u6269\u5c55)']
COUNTY = '\u5b9c\u5dde\u5e02\u8425\u9500\u90e8'
COUNTY_ALIAS = '\u5b9c\u5dde\u8425\u9500\u90e8'
OTHER = '\u90fd\u5b89\u53bf\u8425\u9500\u90e8'
CITY = '\u57ce\u5e02'
RURAL = '\u519c\u6751'
AREA = '\u8857\u9053\uff08\u8857\u533a\uff09'


def workbook(rows, headers=HEADERS):
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    root = ET.Element('worksheet', xmlns=ns)
    data = ET.SubElement(root, 'sheetData')
    for number, values in enumerate([['Synthetic title'], headers, *rows], 1):
        row = ET.SubElement(data, 'row', r=str(number))
        for col, value in enumerate(values):
            cell = ET.SubElement(row, 'c', r=f'{chr(65 + col)}{number}', t='inlineStr')
            ET.SubElement(ET.SubElement(cell, 'is'), 't').text = str(value)
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('xl/workbook.xml', f'<workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Synthetic" sheetId="1" r:id="rId1"/></sheets></workbook>')
        archive.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
        archive.writestr('xl/worksheets/sheet1.xml', ET.tostring(root))
    return output.getvalue()


def customer(code='001', tier=1, county=COUNTY, market=CITY, area=AREA):
    return [code, 'Synthetic customer', county, tier, market, area]


def reviewed(source=3, code='00001', county=COUNTY_ALIAS, market=CITY, area=AREA, quantities=None):
    return dict(sourceRow=source, productCode=code, productName='Synthetic product',
                county=county, marketType=market, businessArea=area,
                assignments={1: 7, 30: 0} if quantities is None else quantities,
                warnings=['Synthetic warning'])


class CustomerAllocationTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(parse_customers, 'Customer allocation engine must be implemented')

    def engine(self, rows=None, review=None, template='county', choices=None):
        dataset = parse_customers(workbook(rows or [customer()]), template, XlsxBook)
        review = review or [reviewed()]
        choices = choices if choices is not None else {str(r['sourceRow']): 'keep' for r in review}
        return AllocationQuery(dataset, {'reviewRows': review}, choices, template)

    def test_header_discovery_code_and_non_pii_metadata(self):
        data = parse_customers(workbook([customer()]), 'county', XlsxBook)
        self.assertEqual(data.summary['headerRow'], 2)
        self.assertEqual(data.summary['validCustomerCount'], 1)
        self.assertNotIn('Synthetic customer', json.dumps(data.summary))
        detail = self.engine().plan_query({}, 'detail').page(1, 20)
        self.assertEqual(detail['rows'][0]['customerCode'], '001')
        self.assertEqual(detail['rows'][0]['productCode'], '00001')

    def test_zip_expansion_and_entry_limits_before_reader(self):
        from unittest.mock import patch
        content = workbook([customer()])
        def forbidden_reader(content):
            self.fail('Unsafe archive reached workbook reader')
        with patch('customer_allocation.MAX_ZIP_ENTRIES', 2):
            with self.assertRaises(ValueError):
                parse_customers(content, 'county', forbidden_reader)
        with patch('customer_allocation.MAX_EXPANDED_BYTES', 10):
            with self.assertRaises(ValueError):
                parse_customers(content, 'county', forbidden_reader)

    def test_only_explicit_metadata_footers_are_skipped(self):
        headers = ['\u5e8f\u53f7', *HEADERS, '\u5907\u6ce8']
        footer = ['\u5236\u8868\u4eba\uff1aSynthetic', *([''] * 6), '\u5236\u8868\u65e5\u671f\uff1a2026-10-08']
        content = workbook([['1', *customer(), ''], footer], headers)
        dataset = parse_customers(content, 'county', XlsxBook)
        self.assertEqual(dataset.summary['footerRowCount'], 1)
        self.assertEqual(dataset.summary['sourceRowCount'], 1)
        self.assertEqual(dataset.summary['invalidCustomerCount'], 0)
        engine = AllocationQuery(dataset, {'reviewRows': [reviewed()]}, {'3': 'keep'}, 'county')
        self.assertFalse(engine.plan_query({}, 'summary').totals['incomplete'])
        for invalid in (
                ['Arbitrary content', *([''] * 7)],
                ['Prefix \u5236\u8868\u4eba\uff1aSynthetic', *([''] * 7)],
                ['\u5236\u8868\u4eba\uff1aSynthetic', *([''] * 6), 'Unexpected footer'],
                ['\u5236\u8868\u4eba\uff1aSynthetic', '', '', COUNTY, '', '', '', '']):
            with self.subTest(invalid=invalid[0]):
                dataset = parse_customers(workbook([invalid], headers), 'county', XlsxBook)
                self.assertEqual(dataset.summary['footerRowCount'], 0)
                self.assertEqual(dataset.summary['invalidCustomerCount'], 1)
    def test_all_templates_match_applied_assignments_and_zero(self):
        for template in ('county', 'district_type', 'county_business', 'market_type'):
            with self.subTest(template=template):
                engine = self.engine([customer(), customer('002', 30)], template=template)
                plan = engine.plan_query({}, 'summary')
                self.assertEqual(plan.totals, dict(productCount=1, customerCount=2,
                    matchedCustomerCount=2, unmatchedCustomerCount=0, quantityTiao=7,
                    quantityXiang=0.028, incomplete=False, detailRowCount=2))
                row = list(plan.iter_rows())[0]
                self.assertEqual((row['matchedCustomers'], row['nonzeroCustomers']), (2, 1))
                detail = engine.plan_query({}, 'detail')
                self.assertEqual(sum(r['quantityTiao'] for r in detail.iter_rows()), 7)
                self.assertEqual('businessArea' in {c['key'] for c in detail.columns}, template == 'county_business')

    def test_missing_required_columns_block_and_optional_dimensions_allowed(self):
        for template in ('county', 'district_type', 'county_business'):
            with self.assertRaises(ValueError):
                parse_customers(workbook([['001', 'Synthetic', 1]], [HEADERS[0], HEADERS[1], HEADERS[3]]), template, XlsxBook)
        data = parse_customers(workbook([['001', 'Synthetic', 1, CITY]],
            [HEADERS[0], HEADERS[1], HEADERS[3], HEADERS[4]]), 'market_type', XlsxBook)
        self.assertEqual(data.summary['validCustomerCount'], 1)
        wrong = HEADERS.copy()
        wrong[5] = '\u5546\u5708\u7c7b\u578b(\u56fd\u6807)'
        with self.assertRaises(ValueError):
            parse_customers(workbook([customer()], wrong), 'county_business', XlsxBook)

    def test_dedup_uses_measurements_not_customer_name(self):
        copy = customer()
        copy[1] = 'Different display name'
        engine = self.engine([customer(), copy])
        self.assertEqual(engine.plan_query({}, 'detail').row_count, 1)
        issues = list(engine.plan_query({}, 'issues').iter_rows())
        self.assertEqual([r['sourceRow'] for r in issues], [4])
        self.assertFalse(engine.plan_query({}, 'summary').totals['incomplete'])

    def test_conflicting_duplicates_excluded_and_possible_range_retained(self):
        engine = self.engine([customer(), customer(tier=30), customer('002', tier='?')])
        plan = engine.plan_query({'minTier': 30, 'maxTier': 30}, 'summary')
        self.assertEqual(plan.totals['customerCount'], 2)
        self.assertEqual(plan.totals['unmatchedCustomerCount'], 2)
        self.assertEqual(plan.totals['quantityTiao'], 0)
        self.assertTrue(plan.totals['incomplete'])
        self.assertEqual(engine.plan_query({}, 'detail').row_count, 0)
        self.assertEqual({r['sourceRow'] for r in engine.plan_query({}, 'issues').iter_rows()}, {3, 4, 5})

    def test_conflicting_customer_issues_export_only_filtered_source_rows(self):
        engine = self.engine([customer(), customer(county=OTHER, tier=30)])
        plan = engine.plan_query({'counties': [COUNTY], 'minTier': 1, 'maxTier': 1}, 'issues')
        self.assertEqual(plan.totals['unmatchedCustomerCount'], 1)
        self.assertTrue(plan.totals['incomplete'])
        self.assertEqual([r['sourceRow'] for r in plan.iter_rows()], [3])

    def test_invalid_and_unknown_business_queries_have_bounded_plan_memory(self):
        import tracemalloc
        review = [reviewed(i + 3, str(i).zfill(5)) for i in range(100)]
        for rows, template in (
            ([customer(str(i), tier='?') for i in range(1000)], 'county'),
            ([customer(str(i), area='Unknown area ' + str(i)) for i in range(1000)], 'county_business'),
        ):
            with self.subTest(template=template):
                engine = self.engine(rows, review, template)
                tracemalloc.start()
                try:
                    plan = engine.plan_query({}, 'issues')
                    peak = tracemalloc.get_traced_memory()[1]
                finally:
                    tracemalloc.stop()
                self.assertEqual(plan.row_count, 100000)
                self.assertLess(peak, 2 * 1024 * 1024)
                self.assertEqual(len(plan.page(10000, 10)['rows']), 10)

    def test_unknown_missing_and_no_code_rows_are_not_zero(self):
        rows = [customer(), customer('002', county='Unknown'), customer('', tier=''), customer('004', market='Unknown')]
        engine = self.engine(rows, template='district_type')
        totals = engine.plan_query({}, 'summary').totals
        self.assertEqual((totals['matchedCustomerCount'], totals['unmatchedCustomerCount']), (1, 3))
        self.assertTrue(totals['incomplete'])
        issues = list(engine.plan_query({'counties': [COUNTY]}, 'issues').iter_rows())
        self.assertEqual(len(issues), 3)
        self.assertTrue(all(r['reason'] and 'productCode' in r for r in issues))

    def test_issue_rows_preserve_unknown_dimension_and_reject_invented_alias(self):
        engine = self.engine([customer(county='Unknown county')])
        issue = next(engine.plan_query({}, 'issues').iter_rows())
        self.assertEqual(issue['county'], 'Unknown county')
        invented = '\u90fd\u5b89\u5e02\u8425\u9500\u90e8'
        engine = self.engine([customer(county=invented)], [reviewed(county=OTHER)])
        self.assertEqual(engine.plan_query({}, 'detail').row_count, 0)

    def test_all_real_county_values_and_extended_business_labels(self):
        from customer_allocation import COUNTY_NAMES
        self.assertEqual(len(COUNTY_NAMES), 11)
        rows = [customer(str(i), county=county) for i, county in enumerate(COUNTY_NAMES)]
        review = [reviewed(county=county) for county in COUNTY_NAMES]
        engine = self.engine(rows, review)
        self.assertEqual(set(engine.filter_options()['counties']), set(COUNTY_NAMES))
        self.assertEqual(engine.plan_query({}, 'detail').row_count, 11)
        for area in (AREA, '\u5c45\u6c11\u533a'):
            engine = self.engine([customer(area=area)], [reviewed(area=area)], 'county_business')
            self.assertEqual(engine.plan_query({}, 'detail').row_count, 1)

    def test_missing_assignment_ambiguous_dimension_and_duplicate_products(self):
        for review in ([reviewed(quantities={30: 0})], [reviewed(), reviewed()],
                       [reviewed(), reviewed(source=4)]):
            engine = self.engine(review=review)
            self.assertEqual(engine.plan_query({}, 'detail').row_count, 0)
            self.assertTrue(engine.plan_query({}, 'summary').totals['incomplete'])
            self.assertGreater(engine.plan_query({}, 'issues').row_count, 0)

    def test_omits_empty_lists_unique_totals_and_filters(self):
        engine = self.engine([customer(), customer('002', 30, OTHER, RURAL)],
                             [reviewed(), reviewed(4, '00002'), reviewed(5, '00003')],
                             choices={'3': 'keep', '4': 'keep', '5': 'omit'})
        self.assertEqual(len(engine.filter_options()['products']), 3)
        omitted = engine.plan_query({'productRows': [5]}, 'detail')
        self.assertEqual(omitted.row_count, 2)
        self.assertEqual(omitted.totals['quantityTiao'], 0)
        self.assertTrue(all(r['allocationStatus'] == '\u5df2\u6392\u9664\u6295\u653e'
                            for r in omitted.iter_rows()))
        self.assertEqual(set(engine.filter_options()['counties']), {COUNTY, OTHER})
        self.assertEqual(engine.plan_query({}, 'detail').totals['customerCount'], 2)
        for field in ('productRows', 'counties', 'markets'):
            plan = engine.plan_query({field: []}, 'summary')
            self.assertEqual(plan.row_count, 0)
            self.assertEqual(plan.totals['quantityTiao'], 0)
        plan = engine.plan_query({'productRows': [3], 'counties': [COUNTY], 'markets': [CITY], 'minTier': 1, 'maxTier': 1}, 'detail')
        self.assertEqual(plan.row_count, 1)
        self.assertEqual(plan.page(2, 1)['rows'], [])
        for filters in ({'minTier': 0}, {'maxTier': 31}, {'minTier': 30, 'maxTier': 1},
                        {'productRows': ['3']}, {'counties': 'x'}, {'markets': [1]},
                        {'minTier': True}, {'surprise': 1}):
            with self.assertRaises(ValueError):
                engine.plan_query(filters, 'detail')
        with self.assertRaises(ValueError):
            engine.plan_query({}, 'bad')
        for page, size in ((0, 10), (1, 0), (True, 1), (1, 201), (1, 4000000)):
            with self.assertRaises(ValueError):
                plan.page(page, size)

    def test_choices_exact_keys_and_snapshot_of_assignments(self):
        with self.assertRaises(ValueError):
            self.engine(choices={3: 'keep'})
        row = reviewed()
        engine = self.engine(review=[row])
        plan = engine.plan_query({}, 'detail')
        row['assignments'][1] = 999
        self.assertEqual(list(plan.iter_rows())[0]['quantityTiao'], 7)

    def test_business_area_is_validated_against_review_template(self):
        engine = self.engine([customer(area='Unknown')], template='county_business')
        self.assertEqual(engine.plan_query({}, 'detail').row_count, 0)
        self.assertTrue(engine.plan_query({}, 'summary').totals['incomplete'])

    def test_customer_issue_reasons_and_validation_are_chinese(self):
        engine = self.engine([customer(), customer('002', tier='?')],
                             [reviewed(quantities={30: 0})])
        for issue in engine.plan_query({}, 'issues').iter_rows():
            self.assertRegex(issue['reason'], '[\u4e00-\u9fff]')
        with self.assertRaisesRegex(ValueError, '[\u4e00-\u9fff]'):
            engine.plan_query({'minTier': 0}, 'detail')

    def test_grouped_large_summary_and_lazy_last_page(self):
        rows = [customer(str(i).zfill(6)) for i in range(20000)]
        review = [reviewed(i + 3, str(i).zfill(5)) for i in range(200)]
        engine = self.engine(rows, review)
        started = time.perf_counter()
        summary = engine.plan_query({}, 'summary')
        detail = engine.plan_query({}, 'detail')
        self.assertEqual(summary.totals['quantityTiao'], 28000000)
        self.assertEqual(detail.row_count, 4000000)
        self.assertEqual(detail.page(400000, 10)['rows'][-1]['customerCode'], '019999')
        self.assertLess(time.perf_counter() - started, 3)

    def test_irrelevant_business_area_does_not_split_county_groups(self):
        from unittest.mock import patch
        engine = self.engine([customer(str(i), area='Area ' + str(i)) for i in range(50)],
                             [reviewed(), reviewed(4, '00002')])
        with patch.object(engine, '_match', wraps=engine._match) as matching:
            self.assertEqual(engine.plan_query({}, 'summary').totals['quantityTiao'], 700)
            self.assertEqual(matching.call_count, 2)


if __name__ == '__main__':
    unittest.main()
