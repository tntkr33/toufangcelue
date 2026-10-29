import unittest
from customer_allocation import AllocationQuery, parse_customers
from server import XlsxBook
from test_customer_allocation import HEADERS, workbook, customer, reviewed


CYCLE_HEADERS = HEADERS + ['\u8ba2\u8d27\u5468\u671f\u7c7b\u578b', '\u5468\u6b21']
WEEKLY = '\u4e00\u5468\u4e00\u6b21'
BIWEEKLY = '\u4e24\u5468\u4e00\u6b21'
MONTHLY = '\u6bcf\u6708\u4e00\u8ba2'
FIRST, SECOND = '\u7b2c\u4e00\u5468', '\u7b2c\u4e8c\u5468'


class CycleTests(unittest.TestCase):
    def test_eligible_count_is_unique_cycle_based_and_keeps_zero_details(self):
        rows = [customer('001') + [WEEKLY, ''], customer('001') + [WEEKLY, ''],
                customer('002') + [BIWEEKLY, FIRST], customer('003') + [BIWEEKLY, SECOND],
                customer('004') + [BIWEEKLY, SECOND], customer('005') + [MONTHLY, ''],
                customer('006') + ['', '']]
        engine = self.engine(rows)
        for parity, monthly, expected in [('odd', False, 2), ('even', False, 3), ('even', True, 4)]:
            for view in ('summary', 'detail', 'issues'):
                plan = engine.plan_query({'weekParity': parity, 'monthlyWeek': monthly}, view)
                self.assertEqual(plan.totals['eligibleCustomerCount'], expected)
                self.assertEqual(plan.totals['customerCount'], 6)
                self.assertEqual(plan.totals['detailRowCount'], 6)
        filtered = engine.plan_query({'weekParity': 'odd', 'monthlyWeek': False, 'minTier': 2}, 'detail')
        self.assertEqual(filtered.totals['eligibleCustomerCount'], 0)

    def test_four_weeks_means_monthly_regardless_of_week_field(self):
        engine = self.engine([customer('001') + ['\u56db\u5468\u4e00\u6b21', SECOND]])
        for parity in ('odd', 'even'):
            for monthly in (False, True):
                filters = {'weekParity': parity, 'monthlyWeek': monthly}
                plan = engine.plan_query(filters, 'detail')
                self.assertEqual(plan.totals['quantityTiao'], 7 if monthly else 0)
                self.assertEqual(plan.totals['cycleReviewCustomerCount'], 0)
                self.assertEqual(plan.row_count, 1)
                self.assertEqual(engine.plan_query(filters, 'issues').row_count, 0)

    def test_invalid_schedule_and_frozen_export_rows(self):
        import tempfile
        from pathlib import Path
        from query_workbook import write_query_workbook
        engine = self.engine([customer('001') + [BIWEEKLY, FIRST],
                              customer('002') + ['', '']])
        for filters in ({'weekParity': 'bad', 'monthlyWeek': False},
                        {'weekParity': 'odd', 'monthlyWeek': 'false'},
                        {'weekParity': 'odd'}):
            with self.assertRaises(ValueError):
                engine.plan_query(filters, 'summary')
        plan = engine.plan_query({'weekParity': 'even', 'monthlyWeek': False}, 'detail')
        engine.plan_query({'weekParity': 'odd', 'monthlyWeek': True}, 'detail')
        self.assertEqual([r['quantityTiao'] for r in plan.iter_rows()], [0, 0])
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'cycle.xlsx'
            write_query_workbook(target, plan.columns, plan.iter_rows(), ['Cycle test'])
            book = XlsxBook(target.read_bytes())
            try:
                rows = [[book.cell_text(c) for _, c in sorted(book.row_cells(r).items())]
                        for r in book.rows(book.sheets[0])]
                start = rows.index([c['label'] for c in plan.columns]) + 1
                index = [c['key'] for c in plan.columns].index('quantityTiao')
                self.assertEqual([float(r[index]) for r in rows[start:]], [0, 0])
            finally:
                book.zf.close()

    def engine(self, rows):
        dataset = parse_customers(workbook(rows, CYCLE_HEADERS), 'county', XlsxBook)
        return AllocationQuery(dataset, {'reviewRows': [reviewed()]}, {'3': 'keep'}, 'county')

    def test_week_and_month_change_quantities_not_customer_presence(self):
        engine = self.engine([customer('001') + [WEEKLY, ''],
                              customer('002') + [BIWEEKLY, FIRST],
                              customer('003') + [BIWEEKLY, SECOND],
                              customer('004') + [MONTHLY, FIRST]])
        for parity, month, expected in [('odd', False, [7, 7, 0, 0]),
                                         ('even', False, [7, 0, 7, 0]),
                                         ('even', True, [7, 0, 7, 7])]:
            filters = {'weekParity': parity, 'monthlyWeek': month}
            plan = engine.plan_query(filters, 'detail')
            self.assertEqual([r['quantityTiao'] for r in plan.iter_rows()], expected)
            self.assertEqual(plan.row_count, 4)
            self.assertEqual(plan.totals['quantityTiao'], sum(expected))
            self.assertEqual(engine.plan_query(filters, 'issues').row_count, 0)

    def test_bad_cycle_zero_in_details_and_reason_in_review_list(self):
        rows = [customer('001') + ['', ''], customer('002') + ['unexpected', FIRST],
                customer('003') + [BIWEEKLY, ''], customer('004') + [BIWEEKLY, 'third']]
        engine = self.engine(rows)
        filters = {'weekParity': 'odd', 'monthlyWeek': False}
        detail = engine.plan_query(filters, 'detail')
        self.assertEqual(detail.row_count, 4)
        self.assertTrue(all(r['quantityTiao'] == 0 for r in detail.iter_rows()))
        self.assertEqual(detail.totals['cycleReviewCustomerCount'], 4)
        issues = list(engine.plan_query(filters, 'issues').iter_rows())
        self.assertEqual(len(issues), 4)
        self.assertTrue(all(r['reason'] and r['customerCode'] and r['customerName'] for r in issues))

    def test_missing_columns_are_cycle_warnings_not_silent_weekly_default(self):
        dataset = parse_customers(workbook([customer()]), 'county', XlsxBook)
        engine = AllocationQuery(dataset, {'reviewRows': [reviewed()]}, {'3': 'keep'}, 'county')
        plan = engine.plan_query({'weekParity': 'odd', 'monthlyWeek': False}, 'detail')
        self.assertEqual(plan.row_count, 1)
        self.assertEqual(plan.totals['quantityTiao'], 0)
        self.assertEqual(plan.totals['cycleReviewCustomerCount'], 1)

    def test_duplicate_cycle_conflicts_not_silently_collapsed(self):
        engine = self.engine([customer('001') + [BIWEEKLY, FIRST],
                              customer('001') + [BIWEEKLY, SECOND]])
        self.assertEqual(engine.dataset.summary['invalidCustomerCount'], 1)
