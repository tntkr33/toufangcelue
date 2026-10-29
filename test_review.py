import unittest
import xlrd
try:
    from . import server as s
except ImportError:
    import server as s


class ReviewTests(unittest.TestCase):
    def test_confirmation_preview_does_not_change_session(self):
        import copy
        data = s.build_demo_xlsx(['商品编码', '商品名称', '本期投放策略'], [['01614', '测试', '1-30档2条']])
        session = {'strategy': data, 'templateKey': 'district_type', 'review': s.analyze_and_apply(data, 'district_type'), 'output': b'cached'}
        before = copy.deepcopy(session)
        row = session['review']['reviewRows'][0]
        result = s.update_review_session(session, {'action': 'numbers', 'sourceRow': 2, 'changes': [{'id': row['id'], 'tier': 30, 'value': 100}]}, preview=True)
        self.assertEqual(result['reviewRows'][0]['assignments'][30], 100)
        self.assertEqual(session, before)
        result = s.update_review_session(session, {'action': 'text', 'sourceRow': 2, 'text': '30档50条'}, preview=True)
        self.assertEqual(result['reviewRows'][0]['assignments'][30], 50)
        self.assertEqual(session, before)

    def test_edits_reach_export_and_reparse_clears_numbers(self):
        data = s.build_demo_xlsx(['商品编码', '商品名称', '本期投放策略'], [['01614', '测试', '1-30档2条']])
        session = {'strategy': data, 'templateKey': 'district_type'}
        initial = s.analyze_and_apply(data, 'district_type')
        session['review'] = initial
        row = initial['reviewRows'][0]
        result = s.update_review_session(session, {'action': 'numbers', 'sourceRow': 2, 'changes': [{'id': row['id'], 'tier': 28, 'value': 7}]})
        self.assertTrue(result['ok'])
        output = s.build_export_for_session(session, zero_choices={str(p['sourceRow']):'keep' for p in session['review']['exportProducts']}, zero_token=session['review']['zeroToken'])['output']
        sheet = xlrd.open_workbook(file_contents=output).sheet_by_index(0)
        self.assertEqual(sheet.cell_value(row['row'] - 1, 10), 7)
        result = s.update_review_session(session, {'action': 'text', 'sourceRow': 2, 'text': '28档-30档9条'})
        self.assertEqual(result['reviewRows'][0]['assignments'][28], 9)
        self.assertEqual(result['reviewRows'][0]['originalText'], '1-30档2条')
        output = s.build_export_for_session(session, zero_choices={str(p['sourceRow']):'keep' for p in session['review']['exportProducts']}, zero_token=session['review']['zeroToken'])['output']
        self.assertEqual(xlrd.open_workbook(file_contents=output).sheet_by_index(0).cell_value(row['row'] - 1, 10), 9)
        with self.assertRaises(ValueError):
            s.update_review_session(session, {'action': 'numbers', 'sourceRow': 2, 'changes': [{'id': row['id'], 'tier': 28, 'value': -1}]})


if __name__ == '__main__':
    unittest.main()
