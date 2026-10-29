import unittest

import xlrd
from strategy_excel_tool import server as s


class ExportChoicesTests(unittest.TestCase):
    def session(self, texts):
        data = s.build_demo_xlsx(['商品编码', '商品名称', '本期投放策略'],
                                [[str(70000+i), '品规'+str(i), t] for i, t in enumerate(texts)])
        return {'strategy': data, 'templateKey': 'district_type',
                'review': s.analyze_and_apply(data, 'district_type')}

    def test_defaults_and_warning_linkage(self):
        texts = ['停投', '按价位段投放', '按价位段投放，上限20条（入网日期晚于9月7日客户不参与测算）',
                 '南丹：按价位段投放。\n凤山县：1-30档3条', '1-30档5条',
                 '按价位段投放。\n凤山县：1-30档3条', '南丹：按价位段投放',
                 '不按价位段投放', '按价位段投放，仅限南丹', '暂未确定是否按价位段投放',
                 '按价位段投放（待确认）', '按价位段投放，按档位补充投放',
                 '按价位段投放，仅市区客户参与', '按价位段投放，商圈A投放5条',
                 '按价位段投放，特殊客户追加供货']
        review = self.session(texts)['review']
        items = review['exportProducts']
        self.assertEqual([p['defaultChoice'] for p in items], ['omit']*3+['keep']*12)
        self.assertEqual(items[3]['category'], 'review')
        self.assertEqual(items[1]['warnings'], review['rules'][1]['warnings'])
        self.assertTrue(items[1]['warnings'])
        self.assertIn('价位段', items[1]['reason'])

    def test_any_product_can_be_excluded_and_keep_preserves_manual_numbers(self):
        session = self.session(['1-30档5条', '按价位段投放', '南丹：按价位段投放。\n凤山县：1-30档3条'])
        row = next(r for r in session['review']['reviewRows'] if r['sourceRow'] == 3)
        s.update_review_session(session, {'action': 'numbers', 'sourceRow': 3,
                                        'changes': [{'id': row['id'], 'tier': 30, 'value': 100}]})
        review = session['review']
        self.assertEqual(review['exportProducts'][1]['defaultChoice'], 'keep')
        choices = {'2': 'omit', '3': 'keep', '4': 'keep'}
        result = s.build_export_for_session(session, zero_choices=choices, zero_token=review['zeroToken'])
        self.assertTrue(result['ok'])
        sheet = xlrd.open_workbook(file_contents=result['output'], formatting_info=True).sheet_by_index(0)
        self.assertTrue(sheet.rowinfo_map[0].hidden)
        self.assertEqual(sheet.nrows, 4+88)
        self.assertEqual(sheet.cell_value(4, 8), 100)
        self.assertEqual({sheet.cell_value(r, 2) for r in range(4, sheet.nrows)}, {'70001', '70002'})
        self.assertTrue(any(sheet.cell_value(r, 8) == 3 for r in range(48, sheet.nrows)))
        again = s.build_export_for_session(session, zero_choices={'2':'keep','3':'keep','4':'keep'}, zero_token=review['zeroToken'])
        self.assertEqual(again['exportSummary']['productCount'], 3)
        self.assertNotEqual(result['output'], again['output'])

    def test_every_export_requires_complete_current_manifest(self):
        session = self.session(['1-30档5条', '按价位段投放'])
        blocked = s.build_export_for_session(session)
        self.assertTrue(blocked['requiresZeroConfirmation'])
        self.assertEqual(len(blocked['exportProducts']), 2)
        token = blocked['zeroToken']
        for choices in ({'3': 'omit'}, {'2':'keep','3':'bad'}, {'2':'keep','3':'keep','999':'omit'}, {'2':'omit','3':'omit'}):
            self.assertFalse(s.build_export_for_session(session, zero_choices=choices, zero_token=token)['ok'])
        s.update_review_session(session, {'action':'text','sourceRow':3,'text':'1-30档6条'})
        self.assertFalse(s.build_export_for_session(session, zero_choices={'2':'keep','3':'omit'}, zero_token=token)['ok'])

    def test_exclusion_in_all_four_templates_preserves_marker(self):
        data = s.build_demo_xlsx(['商品编码','商品名称','本期投放策略'],
                                [['01614','价位段品规','按价位段投放'],['01686','正常品规','1-30档5条']])
        for key in s.TEMPLATES:
            with self.subTest(template=key):
                review = s.analyze_and_apply(data, key)
                session = {'strategy':data, 'templateKey':key, 'review':review}
                result = s.build_export_for_session(session, zero_choices={'2':'omit','3':'keep'}, zero_token=review['zeroToken'])
                self.assertTrue(result['ok'])
                sheet = xlrd.open_workbook(file_contents=result['output'],formatting_info=True).sheet_by_index(0)
                self.assertTrue(sheet.rowinfo_map[0].hidden)
                self.assertEqual(sheet.cell_value(0,0), s.TEMPLATES[key]['marker'][0])
                self.assertEqual(result['exportSummary']['productCount'], 1)
                self.assertEqual(result['exportSummary']['generatedRows'], review['summary']['templateDimensionRows'])
