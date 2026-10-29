import unittest
import xlrd

from strategy_excel_tool import server as s


class ZeroExportTests(unittest.TestCase):
    def session(self, rows):
        data = s.build_demo_xlsx(['商品编码', '商品名称', '本期投放策略'], rows)
        return {'strategy': data, 'templateKey': 'district_type', 'review': s.analyze_and_apply(data, 'district_type')}

    def test_only_explicit_whole_product_stops_are_candidates(self):
        texts = ['停投', '库存不足，停投。', '无库存', '东兰、巴马：停投',
                 '城市、县城：停投', '1-30档0条', '', '待领导确认',
                 '无库存。\n凤山县：1-30档2条', '暂不停投', '无库存，待确认其他安排']
        session = self.session([[str(71000+i), '测试'+str(i), t] for i,t in enumerate(texts)])
        self.assertEqual([c['sourceRow'] for c in session['review']['zeroCandidates']], [2,3,4])

    def test_choices_filter_whole_products_keep_local_stops_and_change_cache(self):
        session = self.session([['01614','整品规停投','停投'],['01686','无库存','无库存'],
                                ['01361','正常','1-30档3条'],['01497','局部停投','东兰、巴马：停投']])
        blocked = s.build_export_for_session(session)
        self.assertTrue(blocked['requiresZeroConfirmation'])
        self.assertNotIn('output', session)
        choices = {'2':'omit','3':'keep','4':'keep','5':'keep'}
        result = s.build_export_for_session(session, zero_choices=choices, zero_token=blocked['zeroToken'])
        self.assertTrue(result['ok'])
        sheet = xlrd.open_workbook(file_contents=result['output'], formatting_info=True).sheet_by_index(0)
        self.assertTrue(sheet.rowinfo_map[0].hidden)
        rows = [sheet.row_values(i) for i in range(4,sheet.nrows)]
        self.assertEqual(len(rows), 3*44)
        self.assertEqual({r[2] for r in rows}, {'01686','01361','01497'})
        self.assertTrue(all(r[8:38] == [0]*30 for r in rows if r[2]=='01686'))
        self.assertTrue(all(r[8:38] == [3]*30 for r in rows if r[2]=='01361'))
        kept = s.build_export_for_session(session, zero_choices={'2':'keep','3':'keep','4':'keep','5':'keep'}, zero_token=blocked['zeroToken'])
        self.assertNotEqual(kept['output'], result['output'])
        self.assertEqual(kept['exportSummary']['productCount'],4)
        self.assertTrue(s.build_export_for_session(session)['requiresZeroConfirmation'])

    def test_manual_edit_invalidates_confirmation_and_nonzero_removes_candidate(self):
        session=self.session([['01614','停投品规','停投'],['01686','正常','1-30档1条']])
        token=session['review']['zeroToken']
        row=session['review']['reviewRows'][0]
        s.update_review_session(session, {'action':'numbers','sourceRow':2,'changes':[{'id':row['id'],'tier':30,'value':100}]})
        self.assertFalse(session['review']['zeroCandidates'])
        old=s.build_export_for_session(session,zero_choices={'2':'omit'},zero_token=token)
        self.assertFalse(old['ok'])
        self.assertNotIn('output',old)
        result=s.build_export_for_session(session,zero_choices={'2':'keep','3':'keep'},zero_token=session['review']['zeroToken'])
        self.assertTrue(result['ok'])
        sheet=xlrd.open_workbook(file_contents=result['output']).sheet_by_index(0)
        self.assertEqual(sheet.cell_value(4,8),100)

    def test_all_omitted_and_incomplete_decisions_are_rejected(self):
        session=self.session([['01614','停投','停投']])
        token=session['review']['zeroToken']
        self.assertFalse(s.build_export_for_session(session,zero_choices={'2':'omit'},zero_token=token)['ok'])
        self.assertFalse(s.build_export_for_session(session,zero_choices={},zero_token=token)['ok'])
        self.assertFalse(s.build_export_for_session(session,zero_choices={'2':'bad'},zero_token=token)['ok'])
        self.assertNotIn('output',session)
