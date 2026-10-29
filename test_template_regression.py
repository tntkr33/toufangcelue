import unittest
import xlrd
from strategy_excel_tool import server as s


class TemplateRegressionTests(unittest.TestCase):
    def test_all_four_import_formats_keep_hidden_identity_and_period(self):
        data = s.build_demo_xlsx(['商品编码', '商品名称', '本期投放策略'], [['01614', 'Test', '1-30档2条']])
        for key, config in s.TEMPLATES.items():
            with self.subTest(template=key):
                result = s.analyze_and_apply(data, key, write_output=True, export_title='202610W2(10.12-10.18)')
                self.assertTrue(result['ok'], result.get('errors'))
                original = xlrd.open_workbook(str(s.TEMPLATE_DIR / config['filename']), formatting_info=True).sheet_by_index(0)
                output = xlrd.open_workbook(file_contents=result['output'], formatting_info=True).sheet_by_index(0)
                self.assertEqual(output.name, original.name)
                self.assertEqual(sorted(output.merged_cells), sorted(original.merged_cells))
                self.assertTrue(output.rowinfo_map[0].hidden)
                self.assertEqual(output.cell_value(0, 0), config['marker'][0])
                self.assertEqual(s.normalize_xls_marker_value(output.cell_value(0, 1)), config['marker'][1])
                self.assertEqual(output.cell_value(1, 0), '202610W2(10.12-10.18)')
                self.assertTrue(result['output'].startswith(bytes.fromhex('D0CF11E0A1B11AE1')))
