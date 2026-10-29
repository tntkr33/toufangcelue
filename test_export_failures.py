import tempfile
import threading
import unittest
import urllib.request
from pathlib import Path
from unittest.mock import patch

from strategy_excel_tool import server as s


class ExportRegressionTests(unittest.TestCase):
    def test_period_validation_requires_confirmation(self):
        for name in ['202610W1(10.15-10.111).xls', '202610W1(99.99-99.99)', '202610W1(10.15-10.11)', '202602W1(02.29-03.01)', '自定义名称']:
            result = s.inspect_export_period(name)
            self.assertTrue(result['requiresConfirmation'], name)
            self.assertEqual(result['period'], s.next_week_default_export_stem())
            self.assertTrue(result['reason'])
        for name in ['', '202610W1(10.01-10.04)', '202612W5(12.28-01.03)', '202802W5(02.28-03.05)']:
            self.assertFalse(s.inspect_export_period(name)['requiresConfirmation'], name)

    def test_period_without_extension_keeps_date_dots(self):
        for name in ['202610W1(10.01-10.04)', '202610W1(10.01-10.04).xls']:
            self.assertEqual(s.export_period_from_filename(name), '202610W1(10.01-10.04)')

    def test_download_from_chinese_directory(self):
        with tempfile.TemporaryDirectory() as root:
            output = Path(root) / '卷烟工具' / 'outputs'
            output.mkdir(parents=True)
            (output / 'test.xls').write_bytes(b'workbook-bytes')
            httpd, _ = s.create_http_server('127.0.0.1', 0, s.Handler)
            thread = threading.Thread(target=httpd.serve_forever, daemon=True)
            thread.start()
            try:
                with patch.object(s, 'OUTPUT_DIR', output):
                    with urllib.request.urlopen('http://127.0.0.1:%s/api/download?name=test.xls' % httpd.server_address[1]) as response:
                        self.assertEqual(response.read(), b'workbook-bytes')
            finally:
                httpd.shutdown()
                httpd.server_close()
                thread.join()
