"""Opt-in smoke test of the built executable; uses synthetic data only."""
import json
import os
from pathlib import Path
import re
import socket
import subprocess
import sys
import time
import unittest
import urllib.request
from strategy_excel_tool import server as s


@unittest.skipUnless(os.environ.get('PACKAGED_TOOL_EXE'), 'Set PACKAGED_TOOL_EXE to verify a Windows release')
class PackagedCustomerTests(unittest.TestCase):
    def test_packaged_version_templates_and_customer_query(self):
        with socket.socket() as sock:
            sock.bind(('127.0.0.1', 0))
            port = sock.getsockname()[1]
        env = dict(os.environ, STRATEGY_TOOL_PORT=str(port))
        # Suppress only the test process's automatic browser launch.
        env['BROWSER'] = '"%s" -c "pass" %%s' % Path(sys.executable).as_posix()
        process = subprocess.Popen([os.environ['PACKAGED_TOOL_EXE']], env=env,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        base = 'http://127.0.0.1:%s' % port
        token = None
        def call(path, payload=None):
            headers = {'X-Tool-Token': token} if token else {}
            if payload is not None:
                headers['Content-Type'] = 'application/json'
            req = urllib.request.Request(base + path, None if payload is None else json.dumps(payload).encode(), headers)
            with urllib.request.urlopen(req, timeout=30) as response:
                return json.load(response)
        def upload(path, fields, file_key, data):
            boundary = 'test-packaged-customer-boundary'
            parts = []
            for key, value in fields.items():
                parts.append(('--%s\r\nContent-Disposition: form-data; name="%s"\r\n\r\n%s\r\n' % (boundary, key, value)).encode())
            parts.extend([('--%s\r\nContent-Disposition: form-data; name="%s"; filename="synthetic.xlsx"\r\nContent-Type: application/octet-stream\r\n\r\n' % (boundary, file_key)).encode(), data, ('\r\n--%s--\r\n' % boundary).encode()])
            req = urllib.request.Request(base + path, b''.join(parts), {'Content-Type': 'multipart/form-data; boundary=' + boundary, 'X-Tool-Token': token})
            with urllib.request.urlopen(req, timeout=30) as response:
                return json.load(response)
        try:
            deadline = time.monotonic() + 30
            while True:
                try:
                    with urllib.request.urlopen(base + '/', timeout=1) as response:
                        page = response.read().decode()
                    break
                except OSError:
                    if process.poll() is not None or time.monotonic() > deadline:
                        self.fail('Packaged tool did not start')
                    time.sleep(.1)
            token = re.search(r"const shutdownToken\s*=\s*['\"]([^'\"]+)", page)[1]
            self.assertEqual(call('/api/version')['version'], 'v2.0')
            self.assertIn('customerTab', page)
            self.assertIn('id="reviewExportBtn"', page)
            self.assertTrue(all(t['exists'] for t in call('/api/templates')['templates']))
            strategy = s.build_demo_xlsx(['商品编码', '商品名称', '本期投放策略'], [['01614', 'Synthetic', '1-30档2条']])
            preview = upload('/api/analyze', {'templateKey': 'district_type'}, 'strategyWorkbook', strategy)
            sid = preview['sessionId']
            customer = s.build_demo_xlsx(['客户编码', '客户名称', '营销部', '档位', '市场类型(扩展)', '订货周期'], [['0001', 'Synthetic', '城区营销部', '30', '城市', '一周一次']])
            state = upload('/api/customer/upload', {'sessionId': sid, 'reviewToken': preview['zeroToken']}, 'customerWorkbook', customer)
            self.assertEqual(state['customerSummary']['validCustomerCount'], 1)
            state = call('/api/customer/choices', {'sessionId': sid, 'zeroToken': preview['zeroToken'], 'choices': {'2': 'keep'}})
            result = call('/api/customer/query', {'sessionId': sid, 'revision': state['revision'], 'view': 'detail', 'filters': {'weekParity': 'odd', 'monthlyWeek': False}})
            self.assertEqual(result['totals']['quantityTiao'], 2)
            self.assertEqual(result['rows'][0]['customerCode'], '0001')
            self.assertEqual(result['rows'][0]['productCode'], '01614')
        finally:
            if token and process.poll() is None:
                call('/api/shutdown', {})
            try:
                process.wait(timeout=15)
            except subprocess.TimeoutExpired:
                process.terminate()
                process.wait(timeout=5)
        self.assertEqual(process.returncode, 0)
