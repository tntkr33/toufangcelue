import json
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest.mock import patch

from strategy_excel_tool import server as s, desktop_actions as d


class CustomerHttpTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.http, _ = s.create_http_server('127.0.0.1', 0, s.Handler)
        self.thread = threading.Thread(target=self.http.serve_forever)
        self.thread.start()
        self.base = 'http://127.0.0.1:%s' % self.http.server_address[1]
        data = s.build_demo_xlsx(['商品编码', '商品名称', '本期投放策略'], [['01614', 'Test', '1-30档2条']])
        self.session = {'strategy': data, 'templateKey': 'district_type', 'review': s.analyze_and_apply(data, 'district_type')}
        self.sid = 'customer-http-test'
        s.SESSIONS[self.sid] = self.session

    def tearDown(self):
        self.http.stopping.set()
        self.http.shutdown()
        self.http.server_close()
        self.thread.join()
        s.SESSIONS.pop(self.sid, None)
        self.tmp.cleanup()

    def request(self, route, payload=None, token=True):
        url = self.base + '/api/customer/' + route
        headers = {'X-Tool-Token': self.http.shutdown_token} if token else {}
        if payload is not None:
            if route == 'query' and not payload.get('queryId'):
                payload = dict(payload, filters=dict(weekParity='odd', monthlyWeek=False, **payload.get('filters', {})))
            data = json.dumps(dict(payload, sessionId=self.sid)).encode()
            headers['Content-Type'] = 'application/json'
        else:
            url += ('&' if '?' in route else '?') + 'sessionId=' + self.sid
            data = None
        try:
            response = urllib.request.urlopen(urllib.request.Request(url, data, headers), timeout=30)
        except urllib.error.HTTPError as exc:
            response = exc
        with response:
            return response.status, json.load(response)

    def test_filtered_export_manual_invalidation_and_security(self):
        self.assertEqual(self.request('state', token=False)[0], 403)
        token = self.session['review']['zeroToken']
        customers = s.build_demo_xlsx(['客户编码', '客户名称', '营销部', '档位', '市场类型(扩展)', '订货周期'], [
            ['0001', 'Customer A', '城区营销部', '30', '城市', '一周一次'],
            ['0002', 'Customer B', '城区营销部', '29', '城市', '一周一次'],
            ['0003', 'Customer C', '凤山县营销部', '30', '县城', '一周一次'],
        ])
        s.CUSTOMERS.upload(self.session, customers, 'customers.xlsx', token)
        code, state = self.request('choices', {'zeroToken': token, 'choices': {'2': 'keep'}})
        self.assertEqual(code, 200, state)
        code, query = self.request('query', {'revision': state['revision'], 'view': 'detail',
            'filters': {'counties': ['城区营销部'], 'minTier': 29, 'maxTier': 30}, 'pageSize': 1})
        self.assertEqual(code, 200, query)
        self.assertEqual(query['totalRows'], 2)
        self.assertEqual(len(query['rows']), 1)
        self.assertEqual(query['totals']['quantityTiao'], 4)
        target = Path(self.tmp.name) / 'filtered.xlsx'
        with patch.object(d, 'choose_save_path', return_value=target):
            code, chosen = self.request('save-location', {'queryId': query['queryId'], 'filename': target.name})
        self.assertEqual(code, 200, chosen)
        code, job = self.request('export-start', {'queryId': query['queryId'], 'saveId': chosen['saveId']})
        self.assertEqual(code, 200, job)
        deadline = time.monotonic() + 20
        while time.monotonic() < deadline:
            _, status = self.request('export-status?jobId=' + job['jobId'])
            if status['status'] != 'running':
                break
            time.sleep(.02)
        self.assertEqual(status['status'], 'complete', status)
        self.assertEqual(status['result']['rowCount'], 2)
        self.assertTrue(target.is_file())
        book = s.XlsxBook(target.read_bytes())
        try:
            rows = [[book.cell_text(c) for _, c in sorted(book.row_cells(r).items())]
                    for r in book.rows(book.sheets[0])]
            header = [c['label'] for c in query['columns']]
            start = rows.index(header) + 1
            data = rows[start:]
            keys = [c['key'] for c in query['columns']]
            self.assertEqual(len(data), 2)
            self.assertEqual({r[keys.index('customerCode')] for r in data}, {'0001', '0002'})
            self.assertEqual(sum(float(r[keys.index('quantityTiao')]) for r in data), query['totals']['quantityTiao'])
            self.assertEqual([r[keys.index('productCode')] for r in data], ['01614', '01614'])
        finally:
            book.zf.close()
        with self.assertRaises(ValueError):
            self.http.saved_files.target(chosen['saveId'], 'another-session')
        row = self.session['review']['reviewRows'][0]
        s.update_review_session(self.session, {'action': 'numbers', 'sourceRow': 2,
            'changes': [{'id': row['id'], 'tier': 30, 'value': 100}]})
        self.assertEqual(self.request('query', {'queryId': query['queryId']})[0], 409)
        self.assertEqual(self.request('export-start', {'queryId': query['queryId'], 'saveId': chosen['saveId']})[0], 409)
        _, state = self.request('state')
        self.assertFalse(state['choicesConfirmed'])
        _, state = self.request('choices', {'zeroToken': state['reviewToken'], 'choices': {'2': 'keep'}})
        _, result = self.request('query', {'revision': state['revision'], 'view': 'summary', 'filters': {'counties': ['城区营销部']}})
        self.assertEqual(result['totals']['quantityTiao'], 102)
