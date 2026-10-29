import threading
import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch
from strategy_excel_tool.customer_service import CustomerService, CustomerStateError
from strategy_excel_tool import desktop_actions as d, query_workbook


class CustomerServiceTests(unittest.TestCase):
    def test_catalog_before_upload_and_default_query_without_confirmation(self):
        from test_customer_allocation import HEADERS, workbook, customer, reviewed
        from server import XlsxBook
        service = CustomerService(XlsxBook, threading.RLock())
        row = reviewed()
        session = {'templateKey': 'county', 'review': {'zeroToken': 't',
            'reviewRows': [row], 'exportProducts': [dict(row, defaultChoice='omit')]}}
        options = service.options(session)
        self.assertEqual(len(options['options']['products']), 1)
        self.assertFalse(options['choicesConfirmed'])
        service.upload(session, workbook([customer() + ['\u4e00\u5468\u4e00\u6b21']], HEADERS + ['\u8ba2\u8d27\u5468\u671f']), 'customers.xlsx', 't')
        with self.assertRaises(CustomerStateError):
            service.query(session, {'revision': 1})
        filters = {'weekParity': 'odd', 'monthlyWeek': False}
        result = service.query(session, {'revision': 1, 'filters': filters})
        self.assertEqual(result['rows'][0]['quantityTiao'], 0)
        self.assertEqual(result['rows'][0]['allocationStatus'], '\u5df2\u6392\u9664\u6295\u653e')
        self.assertFalse(service.state(session)['choicesConfirmed'])
        self.assertEqual(row['assignments'][1], 7)
        service.confirm_choices(session, {'3': 'keep'}, 't')
        with self.assertRaises(CustomerStateError):
            service.get_query(session, result['queryId'])
        self.assertEqual(service.query(session, {'revision': 2, 'filters': filters})['totals']['quantityTiao'], 7)

    def setUp(self):
        self.service = CustomerService(None, threading.RLock())
        self.session = {'review': {'zeroToken': 't', 'exportProducts': [
            {'sourceRow': 4, 'productCode': '001', 'productName': 'Test'}]}}

    def test_choices_are_complete_and_revision_bound(self):
        for choices, token in [({}, 't'), ({'4': 'omit'}, 't'), ({'4': 'keep'}, 'old')]:
            with self.assertRaises(CustomerStateError):
                self.service.confirm_choices(self.session, choices, token)
        state = self.service.confirm_choices(self.session, {'4': 'keep'}, 't')
        self.assertEqual(state['revision'], 1)
        self.assertTrue(state['choicesConfirmed'])
        self.assertEqual(self.service.confirm_choices(self.session, {'4': 'keep'}, 't')['revision'], 1)
        self.service.invalidate(self.session, strategy=True)
        self.assertFalse(self.service.state(self.session)['choicesConfirmed'])

    def test_stale_query_and_export_lock(self):
        self.session['customerEpoch'] = 2
        self.session['customerQueries'] = {'q': {'epoch': 1, 'token': 't'}}
        with self.assertRaises(CustomerStateError):
            self.service.get_query(self.session, 'q')
        self.session['customerExport'] = {'status': 'running'}
        with self.assertRaises(CustomerStateError):
            self.service.invalidate(self.session)
        with self.assertRaises(CustomerStateError):
            self.service.confirm_choices(self.session, {'4': 'keep'}, 't')
        self.assertEqual(self.session['customerEpoch'], 2)

    def test_running_export_pins_revision_and_releases_lifetime_on_failure(self):
        class Plan:
            row_count = 1
            columns = [{'key': 'quantityTiao', 'label': 'Quantity'}]
            totals = {'incomplete': False}
            def iter_rows(self):
                yield {'quantityTiao': 2}
        self.session['customerQueries'] = {'q': {'epoch': 0, 'token': 't', 'plan': Plan(), 'view': 'summary', 'filters': {}}}
        lifetime = d.PageLifetime()
        entered, release, finished = threading.Event(), threading.Event(), threading.Event()
        original_end = lifetime.end_work
        def ended():
            original_end()
            finished.set()
        lifetime.end_work = ended
        def writer(*args, **kwargs):
            entered.set()
            if not release.wait(5):
                raise RuntimeError('test timeout')
            raise OSError('occupied')
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'query.xlsx'
            target.write_bytes(b'original')
            saved = d.SavedFiles()
            with patch.object(d, 'choose_save_path', return_value=target):
                ident = saved.choose('s', target.name, extension='.xlsx')['saveId']
            with patch.object(query_workbook, 'write_query_workbook', side_effect=writer):
                job = self.service.start_export(self.session, {'queryId': 'q', 'sessionId': 's', 'saveId': ident}, saved, lifetime)
                try:
                    self.assertTrue(entered.wait(5))
                    self.assertEqual(lifetime.work, 1)
                    with self.assertRaises(CustomerStateError):
                        self.service.invalidate(self.session, strategy=True)
                    with self.assertRaises(CustomerStateError):
                        self.service.start_export(self.session, {'queryId': 'q', 'sessionId': 's', 'saveId': ident}, saved, lifetime)
                finally:
                    release.set()
                    self.assertTrue(finished.wait(5))
            self.assertEqual(lifetime.work, 0)
            self.assertEqual(self.service.export_status(self.session, job['jobId'])['status'], 'failed')
            self.assertEqual(target.read_bytes(), b'original')
            self.assertEqual(list(Path(tmp).iterdir()), [target])
