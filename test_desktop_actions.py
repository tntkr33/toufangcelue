import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from strategy_excel_tool import desktop_actions as d


class DesktopActionsTests(unittest.TestCase):
    def test_streaming_xlsx_atomic_save(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'query.xlsx'
            target.write_bytes(b'old')
            actions = d.SavedFiles()
            with patch.object(d, 'choose_save_path', return_value=target):
                with self.assertRaises(ValueError):
                    actions.choose('s', target.name)
                ident = actions.choose('s', target.name, extension='.xlsx')['saveId']
            def fail(path):
                path.write_bytes(b'partial')
                raise RuntimeError('test failure')
            with self.assertRaises(RuntimeError):
                actions.write_generated(ident, 's', fail)
            self.assertEqual(target.read_bytes(), b'old')
            self.assertEqual(list(Path(tmp).iterdir()), [target])
            with self.assertRaises(ValueError):
                actions.open(ident, 'open')
            def produce(path):
                path.write_bytes(b'complete')
                return {'rowCount': 5}
            path, result = actions.write_generated(ident, 's', produce)
            self.assertEqual(path, target)
            self.assertEqual(result['rowCount'], 5)
            self.assertEqual(target.read_bytes(), b'complete')

    def test_only_chosen_successfully_written_file_can_open(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / '202610W1(10.01-10.04).xls'
            actions = d.SavedFiles()
            with patch.object(d, 'choose_save_path', return_value=target):
                chosen = actions.choose('session', target.name)
            ident = chosen['saveId']
            with self.assertRaises(ValueError):
                actions.open(ident, 'open')
            with self.assertRaises(ValueError):
                actions.target(ident, 'other-session')
            actions.write(ident, 'session', b'workbook')
            self.assertEqual(target.read_bytes(), b'workbook')
            with patch.object(d, 'open_saved_path') as opened:
                actions.open(ident, 'reveal')
                opened.assert_called_once_with(target, 'reveal')
            with self.assertRaises(ValueError):
                actions.open('../private_templates', 'open')
            target.unlink()
            with self.assertRaises(ValueError):
                actions.open(ident, 'open')

    def test_cancel_and_write_failure_do_not_enable_open(self):
        actions = d.SavedFiles()
        with patch.object(d, 'choose_save_path', return_value=None):
            self.assertTrue(actions.choose('session', 'test.xls')['cancelled'])
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'test.xls'
            target.write_bytes(b'old')
            with patch.object(d, 'choose_save_path', return_value=target):
                ident = actions.choose('session', 'test.xls')['saveId']
            with patch.object(d.os, 'replace', side_effect=PermissionError('occupied')):
                with self.assertRaises(PermissionError):
                    actions.write(ident, 'session', b'new')
            self.assertEqual(target.read_bytes(), b'old')
            self.assertEqual(len(list(Path(tmp).iterdir())), 1)
            with self.assertRaises(ValueError):
                actions.open(ident, 'open')


class PageLifetimeTests(unittest.TestCase):
    def test_reload_and_multiple_pages(self):
        now = [0]
        lifetime = d.PageLifetime(clock=lambda: now[0])
        now[0] = 100
        self.assertFalse(lifetime.should_stop())
        lifetime.connect('one')
        lifetime.disconnect('one')
        now[0] += 9
        self.assertFalse(lifetime.should_stop())
        lifetime.connect('reload')
        now[0] += 100
        self.assertFalse(lifetime.should_stop())
        lifetime.connect('two')
        lifetime.disconnect('reload')
        now[0] += 20
        self.assertFalse(lifetime.should_stop())
        lifetime.disconnect('two')
        now[0] += 10
        self.assertTrue(lifetime.should_stop())

    def test_work_and_native_dialog_protect_from_exit(self):
        now = [0]
        lifetime = d.PageLifetime(clock=lambda: now[0])
        lifetime.connect('one')
        lifetime.begin_work()
        lifetime.disconnect('one')
        now[0] += 100
        self.assertFalse(lifetime.should_stop())
        lifetime.end_work()
        now[0] += 9
        self.assertFalse(lifetime.should_stop())
        now[0] += 1
        self.assertTrue(lifetime.should_stop())
