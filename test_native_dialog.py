"""Opt-in Windows dialog smoke test; operates only on this process's window."""
import ctypes
import os
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

from strategy_excel_tool.desktop_actions import choose_save_path


@unittest.skipUnless(sys.platform == 'win32' and os.environ.get('TEST_NATIVE_DIALOG') == '1',
                     'Set TEST_NATIVE_DIALOG=1 for real Windows save-dialog tests')
class NativeDialogTests(unittest.TestCase):
    def drive_dialog(self, target=None, extension='.xls'):
        from ctypes import wintypes as w
        user = ctypes.WinDLL('user32', use_last_error=True)
        callback_type = ctypes.WINFUNCTYPE(w.BOOL, w.HWND, w.LPARAM)
        user.EnumWindows.argtypes = [callback_type, w.LPARAM]
        user.EnumChildWindows.argtypes = [w.HWND, callback_type, w.LPARAM]
        user.GetWindowTextW.argtypes = [w.HWND, w.LPWSTR, ctypes.c_int]
        user.GetClassNameW.argtypes = [w.HWND, w.LPWSTR, ctypes.c_int]
        user.GetWindowThreadProcessId.argtypes = [w.HWND, ctypes.POINTER(w.DWORD)]
        user.PostMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
        user.SendMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
        done = threading.Event()
        errors = []

        def windows():
            found = []
            @callback_type
            def collect(hwnd, _):
                pid = w.DWORD()
                user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                text = ctypes.create_unicode_buffer(512)
                user.GetWindowTextW(hwnd, text, len(text))
                if pid.value == os.getpid() and text.value == '保存卷烟投放策略':
                    found.append(hwnd)
                return True
            user.EnumWindows(collect, 0)
            return found

        def automate():
            try:
                for _ in range(100):
                    handles = windows()
                    if handles:
                        hwnd = handles[0]
                        if target is None:
                            user.PostMessageW(hwnd, 0x111, 2, 0)
                            return
                        edits = []
                        @callback_type
                        def child(control, _):
                            cls = ctypes.create_unicode_buffer(128)
                            text = ctypes.create_unicode_buffer(1024)
                            user.GetClassNameW(control, cls, len(cls))
                            user.GetWindowTextW(control, text, len(text))
                            if cls.value == 'Edit' and text.value.startswith('native-test'):
                                edits.append(control)
                            return True
                        user.EnumChildWindows(hwnd, child, 0)
                        if edits:
                            text = ctypes.create_unicode_buffer(str(target))
                            user.SendMessageW(edits[0], 0xC, 0, ctypes.addressof(text))
                            user.PostMessageW(hwnd, 0x111, 1, 0)
                            if not done.wait(5):
                                raise AssertionError('Native save did not close')
                            return
                    if done.wait(0.1):
                        return
                raise AssertionError('Native save dialog/control was not found')
            except Exception as exc:
                errors.append(exc)
                for hwnd in windows():
                    user.PostMessageW(hwnd, 0x10, 0, 0)

        thread = threading.Thread(target=automate, daemon=True)
        thread.start()
        try:
            result = choose_save_path('native-test' + extension)
        finally:
            done.set()
            thread.join(12)
        if errors:
            raise errors[0]
        return result

    def test_cancel(self):
        self.assertIsNone(self.drive_dialog())

    def test_choose_unicode_path_does_not_write_before_confirmation(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / '中文测试202610W1(10.01-10.04).xls'
            self.assertEqual(self.drive_dialog(target), target)
            self.assertFalse(target.exists())

    def test_choose_xlsx_uses_xlsx_default_extension(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / '客户查询测试'
            self.assertEqual(self.drive_dialog(target, '.xlsx'), target.with_suffix('.xlsx'))
            self.assertFalse(target.with_suffix('.xlsx').exists())
