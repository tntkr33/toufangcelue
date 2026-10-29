import ctypes
import os
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path


def choose_save_path(filename):
    if sys.platform == 'win32':
        from ctypes import wintypes as w

        class OpenFileName(ctypes.Structure):
            _fields_ = [
                ('lStructSize', w.DWORD), ('hwndOwner', w.HWND), ('hInstance', w.HINSTANCE),
                ('lpstrFilter', w.LPCWSTR), ('lpstrCustomFilter', w.LPWSTR),
                ('nMaxCustFilter', w.DWORD), ('nFilterIndex', w.DWORD),
                ('lpstrFile', w.LPWSTR), ('nMaxFile', w.DWORD),
                ('lpstrFileTitle', w.LPWSTR), ('nMaxFileTitle', w.DWORD),
                ('lpstrInitialDir', w.LPCWSTR), ('lpstrTitle', w.LPCWSTR),
                ('Flags', w.DWORD), ('nFileOffset', w.WORD), ('nFileExtension', w.WORD),
                ('lpstrDefExt', w.LPCWSTR), ('lCustData', ctypes.c_ssize_t),
                ('lpfnHook', ctypes.c_void_p), ('lpTemplateName', w.LPCWSTR),
                ('pvReserved', ctypes.c_void_p), ('dwReserved', w.DWORD), ('FlagsEx', w.DWORD),
            ]

        buffer = ctypes.create_unicode_buffer(32768)
        buffer.value = Path(filename).name
        options = OpenFileName()
        options.lStructSize = ctypes.sizeof(options)
        extension = 'xlsx' if Path(filename).suffix.lower() == '.xlsx' else 'xls'
        options.lpstrFilter = 'Excel (*.%s)\0*.%s\0\0' % (extension, extension)
        options.lpstrFile = ctypes.cast(buffer, w.LPWSTR)
        options.nMaxFile = len(buffer)
        options.nFilterIndex = 1
        options.lpstrDefExt = extension
        options.lpstrTitle = '保存卷烟投放策略'
        # Explorer style, existing parent, overwrite confirmation, unchanged cwd.
        options.Flags = 0x80000 | 0x800 | 0x2 | 0x8
        api = ctypes.WinDLL('comdlg32', use_last_error=True)
        api.GetSaveFileNameW.argtypes = [ctypes.POINTER(OpenFileName)]
        api.GetSaveFileNameW.restype = w.BOOL
        api.CommDlgExtendedError.restype = w.DWORD
        if not api.GetSaveFileNameW(ctypes.byref(options)):
            error = api.CommDlgExtendedError()
            if error:
                raise OSError('无法打开保存窗口，错误码 %s' % error)
            return None
        return Path(buffer.value)
    if sys.platform == 'darwin':
        script = ('on run argv\ntry\nreturn POSIX path of (choose file name with prompt '
                  '"保存卷烟投放策略" default name (item 1 of argv))\n'
                  'on error number -128\nreturn ""\nend try\nend run')
        result = subprocess.run(['osascript', '-e', script, filename], capture_output=True, text=True, check=True)
        return Path(result.stdout.strip()) if result.stdout.strip() else None
    raise OSError('系统另存为目前支持 Windows 和 macOS。')


def open_saved_path(path, action):
    if action not in ('open', 'reveal'):
        raise ValueError('无效的文件操作。')
    if sys.platform == 'win32':
        if action == 'open':
            os.startfile(str(path))
        else:
            subprocess.Popen(['explorer.exe', '/select,', str(path)])
    elif sys.platform == 'darwin':
        subprocess.run(['open', *(['-R'] if action == 'reveal' else []), str(path)], check=True)
    else:
        raise OSError('此系统不支持打开本地文件。')


class SavedFiles:
    def __init__(self):
        self.entries = {}
        self.lock = threading.RLock()
        self.dialog_lock = threading.Lock()

    def choose(self, session_id, filename, extension='.xls'):
        if extension not in ('.xls', '.xlsx'):
            raise ValueError('不支持的文件格式。')
        if not self.dialog_lock.acquire(blocking=False):
            raise ValueError('另一个保存窗口尚未关闭，请先完成或取消。')
        try:
            path = choose_save_path(filename)
        finally:
            self.dialog_lock.release()
        if path is None:
            return {'ok': True, 'cancelled': True}
        path = Path(path).absolute()
        if path.suffix.lower() != extension:
            raise ValueError('文件名必须以 %s 结尾，请重新选择保存位置。' % extension)
        ident = uuid.uuid4().hex
        with self.lock:
            self.entries[ident] = {'path': path, 'session': session_id, 'saved': False}
        return {'ok': True, 'saveId': ident, 'filename': path.name, 'path': str(path)}

    def target(self, ident, session_id):
        with self.lock:
            entry = self.entries.get(ident)
            if not entry or entry['session'] != session_id:
                raise ValueError('保存位置已失效，请重新导出并选择位置。')
            return entry['path']

    def write(self, ident, session_id, data):
        with self.lock:
            path = self.target(ident, session_id)
            self.entries[ident]['saved'] = False
            temporary = None
            try:
                with tempfile.NamedTemporaryFile(dir=path.parent, prefix='.strategy-', suffix='.tmp', delete=False) as f:
                    temporary = Path(f.name)
                    f.write(data)
                os.replace(temporary, path)
            finally:
                if temporary is not None:
                    temporary.unlink(missing_ok=True)
            self.entries[ident]['saved'] = True
            return path

    def write_generated(self, ident, session_id, producer):
        with self.lock:
            path = self.target(ident, session_id)
            self.entries[ident]['saved'] = False
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(dir=path.parent, prefix='.strategy-', suffix='.tmp', delete=False) as f:
                temporary = Path(f.name)
            result = producer(temporary)
            os.replace(temporary, path)
            with self.lock:
                self.entries[ident]['saved'] = True
            return path, result
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)

    def open(self, ident, action):
        with self.lock:
            entry = self.entries.get(ident)
            if not entry or not entry['saved'] or not entry['path'].is_file():
                raise ValueError('文件尚未保存，或已被移动、删除，请重新导出。')
            open_saved_path(entry['path'], action)


class PageLifetime:
    def __init__(self, grace=10, clock=time.monotonic):
        self.grace = grace
        self.clock = clock
        self.pages = set()
        self.work = 0
        self.deadline = None
        self.lock = threading.RLock()

    def connect(self, ident):
        with self.lock:
            self.pages.add(ident)
            self.deadline = None

    def disconnect(self, ident):
        with self.lock:
            self.pages.discard(ident)
            if not self.pages:
                self.deadline = self.clock() + self.grace

    def begin_work(self):
        with self.lock:
            self.work += 1

    def end_work(self):
        with self.lock:
            self.work -= 1
            if not self.pages and self.deadline is not None:
                self.deadline = self.clock() + self.grace

    def should_stop(self):
        with self.lock:
            return (not self.pages and not self.work and self.deadline is not None
                    and self.clock() >= self.deadline)
