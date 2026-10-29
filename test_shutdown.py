import threading
import unittest
import urllib.error
import urllib.request

from strategy_excel_tool import server as s


class ShutdownTests(unittest.TestCase):
    def test_other_tab_cannot_exit_during_file_write_or_dialog(self):
        httpd, _ = s.create_http_server('127.0.0.1', 0, s.Handler)
        thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        thread.start()
        url = 'http://127.0.0.1:%s' % httpd.server_address[1]
        httpd.page_lifetime.begin_work()
        try:
            request = urllib.request.Request(url+'/api/shutdown', data=b'', headers={'X-Tool-Token': httpd.shutdown_token})
            with self.assertRaises(urllib.error.HTTPError) as denied:
                urllib.request.urlopen(request)
            self.assertEqual(denied.exception.code, 409)
            denied.exception.close()
            self.assertTrue(thread.is_alive())
        finally:
            httpd.page_lifetime.end_work()
            httpd.shutdown()
            thread.join(3)
            httpd.server_close()

    def test_only_authorized_post_stops_this_server(self):
        httpd, _ = s.create_http_server('127.0.0.1', 0, s.Handler)
        thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        thread.start()
        url = 'http://127.0.0.1:%s' % httpd.server_address[1]
        try:
            with self.assertRaises(urllib.error.HTTPError) as denied:
                urllib.request.urlopen(urllib.request.Request(url+'/api/shutdown', data=b''))
            self.assertEqual(denied.exception.code, 403)
            denied.exception.close()
            self.assertTrue(thread.is_alive())
            page = urllib.request.urlopen(url).read().decode()
            self.assertIn(httpd.shutdown_token, page)
            request = urllib.request.Request(url+'/api/shutdown', data=b'', headers={'X-Tool-Token': httpd.shutdown_token})
            self.assertEqual(urllib.request.urlopen(request).status, 200)
            thread.join(3)
            self.assertFalse(thread.is_alive())
        finally:
            if thread.is_alive():
                httpd.shutdown()
            httpd.server_close()
