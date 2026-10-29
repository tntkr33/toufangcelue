import unittest

from strategy_excel_tool.server import Handler, create_http_server


class PortTests(unittest.TestCase):
    def test_second_server_uses_different_port(self):
        first, _ = create_http_server("127.0.0.1", 0, Handler)
        self.addCleanup(first.server_close)
        occupied = first.server_address[1]
        second, selected = create_http_server("127.0.0.1", occupied, Handler)
        self.addCleanup(second.server_close)
        self.assertNotEqual(occupied, selected)
        self.assertEqual(second.server_address[1], selected)
