# -*- coding: utf-8 -*-
import os
import sys
import threading
import time
import webbrowser

import server


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    server.TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    server.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    preferred_port = int(os.environ.get("STRATEGY_TOOL_PORT") or "8010")
    try:
        httpd, port = server.create_http_server("127.0.0.1", preferred_port, server.Handler)
    except RuntimeError:
        fallback_port = server.find_available_port("127.0.0.1", 8012, 8099)
        httpd, port = server.create_http_server("127.0.0.1", fallback_port, server.Handler)

    url = "http://127.0.0.1:%s/" % port
    print("Strategy Excel Tool")
    print("Open URL: %s" % url)
    print("Private templates folder: %s" % server.TEMPLATE_DIR)
    print("Outputs folder: %s" % server.OUTPUT_DIR)
    print("Close this window to exit.")

    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
        time.sleep(0.2)


if __name__ == "__main__":
    main()
