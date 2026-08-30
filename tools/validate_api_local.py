"""在临时端口调用统一 API，避免干扰用户正在使用的 Dashboard 进程。"""
import json
import sys
import threading
import urllib.request
from http.server import ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from dashboard_server import Handler

server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
thread = threading.Thread(target=server.serve_forever, daemon=True); thread.start()
try:
    url = f"http://127.0.0.1:{server.server_port}/api/summary?shop=MIYO&start=2026-08-13&end=2026-08-19"
    with urllib.request.urlopen(url, timeout=30) as response:
        payload = json.load(response)
    print(json.dumps({"status": "ok", "metric_version": payload["meta"]["metric_version"], "gmv": payload["metrics"]["gmv"], "refund_amount": payload["metrics"]["refund_amount"], "comparison_period": payload["comparison_period"]}, ensure_ascii=False))
finally:
    server.shutdown(); server.server_close()
