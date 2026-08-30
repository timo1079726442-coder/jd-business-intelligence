"""唯一的 Dashboard 启动入口；端口占用时安全退出，不杀其他 Python 进程。"""
import os
import socket
from pathlib import Path

from dashboard_server import Handler, SERVICE_VERSION, METRIC_VERSION
from http.server import ThreadingHTTPServer

ROOT = Path(__file__).resolve().parent
PORT = 18766
PID_FILE = ROOT / "dashboard.pid"

def main():
    probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        probe.bind(("127.0.0.1", PORT))
    except OSError:
        print(f"Dashboard 已占用 127.0.0.1:{PORT}，未启动第二个实例。请先确认现有 /api/health。")
        return 2
    finally:
        probe.close()
    PID_FILE.write_text(str(os.getpid()), encoding="utf-8")
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"Dashboard PID={os.getpid()} port={PORT} service_version={SERVICE_VERSION} metric_version={METRIC_VERSION}", flush=True)
    try:
        server.serve_forever()
    finally:
        server.server_close()
        if PID_FILE.exists() and PID_FILE.read_text(encoding="utf-8").strip() == str(os.getpid()):
            PID_FILE.unlink()

if __name__ == "__main__":
    raise SystemExit(main())
