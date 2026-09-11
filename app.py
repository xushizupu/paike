"""排课 Web 应用：浏览器保存数据，服务器同步排课。"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import sys
import tempfile
import threading
import uuid
import webbrowser
from collections import deque
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlparse

import excel_io
import scheduler


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

DATA_FILE = BASE_DIR / "基本数据.xlsx"
STATIC_DIR = BASE_DIR / "static"
USAGE_FILE = BASE_DIR / "页面使用说明.json"
MAX_UPLOAD_BYTES = 1 * 1024 * 1024

QUEUE_COND = threading.Condition()
QUEUE_ORDER = deque()
QUEUE_COUNTER = 0


def _load_uploaded_data(filename, content_b64):
    """解析上传文件内容，不在服务器持久保存。"""
    if not filename.lower().endswith(".xlsx"):
        raise ValueError("请上传 .xlsx 格式的模板文件")
    try:
        content = base64.b64decode(content_b64 or "")
    except (ValueError, TypeError) as exc:
        raise ValueError("文件内容无法解析") from exc
    if not content or len(content) > MAX_UPLOAD_BYTES:
        raise ValueError("文件大小不能超过 1MB")
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(content)
            tmp_path = Path(tmp.name)
        return scheduler.load_schedule_data(tmp_path)
    finally:
        if tmp_path is not None:
            try:
                tmp_path.unlink(missing_ok=True)
            except OSError:
                pass


def _enqueue_solve():
    global QUEUE_COUNTER
    with QUEUE_COND:
        QUEUE_COUNTER += 1
        ticket = QUEUE_COUNTER
        QUEUE_ORDER.append(ticket)
    return ticket


def _wait_for_turn(ticket):
    with QUEUE_COND:
        while QUEUE_ORDER and QUEUE_ORDER[0] != ticket:
            QUEUE_COND.wait()


def _release_solve(ticket):
    with QUEUE_COND:
        try:
            QUEUE_ORDER.remove(ticket)
        except ValueError:
            pass
        QUEUE_COND.notify_all()


def _queue_snapshot():
    with QUEUE_COND:
        total = len(QUEUE_ORDER)
    return {
        "total": total,
        "running": 1 if total else 0,
        "waiting": max(0, total - 1),
    }


def _solve_request_body(body):
    file_b64 = body.get("fileBase64")
    if file_b64:
        data = _load_uploaded_data(str(body.get("filename") or "upload.xlsx"), file_b64)
    else:
        data = scheduler.load_schedule_data(DATA_FILE)
    settings = scheduler.normalize_settings(body.get("settings"), data)
    variant = body.get("variant")
    try:
        variant = int(variant)
    except (TypeError, ValueError):
        variant = None
    return scheduler.solve_schedule(data, settings, variant=variant)


class ScheduleHandler(BaseHTTPRequestHandler):
    server_version = "PaikeLocal/1.0"

    def _send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_bytes(self, data, content_type, filename=None, inline=False, cache_control=None):
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        if filename:
            disposition = "inline" if inline else "attachment"
            ascii_name = filename.encode("ascii", "ignore").decode() or "download"
            encoded_name = quote(filename)
            self.send_header(
                "Content-Disposition",
                f'{disposition}; filename="{ascii_name}"; filename*=UTF-8\'\'{encoded_name}',
            )
        if cache_control:
            self.send_header("Cache-Control", cache_control)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _read_body(self):
        length = int(self.headers.get("Content-Length", "0") or "0")
        if length <= 0:
            return {}
        if length > 16 * 1024 * 1024:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return {}

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path in ("/", "/index.html"):
            self._serve_file(STATIC_DIR / "index.html")
            return
        if path.startswith("/static/"):
            relative = path[len("/static/"):]
            target = (STATIC_DIR / relative).resolve()
            if STATIC_DIR.resolve() not in target.parents and target != STATIC_DIR.resolve():
                self._send_json({"error": "forbidden"}, 403)
                return
            self._serve_file(target)
            return
        if path == "/api/ping":
            self._send_json({"ok": True, "pong": True})
            return
        if path == "/api/queue-status":
            self._send_json({"ok": True, **_queue_snapshot()})
            return
        if path in ("/api/data", "/api/settings"):
            self._send_json(self._data_payload())
            return
        if path == "/api/usage":
            usage = {}
            if USAGE_FILE.exists():
                try:
                    usage = json.loads(USAGE_FILE.read_text(encoding="utf-8"))
                except (ValueError, json.JSONDecodeError):
                    usage = {}
            self._send_json({"ok": True, "usage": usage})
            return
        if path == "/api/template":
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            self._send_bytes(DATA_FILE.read_bytes(), content_type, "paike_template.xlsx")
            return
        self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        body = self._read_body()
        if path in ("/api/upload", "/api/parse"):
            filename = str(body.get("filename") or "upload.xlsx")
            try:
                data = _load_uploaded_data(filename, body.get("contentBase64") or "")
            except ValueError as exc:
                self._send_json({"error": str(exc)}, 400)
                return
            settings = scheduler.normalize_settings({}, data)
            self._send_json(
                {
                    "ok": True,
                    "data": data.to_dict(),
                    "settings": scheduler.settings_to_json(settings),
                    "dataSource": {"name": Path(filename).name, "isDemo": False},
                }
            )
            return
        if path == "/api/reset-data":
            self._send_json(self._data_payload())
            return
        if path in ("/api/solve", "/api/solve-sync"):
            ticket = _enqueue_solve()
            try:
                _wait_for_turn(ticket)
                result = _solve_request_body(body)
            except ValueError as exc:
                result = {"ok": False, "errorKind": "validation", "errors": [str(exc)]}
            except Exception as exc:  # noqa: BLE001
                result = {"ok": False, "errorKind": "server", "errors": [str(exc)]}
            finally:
                _release_solve(ticket)
            self._send_json(result)
            return
        if path == "/api/export":
            file_b64 = body.get("fileBase64")
            try:
                data = (
                    _load_uploaded_data(str(body.get("filename") or "upload.xlsx"), file_b64)
                    if file_b64
                    else scheduler.load_schedule_data(DATA_FILE)
                )
            except ValueError as exc:
                self._send_json({"ok": False, "error": str(exc)}, 400)
                return
            settings = scheduler.normalize_settings(body.get("settings"), data)
            result = body.get("result") or {}
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffix = uuid.uuid4().hex[:6]
            issues, files = excel_io.write_result_files_memory(
                data, settings, result, timestamp, suffix
            )
            self._send_json({"ok": True, "files": files, "issues": issues})
            return
        self._send_json({"error": "not found"}, 404)

    def _data_payload(self):
        try:
            data = scheduler.load_schedule_data(DATA_FILE)
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        settings = scheduler.normalize_settings({}, data)
        return {
            "ok": True,
            "data": data.to_dict(),
            "settings": scheduler.settings_to_json(settings),
            "dataSource": {"name": DATA_FILE.name, "isDemo": True},
        }

    def _serve_file(self, target):
        if not target.exists() or not target.is_file():
            self._send_json({"error": "not found"}, 404)
            return
        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        self._send_bytes(target.read_bytes(), content_type, target.name, inline=True, cache_control="no-store")

    def log_message(self, format, *args):  # noqa: A002
        pass


def main():
    requested_port = int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "0.0.0.0")
    port = requested_port
    for candidate in range(requested_port, requested_port + 10):
        try:
            server = ThreadingHTTPServer((host, candidate), ScheduleHandler)
            port = candidate
            break
        except OSError:
            continue
    print(f"排课系统已启动：http://127.0.0.1:{port}")
    print("按 Ctrl+C 停止服务。")
    if not os.environ.get("PORT"):
        threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")


if __name__ == "__main__":
    main()
