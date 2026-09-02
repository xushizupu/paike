"""本地排课 Web 应用。"""

from __future__ import annotations

import json
import base64
import mimetypes
import os
import socketserver
import uuid
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlparse

import excel_io
import scheduler


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "基本数据.xlsx"
CONFIG_FILE = BASE_DIR / "排课设置.json"
OUTPUT_DIR = BASE_DIR / "output"
STATIC_DIR = BASE_DIR / "static"
UPLOAD_DIR = BASE_DIR / "uploads"
DATA_SOURCE_FILE = BASE_DIR / "data_source.json"


def get_data_file():
    if DATA_SOURCE_FILE.exists():
        try:
            info = json.loads(DATA_SOURCE_FILE.read_text(encoding="utf-8"))
            candidate = BASE_DIR / info["path"]
            if candidate.exists():
                return candidate
        except (ValueError, KeyError, TypeError):
            pass
    return DATA_FILE


def get_data_info():
    path = get_data_file()
    return {
        "path": str(path),
        "name": path.name,
        "isDemo": path.resolve() == DATA_FILE.resolve(),
    }


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
        if path == "/api/data":
            self._send_json(self._data_payload())
            return
        if path == "/api/settings":
            self._send_json(self._data_payload())
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
        if path == "/api/upload":
            filename = str(body.get("filename") or "upload.xlsx")
            content_b64 = body.get("contentBase64") or ""
            try:
                content = base64.b64decode(content_b64)
            except (ValueError, TypeError):
                self._send_json({"error": "文件内容无法解析"}, 400)
                return
            UPLOAD_DIR.mkdir(exist_ok=True)
            safe_name = Path(filename).name
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            target = UPLOAD_DIR / f"{timestamp}_{safe_name}"
            target.write_bytes(content)
            try:
                data = scheduler.load_schedule_data(target)
            except Exception as exc:  # noqa: BLE001
                target.unlink(missing_ok=True)
                self._send_json({"error": f"模板格式不正确：{exc}"}, 400)
                return
            DATA_SOURCE_FILE.write_text(
                json.dumps({"path": str(target.relative_to(BASE_DIR))}, ensure_ascii=False),
                encoding="utf-8",
            )
            settings = scheduler.normalize_settings({}, data)
            scheduler.save_settings(CONFIG_FILE, settings)
            self._send_json(self._data_payload())
            return
        if path == "/api/reset-data":
            if DATA_SOURCE_FILE.exists():
                try:
                    info = json.loads(DATA_SOURCE_FILE.read_text(encoding="utf-8"))
                    uploaded = BASE_DIR / info["path"]
                    if UPLOAD_DIR.resolve() in uploaded.resolve().parents:
                        uploaded.unlink(missing_ok=True)
                except (ValueError, KeyError, TypeError):
                    pass
                DATA_SOURCE_FILE.unlink(missing_ok=True)
            data = scheduler.load_schedule_data(DATA_FILE)
            settings = scheduler.normalize_settings({}, data)
            scheduler.save_settings(CONFIG_FILE, settings)
            self._send_json(self._data_payload())
            return
        if path == "/api/solve":
            data = scheduler.load_schedule_data(get_data_file())
            settings = scheduler.normalize_settings(body.get("settings"), data)
            variant = body.get("variant")
            try:
                variant = int(variant)
            except (TypeError, ValueError):
                variant = None
            result = scheduler.solve_schedule(data, settings, variant=variant)
            self._send_json(result)
            return
        if path == "/api/save":
            data = scheduler.load_schedule_data(get_data_file())
            settings = scheduler.normalize_settings(body.get("settings"), data)
            scheduler.save_settings(CONFIG_FILE, settings)
            self._send_json({"ok": True})
            return
        if path == "/api/export":
            data = scheduler.load_schedule_data(get_data_file())
            settings = scheduler.normalize_settings(body.get("settings"), data)
            result = body.get("result") or {}
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffix = uuid.uuid4().hex[:6]
            issues, files = excel_io.write_result_files(
                data, settings, result, OUTPUT_DIR, timestamp, suffix
            )
            payload = {"ok": True, "files": files, "issues": issues}
            self._send_json(payload)
            return
        if path == "/api/download":
            filename = body.get("filename") or ""
            if not filename or ".." in filename:
                self._send_json({"error": "bad filename"}, 400)
                return
            target = (OUTPUT_DIR / filename).resolve()
            if OUTPUT_DIR.resolve() not in target.parents:
                self._send_json({"error": "forbidden"}, 403)
                return
            if not target.exists():
                self._send_json({"error": "not found"}, 404)
                return
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            self._send_bytes(target.read_bytes(), content_type, filename)
            return
        self._send_json({"error": "not found"}, 404)

    def _data_payload(self):
        try:
            data = scheduler.load_schedule_data(get_data_file())
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        settings = scheduler.load_settings(CONFIG_FILE, data)
        return {
            "ok": True,
            "data": data.to_dict(),
            "settings": scheduler.settings_to_json(settings),
            "dataSource": get_data_info(),
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
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")


if __name__ == "__main__":
    main()
