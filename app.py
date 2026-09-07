"""排课 Web 应用：基于会话隔离的多用户版本。"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
from collections import deque
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

import excel_io
import scheduler


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "基本数据.xlsx"
STATIC_DIR = BASE_DIR / "static"
SESSION_ROOT = BASE_DIR / "sessions"
USAGE_FILE = BASE_DIR / "页面使用说明.json"

SESSION_TTL_SECONDS = 3600
MAX_SESSIONS = 200
MAX_UPLOAD_BYTES = 1 * 1024 * 1024
SOLVE_LOCK = threading.Lock()
QUEUE_COND = threading.Condition()
SOLVE_QUEUE = deque()
SOLVE_TASKS = {}
SOLVE_TASK_COUNTER = 0


class SolveTask:
    def __init__(self, task_id, body):
        self.id = task_id
        self.body = body
        self.status = "queued"
        self.position = 1
        self.result = None


def _start_solve_task(body):
    global SOLVE_TASK_COUNTER
    with QUEUE_COND:
        SOLVE_TASK_COUNTER += 1
        task_id = uuid.uuid4().hex
        task = SolveTask(task_id, body)
        SOLVE_TASKS[task_id] = task
        SOLVE_QUEUE.append(task)
        task.position = len(SOLVE_QUEUE)
    threading.Thread(target=_execute_solve_task, args=(task,), daemon=True).start()
    return task


def _execute_solve_task(task):
    with QUEUE_COND:
        while SOLVE_QUEUE and SOLVE_QUEUE[0] is not task:
            QUEUE_COND.wait()
        task.status = "running"
        for idx, item in enumerate(SOLVE_QUEUE, start=1):
            item.position = idx
    try:
        body = task.body
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
        with SOLVE_LOCK:
            result = scheduler.solve_schedule(data, settings, variant=variant)
        task.result = result
    except Exception as exc:  # noqa: BLE001
        task.result = {"ok": False, "errorKind": "server", "errors": [str(exc)]}
    finally:
        task.status = "done" if task.result is not None else "error"
        with QUEUE_COND:
            if SOLVE_QUEUE and SOLVE_QUEUE[0] is task:
                SOLVE_QUEUE.popleft()
            else:
                try:
                    SOLVE_QUEUE.remove(task)
                except ValueError:
                    pass
            for idx, item in enumerate(SOLVE_QUEUE, start=1):
                item.position = idx
            QUEUE_COND.notify_all()
        if len(SOLVE_TASKS) > 200:
            for old_id in list(SOLVE_TASKS)[:50]:
                old_task = SOLVE_TASKS.get(old_id)
                if old_task and old_task.status not in ("queued", "running"):
                    SOLVE_TASKS.pop(old_id, None)


def _load_uploaded_data(filename, content_b64):
    """解析用户上传的 Excel 内容，不持久保存到服务器。"""
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


def _cleanup_sessions():
    """删除空闲超过1小时的会话，并在会话过多时清理最旧的。"""
    SESSION_ROOT.mkdir(exist_ok=True)
    now = time.time()
    entries = []
    for child in SESSION_ROOT.iterdir():
        if not child.is_dir():
            continue
        try:
            mtime = child.stat().st_mtime
        except OSError:
            continue
        if now - mtime > SESSION_TTL_SECONDS:
            shutil.rmtree(child, ignore_errors=True)
        else:
            entries.append((mtime, child))
    entries.sort()
    while len(entries) > MAX_SESSIONS:
        _, old = entries.pop(0)
        shutil.rmtree(old, ignore_errors=True)


class ScheduleHandler(BaseHTTPRequestHandler):
    server_version = "PaikeLocal/1.0"
    _cleanup_counter = 0

    def _send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self._send_session_cookie()
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
        self._send_session_cookie()
        self.end_headers()
        self.wfile.write(data)

    def _send_session_cookie(self):
        if getattr(self, "session_id", None):
            self.send_header("Set-Cookie", f"session_id={self.session_id}; Path=/; HttpOnly; SameSite=Lax")

    def _cookie_value(self, name):
        cookie = self.headers.get("Cookie") or ""
        for part in cookie.split(";"):
            part = part.strip()
            if part.startswith(name + "="):
                return part[len(name) + 1:]
        return None

    def _ensure_session(self):
        sid = self._cookie_value("session_id")
        valid_sid = bool(sid) and bool(re.fullmatch(r"[0-9a-f]{32}", sid))
        session_dir = SESSION_ROOT / sid if valid_sid else None
        if session_dir is None or not session_dir.is_dir():
            sid = uuid.uuid4().hex
            session_dir = SESSION_ROOT / sid
            session_dir.mkdir(parents=True, exist_ok=True)
        else:
            try:
                os.utime(session_dir, None)
            except OSError:
                pass
        self.session_id = sid
        ScheduleHandler._cleanup_counter += 1
        if ScheduleHandler._cleanup_counter % 20 == 0:
            _cleanup_sessions()
        return session_dir

    def _session_dir(self):
        return SESSION_ROOT / self.session_id

    def _data_file(self):
        return DATA_FILE

    def _settings_file(self):
        return self._session_dir() / "settings.json"

    def _output_dir(self):
        return self._session_dir() / "output"

    def _read_body(self):
        length = int(self.headers.get("Content-Length", "0") or "0")
        if length <= 0:
            return {}
        if length > 8 * 1024 * 1024:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return {}

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path in ("/", "/index.html"):
            self._ensure_session()
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
        if path == "/api/data":
            self._ensure_session()
            self._send_json(self._data_payload())
            return
        if path == "/api/settings":
            self._ensure_session()
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
        if path == "/api/task":
            query = parse_qs(urlparse(self.path).query)
            task_id = (query.get("taskId") or [None])[0]
            task = SOLVE_TASKS.get(task_id)
            if not task:
                self._send_json({"ok": False, "error": "任务不存在或已过期"}, 404)
                return
            if task.status in ("queued", "running"):
                self._send_json(
                    {
                        "ok": True,
                        "status": task.status,
                        "position": task.position,
                    }
                )
            else:
                self._send_json(task.result or {"ok": False, "error": "任务无结果"})
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
        session_dir = self._ensure_session()
        if path in ("/api/upload", "/api/parse"):
            filename = str(body.get("filename") or "upload.xlsx")
            content_b64 = body.get("contentBase64") or ""
            try:
                data = _load_uploaded_data(filename, content_b64)
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
            output_dir = session_dir / "output"
            if output_dir.exists():
                shutil.rmtree(output_dir, ignore_errors=True)
            self._send_json(self._data_payload())
            return
        if path == "/api/solve":
            task = _start_solve_task(body)
            self._send_json(
                {
                    "ok": True,
                    "queued": True,
                    "taskId": task.id,
                    "position": task.position,
                }
            )
            return
        if path == "/api/save":
            self._send_json({"ok": True})
            return
        if path == "/api/export":
            file_b64 = body.get("fileBase64")
            if file_b64:
                try:
                    data = _load_uploaded_data(
                        str(body.get("filename") or "upload.xlsx"), file_b64
                    )
                except ValueError as exc:
                    self._send_json({"error": str(exc)}, 400)
                    return
            else:
                data = scheduler.load_schedule_data(DATA_FILE)
            settings = scheduler.normalize_settings(body.get("settings"), data)
            result = body.get("result") or {}
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffix = uuid.uuid4().hex[:6]
            issues, files = excel_io.write_result_files(
                data, settings, result, self._output_dir(), timestamp, suffix
            )
            payload = {"ok": True, "files": files, "issues": issues}
            self._send_json(payload)
            return
        if path == "/api/download":
            filename = body.get("filename") or ""
            if not filename or ".." in filename:
                self._send_json({"error": "bad filename"}, 400)
                return
            target = (self._output_dir() / filename).resolve()
            output_dir = self._output_dir().resolve()
            if output_dir not in target.parents:
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
            data = scheduler.load_schedule_data(self._data_file())
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        settings = scheduler.normalize_settings({}, data)
        path = DATA_FILE
        display_name = DATA_FILE.name
        return {
            "ok": True,
            "data": data.to_dict(),
            "settings": scheduler.settings_to_json(settings),
            "dataSource": {
                "path": str(path),
                "name": display_name,
                "isDemo": True,
            },
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
    SESSION_ROOT.mkdir(exist_ok=True)
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
