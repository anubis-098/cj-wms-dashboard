from __future__ import annotations

import argparse
import json
import os
import shutil
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


def mirror_latest(source: Path, destination: Path) -> dict[str, object]:
    candidates = [
        path for path in source.iterdir()
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(f"No .xlsx files found in {source}")
    latest = max(candidates, key=lambda path: (path.stat().st_mtime_ns, path.name.lower()))
    source_stat = latest.stat()
    destination.mkdir(parents=True, exist_ok=True)
    target = destination / latest.name
    if target.exists():
        target_stat = target.stat()
        if target_stat.st_size == source_stat.st_size and target_stat.st_mtime_ns == source_stat.st_mtime_ns:
            return {"changed": False, "filename": latest.name, "size": source_stat.st_size}

    file_descriptor, temporary_name = tempfile.mkstemp(prefix=".smd-", suffix=".tmp", dir=destination)
    os.close(file_descriptor)
    temporary = Path(temporary_name)
    try:
        shutil.copy2(latest, temporary)
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)
    return {"changed": True, "filename": latest.name, "size": source_stat.st_size}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", default=r"\\10.84.194.51\CJWMSDashboard")
    parser.add_argument("--destination", required=True)
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--token", required=True)
    args = parser.parse_args()
    source = Path(args.source)
    destination = Path(args.destination).resolve()

    class Handler(BaseHTTPRequestHandler):
        def send_json(self, status: int, payload: dict[str, object]) -> None:
            body = json.dumps(payload).encode()
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:  # noqa: N802
            self.send_json(200, {"status": "ok"})

        def do_POST(self) -> None:  # noqa: N802
            if self.path != "/sync":
                self.send_json(404, {"detail": "Not found"})
                return
            if self.headers.get("X-Sync-Token") != args.token:
                self.send_json(403, {"detail": "Forbidden"})
                return
            try:
                self.send_json(200, {"status": "success", "data": mirror_latest(source, destination)})
            except Exception as exc:
                self.send_json(500, {"detail": str(exc)})

        def log_message(self, format: str, *values: object) -> None:
            return

    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
