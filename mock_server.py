#!/usr/bin/env python3
"""Dependency-free Microsoft identity + Graph/SharePoint mock for GoalLens.

This is a contract simulator, not an emulator. It exposes only the endpoints
needed by the GoalLens SharePoint connector and lets tests switch deterministic
scenarios through a local admin endpoint.
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import threading
import time
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse


ROOT = Path(__file__).resolve().parent
FIXTURES = ROOT / "fixtures"

SITE_ID = "goallensdev.sharepoint.com,site-collection-id,site-id"
DRIVE_ID = "drive-goallens-documents"
ROOT_FOLDER_ID = "root"

SCENARIOS = {
    "happy_path",
    "permission_denied",
    "invalid_token",
    "throttled_once",
    "expired_delta",
    "new_version",
    "deleted",
    "restored",
    "download_failure",
}


class MockState:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.scenario = "happy_path"
        self.request_counts: dict[str, int] = {}

    def reset(self, scenario: str = "happy_path") -> None:
        with self._lock:
            self.scenario = scenario
            self.request_counts = {}

    def record(self, key: str) -> int:
        with self._lock:
            count = self.request_counts.get(key, 0) + 1
            self.request_counts[key] = count
            return count

    def snapshot(self) -> dict:
        with self._lock:
            return {
                "scenario": self.scenario,
                "request_counts": dict(self.request_counts),
                "available_scenarios": sorted(SCENARIOS),
            }


def _item(item_id: str, name: str, etag: str, size: int, mime: str) -> dict:
    return {
        "id": item_id,
        "name": name,
        "eTag": etag,
        "size": size,
        "webUrl": f"https://goallensdev.sharepoint.com/sites/GoalLensSandbox/Shared Documents/{name}",
        "lastModifiedDateTime": "2026-09-19T15:30:00Z",
        "parentReference": {"driveId": DRIVE_ID, "id": ROOT_FOLDER_ID},
        "file": {"mimeType": mime},
    }


class GraphMockHandler(BaseHTTPRequestHandler):
    server_version = "GoalLensGraphMock/1.0"

    @property
    def state(self) -> MockState:
        return self.server.state  # type: ignore[attr-defined]

    def log_message(self, fmt: str, *args) -> None:
        if getattr(self.server, "quiet", False):  # type: ignore[attr-defined]
            return
        super().log_message(fmt, *args)

    def _base_url(self) -> str:
        return f"http://{self.headers.get('Host', '127.0.0.1:8765')}"

    def _json(self, status: int, body: dict | list, headers: dict | None = None) -> None:
        payload = json.dumps(body).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(payload)

    def _error(self, status: int, code: str, message: str, headers: dict | None = None) -> None:
        self._json(status, {"error": {"code": code, "message": message}}, headers)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def _require_auth(self) -> bool:
        if self.state.scenario == "invalid_token":
            self._error(HTTPStatus.UNAUTHORIZED, "InvalidAuthenticationToken", "Mock token rejected")
            return False
        if not self.headers.get("Authorization", "").startswith("Bearer "):
            self._error(HTTPStatus.UNAUTHORIZED, "InvalidAuthenticationToken", "Bearer token required")
            return False
        return True

    def _maybe_throttle(self, key: str) -> bool:
        count = self.state.record(key)
        if self.state.scenario == "throttled_once" and count == 1:
            self._error(
                HTTPStatus.TOO_MANY_REQUESTS,
                "TooManyRequests",
                "Mock throttling response",
                {"Retry-After": "1"},
            )
            return True
        return False

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/__admin/scenario":
            body = self._read_json()
            scenario = body.get("scenario")
            if scenario not in SCENARIOS:
                self._json(HTTPStatus.BAD_REQUEST, {"error": "unknown scenario", "available": sorted(SCENARIOS)})
                return
            self.state.reset(scenario)
            self._json(HTTPStatus.OK, self.state.snapshot())
            return

        if path == "/__admin/reset":
            self.state.reset()
            self._json(HTTPStatus.OK, self.state.snapshot())
            return

        if path.endswith("/oauth2/v2.0/token"):
            self._json(
                HTTPStatus.OK,
                {
                    "token_type": "Bearer",
                    "expires_in": 3600,
                    "ext_expires_in": 3600,
                    "access_token": "mock-sharepoint-access-token",
                },
            )
            return

        self._error(HTTPStatus.NOT_FOUND, "NotFound", f"No mock POST route for {path}")

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        query = parse_qs(parsed.query)

        if path == "/health":
            self._json(HTTPStatus.OK, {"status": "ok", "scenario": self.state.scenario})
            return
        if path == "/__admin/state":
            self._json(HTTPStatus.OK, self.state.snapshot())
            return
        if not path.startswith("/v1.0/"):
            self._error(HTTPStatus.NOT_FOUND, "NotFound", f"No mock GET route for {path}")
            return
        if not self._require_auth():
            return
        if self.state.scenario == "permission_denied":
            self._error(HTTPStatus.FORBIDDEN, "accessDenied", "The app has no grant on this site")
            return
        if self._maybe_throttle(f"GET {path}"):
            return

        if path.startswith("/v1.0/sites/") and ":/sites/" in path:
            self._json(
                HTTPStatus.OK,
                {
                    "id": SITE_ID,
                    "name": "GoalLensSandbox",
                    "displayName": "GoalLens Connector Sandbox",
                    "webUrl": "https://goallensdev.sharepoint.com/sites/GoalLensSandbox",
                },
            )
            return

        if path == f"/v1.0/sites/{SITE_ID}/drives":
            self._json(
                HTTPStatus.OK,
                {
                    "value": [
                        {
                            "id": DRIVE_ID,
                            "name": "Documents",
                            "driveType": "documentLibrary",
                            "webUrl": "https://goallensdev.sharepoint.com/sites/GoalLensSandbox/Shared Documents",
                        }
                    ]
                },
            )
            return

        if path == f"/v1.0/drives/{DRIVE_ID}/items/{ROOT_FOLDER_ID}/children":
            self._json(HTTPStatus.OK, {"value": self._initial_items()})
            return

        if path == f"/v1.0/drives/{DRIVE_ID}/root/delta":
            self._handle_delta(query)
            return

        prefix = f"/v1.0/drives/{DRIVE_ID}/items/"
        if path.startswith(prefix) and path.endswith("/content"):
            item_id = path[len(prefix) : -len("/content")]
            self._send_content(item_id)
            return

        self._error(HTTPStatus.NOT_FOUND, "itemNotFound", f"No Graph fixture for {path}")

    def _initial_items(self) -> list[dict]:
        return [
            _item("file-sales-plan", "Sales Growth Plan.txt", '"etag-sales-v1"', 182, "text/plain"),
            _item("file-scorecard", "Revenue KPI Scorecard.csv", '"etag-scorecard-v1"', 151, "text/csv"),
        ]

    def _handle_delta(self, query: dict[str, list[str]]) -> None:
        if self.state.scenario == "expired_delta" and "token" in query:
            self._error(HTTPStatus.GONE, "resyncRequired", "The delta token has expired")
            return

        base = self._base_url()
        if query.get("page") == ["2"]:
            self._json(
                HTTPStatus.OK,
                {
                    "value": [self._initial_items()[1]],
                    "@odata.deltaLink": f"{base}/v1.0/drives/{DRIVE_ID}/root/delta?token=baseline",
                },
            )
            return

        if "token" not in query:
            self._json(
                HTTPStatus.OK,
                {
                    "value": [self._initial_items()[0]],
                    "@odata.nextLink": f"{base}/v1.0/drives/{DRIVE_ID}/root/delta?page=2",
                },
            )
            return

        scenario = self.state.scenario
        if scenario == "new_version":
            value = [_item("file-scorecard", "Revenue KPI Scorecard.csv", '"etag-scorecard-v2"', 171, "text/csv")]
            token = "after-new-version"
        elif scenario == "deleted":
            value = [{"id": "file-scorecard", "deleted": {"state": "deleted"}}]
            token = "after-delete"
        elif scenario == "restored":
            value = [_item("file-scorecard", "Revenue KPI Scorecard.csv", '"etag-scorecard-v1"', 151, "text/csv")]
            token = "after-restore"
        else:
            value = []
            token = "no-changes"
        self._json(
            HTTPStatus.OK,
            {"value": value, "@odata.deltaLink": f"{base}/v1.0/drives/{DRIVE_ID}/root/delta?token={token}"},
        )

    def _send_content(self, item_id: str) -> None:
        if self.state.scenario == "download_failure":
            self._error(HTTPStatus.INTERNAL_SERVER_ERROR, "serviceNotAvailable", "Mock download failure")
            return
        names = {
            "file-sales-plan": "sales-growth-plan.txt",
            "file-scorecard": "revenue-kpi-scorecard.csv",
        }
        fixture_name = names.get(item_id)
        if not fixture_name:
            self._error(HTTPStatus.NOT_FOUND, "itemNotFound", "Unknown mock file")
            return
        if item_id == "file-scorecard" and self.state.scenario == "new_version":
            fixture_name = "revenue-kpi-scorecard-v2.csv"
        payload = (FIXTURES / fixture_name).read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mimetypes.guess_type(fixture_name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def create_server(host: str = "127.0.0.1", port: int = 8765, quiet: bool = False) -> ThreadingHTTPServer:
    server = ThreadingHTTPServer((host, port), GraphMockHandler)
    server.state = MockState()  # type: ignore[attr-defined]
    server.quiet = quiet  # type: ignore[attr-defined]
    return server


def main() -> None:
    parser = argparse.ArgumentParser(description="GoalLens Microsoft Graph/SharePoint mock")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args()
    server = create_server(args.host, args.port, args.quiet)
    print(f"GoalLens SharePoint mock listening on http://{args.host}:{server.server_port}", flush=True)
    print("Admin state: GET /__admin/state", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
