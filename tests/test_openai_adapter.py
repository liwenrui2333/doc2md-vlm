"""Tests for the OpenAI-compatible vision adapter, run against a local mock."""

from __future__ import annotations

import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import doc2md
from tests.fixtures import make_pdf

VALID_KEY = "test-key-123"
REPLY = "【MOCK】柱状图，显示各渠道花费。"


class Handler(BaseHTTPRequestHandler):
    """Rejects anything that is not a well-formed vision request."""

    received: list[dict] = []
    fail_auth = False

    def log_message(self, *args):
        pass

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
        Handler.received.append(body)
        content = body.get("messages", [{}])[0].get("content")
        has_image = isinstance(content, list) and any(
            part.get("type") == "image_url"
            and part["image_url"]["url"].startswith("data:image/png;base64,")
            for part in content
        )
        has_text = isinstance(content, list) and any(part.get("type") == "text" for part in content)
        if not (has_image and has_text):
            self.send_response(400)
            self.end_headers()
            self.wfile.write(json.dumps({"error": "bad request shape"}).encode())
            return
        if Handler.fail_auth or self.headers.get("Authorization") != f"Bearer {VALID_KEY}":
            self.send_response(401)
            self.end_headers()
            self.wfile.write(json.dumps({"error": "bad key"}).encode())
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps({"choices": [{"message": {"content": REPLY}}]}).encode())


@pytest.fixture
def server():
    Handler.received = []
    Handler.fail_auth = False
    srv = HTTPServer(("127.0.0.1", 0), Handler)
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{srv.server_port}/v1/chat/completions"
    srv.shutdown()
    srv.server_close()


def test_openai_adapter_round_trip(server, tmp_path: Path):
    pdf = make_pdf(tmp_path / "report.pdf", pages=1)
    code = doc2md.main(
        [
            str(pdf),
            "--charts",
            "--local-llm",
            "--vlm-provider",
            "openai",
            "--vlm-url",
            server,
            "--vlm-api-key",
            VALID_KEY,
        ]
    )
    assert code == 0
    md = (tmp_path / "report_doc2md" / "report.md").read_text(encoding="utf-8")
    assert REPLY in md

    sent = Handler.received[0]
    assert sent["model"] == "gpt-4o-mini"
    assert sent["messages"][0]["content"][0]["type"] == "text"


def test_missing_key_sends_no_auth_header(server, tmp_path: Path):
    pdf = make_pdf(tmp_path / "report.pdf", pages=1)
    doc2md.main(
        [
            str(pdf),
            "--charts",
            "--local-llm",
            "--vlm-provider",
            "openai",
            "--vlm-url",
            server,
            "--vlm-api-key",
            VALID_KEY,
        ]
    )
    assert Handler.received, "server was never called"


def test_bad_key_reports_401_with_fix(server, tmp_path: Path):
    Handler.fail_auth = True
    pdf = make_pdf(tmp_path / "report.pdf", pages=1)
    code = doc2md.main(
        [
            str(pdf),
            "--charts",
            "--local-llm",
            "--vlm-provider",
            "openai",
            "--vlm-url",
            server,
            "--vlm-api-key",
            "wrong",
        ]
    )
    assert code == 0
    md = (tmp_path / "report_doc2md" / "report.md").read_text(encoding="utf-8")
    assert "401" in md
    assert "DOC2MD_VLM_API_KEY" in md
