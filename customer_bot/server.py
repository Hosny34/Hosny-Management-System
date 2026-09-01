from __future__ import annotations

import argparse
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict
from urllib.parse import parse_qs, urlparse

from .bot import ArabicCustomerBot
from .config import load_config
from .queries import WarehouseCustomerQueries
from .whatsapp import extract_incoming_messages, send_text


def _json_bytes(payload: Dict[str, Any], status: int = 200) -> tuple[int, bytes]:
    return status, json.dumps(payload, ensure_ascii=False).encode("utf-8")


class BotHandler(BaseHTTPRequestHandler):
    queries = WarehouseCustomerQueries(load_config())
    bot = ArabicCustomerBot(queries)

    def _send_json(self, payload: Dict[str, Any], status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_text(self, text: str, status: int = 200) -> None:
        body = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _query(self) -> Dict[str, str]:
        parsed = urlparse(self.path)
        return {k: (v[0] if v else "") for k, v in parse_qs(parsed.query).items()}

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        q = self._query()
        if parsed.path == "/health":
            cfg = self.queries.config
            self._send_json(
                {
                    "ok": True,
                    "warehouse_db_path": str(cfg.warehouse_db_path),
                    "stock_stale_minutes": cfg.stock_stale_minutes,
                    "branches": len(cfg.branches),
                }
            )
            return
        if parsed.path == "/api/branches":
            self._send_json({"branches": self.queries.branch_info(q.get("branch") or None)})
            return
        if parsed.path == "/api/stock":
            self._send_json(
                {
                    "rows": self.queries.search_stock(
                        item_type=q.get("item_type", ""),
                        school=q.get("school", ""),
                        color=q.get("color", ""),
                        size=q.get("size", ""),
                        limit=int(q.get("limit") or 30),
                    )
                }
            )
            return
        if parsed.path == "/api/price":
            self._send_json(
                {
                    "rows": self.queries.search_prices(
                        item_type=q.get("item_type", ""),
                        school=q.get("school", ""),
                        color=q.get("color", ""),
                        size=q.get("size", ""),
                        limit=int(q.get("limit") or 30),
                    )
                }
            )
            return
        if parsed.path == "/api/reservation":
            row = self.queries.reservation_status(branch=q.get("branch", ""), bill_number=q.get("bill_number", ""))
            self._send_json(row or {"error": "reservation not found"}, 200 if row else 404)
            return
        if parsed.path == "/whatsapp/webhook":
            import os

            expected = os.environ.get("WHATSAPP_VERIFY_TOKEN", "").strip()
            if q.get("hub.mode") == "subscribe" and expected and q.get("hub.verify_token") == expected:
                self._send_text(q.get("hub.challenge", ""))
            else:
                self._send_json({"error": "invalid verify token"}, 403)
            return
        self._send_json({"error": "not found"}, 404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b"{}"
        try:
            payload = json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError:
            self._send_json({"error": "invalid json"}, 400)
            return
        if parsed.path == "/api/chat":
            message = str(payload.get("message") or "")
            self._send_json({"parsed": self.bot.parse(message), "reply": self.bot.reply(message)})
            return
        if parsed.path == "/whatsapp/webhook":
            sent = []
            for msg in extract_incoming_messages(payload):
                reply = self.bot.reply(msg["text"])
                sent.append({"to": msg["from"], "reply": reply, "send_result": send_text(msg["from"], reply)})
            self._send_json({"ok": True, "messages": len(sent), "sent": sent})
            return
        self._send_json({"error": "not found"}, 404)

    def log_message(self, fmt: str, *args: Any) -> None:
        return


def main() -> int:
    parser = argparse.ArgumentParser(description="Run Hosny customer bot local server.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8090)
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), BotHandler)
    print(f"Customer bot listening on http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

