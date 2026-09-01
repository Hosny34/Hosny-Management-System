from __future__ import annotations

import os
from typing import Any, Dict

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import PlainTextResponse
from pydantic import BaseModel

try:
    from .bot import ArabicCustomerBot
    from .config import load_config
    from .runtime import make_queries
    from .whatsapp import extract_incoming_messages, send_text
except ImportError:
    from bot import ArabicCustomerBot
    from config import load_config
    from runtime import make_queries
    from whatsapp import extract_incoming_messages, send_text


config = load_config()
queries = make_queries(config)
bot = ArabicCustomerBot(queries)
app = FastAPI(title="Hosny Customer Bot", version="0.1.0")


class ChatRequest(BaseModel):
    message: str


@app.get("/health")
def health() -> Dict[str, Any]:
    return {
        "ok": True,
        "warehouse_db_path": str(config.warehouse_db_path),
        "customer_stock_api_url": config.customer_stock_api_url,
        "stock_stale_minutes": config.stock_stale_minutes,
        "branches": len(config.branches),
    }


@app.get("/api/branches")
def api_branches(branch: str = "") -> Dict[str, Any]:
    return {"branches": queries.branch_info(branch or None)}


@app.get("/api/stock")
def api_stock(
    item_type: str = "",
    school: str = "",
    color: str = "",
    size: str = "",
    limit: int = Query(30, ge=1, le=100),
) -> Dict[str, Any]:
    return {
        "rows": queries.search_stock(
            item_type=item_type,
            school=school,
            color=color,
            size=size,
            limit=limit,
        )
    }


@app.get("/api/price")
def api_price(
    item_type: str = "",
    school: str = "",
    color: str = "",
    size: str = "",
    limit: int = Query(30, ge=1, le=100),
) -> Dict[str, Any]:
    return {
        "rows": queries.search_prices(
            item_type=item_type,
            school=school,
            color=color,
            size=size,
            limit=limit,
        )
    }


@app.get("/api/reservation")
def api_reservation(branch: str, bill_number: str) -> Dict[str, Any]:
    row = queries.reservation_status(branch=branch, bill_number=bill_number)
    if not row:
        raise HTTPException(status_code=404, detail="reservation not found")
    return row


@app.post("/api/chat")
def api_chat(req: ChatRequest) -> Dict[str, Any]:
    parsed = bot.parse(req.message)
    return {"parsed": parsed, "reply": bot.reply(req.message)}


@app.get("/whatsapp/webhook", response_class=PlainTextResponse)
def verify_webhook(
    hub_mode: str = Query("", alias="hub.mode"),
    hub_verify_token: str = Query("", alias="hub.verify_token"),
    hub_challenge: str = Query("", alias="hub.challenge"),
) -> str:
    expected = os.environ.get("WHATSAPP_VERIFY_TOKEN", "").strip()
    if hub_mode == "subscribe" and expected and hub_verify_token == expected:
        return hub_challenge
    raise HTTPException(status_code=403, detail="invalid verify token")


@app.post("/whatsapp/webhook")
async def receive_webhook(request: Request) -> Dict[str, Any]:
    payload = await request.json()
    sent = []
    for msg in extract_incoming_messages(payload):
        reply = bot.reply(msg["text"])
        sent.append({"to": msg["from"], "reply": reply, "send_result": send_text(msg["from"], reply)})
    return {"ok": True, "messages": len(sent), "sent": sent}
