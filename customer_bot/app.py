from __future__ import annotations

import os
from urllib.parse import parse_qs
from typing import Any, Dict

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, PlainTextResponse, Response
from pydantic import BaseModel

try:
    from .bot import ArabicCustomerBot
    from .config import load_config
    from .conversation import MenuConversationBot
    from .runtime import make_queries
    from .whatsapp import extract_incoming_messages, extract_twilio_message, send_text, twiml_message
except ImportError:
    from bot import ArabicCustomerBot
    from config import load_config
    from conversation import MenuConversationBot
    from runtime import make_queries
    from whatsapp import extract_incoming_messages, extract_twilio_message, send_text, twiml_message


config = load_config()
queries = make_queries(config)
bot = ArabicCustomerBot(queries)
menu_bot = MenuConversationBot(queries)
app = FastAPI(title="Hosny Customer Bot", version="0.1.0")


class ChatRequest(BaseModel):
    message: str


def _business_home_html() -> str:
    branch_items = "\n".join(
        f"""
        <li>
          <strong>{branch.get("name", "")}</strong>
          <span>{branch.get("hours", "")}</span>
        </li>
        """
        for branch in config.branches
    )
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Hosny Uniform | حسني للزي المدرسي</title>
  <style>
    :root {{ color-scheme: light; font-family: Tahoma, Arial, sans-serif; }}
    body {{ margin: 0; background: #f4f7fb; color: #172033; line-height: 1.7; }}
    main {{ max-width: 980px; margin: 0 auto; padding: 36px 18px 56px; }}
    header {{ background: #12345a; color: #fff; padding: 36px 22px; border-radius: 8px; }}
    h1 {{ margin: 0 0 8px; font-size: clamp(30px, 6vw, 52px); letter-spacing: 0; }}
    h2 {{ margin: 0 0 12px; font-size: 24px; }}
    section {{ margin-top: 18px; padding: 22px; background: #fff; border: 1px solid #d9e1ec; border-radius: 8px; }}
    ul {{ margin: 0; padding: 0 20px 0 0; }}
    li {{ margin: 8px 0; }}
    li span {{ display: block; color: #59677b; }}
    a {{ color: #0b65c2; }}
    .contact {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 14px; }}
    .contact a {{ display: inline-block; padding: 10px 14px; border-radius: 6px; background: #0b65c2; color: #fff; text-decoration: none; }}
    footer {{ margin-top: 18px; color: #59677b; font-size: 14px; }}
  </style>
</head>
<body>
  <main>
    <header>
      <h1>Hosny Uniform</h1>
      <p>حسني للزي المدرسي - زي مدرسي ومستلزمات مدارس في مصر.</p>
      <div class="contact">
        <a href="https://wa.me/201063062890">واتساب: 01063062890</a>
        <a href="/privacy">سياسة الخصوصية</a>
      </div>
    </header>
    <section>
      <h2>عن النشاط</h2>
      <p>نساعد أولياء الأمور والطلاب في معرفة توفر الزي المدرسي والمقاسات والأسعار في الفروع. خدمة واتساب مخصصة للرد على استفسارات العملاء عن المخزون والفروع.</p>
    </section>
    <section>
      <h2>الفروع</h2>
      <ul>{branch_items}</ul>
    </section>
    <section>
      <h2>التواصل</h2>
      <p>للاستفسار عن توفر المقاسات أو الأسعار، تواصل معنا على واتساب رقم 01063062890.</p>
    </section>
    <footer>© Hosny Uniform - حسني للزي المدرسي</footer>
  </main>
</body>
</html>"""


def _privacy_html() -> str:
    return """<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>سياسة الخصوصية | Hosny Uniform</title>
  <style>
    body { margin: 0; background: #f4f7fb; color: #172033; font-family: Tahoma, Arial, sans-serif; line-height: 1.8; }
    main { max-width: 900px; margin: 0 auto; padding: 34px 18px 54px; }
    section { background: #fff; border: 1px solid #d9e1ec; border-radius: 8px; padding: 24px; }
    h1 { margin-top: 0; letter-spacing: 0; }
    a { color: #0b65c2; }
  </style>
</head>
<body>
  <main>
    <section>
      <h1>سياسة الخصوصية</h1>
      <p>نستخدم بيانات الرسائل الواردة عبر واتساب للرد على استفسارات العملاء عن الفروع، الأسعار، وتوفر المخزون.</p>
      <p>قد نحتفظ برقم الهاتف ومحتوى المحادثة ووقت الرسالة لأغراض خدمة العملاء وتحسين الردود. لا نبيع بيانات العملاء ولا نشاركها مع أطراف خارجية إلا مزودي الخدمة اللازمين لتشغيل واتساب والردود الآلية.</p>
      <p>يمكنك طلب حذف بيانات محادثتك أو إيقاف التواصل معنا عبر إرسال رسالة على واتساب إلى 01063062890.</p>
      <p><a href="/">العودة للرئيسية</a></p>
    </section>
  </main>
</body>
</html>"""


@app.get("/", response_class=HTMLResponse)
def homepage() -> str:
    return _business_home_html()


@app.get("/privacy", response_class=HTMLResponse)
def privacy() -> str:
    return _privacy_html()


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
        reply = menu_bot.reply(msg["from"], msg["text"])
        sent.append({"to": msg["from"], "reply": reply, "send_result": send_text(msg["from"], reply)})
    return {"ok": True, "messages": len(sent), "sent": sent}


@app.post("/twilio/webhook")
async def receive_twilio_webhook(request: Request) -> Response:
    raw_body = (await request.body()).decode("utf-8", errors="replace")
    values = parse_qs(raw_body, keep_blank_values=True)
    form = {key: vals[-1] if vals else "" for key, vals in values.items()}
    msg = extract_twilio_message(form)
    reply = menu_bot.reply(msg["from"], msg["text"]) if msg["text"] else menu_bot.reply(msg["from"], "قائمة")
    return Response(content=twiml_message(reply), media_type="application/xml")
