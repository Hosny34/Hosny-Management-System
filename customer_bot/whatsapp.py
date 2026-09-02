from __future__ import annotations

import os
from html import escape
from typing import Any, Dict, Iterable, List

def extract_incoming_messages(payload: Dict[str, Any]) -> List[Dict[str, str]]:
    messages: List[Dict[str, str]] = []
    for entry in payload.get("entry") or []:
        for change in entry.get("changes") or []:
            value = change.get("value") or {}
            for msg in value.get("messages") or []:
                text = ((msg.get("text") or {}).get("body") or "").strip()
                sender = str(msg.get("from") or "").strip()
                message_id = str(msg.get("id") or "").strip()
                if sender and text:
                    messages.append({"from": sender, "text": text, "id": message_id})
    return messages


def extract_twilio_message(form: Dict[str, str]) -> Dict[str, str]:
    sender = str(form.get("From") or form.get("WaId") or "").strip()
    if sender.startswith("whatsapp:"):
        sender = sender[len("whatsapp:") :]
    return {
        "from": sender,
        "text": str(form.get("Body") or "").strip(),
        "id": str(form.get("MessageSid") or form.get("SmsMessageSid") or "").strip(),
    }


def twiml_message(body: str) -> str:
    safe_body = escape(body[:1600], quote=False)
    return f'<?xml version="1.0" encoding="UTF-8"?><Response><Message>{safe_body}</Message></Response>'


def send_text(to: str, body: str) -> Dict[str, Any]:
    import requests

    token = os.environ.get("WHATSAPP_ACCESS_TOKEN", "").strip()
    phone_id = os.environ.get("WHATSAPP_PHONE_NUMBER_ID", "").strip()
    if not token or not phone_id:
        return {"sent": False, "reason": "WhatsApp credentials are not configured"}
    url = f"https://graph.facebook.com/v20.0/{phone_id}/messages"
    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "messaging_product": "whatsapp",
            "to": to,
            "type": "text",
            "text": {"body": body[:4000]},
        },
        timeout=20,
    )
    return {"sent": resp.ok, "status_code": resp.status_code, "response": resp.text}
