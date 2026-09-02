from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

try:
    from .queries import BRANCH_NAMES, WarehouseCustomerQueries, canonical_branch
except ImportError:
    from queries import BRANCH_NAMES, WarehouseCustomerQueries, canonical_branch


DIGIT_TRANSLATION = str.maketrans(
    {
        "\u0660": "0",
        "\u0661": "1",
        "\u0662": "2",
        "\u0663": "3",
        "\u0664": "4",
        "\u0665": "5",
        "\u0666": "6",
        "\u0667": "7",
        "\u0668": "8",
        "\u0669": "9",
        "\u06f0": "0",
        "\u06f1": "1",
        "\u06f2": "2",
        "\u06f3": "3",
        "\u06f4": "4",
        "\u06f5": "5",
        "\u06f6": "6",
        "\u06f7": "7",
        "\u06f8": "8",
        "\u06f9": "9",
    }
)


def money(value: Any) -> str:
    try:
        return str(int(round(float(value or 0))))
    except (TypeError, ValueError):
        return "0"


def normalize_text(value: Any) -> str:
    text = str(value or "").strip().translate(DIGIT_TRANSLATION)
    text = re.sub(r"[إأآا]", "ا", text)
    text = text.replace("ى", "ي").replace("ة", "ه")
    text = re.sub(r"\s+", " ", text)
    return text


def _find_number(text: str) -> str:
    m = re.search(r"(?:حجز|فاتوره|فاتورة|رقم|#)?\s*(\d{1,8})", text)
    return m.group(1) if m else ""


def _find_size(text: str) -> str:
    prefixed = re.search(
        r"(?:مقاس|مقاسه|سايز|size)\s*(\d{1,2}|[SMLX]{1,3})\b",
        text,
        flags=re.IGNORECASE,
    )
    if prefixed:
        return prefixed.group(1).upper()
    numeric = re.search(r"(?<![\wء-ي])(\d{1,2})(?![\wء-ي])", text, flags=re.IGNORECASE)
    return numeric.group(1).upper() if numeric else ""


def _match_branch(text: str) -> str:
    norm = normalize_text(text)
    for device, name in BRANCH_NAMES.items():
        if normalize_text(name) in norm or normalize_text(name.replace("فرع ", "")) in norm:
            return device
    if "سنتر" in norm or "الشارع الجديد" in norm or "center" in norm:
        return "POS-CEN"
    if "عبور" in norm:
        return "POS-OBO"
    if "زايد" in norm:
        return "POS-ZAY"
    if "بهتيم" in norm:
        return "POS-BAH"
    if "اكتوبر" in norm or "اكتوبر" in norm:
        return "POS-OCT"
    if "جسر" in norm or "سويس" in norm:
        return "POS-GESR"
    return ""


def _match_known(text: str, values: List[str]) -> str:
    norm = normalize_text(text)
    for value in sorted(values, key=len, reverse=True):
        if value and normalize_text(value) in norm:
            return value
    return ""


@dataclass
class ArabicCustomerBot:
    queries: WarehouseCustomerQueries

    def parse(self, message: str) -> Dict[str, str]:
        text = normalize_text(message)
        intent = "unknown"
        if any(w in text for w in ("عنوان", "مكان", "لوكيشن", "location", "فرع")) and not any(w in text for w in ("حجز", "فاتوره", "فاتورة")):
            intent = "branch"
        if any(w in text for w in ("متوفر", "موجود", "عندكم", "فيه", "في ")) and intent == "unknown":
            intent = "stock"
        if any(w in text for w in ("سعر", "بكام", "كام", "ثمن")):
            intent = "price" if intent == "unknown" else intent
        if any(w in text for w in ("حجز", "فاتوره", "فاتورة")):
            intent = "reservation"

        school = _match_known(message, self.queries.known_values("school"))
        item_type = _match_known(message, self.queries.known_values("item_type"))
        color = _match_known(message, self.queries.known_values("color"))
        if color and not any(w in text for w in ("لون", "اللون")):
            norm_color = normalize_text(color)
            if norm_color in normalize_text(item_type) or norm_color in normalize_text(school):
                color = ""

        return {
            "intent": intent,
            "branch": _match_branch(message),
            "bill_number": _find_number(text),
            "size": _find_size(text),
            "school": school,
            "item_type": item_type,
            "color": color,
        }

    def reply(self, message: str) -> str:
        parsed = self.parse(message)
        intent = parsed["intent"]
        if intent == "branch":
            return self._reply_branch(parsed)
        if intent == "reservation":
            return self._reply_reservation(parsed)
        if intent in {"stock", "price"}:
            return self._reply_stock_or_price(parsed, price_only=(intent == "price"))
        return (
            "أهلاً بحضرتك. ممكن تسألني عن توفر صنف، سعر صنف، عنوان فرع، أو حالة حجز.\n"
            "مثال: عندكم تيشيرت صيفي رجاك مقاس 14؟"
        )

    def _reply_branch(self, parsed: Dict[str, str]) -> str:
        rows = self.queries.branch_info(parsed.get("branch") or None)
        if not rows:
            return "لم أجد بيانات هذا الفرع حالياً."
        lines = []
        for b in rows:
            lines.append(
                f"{b['name']}\n"
                f"العنوان: {b['address']}\n"
                f"التليفون: {b['phone']}\n"
                f"المواعيد: {b['hours']}\n"
                f"الخريطة: {b['maps_url']}"
            )
        return "\n\n".join(lines)

    def _reply_reservation(self, parsed: Dict[str, str]) -> str:
        if not parsed.get("branch"):
            return "من فضلك اكتب اسم الفرع أولاً ثم رقم الحجز. مثال: فرع الشارع الجديد حجز 558"
        if not parsed.get("bill_number"):
            return "من فضلك اكتب رقم الحجز بعد اسم الفرع. مثال: فرع العبور حجز 558"
        row = self.queries.reservation_status(branch=parsed["branch"], bill_number=parsed["bill_number"])
        if not row:
            return "لم أجد هذا الحجز في الفرع المحدد. تأكد من اسم الفرع ورقم الحجز."
        items = "، ".join(
            f"{x['item_type']} {x['school']} {x['color']} مقاس {x['size']} عدد {x['qty']}"
            for x in row["items"][:5]
        )
        return (
            f"حجز رقم {row['bill_number']} في {row['branch']}\n"
            f"الحالة: {row['status']}\n"
            f"الإجمالي: {money(row['total'])} جنيه\n"
            f"المدفوع: {money(row['paid'])} جنيه\n"
            f"المتبقي: {money(row['remaining'])} جنيه\n"
            f"الأصناف: {items or 'غير متاحة'}"
        )

    def _reply_stock_or_price(self, parsed: Dict[str, str], *, price_only: bool = False) -> str:
        filters = {
            "item_type": parsed.get("item_type", ""),
            "school": parsed.get("school", ""),
            "color": parsed.get("color", ""),
            "size": parsed.get("size", ""),
        }
        if not any(filters.values()):
            return "من فضلك اكتب المدرسة أو النوع أو اللون أو المقاس المطلوب."
        rows = self.queries.search_stock(**filters, min_count=0 if price_only else 1, limit=20)
        if not rows:
            return "لم أجد نتيجة مطابقة حالياً. ممكن تكتب المدرسة والنوع والمقاس بشكل أوضح؟"
        lines = []
        for r in rows[:10]:
            stale = " - آخر تحديث قديم، يفضل التأكيد قبل الذهاب" if r.get("stale") else ""
            if price_only:
                lines.append(
                    f"{r['branch']}: {r['item_type']} {r['school']} {r['color']} مقاس {r['size']} السعر {money(r['unit_price'])} جنيه{stale}"
                )
            else:
                lines.append(
                    f"{r['branch']}: متوفر {r['count']} قطعة - {r['item_type']} {r['school']} {r['color']} مقاس {r['size']} - السعر {money(r['unit_price'])} جنيه{stale}"
                )
        if len(rows) > 10:
            lines.append(f"ويوجد {len(rows) - 10} نتيجة أخرى. ممكن تحدد اللون أو المقاس أكثر.")
        return "\n".join(lines)
