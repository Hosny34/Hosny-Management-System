from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Protocol

try:
    from .bot import DIGIT_TRANSLATION, money, normalize_text
    from .queries import branch_display_name
except ImportError:
    from bot import DIGIT_TRANSLATION, money, normalize_text
    from queries import branch_display_name


class CustomerQueries(Protocol):
    def branch_info(self, branch: str | None = None) -> List[Dict[str, Any]]:
        ...

    def distinct_values(
        self,
        field: str,
        *,
        source_device: str = "",
        school: str = "",
        item_type: str = "",
        color: str = "",
        min_count: int = 1,
        limit: int = 200,
    ) -> List[str]:
        ...

    def search_stock(
        self,
        *,
        item_type: str = "",
        school: str = "",
        color: str = "",
        size: str = "",
        source_device: str = "",
        min_count: int = 1,
        limit: int = 30,
    ) -> List[Dict[str, Any]]:
        ...


@dataclass
class MenuSession:
    step: str = "main"
    data: Dict[str, str] = field(default_factory=dict)
    options: List[Dict[str, str]] = field(default_factory=list)


class MenuConversationBot:
    def __init__(self, queries: CustomerQueries) -> None:
        self.queries = queries
        self.sessions: Dict[str, MenuSession] = {}

    def reply(self, user_id: str, message: str) -> str:
        key = str(user_id or "default").strip() or "default"
        text = normalize_text(message)
        if self._is_main_menu_request(text):
            return self._reset_to_main(key)
        session = self.sessions.get(key)
        if session is None or session.step == "main":
            return self._handle_main(key, text)
        if session.step == "branch":
            return self._handle_branch(key, text)
        if session.step == "school":
            return self._handle_school(key, text)
        if session.step == "item_type":
            return self._handle_item_type(key, text)
        if session.step == "color":
            return self._handle_color(key, text)
        if session.step == "size":
            return self._handle_size(key, text)
        return self._reset_to_main(key)

    def _is_main_menu_request(self, text: str) -> bool:
        return text in {"0", "قائمه", "قائمة", "منيو", "menu", "main", "الرئيسيه", "الرئيسية"}

    def _reset_to_main(self, key: str) -> str:
        self.sessions[key] = MenuSession(step="main")
        return self._main_menu()

    def _main_menu(self) -> str:
        return "\n".join(
            [
                "أهلاً بحضرتك. اختار رقم من القائمة:",
                "1. البحث عن صنف",
                "2. عناوين الفروع",
                "0. القائمة الرئيسية",
            ]
        )

    def _handle_main(self, key: str, text: str) -> str:
        if text in {"1", "بحث", "صنف", "stock"}:
            return self._show_branches(key)
        if text in {"2", "فروع", "عنوان", "عناوين", "branches"}:
            rows = self.queries.branch_info(None)
            lines = ["عناوين الفروع:"]
            for row in rows:
                lines.append(
                    f"{row['name']}\n"
                    f"العنوان: {row['address']}\n"
                    f"التليفون: {row['phone']}\n"
                    f"المواعيد: {row['hours']}"
                )
            lines.append("0. القائمة الرئيسية")
            self.sessions[key] = MenuSession(step="main")
            return "\n\n".join(lines)
        return self._main_menu()

    def _show_branches(self, key: str) -> str:
        branches = [
            {"id": str(row["device"]), "title": str(row["name"])}
            for row in self.queries.branch_info(None)
            if row.get("device")
        ]
        self.sessions[key] = MenuSession(step="branch", options=branches)
        return self._numbered("اختار الفرع:", branches)

    def _handle_branch(self, key: str, text: str) -> str:
        choice = self._choice(self.sessions[key], text)
        if choice is None:
            return self._invalid_choice(self.sessions[key], "اختار رقم الفرع من القائمة:")
        session = MenuSession(step="school", data={"branch": choice["id"]})
        self.sessions[key] = session
        return self._show_schools(key)

    def _show_schools(self, key: str) -> str:
        session = self.sessions[key]
        schools = self.queries.distinct_values(
            "school",
            source_device=session.data["branch"],
            min_count=0,
            limit=80,
        )
        options = [{"id": value, "title": value} for value in schools]
        session.options = options
        if not options:
            self.sessions[key] = MenuSession(step="main")
            return "لا يوجد مخزون متاح لهذا الفرع حالياً.\n\n" + self._main_menu()
        return self._numbered("اختار المدرسة:", options)

    def _handle_school(self, key: str, text: str) -> str:
        session = self.sessions[key]
        choice = self._choice(session, text)
        if choice is None:
            return self._invalid_choice(session, "اختار رقم المدرسة من القائمة:")
        session.data["school"] = choice["id"]
        has_available_stock = self.queries.search_stock(
            source_device=session.data["branch"],
            school=session.data["school"],
            min_count=1,
            limit=1,
        )
        if not has_available_stock:
            self.sessions[key] = MenuSession(step="main")
            return (
                "غير متوفر حالياً لهذه المدرسة في الفرع المختار.\n"
                "ممكن تسأل مرة أخرى خلال يومين.\n\n"
                + self._main_menu()
            )
        session.step = "item_type"
        return self._show_item_types(key)

    def _show_item_types(self, key: str) -> str:
        session = self.sessions[key]
        item_types = self.queries.distinct_values(
            "item_type",
            source_device=session.data["branch"],
            school=session.data["school"],
            min_count=1,
            limit=50,
        )
        session.options = [{"id": value, "title": value} for value in item_types]
        if not session.options:
            self.sessions[key] = MenuSession(step="main")
            return "لا يوجد أصناف متاحة لهذه المدرسة في الفرع المختار.\n\n" + self._main_menu()
        return self._numbered("اختار النوع:", session.options)

    def _handle_item_type(self, key: str, text: str) -> str:
        session = self.sessions[key]
        choice = self._choice(session, text)
        if choice is None:
            return self._invalid_choice(session, "اختار رقم النوع من القائمة:")
        session.data["item_type"] = choice["id"]
        colors = self.queries.distinct_values(
            "color",
            source_device=session.data["branch"],
            school=session.data["school"],
            item_type=session.data["item_type"],
            min_count=1,
            limit=50,
        )
        if len(colors) <= 1:
            session.data["color"] = colors[0] if colors else ""
            session.step = "size"
            session.options = []
            return self._ask_size(session)
        session.step = "color"
        session.options = [{"id": value, "title": value} for value in colors]
        return self._numbered("اختار اللون:", session.options)

    def _handle_color(self, key: str, text: str) -> str:
        session = self.sessions[key]
        choice = self._choice(session, text)
        if choice is None:
            return self._invalid_choice(session, "اختار رقم اللون من القائمة:")
        session.data["color"] = choice["id"]
        session.step = "size"
        session.options = []
        return self._ask_size(session)

    def _ask_size(self, session: MenuSession) -> str:
        branch = branch_display_name(session.data.get("branch"))
        return (
            f"الفرع: {branch}\n"
            f"المدرسة: {session.data.get('school', '')}\n"
            f"النوع: {session.data.get('item_type', '')}\n"
            "اكتب المقاس المطلوب فقط.\n"
            "0. القائمة الرئيسية"
        )

    def _handle_size(self, key: str, text: str) -> str:
        session = self.sessions[key]
        size = text.translate(DIGIT_TRANSLATION).strip().upper()
        if not size:
            return "اكتب المقاس المطلوب فقط.\n0. القائمة الرئيسية"
        rows = self.queries.search_stock(
            source_device=session.data.get("branch", ""),
            school=session.data.get("school", ""),
            item_type=session.data.get("item_type", ""),
            color=session.data.get("color", ""),
            size=size,
            min_count=1,
            limit=20,
        )
        self.sessions[key] = MenuSession(step="main")
        if not rows:
            return "غير متوفر حالياً في الفرع المختار.\n\n" + self._main_menu()
        lines = ["النتيجة:"]
        for row in rows[:10]:
            stale = " - آخر تحديث قديم، يفضل التأكيد قبل الذهاب" if row.get("stale") else ""
            lines.append(
                f"{row['branch']}: متوفر {row['count']} قطعة - "
                f"{row['item_type']} {row['school']} {row['color']} مقاس {row['size']} - "
                f"السعر {money(row['unit_price'])} جنيه{stale}"
            )
        lines.append("")
        lines.append(self._main_menu())
        return "\n".join(lines)

    def _choice(self, session: MenuSession, text: str) -> Dict[str, str] | None:
        clean_text = text.translate(DIGIT_TRANSLATION).strip()
        if clean_text.isdigit():
            index = int(clean_text)
            if 1 <= index <= len(session.options):
                return session.options[index - 1]
        normalized = normalize_text(clean_text)
        for option in session.options:
            if normalize_text(option["title"]) == normalized:
                return option
        return None

    def _invalid_choice(self, session: MenuSession, heading: str) -> str:
        return "اختيار غير صحيح.\n\n" + self._numbered(heading, session.options)

    def _numbered(self, heading: str, options: List[Dict[str, str]]) -> str:
        lines = [heading]
        for index, option in enumerate(options, start=1):
            lines.append(f"{index}. {option['title']}")
        lines.append("0. القائمة الرئيسية")
        return "\n".join(lines)
