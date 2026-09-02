from __future__ import annotations

import sqlite3
import tempfile
import unittest
import os
from pathlib import Path

from customer_bot.bot import ArabicCustomerBot
from customer_bot.config import BotConfig
from customer_bot.queries import WarehouseCustomerQueries
from customer_bot.whatsapp import extract_twilio_message, twiml_message


class TestCustomerBot(unittest.TestCase):
    def setUp(self) -> None:
        fd, raw_path = tempfile.mkstemp(prefix="customer_bot_", suffix=".sqlite3")
        os.close(fd)
        self.path = Path(raw_path)
        self.addCleanup(lambda: self.path.exists() and self.path.unlink())
        conn = sqlite3.connect(self.path)
        with conn:
            conn.executescript(
                """
                CREATE TABLE pos_stocks_mirror (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source_device TEXT NOT NULL,
                    item_type TEXT NOT NULL,
                    school TEXT NOT NULL,
                    color TEXT NOT NULL,
                    size TEXT NOT NULL,
                    unit_price REAL NOT NULL,
                    count INTEGER NOT NULL,
                    snapshot_at TEXT NOT NULL
                );
                CREATE TABLE pos_stocks_snapshot_meta (
                    source_device TEXT PRIMARY KEY,
                    snapshot_at TEXT NOT NULL,
                    row_count INTEGER NOT NULL DEFAULT 0,
                    total_value REAL NOT NULL DEFAULT 0,
                    app_version TEXT
                );
                CREATE TABLE pos_reservations_mirror (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source_device TEXT NOT NULL,
                    reservation_key TEXT NOT NULL,
                    customer TEXT,
                    item_type TEXT NOT NULL,
                    school TEXT NOT NULL,
                    color TEXT NOT NULL,
                    size TEXT NOT NULL,
                    qty INTEGER NOT NULL,
                    unit_price REAL NOT NULL,
                    total_amount REAL NOT NULL,
                    paid_amount REAL NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'معلق',
                    shift_id INTEGER,
                    last_event_uuid TEXT,
                    updated_at TEXT NOT NULL,
                    UNIQUE(source_device, reservation_key)
                );
                """
            )
            conn.execute(
                """
                INSERT INTO pos_stocks_mirror
                    (source_device,item_type,school,color,size,unit_price,count,snapshot_at)
                VALUES ('POS-OBO','تيشيرت خريفي','البيان','رمادي','10',315,3,'2099-01-01T10:00:00Z')
                """
            )
            conn.execute(
                """
                INSERT INTO pos_stocks_snapshot_meta(source_device,snapshot_at,row_count,total_value)
                VALUES ('POS-OBO','2099-01-01T10:00:00Z',1,945)
                """
            )
            conn.execute(
                """
                INSERT INTO pos_reservations_mirror
                    (source_device,reservation_key,customer,item_type,school,color,size,qty,unit_price,total_amount,paid_amount,status,updated_at)
                VALUES ('POS-CEN','id:558','عميل اختبار','تيشيرت صيفي','رجاك','احمر','10',2,270,540,200,'معلق','2099-01-01T10:00:00Z')
                """
            )
        conn.close()

        cfg = BotConfig(
            warehouse_db_path=self.path,
            stock_stale_minutes=30,
            branches=[
                {
                    "device": "POS-OBO",
                    "name": "فرع العبور",
                    "address": "عنوان العبور",
                    "phone": "010",
                    "maps_url": "https://maps.example/obo",
                    "hours": "من 10 إلى 10",
                },
                {
                    "device": "POS-CEN",
                    "name": "فرع السنتر",
                    "address": "عنوان السنتر",
                    "phone": "011",
                    "maps_url": "https://maps.example/cen",
                    "hours": "من 10 إلى 10",
                },
            ],
        )
        self.bot = ArabicCustomerBot(WarehouseCustomerQueries(cfg))

    def test_stock_reply_uses_all_branches_and_pos_price(self):
        reply = self.bot.reply("عندكم البيان تيشيرت خريفي مقاس 10؟")
        self.assertIn("فرع العبور", reply)
        self.assertIn("متوفر 3 قطعة", reply)
        self.assertIn("السعر 315 جنيه", reply)

    def test_eis_text_does_not_match_size_s(self):
        parsed = self.bot.parse("عندكم EIS ابتدائي تيشيرت خريفي مقاس 10؟")
        self.assertEqual(parsed["size"], "10")
        self.assertNotEqual(parsed["size"], "S")

    def test_twilio_helpers_extract_message_and_escape_xml(self):
        msg = extract_twilio_message(
            {
                "From": "whatsapp:+201111111111",
                "Body": "stock question",
                "MessageSid": "SM123",
            }
        )
        self.assertEqual(msg["from"], "+201111111111")
        self.assertEqual(msg["text"], "stock question")
        self.assertEqual(msg["id"], "SM123")
        self.assertIn("&lt;test&gt;", twiml_message("<test>"))

    def test_branch_info_reply_is_arabic(self):
        reply = self.bot.reply("عنوان فرع العبور")
        self.assertIn("عنوان العبور", reply)
        self.assertIn("الخريطة", reply)

    def test_reservation_lookup_requires_branch_and_bill_number(self):
        reply = self.bot.reply("فرع السنتر حجز 558")
        self.assertIn("حجز رقم 558", reply)
        self.assertIn("الحالة: معلق", reply)
        self.assertIn("المتبقي: 340 جنيه", reply)


if __name__ == "__main__":
    unittest.main()
