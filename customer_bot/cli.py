from __future__ import annotations

import sys

from .bot import ArabicCustomerBot
from .runtime import make_queries


def main() -> int:
    message = " ".join(sys.argv[1:]).strip()
    if not message:
        print("اكتب رسالة العميل بعد الأمر.")
        return 2
    bot = ArabicCustomerBot(make_queries())
    print(bot.reply(message))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
