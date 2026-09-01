from __future__ import annotations

import sys

from .bot import ArabicCustomerBot
from .config import load_config
from .queries import WarehouseCustomerQueries


def main() -> int:
    message = " ".join(sys.argv[1:]).strip()
    if not message:
        print("اكتب رسالة العميل بعد الأمر.")
        return 2
    bot = ArabicCustomerBot(WarehouseCustomerQueries(load_config()))
    print(bot.reply(message))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

