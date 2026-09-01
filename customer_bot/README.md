# Hosny Customer WhatsApp Bot

Local-first customer bot service for answering Arabic WhatsApp questions from
the Warehouse mirror database.

This first version is intentionally separate from the POS and Warehouse desktop
apps. It reads Warehouse data in read-only mode and can be tested locally before
connecting Meta WhatsApp Cloud API.

## Run Locally

No extra install is needed for the local test server:

```powershell
cd G:\TestManagementSystem\Hosny-Management-System
python -m customer_bot.server --host 127.0.0.1 --port 8090
```

FastAPI support also exists in `customer_bot.app`, but requires installing
`customer_bot/requirements.txt`.

Open:

```text
http://127.0.0.1:8090/health
```

## Local Chat Test

```powershell
python -m customer_bot.cli "عندكم رجاك تيشيرت صيفي مقاس 14؟"
python -m customer_bot.cli "عنوان فرع العبور"
python -m customer_bot.cli "فرع السنتر حجز 558"
```

## Current Rules

- Replies are Arabic only.
- Stock searches show all branches.
- Prices are POS mirror prices, not Warehouse profile prices.
- If a branch stock snapshot is older than 30 minutes, the reply warns the
  customer that the update is old.
- Reservation lookup requires branch first, then bill number. The current
  Warehouse reservation mirror may not yet expose all visible POS bill numbers,
  so some old/new UUID-keyed reservations may not be found until we add that
  mapping.

## WhatsApp Setup Later

Set these environment variables before enabling real WhatsApp replies:

```text
WHATSAPP_ACCESS_TOKEN
WHATSAPP_PHONE_NUMBER_ID
WHATSAPP_VERIFY_TOKEN
WAREHOUSE_DB_PATH
```
