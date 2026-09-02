# Hosny Customer WhatsApp Bot

Local-first customer bot service for answering Arabic WhatsApp questions from
the Warehouse mirror database.

This first version is intentionally separate from the POS and Warehouse desktop
apps. It reads Warehouse data in read-only mode and can be tested locally before
connecting Meta WhatsApp Cloud API.

For production, the bot can run on Railway and read a cached stock copy from
the sync server. The Warehouse laptop refreshes that copy with a separate
uploader whenever it is online.

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

## Upload Stock Cache

Set the same secret token on Railway sync server and on the Warehouse laptop:

```powershell
$env:CUSTOMER_STOCK_UPLOAD_TOKEN="choose-a-long-random-secret"
```

Dry-run from the laptop:

```powershell
python -m customer_bot.upload_stock_snapshot --dry-run
```

Upload to the sync server:

```powershell
python -m customer_bot.upload_stock_snapshot --server-url https://web-production-e022.up.railway.app
```

The upload reads `Warehouse/warehouse_data.sqlite3` in read-only mode and sends
only branch stock rows: branch device, item type, school, color, size, POS
price, count, and sync time.

## Railway Bot Data Source

When the bot runs on Railway, point it at the cached stock API:

```text
CUSTOMER_STOCK_API_URL=https://web-production-e022.up.railway.app
```

If this variable is not set, the bot uses the local Warehouse DB path instead.

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

## Twilio WhatsApp Setup

For Twilio Sandbox or a production Twilio WhatsApp sender, set the incoming
message webhook to:

```text
https://acceptable-alignment-production-dae8.up.railway.app/twilio/webhook
```

Use HTTP `POST`. The bot replies with TwiML, so no Twilio secret is required for
basic inbound customer replies.
