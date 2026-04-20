#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SERVER_DB="${SERVER_DB:-$SCRIPT_DIR/../sync_server/Hosny-sync-server/sync_server.sqlite3}"

usage() {
  cat <<'EOF'
Usage:
  bash reset_sync_server.sh pos-zay
  bash reset_sync_server.sh pos-oct
  bash reset_sync_server.sh pos-obo
  bash reset_sync_server.sh pos-gesr
  bash reset_sync_server.sh pos-bah
  bash reset_sync_server.sh pos-cen
  bash reset_sync_server.sh warehouse
  bash reset_sync_server.sh all

Optional environment override:
  SERVER_DB=/full/path/to/sync_server.sqlite3

Notes:
  - Stop the sync server before running this script.
  - "warehouse" targets device name "WAREHOUSE" by default.
  - "all" deletes every device, event, and cursor from the sync server DB.
EOF
}

require_sqlite() {
  if ! command -v sqlite3 >/dev/null 2>&1; then
    echo "Error: sqlite3 is not installed or not in PATH."
    exit 1
  fi
}

require_db() {
  if [[ ! -f "$SERVER_DB" ]]; then
    echo "Error: server DB not found:"
    echo "  $SERVER_DB"
    exit 1
  fi
}

confirm() {
  local prompt="$1"
  read -r -p "$prompt [y/N]: " reply
  [[ "${reply,,}" == "y" ]]
}

reset_device() {
  local device_name="$1"
  local target_scope="$2"

  echo "Resetting device '$device_name' in:"
  echo "  $SERVER_DB"

  sqlite3 "$SERVER_DB" <<EOF
PRAGMA foreign_keys = OFF;
BEGIN TRANSACTION;

DELETE FROM events
WHERE source_device IN (
  SELECT device_uuid FROM devices WHERE device_name = '$device_name'
)
OR target_scope = '$target_scope';

DELETE FROM device_cursors
WHERE device_uuid IN (
  SELECT device_uuid FROM devices WHERE device_name = '$device_name'
);

DELETE FROM devices
WHERE device_name = '$device_name';

COMMIT;
PRAGMA foreign_keys = ON;
EOF

  echo "Done."
}

reset_all() {
  echo "Resetting ALL devices/events/cursors in:"
  echo "  $SERVER_DB"

  sqlite3 "$SERVER_DB" <<'EOF'
PRAGMA foreign_keys = OFF;
BEGIN TRANSACTION;
DELETE FROM events;
DELETE FROM device_cursors;
DELETE FROM devices;
COMMIT;
PRAGMA foreign_keys = ON;
EOF

  echo "Done."
}

main() {
  if [[ $# -ne 1 ]]; then
    usage
    exit 1
  fi

  require_sqlite
  require_db

  case "$1" in
    pos-zay)
      if confirm "This will reset POS-ZAY server state and its targeted events."; then
        reset_device "POS-ZAY" "pos:POS-ZAY"
      else
        echo "Cancelled."
      fi
      ;;
    pos-oct)
      if confirm "This will reset POS-OCT server state and its targeted events."; then
        reset_device "POS-OCT" "pos:POS-OCT"
      else
        echo "Cancelled."
      fi
      ;;
    pos-obo)
      if confirm "This will reset POS-OBO server state and its targeted events."; then
        reset_device "POS-OBO" "pos:POS-OBO"
      else
        echo "Cancelled."
      fi
      ;;
    pos-gesr)
      if confirm "This will reset POS-GESR server state and its targeted events."; then
        reset_device "POS-GESR" "pos:POS-GESR"
      else
        echo "Cancelled."
      fi
      ;;
    pos-bah)
      if confirm "This will reset POS-BAH server state and its targeted events."; then
        reset_device "POS-BAH" "pos:POS-BAH"
      else
        echo "Cancelled."
      fi
      ;;
    pos-cen)
      if confirm "This will reset POS-CEN server state and its targeted events."; then
        reset_device "POS-CEN" "pos:POS-CEN"
      else
        echo "Cancelled."
      fi
      ;;
    warehouse)
      if confirm "This will reset WAREHOUSE server state and its device cursor/history."; then
        reset_device "WAREHOUSE" "warehouse"
      else
        echo "Cancelled."
      fi
      ;;
    all)
      if confirm "This will wipe ALL sync server devices, events, and cursors."; then
        reset_all
      else
        echo "Cancelled."
      fi
      ;;
    -h|--help|help)
      usage
      ;;
    *)
      echo "Unknown mode: $1"
      usage
      exit 1
      ;;
  esac
}

main "$@"
