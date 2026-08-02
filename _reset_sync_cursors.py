"""Reset local sync cursors for all client DBs.

Call this when the Railway sync server was redeployed without a persistent
volume. The server's event log starts over from server_seq=1, but clients
still carry large `last_pulled_seq` values from the previous deployment,
so pulls return 0 events forever.

Does NOT touch sync_inbox (idempotent by event_uuid) or sync_outbox
(acked rows stay acked; local domain state is untouched).
"""
from __future__ import annotations
import os, sqlite3
from pathlib import Path

REPO = Path(__file__).resolve().parent
DBS = [
    REPO / "Warehouse" / "warehouse_data.sqlite3",
    REPO / "POS" / "warehouse_data.sqlite3",
    REPO / "POS-STOCK-MONITOR" / "warehouse_data.sqlite3",
]

for p in DBS:
    p = str(p)
    if not os.path.exists(p):
        print("skip (missing):", p); continue
    c = sqlite3.connect(p)
    try:
        try:
            before = c.execute(
                "SELECT last_pulled_seq FROM sync_state WHERE channel='main'"
            ).fetchone()
        except sqlite3.OperationalError:
            before = None
        with c:
            try:
                c.execute(
                    "UPDATE sync_state SET last_pulled_seq=0, last_pull_at=NULL, "
                    "last_error='cursor reset after server redeploy' "
                    "WHERE channel='main'"
                )
            except sqlite3.OperationalError as e:
                print("skip (no sync_state):", p, "err:", e); continue
        after = c.execute(
            "SELECT last_pulled_seq FROM sync_state WHERE channel='main'"
        ).fetchone()
        print(f"reset {os.path.basename(os.path.dirname(p))}: "
              f"{before[0] if before else None} -> {after[0]}")
    finally:
        c.close()

print("\nDone. Run manual sync on each app now to re-pull all events.")
