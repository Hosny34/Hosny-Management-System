# -*- coding: utf-8 -*-
"""Sync client — talks to the sync server, drains the outbox, fills the inbox.

Phase 2 scope:
- Manual sync cycle: push all pending outbox events, then pull new
  events from the server into sync_inbox.
- Uses stdlib only (urllib, json, sqlite3) so PyInstaller just works.
- Never called from the Tk main loop directly — sync_ui.SyncDialog runs
  it on a background thread and posts progress messages via a queue.
- No event appliers yet. Pulled events land in sync_inbox but are NOT
  applied to domain tables. Phase 3 and 4 add appliers.

Design notes
------------
- The client has no persistent state of its own. Server URL, API key,
  JWT and cursor all live in the app's SQLite DB (device_identity and
  sync_state tables from sync_core).
- Credentials stay in the client DB. There is no OS keychain yet. This
  is acceptable for Phase 2; a future phase will add Windows DPAPI
  encryption.
- JWTs are fetched on demand. If the server rejects a token we transparently
  re-auth once before surfacing the error.
"""

from __future__ import annotations

import hashlib
import json
import os
import random
import sqlite3
import ssl
import sys
import time
import uuid
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple

try:
    import logging_setup
except Exception:
    logging_setup = None  # type: ignore


DEFAULT_TIMEOUT = 60.0  # seconds per HTTP request
PUSH_BATCH_SIZE = 50    # events per /sync/push call
PULL_BATCH_SIZE = 500   # events per /sync/pull call
HTTP_RETRY_ATTEMPTS = 3
try:
    from pos_version import APP_VERSION
except Exception:
    APP_VERSION = ""
DEAD_LETTER_MAX_ATTEMPTS = 5


def format_inbound_event_brief(event_type: str, payload: Any) -> str:
    """One-line Arabic-friendly summary for inbox events (UI / logs)."""
    p = payload if isinstance(payload, dict) else {}
    et = str(event_type or "").strip()

    def _s(x: Any, n: int = 80) -> str:
        t = str(x or "").strip().replace("\n", " ")
        return t if len(t) <= n else t[: n - 1] + "…"

    try:
        if et == "PRICE_UPDATE":
            flt = p.get("filters") if isinstance(p.get("filters"), dict) else {}
            parts = [
                _s(flt.get("item_type"), 40),
                _s(flt.get("school"), 40),
                _s(flt.get("color"), 20),
                _s(flt.get("size"), 20),
            ]
            parts = [x for x in parts if x]
            return f"سعر {p.get('new_price')} — " + " / ".join(parts) if parts else f"سعر {p.get('new_price')}"

        if et == "STOCK_TRANSFER_OUT":
            items = p.get("items") or []
            n = len(items) if isinstance(items, list) else 0
            note = _s(p.get("note"), 60)
            base = f"شحنة {n} بند — {_s(p.get('from_device'), 40)}"
            return base + (f" — {note}" if note else "")

        if et == "CATALOG_UPSERT":
            return f"كتالوج — {_s(p.get('item_type'), 60)}"

        if et == "POS_STOCK_SNAPSHOT":
            rows = p.get("rows") or []
            n = len(rows) if isinstance(rows, list) else 0
            src = _s(p.get("source_device_name") or p.get("source_device"), 40)
            return f"لقطة مخزون {n} صف — {src}" if src else f"لقطة مخزون {n} صف"

        if et == "STOCK_RETURN_TO_WAREHOUSE":
            return f"مرتجع إلى المخزن — {_s(p.get('note'), 80)}"

        if et == "POS_TRANSFER_VIA_WAREHOUSE":
            return f"تحويل عبر المخزن — {_s(p.get('note'), 80)}"

        if et == "SALE_CREATED":
            its = p.get("items") or []
            ni = len(its) if isinstance(its, list) else 0
            cust = _s(p.get("customer"), 40)
            tot = p.get("total")
            return f"فاتورة {ni} بند — {cust}" + (f" — إجمالي {tot}" if tot is not None else "")

        if et in ("SALE_VOIDED", "SALE_RETURNED"):
            bid = p.get("bill_id")
            return f"{et} — فاتورة #{bid}" if bid is not None else et

        return et
    except Exception:
        return et or "حدث"


# --------------------------- Config persistence --------------------------- #

def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")


def _table_row_count(conn: sqlite3.Connection, table: str) -> int:
    exists = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table,),
    ).fetchone()
    if not exists:
        return 0
    return int(conn.execute("SELECT COUNT(*) FROM %s" % table).fetchone()[0] or 0)


def _local_max_inbox_seq(conn: sqlite3.Connection) -> int:
    try:
        row = conn.execute("SELECT COALESCE(MAX(server_seq), 0) FROM sync_inbox").fetchone()
        return int((row[0] if row else 0) or 0)
    except sqlite3.Error:
        return 0


def _has_business_activity(conn: sqlite3.Connection) -> bool:
    for table in (
        "stocks",
        "bills",
        "bill_items",
        "income_bills",
        "income_bill_items",
        "movements",
        "reservations",
        "shifts",
        "returns",
        "return_items",
    ):
        try:
            if _table_row_count(conn, table) > 0:
                return True
        except sqlite3.Error:
            continue
    return False


def load_sync_config(conn: sqlite3.Connection) -> Dict[str, Any]:
    """Return the device identity + cursor state as a single dict."""
    row = conn.execute(
        "SELECT device_uuid, device_name, device_role, api_token, server_url "
        "FROM device_identity WHERE id = 1"
    ).fetchone()
    if row is None:
        return {
            "device_uuid": None, "device_name": None, "device_role": None,
            "api_token": None, "server_url": None,
            "last_pulled_seq": 0, "last_push_at": None, "last_pull_at": None,
            "last_error": None,
        }
    state = conn.execute(
        "SELECT last_pulled_seq, last_push_at, last_pull_at, last_error "
        "FROM sync_state WHERE channel = 'main'"
    ).fetchone()
    if state is None:
        last_seq, last_push, last_pull, last_err = 0, None, None, None
    else:
        last_seq = int(state[0] or 0)
        last_push = state[1]
        last_pull = state[2]
        last_err = state[3]
    local_max_seq = _local_max_inbox_seq(conn)
    if local_max_seq > last_seq:
        last_seq = local_max_seq
    return {
        "device_uuid": row[0],
        "device_name": row[1],
        "device_role": row[2],
        "api_token":   row[3],
        "server_url":  row[4],
        "last_pulled_seq": last_seq,
        "last_push_at":    last_push,
        "last_pull_at":    last_pull,
        "last_error":      last_err,
    }


def save_setup(
    conn: sqlite3.Connection,
    *,
    server_url: str,
    device_name: str,
    api_token: str = "",
) -> None:
    """Persist sync setup: server URL, device name, and optional API key.

    The simplified Railway deployment authenticates with device_name only,
    but we keep the api_token column for backward compatibility.
    """
    now = _utc_now_iso()
    new_name = device_name.strip()
    cfg = load_sync_config(conn)
    old_name = str(cfg.get("device_name") or "").strip()
    identity_change = (
        bool(old_name)
        and old_name.upper() not in {"POS-UNCONFIGURED", "UNCONFIGURED"}
        and old_name != new_name
    )
    if identity_change and _has_business_activity(conn):
        raise SyncError(
            "This database is already configured as %s and contains stock/sales/shift data. "
            "Do not rename it to %s. Close POS, delete warehouse_data.sqlite3, "
            "warehouse_data.sqlite3-wal, and warehouse_data.sqlite3-shm from the POS folder, "
            "then open POS again and configure the correct branch."
            % (old_name, new_name)
        )
    with conn:
        if identity_change:
            conn.execute(
                """
                UPDATE device_identity
                   SET device_uuid = ?, server_url = ?, device_name = ?,
                       api_token = ?, updated_at = ?
                 WHERE id = 1
                """,
                (str(uuid.uuid4()), server_url.rstrip("/"), new_name, (api_token or "").strip(), now),
            )
            conn.execute(
                """
                UPDATE sync_state
                   SET last_pulled_seq = 0,
                       last_push_at = NULL,
                       last_pull_at = NULL,
                       last_error = NULL
                 WHERE channel = 'main'
                """
            )
            conn.execute("DELETE FROM sync_inbox")
            conn.execute("DELETE FROM sync_dead_letter")
        else:
            conn.execute(
                """
                UPDATE device_identity
                   SET server_url = ?, device_name = ?, api_token = ?, updated_at = ?
                 WHERE id = 1
                """,
                (server_url.rstrip("/"), new_name, (api_token or "").strip(), now),
            )


def _set_sync_state(
    conn: sqlite3.Connection,
    *,
    last_pulled_seq: Optional[int] = None,
    last_push_at: Optional[str] = None,
    last_pull_at: Optional[str] = None,
    last_error: Optional[str] = None,
) -> None:
    """Patch sync_state for channel='main'. Only provided fields change."""
    current = conn.execute(
        "SELECT last_pulled_seq, last_push_at, last_pull_at, last_error "
        "FROM sync_state WHERE channel = 'main'"
    ).fetchone()
    cur_seq = int(current[0]) if current else 0
    cur_push = current[1] if current else None
    cur_pull = current[2] if current else None
    cur_err = current[3] if current else None

    new_seq = cur_seq if last_pulled_seq is None else int(last_pulled_seq)
    new_push = cur_push if last_push_at is None else last_push_at
    new_pull = cur_pull if last_pull_at is None else last_pull_at
    new_err = last_error  # explicit pass-through: caller can clear with None

    with conn:
        cur = conn.execute(
            """
            UPDATE sync_state
               SET last_pulled_seq = ?,
                   last_push_at    = ?,
                   last_pull_at    = ?,
                   last_error      = ?
             WHERE channel = 'main'
            """,
            (new_seq, new_push, new_pull, new_err),
        )
        if cur.rowcount == 0:
            conn.execute(
                """
                INSERT OR IGNORE INTO sync_state
                    (channel, last_pulled_seq, last_push_at, last_pull_at, last_error)
                VALUES ('main', ?, ?, ?, ?)
                """,
            (new_seq, new_push, new_pull, new_err),
        )


def _sync_meta_get(conn: sqlite3.Connection, key: str) -> Optional[str]:
    try:
        row = conn.execute("SELECT value FROM sync_meta WHERE key = ?", (key,)).fetchone()
        return None if row is None else row[0]
    except sqlite3.OperationalError:
        return None


def _sync_meta_set(conn: sqlite3.Connection, key: str, value: str) -> None:
    now = _utc_now_iso()
    try:
        with conn:
            conn.execute(
                """
                INSERT INTO sync_meta(key, value, updated_at)
                VALUES(?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET
                    value = excluded.value,
                    updated_at = excluded.updated_at
                """,
                (key, value, now),
            )
    except sqlite3.OperationalError:
        pass


def _snapshot_source_from_payload(payload: Dict[str, Any], fallback: Any = None) -> str:
    return str(
        payload.get("source_device_name")
        or payload.get("source_device")
        or payload.get("__source_device__")
        or fallback
        or ""
    ).strip()


def _checkpoint_wal_if_large(conn: sqlite3.Connection, *, max_wal_bytes: int = 32 * 1024 * 1024) -> None:
    try:
        row = conn.execute("PRAGMA database_list").fetchone()
        db_path = row[2] if row and len(row) > 2 else ""
        if not db_path:
            return
        wal_path = db_path + "-wal"
        if os.path.exists(wal_path) and os.path.getsize(wal_path) >= max_wal_bytes:
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        else:
            conn.execute("PRAGMA wal_checkpoint(PASSIVE)")
    except Exception:
        pass


# ------------------------------ HTTP helpers ------------------------------ #

class SyncError(Exception):
    """Any recoverable client-side sync failure. Carries a short reason."""


def _candidate_ca_files() -> List[str]:
    """CA bundle locations, ordered from deploy-local to Python defaults."""
    candidates: List[str] = []
    bases = []
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(getattr(sys, "executable", "")))
        if exe_dir:
            bases.append(exe_dir)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bases.append(str(meipass))
    bases.append(os.path.dirname(os.path.abspath(__file__)))

    for base in bases:
        candidates.append(os.path.join(base, "cacert.pem"))

    try:
        import certifi  # type: ignore

        candidates.append(certifi.where())
    except Exception:
        pass

    seen = set()
    existing = []
    for path in candidates:
        norm = os.path.abspath(path)
        if norm in seen:
            continue
        seen.add(norm)
        if os.path.isfile(norm):
            existing.append(norm)
    return existing


def _build_ssl_ctx(verify: bool) -> Optional[ssl.SSLContext]:
    if verify:
        for cafile in _candidate_ca_files():
            try:
                return ssl.create_default_context(cafile=cafile)
            except Exception:
                continue
        return ssl.create_default_context()
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _http_request(
    method: str,
    url: str,
    *,
    body: Optional[Dict[str, Any]] = None,
    token: Optional[str] = None,
    timeout: float = DEFAULT_TIMEOUT,
    verify_tls: bool = True,
) -> Tuple[int, Dict[str, Any]]:
    data: Optional[bytes] = None
    headers = {"Accept": "application/json"}
    if body is not None:
        data = json.dumps(body, ensure_ascii=False, default=str).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
    if token:
        headers["Authorization"] = "Bearer " + token

    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    ctx = _build_ssl_ctx(verify_tls) if url.lower().startswith("https://") else None
    for attempt in range(1, HTTP_RETRY_ATTEMPTS + 1):
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                raw = resp.read().decode("utf-8")
                return resp.status, (json.loads(raw) if raw else {})
        except urllib.error.HTTPError as e:
            can_retry = int(getattr(e, "code", 0) or 0) in (408, 429, 500, 502, 503, 504)
            if can_retry and attempt < HTTP_RETRY_ATTEMPTS:
                time.sleep((0.35 * (2 ** (attempt - 1))) + random.uniform(0.0, 0.25))
                continue
            try:
                detail = json.loads(e.read().decode("utf-8"))
            except Exception:
                detail = {"detail": str(e)}
            raise SyncError(f"HTTP {e.code}: {detail.get('detail', detail)}")
        except urllib.error.URLError as e:
            if attempt < HTTP_RETRY_ATTEMPTS:
                time.sleep((0.35 * (2 ** (attempt - 1))) + random.uniform(0.0, 0.25))
                continue
            raise SyncError(f"network error: {e.reason}")
        except TimeoutError:
            if attempt < HTTP_RETRY_ATTEMPTS:
                time.sleep((0.35 * (2 ** (attempt - 1))) + random.uniform(0.0, 0.25))
                continue
            raise SyncError("timeout")


# ------------------------------ SyncClient ------------------------------- #

ProgressFn = Callable[[str], None]


class SyncClient:
    """Stateless sync helper driven by the app DB."""

    def __init__(self, conn: sqlite3.Connection, verify_tls: bool = True) -> None:
        self.conn = conn
        self.verify_tls = verify_tls
        try:
            self.conn.execute("PRAGMA journal_mode=WAL;")
            self.conn.execute("PRAGMA busy_timeout=30000;")
        except Exception:
            pass

    # ---- auth ----

    def _token_url(self, cfg: Dict[str, Any]) -> str:
        return cfg["server_url"].rstrip("/") + "/v1/auth/token"

    def _fetch_jwt(self, cfg: Dict[str, Any]) -> str:
        body = {"device_name": cfg["device_name"]}
        if cfg.get("api_token"):
            body["api_key"] = cfg["api_token"]
        status, body = _http_request(
            "POST",
            self._token_url(cfg),
            body=body,
            verify_tls=self.verify_tls,
        )
        if status != 200 or "access_token" not in body:
            raise SyncError("login failed: unexpected response")
        return body["access_token"]

    # ---- public ----

    def test_connection(self) -> Dict[str, Any]:
        """Call /v1/health and /v1/auth/token. Used by the setup dialog."""
        cfg = load_sync_config(self.conn)
        cycle_id = "c" + _utc_now_iso().replace("-", "").replace(":", "").replace(".", "")
        if not cfg.get("server_url"):
            raise SyncError("server URL is not configured")
        health_url = cfg["server_url"].rstrip("/") + "/v1/health"
        status, body = _http_request(
            "GET", health_url, verify_tls=self.verify_tls
        )
        if status != 200:
            raise SyncError(f"health check failed: HTTP {status}")
        if not cfg.get("device_name"):
            return {"health": body, "auth": None}
        token = self._fetch_jwt(cfg)
        return {"health": body, "auth_ok": True, "token_prefix": token[:16] + "…"}

    def wait_for_updates(self, *, timeout_s: int = 25) -> Dict[str, Any]:
        """Long-poll server for incoming updates visible to this device."""
        cfg = load_sync_config(self.conn)
        if not cfg.get("server_url") or not cfg.get("device_name"):
            return {
                "has_updates": False,
                "next_seq": int(cfg.get("last_pulled_seq") or 0),
                "max_server_seq": 0,
            }

        token = self._fetch_jwt(cfg)
        since = int(cfg.get("last_pulled_seq") or 0)
        wait_url = (
            cfg["server_url"].rstrip("/")
            + "/v1/sync/wait?"
            + urllib.parse.urlencode(
                {"since": since, "timeout_s": int(max(1, min(timeout_s, 60)))}
            )
        )
        status, body = _http_request(
            "GET",
            wait_url,
            token=token,
            timeout=float(max(5, timeout_s + 5)),
            verify_tls=self.verify_tls,
        )
        if status != 200:
            raise SyncError(f"wait returned HTTP {status}")
        return {
            "has_updates": bool(body.get("has_updates")),
            "next_seq": int(body.get("next_seq") or since),
            "max_server_seq": int(body.get("max_server_seq") or 0),
        }

    def run_cycle(self, progress: Optional[ProgressFn] = None) -> Dict[str, Any]:
        """Run one full push + pull + apply cycle. Returns a summary dict.

        Safe to call from a worker thread. All UI updates must be posted
        via the supplied `progress` callback (which should marshal back
        to the Tk main loop).
        """
        def note(msg: str) -> None:
            if progress is not None:
                progress(msg)

        cfg = load_sync_config(self.conn)
        cycle_id = "c" + _utc_now_iso().replace("-", "").replace(":", "").replace(".", "")
        if not cfg.get("server_url") or not cfg.get("device_name"):
            raise SyncError("sync is not configured yet — open Setup first")
        started = time.time()
        try:
            if logging_setup is not None:  # type: ignore[name-defined]
                logging_setup.log_event(  # type: ignore[union-attr]
                    "sync.cycle.start",
                    cycle_id=cycle_id,
                    device_name=cfg.get("device_name"),
                    device_role=cfg.get("device_role"),
                    last_pulled_seq=cfg.get("last_pulled_seq"),
                    server_url=cfg.get("server_url"),
                )
        except Exception:
            pass

        note(f"[{cycle_id}] تسجيل الدخول...")
        token = self._fetch_jwt(cfg)

        # Phase 3: POS emits a full stock snapshot before pushing, so
        # the warehouse always has a fresh mirror after the round-trip.
        required_stock_snapshot_uuid: Optional[str] = None
        if (cfg.get("device_role") or "").lower() == "pos":
            try:
                note(f"[{cycle_id}] تحديث لقطة المخزون...")
                required_stock_snapshot_uuid = self.emit_stock_snapshot_event(cfg)
            except Exception as e:
                note(f"تعذّر إنشاء لقطة المخزون: {e}")
                _set_sync_state(self.conn, last_error=f"stock snapshot: {e}")
                raise SyncError(f"stock snapshot failed: {e}")
            try:
                note(f"[{cycle_id}] تحديث ملخص جرد نقطة البيع...")
                self.emit_stock_audit_snapshot_event(cfg)
            except Exception as e:
                note(f"تعذّر إنشاء ملخص جرد نقطة البيع: {e}")
            try:
                note(f"[{cycle_id}] تحديث ملخص مالي اليوم...")
                self.emit_financial_snapshot_event(cfg)
            except Exception as e:
                note(f"تعذّر إنشاء ملخص مالي اليوم: {e}")

        note(f"[{cycle_id}] جارٍ رفع الأحداث...")
        push_stats = self._push_loop(cfg, token, note)
        if required_stock_snapshot_uuid:
            row = self.conn.execute(
                "SELECT status, COALESCE(last_error, '') FROM sync_outbox WHERE event_uuid = ?",
                (required_stock_snapshot_uuid,),
            ).fetchone()
            status = str(row[0] if row else "").strip().lower()
            if status != "acked":
                err = str(row[1] if row else "").strip()
                detail = f" ({err})" if err else ""
                _set_sync_state(self.conn, last_error=f"stock snapshot not pushed{detail}")
                raise SyncError(f"stock snapshot was not pushed{detail}")

        note(f"[{cycle_id}] جارٍ تنزيل الأحداث...")
        pull_stats = self._pull_loop(cfg, token, note)

        # Phase 3: warehouse refreshes its known-device cache from the
        # server's status endpoint so the bill dialog can offer POS
        # names in the Customer dropdown. Non-fatal on failure.
        if (cfg.get("device_role") or "").lower() == "warehouse":
            try:
                note(f"[{cycle_id}] تحديث قائمة الفروع...")
                self.refresh_device_list(cfg, token)
            except Exception as e:
                note(f"تعذّر تحديث قائمة الفروع: {e}")

        note(f"[{cycle_id}] جارٍ تطبيق الأحداث الواردة...")
        apply_stats = self.apply_inbox(cfg, note)

        post_apply_push_stats = {"pushed": 0, "duplicates": 0}
        if (cfg.get("device_role") or "").lower() == "pos":
            changed_event_types = {
                str(e.get("event_type") or "")
                for e in (apply_stats.get("applied_events") or [])
                if isinstance(e, dict)
            }
            inventory_changing_events = {
                "PRICE_UPDATE",
                "STOCK_TRANSFER_OUT",
                "STOCK_TRANSFER_CANCELLED",
                "BRANCH_STOCK_RECLASSIFIED",
                "BRANCH_CATALOG_DELETED",
                "CATALOG_UPSERT",
                "SPEC_RENAMED",
                "POS_OWNERSHIP_SNAPSHOT",
            }
            if changed_event_types & inventory_changing_events:
                try:
                    note(f"[{cycle_id}] تحديث لقطة المخزون بعد تطبيق التغييرات...")
                    self.emit_stock_snapshot_event(cfg)
                    self.emit_stock_audit_snapshot_event(cfg)
                    post_apply_push_stats = self._push_loop(cfg, token, note)
                except Exception as e:
                    note(f"تعذّر رفع لقطة المخزون بعد التطبيق: {e}")
        _checkpoint_wal_if_large(self.conn)

        summary = {
            "pushed":     push_stats["pushed"] + post_apply_push_stats["pushed"],
            "duplicates": push_stats["duplicates"] + post_apply_push_stats["duplicates"],
            "pulled":     pull_stats["pulled"],
            "next_seq":   pull_stats["next_seq"],
            "applied":    apply_stats["applied"],
            "skipped":    apply_stats["skipped"],
            "apply_errors": apply_stats["errors"],
            "applied_events":  apply_stats.get("applied_events") or [],
            "skipped_events":  apply_stats.get("skipped_events") or [],
            "error_events":    apply_stats.get("error_events") or [],
            "dead_lettered": apply_stats.get("dead_lettered", 0),
            "cycle_id": cycle_id,
        }
        note(
            f"تم: رفع {summary['pushed']} • تنزيل {summary['pulled']} "
            f"• تطبيق {summary['applied']}"
            + (f" • فشل {summary['apply_errors']}" if summary["apply_errors"] else "")
            + (f" • DLQ {summary['dead_lettered']}" if summary["dead_lettered"] else "")
        )
        try:
            if logging_setup is not None:  # type: ignore[name-defined]
                logging_setup.log_event("sync.cycle.done", elapsed_ms=int((time.time() - started) * 1000), **summary)  # type: ignore[union-attr]
        except Exception:
            pass
        return summary

    # ---- apply (Phase 3) ----

    def apply_inbox(
        self,
        cfg: Optional[Dict[str, Any]] = None,
        progress: Optional[ProgressFn] = None,
    ) -> Dict[str, Any]:
        """Apply every not-yet-applied event in `sync_inbox` using the
        role-specific applier registry.

        Dispatch contract:
            - Only events with `apply_status IS NULL OR apply_status='error'`
              are picked up (retries on previously-failed events are safe
              because appliers are idempotent).
            - Successful applies mark the row with apply_status='ok' and
              apply_at = now.
            - Unknown event types get apply_status='deferred'. They stay
              recoverable; when a later version registers an applier, the
              same row is picked up again without a cleanup script.
            - An applier that raises ApplyError gets apply_status='error'
              + apply_error filled in; the transaction is rolled back so
              no partial domain mutation lands.
            - An applier that raises ANY other exception is treated the
              same way (defensive).
        """
        try:
            import sync_appliers
        except Exception as e:
            raise SyncError(f"sync_appliers unavailable: {e}")

        if cfg is None:
            cfg = load_sync_config(self.conn)
        registry = sync_appliers.for_role(cfg.get("device_role"))

        def note(msg: str) -> None:
            if progress is not None:
                progress(msg)

        # Fetch all pending + previously-failed events, ordered by
        # server_seq so dependent events apply in the right order.
        rows = self.conn.execute(
            """
            SELECT event_uuid, event_type, server_seq, payload_json,
                   COALESCE(apply_attempts, 0), source_device
              FROM sync_inbox
             WHERE apply_status IS NULL
                OR apply_status IN ('error', 'deferred')
             ORDER BY server_seq ASC
            """
        ).fetchall()

        if not rows:
            return {
                "applied": 0, "skipped": 0, "errors": 0,
                "applied_events": [], "skipped_events": [], "error_events": [],
            }

        applied = 0
        skipped = 0
        errors = 0
        dead_lettered = 0
        now = _utc_now_iso()
        applied_events: List[Dict[str, Any]] = []
        skipped_events: List[Dict[str, Any]] = []
        error_events: List[Dict[str, Any]] = []
        latest_snapshot_by_source: Dict[str, Tuple[int, str]] = {}
        superseded_snapshots = set()

        for r in rows:
            event_uuid, event_type, server_seq, payload_json, _attempts, source_dev = \
                r[0], r[1], int(r[2]), r[3], int(r[4] or 0), r[5]
            if event_type != "POS_STOCK_SNAPSHOT":
                continue
            try:
                payload = json.loads(payload_json or "{}")
            except Exception:
                payload = {}
            source = _snapshot_source_from_payload(payload, source_dev)
            if not source:
                continue
            previous = latest_snapshot_by_source.get(source)
            if previous is not None:
                superseded_snapshots.add(previous[1])
            latest_snapshot_by_source[source] = (server_seq, event_uuid)

        if superseded_snapshots:
            with self.conn:
                for r in rows:
                    event_uuid, event_type, server_seq, _payload_json, attempts, source_dev = \
                        r[0], r[1], int(r[2]), r[3], int(r[4] or 0), r[5]
                    if event_uuid not in superseded_snapshots:
                        continue
                    self.conn.execute(
                        """
                        UPDATE sync_inbox
                           SET apply_status = 'skipped',
                               apply_at = ?,
                               apply_error = 'superseded by newer stock snapshot',
                               apply_attempts = ?
                         WHERE event_uuid = ?
                        """,
                        (now, attempts + 1, event_uuid),
                    )
                    skipped += 1
                    skipped_events.append({
                        "event_uuid": str(event_uuid),
                        "event_type": str(event_type),
                        "server_seq": int(server_seq),
                        "source_device": str(source_dev or "") or None,
                        "reason": "superseded by newer stock snapshot",
                        "summary": str(event_type),
                    })

        for r in rows:
            event_uuid, event_type, server_seq, payload_json, attempts, source_dev = \
                r[0], r[1], int(r[2]), r[3], int(r[4] or 0), r[5]
            if event_uuid in superseded_snapshots:
                continue

            applier = registry.get(event_type)
            if applier is None:
                # Leave the event recoverable for future app versions.
                with self.conn:
                    self.conn.execute(
                        """
                        UPDATE sync_inbox
                           SET apply_status = 'deferred',
                               apply_at = ?,
                               apply_error = 'no applier for this device role',
                               apply_attempts = ?
                         WHERE event_uuid = ?
                        """,
                        (now, attempts + 1, event_uuid),
                    )
                skipped += 1
                skipped_events.append({
                    "event_uuid": str(event_uuid),
                    "event_type": str(event_type),
                    "server_seq": int(server_seq),
                    "source_device": str(source_dev or "") or None,
                    "reason": "no applier for this device role",
                    "summary": str(event_type),
                })
                continue

            try:
                payload = json.loads(payload_json or "{}")
            except Exception:
                payload = {}
            target_scope = str(payload.get("__target_scope__") or "").strip()
            device_name = str(cfg.get("device_name") or "").strip()
            if target_scope.lower().startswith("pos:"):
                target_device = target_scope[4:].strip()
                if device_name and target_device.casefold() != device_name.casefold():
                    reason = f"targeted to {target_device}, current device is {device_name}"
                    with self.conn:
                        self.conn.execute(
                            """
                            UPDATE sync_inbox
                               SET apply_status = 'skipped',
                                   apply_at = ?,
                                   apply_error = ?,
                                   apply_attempts = ?
                             WHERE event_uuid = ?
                            """,
                            (now, reason, attempts + 1, event_uuid),
                        )
                    skipped += 1
                    skipped_events.append({
                        "event_uuid": str(event_uuid),
                        "event_type": str(event_type),
                        "server_seq": int(server_seq),
                        "source_device": str(source_dev or "") or None,
                        "reason": reason,
                        "summary": format_inbound_event_brief(str(event_type), payload),
                    })
                    continue

            note(f"تطبيق {event_type} (seq={server_seq})...")
            try:
                # Each apply runs in its own transaction so a failure on
                # one event doesn't corrupt a later event's work.
                with self.conn:
                    result = applier(self.conn, payload, event_uuid) or {}
                    if result.get("skipped"):
                        status = "skipped"
                    else:
                        status = "ok"
                    self.conn.execute(
                        """
                        UPDATE sync_inbox
                           SET apply_status   = ?,
                               apply_at       = ?,
                               apply_error    = NULL,
                               apply_attempts = ?
                         WHERE event_uuid = ?
                        """,
                        (status, now, attempts + 1, event_uuid),
                    )
                if status == "ok":
                    applied += 1
                    applied_events.append({
                        "event_uuid": str(event_uuid),
                        "event_type": str(event_type),
                        "server_seq": int(server_seq),
                        "source_device": str(source_dev or "") or None,
                        "summary": format_inbound_event_brief(str(event_type), payload),
                    })
                else:
                    skipped += 1
                    skipped_events.append({
                        "event_uuid": str(event_uuid),
                        "event_type": str(event_type),
                        "server_seq": int(server_seq),
                        "source_device": str(source_dev or "") or None,
                        "reason": str((result or {}).get("reason") or "skipped by applier"),
                        "summary": format_inbound_event_brief(str(event_type), payload),
                    })
            except Exception as e:
                errors += 1
                err_text = str(e)[:500]
                next_attempts = attempts + 1
                # Separate write so the error lands even though the
                # applier's transaction rolled back.
                with self.conn:
                    if next_attempts >= DEAD_LETTER_MAX_ATTEMPTS:
                        cur = self.conn.execute(
                            """
                            UPDATE sync_dead_letter
                               SET apply_error = ?,
                                   attempts = ?,
                                   last_failed_at = ?
                             WHERE event_uuid = ?
                            """,
                            (err_text, next_attempts, now, event_uuid),
                        )
                        if cur.rowcount == 0:
                            self.conn.execute(
                                """
                                INSERT OR IGNORE INTO sync_dead_letter
                                    (event_uuid, event_type, server_seq, source_device, payload_json,
                                     apply_error, attempts, first_failed_at, last_failed_at)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """,
                                (
                                    event_uuid, event_type, int(server_seq), source_dev, payload_json,
                                    err_text, next_attempts, now, now,
                                ),
                            )
                        self.conn.execute(
                            """
                            UPDATE sync_dead_letter
                               SET apply_error = ?,
                                   attempts = ?,
                                   last_failed_at = ?
                             WHERE event_uuid = ?
                            """,
                            (err_text, next_attempts, now, event_uuid),
                        )
                        self.conn.execute(
                            """
                            UPDATE sync_inbox
                               SET apply_status   = 'dead',
                                   apply_error    = ?,
                                   apply_attempts = ?
                             WHERE event_uuid = ?
                            """,
                            (err_text, next_attempts, event_uuid),
                        )
                        dead_lettered += 1
                    else:
                        self.conn.execute(
                            """
                            UPDATE sync_inbox
                               SET apply_status   = 'error',
                                   apply_error    = ?,
                                   apply_attempts = ?
                             WHERE event_uuid = ?
                            """,
                            (err_text, next_attempts, event_uuid),
                        )
                note(f"فشل تطبيق {event_type}: {err_text}")
                try:
                    payload_err = json.loads(payload_json or "{}")
                except Exception:
                    payload_err = {}
                error_events.append({
                    "event_uuid": str(event_uuid),
                    "event_type": str(event_type),
                    "server_seq": int(server_seq),
                    "source_device": str(source_dev or "") or None,
                    "error": err_text,
                    "summary": format_inbound_event_brief(str(event_type), payload_err),
                })

        return {
            "applied": applied,
            "skipped": skipped,
            "errors": errors,
            "dead_lettered": dead_lettered,
            "applied_events": applied_events,
            "skipped_events": skipped_events,
            "error_events": error_events,
        }

    # ---- push ----

    def _push_loop(
        self,
        cfg: Dict[str, Any],
        token: str,
        note: ProgressFn,
    ) -> Dict[str, int]:
        pushed_total = 0
        dup_total = 0
        push_url = cfg["server_url"].rstrip("/") + "/v1/sync/push"

        has_target = self._has_target_scope_column()

        while True:
            if has_target:
                rows = self.conn.execute(
                    """
                    SELECT local_seq, event_uuid, event_type, payload_json,
                           created_at, target_scope
                      FROM sync_outbox
                     WHERE status = 'pending'
                     ORDER BY local_seq ASC
                     LIMIT ?
                    """,
                    (PUSH_BATCH_SIZE,),
                ).fetchall()
            else:
                rows = self.conn.execute(
                    """
                    SELECT local_seq, event_uuid, event_type, payload_json,
                           created_at
                      FROM sync_outbox
                     WHERE status = 'pending'
                     ORDER BY local_seq ASC
                     LIMIT ?
                    """,
                    (PUSH_BATCH_SIZE,),
                ).fetchall()
            if not rows:
                break

            batch = []
            for r in rows:
                try:
                    payload = json.loads(r[3] or "{}")
                except Exception:
                    payload = {"_raw": r[3]}
                target_scope = None
                if has_target and len(r) > 5:
                    target_scope = r[5]
                # Legacy-path fallback: if the record was written before
                # the target_scope column existed, it may carry the
                # scope inside the payload.
                if target_scope is None and isinstance(payload, dict):
                    target_scope = payload.pop("__target_scope__", None)
                entry = {
                    "event_uuid": r[1],
                    "event_type": r[2],
                    "payload":    payload,
                    "created_at": r[4],
                }
                if target_scope:
                    entry["target_scope"] = target_scope
                batch.append(entry)

            note(f"رفع دفعة من {len(batch)} حدث...")
            try:
                status, body = _http_request(
                    "POST",
                    push_url,
                    body={"events": batch},
                    token=token,
                    verify_tls=self.verify_tls,
                )
            except SyncError as e:
                # Mark the batch as failed-this-round and bail out.
                self._mark_failed([r[0] for r in rows], str(e))
                _set_sync_state(self.conn, last_error=f"push: {e}")
                raise
            if status != 200:
                raise SyncError(f"push returned HTTP {status}")

            pushed_total += int(body.get("accepted", 0))
            dup_total += int(body.get("duplicates", 0))
            self._mark_sent([r[0] for r in rows])

            # If the whole batch was the whole remaining outbox, next
            # iteration will find 0 rows and exit.

        _set_sync_state(
            self.conn,
            last_push_at=_utc_now_iso(),
            last_error=None,
        )
        return {"pushed": pushed_total, "duplicates": dup_total}

    def _mark_sent(self, local_seqs: List[int]) -> None:
        if not local_seqs:
            return
        now = _utc_now_iso()
        placeholders = ",".join("?" * len(local_seqs))
        with self.conn:
            self.conn.execute(
                "UPDATE sync_outbox SET status='acked', sent_at=?, last_error=NULL "
                "WHERE local_seq IN (" + placeholders + ")",
                (now, *local_seqs),
            )

    def _mark_failed(self, local_seqs: List[int], err: str) -> None:
        if not local_seqs:
            return
        placeholders = ",".join("?" * len(local_seqs))
        with self.conn:
            self.conn.execute(
                "UPDATE sync_outbox SET attempts = attempts + 1, last_error = ? "
                "WHERE local_seq IN (" + placeholders + ")",
                (err[:500], *local_seqs),
            )

    # ---- pull ----

    def _pull_loop(
        self,
        cfg: Dict[str, Any],
        token: str,
        note: ProgressFn,
    ) -> Dict[str, int]:
        pulled_total = 0
        cursor = int(cfg.get("last_pulled_seq") or 0)
        local_max_seq = _local_max_inbox_seq(self.conn)
        if local_max_seq > cursor:
            cursor = local_max_seq
            _set_sync_state(
                self.conn,
                last_pulled_seq=cursor,
                last_error=None,
            )
            note(f"استعادة مؤشر المزامنة إلى seq={cursor} من الأحداث المحفوظة محلياً...")
        pull_url_base = cfg["server_url"].rstrip("/") + "/v1/sync/pull"
        recovery_attempted = False

        while True:
            qs = urllib.parse.urlencode({"since": cursor, "limit": PULL_BATCH_SIZE})
            try:
                status, body = _http_request(
                    "GET",
                    pull_url_base + "?" + qs,
                    token=token,
                    verify_tls=self.verify_tls,
                )
            except SyncError as e:
                _set_sync_state(self.conn, last_error=f"pull: {e}")
                raise
            if status != 200:
                raise SyncError(f"pull returned HTTP {status}")

            events = body.get("events", []) or []
            # Detect a server-side reset: when the server reports its own
            # highest server_seq is lower than our cursor, its event log
            # was wiped (common with Railway ephemeral storage). In that
            # case we rewind to 0 so the next iteration re-pulls from the
            # start. Inbox inserts are idempotent via event_uuid.
            srv_max = body.get("max_server_seq")
            if (
                not recovery_attempted
                and srv_max is not None
                and cursor > 0
                and int(srv_max) < cursor
            ):
                recovery_attempted = True
                note("اكتشاف إعادة ضبط للسيرفر — إعادة ضبط المؤشر وإعادة التنزيل…")
                cursor = 0
                continue
            if not events:
                break

            note(f"تنزيل دفعة من {len(events)} حدث...")
            now = _utc_now_iso()
            with self.conn:
                for ev in events:
                    try:
                        payload_obj = dict(ev.get("payload") or {})
                        sd = ev.get("source_device")
                        if sd:
                            payload_obj["__source_device__"] = sd
                        target_scope = ev.get("target_scope")
                        if target_scope:
                            payload_obj["__target_scope__"] = target_scope
                        self.conn.execute(
                            """
                            INSERT OR IGNORE INTO sync_inbox
                                (event_uuid, event_type, server_seq,
                                 source_device, payload_json, applied_at, server_created_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                ev["event_uuid"],
                                ev["event_type"],
                                int(ev["server_seq"]),
                                ev.get("source_device"),
                                json.dumps(payload_obj, ensure_ascii=False, default=str),
                                now,
                                ev.get("created_at"),
                            ),
                        )
                    except sqlite3.IntegrityError:
                        # Already applied in a previous cycle — idempotent.
                        pass
            pulled_total += len(events)
            cursor = int(body.get("next_seq", cursor))

            if len(events) < PULL_BATCH_SIZE:
                break

        _set_sync_state(
            self.conn,
            last_pulled_seq=cursor,
            last_pull_at=_utc_now_iso(),
            last_error=None,
        )
        return {"pulled": pulled_total, "next_seq": cursor}

    # ---- helpers (Phase 3) ----

    def _has_target_scope_column(self) -> bool:
        """True if sync_outbox has the Phase-3 target_scope column."""
        try:
            cur = self.conn.execute("PRAGMA table_info(sync_outbox)")
            return any(r[1] == "target_scope" for r in cur.fetchall())
        except sqlite3.OperationalError:
            return False

    def emit_stock_snapshot_event(
        self, cfg: Optional[Dict[str, Any]] = None
    ) -> Optional[str]:
        """POS-side: snapshot the full stocks table and append a
        STOCK_SNAPSHOT outbox event targeted at 'warehouse'.

        Called at the start of every POS sync cycle so the warehouse
        has a fresh mirror. If the current snapshot is unchanged and an
        identical snapshot is already pending, reuse it instead of
        appending another large upload payload.

        Returns the event_uuid when a new or still-pending snapshot must
        be acknowledged. If the current unchanged snapshot was already
        acknowledged, returns None.
        """
        if cfg is None:
            cfg = load_sync_config(self.conn)
        if (cfg.get("device_role") or "").lower() != "pos":
            return None
        device_name = cfg.get("device_name") or "POS-UNCONFIGURED"

        try:
            rows = self.conn.execute(
                """
                SELECT item_type, school, color, size, unit_price, SUM(count) AS total
                  FROM stocks
                 GROUP BY item_type, school, color, size, unit_price
                 ORDER BY item_type, school, color, size, unit_price
                """
            ).fetchall()
        except sqlite3.OperationalError as e:
            raise SyncError(f"cannot read stocks for stock snapshot: {e}")

        snapshot_at = _utc_now_iso()
        snapshot_rows = [
            {
                "item_type":  r[0],
                "school":     r[1],
                "color":      r[2],
                "size":       r[3],
                "unit_price": float(r[4] or 0),
                "count":      int(r[5] or 0),
            }
            for r in rows
        ]
        snapshot_hash = hashlib.sha256(
            json.dumps(snapshot_rows, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
        ).hexdigest()
        meta_key = f"pos_stock_snapshot_hash:{device_name}"
        previous_hash = _sync_meta_get(self.conn, meta_key)
        if previous_hash == snapshot_hash:
            try:
                existing = self.conn.execute(
                    """
                    SELECT event_uuid
                      FROM sync_outbox
                     WHERE event_type = 'POS_STOCK_SNAPSHOT'
                       AND status = 'pending'
                       AND payload_json LIKE ?
                     ORDER BY local_seq DESC
                     LIMIT 1
                    """,
                    (f"%{snapshot_hash}%",),
                ).fetchone()
            except sqlite3.OperationalError:
                existing = None
            if existing and existing[0]:
                try:
                    if logging_setup is not None:  # type: ignore[name-defined]
                        logging_setup.log_event(  # type: ignore[union-attr]
                            "sync.stock_snapshot.reuse_pending",
                            device_name=device_name,
                            event_uuid=str(existing[0]),
                            snapshot_hash=snapshot_hash,
                        )
                except Exception:
                    pass
                return str(existing[0])
            return None

        payload = {
            "source_device_name": device_name,
            "snapshot_at":        snapshot_at,
            "app_version":        APP_VERSION,
            "includes_zero_rows":  True,
            "snapshot_hash":      snapshot_hash,
            "rows":               snapshot_rows,
        }

        # Append directly to sync_outbox with scope = 'warehouse' so
        # the server routes it only to the single WH device.
        import uuid as _uuid
        event_uuid = str(_uuid.uuid4())
        now = _utc_now_iso()

        has_target = self._has_target_scope_column()
        try:
            with self.conn:
                if has_target:
                    self.conn.execute(
                    """
                    INSERT INTO sync_outbox
                        (event_uuid, event_type, payload_json,
                         created_at, status, attempts, target_scope)
                    VALUES (?, 'POS_STOCK_SNAPSHOT', ?, ?, 'pending', 0, 'warehouse')
                    """,
                    (
                        event_uuid,
                        json.dumps(payload, ensure_ascii=False, default=str),
                        now,
                    ),
                    )
                else:
                    # Legacy schema — stash the scope inside the payload.
                    payload["__target_scope__"] = "warehouse"
                    self.conn.execute(
                    """
                    INSERT INTO sync_outbox
                        (event_uuid, event_type, payload_json,
                         created_at, status, attempts)
                    VALUES (?, 'POS_STOCK_SNAPSHOT', ?, ?, 'pending', 0)
                    """,
                    (
                        event_uuid,
                        json.dumps(payload, ensure_ascii=False, default=str),
                        now,
                    ),
                    )
        except sqlite3.OperationalError as e:
            raise SyncError(f"cannot queue stock snapshot: {e}")

        _sync_meta_set(self.conn, meta_key, snapshot_hash)
        try:
            if logging_setup is not None:  # type: ignore[name-defined]
                logging_setup.log_event(  # type: ignore[union-attr]
                    "sync.stock_snapshot.queued",
                    device_name=device_name,
                    event_uuid=event_uuid,
                    rows=len(snapshot_rows),
                    total_qty=sum(int(r.get("count") or 0) for r in snapshot_rows),
                    total_value=sum(float(r.get("unit_price") or 0) * int(r.get("count") or 0) for r in snapshot_rows),
                    snapshot_hash=snapshot_hash,
                    snapshot_at=snapshot_at,
                    app_version=APP_VERSION,
                )
        except Exception:
            pass
        return event_uuid

    def emit_stock_audit_snapshot_event(
        self, cfg: Optional[Dict[str, Any]] = None
    ) -> Optional[str]:
        """POS-side: mirror stock-audit reports to warehouse for reporting only."""
        if cfg is None:
            cfg = load_sync_config(self.conn)
        if (cfg.get("device_role") or "").lower() != "pos":
            return None
        device_name = cfg.get("device_name") or "POS-UNCONFIGURED"

        try:
            reports = self.conn.execute(
                """
                SELECT id, created_at, reason, diff_count, total_diff, total_value
                  FROM stock_audit_reports
                 ORDER BY id ASC
                """
            ).fetchall()
        except sqlite3.OperationalError:
            return None

        report_rows = []
        for report in reports:
            report_id = int(report[0])
            lines = self.conn.execute(
                """
                SELECT item_type, school, color, size, expected, actual,
                       diff, unit_price, diff_value
                  FROM stock_audit_report_lines
                 WHERE report_id = ?
                 ORDER BY id ASC
                """,
                (report_id,),
            ).fetchall()
            report_rows.append({
                "audit_uuid": f"{device_name}:{report_id}:{str(report[1] or '').strip()}",
                "report_id": report_id,
                "created_at": report[1],
                "reason": report[2],
                "diff_count": int(report[3] or 0),
                "total_diff": int(report[4] or 0),
                "total_value": float(report[5] or 0),
                "lines": [
                    {
                        "item_type": row[0],
                        "school": row[1],
                        "color": row[2],
                        "size": row[3],
                        "expected": int(row[4] or 0),
                        "actual": int(row[5] or 0),
                        "diff": int(row[6] or 0),
                        "unit_price": float(row[7] or 0),
                        "diff_value": float(row[8] or 0),
                    }
                    for row in lines
                ],
            })

        snapshot_hash = hashlib.sha256(
            json.dumps(report_rows, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
        ).hexdigest()
        meta_key = f"pos_stock_audit_snapshot_hash:{device_name}"
        if _sync_meta_get(self.conn, meta_key) == snapshot_hash:
            return None

        payload = {
            "source_device_name": device_name,
            "snapshot_at": _utc_now_iso(),
            "app_version": APP_VERSION,
            "snapshot_hash": snapshot_hash,
            "reports": report_rows,
        }

        import uuid as _uuid
        event_uuid = str(_uuid.uuid4())
        now = _utc_now_iso()
        has_target = self._has_target_scope_column()
        try:
            if has_target:
                self.conn.execute(
                    """
                    INSERT INTO sync_outbox
                        (event_uuid, event_type, payload_json,
                         created_at, status, attempts, target_scope)
                    VALUES (?, 'POS_STOCK_AUDIT_SNAPSHOT', ?, ?, 'pending', 0, 'warehouse')
                    """,
                    (event_uuid, json.dumps(payload, ensure_ascii=False, default=str), now),
                )
            else:
                payload["__target_scope__"] = "warehouse"
                self.conn.execute(
                    """
                    INSERT INTO sync_outbox
                        (event_uuid, event_type, payload_json,
                         created_at, status, attempts)
                    VALUES (?, 'POS_STOCK_AUDIT_SNAPSHOT', ?, ?, 'pending', 0)
                    """,
                    (event_uuid, json.dumps(payload, ensure_ascii=False, default=str), now),
                )
        except sqlite3.OperationalError:
            return None
        _sync_meta_set(self.conn, meta_key, snapshot_hash)
        return event_uuid

    def _insert_warehouse_targeted_event(
        self,
        event_type: str,
        payload: Dict[str, Any],
    ) -> Optional[str]:
        import uuid as _uuid

        event_uuid = str(_uuid.uuid4())
        now = _utc_now_iso()
        try:
            if self._has_target_scope_column():
                self.conn.execute(
                    """
                    INSERT INTO sync_outbox
                        (event_uuid, event_type, payload_json,
                         created_at, status, attempts, target_scope)
                    VALUES (?, ?, ?, ?, 'pending', 0, 'warehouse')
                    """,
                    (event_uuid, event_type, json.dumps(payload, ensure_ascii=False, default=str), now),
                )
            else:
                scoped_payload = dict(payload)
                scoped_payload["__target_scope__"] = "warehouse"
                self.conn.execute(
                    """
                    INSERT INTO sync_outbox
                        (event_uuid, event_type, payload_json,
                         created_at, status, attempts)
                    VALUES (?, ?, ?, ?, 'pending', 0)
                    """,
                    (event_uuid, event_type, json.dumps(scoped_payload, ensure_ascii=False, default=str), now),
                )
        except sqlite3.OperationalError:
            return None
        return event_uuid

    def emit_financial_snapshot_event(
        self, cfg: Optional[Dict[str, Any]] = None
    ) -> Optional[str]:
        """POS-side daily cash/Visa snapshot for the Warehouse monitor."""
        if cfg is None:
            cfg = load_sync_config(self.conn)
        if (cfg.get("device_role") or "").lower() != "pos":
            return None
        device_name = cfg.get("device_name") or "POS-UNCONFIGURED"
        day = datetime.now().strftime("%Y-%m-%d")

        cash_total = 0.0
        visa_total = 0.0

        def add_amount(amount: Any, payment_method: Any) -> None:
            nonlocal cash_total, visa_total
            try:
                value = float(amount or 0.0)
            except (TypeError, ValueError):
                value = 0.0
            method = str(payment_method or "CASH").strip().upper()
            if method == "VISA":
                visa_total += value
            else:
                cash_total += value

        try:
            bill_rows = self.conn.execute(
                """
                SELECT COALESCE(bill_type, 'SALE') AS bill_type,
                       COALESCE(status, 'CONFIRMED') AS status,
                       COALESCE(payment_method, 'CASH') AS payment_method,
                       COALESCE(total, 0) AS total
                  FROM bills
                 WHERE substr(COALESCE(created_at, ''), 1, 10) = ?
                """,
                (day,),
            ).fetchall()
        except sqlite3.OperationalError:
            bill_rows = []
        for row in bill_rows:
            bill_type = str(row[0] or "SALE").upper()
            status = str(row[1] or "CONFIRMED").upper()
            if status == "VOID":
                continue
            amount = float(row[3] or 0.0)
            if bill_type == "RETURN":
                amount = -abs(amount)
            elif bill_type == "EXCHANGE":
                amount = amount
            else:
                amount = abs(amount)
            add_amount(amount, row[2])

        try:
            movement_rows = self.conn.execute(
                """
                SELECT m.direction,
                       COALESCE(m.unit_price, 0) AS amount,
                       COALESCE(m.payment_method, b.payment_method, 'CASH') AS payment_method,
                       COALESCE(b.created_at, '') AS bill_created_at
                  FROM movements m
             LEFT JOIN bills b ON b.id = m.bill_id
                 WHERE substr(COALESCE(m.ts, ''), 1, 10) = ?
                   AND m.direction IN ('RESERVE_PAY', 'DELIVER_PAY', 'RESERVE_REFUND', 'VOID_PAY')
                """,
                (day,),
            ).fetchall()
        except sqlite3.OperationalError:
            movement_rows = []
        for row in movement_rows:
            direction = str(row[0] or "").upper()
            amount = float(row[1] or 0.0)
            if direction == "VOID_PAY" and str(row[3] or "")[:10] == day:
                continue
            if direction == "VOID_PAY":
                amount = -amount
            elif direction == "RESERVE_REFUND":
                amount = -abs(amount)
            else:
                amount = abs(amount)
            add_amount(amount, row[2])

        payload = {
            "source_device_name": device_name,
            "snapshot_at": _utc_now_iso(),
            "app_version": APP_VERSION,
            "day": day,
            "cash_total": round(float(cash_total), 2),
            "visa_total": round(float(visa_total), 2),
            "total_collected": round(float(cash_total + visa_total), 2),
        }
        return self._insert_warehouse_targeted_event("POS_FINANCIAL_SNAPSHOT", payload)

    def refresh_device_list(
        self,
        cfg: Optional[Dict[str, Any]] = None,
        token: Optional[str] = None,
    ) -> int:
        """Warehouse-side: fetch /v1/sync/status and refresh the
        `known_devices` cache so the bill dialog can offer POS device
        names in the Customer field even offline.

        Returns the number of devices written. Raises SyncError on
        transport/auth issues; callers should treat non-warehouse
        roles as no-ops and not call this.
        """
        if cfg is None:
            cfg = load_sync_config(self.conn)
        if (cfg.get("device_role") or "").lower() != "warehouse":
            return 0
        if not cfg.get("server_url"):
            raise SyncError("no server_url configured")
        if token is None:
            token = self._fetch_jwt(cfg)

        status_url = cfg["server_url"].rstrip("/") + "/v1/sync/status"
        _, body = _http_request(
            "GET", status_url, token=token, verify_tls=self.verify_tls
        )
        devices = body.get("devices", []) or []

        now = _utc_now_iso()
        written = 0
        with self.conn:
            for d in devices:
                name = (d.get("device_name") or "").strip()
                role = (d.get("role") or "").strip()
                if not name or not role:
                    continue
                cur = self.conn.execute(
                    """
                    UPDATE known_devices
                       SET device_uuid  = ?,
                           role         = ?,
                           last_seen_at = ?,
                           updated_at   = ?
                     WHERE device_name = ?
                    """,
                    (
                        d.get("device_uuid"),
                        role,
                        d.get("last_seen_at"),
                        now,
                        name,
                    ),
                )
                if cur.rowcount == 0:
                    self.conn.execute(
                        """
                        INSERT OR IGNORE INTO known_devices
                            (device_name, device_uuid, role, last_seen_at, updated_at)
                        VALUES (?, ?, ?, ?, ?)
                        """,
                        (
                            name,
                            d.get("device_uuid"),
                            role,
                            d.get("last_seen_at"),
                            now,
                        ),
                    )
                written += 1
        return written
