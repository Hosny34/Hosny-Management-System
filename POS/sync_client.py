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

import json
import random
import sqlite3
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple


DEFAULT_TIMEOUT = 15.0  # seconds per HTTP request
PUSH_BATCH_SIZE = 200   # events per /sync/push call
PULL_BATCH_SIZE = 500   # events per /sync/pull call
HTTP_RETRY_ATTEMPTS = 3
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
    with conn:
        conn.execute(
            """
            UPDATE device_identity
               SET server_url = ?, device_name = ?, api_token = ?, updated_at = ?
             WHERE id = 1
            """,
            (server_url.rstrip("/"), device_name.strip(), (api_token or "").strip(), now),
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
        conn.execute(
            """
            INSERT INTO sync_state (channel, last_pulled_seq, last_push_at, last_pull_at, last_error)
                 VALUES ('main', ?, ?, ?, ?)
            ON CONFLICT(channel) DO UPDATE SET
                last_pulled_seq = excluded.last_pulled_seq,
                last_push_at    = excluded.last_push_at,
                last_pull_at    = excluded.last_pull_at,
                last_error      = excluded.last_error
            """,
            (new_seq, new_push, new_pull, new_err),
        )


# ------------------------------ HTTP helpers ------------------------------ #

class SyncError(Exception):
    """Any recoverable client-side sync failure. Carries a short reason."""


def _build_ssl_ctx(verify: bool) -> Optional[ssl.SSLContext]:
    if verify:
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

        note(f"[{cycle_id}] تسجيل الدخول...")
        token = self._fetch_jwt(cfg)

        # Phase 3: POS emits a full stock snapshot before pushing, so
        # the warehouse always has a fresh mirror after the round-trip.
        if (cfg.get("device_role") or "").lower() == "pos":
            try:
                note(f"[{cycle_id}] تحديث لقطة المخزون...")
                self.emit_stock_snapshot_event(cfg)
            except Exception as e:
                note(f"تعذّر إنشاء لقطة المخزون: {e}")

        note(f"[{cycle_id}] جارٍ رفع الأحداث...")
        push_stats = self._push_loop(cfg, token, note)

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

        summary = {
            "pushed":     push_stats["pushed"],
            "duplicates": push_stats["duplicates"],
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
            - Unknown event types get apply_status='skipped'. They stay in
              the inbox; Phase 4 will register appliers for them and
              re-drain in the next cycle.
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
                OR apply_status = 'error'
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

        for r in rows:
            event_uuid, event_type, server_seq, payload_json, attempts, source_dev = \
                r[0], r[1], int(r[2]), r[3], int(r[4] or 0), r[5]

            applier = registry.get(event_type)
            if applier is None:
                # Leave the event in the inbox but record that we looked
                # at it. When a later phase adds a handler, bump the
                # migration to reset `apply_status` to NULL for these
                # rows so they get picked up again.
                with self.conn:
                    self.conn.execute(
                        "UPDATE sync_inbox SET apply_status = 'skipped' "
                        "WHERE event_uuid = ?",
                        (event_uuid,),
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
                        self.conn.execute(
                            """
                            INSERT INTO sync_dead_letter
                                (event_uuid, event_type, server_seq, source_device, payload_json,
                                 apply_error, attempts, first_failed_at, last_failed_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ON CONFLICT(event_uuid) DO UPDATE SET
                                apply_error = excluded.apply_error,
                                attempts = excluded.attempts,
                                last_failed_at = excluded.last_failed_at
                            """,
                            (
                                event_uuid, event_type, int(server_seq), source_dev, payload_json,
                                err_text, next_attempts, now, now,
                            ),
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
                        self.conn.execute(
                            """
                            INSERT INTO sync_inbox
                                (event_uuid, event_type, server_seq,
                                 source_device, payload_json, applied_at)
                            VALUES (?, ?, ?, ?, ?, ?)
                            """,
                            (
                                ev["event_uuid"],
                                ev["event_type"],
                                int(ev["server_seq"]),
                                ev.get("source_device"),
                                json.dumps(payload_obj, ensure_ascii=False, default=str),
                                now,
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
        always has a fresh mirror. Deduplication across cycles is a
        Phase-5 optimisation; for now a snapshot is emitted every time
        even if stocks haven't changed.

        Returns the event_uuid, or None on soft failure (e.g. no
        stocks table). Never raises on domain issues — the cycle
        should continue even if the snapshot can't be emitted.
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
                 WHERE count > 0
                 GROUP BY item_type, school, color, size, unit_price
                """
            ).fetchall()
        except sqlite3.OperationalError:
            return None

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

        payload = {
            "source_device_name": device_name,
            "snapshot_at":        snapshot_at,
            "rows":               snapshot_rows,
        }

        # Append directly to sync_outbox with scope = 'warehouse' so
        # the server routes it only to the single WH device.
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
        except sqlite3.OperationalError:
            return None

        return event_uuid

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
                self.conn.execute(
                    """
                    INSERT INTO known_devices
                        (device_name, device_uuid, role, last_seen_at, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(device_name) DO UPDATE SET
                        device_uuid  = excluded.device_uuid,
                        role         = excluded.role,
                        last_seen_at = excluded.last_seen_at,
                        updated_at   = excluded.updated_at
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
