# -*- coding: utf-8 -*-
"""Background periodic sync for warehouse/POS Tk apps.

Runs ``SyncClient.run_cycle`` on a worker thread every 10 minutes.
Automatic sync is intentionally quiet: if the internet/server is down,
cashiers should not see errors; the next attempt is simply scheduled.
Manual sync still reports errors through ``sync_ui``.
"""

from __future__ import annotations

import os
import queue
import sqlite3
import threading
import time
import tkinter as tk
import zlib
from typing import Any, Dict, Optional

import sync_client

# 10-minute base interval for all clients.
DEFAULT_BASE_MS = 10 * 60 * 1000
FIRST_DELAY_BASE_MS = 8_000
OFFSET_STEP_MS = 12_000
OFFSET_SLOT_COUNT = 20


class PeriodicSyncController:
    def __init__(self, root: tk.Misc, db_path: str, *, verify_tls: bool = True) -> None:
        self.root = root
        self.db_path = os.path.abspath(db_path)
        self.verify_tls = verify_tls
        self._q: "queue.Queue[Any]" = queue.Queue()
        self._failures = 0
        self._after_id: Optional[str] = None
        self._thr: Optional[threading.Thread] = None
        self._wait_thr: Optional[threading.Thread] = None
        self._last_nudge_at = 0.0

    def start(self) -> None:
        self._schedule_after(self._startup_delay_ms())
        self._pump_queue()
        self._start_waiter()

    def _startup_delay_ms(self) -> int:
        """Stagger first sync attempts to reduce overlap across apps/devices."""
        seed = self.db_path
        try:
            conn = sqlite3.connect(self.db_path, timeout=30.0, isolation_level=None, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            try:
                conn.execute("PRAGMA busy_timeout=30000;")
                cfg = sync_client.load_sync_config(conn)
                if (cfg or {}).get("device_name"):
                    seed = str(cfg.get("device_name"))
            finally:
                conn.close()
        except Exception:
            pass
        slot = zlib.crc32(seed.encode("utf-8")) % OFFSET_SLOT_COUNT
        return FIRST_DELAY_BASE_MS + (slot * OFFSET_STEP_MS)

    def _schedule_after(self, delay_ms: int) -> None:
        try:
            if self._after_id is not None:
                self.root.after_cancel(self._after_id)
        except Exception:
            pass
        delay_ms = max(3_000, int(delay_ms))
        self._after_id = self.root.after(delay_ms, self._tick)

    def _next_interval_ms(self) -> int:
        return DEFAULT_BASE_MS

    def _pump_queue(self) -> None:
        try:
            while True:
                msg = self._q.get_nowait()
                kind = msg[0]
                if kind == "summary":
                    try:
                        import sync_ui

                        sync_ui._notify_host_synced(self.root)
                    except Exception:
                        pass
        except queue.Empty:
            pass
        try:
            self.root.after(250, self._pump_queue)
        except Exception:
            pass

    def _tick(self) -> None:
        self._after_id = None
        if self._thr is not None and self._thr.is_alive():
            self._schedule_after(60_000)
            return

        try:
            conn = sqlite3.connect(self.db_path, timeout=30.0, isolation_level=None, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            try:
                conn.execute("PRAGMA busy_timeout=30000;")
                cfg = sync_client.load_sync_config(conn)
            finally:
                conn.close()
        except Exception:
            cfg = {}

        if not (cfg or {}).get("server_url") or not (cfg or {}).get("device_name"):
            self._schedule_after(self._next_interval_ms())
            return

        def worker() -> None:
            summary: Optional[Dict[str, Any]] = None
            err: Optional[str] = None
            try:
                conn2 = sqlite3.connect(self.db_path, timeout=30.0, isolation_level=None, check_same_thread=False)
                conn2.row_factory = sqlite3.Row
                try:
                    conn2.execute("PRAGMA busy_timeout=30000;")
                    client = sync_client.SyncClient(conn2, verify_tls=self.verify_tls)
                    summary = client.run_cycle(progress=None)
                finally:
                    conn2.close()
            except sync_client.SyncError as e:
                err = str(e)
            except Exception as e:
                err = f"unexpected: {e}"

            def post() -> None:
                if err:
                    self._failures += 1
                else:
                    self._failures = 0
                    self._q.put(("summary", summary or {}))
                self._schedule_after(self._next_interval_ms())

            try:
                self.root.after(0, post)
            except Exception:
                pass

        self._thr = threading.Thread(target=worker, daemon=True)
        self._thr.start()

    def _start_waiter(self) -> None:
        if self._wait_thr is not None and self._wait_thr.is_alive():
            return

        def waiter() -> None:
            while True:
                try:
                    conn = sqlite3.connect(self.db_path, timeout=30.0, isolation_level=None, check_same_thread=False)
                    conn.row_factory = sqlite3.Row
                    try:
                        conn.execute("PRAGMA busy_timeout=30000;")
                        cfg = sync_client.load_sync_config(conn)
                        if not (cfg or {}).get("server_url") or not (cfg or {}).get("device_name"):
                            time.sleep(10.0)
                            continue
                        client = sync_client.SyncClient(conn, verify_tls=self.verify_tls)
                        res = client.wait_for_updates(timeout_s=25)
                    finally:
                        conn.close()
                except Exception:
                    time.sleep(4.0)
                    continue

                if not bool((res or {}).get("has_updates")):
                    continue

                def _wake() -> None:
                    now = time.monotonic()
                    if now - self._last_nudge_at < 2.0:
                        return
                    self._last_nudge_at = now
                    if self._thr is not None and self._thr.is_alive():
                        return
                    self._schedule_after(250)

                try:
                    self.root.after(0, _wake)
                except Exception:
                    return

        self._wait_thr = threading.Thread(target=waiter, daemon=True)
        self._wait_thr.start()


def attach_periodic_sync(root: tk.Misc, db_path: str, *, verify_tls: bool = True) -> None:
    """Start at most one periodic sync controller on ``root``."""
    attr = "_periodic_sync_controller_v1"
    if getattr(root, attr, None) is not None:
        return
    ctl = PeriodicSyncController(root, db_path, verify_tls=verify_tls)
    setattr(root, attr, ctl)
    ctl.start()
