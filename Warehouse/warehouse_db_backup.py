# -*- coding: utf-8 -*-
"""Hourly local SQLite backups for Hosny Warehouse."""

from __future__ import annotations

import logging
import os
import queue
import re
import sqlite3
import threading
import time
import tkinter as tk
from pathlib import Path
from typing import Any, Optional


BACKUP_INTERVAL_MS = 60 * 60 * 1000
FIRST_BACKUP_DELAY_MS = 15 * 60 * 1000

LOG = logging.getLogger("hosny")


def _safe_name(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"[^A-Za-z0-9_.-]+", "-", text).strip("-_.")
    return text[:40] or "WAREHOUSE"


def _device_name(db_path: str) -> str:
    try:
        conn = sqlite3.connect(os.path.abspath(db_path), timeout=5.0)
        try:
            row = conn.execute(
                "SELECT device_name FROM device_identity WHERE id = 1"
            ).fetchone()
            if row and row[0]:
                return _safe_name(row[0])
        finally:
            conn.close()
    except Exception:
        pass
    return "WAREHOUSE"


def _integrity_ok(conn: sqlite3.Connection) -> bool:
    try:
        row = conn.execute("PRAGMA quick_check").fetchone()
        return bool(row and str(row[0]).lower() == "ok")
    except Exception:
        return False


def create_backup(db_path: str, backup_root: Optional[str] = None) -> str:
    source_path = os.path.abspath(db_path)
    if not os.path.isfile(source_path):
        raise FileNotFoundError(source_path)

    base_dir = os.path.dirname(source_path)
    root = Path(backup_root or os.path.join(base_dir, "db_backups", "hourly"))
    root.mkdir(parents=True, exist_ok=True)

    device = _device_name(source_path)
    tmp_path = root / ("warehouse_data-%s-latest.sqlite3.tmp" % device)
    final_path = root / ("warehouse_data-%s-latest.sqlite3" % device)

    src = sqlite3.connect(source_path, timeout=30.0, isolation_level=None)
    try:
        src.execute("PRAGMA busy_timeout=30000;")
        if not _integrity_ok(src):
            raise RuntimeError("source database quick_check failed")
        dst = sqlite3.connect(str(tmp_path), timeout=30.0)
        try:
            src.backup(dst, pages=256, sleep=0.05)
            dst.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            if not _integrity_ok(dst):
                raise RuntimeError("backup database quick_check failed")
        finally:
            dst.close()
    finally:
        src.close()

    os.replace(str(tmp_path), str(final_path))
    return str(final_path)


def prune_backups(db_path: str, backup_root: Optional[str] = None) -> None:
    base_dir = os.path.dirname(os.path.abspath(db_path))
    root = Path(backup_root or os.path.join(base_dir, "db_backups", "hourly"))
    if not root.is_dir():
        return

    device = _device_name(db_path)
    keep_name = "warehouse_data-%s-latest.sqlite3" % device
    for path in root.glob("warehouse_data-*.sqlite3*"):
        try:
            if path.name != keep_name:
                path.unlink()
        except Exception:
            LOG.exception("failed to prune warehouse DB backup: %s", path)


class HourlyDbBackupController:
    def __init__(self, root: tk.Misc, db_path: str) -> None:
        self.root = root
        self.db_path = os.path.abspath(db_path)
        self._after_id: Optional[str] = None
        self._worker: Optional[threading.Thread] = None
        self._q: "queue.Queue[tuple]" = queue.Queue()

    def start(self) -> None:
        self._schedule(FIRST_BACKUP_DELAY_MS)
        self._pump_queue()

    def _schedule(self, delay_ms: int) -> None:
        try:
            if self._after_id is not None:
                self.root.after_cancel(self._after_id)
        except Exception:
            pass
        self._after_id = self.root.after(max(10_000, int(delay_ms)), self._tick)

    def _tick(self) -> None:
        self._after_id = None
        if bool(getattr(self.root, "_hosny_sync_active", False)):
            self._schedule(10 * 60 * 1000)
            return
        if self._worker is not None and self._worker.is_alive():
            self._schedule(5 * 60 * 1000)
            return

        def worker() -> None:
            started = time.monotonic()
            try:
                path = create_backup(self.db_path)
                prune_backups(self.db_path)
                self._q.put(("ok", path, time.monotonic() - started))
            except Exception as ex:
                self._q.put(("error", str(ex)))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()
        self._schedule(BACKUP_INTERVAL_MS)

    def _pump_queue(self) -> None:
        try:
            while True:
                msg = self._q.get_nowait()
                if msg[0] == "ok":
                    LOG.info("Warehouse DB hourly backup created: %s (%.1fs)", msg[1], msg[2])
                else:
                    LOG.error("Warehouse DB hourly backup failed: %s", msg[1])
        except queue.Empty:
            pass
        try:
            self.root.after(1000, self._pump_queue)
        except Exception:
            pass


def attach_hourly_db_backup(root: tk.Misc, db_path: str) -> None:
    attr = "_hourly_db_backup_controller_v1"
    if getattr(root, attr, None) is not None:
        return
    ctl = HourlyDbBackupController(root, db_path)
    setattr(root, attr, ctl)
    ctl.start()
