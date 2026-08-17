# -*- coding: utf-8 -*-
"""Persistent crash / error logging for the desktop apps.

Usage (at the top of the entry script, before any Tk code runs):

    import logging_setup
    logging_setup.install_crash_logging("HosnyWarehouse")

Behaviour:

- Writes rotating log files to ``<exe-dir>/logs/<app>.log`` (up to 5 x 2 MB).
- Captures ``sys.excepthook`` so uncaught exceptions are recorded.
- Overrides ``tk.Tk.report_callback_exception`` / ``tk.Misc.report_callback_exception``
  so errors raised inside Tk event callbacks (where GUI crashes usually live)
  are recorded and surfaced as a messagebox instead of silently killing the app.
- Re-points ``sys.stdout`` and ``sys.stderr`` into the logger, so the many
  existing ``traceback.print_exc()`` calls also end up in the log file.
"""

from __future__ import annotations

import logging
import logging.handlers
import os
import platform
import socket
import sys
import time
from contextlib import contextmanager
from typing import Any, Dict, Iterator, Optional


_INSTALLED = False
_APP_NAME = "app"
_LOG_PATH: Optional[str] = None
_LOGGER: Optional[logging.Logger] = None
_CRASH_POPUP_SHOWN = False
_CONTEXT: Dict[str, Any] = {}
_SENSITIVE_KEYS = {"password", "api_key", "apikey", "token", "secret", "authorization", "key"}


def _resolve_base_dir() -> str:
    """Where to put the ``logs/`` folder.

    - Frozen (PyInstaller) build: the directory containing the .exe.
    - Source run: the directory containing the entry script (``sys.argv[0]``)
      or, as a last resort, this file's own directory.
    """
    if getattr(sys, "frozen", False):
        exe_path = os.path.abspath(getattr(sys, "executable", ""))
        if exe_path:
            return os.path.dirname(exe_path)
    script = os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else ""
    if script and os.path.isfile(script):
        return os.path.dirname(script)
    return os.path.dirname(os.path.abspath(__file__))


class _LogStream:
    """Minimal file-like wrapper that routes writes into a logger."""

    def __init__(self, logger: logging.Logger, level: int) -> None:
        self._logger = logger
        self._level = level
        self._buf = ""

    def write(self, s):  # type: ignore[override]
        if s is None:
            return 0
        if not isinstance(s, str):
            try:
                s = s.decode("utf-8", "replace")
            except Exception:
                s = str(s)
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            line = line.rstrip()
            if line:
                try:
                    self._logger.log(self._level, line)
                except Exception:
                    pass
        return len(s)

    def flush(self):  # type: ignore[override]
        if self._buf.strip():
            try:
                self._logger.log(self._level, self._buf.rstrip())
            except Exception:
                pass
        self._buf = ""

    def isatty(self):  # type: ignore[override]
        return False


def _show_crash_popup(exc_type, exc_value) -> None:
    """Best-effort error popup. Silent if Tk is not available yet."""
    global _CRASH_POPUP_SHOWN
    if _CRASH_POPUP_SHOWN:
        return
    summary = f"{getattr(exc_type, '__name__', 'Error')}: {exc_value}"
    path = _LOG_PATH or "logs"
    title = "خطأ غير متوقع"
    body = f"حدث خطأ غير متوقع:\n{summary}\n\nتم حفظ التفاصيل في:\n{path}"
    try:
        import tkinter as tk
        from tkinter import messagebox

        # Avoid Tk popups after teardown (prevents: application has been destroyed).
        root = tk._default_root  # type: ignore[attr-defined]
        if root is None:
            raise RuntimeError("Tk root is unavailable")
        if not bool(root.winfo_exists()):
            raise RuntimeError("Tk root is already destroyed")
        messagebox.showerror(
            title,
            body,
            parent=root,
        )
        _CRASH_POPUP_SHOWN = True
    except Exception:
        # Fall back to a native Windows popup when Tk cannot show dialogs.
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(0, body, title, 0x10)
            _CRASH_POPUP_SHOWN = True
        except Exception:
            pass


def _safe_detail(value: Any, depth: int = 0) -> Any:
    if depth > 4:
        return "<max-depth>"
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, dict):
        out: Dict[str, Any] = {}
        for key, item in value.items():
            key_text = str(key)
            if key_text.strip().lower() in _SENSITIVE_KEYS:
                out[key_text] = "<redacted>"
            else:
                out[key_text] = _safe_detail(item, depth + 1)
        return out
    if isinstance(value, (list, tuple, set)):
        seq = list(value)
        out = [_safe_detail(item, depth + 1) for item in seq[:80]]
        if len(seq) > 80:
            out.append({"truncated": len(seq) - 80})
        return out
    text = str(value)
    return text if len(text) <= 500 else text[:500] + "...<truncated>"


def _format_details(details: Dict[str, Any]) -> str:
    if not details:
        return ""
    try:
        import json

        return " | " + json.dumps(_safe_detail(details), ensure_ascii=False, sort_keys=True, default=str)
    except Exception:
        return " | " + str(details)


def configure_context(**context: Any) -> None:
    clean = {str(k): _safe_detail(v) for k, v in context.items() if v not in (None, "")}
    _CONTEXT.update(clean)
    get_logger().info("context updated%s", _format_details(clean))


def log_event(event: str, **details: Any) -> None:
    payload = dict(_CONTEXT)
    payload.update(details)
    get_logger().info("event=%s%s", event, _format_details(payload))


def log_exception(event: str, **details: Any) -> None:
    payload = dict(_CONTEXT)
    payload.update(details)
    get_logger().exception("event=%s%s", event, _format_details(payload))


@contextmanager
def log_operation(event: str, **details: Any) -> Iterator[None]:
    started = time.time()
    log_event(event + ".start", **details)
    try:
        yield
    except Exception:
        log_exception(event + ".failed", elapsed_ms=int((time.time() - started) * 1000), **details)
        raise
    else:
        log_event(event + ".done", elapsed_ms=int((time.time() - started) * 1000), **details)


def _show_crash_popup(exc_type, exc_value) -> None:
    """Best-effort error popup. Silent if Tk is not available yet."""
    global _CRASH_POPUP_SHOWN
    if _CRASH_POPUP_SHOWN:
        return
    summary = f"{getattr(exc_type, '__name__', 'Error')}: {exc_value}"
    path = _LOG_PATH or "logs"
    title = "خطأ غير متوقع"
    body = f"حدث خطأ غير متوقع:\n{summary}\n\nتم حفظ التفاصيل في:\n{path}"
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk._default_root  # type: ignore[attr-defined]
        if root is None or not bool(root.winfo_exists()):
            raise RuntimeError("Tk root is unavailable")
        messagebox.showerror(title, body, parent=root)
        _CRASH_POPUP_SHOWN = True
    except Exception:
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(0, body, title, 0x10)
            _CRASH_POPUP_SHOWN = True
        except Exception:
            pass


def install_crash_logging(app_name: str, **context: Any) -> str:
    """Install file logging + global exception hooks. Returns the log path."""
    global _INSTALLED, _APP_NAME, _LOG_PATH, _LOGGER
    if _INSTALLED:
        return _LOG_PATH or ""

    _APP_NAME = app_name or "app"
    base = _resolve_base_dir()
    log_dir = os.path.join(base, "logs")
    try:
        os.makedirs(log_dir, exist_ok=True)
    except Exception:
        log_dir = base
    _LOG_PATH = os.path.join(log_dir, f"{_APP_NAME}.log")

    logger = logging.getLogger("hosny")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    already_bound = any(
        isinstance(h, logging.handlers.RotatingFileHandler)
        and getattr(h, "_hosny_log", False)
        for h in logger.handlers
    )
    if not already_bound:
        try:
            handler = logging.handlers.RotatingFileHandler(
                _LOG_PATH,
                maxBytes=10_000_000,
                backupCount=10,
                encoding="utf-8",
                delay=True,
            )
            handler.setFormatter(
                logging.Formatter(
                    "%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s",
                    datefmt="%Y-%m-%d %H:%M:%S",
                )
            )
            handler._hosny_log = True  # type: ignore[attr-defined]
            logger.addHandler(handler)
        except Exception:
            # If we cannot open the log file we at least keep the hooks below.
            pass

    _LOGGER = logger
    if context:
        configure_context(**context)
    logger.info("=" * 60)
    logger.info(
        "%s starting up (frozen=%s, pid=%s, cwd=%s, exe=%s, python=%s, platform=%s, host=%s, log=%s)",
        _APP_NAME,
        bool(getattr(sys, "frozen", False)),
        os.getpid(),
        os.getcwd(),
        getattr(sys, "executable", ""),
        sys.version.replace("\n", " "),
        platform.platform(),
        socket.gethostname(),
        _LOG_PATH,
    )
    if sys.argv:
        logger.info("argv=%s", sys.argv)

    try:
        sys.stdout = _LogStream(logger, logging.INFO)
    except Exception:
        pass
    try:
        sys.stderr = _LogStream(logger, logging.ERROR)
    except Exception:
        pass

    def _excepthook(exc_type, exc_value, exc_tb):
        try:
            logger.error(
                "UNCAUGHT EXCEPTION",
                exc_info=(exc_type, exc_value, exc_tb),
            )
        except Exception:
            pass
        _show_crash_popup(exc_type, exc_value)

    sys.excepthook = _excepthook

    try:
        import threading

        def _thread_excepthook(args):
            try:
                logger.error(
                    "UNCAUGHT THREAD EXCEPTION in %s",
                    getattr(args.thread, "name", "?"),
                    exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
                )
            except Exception:
                pass

        threading.excepthook = _thread_excepthook  # type: ignore[attr-defined]
    except Exception:
        pass

    try:
        import tkinter as tk

        def _tk_report_exc(self, exc, val, tb):  # noqa: ARG001
            try:
                logger.error(
                    "TK CALLBACK EXCEPTION",
                    exc_info=(exc, val, tb),
                )
            except Exception:
                pass
            _show_crash_popup(exc, val)

        tk.Tk.report_callback_exception = _tk_report_exc
        try:
            tk.Misc.report_callback_exception = _tk_report_exc  # type: ignore[assignment]
        except Exception:
            pass
    except Exception:
        pass

    _INSTALLED = True
    return _LOG_PATH


def get_log_path() -> str:
    return _LOG_PATH or ""


def get_logger() -> logging.Logger:
    return _LOGGER or logging.getLogger("hosny")
