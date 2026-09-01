# -*- coding: utf-8 -*-
"""Separate updater process for Hosny POS.

This script is designed to run outside HosnyPOS.exe. It waits for POS to
close, unpacks staged source files, runs POS/build_pos.bat locally, promotes
dist/HosnyPOS.exe and dist/cacert.pem, recreates the desktop shortcut, and
restarts POS.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional


LIVE_EXE = "HosnyPOS.exe"
LIVE_CERT = "cacert.pem"
BUILD_BAT = "build_pos.bat"
SHORTCUT_NAME = "Hosny POS.lnk"
_PROGRESS_CALLBACK = None


def _log(path: Path, message: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with path.open("a", encoding="utf-8") as f:
        f.write("[%s] %s\n" % (stamp, message))
    if _PROGRESS_CALLBACK is not None:
        try:
            _PROGRESS_CALLBACK(message)
        except Exception:
            pass


def _wait_for_process(pid: int, timeout_s: int, log_path: Path) -> None:
    if pid <= 0:
        return
    try:
        synchronize = 0x00100000
        handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, int(pid))
        if not handle:
            return
        try:
            wait_ms = int(max(1, timeout_s) * 1000)
            result = ctypes.windll.kernel32.WaitForSingleObject(handle, wait_ms)
            if result == 0x00000102:
                _log(log_path, "POS process did not exit before timeout; forcing close")
                try:
                    subprocess.call(
                        ["taskkill.exe", "/PID", str(pid), "/T", "/F"],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        shell=False,
                    )
                    ctypes.windll.kernel32.WaitForSingleObject(handle, 10000)
                except Exception as ex:
                    _log(log_path, "forced close failed: %s" % ex)
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)
    except Exception as ex:
        _log(log_path, "could not wait for POS process: %s" % ex)


def _copy_tree_contents(src: Path, dst: Path, log_path: Path) -> None:
    for item in src.iterdir():
        if item.name in {"dist", "build", "backups", "__pycache__"}:
            continue
        if item.name == LIVE_EXE:
            continue
        target = dst / item.name
        if item.is_dir():
            if target.exists():
                shutil.rmtree(str(target))
            shutil.copytree(str(item), str(target))
        else:
            shutil.copy2(str(item), str(target))
        _log(log_path, "copied %s" % item.name)


def _backup_files(pos_dir: Path, source_dir: Path, version: str, log_path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = pos_dir / "backups" / ("%s_%s" % (version, stamp))
    backup_dir.mkdir(parents=True, exist_ok=True)
    names = {LIVE_EXE, LIVE_CERT, BUILD_BAT}
    for item in source_dir.iterdir():
        names.add(item.name)
    for name in sorted(names):
        src = pos_dir / name
        if not src.exists():
            continue
        dst = backup_dir / name
        if src.is_dir():
            shutil.copytree(str(src), str(dst))
        else:
            shutil.copy2(str(src), str(dst))
    _log(log_path, "backup created at %s" % backup_dir)
    return backup_dir


def _restore_backup(backup_dir: Path, pos_dir: Path, log_path: Path) -> None:
    if not backup_dir.exists():
        return
    for item in backup_dir.iterdir():
        target = pos_dir / item.name
        if target.exists():
            if target.is_dir():
                shutil.rmtree(str(target))
            else:
                target.unlink()
        if item.is_dir():
            shutil.copytree(str(item), str(target))
        else:
            shutil.copy2(str(item), str(target))
    _log(log_path, "backup restored from %s" % backup_dir)


def _extract_package(package_path: Path, work_dir: Path) -> Path:
    extract_dir = work_dir / "source"
    if extract_dir.exists():
        shutil.rmtree(str(extract_dir))
    extract_dir.mkdir(parents=True)
    with zipfile.ZipFile(str(package_path), "r") as zf:
        zf.extractall(str(extract_dir))
    nested = extract_dir / "POS"
    return nested if nested.is_dir() else extract_dir


def _run_build(pos_dir: Path, log_path: Path, version: str = "") -> None:
    build_bat = pos_dir / BUILD_BAT
    if not build_bat.is_file():
        raise RuntimeError("build_pos.bat was not found in POS folder")
    _log(log_path, "running %s" % build_bat)
    env = os.environ.copy()
    env["HOSNY_AUTO_UPDATE"] = "1"
    for name in ("TCL_LIBRARY", "TK_LIBRARY"):
        if name in env:
            _log(log_path, "clearing inherited %s before build" % name)
            env.pop(name, None)
    safe_version = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in str(version or "unknown"))
    build_log = pos_dir / "logs" / ("pos_update_build_%s.log" % safe_version)
    build_log.parent.mkdir(parents=True, exist_ok=True)
    _log(log_path, "build output log: %s" % build_log)
    with build_log.open("w", encoding="utf-8", errors="replace") as out:
        out.write("Running %s\n\n" % build_bat)
        result = subprocess.call(
            ["cmd.exe", "/c", str(build_bat)],
            cwd=str(pos_dir),
            env=env,
            stdout=out,
            stderr=subprocess.STDOUT,
            shell=False,
        )
    if result != 0:
        raise RuntimeError(
            "build_pos.bat failed with exit code %s. See %s. "
            "Make sure this POS PC has a Python installation that can import tkinter."
            % (result, build_log)
        )


def _close_remaining_live_processes(log_path: Path) -> None:
    _log(log_path, "closing remaining HosnyPOS.exe processes")
    try:
        result = subprocess.call(
            ["taskkill.exe", "/IM", LIVE_EXE, "/T", "/F"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
        )
        _log(log_path, "taskkill HosnyPOS.exe exit code: %s" % result)
    except Exception as kill_ex:
        _log(log_path, "taskkill by image name failed: %s" % kill_ex)


def _wait_for_file_release(path: Path, log_path: Path, timeout_s: int = 20) -> bool:
    if not path.exists():
        return True
    deadline = time.time() + max(1, timeout_s)
    last_error = None
    forced = False
    while time.time() < deadline:
        try:
            probe = path.with_name(path.name + ".locktest")
            path.rename(probe)
            probe.rename(path)
            return True
        except OSError as ex:
            last_error = ex
            if not forced and time.time() > deadline - max(1, timeout_s) + 5:
                forced = True
                _log(log_path, "live executable is still locked before promotion")
                _close_remaining_live_processes(log_path)
            time.sleep(0.5)
    _log(log_path, "file is still locked after waiting: %s (%s)" % (path, last_error))
    return False


def _copy_exe_with_retries(src: Path, dst: Path, log_path: Path, attempts: int = 8) -> bool:
    last_error = None
    for attempt in range(1, max(1, attempts) + 1):
        try:
            shutil.copy2(str(src), str(dst))
            _log(log_path, "executable promoted to %s" % dst)
            return True
        except OSError as ex:
            last_error = ex
            if attempt == 2:
                _close_remaining_live_processes(log_path)
            _log(log_path, "copy executable attempt %s failed: %s" % (attempt, ex))
            time.sleep(1.0)
    _log(log_path, "copy executable failed after retries: %s" % last_error)
    return False


def _promote_dist(pos_dir: Path, log_path: Path, version: str = "") -> Path:
    dist_exe = pos_dir / "dist" / LIVE_EXE
    dist_cert = pos_dir / "dist" / LIVE_CERT
    if not dist_exe.is_file():
        raise RuntimeError("dist/HosnyPOS.exe was not created")
    live_exe = pos_dir / LIVE_EXE
    promoted_exe = live_exe
    _wait_for_file_release(live_exe, log_path)
    if not _copy_exe_with_retries(dist_exe, live_exe, log_path):
        safe_version = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in str(version or "updated"))
        fallback_exe = pos_dir / ("HosnyPOS-%s.exe" % safe_version)
        if fallback_exe.exists():
            fallback_exe = pos_dir / ("HosnyPOS-%s-%s.exe" % (safe_version, datetime.now().strftime("%Y%m%d_%H%M%S")))
        shutil.copy2(str(dist_exe), str(fallback_exe))
        promoted_exe = fallback_exe
        _log(
            log_path,
            "live executable stayed locked; using fallback executable for this update: %s" % fallback_exe,
        )
    if dist_cert.is_file():
        shutil.copy2(str(dist_cert), str(pos_dir / LIVE_CERT))
    _log(log_path, "dist output promoted")
    return promoted_exe


def _create_shortcut(pos_dir: Path, log_path: Path, exe_path: Optional[Path] = None) -> None:
    target_exe = exe_path or (pos_dir / LIVE_EXE)
    try:
        import win32com.client  # type: ignore

        shell = win32com.client.Dispatch("WScript.Shell")
        desktop_paths = []
        for raw_path in (
            shell.SpecialFolders("Desktop"),
            os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
            os.path.join(os.environ.get("PUBLIC", ""), "Desktop"),
        ):
            if not raw_path:
                continue
            path = Path(raw_path)
            if path not in desktop_paths:
                desktop_paths.append(path)
        existing = [desktop / SHORTCUT_NAME for desktop in desktop_paths if (desktop / SHORTCUT_NAME).exists()]
        target_path = existing[0] if existing else (desktop_paths[0] / SHORTCUT_NAME if desktop_paths else None)
        if target_path is None:
            raise RuntimeError("no desktop shortcut location could be updated")
        try:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            shortcut = shell.CreateShortcut(str(target_path))
            shortcut.TargetPath = str(target_exe)
            shortcut.WorkingDirectory = str(pos_dir)
            shortcut.IconLocation = str(target_exe)
            shortcut.Save()
            _log(log_path, "shortcut updated: %s" % target_path)
        except Exception as item_ex:
            raise RuntimeError("shortcut update failed for %s: %s" % (target_path, item_ex))
        for shortcut_path in existing[1:]:
            try:
                shortcut_path.unlink()
                _log(log_path, "duplicate shortcut removed: %s" % shortcut_path)
            except Exception as item_ex:
                _log(log_path, "duplicate shortcut remove failed for %s: %s" % (shortcut_path, item_ex))
    except Exception as ex:
        _log(log_path, "shortcut update failed: %s" % ex)


def _restart_pos(pos_dir: Path, log_path: Path, exe_path: Optional[Path] = None) -> None:
    exe_path = exe_path or (pos_dir / LIVE_EXE)
    if not exe_path.is_file():
        raise RuntimeError("cannot restart POS because HosnyPOS.exe was not found")
    try:
        proc = subprocess.Popen(
            [str(exe_path)],
            cwd=str(pos_dir),
            close_fds=True,
        )
        _log(log_path, "restart launched: pid=%s path=%s" % (getattr(proc, "pid", ""), exe_path))
        return
    except Exception as ex:
        _log(log_path, "restart via Popen failed: %s" % ex)

    if os.name == "nt":
        try:
            os.startfile(str(exe_path))  # type: ignore[attr-defined]
            _log(log_path, "restart launched via os.startfile: %s" % exe_path)
            return
        except Exception as ex:
            _log(log_path, "restart via os.startfile failed: %s" % ex)

    raise RuntimeError("failed to restart POS")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pos-dir", required=True)
    parser.add_argument("--package", required=True)
    parser.add_argument("--pid", type=int, default=0)
    parser.add_argument("--version", default="")
    parser.add_argument("--restart", action="store_true")
    parser.add_argument("--gui", action="store_true")
    return parser


def _run_update(args) -> int:
    pos_dir = Path(args.pos_dir).resolve()
    package_path = Path(args.package).resolve()
    log_path = pos_dir / "logs" / "pos_updater.log"
    backup_dir = None

    try:
        _log(log_path, "update started")
        _wait_for_process(args.pid, 30, log_path)
        work_dir = package_path.parent / "work"
        source_dir = _extract_package(package_path, work_dir)
        backup_dir = _backup_files(pos_dir, source_dir, args.version or "unknown", log_path)
        _copy_tree_contents(source_dir, pos_dir, log_path)
        _run_build(pos_dir, log_path, args.version or "unknown")
        promoted_exe = _promote_dist(pos_dir, log_path, args.version or "unknown")
        _create_shortcut(pos_dir, log_path, promoted_exe)
        if args.version:
            with (pos_dir / "current_version.json").open("w", encoding="utf-8") as f:
                json.dump({"version": args.version, "updated_at": datetime.now().isoformat()}, f)
        if args.restart:
            _restart_pos(pos_dir, log_path, promoted_exe)
        _log(log_path, "update completed")
        return 0
    except Exception as ex:
        _log(log_path, "update failed: %s" % ex)
        try:
            if backup_dir is not None:
                _restore_backup(backup_dir, pos_dir, log_path)
        except Exception as restore_ex:
            _log(log_path, "rollback failed: %s" % restore_ex)
        if args.restart and (pos_dir / LIVE_EXE).is_file():
            try:
                _restart_pos(pos_dir, log_path)
            except Exception as restart_ex:
                _log(log_path, "restart after rollback failed: %s" % restart_ex)
        return 1


def _run_update_gui(args) -> int:
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:
        return _run_update(args)

    global _PROGRESS_CALLBACK
    events = queue.Queue()
    done = {"rc": None}

    def publish(message: str) -> None:
        events.put(("log", message))

    _PROGRESS_CALLBACK = publish

    root = tk.Tk()
    root.title("Hosny POS Update")
    root.geometry("520x300")
    root.resizable(False, False)
    root.protocol("WM_DELETE_WINDOW", lambda: None)

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill="both", expand=True)

    title = ttk.Label(frame, text="جاري تحديث نقطة البيع", font=("Tahoma", 13, "bold"))
    title.pack(anchor="e")
    status = ttk.Label(frame, text="يتم تجهيز التحديث...", font=("Tahoma", 10))
    status.pack(anchor="e", pady=(8, 8))

    progress = ttk.Progressbar(frame, mode="indeterminate")
    progress.pack(fill="x", pady=(0, 10))
    progress.start(12)

    text = tk.Text(frame, height=9, wrap="word", state="disabled", font=("Consolas", 9))
    text.pack(fill="both", expand=True)

    close_button = ttk.Button(frame, text="إغلاق", command=root.destroy)
    close_button.pack(anchor="w", pady=(10, 0))
    close_button.configure(state="disabled")

    def append_line(message: str) -> None:
        status.configure(text=message)
        text.configure(state="normal")
        text.insert("end", message + "\n")
        text.see("end")
        text.configure(state="disabled")

    def worker() -> None:
        rc = 1
        try:
            rc = _run_update(args)
        finally:
            events.put(("done", rc))

    def pump() -> None:
        try:
            while True:
                kind, payload = events.get_nowait()
                if kind == "log":
                    append_line(str(payload))
                elif kind == "done":
                    done["rc"] = int(payload)
                    progress.stop()
                    if int(payload) == 0:
                        status.configure(text="تم التحديث بنجاح. سيتم تشغيل نقطة البيع الآن.")
                        root.after(1800, root.destroy)
                    else:
                        status.configure(text="فشل التحديث. تم تسجيل التفاصيل في ملف السجل.")
                        close_button.configure(state="normal")
                        root.after(8000, root.destroy)
        except queue.Empty:
            pass
        if done["rc"] is None:
            root.after(100, pump)

    threading.Thread(target=worker, daemon=True).start()
    root.after(100, pump)
    root.mainloop()
    _PROGRESS_CALLBACK = None
    return int(done["rc"] if done["rc"] is not None else 1)


def main(argv=None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    if getattr(args, "gui", False):
        return _run_update_gui(args)
    return _run_update(args)


if __name__ == "__main__":
    raise SystemExit(main())
