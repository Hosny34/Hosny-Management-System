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
import shutil
import subprocess
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path


LIVE_EXE = "HosnyPOS.exe"
LIVE_CERT = "cacert.pem"
BUILD_BAT = "build_pos.bat"
SHORTCUT_NAME = "Hosny POS.lnk"


def _log(path: Path, message: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with path.open("a", encoding="utf-8") as f:
        f.write("[%s] %s\n" % (stamp, message))


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


def _run_smoke_test(exe_path: Path, cwd: Path, log_path: Path) -> None:
    _log(log_path, "verifying prebuilt executable: %s" % exe_path)
    result = subprocess.call(
        [str(exe_path), "--smoke-import"],
        cwd=str(cwd),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=False,
    )
    if result != 0:
        raise RuntimeError("prebuilt HosnyPOS.exe smoke test failed with exit code %s" % result)


def _prepare_prebuilt_dist(source_dir: Path, pos_dir: Path, log_path: Path) -> bool:
    marker = source_dir / "prebuilt_update.json"
    source_exe = source_dir / "prebuilt" / LIVE_EXE
    if not source_exe.is_file():
        source_exe = source_dir / LIVE_EXE
    if not marker.is_file() or not source_exe.is_file():
        return False
    _run_smoke_test(source_exe, source_dir, log_path)
    dist_dir = pos_dir / "dist"
    dist_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(source_exe), str(dist_dir / LIVE_EXE))
    source_cert = source_dir / LIVE_CERT
    if source_cert.is_file():
        shutil.copy2(str(source_cert), str(dist_dir / LIVE_CERT))
    _log(log_path, "prebuilt package prepared in dist")
    return True


def _wait_for_file_release(path: Path, log_path: Path, timeout_s: int = 60) -> None:
    if not path.exists():
        return
    deadline = time.time() + max(1, timeout_s)
    last_error = None
    forced = False
    while time.time() < deadline:
        try:
            probe = path.with_name(path.name + ".locktest")
            path.rename(probe)
            probe.rename(path)
            return
        except OSError as ex:
            last_error = ex
            if not forced and time.time() > deadline - max(1, timeout_s) + 5:
                forced = True
                _log(log_path, "live executable is still locked; closing remaining HosnyPOS.exe processes")
                try:
                    subprocess.call(
                        ["taskkill.exe", "/IM", LIVE_EXE, "/T", "/F"],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        shell=False,
                    )
                except Exception as kill_ex:
                    _log(log_path, "taskkill by image name failed: %s" % kill_ex)
            time.sleep(0.5)
    raise RuntimeError("file is still locked after waiting: %s (%s)" % (path, last_error))


def _promote_dist(pos_dir: Path, log_path: Path) -> None:
    dist_exe = pos_dir / "dist" / LIVE_EXE
    dist_cert = pos_dir / "dist" / LIVE_CERT
    if not dist_exe.is_file():
        raise RuntimeError("dist/HosnyPOS.exe was not created")
    _wait_for_file_release(pos_dir / LIVE_EXE, log_path)
    shutil.copy2(str(dist_exe), str(pos_dir / LIVE_EXE))
    if dist_cert.is_file():
        shutil.copy2(str(dist_cert), str(pos_dir / LIVE_CERT))
    _log(log_path, "dist output promoted")


def _create_shortcut(pos_dir: Path, log_path: Path) -> None:
    try:
        import win32com.client  # type: ignore

        desktop = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))
        shortcut_path = desktop / SHORTCUT_NAME
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(str(shortcut_path))
        shortcut.TargetPath = str(pos_dir / LIVE_EXE)
        shortcut.WorkingDirectory = str(pos_dir)
        shortcut.IconLocation = str(pos_dir / LIVE_EXE)
        shortcut.Save()
        _log(log_path, "shortcut updated: %s" % shortcut_path)
    except Exception as ex:
        _log(log_path, "shortcut update failed: %s" % ex)


def _restart_pos(pos_dir: Path, log_path: Path) -> None:
    exe_path = pos_dir / LIVE_EXE
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


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pos-dir", required=True)
    parser.add_argument("--package", required=True)
    parser.add_argument("--pid", type=int, default=0)
    parser.add_argument("--version", default="")
    parser.add_argument("--restart", action="store_true")
    args = parser.parse_args(argv)

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
        if not _prepare_prebuilt_dist(source_dir, pos_dir, log_path):
            _run_build(pos_dir, log_path, args.version or "unknown")
        _promote_dist(pos_dir, log_path)
        _create_shortcut(pos_dir, log_path)
        if args.version:
            with (pos_dir / "current_version.json").open("w", encoding="utf-8") as f:
                json.dump({"version": args.version, "updated_at": datetime.now().isoformat()}, f)
        if args.restart:
            _restart_pos(pos_dir, log_path)
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


if __name__ == "__main__":
    raise SystemExit(main())
