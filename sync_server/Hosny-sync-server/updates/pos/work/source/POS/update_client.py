# -*- coding: utf-8 -*-
"""POS update check/download helper.

Uses the existing sync server credentials from the local POS database.
Installation is intentionally handled by a separate updater process.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Any, Dict, Optional

import sync_client
from pos_version import APP_VERSION


class UpdateError(RuntimeError):
    pass


def _version_key(value: Any) -> tuple:
    text = str(value or "").strip()
    parts = re.findall(r"\d+", text)
    return tuple(int(p) for p in parts) if parts else (0,)


def is_newer_version(remote: Any, local: Any = APP_VERSION) -> bool:
    return _version_key(remote) > _version_key(local)


def _auth_token(conn: sqlite3.Connection, verify_tls: bool = True) -> tuple:
    cfg = sync_client.load_sync_config(conn)
    if not cfg.get("server_url") or not cfg.get("device_name"):
        raise UpdateError("sync is not configured")
    client = sync_client.SyncClient(conn, verify_tls=verify_tls)
    token = client._fetch_jwt(cfg)  # same auth path used by normal sync
    return cfg, token


def check_for_update(
    conn: sqlite3.Connection,
    *,
    verify_tls: bool = True,
) -> Dict[str, Any]:
    cfg, token = _auth_token(conn, verify_tls=verify_tls)
    url = cfg["server_url"].rstrip("/") + "/v1/updates/pos/latest"
    status, body = sync_client._http_request(
        "GET",
        url,
        token=token,
        verify_tls=verify_tls,
    )
    if status != 200:
        raise UpdateError("update check failed: HTTP %s" % status)
    manifest = body.get("manifest") if body.get("available") else None
    remote_version = (manifest or {}).get("version")
    return {
        "local_version": APP_VERSION,
        "available": bool(manifest) and is_newer_version(remote_version, APP_VERSION),
        "manifest": manifest,
        "server_time": body.get("server_time"),
    }


def _expected_sha256(manifest: Dict[str, Any]) -> Optional[str]:
    hashes = manifest.get("hashes")
    if isinstance(hashes, dict):
        value = hashes.get("package_sha256") or hashes.get("sha256")
        return str(value).strip().lower() if value else None
    value = manifest.get("package_sha256")
    return str(value).strip().lower() if value else None


def _package_url(cfg: Dict[str, Any], manifest: Dict[str, Any]) -> str:
    explicit = str(manifest.get("package_url") or "").strip()
    if explicit:
        return explicit
    filename = str(manifest.get("package_file") or "").strip()
    if not filename:
        raise UpdateError("update manifest does not define package_file")
    return (
        cfg["server_url"].rstrip("/")
        + "/v1/updates/pos/package/"
        + urllib.parse.quote(os.path.basename(filename))
    )


def download_update_package(
    conn: sqlite3.Connection,
    manifest: Dict[str, Any],
    *,
    staging_root: Optional[str] = None,
    verify_tls: bool = True,
) -> Dict[str, Any]:
    cfg, token = _auth_token(conn, verify_tls=verify_tls)
    version = str(manifest.get("version") or "").strip()
    if not version:
        raise UpdateError("update manifest does not define version")

    root = staging_root or os.path.join(tempfile.gettempdir(), "HosnyPOSUpdates")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target_dir = os.path.abspath(os.path.join(root, version + "_" + stamp))
    os.makedirs(target_dir, exist_ok=True)

    package_name = os.path.basename(str(manifest.get("package_file") or ("HosnyPOS-" + version + ".zip")))
    package_path = os.path.join(target_dir, package_name)
    url = _package_url(cfg, manifest)

    headers = {"Authorization": "Bearer " + token}
    req = urllib.request.Request(url, headers=headers, method="GET")
    ctx = sync_client._build_ssl_ctx(verify_tls) if url.lower().startswith("https://") else None
    digest = hashlib.sha256()
    try:
        with urllib.request.urlopen(req, timeout=120, context=ctx) as resp:
            with open(package_path, "wb") as f:
                while True:
                    chunk = resp.read(1024 * 1024)
                    if not chunk:
                        break
                    digest.update(chunk)
                    f.write(chunk)
    except Exception as ex:
        raise UpdateError("failed to download update package: %s" % ex)

    actual = digest.hexdigest().lower()
    expected = _expected_sha256(manifest)
    if expected and actual != expected:
        raise UpdateError("update package hash mismatch")

    manifest_path = os.path.join(target_dir, "update_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2, sort_keys=True)

    return {
        "version": version,
        "staging_dir": target_dir,
        "package_path": package_path,
        "manifest_path": manifest_path,
        "sha256": actual,
    }


def _runtime_pos_dir() -> str:
    if getattr(sys, "frozen", False):
        exe = os.path.abspath(getattr(sys, "executable", ""))
        if exe:
            return os.path.dirname(exe)
    return os.path.dirname(os.path.abspath(__file__))


def _append_update_launch_log(pos_dir: str, message: str) -> None:
    try:
        log_dir = os.path.join(pos_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        path = os.path.join(log_dir, "pos_update_launch.log")
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(path, "a", encoding="utf-8") as f:
            f.write("[%s] %s\n" % (stamp, message))
    except Exception:
        pass


def _candidate_python_commands() -> list:
    return [
        ["py", "-3.10"],
        ["py", "-3"],
        ["python"],
    ]


def _find_python_command() -> list:
    for cmd in _candidate_python_commands():
        try:
            rc = subprocess.call(
                cmd + ["-V"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                shell=False,
            )
            if rc == 0:
                return cmd
        except Exception:
            continue
    raise UpdateError("no Python launcher was found to start the updater")


def launch_updater(download_result: Dict[str, Any], *, restart: bool = True) -> Dict[str, Any]:
    pos_dir = _runtime_pos_dir()
    updater_script = os.path.join(pos_dir, "pos_updater.py")
    package_path = os.path.abspath(str(download_result.get("package_path") or ""))
    version = str(download_result.get("version") or "").strip()

    if not os.path.isfile(updater_script):
        raise UpdateError("pos_updater.py was not found in the POS folder")
    if not os.path.isfile(package_path):
        raise UpdateError("downloaded update package was not found")
    if not version:
        raise UpdateError("downloaded update result does not include version")

    cmd = (
        _find_python_command()
        + [
            updater_script,
            "--pos-dir",
            pos_dir,
            "--package",
            package_path,
            "--pid",
            str(os.getpid()),
            "--version",
            version,
        ]
    )
    if restart:
        cmd.append("--restart")

    creationflags = 0
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    _append_update_launch_log(pos_dir, "launching updater: %s" % " ".join('"%s"' % c if " " in c else c for c in cmd))
    proc = subprocess.Popen(
        cmd,
        cwd=pos_dir,
        close_fds=True,
        creationflags=creationflags,
    )
    _append_update_launch_log(pos_dir, "updater process started: pid=%s" % getattr(proc, "pid", ""))
    return {
        "version": version,
        "pos_dir": pos_dir,
        "package_path": package_path,
        "command": cmd,
        "pid": getattr(proc, "pid", None),
    }
