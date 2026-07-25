from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys
import urllib.error
import urllib.request
from typing import Dict, List, Optional, Tuple


MODE_CONFIG = {
    "pos-zay": {"device_name": "POS-ZAY", "target_scope": "pos:POS-ZAY", "wipe_local": True},
    "pos-oct": {"device_name": "POS-OCT", "target_scope": "pos:POS-OCT", "wipe_local": True},
    "pos-obo": {"device_name": "POS-OBO", "target_scope": "pos:POS-OBO", "wipe_local": True},
    "pos-gesr": {"device_name": "POS-GESR", "target_scope": "pos:POS-GESR", "wipe_local": True},
    "pos-bah": {"device_name": "POS-BAH", "target_scope": "pos:POS-BAH", "wipe_local": True},
    "pos-cen": {"device_name": "POS-CEN", "target_scope": "pos:POS-CEN", "wipe_local": True},
    "warehouse": {"device_name": "WAREHOUSE", "target_scope": "warehouse", "wipe_local": False},
    "all": {"device_name": "", "target_scope": "", "wipe_local": False},
}


def _script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _default_server_db() -> str:
    return os.path.abspath(
        os.path.join(_script_dir(), "..", "sync_server", "Hosny-sync-server", "sync_server.sqlite3")
    )


def _default_server_url() -> str:
    return (os.environ.get("SERVER_URL") or "https://web-production-e022.up.railway.app").strip().rstrip("/")


def _confirm(prompt: str) -> bool:
    reply = input("%s Type Y to continue: " % prompt).strip().upper()
    return reply == "Y"


def _post_json(url: str, body: Dict[str, object], headers: Dict[str, str]) -> Dict[str, object]:
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers=dict({"Content-Type": "application/json; charset=utf-8", "Accept": "application/json"}, **headers),
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw or "{}")
    except urllib.error.HTTPError as ex:
        try:
            detail = ex.read().decode("utf-8")
        except Exception:
            detail = str(ex)
        raise RuntimeError("HTTP %s: %s" % (ex.code, detail))
    except urllib.error.URLError as ex:
        raise RuntimeError("network error: %s" % ex.reason)


def _remote_reset(mode: str, device_name: str, target_scope: str) -> Dict[str, object]:
    base_url = _default_server_url()
    if mode == "all":
        return _post_json(base_url + "/v1/admin/reset-all", {"confirm": "RESET ALL"}, {})
    return _post_json(
        base_url + "/v1/admin/reset-device",
        {"device_name": device_name, "target_scope": target_scope},
        {},
    )


def _local_server_reset(mode: str, device_name: str, target_scope: str) -> Dict[str, object]:
    server_db = (os.environ.get("SERVER_DB") or _default_server_db()).strip()
    if not os.path.exists(server_db):
        raise RuntimeError("server DB not found: %s" % server_db)
    conn = sqlite3.connect(server_db)
    try:
        with conn:
            conn.execute("PRAGMA foreign_keys = OFF")
            if mode == "all":
                events_deleted = int(conn.execute("SELECT COUNT(*) FROM events").fetchone()[0] or 0)
                cursors_deleted = int(conn.execute("SELECT COUNT(*) FROM device_cursors").fetchone()[0] or 0)
                devices_deleted = int(conn.execute("SELECT COUNT(*) FROM devices").fetchone()[0] or 0)
                conn.execute("DELETE FROM events")
                conn.execute("DELETE FROM device_cursors")
                conn.execute("DELETE FROM devices")
            else:
                row = conn.execute("SELECT device_uuid FROM devices WHERE device_name = ?", (device_name,)).fetchone()
                device_uuid = str(row[0]) if row else ""
                events_deleted = int(
                    conn.execute(
                        "SELECT COUNT(*) FROM events WHERE source_device = ? OR target_scope = ?",
                        (device_uuid, target_scope),
                    ).fetchone()[0]
                    or 0
                )
                cursors_deleted = int(
                    conn.execute(
                        "SELECT COUNT(*) FROM device_cursors WHERE device_uuid = ?",
                        (device_uuid,),
                    ).fetchone()[0]
                    or 0
                )
                devices_deleted = int(
                    conn.execute("SELECT COUNT(*) FROM devices WHERE device_name = ?", (device_name,)).fetchone()[0]
                    or 0
                )
                if device_uuid:
                    conn.execute("DELETE FROM events WHERE source_device = ? OR target_scope = ?", (device_uuid, target_scope))
                    conn.execute("DELETE FROM device_cursors WHERE device_uuid = ?", (device_uuid,))
                else:
                    conn.execute("DELETE FROM events WHERE target_scope = ?", (target_scope,))
                conn.execute("DELETE FROM devices WHERE device_name = ?", (device_name,))
            conn.execute("PRAGMA foreign_keys = ON")
        return {
            "ok": True,
            "mode": mode,
            "events_deleted": events_deleted,
            "device_cursors_deleted": cursors_deleted,
            "devices_deleted": devices_deleted,
            "server": "local-sqlite",
            "server_db": server_db,
        }
    finally:
        conn.close()


def _candidate_local_dbs() -> List[str]:
    explicit = (os.environ.get("LOCAL_DB") or "").strip()
    if explicit:
        return [os.path.abspath(explicit)]
    roots = [
        os.getcwd(),
        _script_dir(),
        os.path.abspath(os.path.join(_script_dir(), "..")),
        os.path.abspath(os.path.join(_script_dir(), "..", "POS")),
    ]
    out = []
    seen = set()
    for root in roots:
        if not os.path.isdir(root):
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d.lower() not in ("build", "dist", "__pycache__")]
            if "warehouse_data.sqlite3" not in filenames:
                continue
            path = os.path.abspath(os.path.join(dirpath, "warehouse_data.sqlite3"))
            if path not in seen:
                seen.add(path)
                out.append(path)
    return out


def _db_device_name(path: str) -> str:
    try:
        conn = sqlite3.connect(path)
        try:
            row = conn.execute("SELECT device_name FROM device_identity WHERE id = 1").fetchone()
            return str(row[0] if row else "").strip().upper()
        finally:
            conn.close()
    except Exception:
        return ""


def _resolve_local_db(device_name: str) -> Optional[str]:
    matches = []
    for path in _candidate_local_dbs():
        if _db_device_name(path) == device_name.upper():
            matches.append(path)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise RuntimeError("multiple local DBs matched %s: %s" % (device_name, matches))
    return None


def _wipe_local_db(device_name: str) -> Dict[str, object]:
    db_path = _resolve_local_db(device_name)
    if not db_path:
        return {"ok": False, "skipped": True, "reason": "local DB not found for %s" % device_name}
    deleted = []
    for suffix in ("", "-wal", "-shm"):
        path = db_path + suffix
        if os.path.exists(path):
            os.remove(path)
            deleted.append(path)
    return {"ok": True, "deleted_files": deleted, "local_db": db_path}


def _run(mode: str) -> int:
    cfg = MODE_CONFIG[mode]
    device_name = str(cfg["device_name"])
    target_scope = str(cfg["target_scope"])
    prompt = "This will fully reset %s." % (device_name or "ALL DEVICES")
    if mode == "all":
        prompt += " It wipes all sync server devices/events/cursors."
    else:
        prompt += " It wipes remote sync history"
        if cfg["wipe_local"]:
            prompt += " and deletes the local POS DB"
        prompt += "."
    if not _confirm(prompt):
        print("Cancelled.")
        return 0

    remote_result: Optional[Dict[str, object]] = None
    local_result: Optional[Dict[str, object]] = None
    errors: List[str] = []

    try:
        if (os.environ.get("SERVER_URL") or "").strip():
            remote_result = _remote_reset(mode, device_name, target_scope)
        else:
            remote_result = _local_server_reset(mode, device_name, target_scope)
    except Exception as ex:
        errors.append("remote/server reset failed: %s" % ex)

    if cfg["wipe_local"]:
        try:
            local_result = _wipe_local_db(device_name)
        except Exception as ex:
            errors.append("local DB wipe failed: %s" % ex)

    if remote_result:
        print("Server reset result:")
        print(json.dumps(remote_result, indent=2, ensure_ascii=False))
    if local_result:
        print("Local wipe result:")
        print(json.dumps(local_result, indent=2, ensure_ascii=False))

    if errors:
        print("")
        for err in errors:
            print("ERROR:", err)
        return 1

    print("")
    print("Reset completed successfully.")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Reset POS/warehouse sync state locally and/or remotely.")
    parser.add_argument("mode", choices=sorted(MODE_CONFIG.keys()))
    args = parser.parse_args(argv)
    return _run(args.mode)


if __name__ == "__main__":
    sys.exit(main())
