# -*- coding: utf-8 -*-
"""Create a POS update zip and publish latest.json for the sync server."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path


EXCLUDE_DIRS = {"dist", "build", "backups", "__pycache__", ".git"}
VERSION_RE = re.compile(r"^\d+(?:\.\d+)+$")

INCLUDE_FILES = {
    "HosnyPOS.py",
    "logging_setup.py",
    "sync_appliers.py",
    "sync_client.py",
    "sync_core.py",
    "sync_periodic.py",
    "sync_ui.py",
    "update_client.py",
    "verify_pe_arch.py",
    "pos_updater.py",
    "pos_version.py",
    "pos_db_backup.py",
    "HosnyPOS.spec",
    "HosnyPOS-win7-x86.spec",
    "HosnyPOS-win7-x64.spec",
    "build_pos.bat",
    "build_pos_win7_x86.bat",
    "build_pos_win7_x64.bat",
    "requirements-win7-x64.txt",
    "requirements-win7-x86.txt",
}


def _iter_files(pos_dir: Path):
    dist_cert = pos_dir / "dist" / "cacert.pem"
    live_cert = pos_dir / "cacert.pem"
    if dist_cert.is_file():
        yield dist_cert, Path("cacert.pem")
    elif live_cert.is_file():
        yield live_cert, Path("cacert.pem")
    for name in sorted(INCLUDE_FILES):
        path = pos_dir / name
        if path.is_file():
            yield path, Path(name)
    fonts_dir = pos_dir / "Fonts"
    if fonts_dir.is_dir():
        for root, dirs, files in os.walk(str(fonts_dir)):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
            root_path = Path(root)
            for name in sorted(files):
                path = root_path / name
                if path.is_file():
                    yield path, path.relative_to(pos_dir)


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    repo_dir = Path(__file__).parent.parent
    parser.add_argument("--pos-dir", default=str(repo_dir / "POS"))
    parser.add_argument("--server-dir", default=str(repo_dir / "sync_server" / "Hosny-sync-server"))
    parser.add_argument("--version", required=True)
    parser.add_argument("--notes", default="")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args(argv)
    if not VERSION_RE.match(str(args.version or "").strip()):
        raise RuntimeError("POS update version must be numeric dotted like 1.2 or 2026.08.02.1, not %r" % args.version)
    if args.verbose:
        print("publisher started", flush=True)

    pos_dir = Path(os.path.abspath(args.pos_dir))
    out_dir = Path(os.path.abspath(args.server_dir)) / "updates" / "pos"
    out_dir.mkdir(parents=True, exist_ok=True)

    package_name = "HosnyPOS-%s.zip" % args.version
    package_path = out_dir / package_name
    files_to_package = list(_iter_files(pos_dir))
    forbidden = [
        str(rel).replace("\\", "/")
        for _, rel in files_to_package
        if "HosnyWarehouse" in str(rel) or str(rel).replace("\\", "/") == "part2_ui.py"
    ]
    if forbidden:
        raise RuntimeError(
            "POS update package must not include warehouse-only files: %s"
            % ", ".join(sorted(forbidden))
        )
    if args.verbose:
        print("packaging %s files into %s" % (len(files_to_package), package_path), flush=True)
    with zipfile.ZipFile(str(package_path), "w", compression=zipfile.ZIP_STORED) as zf:
        for path, rel in files_to_package:
            if args.verbose:
                print("adding %s" % rel, flush=True)
            zf.write(str(path), "POS/" + str(rel).replace("\\", "/"))

    files = sorted("POS/" + str(rel).replace("\\", "/") for _, rel in files_to_package)
    manifest = {
        "app": "POS",
        "version": args.version,
        "package_file": package_name,
        "required_files": ["POS/HosnyPOS.py", "POS/HosnyPOS.spec", "POS/build_pos.bat"],
        "files": files,
        "hashes": {"package_sha256": _sha256(package_path)},
        "notes": args.notes,
        "published_at": datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec="seconds") + "Z",
    }
    with (out_dir / "latest.json").open("w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2, sort_keys=True)
    print("created", package_path)
    print("updated", out_dir / "latest.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
