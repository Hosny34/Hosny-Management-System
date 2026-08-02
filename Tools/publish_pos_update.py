# -*- coding: utf-8 -*-
"""Bump POS version and publish a POS update package."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import subprocess
from datetime import datetime, timezone
from pathlib import Path


VERSION_RE = re.compile(r"^\d+(?:\.\d+)+$")


def _repo_dir() -> Path:
    return Path(__file__).parent.parent


def _validate_numeric_version(version: str) -> str:
    version = (version or "").strip()
    if not VERSION_RE.match(version):
        raise RuntimeError("POS version must be numeric dotted like 1.2 or 2026.08.02.1, not %r" % version)
    return version


def _read_version_from_json(path: Path) -> str:
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return str(data.get("version") or "").strip()
    except Exception:
        return ""


def _read_version_from_py(path: Path) -> str:
    try:
        text = path.read_text(encoding="utf-8")
    except Exception:
        return ""
    match = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', text)
    return match.group(1).strip() if match else ""


def _next_numeric_version(repo: Path) -> str:
    candidates = [
        _read_version_from_json(repo / "POS" / "current_version.json"),
        _read_version_from_py(repo / "POS" / "pos_version.py"),
        _read_version_from_json(repo / "sync_server" / "Hosny-sync-server" / "updates" / "pos" / "latest.json"),
    ]
    versions = []
    for value in candidates:
        if VERSION_RE.match(value):
            raw_parts = value.split(".")
            parts = tuple(int(part) for part in raw_parts)
            widths = tuple(len(part) for part in raw_parts)
            versions.append((parts, widths))
    if not versions:
        return "2026.08.02.1"
    latest, widths = max(versions, key=lambda item: item[0])
    latest = list(latest)
    latest[-1] += 1
    return ".".join(str(part).zfill(widths[idx]) for idx, part in enumerate(latest))


def _write_version(repo: Path, version: str) -> None:
    version = _validate_numeric_version(version)
    version_file = repo / "POS" / "pos_version.py"
    version_file.write_text(
        '# -*- coding: utf-8 -*-\n\nAPP_VERSION = "%s"\n' % version,
        encoding="utf-8",
    )
    current_file = repo / "POS" / "current_version.json"
    current_file.write_text(
        json.dumps(
            {
                "version": version,
                "updated_at": datetime.now(timezone.utc).replace(tzinfo=None).isoformat(timespec="seconds") + "Z",
            },
            ensure_ascii=False,
        ) + "\n",
        encoding="utf-8",
    )


def _load_package_tool(repo: Path):
    path = repo / "Tools" / "create_pos_update_package.py"
    spec = importlib.util.spec_from_file_location("create_pos_update_package", str(path))
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load create_pos_update_package.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _build_pos(repo: Path) -> None:
    pos_dir = repo / "POS"
    marker = pos_dir / "prebuilt_update.json"
    if marker.exists():
        marker.unlink()
    env = os.environ.copy()
    env["HOSNY_AUTO_UPDATE"] = "1"
    result = subprocess.call(
        ["cmd.exe", "/c", str(pos_dir / "build_pos.bat")],
        cwd=str(pos_dir),
        env=env,
        shell=False,
    )
    if result != 0:
        raise RuntimeError("build_pos.bat failed with exit code %s" % result)
    dist_exe = pos_dir / "dist" / "HosnyPOS.exe"
    dist_cert = pos_dir / "dist" / "cacert.pem"
    if not dist_exe.is_file():
        raise RuntimeError("dist/HosnyPOS.exe was not created")
    if not dist_cert.is_file():
        print("warning: dist/cacert.pem was not created; package will use POS/cacert.pem if available")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--notes", default="POS update")
    parser.add_argument("--version", default="")
    args = parser.parse_args(argv)

    repo = _repo_dir()
    version = _validate_numeric_version((args.version or "").strip() or _next_numeric_version(repo))
    _write_version(repo, version)
    print("POS version:", version)
    _build_pos(repo)

    tool = _load_package_tool(repo)
    return int(tool.main(["--version", version, "--notes", args.notes]) or 0)


if __name__ == "__main__":
    raise SystemExit(main())
