# -*- coding: utf-8 -*-
"""Upload the prepared POS update package to the sync-server Git repository."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
from pathlib import Path


SYNC_REPO_URL = "https://github.com/Hosny34/Hosny-sync-server.git"


def _repo_dir() -> Path:
    return Path(__file__).parent.parent


def _run(args, cwd=None) -> None:
    print("+ " + " ".join(str(arg) for arg in args), flush=True)
    subprocess.check_call([str(arg) for arg in args], cwd=str(cwd) if cwd else None)


def _load_manifest(repo: Path) -> dict:
    manifest_path = repo / "sync_server" / "Hosny-sync-server" / "updates" / "pos" / "latest.json"
    if not manifest_path.is_file():
        raise RuntimeError("latest.json was not found: %s" % manifest_path)
    with manifest_path.open("r", encoding="utf-8") as f:
        manifest = json.load(f)
    package = manifest.get("package_file")
    if not package:
        raise RuntimeError("latest.json does not contain package_file")
    package_path = manifest_path.parent / package
    if not package_path.is_file():
        raise RuntimeError("package file was not found: %s" % package_path)
    return {
        "manifest": manifest,
        "manifest_path": manifest_path,
        "package_path": package_path,
        "package": package,
        "version": manifest.get("version") or "unknown",
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo-url", default=SYNC_REPO_URL)
    args = parser.parse_args(argv)

    repo = _repo_dir()
    info = _load_manifest(repo)
    tmp_root = Path(tempfile.mkdtemp(prefix="hosny-sync-deploy-"))
    try:
        _run(["git", "clone", args.repo_url, tmp_root])
        dest = tmp_root / "updates" / "pos"
        dest.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(info["manifest_path"]), str(dest / "latest.json"))
        shutil.copy2(str(info["package_path"]), str(dest / info["package"]))

        _run(["git", "add", "updates/pos/latest.json", "updates/pos/%s" % info["package"]], cwd=tmp_root)
        status = subprocess.check_output(["git", "status", "--short"], cwd=str(tmp_root))
        if not status.strip():
            print("No upload needed. Sync server already has POS update %s." % info["version"])
            return 0

        _run(["git", "commit", "-m", "Publish POS update %s" % info["version"]], cwd=tmp_root)
        _run(["git", "push", "origin", "main"], cwd=tmp_root)
        print("Uploaded POS update %s successfully." % info["version"])
        return 0
    finally:
        shutil.rmtree(str(tmp_root), ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
