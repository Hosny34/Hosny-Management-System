# -*- coding: utf-8 -*-
"""Verify that a Windows executable has the expected PE machine type."""

from __future__ import annotations

import sys
from pathlib import Path


MACHINES = {
    "x86": 0x014C,
    "x64": 0x8664,
}


def main(argv=None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) != 2 or args[1] not in MACHINES:
        print("usage: verify_pe_arch.py EXE_PATH x86|x64")
        return 2

    path = Path(args[0])
    expected_name = args[1]
    expected = MACHINES[expected_name]
    with path.open("rb") as f:
        header = f.read(64)
        if len(header) < 64 or header[:2] != b"MZ":
            print("not a Windows executable")
            return 1
        pe_offset = int.from_bytes(header[0x3C:0x40], "little")
        f.seek(pe_offset)
        pe_header = f.read(6)
    if len(pe_header) < 6 or pe_header[:4] != b"PE\0\0":
        print("invalid PE header")
        return 1
    machine = int.from_bytes(pe_header[4:6], "little")
    print("PE machine: 0x%04x" % machine)
    if machine != expected:
        print("Expected %s machine: 0x%04x" % (expected_name, expected))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
