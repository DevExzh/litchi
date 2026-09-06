#!/usr/bin/env python3
"""Replay the sealed 0433 bundle and retain an external portable receipt.

The bound verifier performs its mutation checks in temporary copies.  This
driver launches that verifier from outside the bundle, keeps ``SHA256SUMS``
immutable, and writes the terminal log/receipt only after the verifier exits.
The new receipt and log are intentionally outside the prior seal and require a
fresh seal if they are to become part of a subsequent inventory.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile


ROOT = Path(__file__).resolve().parent
DRIVERS = (
    "replay.py",
    "capture.py",
    "check.py",
    "derive.py",
    "verify.py",
    "verify-report.py",
    "seal.py",
    "protocol.json",
    "planned-checks.json",
    "profile/capture.py",
    "profile/report.py",
)


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def reject_running_receipts() -> None:
    running: list[str] = []
    for path in sorted(ROOT.rglob("*.json")):
        name = path.name.lower()
        if not (
            "receipt" in name
            or name.startswith("capture-state")
            or path.parent.name == "checks"
        ):
            continue
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError):
            continue
        if isinstance(value, dict) and value.get("status") == "running":
            running.append(path.relative_to(ROOT).as_posix())
    if running:
        raise RuntimeError(
            "refusing replay while receipts are running: " + ", ".join(running)
        )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tag", required=True)
    parser.add_argument("--stage", choices=("precleanup", "aftercleanup", "final"), default="final")
    args = parser.parse_args()
    if re.fullmatch(r"[a-z0-9_-]+", args.tag) is None:
        raise ValueError("--tag must contain only lowercase letters, digits, _ or -")

    reject_running_receipts()
    receipt = ROOT / "checks" / f"{args.tag}.json"
    log = ROOT / "checks" / f"{args.tag}.log"
    if receipt.exists() or log.exists() or log.with_suffix(".log.gz").exists():
        raise FileExistsError(f"replay receipt or log already exists for {args.tag}")
    inventory = ROOT / "SHA256SUMS"
    if not inventory.is_file():
        raise FileNotFoundError("seal the complete 0433 bundle before replay")

    drivers = {name: digest(ROOT / name) for name in DRIVERS}
    inventory_before = digest(inventory)
    command = [
        sys.executable,
        "-B",
        str(ROOT / "verify.py"),
        "--portable-check",
        "--require-inventory",
        "--stage", args.stage,
    ]
    started = now()
    with tempfile.TemporaryDirectory(prefix="litchi-0433-replay-") as temporary:
        temporary_log = Path(temporary) / "replay.log"
        with temporary_log.open("xb") as stream:
            try:
                result = subprocess.run(
                    command,
                    cwd=ROOT.parent,
                    stdout=stream,
                    stderr=subprocess.STDOUT,
                )
                exit_code = result.returncode
            except OSError as error:
                stream.write(f"replay could not launch verifier: {error}\n".encode())
                exit_code = 127
        raw = temporary_log.read_bytes()

    drivers_unchanged = drivers == {name: digest(ROOT / name) for name in DRIVERS}
    inventory_unchanged = inventory_before == digest(inventory)
    passed = exit_code == 0 and drivers_unchanged and inventory_unchanged
    log.write_bytes(raw)
    record = {
        "change": 433,
        "status": "pass" if passed else "failed",
        "command": command,
        "cwd": str(ROOT.parent),
        "portable_flag": "--portable-check",
        "require_inventory": True,
        "stage": args.stage,
        "exit_code": exit_code,
        "started_utc": started,
        "finished_utc": now(),
        "driver_hashes": drivers,
        "drivers_unchanged": drivers_unchanged,
        "replayed_inventory_sha256": inventory_before,
        "inventory_unchanged": inventory_unchanged,
        "inventory_rewritten": False,
        "log": {
            "path": str(log.relative_to(ROOT)),
            "bytes": len(raw),
            "sha256": hashlib.sha256(raw).hexdigest(),
        },
        "scope": "Terminal external portable replay receipt; the new receipt/log are outside the prior seal and require resealing.",
    }
    receipt.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": record["status"], "exit_code": exit_code}))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
