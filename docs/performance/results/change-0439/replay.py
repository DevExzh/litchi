#!/usr/bin/env python3
"""Replay 0439 portable mutation probes without mutating the source bundle."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from typing import Any

ROOT = Path(__file__).resolve().parent
CHANGE = 439


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def contract_drivers() -> list[str]:
    path = ROOT / "lifecycle-contract.json"
    if not path.is_file():
        return ["capture.py", "profile.py", "lifecycle.py", "seal.py", "replay.py", "portable-probes.py", "cleanup.py"]
    value = load(path)
    drivers = value.get("drivers")
    if not isinstance(drivers, list) or any(not isinstance(item, str) for item in drivers):
        raise ValueError("lifecycle contract drivers must be a string list")
    return drivers


def inventory(path: Path) -> dict[str, str]:
    rows: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if "  " not in line:
            raise ValueError("malformed SHA256SUMS row")
        expected, name = line.split("  ", 1)
        if len(expected) != 64 or name in rows or Path(name).is_absolute() or ".." in Path(name).parts:
            raise ValueError("unsafe inventory row")
        member = (ROOT / name).resolve()
        if not member.is_file() or sha(member) != expected.lower():
            raise ValueError(f"stale inventory member: {name}")
        rows[name] = expected.lower()
    actual = {
        rel(path) for path in ROOT.rglob("*")
        if path.is_file() and path.name != "SHA256SUMS"
    }
    if actual != set(rows):
        raise ValueError("SHA256SUMS does not cover the bundle exactly")
    return rows


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tag", required=True)
    parser.add_argument("--stage", choices=("precleanup", "aftercleanup", "final"), default="final")
    args = parser.parse_args(argv)
    if not args.tag or any(ch not in "abcdefghijklmnopqrstuvwxyz0123456789_-" for ch in args.tag):
        parser.error("--tag must contain lowercase letters, digits, _ or -")
    try:
        inventory_path = ROOT / "SHA256SUMS"
        if not inventory_path.is_file():
            raise ValueError("seal the bundle before portable replay")
        drivers = contract_drivers()
        before_drivers = {name: sha(ROOT / name) for name in drivers if (ROOT / name).is_file()}
        before_inventory = inventory(inventory_path)
        command = [sys.executable, "-B", str(ROOT / "portable-probes.py"), "--stage", args.stage]
        started = now()
        result = subprocess.run(command, cwd=ROOT.parent, capture_output=True, text=True, check=False)
        after_drivers = {name: sha(ROOT / name) for name in drivers if (ROOT / name).is_file()}
        after_inventory = inventory(inventory_path)
        passed = (
            result.returncode == 0
            and before_drivers == after_drivers
            and before_inventory == after_inventory
        )
        row = {
            "schema": "litchi-0439-replay-receipt-v1",
            "change": CHANGE,
            "status": "pass" if passed else "failed",
            "tag": args.tag,
            "stage": args.stage,
            "command": command,
            "exit_code": result.returncode,
            "started_utc": started,
            "finished_utc": now(),
            "driver_hashes": before_drivers,
            "driver_hashes_after": after_drivers,
            "drivers_unchanged": before_drivers == after_drivers,
            "inventory_unchanged": before_inventory == after_inventory,
            "output": result.stdout + result.stderr,
            "probe_count": json.loads(result.stdout).get("count", 0) if result.returncode == 0 else 0,
            "scope": "portable mutation probes run only in independent temporary copies",
        }
        checks = ROOT / "checks"
        checks.mkdir(exist_ok=True)
        receipt = checks / f"replay-{args.tag}.json"
        log = checks / f"replay-{args.tag}.log"
        if receipt.exists() or log.exists():
            raise FileExistsError(f"replay tag already exists: {args.tag}")
        log.write_text(row["output"], encoding="utf-8")
        row["log"] = {"path": rel(log), "bytes": log.stat().st_size, "sha256": sha(log)}
        receipt.write_text(json.dumps(row, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": row["status"], "exit_code": result.returncode}, sort_keys=True))
        return 0 if passed else 1
    except (OSError, TypeError, ValueError, json.JSONDecodeError, subprocess.SubprocessError) as error:
        print(f"REPLAY INVALID: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
