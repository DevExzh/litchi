#!/usr/bin/env python3
"""Deterministically compress retained logs/profiles and seal a 0448 bundle."""

from __future__ import annotations

import gzip
import hashlib
import io
import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
CHANGE = 448
COMPRESSIBLE_PROFILE_NAMES = frozenset({
    "perf-stat.txt", "perf-report.txt", "perf-script.txt",
    "perf.stat.txt", "perf.report.txt", "perf.script.txt",
})


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def sha_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def sha(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def compressible(path: Path) -> bool:
    return path.name.endswith(".log") or path.name == "perf.data" or path.name in COMPRESSIBLE_PROFILE_NAMES


def reject_running_receipts() -> None:
    running: list[str] = []
    for path in sorted(ROOT.rglob("*.json"), key=relative):
        name = path.name.lower()
        if not ("receipt" in name or name.startswith("capture-state") or path.parent.name == "checks"):
            continue
        try:
            value = load_json(path)
        except (OSError, UnicodeError, json.JSONDecodeError):
            continue
        if isinstance(value, dict) and value.get("status") == "running":
            running.append(relative(path))
    if running:
        raise RuntimeError("refusing to seal while receipts are running: " + ", ".join(running))


def deterministic_gzip(raw: bytes) -> bytes:
    output = io.BytesIO()
    with gzip.GzipFile(filename="", mode="wb", fileobj=output, compresslevel=9, mtime=0) as stream:
        stream.write(raw)
    stored = output.getvalue()
    if gzip.decompress(stored) != raw:
        raise RuntimeError("gzip round trip changed retained bytes")
    return stored


def load_records() -> dict[str, dict[str, Any]]:
    path = ROOT / "compression.json"
    if not path.is_file():
        return {}
    value = load_json(path)
    if not isinstance(value, dict):
        raise RuntimeError("compression.json must contain an object")
    return value


def compress_logs() -> dict[str, dict[str, Any]]:
    records = load_records()
    candidates = sorted(
        (path for path in ROOT.rglob("*") if path.is_file() and compressible(path)),
        key=relative,
    )
    for source in candidates:
        target = source.with_name(source.name + ".gz")
        if target.exists():
            raise RuntimeError(f"compressed destination already exists: {relative(target)}")
    for source in candidates:
        raw = source.read_bytes()
        stored = deterministic_gzip(raw)
        target = source.with_name(source.name + ".gz")
        target.write_bytes(stored)
        records[relative(target)] = {
            "original_path": relative(source),
            "original_sha256": sha_bytes(raw),
            "original_bytes": len(raw),
            "stored_sha256": sha_bytes(stored),
            "stored_bytes": len(stored),
        }
        source.unlink()
    (ROOT / "compression.json").write_text(
        json.dumps(records, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    return records


def write_inventory() -> int:
    inventory = ROOT / "SHA256SUMS"
    members = sorted(
        path for path in ROOT.rglob("*")
        if path.is_file() and path.name != "SHA256SUMS"
    )
    inventory.write_text(
        "".join(f"{sha(path)}  {relative(path)}\n" for path in members),
        encoding="utf-8",
    )
    return len(members)


def seal_bundle() -> dict[str, Any]:
    reject_running_receipts()
    compression = compress_logs()
    members = write_inventory()
    return {
        "status": "pass",
        "change": CHANGE,
        "compression_records": len(compression),
        "inventory_members": members,
        "inventory": "SHA256SUMS",
        "scope": "deterministic gzip for logs/perf artifacts; all remaining bundle files are covered byte-for-byte",
    }


def main() -> int:
    try:
        print(json.dumps(seal_bundle(), indent=2, sort_keys=True))
    except (OSError, TypeError, ValueError, RuntimeError, json.JSONDecodeError) as error:
        print(f"SEAL INVALID: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
