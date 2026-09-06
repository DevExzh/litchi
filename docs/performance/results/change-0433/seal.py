#!/usr/bin/env python3
"""Losslessly compress 0433 logs/profiles and seal the complete bundle inventory.

Sealing is a terminal bundle operation.  It refuses to mutate the bundle when
any capture or command receipt is still running.  The inventory deliberately
does not include ``SHA256SUMS`` itself; adding a later replay receipt therefore
requires a subsequent seal and cannot silently rewrite the existing seal.
"""

from __future__ import annotations

import gzip
import hashlib
import io
import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def json_object(path: Path) -> dict[str, Any] | None:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


def reject_running_receipts() -> None:
    """Reject active receipts before touching any bundle file.

    Capture receipts live below ``before``/``after`` while command receipts
    live below ``checks``.  Capture state files are included because they are
    the enclosing role's completion receipt even when no lane receipt exists.
    """

    running: list[str] = []
    for path in sorted(ROOT.rglob("*.json")):
        name = path.name.lower()
        is_receipt = (
            "receipt" in name
            or name.startswith("capture-state")
            or path.parent.name == "checks"
        )
        if not is_receipt or path.name in {"compression.json", "expected-checks.json"}:
            continue
        value = json_object(path)
        if value is not None and value.get("status") == "running":
            running.append(relative(path))
    if running:
        raise RuntimeError(
            "refusing to seal while receipts are running: " + ", ".join(running)
        )


def compression_inputs() -> list[Path]:
    paths = {path for path in ROOT.rglob("*.log") if path.is_file()}
    for build in ("before", "after"):
        profiles = ROOT / build / "profiles"
        if profiles.is_dir():
            paths.update(path for path in profiles.rglob("*.data") if path.is_file())
            paths.update(path for path in profiles.rglob("perf-report.txt") if path.is_file())
    return sorted(paths, key=relative)


def deterministic_gzip(raw: bytes) -> bytes:
    output = io.BytesIO()
    with gzip.GzipFile(
        filename="", mode="wb", fileobj=output, compresslevel=9, mtime=0
    ) as stream:
        stream.write(raw)
    stored = output.getvalue()
    if gzip.decompress(stored) != raw:
        raise RuntimeError("gzip round trip changed a log or profile")
    return stored


def refresh_compression(inputs: list[Path]) -> dict[str, dict[str, Any]]:
    compression = ROOT / "compression.json"
    records: dict[str, dict[str, Any]] = {}
    if compression.exists():
        value = json.loads(compression.read_text(encoding="utf-8"))
        if not isinstance(value, dict):
            raise ValueError("compression.json must contain an object")
        records = value
    for path in inputs:
        stored_path = path.with_suffix(path.suffix + ".gz")
        if stored_path.exists():
            raise FileExistsError(f"compressed output already exists: {relative(stored_path)}")
    for path in inputs:
        stored_path = path.with_suffix(path.suffix + ".gz")
        raw = path.read_bytes()
        stored = deterministic_gzip(raw)
        stored_path.write_bytes(stored)
        records[relative(stored_path)] = {
            "original_path": relative(path),
            "original_sha256": sha(raw),
            "original_bytes": len(raw),
            "stored_sha256": sha(stored),
            "stored_bytes": len(stored),
        }
        path.unlink()
    compression.write_text(
        json.dumps(records, sort_keys=True, indent=2) + "\n", encoding="utf-8"
    )
    return records


def expected_checks() -> dict[str, str]:
    checks: dict[str, str] = {}
    for path in sorted(ROOT.rglob("*.json"), key=relative):
        if path.name in {"compression.json", "expected-checks.json"}:
            continue
        value = json_object(path)
        if value is None or "source_before" not in value:
            continue
        status = value.get("status")
        if status not in {"pass", "failed"}:
            raise ValueError(f"check receipt has unsupported terminal status: {relative(path)}")
        key = path.relative_to(ROOT).with_suffix("").as_posix()
        checks[key] = status
    return checks


def refresh_inventory() -> int:
    files = sorted(
        (path for path in ROOT.rglob("*") if path.is_file() and path.name != "SHA256SUMS"),
        key=relative,
    )
    (ROOT / "SHA256SUMS").write_text(
        "".join(f"{sha(path.read_bytes())}  {relative(path)}\n" for path in files),
        encoding="utf-8",
    )
    return len(files)


def main() -> int:
    reject_running_receipts()
    inputs = compression_inputs()
    records = refresh_compression(inputs)
    checks = expected_checks()
    (ROOT / "expected-checks.json").write_text(
        json.dumps(checks, sort_keys=True, indent=2) + "\n", encoding="utf-8"
    )
    inventory_files = refresh_inventory()
    print(
        json.dumps(
            {
                "inventory_files": inventory_files,
                "command_receipts": len(checks),
                "compressed_logs_and_profiles": len(records),
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
