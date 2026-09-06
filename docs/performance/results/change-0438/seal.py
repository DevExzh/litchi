#!/usr/bin/env python3
"""Seal the 0438 ODP bundle with deterministic profile/log compression and hashes.

This is a terminal bundle operation.  It rejects active command/capture
receipts before changing anything, compresses only artifacts with a retained
reader contract, records the exact terminal check statuses, and writes a
complete inventory excluding ``SHA256SUMS`` itself.  Patch files remain
  uncompressed because no 0438 consumer has a patch-gzip contract; they are
still covered byte-for-byte by the inventory.
"""

from __future__ import annotations

import gzip
import hashlib
import io
import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
EXCLUDED_METADATA = frozenset({"compression.json", "expected-checks.json"})
TERMINAL_CHECK_STATUSES = frozenset({"pass", "failed"})
PROFILE_MARKER_NAME = "workload-verify.json"
PROFILE_ROOTS = frozenset({"profiles"})

# Keep these names explicit.  In particular, do not add ``*.patch`` here:
# patch metadata and its original bytes must remain directly readable.
PROFILE_TEXT_NAMES = frozenset(
    {
        "perf-report.txt",
        "perf.report.txt",
        "perf-stat.txt",
        "perf.stat.txt",
        "perf-script.txt",
        "perf.script.txt",
    }
)


def is_compressible_name(path: Path) -> bool:
    """Recognize raw members covered by the retained gzip contract."""

    name = path.name[:-3] if path.name.endswith(".gz") else path.name
    return name.endswith(".log") or name == "perf.data" or name in PROFILE_TEXT_NAMES


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


def is_known_profile_marker(path: Path) -> bool:
    """Recognize only the profile verifier's deliberate non-JSON marker.

    ``profile/capture.py`` stores verifier stdout in a JSON-suffixed marker.
    A marker with any other bytes is malformed and must not be silently
    treated as an unrelated JSON file.
    """

    parts = relative(path).split("/")
    if (
        len(parts) >= 3
        and parts[0] in PROFILE_ROOTS
        and parts[-1] == PROFILE_MARKER_NAME
    ):
        if path.read_bytes() != b"VALID\n":
            raise ValueError(f"known profile marker is not exactly VALID: {relative(path)}")
        return True
    return False


def reject_running_receipts() -> None:
    """Reject active receipts before any compression or inventory mutation."""

    running: list[str] = []
    for path in sorted(ROOT.rglob("*.json"), key=relative):
        if path.name in EXCLUDED_METADATA:
            continue
        if is_known_profile_marker(path):
            continue
        name = path.name.lower()
        is_receipt = (
            "receipt" in name
            or name.startswith("capture-state")
            or path.parent.name == "checks"
        )
        if not is_receipt:
            continue
        value = json_object(path)
        if value is not None and value.get("status") == "running":
            running.append(relative(path))
    if running:
        raise RuntimeError(
            "refusing to seal while receipts are running: " + ", ".join(running)
        )


def compression_inputs() -> list[Path]:
    """Return only log/profile artifacts with a known gzip reader contract."""

    paths: set[Path] = set()
    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        if path.name.endswith(".log"):
            paths.add(path)
        elif path.name == "perf.data" or path.name in PROFILE_TEXT_NAMES:
            paths.add(path)
        # Deliberately omit .patch/.patch.gz: no consumer is authorized to
        # reinterpret patch metadata through a compressed sidecar yet.
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


def load_compression_records(path: Path) -> dict[str, dict[str, Any]]:
    if not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"invalid compression.json: {error}") from error
    if not isinstance(value, dict):
        raise ValueError("compression.json must contain an object")
    return value


def refresh_compression(inputs: list[Path]) -> dict[str, dict[str, Any]]:
    """Compress each new input once and retain original metadata.

    Destination existence is checked for the complete input set before the
    first write.  This prevents a rerun from silently mixing old and new
    compression records.  Each record binds both the original bytes and the
    deterministic gzip bytes; original paths disappear only after their
    corresponding stored bytes have been written and round-trip checked.
    """

    compression = ROOT / "compression.json"
    records = load_compression_records(compression)
# A copied historical gzip sidecar, if present, is retained through the
    # existing compression inventory.  0438 does not require a strict-debt
    # baseline or any other particular pre-existing gzip member.
    for stored_path in sorted(ROOT.rglob("*.gz"), key=relative):
        if not is_compressible_name(stored_path):
            raise ValueError(f"unsupported pre-existing gzip member: {relative(stored_path)}")
        stored_name = relative(stored_path)
        try:
            raw = gzip.decompress(stored_path.read_bytes())
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            raise ValueError(f"invalid pre-existing gzip member: {stored_name}: {error}") from error
        record = {
            "original_path": stored_name[:-3],
            "original_sha256": sha(raw),
            "original_bytes": len(raw),
            "stored_sha256": sha(stored_path.read_bytes()),
            "stored_bytes": stored_path.stat().st_size,
        }
        if stored_name in records and records[stored_name] != record:
            raise ValueError(f"compression record differs for retained gzip: {stored_name}")
        records.setdefault(stored_name, record)
    destinations: list[tuple[Path, Path]] = []
    seen: set[Path] = set()
    for path in inputs:
        stored_path = path.with_suffix(path.suffix + ".gz")
        if stored_path in seen:
            raise ValueError(f"duplicate compression destination: {relative(stored_path)}")
        seen.add(stored_path)
        if stored_path.exists():
            raise FileExistsError(
                f"compressed output already exists: {relative(stored_path)}"
            )
        destinations.append((path, stored_path))

    for path, stored_path in destinations:
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
    """Mirror the verifier's source-before receipt set.

    A profile workload marker is the only known JSON-suffixed non-JSON
    artifact.  All actual receipts carrying ``source_before`` must have a
    terminal status understood by the bundle verifier.
    """

    checks: dict[str, str] = {}
    for path in sorted(ROOT.rglob("*.json"), key=relative):
        if path.name in EXCLUDED_METADATA:
            continue
        if is_known_profile_marker(path):
            continue
        value = json_object(path)
        if value is None:
            name = path.name.lower()
            if (
                path.parent.name == "checks"
                or "receipt" in name
                or name.startswith("capture-state")
            ):
                raise ValueError(f"invalid JSON receipt: {relative(path)}")
            continue
        if "source_before" not in value:
            continue
        status = value.get("status")
        if status not in TERMINAL_CHECK_STATUSES:
            raise ValueError(
                f"check receipt has unsupported terminal status: {relative(path)}"
            )
        key = path.relative_to(ROOT).with_suffix("").as_posix()
        if key in checks:
            raise ValueError(f"duplicate expected check: {key}")
        checks[key] = status
    return checks


def refresh_inventory() -> int:
    """Hash every bundle file except the inventory being written."""

    files = sorted(
        (
            path
            for path in ROOT.rglob("*")
            if path.is_file() and path.name != "SHA256SUMS"
        ),
        key=relative,
    )
    (ROOT / "SHA256SUMS").write_text(
        "".join(f"{sha(path.read_bytes())}  {relative(path)}\n" for path in files),
        encoding="utf-8",
    )
    return len(files)


def main() -> int:
    reject_running_receipts()
    # Validate custody before deleting any raw log.  A malformed receipt
    # therefore cannot leave a half-sealed bundle behind.
    checks = expected_checks()
    inputs = compression_inputs()
    records = refresh_compression(inputs)
    (ROOT / "expected-checks.json").write_text(
        json.dumps(checks, sort_keys=True, indent=2) + "\n", encoding="utf-8"
    )
    inventory_files = refresh_inventory()
    print(
        json.dumps(
            {
                "inventory_files": inventory_files,
                "command_receipts": len(checks),
                "compressed_artifacts": len(records),
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
