#!/usr/bin/env python3
"""Replay the sealed 0435 bundle from outside the bundle directory.

Copied-bundle probes invoke the verifier with portable and inventory flags;
``check.py`` is intentionally not a wrapper around replay.  A terminal replay
log and receipt are written only after the verifier process exits, including a
failed exit.  They are outside the preceding seal and therefore require a
later seal if they are to become inventory members.
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
CHANGE = 435
LIFECYCLE_STAGES = ("precleanup", "aftercleanup", "final")
PROFILE_MARKER_NAME = "workload-verify.json"
PROFILE_ROOTS = frozenset({"profiles"})

# These files are the immutable replay inputs and drivers.  A replay run is
# rejected if one is absent, so a historical pass cannot silently bind to a
# different current driver set.  Retain an explicit copy/version before
# replaying historical evidence under a new tag.
DRIVERS = (
    "replay.py",
    "capture.py",
    "pilot.py",
    "check.py",
    "summary.py",
    "profile-summary.py",
    "verify.py",
    "seal.py",
    "protocol.json",
    "frozen-inputs.json",
    "planned-checks.json",
    "oracle/protocol.json",
    "oracle/verify-report.py",
    "profile.py",
    "verify-report.py",
    "candidate-compare-strict.py",
    "compare-strict.py",
    "preparatory-summary.py",
    "oracle-probes.py",
    "save-binaries.py",
    "freeze.py",
    "build-descriptors.py",
    "cleanup.py",
    "lifecycle.py",
    "portable-probes.py",
)


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def is_known_profile_marker(path: Path) -> bool:
    parts = relative(path).split("/")
    return (
        len(parts) >= 3
        and parts[0] in PROFILE_ROOTS
        and parts[-1] == PROFILE_MARKER_NAME
    )


def reject_running_receipts() -> None:
    running: list[str] = []
    for path in sorted(ROOT.rglob("*.json"), key=relative):
        if path.name in {"compression.json", "expected-checks.json"}:
            continue
        if is_known_profile_marker(path):
            if path.read_bytes() == b"VALID\n":
                continue
            raise ValueError(f"known profile marker is not exactly VALID: {relative(path)}")
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
            running.append(relative(path))
    if running:
        raise RuntimeError(
            "refusing replay while receipts are running: " + ", ".join(running)
        )


def driver_snapshot() -> dict[str, str]:
    result: dict[str, str] = {}
    for name in DRIVERS:
        path = ROOT / name
        if not path.is_file():
            raise FileNotFoundError(f"replay driver/input is missing: {name}")
        result[name] = digest(path)
    return result


def driver_snapshot_after() -> dict[str, str | None]:
    result: dict[str, str | None] = {}
    for name in DRIVERS:
        path = ROOT / name
        result[name] = digest(path) if path.is_file() else None
    return result


def inventory_snapshot(path: Path) -> dict[str, str]:
    """Read and hash every member named by the sealed inventory."""

    rows: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if "  " not in line:
            raise ValueError("sealed inventory contains a malformed line")
        expected, name = line.split("  ", 1)
        if (
            len(expected) != 64
            or any(char not in "0123456789abcdefABCDEF" for char in expected)
            or not name
            or name in rows
            or name == "SHA256SUMS"
            or Path(name).is_absolute()
        ):
            raise ValueError("sealed inventory contains a duplicate or unsafe member")
        member = (ROOT / name).resolve()
        if not member.is_relative_to(ROOT.resolve()):
            raise ValueError("sealed inventory member escapes the bundle")
        if not member.is_file() or digest(member) != expected.lower():
            raise ValueError(f"sealed inventory member is stale or missing: {name}")
        rows[name] = expected.lower()
    actual = {
        relative(candidate)
        for candidate in ROOT.rglob("*")
        if candidate.is_file() and candidate.name != "SHA256SUMS"
    }
    if actual != set(rows):
        raise ValueError("sealed inventory does not cover the bundle exactly")
    return rows


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tag", required=True)
    parser.add_argument(
        "--stage",
        choices=LIFECYCLE_STAGES,
        default=None,
        help="optional portable evidence lifecycle gate passed to verify.py",
    )
    args = parser.parse_args(argv)
    if re.fullmatch(r"[a-z0-9_-]+", args.tag) is None:
        raise ValueError("--tag must contain only lowercase letters, digits, _ or -")

    reject_running_receipts()
    receipt = ROOT / "checks" / f"{args.tag}.json"
    log = ROOT / "checks" / f"{args.tag}.log"
    compressed_log = log.with_suffix(log.suffix + ".gz")
    if receipt.exists() or log.exists() or compressed_log.exists():
        if receipt.is_file():
            try:
                previous = json.loads(receipt.read_text(encoding="utf-8"))
            except (OSError, UnicodeError, json.JSONDecodeError):
                previous = None
            if isinstance(previous, dict) and previous.get("status") == "pass":
                raise FileExistsError(
                    f"historical successful replay {args.tag!r} is immutable; "
                    "retain its exact drivers or use a new tag"
                )
        raise FileExistsError(f"replay receipt or log already exists for {args.tag}")

    inventory = ROOT / "SHA256SUMS"
    if not inventory.is_file():
        raise FileNotFoundError("seal the complete 0435 bundle before replay")
    drivers = driver_snapshot()
    inventory_before = digest(inventory)
    inventory_members_before = inventory_snapshot(inventory)
    command = [
        sys.executable,
        "-B",
        str(ROOT / "portable-probes.py"),
    ]
    if args.stage is not None:
        command.extend(("--stage", args.stage))
    effective_stage = args.stage or "final"
    started = now()
    launch_error: str | None = None
    with tempfile.TemporaryDirectory(prefix="litchi-0435-replay-") as temporary:
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
                launch_error = f"replay could not launch verifier: {error}"
                stream.write((launch_error + "\n").encode("utf-8", errors="replace"))
                exit_code = 127
        raw = temporary_log.read_bytes()

    drivers_after = driver_snapshot_after()
    drivers_unchanged = drivers == drivers_after
    inventory_unchanged = inventory_before == digest(inventory)
    try:
        inventory_members_after = inventory_snapshot(inventory)
    except (OSError, UnicodeError, ValueError):
        inventory_members_after = {}
    inventory_members_unchanged = inventory_members_before == inventory_members_after
    inventory_rewritten = not inventory_unchanged
    passed = (
        exit_code == 0
        and drivers_unchanged
        and inventory_unchanged
        and inventory_members_unchanged
    )

    # These writes happen only after verify.py has returned.  A failed verifier
    # therefore leaves an auditable terminal receipt rather than disappearing
    # as an unrecorded subprocess error.
    log.write_bytes(raw)
    record: dict[str, object] = {
        "change": CHANGE,
        "status": "pass" if passed else "failed",
        "command": command,
        "cwd": str(ROOT.parent),
        "portable_flag": "--portable-check",
        "require_inventory": True,
        "stage": effective_stage,
        "exit_code": exit_code,
        "started_utc": started,
        "finished_utc": now(),
        "driver_hashes": drivers,
        "driver_hashes_after": drivers_after,
        "drivers_unchanged": drivers_unchanged,
        "replayed_inventory_sha256": inventory_before,
        "inventory_unchanged": inventory_unchanged,
        "inventory_rewritten": inventory_rewritten,
        "inventory_members_unchanged": inventory_members_unchanged,
        "log": {
            "path": str(log.relative_to(ROOT)),
            "bytes": len(raw),
            "sha256": hashlib.sha256(raw).hexdigest(),
        },
        "scope": "Terminal copied-bundle portable replay with eight mutation probes; receipt/log are outside the preceding seal and require resealing.",
    }
    if launch_error is not None:
        record["launch_error"] = launch_error
    receipt.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": record["status"], "exit_code": exit_code}, sort_keys=True))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
