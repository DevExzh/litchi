#!/usr/bin/env python3
"""Offline verifier for the retained 0917 real-DOC packet.

Replay reads only the frozen manifest, retained small reports/logs and the
external raw artifacts named by their hashes.  It never starts a probe or
rebuilds a binary.  Re-running reduction from the child reports makes a
tampered report, receipt, timestamp, process metric, or schedule binding fail
closed.  Use ``--packet`` when checking a copied packet in another checkout.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys


class ReplayError(RuntimeError):
    pass


def load_capture(path: Path):
    script = path / "capture.py"
    spec = importlib.util.spec_from_file_location("litchi_perf_0917_capture", script)
    if spec is None or spec.loader is None:
        raise ReplayError(f"cannot load capture helper: {script}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def read_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReplayError(f"invalid JSON {path}: {error}") from error


def replay(packet: Path) -> dict[str, object]:
    packet = packet.resolve()
    capture = load_capture(packet)
    # The helper's constants are derived from its own location.  This lets a
    # copied packet replay after checkout relocation while retaining the
    # external artifact paths frozen in each descriptor.
    manifest = capture.manifest()
    protocol = capture.protocol()
    native_schedule = capture.native_schedule(protocol)
    allocator_schedule = capture.allocation_schedule(protocol)
    native = capture.reduction_for_lane(manifest, "native")
    allocator = capture.reduction_for_lane(manifest, "allocator")
    reduced_path = packet / "reduced.json"
    retained = read_json(reduced_path)
    if not isinstance(retained, dict):
        raise ReplayError("reduced.json is not an object")
    if retained.get("schema") != "litchi.performance.0917.reduction.v1":
        raise ReplayError("reduced schema mismatch")
    if retained.get("protocol") != manifest["protocol"]:
        raise ReplayError("reduced protocol provenance mismatch")
    retained_manifest = retained.get("manifest")
    if not isinstance(retained_manifest, dict):
        raise ReplayError("reduced manifest provenance missing")
    if retained_manifest != capture.manifest_descriptor():
        raise ReplayError("reduced manifest descriptor differs from frozen manifest")
    capture.verify_descriptor(retained_manifest, "reduced manifest")
    if retained.get("native") != native:
        raise ReplayError("native reduction differs from deterministic replay")
    if retained.get("allocator") != allocator:
        raise ReplayError("allocator reduction differs from deterministic replay")
    # Ensure every expected child was consumed, and no forged extra receipt can
    # hide outside the frozen schedule.  Extra files are retained for failed
    # attempts, but they must be visibly marked as such by the packet owner.
    log_dir = packet / "logs"
    expected = {str(x["label"]) for x in native_schedule + allocator_schedule}
    actual = {p.stem for p in log_dir.glob("*.json") if p.is_file()}
    if actual != expected:
        missing = sorted(expected - actual)
        extra = sorted(actual - expected)
        raise ReplayError(f"child receipt set differs; missing={missing}, extra={extra}")
    packet_files = [p for p in packet.rglob("*") if p.is_file() and not p.is_symlink()]
    large = [str(p) for p in packet_files if p.stat().st_size >= capture.MAX_PACKET_ARTIFACT
             and p.name not in {"capture.py", "replay.py"}]
    if large:
        raise ReplayError(f"packet contains evidence at or above 1 MiB: {large}")
    return {"status": "verified", "native_children": len(native_schedule),
            "allocator_children": len(allocator_schedule),
            "packet": str(packet), "reduction": str(reduced_path)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("replay", nargs="?", choices=("replay",), default="replay")
    parser.add_argument("--packet", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        print(json.dumps(replay(args.packet), sort_keys=True))
    except (ReplayError, OSError, KeyError, TypeError, ValueError, AttributeError) as error:
        print(f"0917 replay failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
