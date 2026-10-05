#!/usr/bin/env python3
"""Postcapture reduction adapter for the frozen 0917 packet.

The captured runner is immutable.  This adapter binds its exact SHA-256,
applies one asserted in-memory correction to the native reduction loop, and
then invokes the frozen reduction or replay implementation without writing the
runner, protocol, manifest, or raw evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import types
from typing import Any


PACKET = Path(__file__).resolve().parent
FROZEN_CAPTURE_SHA256 = "fde5219539cf538d51b8bf92ec5527fd9959ac0f39eb2ea5412b8ddf8cd7a894"
OLD_LOOP = "for base_item, (_, item) in zip(baseline_rows, rows, strict=True):"
FIXED_LOOP = "for (_, base_item), (_, item) in zip(baseline_rows, rows, strict=True):"


class AnalysisError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AnalysisError(message)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_fixed_capture(packet: Path) -> types.ModuleType:
    capture_path = packet / "capture.py"
    require(capture_path.is_file() and not capture_path.is_symlink(),
            f"missing frozen capture: {capture_path}")
    actual = sha256(capture_path)
    require(actual == FROZEN_CAPTURE_SHA256,
            f"frozen capture SHA changed: {actual}")
    manifest_path = packet / "manifest.json"
    require(manifest_path.is_file(), f"missing frozen manifest: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    identity = manifest.get("capture_script") if isinstance(manifest, dict) else None
    require(isinstance(identity, dict)
            and identity.get("sha256") == FROZEN_CAPTURE_SHA256,
            "manifest is not bound to the expected frozen capture")
    source = capture_path.read_text(encoding="utf-8")
    require(source.count(OLD_LOOP) == 1,
            "frozen reduction loop occurrence count is not exactly one")
    fixed_source = source.replace(OLD_LOOP, FIXED_LOOP, 1)
    require(fixed_source.count(OLD_LOOP) == 0
            and fixed_source.count(FIXED_LOOP) == 1,
            "in-memory reduction correction was not unique")
    module = types.ModuleType("litchi_perf_0917_fixed_capture")
    module.__file__ = str(capture_path)
    module.__package__ = ""
    code = compile(fixed_source, str(capture_path), "exec")
    exec(code, module.__dict__)
    return module


def load_frozen_replay(packet: Path, capture: types.ModuleType) -> types.ModuleType:
    replay_path = packet / "replay.py"
    require(replay_path.is_file() and not replay_path.is_symlink(),
            f"missing frozen replay: {replay_path}")
    spec = importlib.util.spec_from_file_location("litchi_perf_0917_frozen_replay",
                                                  replay_path)
    require(spec is not None and spec.loader is not None,
            "cannot load frozen replay")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The frozen replay algorithm remains unchanged; only its loader is bound
    # to the corrected in-memory capture module.
    module.load_capture = lambda _packet: capture
    return module


def reduce_packet(packet: Path, output: Path | None) -> None:
    capture = load_fixed_capture(packet)
    capture.reduce(argparse.Namespace(output=output))


def replay_packet(packet: Path) -> dict[str, Any]:
    capture = load_fixed_capture(packet)
    replay = load_frozen_replay(packet, capture)
    return replay.replay(packet)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    reduce_parser = sub.add_parser("reduce")
    reduce_parser.add_argument("--output", type=Path)
    sub.add_parser("replay")
    args = parser.parse_args()
    try:
        if args.action == "reduce":
            reduce_packet(PACKET, args.output)
        else:
            print(json.dumps(replay_packet(PACKET), sort_keys=True))
    except (AnalysisError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"0917 analysis failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
