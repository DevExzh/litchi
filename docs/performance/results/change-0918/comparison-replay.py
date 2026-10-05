#!/usr/bin/env python3
"""Offline replay for the retained 0918 DOC comparison packet."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys


class ReplayError(RuntimeError):
    pass


def require_reconstructed_hierarchy(packet: Path) -> Path:
    """Require the same copied repository hierarchy used by prior packets."""
    try:
        repository = packet.parents[3]
    except IndexError as error:
        raise ReplayError(
            "comparison replay requires docs/performance/results/change-0918 hierarchy"
        ) from error
    expected_packet = (repository / "docs/performance/results/change-0918").resolve()
    if packet != expected_packet:
        raise ReplayError(
            "comparison replay requires packet at "
            "<root>/docs/performance/results/change-0918; "
            "copy the unchanged 0917 validator/probe and fixtures under <root>"
        )
    required = (
        repository / "docs/performance/results/change-0917/capture.py",
        repository / "docs/performance/results/change-0917/probe",
        repository / "test-data/ole/doc/NoHeadFoot.doc",
        repository / "test-data/ole/doc/FloatingPictures.doc",
    )
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise ReplayError(
            "reconstructed replay inputs are missing (original checkout is not a fallback): "
            + ", ".join(missing)
        )
    return repository


def load_runner(packet: Path):
    script = packet / "comparison.py"
    spec = importlib.util.spec_from_file_location("litchi_perf_0918_comparison", script)
    if spec is None or spec.loader is None:
        raise ReplayError(f"cannot load comparison runner: {script}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def replay(packet: Path) -> dict[str, object]:
    packet = packet.resolve()
    require_reconstructed_hierarchy(packet)
    runner = load_runner(packet)
    # Relative packet descriptors resolve below the reconstructed root.  The
    # explicit hierarchy check prevents a copied packet from silently using
    # the original checkout for validator, probe, or fixture inputs.
    result = runner.replay(argparse.Namespace(reduced=None))
    return {"status": "verified", "packet": str(packet), **result}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path,
                        default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        print(json.dumps(replay(args.packet), sort_keys=True))
    except Exception as error:
        print(f"0918 replay failed: {error}", file=sys.stderr)
        return 1
    return 0
if __name__ == "__main__":
    raise SystemExit(main())
