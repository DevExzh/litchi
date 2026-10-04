#!/usr/bin/env python3
"""Offline 0898 custody and semantic replay through the retained verifier."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent


def _load(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load helper: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


experiment = _load("litchi_mce_codec_emitted_name_output_span_experiment_0898_verify_experiment", PACKET / "experiment.py")
_BASE_VERIFIER = _load("litchi_mce_codec_emitted_name_output_span_experiment_0898_base_verifier", PACKET.parent / "change-0894" / "verify.py")
_BASE_VERIFIER.experiment = experiment
_BASE_VERIFIER.SCHEMA = experiment.VERIFICATION_SCHEMA


def verify(plan_path: Path, experiment_root: Path, output: Path | None = None) -> dict[str, Any]:
    return _BASE_VERIFIER.verify(Path(plan_path).resolve(), Path(experiment_root).resolve(), None if output is None else Path(output).resolve())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--experiment-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(verify(args.plan, args.experiment_root, args.output), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
