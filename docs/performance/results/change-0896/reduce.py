#!/usr/bin/env python3
"""Reduce 0896 paired children through the retained 0894 reducer."""

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


experiment = _load("litchi_qname_experiment_0896_reduce_experiment", PACKET / "experiment.py")
sys.modules["experiment"] = experiment
_BASE_REDUCER = _load("litchi_qname_experiment_0896_base_reducer", PACKET.parent / "change-0887" / "reduce.py")
_BASE_REDUCER.experiment = experiment
_BASE_REDUCER.SCHEMA = experiment.REDUCTION_SCHEMA
_BASE_REDUCER.BOOTSTRAP_DRAWS = experiment.BOOTSTRAP_DRAWS
_BASE_REDUCER.BOOTSTRAP_LOW_INDEX = experiment.BOOTSTRAP_CI_INDICES[0]
_BASE_REDUCER.BOOTSTRAP_HIGH_INDEX = experiment.BOOTSTRAP_CI_INDICES[1]


def reduce(plan_path: Path, experiment_root: Path, output: Path | None = None) -> dict[str, Any]:
    return _BASE_REDUCER.reduce(Path(plan_path).resolve(), Path(experiment_root).resolve(), None if output is None else Path(output).resolve())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--experiment-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(reduce(args.plan, args.experiment_root, args.output), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
