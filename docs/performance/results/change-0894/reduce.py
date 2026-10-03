#!/usr/bin/env python3
"""Reduce the retained 0894 children with the reviewed paired reducer.

The reducer is deliberately a small adapter around the 0887 implementation.
The local :mod:`experiment` module supplies the new source/binary identity and
strict 0893 worksheet-guard validation; this file changes only packet schema
and bootstrap seed through the imported implementation.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys


PACKET = Path(__file__).resolve().parent
LOCAL_EXPERIMENT = PACKET / "experiment.py"
experiment_spec = importlib.util.spec_from_file_location("litchi_selected_metadata_experiment_0894", LOCAL_EXPERIMENT)
if experiment_spec is None or experiment_spec.loader is None:
    raise ImportError(f"cannot load local experiment helper: {LOCAL_EXPERIMENT}")
experiment = importlib.util.module_from_spec(experiment_spec)
sys.modules[experiment_spec.name] = experiment
experiment_spec.loader.exec_module(experiment)
sys.modules["experiment"] = experiment

SOURCE = PACKET.parent / "change-0887" / "reduce.py"
reduce_spec = importlib.util.spec_from_file_location("litchi_selected_metadata_reduce_0887", SOURCE)
if reduce_spec is None or reduce_spec.loader is None:
    raise ImportError(f"cannot load reviewed reducer: {SOURCE}")
IMPL = importlib.util.module_from_spec(reduce_spec)
sys.modules[reduce_spec.name] = IMPL
reduce_spec.loader.exec_module(IMPL)

IMPL.SCHEMA = "litchi.xlsx-selected-metadata-reduction-0894.v1"
IMPL.BOOTSTRAP_SEED = experiment.BOOTSTRAP_SEED
IMPL.BOOTSTRAP_DRAWS = experiment.BOOTSTRAP_DRAWS
IMPL.BOOTSTRAP_LOW_INDEX, IMPL.BOOTSTRAP_HIGH_INDEX = experiment.BOOTSTRAP_CI_INDICES

for name in dir(IMPL):
    if not name.startswith("__"):
        globals()[name] = getattr(IMPL, name)


if __name__ == "__main__":
    raise SystemExit(IMPL.main())
