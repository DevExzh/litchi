#!/usr/bin/env python3
"""Reduce 0891 whole-child pairs with the reviewed 0887 reducer."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys


PACKET = Path(__file__).resolve().parent
import experiment  # noqa: E402,F401  # local managed-copy adapter

SOURCE = PACKET.parent / "change-0887" / "reduce.py"
SPEC = importlib.util.spec_from_file_location("litchi_x14ac_reduce_0887", SOURCE)
if SPEC is None or SPEC.loader is None:
    raise ImportError(f"cannot load reviewed reducer: {SOURCE}")
IMPL = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = IMPL
SPEC.loader.exec_module(IMPL)
IMPL.SCHEMA = "litchi.xlsx-x14ac-reduction-0891.v1"

for name in dir(IMPL):
    if not name.startswith("__"):
        globals()[name] = getattr(IMPL, name)


if __name__ == "__main__":
    raise SystemExit(IMPL.main())
