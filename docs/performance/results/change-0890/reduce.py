#!/usr/bin/env python3
"""Reduce retained 0890 children with the reviewed 0887 paired reducer."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys


PACKET = Path(__file__).resolve().parent
# Importing the local experiment wrapper first makes the reviewed reducer's
# normal ``import experiment`` resolve to the 0890 managed-copy adapter.
import experiment  # noqa: E402,F401

_SOURCE = PACKET.parent / "change-0887" / "reduce.py"
_SPEC = importlib.util.spec_from_file_location("litchi_mce_reduce_0887", _SOURCE)
if _SPEC is None or _SPEC.loader is None:
    raise ImportError(f"cannot load reviewed reducer: {_SOURCE}")
_IMPL = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = _IMPL
_SPEC.loader.exec_module(_IMPL)
_IMPL.SCHEMA = "litchi.xlsx-mce-expanded-name-cache-reduction-0890.v1"
for _name in dir(_IMPL):
    if _name.startswith("__"):
        continue
    globals()[_name] = getattr(_IMPL, _name)


if __name__ == "__main__":
    raise SystemExit(_IMPL.main())
