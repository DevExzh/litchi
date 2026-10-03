#!/usr/bin/env python3
"""Run the 0890 matched MCE experiment through managed executable copies.

The capture implementation is the reviewed 0887 driver with only protocol
identity, packet paths, and binary resolution adapted here.  Reusing that
driver keeps the timed scope, fresh-child behavior, semantic fences, and
counterbalanced AB/BA schedule identical.  The plan carries the retained
source binaries separately from ``managed_binaries``; every timed command
resolves the latter, whose byte identity is checked against its source.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
_SOURCE = PACKET.parent / "change-0887" / "experiment.py"
_SPEC = importlib.util.spec_from_file_location("litchi_mce_experiment_0887", _SOURCE)
if _SPEC is None or _SPEC.loader is None:
    raise ImportError(f"cannot load reviewed capture driver: {_SOURCE}")
_IMPL = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = _IMPL
_SPEC.loader.exec_module(_IMPL)


_IMPL.PACKET = PACKET
_IMPL.BASELINE_PACKET = PACKET.parent / "change-0885"
_IMPL.BASELINE_QUALIFICATION = PACKET.parent / "change-0887" / "baseline-qualification"
_IMPL.ARCHIVE_CASES = {
    case: _IMPL.BASELINE_QUALIFICATION / case for case in _IMPL.CASES
}
_IMPL.SCHEMA = "litchi.xlsx-mce-expanded-name-cache-experiment-0890.v1"
_IMPL.COMMAND_SCHEMA = "litchi.xlsx-mce-expanded-name-cache-child-command-0890.v1"
_IMPL.RECEIPT_SCHEMA = "litchi.xlsx-mce-expanded-name-cache-child-receipt-0890.v1"
_IMPL.EXPERIMENT_SCHEMA = "litchi.xlsx-mce-expanded-name-cache-experiment-receipt-0890.v1"
_IMPL.BOOTSTRAP_SEED = 890
_ORIGINAL_VALIDATE_PLAN = _IMPL.validate_plan


def _managed_item(raw: Any) -> dict[str, Any]:
    if isinstance(raw, dict) and isinstance(raw.get("copy"), dict):
        return raw["copy"]
    if isinstance(raw, dict):
        return raw
    raise ValueError(f"invalid managed executable descriptor: {raw!r}")


def _declared_binary(item: dict[str, Any], label: str) -> dict[str, Any]:
    path = Path(item.get("path", ""))
    if not path.is_absolute():
        path = PACKET / path
    identity = {
        "path": str(path.resolve()),
        "bytes": item.get("bytes"),
        "sha256": item.get("sha256"),
        "mode": item.get("mode"),
        "executable": item.get("executable"),
    }
    if not isinstance(identity["bytes"], int) or not isinstance(identity["sha256"], str):
        raise ValueError(f"{label} descriptor is incomplete")
    if identity["executable"] is not True:
        raise ValueError(f"{label} descriptor is not executable")
    return {"path": identity["path"], "identity": identity, "declared": item}


def _managed_binary(plan: dict[str, Any], role: str, mode: str) -> dict[str, Any]:
    binaries = plan.get("managed_binaries")
    if isinstance(binaries, dict) and not plan.get("retained_only"):
        side = binaries.get(role)
        if not isinstance(side, dict):
            raise ValueError(f"plan has no managed {role} binaries")
        item = _managed_item(side.get(mode))
        path = Path(item.get("path", ""))
        if not path.is_absolute():
            path = PACKET / path
        if path.is_file():
            return _IMPL._resolve_descriptor(PACKET, item, f"managed/{role}/{mode}")
    if isinstance(binaries, dict):
        side = binaries.get(role)
        if not isinstance(side, dict):
            raise ValueError(f"plan has no managed {role} binaries")
        return _declared_binary(_managed_item(side.get(mode)), f"managed/{role}/{mode}")
    source_side = plan.get("binaries", {}).get(role)
    if not isinstance(source_side, dict):
        raise ValueError(f"plan has no retained {role} binaries")
    return _IMPL._resolve_descriptor(PACKET, source_side.get(mode), f"retained/{role}/{mode}")


def _validate_managed(plan: dict[str, Any]) -> None:
    source_binaries = plan.get("binaries")
    managed_binaries = plan.get("managed_binaries")
    if not isinstance(source_binaries, dict) or not isinstance(managed_binaries, dict):
        raise ValueError("plan must contain source and managed binary descriptors")
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            source = _IMPL._entry(source_binaries.get(role, {}).get(mode))
            managed = _managed_item(managed_binaries.get(role, {}).get(mode))
            if source.get("sha256") != managed.get("sha256") or source.get("bytes") != managed.get("bytes"):
                raise ValueError(f"managed {role}/{mode} differs from its source binary")
            if managed.get("executable") is not True:
                raise ValueError(f"managed {role}/{mode} is not executable")
            _managed_binary(plan, role, mode)


def validate_plan(plan: dict[str, Any]) -> None:
    _ORIGINAL_VALIDATE_PLAN(plan)
    if plan.get("retained_only"):
        for role in ("before", "after"):
            for mode in ("native", "allocator"):
                _IMPL._resolve_descriptor(PACKET, plan["binaries"][role][mode], f"retained/{role}/{mode}")
                source = _IMPL._entry(plan["binaries"][role][mode])
                managed = _managed_item(plan["managed_binaries"][role][mode])
                if source.get("sha256") != managed.get("sha256") or source.get("bytes") != managed.get("bytes"):
                    raise ValueError(f"retained managed {role}/{mode} identity differs from source")
                _declared_binary(managed, f"managed/{role}/{mode}")
    else:
        _validate_managed(plan)


def load_plan(path: Path) -> dict[str, Any]:
    plan = _IMPL._read(path)
    if not isinstance(plan, dict):
        raise ValueError("plan must be an object")
    managed = plan.get("managed_binaries")
    if not isinstance(managed, dict):
        raise ValueError("plan must contain managed executable copies")
    managed_paths = [
        Path(_managed_item(managed[role][mode])["path"])
        for role in ("before", "after")
        for mode in ("native", "allocator")
    ]
    present = [path.exists() for path in managed_paths]
    if any(present) and not all(present):
        raise ValueError("managed executable custody is partial")
    plan["retained_only"] = not any(present)
    validate_plan(plan)
    _IMPL._load_expected_corpora()
    return plan


_ORIGINAL_RUN = _IMPL.run


def run(plan_path: Path, output_root: Path, *, popen: Any = None) -> dict[str, Any]:
    plan = _IMPL._read(plan_path)
    if not isinstance(plan, dict):
        raise ValueError("plan must be an object")
    plan["retained_only"] = False
    _validate_managed(plan)
    if popen is None:
        return _ORIGINAL_RUN(plan_path, output_root)
    return _ORIGINAL_RUN(plan_path, output_root, popen=popen)


# The reviewed driver resolves binaries through this module-global function in
# run(), load_plan(), and the semantic child checks.  Replace that one seam,
# then expose the remaining implementation and constants for the reducer.
_IMPL._binary = _managed_binary
_IMPL.validate_plan = validate_plan
_IMPL.load_plan = load_plan
_IMPL.run = run
for _name in dir(_IMPL):
    if _name.startswith("__") or _name in {"_binary", "validate_plan", "load_plan", "run"}:
        continue
    globals()[_name] = getattr(_IMPL, _name)
globals()["_binary"] = _managed_binary
globals()["validate_plan"] = validate_plan
globals()["load_plan"] = load_plan
globals()["run"] = run
globals()["PACKET"] = PACKET


if __name__ == "__main__":
    raise SystemExit(_IMPL.main())
