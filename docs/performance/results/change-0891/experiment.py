#!/usr/bin/env python3
"""Run the reviewed 0887 paired workload through 0891 managed copies.

The workload driver, semantic report checks, fixed four-case corpus, 72-child
AB/BA schedule, and fresh-child behavior are imported from the reviewed 0887
driver. This adapter changes only packet identity, binary custody, and the
bootstrap seed. It never builds or mutates source files.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
SOURCE = PACKET.parent / "change-0887" / "experiment.py"
SPEC = importlib.util.spec_from_file_location("litchi_x14ac_experiment_0887", SOURCE)
if SPEC is None or SPEC.loader is None:
    raise ImportError(f"cannot load reviewed experiment driver: {SOURCE}")
IMPL = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = IMPL
SPEC.loader.exec_module(IMPL)

IMPL.PACKET = PACKET
IMPL.BASELINE_PACKET = PACKET.parent / "change-0885"
IMPL.BASELINE_QUALIFICATION = PACKET.parent / "change-0887" / "baseline-qualification"
IMPL.ARCHIVE_CASES = {case: IMPL.BASELINE_QUALIFICATION / case for case in IMPL.CASES}
IMPL.SCHEMA = "litchi.xlsx-x14ac-experiment-0891.v1"
IMPL.COMMAND_SCHEMA = "litchi.xlsx-x14ac-child-command-0891.v1"
IMPL.RECEIPT_SCHEMA = "litchi.xlsx-x14ac-child-receipt-0891.v1"
IMPL.EXPERIMENT_SCHEMA = "litchi.xlsx-x14ac-experiment-receipt-0891.v1"
IMPL.BOOTSTRAP_SEED = 891
ORIGINAL_VALIDATE_PLAN = IMPL.validate_plan


def managed_item(raw: Any) -> dict[str, Any]:
    if isinstance(raw, dict) and isinstance(raw.get("copy"), dict):
        return raw["copy"]
    if isinstance(raw, dict):
        return raw
    raise ValueError(f"invalid managed executable descriptor: {raw!r}")


def declared_binary(item: dict[str, Any], label: str) -> dict[str, Any]:
    path = Path(str(item.get("path", "")))
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


def managed_binary(plan: dict[str, Any], role: str, mode: str) -> dict[str, Any]:
    managed = plan.get("managed_binaries")
    if not isinstance(managed, dict):
        raise ValueError("plan has no managed executable descriptors")
    side = managed.get(role)
    if not isinstance(side, dict):
        raise ValueError(f"plan has no managed {role} binaries")
    item = managed_item(side.get(mode))
    path = Path(str(item.get("path", "")))
    if not path.is_absolute():
        path = PACKET / path
    if plan.get("retained_only"):
        return declared_binary(item, f"managed/{role}/{mode}")
    if not path.is_file():
        raise ValueError(f"managed/{role}/{mode} executable is missing: {path}")
    return IMPL._resolve_descriptor(PACKET, item, f"managed/{role}/{mode}")


def validate_managed(plan: dict[str, Any]) -> None:
    source = plan.get("binaries")
    managed = plan.get("managed_binaries")
    if not isinstance(source, dict) or not isinstance(managed, dict):
        raise ValueError("plan must contain source and managed binaries")
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            source_item = IMPL._entry(source.get(role, {}).get(mode))
            managed_item_value = managed_item(managed.get(role, {}).get(mode))
            if source_item.get("sha256") != managed_item_value.get("sha256") or source_item.get("bytes") != managed_item_value.get("bytes"):
                raise ValueError(f"managed {role}/{mode} differs from source binary")
            if managed_item_value.get("executable") is not True:
                raise ValueError(f"managed {role}/{mode} is not executable")
            managed_binary(plan, role, mode)


def validate_plan(plan: dict[str, Any]) -> None:
    ORIGINAL_VALIDATE_PLAN(plan)
    validate_managed(plan)
    if plan.get("bootstrap_seed") != 891:
        raise ValueError("0891 bootstrap seed is not 891")


def load_plan(path: Path) -> dict[str, Any]:
    plan = IMPL._read(path)
    if not isinstance(plan, dict):
        raise ValueError("plan must be an object")
    managed = plan.get("managed_binaries")
    if not isinstance(managed, dict):
        raise ValueError("plan must contain managed executable copies")
    paths = []
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            path = Path(str(managed_item(managed[role][mode])["path"]))
            paths.append(path if path.is_absolute() else PACKET / path)
    present = [path.exists() for path in paths]
    if any(present) and not all(present):
        raise ValueError("managed executable custody is partial")
    plan["retained_only"] = not any(present)
    validate_plan(plan)
    IMPL._load_expected_corpora()
    return plan


ORIGINAL_RUN = IMPL.run


def run(plan_path: Path, output_root: Path, *, popen: Any = None) -> dict[str, Any]:
    plan = IMPL._read(plan_path)
    if not isinstance(plan, dict):
        raise ValueError("plan must be an object")
    plan["retained_only"] = False
    validate_managed(plan)
    if popen is None:
        return ORIGINAL_RUN(plan_path, output_root)
    return ORIGINAL_RUN(plan_path, output_root, popen=popen)


IMPL._binary = managed_binary
IMPL.validate_plan = validate_plan
IMPL.load_plan = load_plan
IMPL.run = run
for name in dir(IMPL):
    if name.startswith("__") or name in {"_binary", "validate_plan", "load_plan", "run"}:
        continue
    globals()[name] = getattr(IMPL, name)
globals().update({"_binary": managed_binary, "validate_plan": validate_plan, "load_plan": load_plan, "run": run, "PACKET": PACKET})


if __name__ == "__main__":
    raise SystemExit(IMPL.main())
