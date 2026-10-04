#!/usr/bin/env python3
"""Create the immutable 0896 qualification plan without running anything.

The bindings file is supplied by the source/build/quality owner.  It may use
either a top-level ``bindings`` object or top-level binding names; unknown
review receipts are retained in the plan so later custody checks can bind them
without changing this adapter.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
BASE_PLAN_PATH = PACKET.parent / "change-0894" / "plan.json"
_SPEC = importlib.util.spec_from_file_location("litchi_qname_experiment_0896_freeze", PACKET / "experiment.py")
if _SPEC is None or _SPEC.loader is None:
    raise ImportError("cannot load local 0896 experiment helper")
experiment = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = experiment
_SPEC.loader.exec_module(experiment)


def _read(path: Path) -> Any:
    with Path(path).open(encoding="utf-8") as handle:
        return json.load(handle)


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _load_base() -> tuple[Any, dict[str, Any]]:
    spec = importlib.util.spec_from_file_location("litchi_qname_experiment_0896_base0894", PACKET.parent / "change-0894" / "experiment.py")
    if spec is None or spec.loader is None:
        raise ImportError("cannot load retained 0894 experiment helper")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    plan = module.load_plan(BASE_PLAN_PATH)
    return module, plan


def _resolve(raw: Any, base: Path, label: str) -> Path:
    if isinstance(raw, dict):
        raw = raw.get("path")
    _require(isinstance(raw, str), f"{label} path is missing")
    path = Path(raw)
    if not path.is_absolute():
        path = base / path
    _require(path.is_file() and not path.is_symlink(), f"{label} is missing or symlinked: {path}")
    return path.resolve()


def _normalized(raw: Any, base: Path, label: str, *, kind: str | None = None, role: str | None = None) -> dict[str, Any]:
    source = dict(raw) if isinstance(raw, dict) else {"path": raw}
    if isinstance(source.get("identity"), dict):
        identity = dict(source.pop("identity"))
        identity.update(source)
        source = identity
    path = _resolve(source.get("path"), base, label)
    actual = experiment.descriptor(path)
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in source:
            _require(source[key] == actual[key], f"{label} {key} mismatch")
    result = dict(actual)
    for key, value in source.items():
        if key not in result and key != "identity":
            result[key] = value
    if kind is not None:
        result["kind"] = kind
    if role is not None:
        result["role"] = role
    return result


def _bindings(raw: dict[str, Any], input_path: Path, base_plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    supplied = raw.get("bindings") if isinstance(raw.get("bindings"), dict) else raw
    _require(isinstance(supplied, dict), "bindings JSON must contain an object")
    inherited = base_plan.get("bindings", {})
    names = {
        "source_manifest": ("source_manifest", "source"),
        "source_identity": ("source_identity", "source"),
        "candidate_freeze": ("candidate_freeze", "candidate_source", "candidate-source"),
        "baseline_build": ("baseline_build",),
        "candidate_build": ("candidate_build",),
        "baseline_quality": ("baseline_quality",),
        "candidate_quality": ("candidate_quality", "quality"),
        "guard_reference": ("guard_reference",),
    }
    optional_names = {
        "source_review": ("source_review", "source-review"),
        "candidate_review": ("candidate_review", "candidate-review", "report_review", "report-review"),
        "methodology_review": ("methodology_review", "methodology-review"),
    }
    result: dict[str, dict[str, Any]] = {}
    for name, alternatives in names.items():
        entry = None
        for candidate in alternatives:
            if candidate in supplied:
                entry = supplied[candidate]
                break
        if entry is None and name in inherited:
            entry = inherited[name]
        _require(entry is not None, f"bindings JSON omits {name}")
        inherited_item = inherited.get(name, {}) if isinstance(inherited, dict) else {}
        source = dict(entry) if isinstance(entry, dict) else {"path": entry}
        if isinstance(inherited_item, dict):
            for key in ("kind", "role", "review_scope", "binary_role", "build_binding", "candidate_binding", "source_binding", "archive_binding"):
                if key not in source and key in inherited_item:
                    source[key] = inherited_item[key]
        source.setdefault("kind", "evidence")
        source.setdefault("role", "source" if name.startswith("source") or name == "guard_reference" else ("candidate" if "candidate" in name or name == "methodology_review" else "baseline"))
        result[name] = _normalized(source, input_path.parent, f"binding {name}", kind=source["kind"], role=source["role"])

    for name, alternatives in optional_names.items():
        entry = next((supplied[candidate] for candidate in alternatives if candidate in supplied), None)
        if entry is None:
            continue
        source = dict(entry) if isinstance(entry, dict) else {"path": entry}
        source.setdefault("kind", "review")
        source.setdefault("role", "review")
        result[name] = _normalized(source, input_path.parent, f"binding {name}", kind=source["kind"], role=source["role"])

    # Retain additional document/review receipts verbatim as descriptor-bound
    # entries.  They are deliberately not interpreted by this harness.
    for name, entry in supplied.items():
        if name in result or name in {"schema", "status", "base_commit", "evaluation_head", "candidate", "binaries", "fixtures"}:
            continue
        if isinstance(entry, (str, dict)) and (isinstance(entry, str) or "path" in entry or "identity" in entry):
            source = dict(entry) if isinstance(entry, dict) else {"path": entry}
            source.setdefault("kind", "additional_evidence")
            source.setdefault("role", "review")
            result[name] = _normalized(source, input_path.parent, f"binding {name}", kind=source["kind"], role=source["role"])
    return result


def _candidate_value(raw: dict[str, Any], bindings: dict[str, dict[str, Any]], key: str, default: Any = None) -> Any:
    candidate = raw.get("candidate")
    if isinstance(candidate, dict) and key in candidate:
        return candidate[key]
    if key in raw:
        return raw[key]
    if key == "paths":
        receipt = _read(Path(bindings["candidate_freeze"]["path"]))
        for name in ("candidate_paths", "paths", "changed_paths"):
            if isinstance(receipt.get(name), list):
                return receipt[name]
        overlay = receipt.get("overlay") or receipt.get("files")
        if isinstance(overlay, dict):
            return list(overlay)
    return default


def _binary_from_build(build: dict[str, Any], mode: str, base: Path, build_binding: dict[str, Any], label: str) -> dict[str, Any]:
    candidates: list[Any] = []
    binaries = build.get("binaries")
    if isinstance(binaries, dict):
        candidates.append(binaries.get(mode))
    candidates.extend((build.get(mode), build.get(f"{mode}_binary")))
    value = next((item for item in candidates if item is not None), None)
    _require(value is not None, f"{label} binary descriptor is missing")
    item = _normalized(value, base, label)
    expected_role = "native" if mode == "native" else "allocator"
    expected_name = "litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc"
    expected_features = [] if mode == "native" else ["allocator-metrics"]
    _require(item.get("role") == expected_role and item.get("name") == expected_name, f"{label} role/name is not normalized")
    _require(item.get("features") == expected_features, f"{label} feature set is not normalized")
    item.setdefault("build_receipt", build_binding)
    depfiles = build.get("depfiles")
    if isinstance(depfiles, dict) and mode in depfiles:
        item.setdefault("depfile", _normalized(depfiles[mode], base, f"{label} depfile"))
    _require(isinstance(item.get("depfile"), dict), f"{label} depfile binding is missing")
    _require(isinstance(item.get("elf_build_id"), str) and item["elf_build_id"], f"{label} ELF build ID is missing")
    item.setdefault("evidence_bindings", {"build": build_binding["path"], "source": "source_manifest", "quality": "candidate_quality"})
    return item


def _baseline_binaries(base_plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    result = {}
    for mode in ("native", "allocator"):
        item = dict(base_plan["binaries"]["before"][mode])
        item["evidence_bindings"] = dict(item.get("evidence_bindings", {}))
        result[mode] = item
    return result


def freeze(bindings_path: Path, output: Path, base_plan_path: Path = BASE_PLAN_PATH) -> dict[str, Any]:
    bindings_path = Path(bindings_path).resolve()
    output = Path(output).resolve()
    _require(bindings_path.is_file() and not bindings_path.is_symlink(), "bindings JSON is missing")
    _require(not output.exists() and not output.is_symlink(), f"plan output already exists: {output}")
    supplied = _read(bindings_path)
    _require(isinstance(supplied, dict), "bindings JSON root must be an object")
    base_module, base_plan = _load_base()
    _require(Path(base_plan_path).resolve() == BASE_PLAN_PATH.resolve(), "0896 must use retained 0894 baseline plan")
    _require(base_plan.get("base_commit") == experiment.BASELINE_COMMIT, "retained baseline plan commit mismatch")
    plan_bindings = _bindings(supplied, bindings_path, base_plan)
    candidate_build = _read(Path(plan_bindings["candidate_build"]["path"]))
    _require(isinstance(candidate_build, dict) and candidate_build.get("status") == "pass", "candidate build receipt does not pass")
    _require(candidate_build.get("guard_contract") in (None, experiment.GUARD_CONTRACT) or candidate_build.get("timed_harness_guard_contract") == experiment.GUARD_CONTRACT, "candidate build guard contract mismatch")
    candidate_binaries = {
        mode: _binary_from_build(candidate_build, mode, bindings_path.parent, plan_bindings["candidate_build"], f"candidate/{mode}")
        for mode in ("native", "allocator")
    }
    baseline_binaries = _baseline_binaries(base_plan)
    for mode in ("native", "allocator"):
        _require(baseline_binaries[mode]["sha256"] == base_plan["binaries"]["before"][mode]["sha256"], f"baseline {mode} was not inherited byte-for-byte")
        candidate_binaries[mode]["evidence_bindings"] = dict(candidate_binaries[mode]["evidence_bindings"])
    paths = _candidate_value(supplied, plan_bindings, "paths")
    _require(isinstance(paths, list) and "crates/litchi-ooxml-common/src/mce/stream.rs" in paths, "candidate paths must include the MCE stream owner")
    hypothesis = _candidate_value(supplied, plan_bindings, "hypothesis", experiment.HYPOTHESIS)
    _require(hypothesis == experiment.HYPOTHESIS, "candidate hypothesis is not the approved QName route")
    fixtures = {name: experiment.descriptor(path) for name, path in experiment._fixture_paths().items()}
    helper_hashes = {name: experiment.descriptor(PACKET / name) for name in ("experiment.py", "freeze.py", "reduce.py", "verify.py")}
    reused_helper_hashes = {name: experiment.descriptor(path) for name, path in experiment.REUSED_HELPER_PATHS.items()}
    source_binding = plan_bindings["source_identity"]
    source_receipt = _read(Path(source_binding["path"]))
    _require(isinstance(source_receipt, dict) and source_receipt.get("status") == "pass", "source receipt does not pass")
    candidate_source_receipt = _read(Path(plan_bindings["candidate_freeze"]["path"]))
    candidate_overlay = candidate_source_receipt.get("overlay") or candidate_source_receipt.get("files")
    _require(isinstance(candidate_overlay, dict) and candidate_overlay, "candidate source overlay is missing")
    plan = {
        "schema": experiment.SCHEMA,
        "status": "frozen",
        "frozen_utc": experiment._utc_now(),
        "base_commit": experiment.BASELINE_COMMIT,
        "evaluation_head": str(supplied.get("evaluation_head", experiment.RESTORED_HEAD)),
        "hypothesis": hypothesis,
        "profile": experiment.PROFILE,
        "cpu_affinity": experiment.CPU_AFFINITY,
        "cache": "warm",
        "cases": list(experiment.CASES),
        "xlsx_shape": "medium",
        "native_blocks": experiment.NATIVE_BLOCKS,
        "native_warmup": experiment.NATIVE_WARMUP,
        "native_samples": experiment.NATIVE_SAMPLES,
        "allocator_blocks": experiment.ALLOCATOR_BLOCKS,
        "allocator_warmup": experiment.ALLOCATOR_WARMUP,
        "allocator_samples": experiment.ALLOCATOR_SAMPLES,
        "total_children": experiment.TOTAL_CHILDREN,
        "total_samples": experiment.TOTAL_SAMPLES,
        "bootstrap_seed": experiment.BOOTSTRAP_SEED,
        "bootstrap_draws": experiment.BOOTSTRAP_DRAWS,
        "bootstrap_ci_indices": list(experiment.BOOTSTRAP_CI_INDICES),
        "fresh_child": True,
        "fail_fast": True,
        "guard_contract": experiment.GUARD_CONTRACT,
        "build_configuration": base_module.BUILD_CONFIGURATION,
        "source": {"receipt": source_binding, "baseline_commit": experiment.BASELINE_COMMIT, "evaluation_head": str(supplied.get("evaluation_head", experiment.RESTORED_HEAD))},
        "candidate": {"hypothesis": hypothesis, "paths": paths, "production_adoption": False, "source": plan_bindings["candidate_freeze"], "overlay": candidate_overlay},
        "builds": {"before": base_plan["builds"]["before"], "after": plan_bindings["candidate_build"]},
        "bindings": plan_bindings,
        "fixtures": fixtures,
        "binaries": {"before": baseline_binaries, "after": candidate_binaries},
        "helper_hashes": helper_hashes,
        "reused_helper_hashes": reused_helper_hashes,
        "rows": experiment._schedule(),
        "primary_gate": {"median_ratio_at_most": 0.98, "ci95_upper_below": 1.0},
        "review_flags": {"native_p50_spread_above": experiment.P50_SPREAD_REVIEW_THRESHOLD, "pair_ratio_above": experiment.PAIR_RATIO_REVIEW_THRESHOLD, "rss_causal_claim": False},
        "performance_claim": None,
        "production_adoption": False,
    }
    experiment._validate_plan(plan)
    experiment._write_json(output, plan)
    return plan


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bindings", type=Path, required=True)
    parser.add_argument("--output", type=Path, default=PACKET / "plan.json")
    parser.add_argument("--base-plan", type=Path, default=BASE_PLAN_PATH)
    args = parser.parse_args()
    print(json.dumps(freeze(args.bindings, args.output, args.base_plan), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
