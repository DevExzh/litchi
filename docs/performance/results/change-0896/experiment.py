#!/usr/bin/env python3
"""0896 protocol adapter around the retained strict 0894 workload harness.

This module deliberately does not duplicate the child validator.  It loads the
0894 harness under an isolated module name, redirects the harness packet and
schema globals, and supplies a new plan loader for the QName candidate.  The
workload runner therefore retains the exact 0893 guard vectors, report oracle,
allocation checks, command construction, and alternating AB/BA schedule.
"""

from __future__ import annotations

import datetime as _datetime
import hashlib
import importlib.util
import json
from pathlib import Path
import stat
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
BASE_PACKET = PACKET.parent / "change-0894"
BASE_EXPERIMENT = BASE_PACKET / "experiment.py"
_SPEC = importlib.util.spec_from_file_location("litchi_selected_metadata_experiment_0896_base", BASE_EXPERIMENT)
if _SPEC is None or _SPEC.loader is None:
    raise ImportError(f"cannot load retained 0894 experiment helper: {BASE_EXPERIMENT}")
_BASE = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = _BASE
_SPEC.loader.exec_module(_BASE)


SCHEMA = "litchi.xlsx-qname-expansion-experiment-0896.v1"
COMMAND_SCHEMA = "litchi.xlsx-qname-expansion-child-command-0896.v1"
RECEIPT_SCHEMA = "litchi.xlsx-qname-expansion-child-receipt-0896.v1"
EXPERIMENT_SCHEMA = "litchi.xlsx-qname-expansion-experiment-receipt-0896.v1"
VERIFICATION_SCHEMA = "litchi.xlsx-qname-expansion-verification-0896.v1"
REDUCTION_SCHEMA = "litchi.xlsx-qname-expansion-reduction-0896.v1"
PROFILE = _BASE.PROFILE
CPU_AFFINITY = _BASE.CPU_AFFINITY
CHILD_ENV = dict(_BASE.CHILD_ENV)
CASES = tuple(_BASE.CASES)
PRIMARY_CASE = _BASE.PRIMARY_CASE
NATIVE_BLOCKS = _BASE.NATIVE_BLOCKS
NATIVE_WARMUP = _BASE.NATIVE_WARMUP
NATIVE_SAMPLES = _BASE.NATIVE_SAMPLES
ALLOCATOR_BLOCKS = _BASE.ALLOCATOR_BLOCKS
ALLOCATOR_WARMUP = _BASE.ALLOCATOR_WARMUP
ALLOCATOR_SAMPLES = _BASE.ALLOCATOR_SAMPLES
TOTAL_NATIVE_CHILDREN = _BASE.TOTAL_NATIVE_CHILDREN
TOTAL_ALLOCATOR_CHILDREN = _BASE.TOTAL_ALLOCATOR_CHILDREN
TOTAL_CHILDREN = _BASE.TOTAL_CHILDREN
TOTAL_SAMPLES = TOTAL_NATIVE_CHILDREN * NATIVE_SAMPLES + TOTAL_ALLOCATOR_CHILDREN * ALLOCATOR_SAMPLES
BOOTSTRAP_SEED = 896
BOOTSTRAP_DRAWS = _BASE.BOOTSTRAP_DRAWS
BOOTSTRAP_CI_INDICES = _BASE.BOOTSTRAP_CI_INDICES
PAIR_RATIO_REVIEW_THRESHOLD = _BASE.PAIR_RATIO_REVIEW_THRESHOLD
P50_SPREAD_REVIEW_THRESHOLD = _BASE.P50_SPREAD_REVIEW_THRESHOLD
GUARD_CONTRACT = _BASE.GUARD_CONTRACT
BASELINE_COMMIT = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
RESTORED_HEAD = "3ff493b638b7aa5860d97858dae8bd82bac0f2d5"
HYPOTHESIS = "mce-expand-borrowed-validated-qname-0896"
REUSED_HELPER_PATHS = {
    "experiment-0894.py": BASE_PACKET / "experiment.py",
    "reduce-0894.py": BASE_PACKET / "reduce.py",
    "reduce-0887.py": BASE_PACKET.parent / "change-0887" / "reduce.py",
    "verify-0894.py": BASE_PACKET / "verify.py",
}

# Repoint all retained functions that read module globals.  In particular,
# _BASE.run calls its own global load_plan and _BASE._binary, while the local
# loader below supplies the 0896 schema and evidence contract.
_BASE.PACKET = PACKET
_BASE.SCHEMA = SCHEMA
_BASE.COMMAND_SCHEMA = COMMAND_SCHEMA
_BASE.RECEIPT_SCHEMA = RECEIPT_SCHEMA
_BASE.EXPERIMENT_SCHEMA = EXPERIMENT_SCHEMA
_BASE.BASELINE_QUALIFICATION = BASE_PACKET.parent / "change-0887" / "baseline-qualification"
_BASE.GUARD_REFERENCE = BASE_PACKET.parent / "change-0893" / "qualification-after-cleanup" / "report.json"


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _read(path: Path) -> Any:
    with Path(path).open(encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with Path(path).open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    _require(path.is_file() and not path.is_symlink(), f"expected regular non-symlink file: {path}")
    path = path.resolve()
    info = path.stat()
    _require(stat.S_ISREG(info.st_mode), f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": _sha(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def _same_descriptor(left: Any, right: Any, label: str) -> None:
    _require(isinstance(left, dict) and isinstance(right, dict), f"{label} descriptor is missing")
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        _require(left.get(key) == right.get(key), f"{label} {key} mismatch")


def _regular_path(raw: Any, base: Path, label: str, *, directory: bool = False) -> Path:
    _require(isinstance(raw, (str, Path)), f"{label} path is missing")
    path = Path(raw)
    if not path.is_absolute():
        path = base / path
    _require(not path.is_symlink(), f"{label} must not be a symlink: {path}")
    _require(path.is_dir() if directory else path.is_file(), f"{label} is missing: {path}")
    return path.resolve()


def _entry(raw: Any) -> dict[str, Any]:
    if isinstance(raw, dict) and isinstance(raw.get("identity"), dict):
        value = dict(raw["identity"])
        value.update({key: item for key, item in raw.items() if key != "identity"})
        return value
    _require(isinstance(raw, dict), "descriptor entry must be an object")
    return raw


def _descriptor_entry(raw: Any, base: Path, label: str) -> dict[str, Any]:
    item = _entry(raw)
    path = _regular_path(item.get("path"), base, label)
    actual = descriptor(path)
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in item:
            _require(item[key] == actual[key], f"{label} {key} mismatch")
    result = dict(actual)
    for key, value in item.items():
        if key not in result and key != "identity":
            result[key] = value
    return result


def _nested_descriptor(raw: Any, label: str) -> dict[str, Any]:
    """Validate a descriptor retained inside a binary/build receipt."""
    _require(isinstance(raw, dict), f"{label} descriptor is missing")
    if isinstance(raw.get("identity"), dict):
        item = dict(raw["identity"])
        item.update({key: value for key, value in raw.items() if key != "identity"})
    else:
        item = raw
    path = _regular_path(item.get("path"), PACKET, label)
    actual = descriptor(path)
    _same_descriptor(item, actual, label)
    return actual


def _binding_map(plan: dict[str, Any]) -> dict[str, Any]:
    bindings = plan.get("bindings")
    _require(isinstance(bindings, dict) and bindings, "plan evidence bindings are missing")
    return bindings


def _validate_binding_map(plan: dict[str, Any]) -> None:
    bindings = _binding_map(plan)
    required = {
        "source_manifest", "source_identity", "baseline_build", "candidate_build",
        "baseline_quality", "candidate_quality", "candidate_freeze", "guard_reference",
    }
    _require(required <= set(bindings), "plan bindings omit required source/build/quality evidence")
    for name, raw in bindings.items():
        _require(isinstance(raw, dict), f"binding {name} must be an object")
        _require(isinstance(raw.get("kind"), str) and isinstance(raw.get("role"), str), f"binding {name} role metadata is missing")
        actual = _descriptor_entry(raw, PACKET, f"binding {name}")
        _same_descriptor(raw, actual, f"binding {name}")
        if Path(actual["path"]).suffix.lower() == ".json":
            value = _read(Path(actual["path"]))
            _require(isinstance(value, dict), f"binding {name} JSON root is not an object")
            if name.endswith("build") or name.endswith("quality") or name.endswith("review") or name in {"candidate_freeze", "source_manifest", "source_identity"}:
                _require(value.get("status") in {"pass", "approved", "approved_for_measurement", "reviewed", "frozen"}, f"binding {name} is not an approved receipt")


def _fixture_paths() -> dict[str, Path]:
    return _BASE._fixture_paths()


def _validate_fixture_bindings(plan: dict[str, Any]) -> None:
    fixtures = plan.get("fixtures")
    expected = _fixture_paths()
    _require(isinstance(fixtures, dict) and set(fixtures) == set(expected), "fixture binding set is incomplete")
    for name, path in expected.items():
        _same_descriptor(fixtures[name], descriptor(path), f"fixture {name}")


def _validate_json_crosslinks(plan: dict[str, Any]) -> None:
    bindings = _binding_map(plan)
    source = _read(Path(bindings["source_identity"]["path"]))
    baseline = _read(Path(bindings["baseline_build"]["path"]))
    candidate = _read(Path(bindings["candidate_build"]["path"]))
    for label, value in (("source", source), ("baseline build", baseline), ("candidate build", candidate)):
        _require(isinstance(value, dict), f"{label} receipt is not an object")
        for key in ("baseline_commit", "source_revision"):
            if key in value:
                _require(value[key] == BASELINE_COMMIT, f"{label} {key} does not bind restored production")
    candidate_source = _read(Path(bindings["candidate_freeze"]["path"]))
    quality = _read(Path(bindings["candidate_quality"]["path"]))
    _require(plan.get("candidate", {}).get("hypothesis") == HYPOTHESIS, "candidate hypothesis is not the frozen QName route")
    candidate_paths = plan.get("candidate", {}).get("paths")
    _require(isinstance(candidate_paths, list) and "crates/litchi-ooxml-common/src/mce/stream.rs" in candidate_paths, "candidate source inventory omits MCE stream")
    _require(plan.get("candidate", {}).get("production_adoption") is False, "candidate production adoption must remain false before review")
    _validate_build_configuration(candidate, "candidate build")
    _validate_candidate_source_overlay(candidate_source, plan.get("candidate", {}).get("overlay"))
    candidate_source_identity = (bindings["candidate_freeze"]["path"], bindings["candidate_freeze"]["sha256"])
    _require(_source_link(candidate) == candidate_source_identity, "candidate build source link differs from candidate source receipt")
    _require(_source_link(quality) == candidate_source_identity, "candidate quality source link differs from candidate source receipt")
    candidate_binaries = candidate.get("binaries") if isinstance(candidate.get("binaries"), dict) else candidate
    for mode in ("native", "allocator"):
        raw = candidate_binaries.get(mode)
        _require(raw is not None, f"candidate build omits {mode} binary")
        receipt_binary = raw.get("identity") if isinstance(raw, dict) and isinstance(raw.get("identity"), dict) else raw
        planned = plan.get("binaries", {}).get("after", {}).get(mode)
        _require(isinstance(planned, dict), f"plan omits candidate/{mode} binary")
        _same_descriptor(receipt_binary, planned, f"candidate build {mode} binary")


def _source_link(value: dict[str, Any]) -> tuple[str, str] | None:
    """Return the first explicit source receipt identity in a receipt."""
    for key in ("source", "candidate_source", "source_receipt", "source_identity"):
        raw = value.get(key)
        if isinstance(raw, dict):
            item = raw.get("identity") if isinstance(raw.get("identity"), dict) else raw
            if isinstance(item, dict) and isinstance(item.get("path"), str) and isinstance(item.get("sha256"), str):
                return (item["path"], item["sha256"])
        elif isinstance(raw, str):
            path = Path(raw)
            if path.is_file() and not path.is_symlink():
                actual = descriptor(path)
                return (actual["path"], actual["sha256"])
    return None


def _validate_build_configuration(build: dict[str, Any], label: str) -> None:
    raw = build.get("flags") or build.get("build_flags") or build.get("configuration")
    if not isinstance(raw, dict):
        raw = {}
        environment = build.get("environment") or {}
        compiler = str(build.get("compiler", ""))
        raw.update({
            "toolchain": build.get("toolchain") or ("1.95.0" if "1.95.0" in compiler else None),
            "profile": build.get("profile") or ("release" if "--release" in build.get("argv", []) or "CARGO_PROFILE_RELEASE_OPT_LEVEL" in environment else None),
            "opt_level": environment.get("CARGO_PROFILE_RELEASE_OPT_LEVEL"),
            "debug": environment.get("CARGO_PROFILE_RELEASE_DEBUG"),
            "lto": environment.get("CARGO_PROFILE_RELEASE_LTO"),
            "codegen_units": environment.get("CARGO_PROFILE_RELEASE_CODEGEN_UNITS"),
            "jobs": environment.get("CARGO_BUILD_JOBS"),
        })
    def pick(*names: str) -> Any:
        for name in names:
            if name in raw:
                return raw[name]
            if name in build:
                return build[name]
        return None
    observed = {
        "toolchain": pick("rust_toolchain", "toolchain"),
        "profile": pick("profile"),
        "opt_level": pick("opt_level", "optimization_level"),
        "debug": pick("debug", "debuginfo", "debug_info"),
        "lto": pick("lto"),
        "codegen_units": pick("codegen_units", "codegen-units"),
        "jobs": pick("jobs", "build_jobs"),
    }
    _require(str(observed["toolchain"]).startswith("1.95.0"), f"{label} toolchain is not Rust 1.95")
    _require(observed["profile"] == "release", f"{label} profile is not release")
    _require(observed["opt_level"] in (3, "3"), f"{label} opt-level is not 3")
    _require(observed["debug"] in (1, "1", True), f"{label} debug info is not enabled")
    _require(observed["lto"] == "fat", f"{label} LTO is not fat")
    _require(observed["codegen_units"] in (1, "1"), f"{label} codegen-units is not 1")
    _require(observed["jobs"] in (2, "2"), f"{label} build jobs are not 2")


def _validate_candidate_source_overlay(source: dict[str, Any], planned_overlay: Any) -> None:
    overlay = source.get("overlay") or source.get("files")
    _require(isinstance(overlay, dict) and overlay, "candidate source overlay is missing")
    _require(isinstance(planned_overlay, dict), "plan candidate overlay is missing")
    _require(set(overlay) == set(planned_overlay), "candidate overlay inventory changed")
    for relative, raw in overlay.items():
        actual = _nested_descriptor(raw, f"candidate overlay {relative}")
        planned = planned_overlay[relative]
        _same_descriptor(planned, actual, f"candidate overlay {relative}")


def _validate_binary(raw: Any, label: str, expected_role: str, expected_name: str, expected_features: list[str]) -> dict[str, Any]:
    item = _descriptor_entry(raw, PACKET, label)
    _require(item.get("executable") is True, f"{label} is not executable")
    _require(item.get("role") == expected_role, f"{label} role metadata mismatch")
    _require(item.get("name") == expected_name, f"{label} binary name mismatch")
    _require(item.get("features") == expected_features, f"{label} feature metadata mismatch")
    _require(isinstance(item.get("build_receipt"), dict), f"{label} build receipt binding is missing")
    _require(isinstance(item.get("depfile"), dict), f"{label} depfile binding is missing")
    _require(isinstance(item.get("elf_build_id"), str) and item["elf_build_id"], f"{label} ELF build ID is missing")
    _require(isinstance(item.get("evidence_bindings"), dict), f"{label} evidence binding map is missing")
    _nested_descriptor(item["build_receipt"], f"{label} build receipt")
    _nested_descriptor(item["depfile"], f"{label} depfile")
    evidence = item["evidence_bindings"]
    _require(all(isinstance(key, str) and isinstance(value, (str, dict)) for key, value in evidence.items()), f"{label} evidence bindings are malformed")
    return item


def _validate_plan(plan: dict[str, Any]) -> None:
    _require(plan.get("schema") == SCHEMA, f"unexpected plan schema: {plan.get('schema')!r}")
    _require(plan.get("status") == "frozen", "0896 plan is not frozen")
    _require(plan.get("base_commit") == BASELINE_COMMIT, "plan baseline commit is not restored production")
    _require(plan.get("evaluation_head") == RESTORED_HEAD, "plan evaluation head is not the retained 0895 production state")
    _require(plan.get("profile") == PROFILE and str(plan.get("cpu_affinity")) == CPU_AFFINITY, "execution profile/CPU is not frozen")
    _require(plan.get("cache") == "warm" and plan.get("cases") == list(CASES) and plan.get("xlsx_shape") == "medium", "workload scope changed")
    _require(plan.get("native_blocks") == NATIVE_BLOCKS and plan.get("native_warmup") == NATIVE_WARMUP and plan.get("native_samples") == NATIVE_SAMPLES, "native protocol mismatch")
    _require(plan.get("allocator_blocks") == ALLOCATOR_BLOCKS and plan.get("allocator_warmup") == ALLOCATOR_WARMUP and plan.get("allocator_samples") == ALLOCATOR_SAMPLES, "allocator protocol mismatch")
    _require(plan.get("total_children") == TOTAL_CHILDREN and plan.get("total_samples") == TOTAL_SAMPLES, "population mismatch")
    _require(plan.get("bootstrap_seed") == BOOTSTRAP_SEED and plan.get("bootstrap_draws") == BOOTSTRAP_DRAWS and plan.get("bootstrap_ci_indices") == list(BOOTSTRAP_CI_INDICES), "bootstrap protocol mismatch")
    _require(plan.get("fresh_child") is True and plan.get("fail_fast") is True and plan.get("guard_contract") == GUARD_CONTRACT, "guard/fail-fast contract mismatch")
    _require(plan.get("build_configuration") == _BASE.BUILD_CONFIGURATION, "build configuration mismatch")
    _validate_binding_map(plan)
    _validate_json_crosslinks(plan)
    _validate_fixture_bindings(plan)
    binaries = plan.get("binaries")
    _require(isinstance(binaries, dict) and set(binaries) == {"before", "after"}, "plan binary roles are incomplete")
    for side in ("before", "after"):
        values = binaries[side]
        _require(isinstance(values, dict) and set(values) == {"native", "allocator"}, f"{side} binary lanes are incomplete")
        _validate_binary(values["native"], f"{side}/native", "native", "litchi-perf-baseline", [])
        _validate_binary(values["allocator"], f"{side}/allocator", "allocator", "litchi-perf-baseline-alloc", ["allocator-metrics"])
    helpers = plan.get("helper_hashes")
    _require(isinstance(helpers, dict), "packet helper hashes are missing")
    for name in ("experiment.py", "freeze.py", "reduce.py", "verify.py"):
        _require(name in helpers, f"helper hash missing: {name}")
        _same_descriptor(helpers[name], descriptor(PACKET / name), f"helper {name}")
    reused = plan.get("reused_helper_hashes")
    _require(isinstance(reused, dict) and set(reused) == set(REUSED_HELPER_PATHS), "reused helper hash set is incomplete")
    for name, path in REUSED_HELPER_PATHS.items():
        _same_descriptor(reused[name], descriptor(path), f"reused helper {name}")
    gate = plan.get("primary_gate")
    _require(gate == {"median_ratio_at_most": 0.98, "ci95_upper_below": 1.0}, "primary p50 gate changed")
    _require(plan.get("review_flags") == {"native_p50_spread_above": P50_SPREAD_REVIEW_THRESHOLD, "pair_ratio_above": PAIR_RATIO_REVIEW_THRESHOLD, "rss_causal_claim": False}, "review flag thresholds changed")
    _require(plan.get("rows") == _schedule(), "plan schedule differs from fixed AB/BA schedule")


def load_plan(path: Path) -> dict[str, Any]:
    path = Path(path).resolve()
    _require(path.is_file() and not path.is_symlink(), f"plan is missing: {path}")
    plan = _read(path)
    _require(isinstance(plan, dict), "plan must be an object")
    _validate_plan(plan)
    # The retained implementation resolves binary identity through its own
    # helper.  Calling it only after the new plan contract has passed keeps
    # the strict report path intact without accepting an 0894 plan silently.
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            _BASE._binary(plan, role, mode)
    _BASE._load_expected_corpora()
    _BASE._guard_reference()
    return plan


def _binary(plan: dict[str, Any], role: str, mode: str) -> dict[str, Any]:
    return _BASE._binary(plan, role, mode)


def _load_expected_corpora() -> dict[str, dict[str, Any]]:
    return _BASE._load_expected_corpora()


def _schedule() -> list[dict[str, Any]]:
    return _BASE._schedule()


def validate_schedule(rows: list[dict[str, Any]]) -> None:
    _BASE.validate_schedule(rows)


def _command(binary: str, case: str, warmup: int, samples: int, report: Path, producer: Path, peak: Path) -> list[str]:
    return _BASE._command(binary, case, warmup, samples, report, producer, peak)


def _validate_child(report: dict[str, Any], producer: dict[str, Any] | None, case: str, mode: str, warmup: int, samples: int, expected: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    return _BASE._validate_child(report, producer, case, mode, warmup, samples, expected, binary)


def _read_peak(path: Path) -> int:
    return _BASE._read_peak(path)


def _parse_time(value: str) -> _datetime.datetime:
    return _BASE._parse_time(value)


def _artifacts(folder: Path) -> dict[str, dict[str, Any]]:
    return _BASE._artifacts(folder)


def run(plan_path: Path, output_root: Path, *, popen: Any = None) -> dict[str, Any]:
    """Run the retained strict harness; root owns the only real invocation."""
    if popen is None:
        return _BASE.run(Path(plan_path).resolve(), Path(output_root).resolve())
    return _BASE.run(Path(plan_path).resolve(), Path(output_root).resolve(), popen=popen)


# Install the adapter only after its definition exists.  The retained runner
# resolves ``load_plan`` through its module globals, so this assignment is the
# explicit schema boundary that prevents a silent fallback to 0894 validation.
_BASE.load_plan = load_plan


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(run(args.plan, args.output_root), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
