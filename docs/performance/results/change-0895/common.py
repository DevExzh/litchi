#!/usr/bin/env python3
"""Small custody helpers shared by the 0895 heaptrack packet.

The packet deliberately reuses the sealed 0894 experiment validator.  This
module only adapts its descriptors and report validator; it never launches a
process.  Every writer in this packet uses create-only JSON output so a
partial capture cannot be silently replaced.
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
PARENT_PACKET = PACKET.parent / "change-0894"
BASE_PLAN_PATH = PARENT_PACKET / "plan.json"

_SPEC = importlib.util.spec_from_file_location(
    "litchi_selected_metadata_experiment_0894_for_0895", PARENT_PACKET / "experiment.py"
)
if _SPEC is None or _SPEC.loader is None:
    raise ImportError(f"cannot load sealed 0894 experiment helper: {PARENT_PACKET / 'experiment.py'}")
experiment = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = experiment
_SPEC.loader.exec_module(experiment)


SCHEMA = "litchi.xlsx-selected-metadata-heaptrack-0895.v1"
CAPTURE_SCHEMA = "litchi.xlsx-selected-metadata-heaptrack-capture-0895.v1"
EXPORT_SCHEMA = "litchi.xlsx-selected-metadata-heaptrack-export-0895.v1"
REDUCTION_SCHEMA = "litchi.xlsx-selected-metadata-heaptrack-reduction-0895.v1"
VERIFY_SCHEMA = "litchi.xlsx-selected-metadata-heaptrack-verification-0895.v1"

CPU_AFFINITY = "12"
PRIMARY_CASE = "xlsx_producer_medium_source_full_cell_scan"
WARMUP = 0
SAMPLES = 20
ROLES = ("baseline", "candidate", "candidate", "baseline")
TOTAL_ROWS = len(ROLES)
TOTAL_SAMPLES = TOTAL_ROWS * SAMPLES

ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
}

# This is the invocation used by the retained 0895 symbol-quality preflight.
# It requests allocation-cost stacks, no peak-cost entries, full symbol names,
# and no filtered caller scope.  A root run may add a version-supported
# ``--merge-backtraces=0`` token at freeze time; the exact argv is then pinned
# in the 0895 plan and replayed byte-for-byte by export/verify.
DEFAULT_PRINT_ARGS = [
    "-f", "{profile}", "-t", "0", "-m", "0", "-a", "1", "-p", "0",
    "-T", "0", "-n", "30", "-s", "10",
    "--flamegraph-cost-type", "allocations",
    "-F", "{stacks}",
    "-H", "{histogram}",
]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"expected regular non-symlink file: {path}")
    path = path.resolve()
    info = path.stat()
    require(stat.S_ISREG(info.st_mode), f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def read_json(path: Path) -> Any:
    with Path(path).open(encoding="utf-8") as handle:
        return json.load(handle)


def write_json(path: Path, value: Any) -> None:
    path = Path(path)
    require(not path.exists() and not path.is_symlink(), f"output already exists or is symlinked: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def same_descriptor(left: Any, right: Any, label: str) -> None:
    require(isinstance(left, dict) and isinstance(right, dict), f"{label} descriptor is missing")
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        require(left.get(key) == right.get(key), f"{label} {key} mismatch")


def regular_path(raw: Any, base: Path, label: str, *, directory: bool = False) -> Path:
    require(isinstance(raw, (str, Path)), f"{label} path is missing")
    path = Path(raw)
    if not path.is_absolute():
        path = base / path
    require(not path.is_symlink(), f"{label} must not be a symlink: {path}")
    require(path.is_dir() if directory else path.is_file(), f"{label} is missing: {path}")
    return path.resolve()


def entry(value: Any) -> dict[str, Any]:
    if isinstance(value, str):
        return {"path": value}
    require(isinstance(value, dict), f"descriptor is not an object: {value!r}")
    if isinstance(value.get("identity"), dict):
        merged = dict(value["identity"])
        merged.update({key: item for key, item in value.items() if key != "identity"})
        return merged
    return value


def load_base_plan(path: Path = BASE_PLAN_PATH) -> dict[str, Any]:
    path = Path(path).resolve()
    require(path == BASE_PLAN_PATH.resolve(), f"0894 base plan must be the retained packet plan: {path}")
    plan = read_json(path)
    require(isinstance(plan, dict), "0894 base plan must be an object")
    experiment.load_plan(path)
    require(plan.get("schema") == "litchi.xlsx-selected-metadata-experiment-0894.v1", "0894 plan schema mismatch")
    require(plan.get("status") == "frozen", "0894 plan is not frozen")
    return plan


def exact_0894_helper_hashes(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    helpers = plan.get("helper_hashes")
    require(isinstance(helpers, dict) and len(helpers) == 7, "0894 plan must expose exactly seven helper hashes")
    expected = {
        "experiment.py", "freeze.py", "quality.py", "quality_source.py",
        "reduce-0887.py", "reduce.py", "verify.py",
    }
    require(set(helpers) == expected, "0894 helper hash set changed")
    for name, item in helpers.items():
        require(isinstance(item, dict), f"0894 helper descriptor is invalid: {name}")
        actual = descriptor(Path(item["path"]))
        same_descriptor(item, actual, f"0894 helper {name}")
    return helpers


def binary(plan: dict[str, Any], role: str) -> dict[str, Any]:
    require(role in ("baseline", "candidate"), f"invalid 0895 role: {role}")
    side = plan["binaries"]["before" if role == "baseline" else "after"]
    raw = entry(side["native"])
    path = regular_path(raw.get("path"), PACKET, f"0894 {role} native binary")
    actual = descriptor(path)
    same_descriptor(raw, actual, f"0894 {role} native binary")
    require(raw.get("role") == "native" and raw.get("features") == [], f"0894 {role} binary lane mismatch")
    return {**raw, "path": str(path), "identity": actual}


def _binding_path(plan: dict[str, Any], name: str) -> Path:
    bindings = plan.get("bindings")
    require(isinstance(bindings, dict) and name in bindings, f"0894 binding is missing: {name}")
    raw = entry(bindings[name])
    path = regular_path(raw.get("path"), PARENT_PACKET, f"0894 binding {name}")
    same_descriptor(raw, descriptor(path), f"0894 binding {name}")
    return path


def evidence_files(plan: dict[str, Any]) -> dict[str, Path]:
    """Return the exact decision, source, build, freeze, corpus and quality files."""
    result = {
        "plan": BASE_PLAN_PATH,
        "bindings": PARENT_PACKET / "bindings.json",
        "candidate_freeze": _binding_path(plan, "candidate_freeze"),
        "baseline_build": _binding_path(plan, "baseline_build"),
        "candidate_build": _binding_path(plan, "candidate_build"),
        "baseline_quality": _binding_path(plan, "baseline_quality"),
        "candidate_quality": _binding_path(plan, "candidate_quality"),
        "source_manifest": _binding_path(plan, "source_manifest"),
        "decision": PARENT_PACKET / "decision.json",
        "candidate_review": PARENT_PACKET / "candidate-review.json",
        "capture_review": PARENT_PACKET / "capture-review.json",
        "final_review": PARENT_PACKET / "final-review.json",
        "corpus_run": PARENT_PACKET / "corpus-v1" / "run.json",
        "corpus_review": PARENT_PACKET / "corpus-v1" / "corpus-review.json",
        "corpus_comparison": PARENT_PACKET / "corpus-v1" / "comparison-v2.json",
        "corpus_cleanup_verification": PARENT_PACKET / "corpus-v1" / "verification-after-cleanup.json",
        "quality_source": PARENT_PACKET / "quality-source.json",
        "methodology_review": PARENT_PACKET / "methodology-review.json",
    }
    for name, path in result.items():
        require(path.is_file() and not path.is_symlink(), f"0894 evidence file is missing: {name}: {path}")
    return {name: path.resolve() for name, path in result.items()}


def evidence_descriptors(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {name: descriptor(path) for name, path in evidence_files(plan).items()}


def assert_rejected_0894(plan: dict[str, Any]) -> None:
    decision = read_json(evidence_files(plan)["decision"])
    require(decision.get("status") == "complete" and decision.get("decision") == "reject", "0894 decision is not a completed rejection")
    final = read_json(evidence_files(plan)["final_review"])
    require(final.get("status") == "pass" and final.get("decision") == "reject", "0894 final review is not a passing rejection")
    require(final.get("performance_claim") is None, "0894 final review contains a performance claim")


def expected_corpora() -> dict[str, dict[str, Any]]:
    return experiment._load_expected_corpora()


def expected_binary_for_experiment(plan: dict[str, Any], role: str) -> dict[str, Any]:
    return experiment._binary(plan, "before" if role == "baseline" else "after", "native")


def render_print_args(template: list[str], *, profile: Path, stacks: Path, histogram: Path) -> list[str]:
    values = {"profile": str(profile), "stacks": str(stacks), "histogram": str(histogram)}
    result = []
    for token in template:
        try:
            result.append(token.format(**values))
        except KeyError as error:
            raise ValueError(f"unknown heaptrack_print template field: {error.args[0]}") from error
    return result


def absolute_executable(raw: str, label: str) -> Path:
    path = Path(raw)
    require(path.is_absolute(), f"{label} must be an absolute path: {path}")
    require(path.is_file() and not path.is_symlink(), f"{label} must be a regular non-symlink file: {path}")
    require(path.stat().st_mode & stat.S_IXUSR, f"{label} is not executable: {path}")
    return path.resolve()


def load_plan(path: Path) -> dict[str, Any]:
    """Load and rebind the frozen 0895 plan to the current 0894 packet."""
    path = Path(path).resolve()
    require(path.is_file() and not path.is_symlink(), f"0895 plan is missing: {path}")
    plan = read_json(path)
    require(isinstance(plan, dict), "0895 plan must be an object")
    require(plan.get("schema") == SCHEMA and plan.get("status") == "frozen", "0895 plan schema/status mismatch")
    require(plan.get("cpu_affinity") == CPU_AFFINITY and plan.get("cache") == "warm", "0895 CPU/cache contract mismatch")
    require(plan.get("case") == PRIMARY_CASE and plan.get("warmup") == WARMUP and plan.get("samples") == SAMPLES, "0895 case protocol mismatch")
    require(plan.get("rows") == [
        {"ordinal": ordinal, "role": role, "case": PRIMARY_CASE, "warmup": WARMUP, "samples": SAMPLES, "pair": f"profile-{ordinal:02d}"}
        for ordinal, role in enumerate(ROLES)
    ], "0895 role schedule mismatch")
    require(plan.get("total_rows") == TOTAL_ROWS and plan.get("total_measured_samples") == TOTAL_SAMPLES, "0895 population mismatch")
    require(plan.get("fresh_child") is True and plan.get("fail_fast") is True, "0895 fail-fast/fresh-child contract missing")
    require(plan.get("candidate_adoption") is False and plan.get("production_adoption") is False, "0895 adoption flag changed")
    base_plan = load_base_plan()
    assert_rejected_0894(base_plan)
    base_ref = plan.get("base_plan")
    same_descriptor(base_ref, descriptor(BASE_PLAN_PATH), "0895 base plan")
    require(plan.get("base_plan_schema") == base_plan.get("schema"), "0895 base plan schema binding mismatch")
    helpers = exact_0894_helper_hashes(base_plan)
    require(plan.get("helper_hashes") == helpers, "0895 helper descriptor set/hash changed")
    require(plan.get("helper_sha256") == {name: item["sha256"] for name, item in helpers.items()}, "0895 helper SHA map changed")
    packet_helpers = plan.get("packet_helpers")
    expected_packet_helpers = {
        name: descriptor(PACKET / name)
        for name in ("common.py", "freeze.py", "capture.py", "export.py", "reduce.py", "verify.py", "attribution.py")
    }
    require(packet_helpers == expected_packet_helpers, "0895 packet helper descriptors changed")
    require(plan.get("evidence") == evidence_descriptors(base_plan), "0895 0894 evidence descriptors changed")
    for role in ("baseline", "candidate"):
        observed = plan.get("binaries", {}).get(role)
        expected = binary(base_plan, role)
        same_descriptor(observed, expected["identity"], f"0895 {role} binary")
        require(observed.get("path") == expected.get("path"), f"0895 {role} binary path changed")
    tools = plan.get("tools")
    require(isinstance(tools, dict), "0895 tool bindings are missing")
    for name in ("heaptrack", "heaptrack_print"):
        item = tools.get(name)
        require(isinstance(item, dict), f"0895 {name} binding is missing")
        tool_path = regular_path(item.get("path"), PACKET, f"0895 {name}")
        same_descriptor(item.get("identity"), descriptor(tool_path), f"0895 {name}")
        require(isinstance(item.get("version"), str) and item["version"].strip(), f"0895 {name} version is missing")
    same_descriptor(tools.get("taskset_identity"), descriptor(Path("/usr/bin/taskset")), "0895 taskset")
    require(plan.get("owner_categories") == ["legacycodec", "mce_stream", "selected_worksheet", "x14ac", "other"], "0895 owner categories mismatch")
    require(plan.get("print_argv_template") == DEFAULT_PRINT_ARGS, "0895 print argv template mismatch")
    require(plan.get("profile_suffix") == ".zst", "0895 profile suffix is not the pinned heaptrack suffix")
    return plan
