#!/usr/bin/env python3
"""Validate and attribute the 0550 XLSX ``MultiSourceEdit::commit`` lane.

The 0550 profile command collects with Callgrind disabled at process start and
toggles only the exact ``MultiSourceEdit::commit`` symbol.  A source-backed
runner invokes that symbol several times in lifecycle gates before the one
operation selected by the profile case.  The numbered dumps are therefore
classified from their positive incoming parent/ancestor edges.  Their ordinal
is checked after classification, but is never used as the scope evidence.

This module is deliberately read-only unless ``--create-annotations`` is
provided.  It validates the build and child receipts, source manifests,
binary identities, profile/native result identities, all lifecycle and
measured raw dumps, and the zero-Ir termination dump.  The measured dump is
annotated with both inclusive and self views.  Immediate owner children are
reported as disjoint raw edge partitions; nested validation, reduced parser,
rewrite, and merge rows are retained as overlapping attribution diagnostics.
No nested inclusive value is interpreted as a removable fraction or a
latency result.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
HELPER_PATH = HERE.parent / "change-0521" / "analyze_profiles.py"
IMPORTED_RAW_PATH = HERE.parent / "change-0519" / "analyze_profiles.py"
IMPORTED_EDGES_PATH = HERE.parent / "change-0519" / "compare_profile_lanes.py"
SCOPE_REVIEW_PATH = HERE / "scope-review.md"
CAPTURE_AMENDMENT_PATH = HERE / "capture-amendment-inputs.json"
STAGE = "baseline"
STAGES = (STAGE,)

OWNER = "litchi_xlsx::cell_values::source::MultiSourceEdit::commit"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
PROFILE_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"

# The source review establishes that the vendor-extension lifecycle gate has
# one extra owner invocation.  This is a validation of the expected harness
# matrix, while the role itself is always established from raw positive edges.
LIFECYCLE_COUNTS = {
    "medium": 3,
    "dense-sparse": 3,
    "noncompact": 3,
    "vendor-extension": 4,
}

# A function can be emitted under either the public wrapper's short Rust name
# or its implementation name after monomorphization.  The first entry in each
# tuple is the source-level target; aliases are accepted only to report a
# codegen spelling, never to silently substitute a different owner.
NESTED_TARGETS: dict[str, tuple[str, ...]] = {
    "validation": (
        "litchi_xlsx::cell_values::validation::validate_xml",
    ),
    "reduced_readback": (
        "litchi_xlsx::raw::worksheet::edit::reduced_readback",
    ),
    "reduced_parser": (
        "litchi_xlsx::raw::worksheet::parse",
        "litchi_xlsx::raw::worksheet::codec::<impl litchi_xlsx::raw::worksheet::model::Parser>::parse",
    ),
    "rewrite": (
        "litchi_xlsx::raw::worksheet::edit::package::rewrite",
        "litchi_xlsx::raw::worksheet::edit::package::rewrite_value_only_with_provenance",
    ),
    "merge": (
        "litchi_xlsx::cell::Store::merge_omitted_cells",
    ),
}

TIMING_KEYS = frozenset(
    {
        "elapsed_ns",
        "open_ns",
        "plan_ns",
        "commit_ns",
        "publication_ns",
        "reopen_ns",
        "commit_allocation_metrics",
        "publication_allocation_metrics",
    }
)

FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([\d,]+)")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")
PART_RE = re.compile(r"^part:\s*(\d+)$")
TRIGGER_RE = re.compile(r"^desc:\s+Trigger:\s+(.*)$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
STAR_RE = re.compile(r"^\s*([\d,]+)\s+\*\s+(.*)$")
EDGE_RE = re.compile(r"^\s*([\d,]+)\s+>\s+(.*)$")
WARNING_RE = re.compile(
    r"^(?:warning|overflow|error|failed|notice)(?:\s|:|$)", re.IGNORECASE
)
PERL_ENV = {"PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"}
HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
SCHEMA = "xlsx_0550_source_commit_attribution_v1"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside the evidence directory: {path}") from error


def regular(path: Path, label: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"{label} is not a regular file")


def digest_value(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase SHA-256")
    return value


def load_helper() -> Any:
    """Load the immutable 0521 parser/annotation helper by path."""

    regular(HELPER_PATH, "retained profile helper")
    previous = sys.modules.get("analyze_profiles")
    spec = importlib.util.spec_from_file_location(
        "litchi_xlsx_0521_profile_helper_for_0550", HELPER_PATH
    )
    require(spec is not None and spec.loader is not None,
            f"cannot load retained profile helper: {HELPER_PATH}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    finally:
        if previous is None:
            sys.modules.pop("analyze_profiles", None)
        else:
            sys.modules["analyze_profiles"] = previous
    # Error labels and relative paths in the helper are otherwise bound to
    # its old evidence directory.  The parser itself remains immutable.
    module.HERE = HERE
    return module


HELPER = load_helper()


def imported_helper_hashes() -> dict[str, str]:
    """Validate the complete immutable helper chain loaded above."""

    paths = {
        str(HELPER_PATH.relative_to(HERE.parent)): HELPER_PATH,
        str(IMPORTED_RAW_PATH.relative_to(HERE.parent)): IMPORTED_RAW_PATH,
        str(IMPORTED_EDGES_PATH.relative_to(HERE.parent)): IMPORTED_EDGES_PATH,
    }
    for label, path in paths.items():
        regular(path, f"retained profile helper {label}")
    require(Path(HELPER.RAW.__file__).resolve() == IMPORTED_RAW_PATH.resolve(),
            "0521 raw helper did not load the bound 0519 parser")
    require(Path(HELPER.EDGES.__file__).resolve() == IMPORTED_EDGES_PATH.resolve(),
            "0521 raw helper did not load the bound 0519 edge helper")
    return {label: sha256(path) for label, path in paths.items()}


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(isinstance(plan.get("revision"), str) and plan["revision"],
            "plan revision is missing")
    require(plan.get("cpu") == 2, "plan CPU differs from the pinned CPU")
    require(plan.get("scope") ==
            "Current source-backed XLSX MultiSourceEdit commit attribution; no optimization or speedup claim",
            "plan scope differs from the frozen attribution scope")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "plan.profile is missing")
    require(profile.get("owner") == OWNER, "profile owner differs from the exact owner")
    require(profile.get("parent") == MEASURED_PARENT,
            "profile measured parent differs from the frozen runner")
    require(profile.get("lifecycle") == LIFECYCLE_PARENT,
            "profile lifecycle parent differs from the frozen runner")
    require(profile.get("case") == PROFILE_CASE,
            "profile case differs from the frozen source-backed case")
    require(profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 1,
            "profile repeat/warmup/sample matrix differs")
    shapes = plan.get("shapes")
    require(shapes == list(LIFECYCLE_COUNTS),
            "profile shape matrix differs from the frozen four-shape matrix")
    require(plan.get("cases") == [
        "xlsx_source_backed_cell_values_one_edit_save",
        PROFILE_CASE,
    ], "source-backed case matrix differs from the frozen plan")
    owned = plan.get("owned_paths")
    require(owned == ["/home/zhuhe/litchi-goal-0550-target"],
            "owned target differs from the frozen plan")
    return plan


def validate_capture_amendment() -> dict[str, Any]:
    """Bind the approved failed-profile continuation to this replay.

    The amendment changes only the Valgrind launch options.  Keeping its
    envelope and every input hash in the report prevents a successful child
    from being mistaken for the preserved pre-amendment failure.
    """

    regular(CAPTURE_AMENDMENT_PATH, "capture amendment inputs")
    amendment = read_json(CAPTURE_AMENDMENT_PATH)
    require(isinstance(amendment, dict), "capture amendment inputs are not an object")
    require(set(amendment) == {"utc", "reason", "inputs", "failed_artifacts"},
            "capture amendment input envelope differs")
    require(isinstance(amendment.get("utc"), str) and amendment["utc"],
            "capture amendment timestamp is missing")
    require(isinstance(amendment.get("reason"), str) and amendment["reason"].strip(),
            "capture amendment reason is missing")
    expected_paths = {
        "plan.json": PLAN_PATH,
        "run.py": RUN_PATH,
        "capture.py": HERE / "capture.py",
        "capture_amended.py": HERE / "capture_amended.py",
        "resume.py": HERE / "resume.py",
        "failed-attempts/manifest.json": HERE / "failed-attempts" / "manifest.json",
    }
    inputs = amendment["inputs"]
    require(isinstance(inputs, dict) and set(inputs) == set(expected_paths),
            "capture amendment input inventory differs")
    input_hashes: dict[str, str] = {}
    for name, path in expected_paths.items():
        regular(path, f"capture amendment input {name}")
        digest = digest_value(inputs.get(name), f"capture amendment input {name}")
        require(digest == sha256(path), f"capture amendment input hash differs: {name}")
        input_hashes[name] = digest
    failed_artifacts = amendment["failed_artifacts"]
    manifest = read_json(expected_paths["failed-attempts/manifest.json"])
    require(failed_artifacts == manifest,
            "capture amendment failed-artifact custody differs")
    return {
        "file": relative(CAPTURE_AMENDMENT_PATH),
        "sha256": sha256(CAPTURE_AMENDMENT_PATH),
        "inputs": input_hashes,
        "failed_artifacts": failed_artifacts,
        "validation": {
            "envelope_bound": True,
            "all_amended_input_hashes_match": True,
            "failed_artifact_custody_bound": True,
        },
    }


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    return [
        {
            "name": f"profile-r{repeat}-{shape}-c0",
            "repeat": repeat,
            "shape": shape,
            "case": profile["case"],
            "warmup": profile["warmup"],
            "samples": profile["samples"],
        }
        for repeat in range(1, profile["repeats"] + 1)
        for shape in plan["shapes"]
    ]


def native_job(job: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    # capture.py launches one native process per case.  The cN suffix is the
    # index in the frozen matrix, while the report's configuration contains
    # the single case supplied to that process.
    case_index = plan["cases"].index(job["case"])
    return {
        "name": f"native-r{job['repeat']}-{job['shape']}-c{case_index}",
        "repeat": job["repeat"],
        "shape": job["shape"],
        "cases": [job["case"]],
        "warmup": plan["native"]["warmup"],
        "samples": plan["native"]["samples"],
    }


def option(command: list[str], name: str) -> str:
    values: list[str] = []
    for index, item in enumerate(command):
        if item == name:
            require(index + 1 < len(command), f"command omits value for {name}")
            values.append(command[index + 1])
        elif item.startswith(name + "="):
            values.append(item.split("=", 1)[1])
    require(len(values) == 1, f"command omits or repeats {name}")
    return values[0]


def validate_host(path: Path, label: str) -> None:
    value = read_json(path)
    require(isinstance(value, dict), f"{label}: host observation is not an object")
    try:
        observed = datetime.datetime.fromisoformat(value["observed_utc"])
    except (KeyError, TypeError, ValueError) as error:
        raise EvidenceError(f"{label}: host timestamp is malformed") from error
    require(observed.tzinfo is not None, f"{label}: host timestamp has no timezone")
    processes = value.get("compiler_processes")
    require(isinstance(processes, list), f"{label}: compiler observation is not a list")
    for index, process in enumerate(processes):
        require(isinstance(process, dict), f"{label}: compiler process {index} is malformed")
        require(isinstance(process.get("pid"), int) and not isinstance(process["pid"], bool)
                and process["pid"] > 0, f"{label}: compiler pid {index} is invalid")
        require(process.get("comm") in {"cargo", "rustc"},
                f"{label}: compiler command {index} is unexpected")
        require(isinstance(process.get("cwd"), str) and process["cwd"],
                f"{label}: compiler cwd {index} is missing")
    require(value.get("scope") == HOST_SCOPE, f"{label}: host scope differs")


def validate_artifacts(folder: Path, receipt: dict[str, Any], expected: set[str],
                       label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: receipt artifacts are missing")
    actual = set(artifacts)
    require(actual == expected,
            f"{label}: artifact set differs; missing={sorted(expected - actual)}, "
            f"extra={sorted(actual - expected)}")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label}: artifact path is not stage-local: {filename!r}")
        digest_value(digest, f"{label}/{filename} receipt hash")
        path = folder / filename
        regular(path, f"{label}/{filename}")
        require(sha256(path) == digest, f"{label}/{filename} hash differs")
        if filename.endswith(".host.json"):
            validate_host(path, f"{label}/{filename}")


def cleanup_binary_hash(cleanup: dict[str, Any], binary_path: Path) -> str | None:
    """Find a cleanup custody digest for a removed stage binary."""

    keys = (
        "binary_sha256_by_path",
        "binary_sha256_by_kind",
        "binary_hashes",
        "binary_sha256",
    )
    for key in keys:
        value = cleanup.get(key)
        if isinstance(value, str) and key == "binary_sha256":
            return value
        if not isinstance(value, dict):
            continue
        for candidate in (
            str(binary_path), binary_path.as_posix(), binary_path.name,
            binary_path.parent.name, "normal", "normal/binary",
        ):
            found = value.get(candidate)
            if isinstance(found, str):
                return found
    return None


def validate_cleanup_for_binary(plan: dict[str, Any], binary_path: Path,
                                expected_sha: str) -> dict[str, Any]:
    path = HERE / "cleanup.json"
    regular(path, "cleanup custody record")
    cleanup = read_json(path)
    require(isinstance(cleanup, dict), "cleanup custody record is not an object")
    require(cleanup.get("plan_sha256") == sha256(PLAN_PATH),
            "cleanup custody plan binding differs")
    require(cleanup.get("owned_paths_absent") is True,
            "cleanup does not prove owned target absence")
    require(cleanup.get("removed", cleanup.get("removed_paths")) == plan["owned_paths"],
            "cleanup does not name the exact owned target")
    require(cleanup.get("accessible_process_references") == [],
            "cleanup retains process references")
    require(all(not os.path.lexists(item) for item in plan["owned_paths"]),
            "owned target remains after claimed cleanup")
    require(cleanup_binary_hash(cleanup, binary_path) == expected_sha,
            "cleanup does not custody the removed normal binary hash")
    return {"file": relative(path), "sha256": sha256(path),
            "owned_paths_absent": True, "removed": plan["owned_paths"]}


def validate_build(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    require(folder.is_dir(), f"{stage}: stage directory is missing")
    manifest_path = folder / "source-manifest.json"
    receipt_path = folder / "build-normal.receipt.json"
    binary_meta_path = folder / "binary-normal.json"
    for path, label in (
        (manifest_path, f"{stage} source manifest"),
        (receipt_path, f"{stage} build receipt"),
        (binary_meta_path, f"{stage} binary identity"),
    ):
        regular(path, label)
    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict) and manifest,
            f"{stage}: source manifest is empty")
    for name, digest in manifest.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"{stage}: source manifest path is invalid")
        digest_value(digest, f"{stage} source manifest {name}")
    manifest_sha = sha256(manifest_path)
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), f"{stage}: build receipt is not an object")
    require(receipt.get("exit_code") == 0, f"{stage}: normal build failed")
    require(receipt.get("binary_sha256") is None,
            f"{stage}: normal build receipt has a child binary hash")
    require(receipt.get("source_manifest_sha256") == manifest_sha,
            f"{stage}: build/source manifest binding differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{stage}: build/plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{stage}: build/run binding differs")
    if "execution_stage" in receipt:
        require(receipt["execution_stage"] == stage,
                f"{stage}: build execution stage differs")
    validate_artifacts(folder, receipt,
                       {"build-normal.host.json", "build-normal.stdout",
                        "build-normal.stderr"},
                       f"{stage}/build-normal")
    binary = read_json(binary_meta_path)
    require(isinstance(binary, dict), f"{stage}: binary identity is not an object")
    binary_sha = digest_value(binary.get("sha256"), f"{stage} binary identity")
    require(isinstance(binary.get("bytes"), int) and not isinstance(binary["bytes"], bool)
            and binary["bytes"] > 0, f"{stage}: binary size is invalid")
    require(binary.get("source_manifest_sha256") == manifest_sha,
            f"{stage}: binary/source manifest binding differs")
    require(binary.get("build_receipt_sha256") == sha256(receipt_path),
            f"{stage}: binary/build binding differs")
    binary_path = Path(binary.get("path", ""))
    require(binary_path.is_absolute(), f"{stage}: binary path is not absolute")
    if binary_path.is_file():
        regular(binary_path, f"{stage} retained binary")
        require(binary_path.stat().st_size == binary["bytes"],
                f"{stage}: retained binary size differs")
        require(sha256(binary_path) == binary_sha,
                f"{stage}: retained binary hash differs")
    else:
        # Validate post-cleanup custody when the owned target has already
        # been removed, but do not emit filesystem-presence state.  Omitting
        # that transient detail keeps a pre-cleanup report byte-identical when
        # the same replay runs after cleanup.
        validate_cleanup_for_binary(plan, binary_path, binary_sha)
    return {
        "stage": stage,
        "source_manifest": relative(manifest_path),
        "source_manifest_sha256": manifest_sha,
        "build_receipt": relative(receipt_path),
        "build_receipt_sha256": sha256(receipt_path),
        "binary_identity": relative(binary_meta_path),
        "binary_sha256": binary_sha,
        "binary_bytes": binary["bytes"],
        "binary_path": str(binary_path),
    }


def validate_execution_binding(receipt: dict[str, Any], stage: str, label: str) -> dict[str, Any]:
    execution_stage = receipt.get("execution_stage")
    if execution_stage is None:
        return {"stage": None, "manifest": None, "manifest_sha256": None}
    require(execution_stage in STAGES, f"{label}: unsupported execution stage")
    require(execution_stage == stage,
            f"{label}: execution stage {execution_stage!r} differs from {stage!r}")
    manifest = HERE / execution_stage / "source-manifest.json"
    regular(manifest, f"{label} execution source manifest")
    digest = digest_value(receipt.get("execution_manifest_sha256"),
                          f"{label} execution manifest hash")
    require(digest == sha256(manifest), f"{label}: execution manifest hash differs")
    return {"stage": execution_stage, "manifest": relative(manifest),
            "manifest_sha256": digest}


def validate_profile_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any],
                             metadata: dict[str, Any], numbered: list[tuple[int, Path]]) -> dict[str, Any]:
    folder = HERE / stage
    name = job["name"]
    path = folder / f"{name}.receipt.json"
    regular(path, relative(path))
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{relative(path)}: receipt is not an object")
    require(receipt.get("exit_code") == 0, f"{relative(path)}: child failed")
    require(receipt.get("binary_sha256") == metadata["binary_sha256"],
            f"{relative(path)}: binary binding differs")
    require(receipt.get("source_manifest_sha256") == metadata["source_manifest_sha256"],
            f"{relative(path)}: source binding differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{relative(path)}: plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{relative(path)}: run binding differs")
    execution = validate_execution_binding(receipt, stage, relative(path))
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{relative(path)}: command is not an argv list")
    require(command[:3] == ["taskset", "-c", str(plan["cpu"])],
            f"{relative(path)}: command is not pinned to the plan CPU")
    for token in ("/usr/bin/time", "-v", "valgrind", "--tool=callgrind",
                  "--collect-atstart=no", f"--toggle-collect={OWNER}",
                  f"--zero-before={OWNER}", f"--dump-after={OWNER}"):
        require(token in command, f"{relative(path)}: command omits {token}")
    require(option(command, "--callgrind-out-file") ==
            str(folder / f"{name}.callgrind"),
            f"{relative(path)}: Callgrind output path differs")
    require(str(metadata["binary_path"]) in command,
            f"{relative(path)}: binary path differs")
    require(option(command, "--case") == job["case"],
            f"{relative(path)}: case differs")
    require(option(command, "--xlsx-cell-crud-shape") == job["shape"],
            f"{relative(path)}: shape differs")
    require(option(command, "--warmup") == str(job["warmup"]),
            f"{relative(path)}: warmup differs")
    require(option(command, "--samples") == str(job["samples"]),
            f"{relative(path)}: samples differs")
    require(option(command, "--json") == str(folder / f"{name}.json"),
            f"{relative(path)}: result path differs")
    require(option(command, "--corpus-manifest") == str(folder / f"{name}.catalog.json"),
            f"{relative(path)}: corpus path differs")
    expected = {
        f"{name}.host.json", f"{name}.json", f"{name}.catalog.json",
        f"{name}.stdout", f"{name}.stderr", f"{name}.callgrind",
        *(f"{name}.callgrind.{number}" for number, _ in numbered),
    }
    validate_artifacts(folder, receipt, expected, f"{stage}/{name}")
    return {
        "file": relative(path),
        "sha256": sha256(path),
        "execution": execution,
        "artifact_count": len(expected),
    }


def validate_native_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any],
                            metadata: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    native = native_job(job, plan)
    name = native["name"]
    path = folder / f"{name}.receipt.json"
    report_path = folder / f"{name}.json"
    regular(path, relative(path))
    regular(report_path, relative(report_path))
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{relative(path)}: receipt is not an object")
    require(receipt.get("exit_code") == 0, f"{relative(path)}: native child failed")
    require(receipt.get("binary_sha256") == metadata["binary_sha256"],
            f"{relative(path)}: native binary binding differs")
    require(receipt.get("source_manifest_sha256") == metadata["source_manifest_sha256"],
            f"{relative(path)}: native source binding differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{relative(path)}: native plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{relative(path)}: native run binding differs")
    execution = validate_execution_binding(receipt, stage, relative(path))
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{relative(path)}: native command is not an argv list")
    require(command[:3] == ["taskset", "-c", str(plan["cpu"])],
            f"{relative(path)}: native command is not pinned")
    require("/usr/bin/time" in command and "-v" in command,
            f"{relative(path)}: native command has no /usr/bin/time -v")
    require(str(metadata["binary_path"]) in command,
            f"{relative(path)}: native binary path differs")
    require(option(command, "--case") == ",".join(native["cases"]),
            f"{relative(path)}: native case matrix differs")
    require(option(command, "--xlsx-cell-crud-shape") == job["shape"],
            f"{relative(path)}: native shape differs")
    require(option(command, "--warmup") == str(native["warmup"]),
            f"{relative(path)}: native warmup differs")
    require(option(command, "--samples") == str(native["samples"]),
            f"{relative(path)}: native samples differs")
    require(option(command, "--json") == str(report_path),
            f"{relative(path)}: native result path differs")
    require(option(command, "--corpus-manifest") == str(folder / f"{name}.catalog.json"),
            f"{relative(path)}: native corpus path differs")
    expected = {
        f"{name}.host.json", f"{name}.json", f"{name}.catalog.json",
        f"{name}.stdout", f"{name}.stderr",
    }
    validate_artifacts(folder, receipt, expected, f"{stage}/{name}")
    return {
        "name": name,
        "cases": native["cases"],
        "file": relative(report_path),
        "sha256": sha256(report_path),
        "receipt": relative(path),
        "receipt_sha256": sha256(path),
        "execution": execution,
    }


def stable_identity_value(value: Any, label: str, key: str | None = None) -> Any:
    if key is not None and (
        key in TIMING_KEYS or key.endswith("_ns") or
        "allocation_metrics" in key or key.startswith("elapsed")
    ):
        return None
    if isinstance(value, dict):
        result: dict[str, Any] = {}
        for child_key, child in sorted(value.items()):
            normalized = stable_identity_value(child, f"{label}.{child_key}", child_key)
            if normalized is not None:
                result[child_key] = normalized
        return result
    if isinstance(value, list):
        normalized = [stable_identity_value(item, f"{label}[{index}]")
                      for index, item in enumerate(value)]
        if not normalized:
            return []
        if all(item == normalized[0] for item in normalized):
            return normalized[0]
        # A non-constant list outside known timing fields is itself useful
        # identity evidence.  Keep it instead of inventing a scalar.
        return normalized
    return value


def validate_result(path: Path, plan: dict[str, Any], stage: str, job: dict[str, Any],
                    binary_sha: str, cases: list[str], samples: int,
                    warmup: int) -> tuple[dict[str, Any], dict[str, Any]]:
    label = relative(path)
    report = read_json(path)
    require(isinstance(report, dict), f"{label}: report is not an object")
    require(report.get("schema_version") == 1, f"{label}: schema version differs")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("binary") == "litchi-perf-baseline",
            f"{label}: tool binary differs")
    require(tool.get("profile") == "release", f"{label}: report is not release profile")
    require(tool.get("instrumentation") == "none",
            f"{label}: report is instrumented")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("binary_sha256") == binary_sha,
            f"{label}: report binary identity differs")
    environment = report.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == plan["revision"],
            f"{label}: report revision differs")
    if "cpu_affinity" in environment:
        require(str(environment["cpu_affinity"]) == str(plan["cpu"]),
                f"{label}: report CPU affinity differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration is missing")
    require(configuration.get("cases") == cases,
            f"{label}: case configuration differs")
    shape_key = "xlsx_cell_crud_shapes"
    if shape_key not in configuration:
        shape_key = "xlsx_cell_crud_shapes"
    require(configuration.get(shape_key) == [job["shape"]],
            f"{label}: shape configuration differs")
    require(configuration.get("samples_per_case") == samples,
            f"{label}: sample configuration differs")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{label}: warmup configuration differs")
    rows = report.get("results")
    require(isinstance(rows, list) and len(rows) == 1,
            f"{label}: expected exactly one result row")
    catalog_path = path.with_name(path.stem + ".catalog.json")
    regular(catalog_path, relative(catalog_path))
    catalog = read_json(catalog_path)
    catalog_reference = report.get("corpus_catalog")
    require(isinstance(catalog, dict) and isinstance(catalog_reference, dict),
            f"{label}: corpus catalog identity is missing")
    require(catalog_reference.get("catalog_sha256") == catalog.get("catalog_sha256")
            and catalog_reference.get("content_set_sha256") == catalog.get("content_set_sha256"),
            f"{label}: report/catalog identity differs")
    matching = [row for row in rows if isinstance(row, dict) and row.get("case") == job["case"]]
    require(len(matching) == 1,
            f"{label}: expected one result row for {job['case']}, got {len(matching)}")
    row = matching[0]
    corpus = row.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == job["shape"],
            f"{label}: corpus shape differs")
    output = row.get("output_sha256")
    require(isinstance(output, str) and SHA256_RE.fullmatch(output),
            f"{label}: output digest is malformed")
    sink = row.get("sink")
    require(isinstance(sink, dict), f"{label}: sink identity is missing")
    source = row.get("source")
    require(isinstance(source, dict), f"{label}: source identity is missing")
    logical = {
        "corpus": corpus,
        "sink": sink,
        "source": stable_identity_value(source, f"{label}.source"),
        "output_sha256": output,
    }
    return report, logical


def numbered_paths(folder: Path, name: str) -> list[tuple[int, Path]]:
    prefix = f"{name}.callgrind."
    found: list[tuple[int, Path]] = []
    for path in folder.glob(prefix + "*"):
        suffix = path.name[len(prefix):]
        if suffix.isdigit() and path.is_file() and not path.is_symlink():
            found.append((int(suffix), path))
    found.sort()
    require(found, f"{folder.name}/{name}: no numbered Callgrind dumps")
    require([number for number, _ in found] == list(range(1, len(found) + 1)),
            f"{folder.name}/{name}: numbered dumps are not contiguous")
    return found


def parse_cost(line: str, event_index: int) -> int | None:
    try:
        return HELPER._cost(line, event_index)
    except (AttributeError, TypeError, ValueError):
        fields = line.split()
        if len(fields) <= event_index + 1:
            return None
        position = fields[0]
        if not position or (position[0] not in "+-*" and not position[0].isdigit()):
            return None
        try:
            return int(fields[event_index + 1])
        except ValueError:
            return None


def parse_raw(path: Path) -> dict[str, Any]:
    """Parse function sections, self costs, and raw call edges."""

    text = read_text(path)
    events: list[str] = []
    summary: int | None = None
    names: dict[int, str] = {}
    functions: dict[int, dict[str, Any]] = {}
    current: int | None = None
    pending_callee: int | None = None
    pending_calls: int | None = None
    pending_line: int | None = None
    for line_number, raw_line in enumerate(text.splitlines(), 1):
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("events:"):
            events = line.split(":", 1)[1].split()
            continue
        match = SUMMARY_RE.match(line)
        if match:
            values: list[int] = []
            for token in match.group(1).split():
                try:
                    values.append(int(token.replace(",", "")))
                except ValueError:
                    break
            require(values, f"{relative(path)}: summary has no integer")
            summary = values[0]
            continue
        function_match = FUNCTION_RE.match(line)
        if function_match:
            kind, text_id, function_name = function_match.groups()
            function_id = int(text_id)
            function = functions.setdefault(
                function_id, {"name": "", "self_ir": 0, "edges": []}
            )
            if function_name:
                names[function_id] = function_name
                function["name"] = function_name
            if kind == "fn":
                current = function_id
                pending_callee = pending_calls = pending_line = None
            else:
                require(current is not None,
                        f"{relative(path)}:{line_number}: cfn outside fn")
                pending_callee = function_id
                pending_calls = None
                pending_line = line_number
            continue
        calls_match = CALLS_RE.match(line)
        if calls_match:
            require(current is not None and pending_callee is not None,
                    f"{relative(path)}:{line_number}: calls outside cfn")
            require(pending_calls is None,
                    f"{relative(path)}:{line_number}: repeated calls record")
            pending_calls = int(calls_match.group(1).replace(",", ""))
            continue
        if line.startswith((
            "creator:", "cmd:", "pid:", "version:", "positions:", "part:",
            "ob=", "cob=", "fl=", "cfi=", "fi=", "fe=", "totals:",
        )):
            continue
        if current is None or events != ["Ir"]:
            continue
        cost = parse_cost(line, 0)
        if cost is None:
            continue
        if pending_callee is not None:
            functions[current]["edges"].append({
                "callee_id": pending_callee,
                "calls": pending_calls,
                "inclusive_ir": cost,
                "line": pending_line,
            })
        else:
            functions[current]["self_ir"] += cost
        pending_callee = pending_calls = pending_line = None
    require(events == ["Ir"], f"{relative(path)}: event set is not exactly Ir")
    require(summary is not None, f"{relative(path)}: summary is missing")
    for function_id, function in functions.items():
        if not function["name"]:
            function["name"] = names.get(function_id, "")
    selected_ids = [
        function_id for function_id, function in functions.items()
        if function["name"] == OWNER
    ]
    require(len(selected_ids) == 1,
            f"{relative(path)}: exact owner matched {selected_ids}")
    selected_id = selected_ids[0]
    incoming: list[dict[str, Any]] = []
    for parent_id, function in functions.items():
        for edge in function["edges"]:
            if (edge["callee_id"] == selected_id and edge["calls"] is not None
                    and edge["calls"] > 0 and edge["inclusive_ir"] > 0):
                incoming.append({
                    "caller_id": parent_id,
                    "caller": function["name"] or names.get(parent_id, ""),
                    "callee_id": selected_id,
                    "callee": OWNER,
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "line": edge["line"],
                })
    require(len(incoming) == 1,
            f"{relative(path)}: exact owner has {len(incoming)} positive incoming edges")
    direct_ir = sum(edge["inclusive_ir"] for edge in functions[selected_id]["edges"])
    self_ir = functions[selected_id]["self_ir"]
    edge = incoming[0]
    require(edge["inclusive_ir"] == summary,
            f"{relative(path)}: owner edge does not equal summary")
    require(self_ir + direct_ir == edge["inclusive_ir"],
            f"{relative(path)}: owner self plus direct equation does not balance")
    return {
        "file": relative(path),
        "sha256": sha256(path),
        "summary_ir": summary,
        "events": events,
        "functions": functions,
        "owner_id": selected_id,
        "owner_incoming": edge,
        "owner_self_ir": self_ir,
        "owner_direct_ir": direct_ir,
        "warnings": [line for line in text.splitlines() if WARNING_RE.search(line)],
        "validation": {
            "event_set_exactly_ir": True,
            "exact_owner_symbol_present_once": True,
            "one_positive_owner_incoming_edge": True,
            "owner_edge_matches_summary": True,
            "owner_self_plus_direct_matches_incoming": True,
        },
    }


def name_matches(actual: str, expected: str) -> bool:
    return actual == expected or actual.endswith("::" + expected.rsplit("::", 1)[-1])


def positive_ancestor_path(parsed: dict[str, Any], target_id: int,
                           ancestor: str) -> dict[str, Any] | None:
    """Find a bounded positive-Ir path from an ancestor to a target."""

    functions = parsed["functions"]
    starts = sorted(
        function_id for function_id, function in functions.items()
        if name_matches(function["name"], ancestor)
    )
    if not starts:
        return None
    queue: list[tuple[int, list[int], list[dict[str, Any]]]] = [
        (function_id, [function_id], []) for function_id in starts
    ]
    visited = set(starts)
    while queue:
        current, ids, edges = queue.pop(0)
        if current == target_id:
            return {
                "ancestor": ancestor,
                "function_ids": ids,
                "functions": [functions.get(item, {}).get("name", "") for item in ids],
                "edges": edges,
                "depth": len(edges),
                "max_depth": 8,
            }
        if len(edges) >= 8:
            continue
        function = functions.get(current)
        if function is None:
            continue
        outgoing = sorted(
            (
                edge for edge in function["edges"]
                if edge["inclusive_ir"] > 0 and edge["calls"] is not None
                and edge["calls"] >= 0
            ),
            key=lambda edge: (
                functions.get(edge["callee_id"], {}).get("name", ""),
                edge["callee_id"], edge["line"] or 0,
            ),
        )
        for edge in outgoing:
            child = edge["callee_id"]
            if child in visited:
                continue
            visited.add(child)
            queue.append((
                child,
                [*ids, child],
                [*edges, {
                    "caller_id": current,
                    "caller": functions.get(current, {}).get("name", ""),
                    "callee_id": child,
                    "callee": functions.get(child, {}).get("name", ""),
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "line": edge["line"],
                }],
            ))
    return None


def classify_dump(parsed: dict[str, Any], number: int) -> dict[str, Any]:
    owner_edge = parsed["owner_incoming"]
    parent_id = owner_edge["caller_id"]
    caller = owner_edge["caller"]
    direct_lifecycle = name_matches(caller, LIFECYCLE_PARENT)
    direct_measured = name_matches(caller, MEASURED_PARENT)
    require(not (direct_lifecycle and direct_measured),
            f"{parsed['file']}: owner caller matches both scope parents")
    if direct_lifecycle:
        lifecycle_path = {
            "ancestor": LIFECYCLE_PARENT,
            "function_ids": [parent_id],
            "functions": [caller],
            "edges": [],
            "depth": 0,
            "max_depth": 8,
            "direct_owner_parent": True,
        }
        measured_path = None
    elif direct_measured:
        lifecycle_path = None
        measured_path = {
            "ancestor": MEASURED_PARENT,
            "function_ids": [parent_id],
            "functions": [caller],
            "edges": [],
            "depth": 0,
            "max_depth": 8,
            "direct_owner_parent": True,
        }
    else:
        # Only an unlabelled immediate caller needs ancestry search.  A
        # collection-off context edge can retain positive Ir while reporting
        # calls=0 (for example, the measured runner surrounding a lifecycle
        # gate); searching from an already-labelled caller would therefore
        # make a lifecycle dump appear to belong to the measured runner.
        lifecycle_path = positive_ancestor_path(parsed, parent_id, LIFECYCLE_PARENT)
        measured_path = positive_ancestor_path(parsed, parent_id, MEASURED_PARENT)
        require(not (lifecycle_path is not None and measured_path is not None),
                f"{parsed['file']}: owner parent is reachable from both lifecycle and measured runners")
    if measured_path is not None:
        role = "measured"
        ancestor = MEASURED_PARENT
        selected_path = measured_path
    elif lifecycle_path is not None:
        role = "lifecycle"
        ancestor = LIFECYCLE_PARENT
        selected_path = lifecycle_path
    else:
        raise EvidenceError(
            f"{parsed['file']}: positive owner parent {caller!r} has no allowed lifecycle/measured ancestor"
        )
    require(owner_edge["calls"] == 1,
            f"{parsed['file']}: classified {role} owner edge has {owner_edge['calls']} calls")
    return {
        "number": number,
        "role": role,
        "owner_parent": caller,
        "owner_calls": owner_edge["calls"],
        "summary_ir": parsed["summary_ir"],
        "owner_incoming": owner_edge,
        "positive_ancestor": selected_path,
        "lifecycle_ancestor": lifecycle_path,
        "measured_ancestor": measured_path,
        "file": parsed["file"],
        "sha256": parsed["sha256"],
        "warnings": parsed["warnings"],
        "validation": {
            "positive_parent_or_ancestor_classified": True,
            "lifecycle_excluded_or_measured_included": True,
            "single_owner_call": True,
        },
    }


def validate_numbered_dump(path: Path, number: int) -> tuple[dict[str, Any], dict[str, Any]]:
    text = read_text(path)
    label = relative(path)
    require(HELPER.part_number(text, label) == number,
            f"{label}: part number differs from suffix")
    require(HELPER.trigger(text, label) == f"--dump-after={OWNER}",
            f"{label}: trigger is not the exact owner")
    parsed = parse_raw(path)
    classification = classify_dump(parsed, number)
    return {
        **classification,
        "owner_self_ir": parsed["owner_self_ir"],
        "owner_direct_ir": parsed["owner_direct_ir"],
        "validation": {
            **parsed["validation"],
            **classification["validation"],
            "part_matches_suffix": True,
            "exact_owner_trigger": True,
        },
    }, parsed


def validate_termination(path: Path, expected_part: int) -> dict[str, Any]:
    regular(path, relative(path))
    text = read_text(path)
    label = relative(path)
    require(HELPER.part_number(text, label) == expected_part,
            f"{label}: termination part differs from expected {expected_part}")
    require(HELPER.trigger(text, label) == "Program termination",
            f"{label}: termination trigger differs")
    lines = text.splitlines()
    require(lines.count("events: Ir") == 1 and all(
        not line.startswith("events:") or line == "events: Ir" for line in lines
    ), f"{label}: termination event set is not exactly Ir")
    values = [
        int(match.group(1).replace(",", ""))
        for line in lines if (match := SUMMARY_RE.match(line.strip()))
    ]
    require(len(values) == 1 and values[0] == 0,
            f"{label}: termination summary is {values}, expected zero")
    return {
        "file": label,
        "sha256": sha256(path),
        "part": expected_part,
        "trigger": "Program termination",
        "summary_ir": 0,
        "validation": {"termination_event_set_exactly_ir": True,
                        "termination_zero_ir": True},
    }


def annotation_command(path: Path, inclusive: bool) -> list[str]:
    return [
        "callgrind_annotate", "--auto=no", "--threshold=100",
        "--show-percs=no", "--inclusive=" + ("yes" if inclusive else "no"),
        "--tree=both", str(path),
    ]


def run_annotation(path: Path, inclusive: bool) -> tuple[str, list[str]]:
    command = annotation_command(path, inclusive)
    environment = dict(os.environ)
    environment.update(PERL_ENV)
    try:
        process = subprocess.run(command, capture_output=True, text=True,
                                 check=True, env=environment)
    except (OSError, subprocess.CalledProcessError) as error:
        raise EvidenceError(f"callgrind_annotate failed for {relative(path)}: {error}") from error
    require(process.stderr == "", f"callgrind_annotate emitted stderr for {relative(path)}")
    return process.stdout, command


def display_name(text: str) -> str:
    text = text.rsplit(" [", 1)[0].strip()
    text = re.sub(r"\s+\([\d,]+x\)$", "", text)
    if text.startswith("???:"):
        return text[4:]
    return text.rsplit(":", 1)[-1] if text.startswith(("./", "/")) else text


def parse_annotation(text: str, selected: str, label: str) -> dict[str, Any]:
    rows: list[tuple[int, int]] = []
    lines = text.splitlines()
    for index, line in enumerate(lines):
        match = STAR_RE.match(line)
        if match and display_name(match.group(2)) == selected:
            rows.append((index, int(match.group(1).replace(",", ""))))
    require(len(rows) == 1,
            f"{label}: expected one annotation row for {selected}, got {rows}")
    index, selected_ir = rows[0]
    direct: list[dict[str, Any]] = []
    for line in lines[index + 1:]:
        match = EDGE_RE.match(line)
        if not match:
            break
        rendered = match.group(2)
        calls_match = re.search(r"\(([\d,]+)x\)", rendered)
        direct.append({
            "name": display_name(rendered),
            "inclusive_ir": int(match.group(1).replace(",", "")),
            "calls": int(calls_match.group(1).replace(",", ""))
            if calls_match else None,
        })
    return {"selected_ir": selected_ir, "direct": direct}


def direct_name_map(edges: list[dict[str, Any]]) -> dict[str, int]:
    values: dict[str, int] = {}
    for edge in edges:
        values[edge["name"]] = values.get(edge["name"], 0) + edge["inclusive_ir"]
    return values


def write_exact(path: Path, content: str) -> None:
    path = path.resolve()
    label = relative(path) if path.is_relative_to(HERE) else str(path)
    encoded = content.encode("utf-8")
    if path.exists():
        regular(path, label)
        require(path.read_bytes() == encoded,
                f"{label} differs from deterministic replay")
        return
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
    except FileExistsError:
        require(path.read_bytes() == encoded,
                f"{label} differs from concurrent deterministic replay")


def annotate_measured(folder: Path, name: str, selected: Path,
                      expected_ir: int, create_annotations: bool) -> dict[str, Any]:
    inclusive_text, inclusive_command = run_annotation(selected, True)
    self_text, self_command = run_annotation(selected, False)
    inclusive_path = folder / f"{name}.inclusive.txt"
    self_path = folder / f"{name}.self.txt"
    if create_annotations:
        write_exact(inclusive_path, inclusive_text)
        write_exact(self_path, self_text)
    else:
        regular(inclusive_path, relative(inclusive_path))
        regular(self_path, relative(self_path))
        require(read_text(inclusive_path) == inclusive_text,
                f"{relative(inclusive_path)} differs from deterministic replay")
        require(read_text(self_path) == self_text,
                f"{relative(self_path)} differs from deterministic replay")
    inclusive = parse_annotation(inclusive_text, OWNER, relative(inclusive_path))
    exclusive = parse_annotation(self_text, OWNER, relative(self_path))
    inc_map = direct_name_map(inclusive["direct"])
    self_map = direct_name_map(exclusive["direct"])
    require(inclusive["selected_ir"] == expected_ir,
            f"{relative(inclusive_path)}: owner Ir differs from raw summary")
    require(inc_map == self_map,
            f"{name}: inclusive/self direct children differ")
    require(exclusive["selected_ir"] + sum(self_map.values()) == inclusive["selected_ir"],
            f"{name}: owner self plus direct annotation Ir does not balance")
    return {
        "selected_dump": relative(selected),
        "environment": dict(PERL_ENV),
        "command": {
            "inclusive": [*inclusive_command[:-1], relative(selected)],
            "self": [*self_command[:-1], relative(selected)],
        },
        "inclusive": {
            "file": relative(inclusive_path),
            "sha256": sha256(inclusive_path),
            "owner_ir": inclusive["selected_ir"],
            "owner_self_ir": exclusive["selected_ir"],
            "direct_callee_ir": dict(sorted(inc_map.items())),
        },
        "self": {"file": relative(self_path), "sha256": sha256(self_path)},
        "validation": {
            "owner_matches_raw_summary": True,
            "self_plus_direct_equation": True,
            "inclusive_and_self_direct_children_match": True,
            "deterministic_annotation_environment": True,
        },
    }


def immediate_children(parsed: dict[str, Any], annotation: dict[str, Any]) -> dict[str, Any]:
    """Return disjoint immediate owner edge partitions keyed by function ID."""

    functions = parsed["functions"]
    owner = functions[parsed["owner_id"]]
    by_id: dict[int, dict[str, Any]] = {}
    for edge in owner["edges"]:
        child_id = edge["callee_id"]
        child = by_id.setdefault(child_id, {
            "function_id": child_id,
            "name": functions.get(child_id, {}).get("name", ""),
            "calls": 0,
            "edge_count": 0,
            "inclusive_ir": 0,
            "child_self_ir": functions.get(child_id, {}).get("self_ir", 0),
        })
        child["calls"] += edge["calls"] or 0
        child["edge_count"] += 1
        child["inclusive_ir"] += edge["inclusive_ir"]
    rows = sorted(by_id.values(), key=lambda item: (-item["inclusive_ir"], item["function_id"]))
    require(sum(item["inclusive_ir"] for item in rows) == parsed["owner_direct_ir"],
            f"{parsed['file']}: immediate owner edges do not partition direct Ir")
    annotation_map = annotation["inclusive"]["direct_callee_ir"]
    raw_map = direct_name_map([
        {"name": item["name"], "inclusive_ir": item["inclusive_ir"]}
        for item in rows
    ])
    require(annotation_map == raw_map,
            f"{parsed['file']}: annotation direct children differ from raw edges")
    return {
        "scope": "Immediate raw owner child edges; each function ID is one disjoint edge partition",
        "rows": rows,
        "direct_ir": parsed["owner_direct_ir"],
        "sum_inclusive_ir": sum(item["inclusive_ir"] for item in rows),
        "validation": {
            "function_id_rows_are_disjoint": True,
            "sum_equals_owner_direct_ir": True,
            "annotation_edges_match_raw_edges": True,
            "nested_descendants_not_added_to_owner_partition": True,
        },
    }


def owner_reachable(parsed: dict[str, Any]) -> tuple[set[int], dict[int, int]]:
    functions = parsed["functions"]
    root = parsed["owner_id"]
    reachable = {root}
    depths = {root: 0}
    queue = [root]
    while queue:
        current = queue.pop(0)
        outgoing = sorted(
            (edge for edge in functions[current]["edges"]
             if edge["inclusive_ir"] > 0 and edge["calls"] is not None
             and edge["calls"] >= 0),
            key=lambda edge: (edge["callee_id"], edge["line"] or 0),
        )
        for edge in outgoing:
            child = edge["callee_id"]
            if child not in reachable:
                reachable.add(child)
                depths[child] = depths[current] + 1
                queue.append(child)
    return reachable, depths


def nested_target(parsed: dict[str, Any], category: str) -> dict[str, Any]:
    functions = parsed["functions"]
    reachable, depths = owner_reachable(parsed)
    names = NESTED_TARGETS[category]
    ids = sorted(
        function_id for function_id in reachable
        if functions[function_id].get("name") in names
    )
    all_ids = sorted(
        function_id for function_id, function in functions.items()
        if function.get("name") in names
    )
    incoming: list[dict[str, Any]] = []
    outside: list[dict[str, Any]] = []
    selected = set(ids)
    for parent_id, function in functions.items():
        for edge in function["edges"]:
            if edge["callee_id"] not in selected or edge["calls"] is None \
                    or edge["calls"] <= 0 or edge["inclusive_ir"] <= 0:
                continue
            value = {
                "caller_id": parent_id,
                "caller": function.get("name", ""),
                "callee_id": edge["callee_id"],
                "callee": functions[edge["callee_id"]].get("name", ""),
                "calls": edge["calls"],
                "inclusive_ir": edge["inclusive_ir"],
                "line": edge["line"],
            }
            (incoming if parent_id in reachable else outside).append(value)
    incoming.sort(key=lambda row: (row["callee"], row["caller"], row["line"] or 0))
    outside.sort(key=lambda row: (row["callee"], row["caller"], row["line"] or 0))
    direct_children: list[dict[str, Any]] = []
    for function_id in ids:
        child_map: dict[int, dict[str, Any]] = {}
        for edge in functions[function_id]["edges"]:
            child = child_map.setdefault(edge["callee_id"], {
                "function_id": edge["callee_id"],
                "name": functions.get(edge["callee_id"], {}).get("name", ""),
                "calls": 0,
                "inclusive_ir": 0,
            })
            child["calls"] += edge["calls"] or 0
            child["inclusive_ir"] += edge["inclusive_ir"]
        direct_children.append({
            "function_id": function_id,
            "name": functions[function_id].get("name", ""),
            "children": sorted(child_map.values(),
                                key=lambda row: (-row["inclusive_ir"], row["function_id"])),
        })
    self_ir = sum(functions[function_id]["self_ir"] for function_id in ids)
    direct_ir = sum(
        edge["inclusive_ir"] for function_id in ids for edge in functions[function_id]["edges"]
    )
    return {
        "category": category,
        "source_level_targets": list(names),
        "matched_function_ids": ids,
        "all_raw_matching_function_ids": all_ids,
        "present_in_owner_graph": bool(ids),
        "out_of_line_positive_edge": bool(incoming),
        "minimum_owner_depth": min((depths[item] for item in ids), default=None),
        "inclusive_ir": sum(row["inclusive_ir"] for row in incoming),
        "self_ir": self_ir,
        "direct_ir": direct_ir,
        "calls": sum(row["calls"] for row in incoming),
        "positive_incoming_edges": incoming,
        "positive_incoming_edges_outside_owner_graph": outside,
        "function_direct_children": direct_children,
        "overlap_policy": (
            "Nested inclusive/self rows overlap their callers and descendants; "
            "they are reported as attribution diagnostics and are not summed into "
            "a removable fraction."
        ),
        "absence_reason": None if ids else (
            "target absent from this dump's positive owner-reachable graph; "
            "the compiler may have inlined or omitted the source-level helper"
        ),
        "validation": {
            "owner_reachable_positive_edges_define_scope": True,
            "target_presence_or_absence_recorded": True,
            "incoming_edges_restricted_to_owner_graph": True,
            "nested_rows_not_treated_as_disjoint": True,
        },
    }


def nested_attribution(parsed: dict[str, Any]) -> dict[str, Any]:
    return {
        "scope": (
            "Positive owner-reachable Callgrind graph for the one measured commit; "
            "validation, reduced parser/readback, rewrite, and merge are named rows"
        ),
        "targets": {category: nested_target(parsed, category)
                    for category in NESTED_TARGETS},
        "validation": {
            "all_named_targets_reported": True,
            "owner_graph_is_positive_edge_scoped": True,
            "inclusive_rows_are_explicitly_overlapping": True,
            "no_removable_fraction_inferred": True,
        },
    }


def analyze_profile(stage: str, plan: dict[str, Any], metadata: dict[str, Any],
                    job: dict[str, Any], create_annotations: bool) -> dict[str, Any]:
    folder = HERE / stage
    name = job["name"]
    numbered = numbered_paths(folder, name)
    profile_receipt = validate_profile_receipt(stage, job, plan, metadata, numbered)
    profile_path = folder / f"{name}.json"
    regular(profile_path, relative(profile_path))
    profile_report, profile_identity = validate_result(
        profile_path, plan, stage, job, metadata["binary_sha256"],
        [job["case"]], job["samples"], job["warmup"]
    )
    native_meta = validate_native_receipt(stage, job, plan, metadata)
    native_report, native_identity = validate_result(
        HERE / native_meta["file"], plan, stage, job, metadata["binary_sha256"],
        native_meta["cases"], plan["native"]["samples"],
        plan["native"]["warmup"]
    )
    _ = profile_report, native_report
    require(profile_identity == native_identity,
            f"{name}: profile/native corpus, sink, source, or output identity differs")
    dumps: list[dict[str, Any]] = []
    parsed_by_number: dict[int, dict[str, Any]] = {}
    for number, path in numbered:
        dump, parsed = validate_numbered_dump(path, number)
        dumps.append(dump)
        parsed_by_number[number] = parsed
    lifecycle = [dump for dump in dumps if dump["role"] == "lifecycle"]
    measured = [dump for dump in dumps if dump["role"] == "measured"]
    expected_lifecycle = LIFECYCLE_COUNTS[job["shape"]]
    require(len(lifecycle) == expected_lifecycle,
            f"{name}: expected {expected_lifecycle} lifecycle dumps, found {len(lifecycle)}")
    require(len(measured) == 1,
            f"{name}: expected exactly one measured commit dump, found {len(measured)}")
    require(len(dumps) == expected_lifecycle + 1,
            f"{name}: total owner dump count differs from lifecycle plus one measured")
    measured_dump = measured[0]
    require(measured_dump["number"] == expected_lifecycle + 1,
            f"{name}: measured dump number differs after positive-edge classification")
    require([dump["role"] for dump in dumps] ==
            ["lifecycle"] * expected_lifecycle + ["measured"],
            f"{name}: lifecycle/measured dump order differs")
    selected_path = HERE / measured_dump["file"]
    parsed = parsed_by_number[measured_dump["number"]]
    annotation = annotate_measured(
        folder, name, selected_path, measured_dump["summary_ir"], create_annotations
    )
    children = immediate_children(parsed, annotation)
    nested = nested_attribution(parsed)
    termination = validate_termination(folder / f"{name}.callgrind", len(dumps) + 1)
    return {
        "name": name,
        "stage": stage,
        "repeat": job["repeat"],
        "shape": job["shape"],
        "case": job["case"],
        "owner": OWNER,
        "receipt": profile_receipt,
        "profile_result": {"file": relative(profile_path),
                            "sha256": sha256(profile_path)},
        "native_result": native_meta,
        "profile_native_identity": profile_identity,
        "raw_dumps": dumps,
        "excluded_dumps": [
            {**dump, "excluded_from_measured_attribution": True,
             "exclusion_reason": "lifecycle gate outside selected measured runner"}
            for dump in lifecycle
        ],
        "measured_dump": measured_dump,
        "termination": termination,
        "annotations": annotation,
        "commit_attribution": {
            "owner": OWNER,
            "caller": measured_dump["owner_parent"],
            "calls": measured_dump["owner_calls"],
            "inclusive_ir": measured_dump["summary_ir"],
            "self_ir": parsed["owner_self_ir"],
            "direct_ir": parsed["owner_direct_ir"],
            "immediate_children": children,
            "accounting": {
                "self_plus_direct_ir": parsed["owner_self_ir"] + parsed["owner_direct_ir"],
                "owner_inclusive_ir": parsed["summary_ir"],
                "balances": parsed["owner_self_ir"] + parsed["owner_direct_ir"]
                == parsed["summary_ir"],
            },
        },
        "nested_attribution": nested,
        "validation": {
            "exact_owner_scope": True,
            "positive_ancestor_scope_classification": True,
            "lifecycle_dumps_excluded": len(lifecycle) == expected_lifecycle,
            "exactly_one_measured_commit_dump": len(measured) == 1,
            "shape_specific_lifecycle_count": True,
            "termination_zero_ir": True,
            "profile_native_identity_equal": True,
            "immediate_children_disjoint": True,
            "nested_targets_attributed_without_removable_fraction": True,
        },
        "limitations": [
            "Callgrind Ir is a mechanism diagnostic, not native latency, hardware cycles, or allocation counts.",
            "The native commit phase includes edit staging; the exact owner profile excludes staging and publication.",
            "Nested inclusive rows overlap and do not identify removable work.",
        ],
    }


def analyze_stage(stage: str, plan: dict[str, Any], create_annotations: bool) -> dict[str, Any]:
    metadata = validate_build(stage, plan)
    profiles = [analyze_profile(stage, plan, metadata, job, create_annotations)
                for job in profile_jobs(plan)]
    require(len(profiles) == 8, f"{stage}: profile matrix is incomplete")
    require(sum(len(row["measured_dump"]) > 0 for row in profiles) == 8,
            f"{stage}: measured profile count is incomplete")
    return {
        "stage": stage,
        "status": "pass",
        "metadata": metadata,
        "profile_count": len(profiles),
        "measured_commit_count": sum(
            1 for row in profiles if row["validation"]["exactly_one_measured_commit_dump"]
        ),
        "lifecycle_dump_count": sum(len(row["excluded_dumps"]) for row in profiles),
        "profiles": profiles,
        "validation": {
            "source_binary_receipt_bindings": True,
            "eight_profile_jobs": True,
            "all_native_counterparts_valid": True,
            "all_lifecycle_and_measured_roles_from_positive_edges": True,
            "one_measured_commit_per_job": True,
            "all_termination_dumps_zero_ir": True,
            "all_nested_attribution_rows_present": True,
        },
    }


def repeat_metric(first: int, second: int) -> dict[str, Any]:
    delta = second - first
    result = {
        "repeat_1": first,
        "repeat_2": second,
        "delta_ir": delta,
        "repeat_2_over_repeat_1": None,
        "observed_delta_percent": None,
    }
    if first != 0:
        result["repeat_2_over_repeat_1"] = second / first
        result["observed_delta_percent"] = (second / first - 1.0) * 100.0
    return result


def compare_repeats(profiles: list[dict[str, Any]]) -> dict[str, Any]:
    """Report same-binary repeat drift by semantic function names."""

    by_shape = {(row["shape"], row["repeat"]): row for row in profiles}
    require(set(by_shape) == {
        (shape, repeat) for shape in LIFECYCLE_COUNTS for repeat in (1, 2)
    }, "baseline repeat matrix is incomplete")
    rows: list[dict[str, Any]] = []
    for shape in sorted(LIFECYCLE_COUNTS):
        first, second = by_shape[(shape, 1)], by_shape[(shape, 2)]
        require(first["profile_native_identity"] == second["profile_native_identity"],
                f"{shape}: repeat output identity differs")
        first_children = {
            row["name"]: row["inclusive_ir"]
            for row in first["commit_attribution"]["immediate_children"]["rows"]
        }
        second_children = {
            row["name"]: row["inclusive_ir"]
            for row in second["commit_attribution"]["immediate_children"]["rows"]
        }
        child_names = sorted(set(first_children) | set(second_children))
        first_targets = first["nested_attribution"]["targets"]
        second_targets = second["nested_attribution"]["targets"]
        rows.append({
            "shape": shape,
            "repeat_1_profile": first["name"],
            "repeat_2_profile": second["name"],
            "profile_native_identity_equal": True,
            "commit_ir": {
                "inclusive": repeat_metric(
                    first["commit_attribution"]["inclusive_ir"],
                    second["commit_attribution"]["inclusive_ir"],
                ),
                "self": repeat_metric(
                    first["commit_attribution"]["self_ir"],
                    second["commit_attribution"]["self_ir"],
                ),
                "direct": repeat_metric(
                    first["commit_attribution"]["direct_ir"],
                    second["commit_attribution"]["direct_ir"],
                ),
            },
            "immediate_child_ir": [
                {
                    "name": name,
                    "repeat_1": first_children.get(name, 0),
                    "repeat_2": second_children.get(name, 0),
                    "metric": repeat_metric(
                        first_children.get(name, 0), second_children.get(name, 0)
                    ),
                }
                for name in child_names
            ],
            "nested_targets": {
                category: {
                    "inclusive_ir": repeat_metric(
                        first_targets[category]["inclusive_ir"],
                        second_targets[category]["inclusive_ir"],
                    ),
                    "self_ir": repeat_metric(
                        first_targets[category]["self_ir"],
                        second_targets[category]["self_ir"],
                    ),
                    "direct_ir": repeat_metric(
                        first_targets[category]["direct_ir"],
                        second_targets[category]["direct_ir"],
                    ),
                }
                for category in NESTED_TARGETS
            },
        })
    return {
        "available": True,
        "profiles": rows,
        "profile_count": len(rows),
        "scope": "Same-baseline-binary repeat drift; observed changes are diagnostic only",
        "validation": {
            "repeat_matrix_matches": True,
            "profile_native_identity_parity": True,
            "immediate_child_partitions_compared_by_name": True,
            "nested_targets_compared_without_removable_fraction": True,
        },
    }


def analyze(create_annotations: bool = False) -> dict[str, Any]:
    plan = plan_data()
    amendment = validate_capture_amendment()
    imports = imported_helper_hashes()
    stages = {STAGE: analyze_stage(STAGE, plan, create_annotations)}
    repeat_drift = compare_repeats(stages[STAGE]["profiles"])
    return {
        "schema": SCHEMA,
        "status": "pass",
        "scope": plan["scope"],
        "performance_claim": "none",
        "selected_function": OWNER,
        "lifecycle_parent": LIFECYCLE_PARENT,
        "measured_parent": MEASURED_PARENT,
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "run_script": relative(RUN_PATH),
        "run_script_sha256": sha256(RUN_PATH),
        "scope_review": {
            "file": relative(SCOPE_REVIEW_PATH) if SCOPE_REVIEW_PATH.is_file() else None,
            "sha256": sha256(SCOPE_REVIEW_PATH) if SCOPE_REVIEW_PATH.is_file() else None,
        },
        "capture_amendment_inputs": amendment,
        "retained_helper": {"file": str(HELPER_PATH.relative_to(HERE.parent)),
                            "sha256": sha256(HELPER_PATH)},
        "imported_hashes": imports,
        "stage_selection": [STAGE],
        "stages": stages,
        "repeat_drift": repeat_drift,
        "limitations": [
            "The profile lane has eight jobs, one measured commit dump per job, and shape-dependent lifecycle dumps.",
            "Lifecycle/setup calls are retained for scope evidence and excluded from measured commit attribution.",
            "Inclusive descendants overlap; no removable fraction, latency, allocation, cold-cache, or native-Office claim follows.",
        ],
    }


def write_output(document: dict[str, Any], output: Path) -> None:
    content = json.dumps(document, indent=2, sort_keys=True) + "\n"
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    write_exact(output, content)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--create-annotations", action="store_true",
                        help="create deterministic measured-dump annotation sidecars")
    parser.add_argument("--output", type=Path,
                        default=HERE / "profile-analysis.json",
                        help="exclusive-create JSON output destination")
    args = parser.parse_args(argv)
    try:
        document = analyze(args.create_annotations)
        write_output(document, args.output)
    except (EvidenceError, OSError, TypeError, ValueError, json.JSONDecodeError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0550 XLSX commit attribution verified: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
