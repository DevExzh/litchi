"""Verify the source-bound 0525 XLSX reconstruction evidence.

The capture campaign is intentionally staged.  Before all of the planned
artifacts exist this verifier returns a structured ``incomplete`` document;
it never turns a partial campaign into a passing comparison.  Once the
campaign is complete it replays the source patches, receipt custody, frozen
numeric/profile/hardware analyzers, and the final disposition without
assuming whether the candidate was retained or reverted.
"""

from __future__ import annotations

import argparse
import ast
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import tempfile
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
NUMERIC_PATH = HERE / "analyze.py"
PROFILE_PATH = HERE / "analyze_profiles.py"
HARDWARE_PATH = HERE / "analyze_hardware.py"
EAGER_PATH = HERE / "eager_guard.py"
EAGER_ANALYZER_PATH = HERE / "analyze_eager_guard.py"
EAGER_CONFIRMATION_PLAN_PATH = HERE / "eager-confirmation-plan.json"
EAGER_CONFIRMATION_WRAPPER_PATH = HERE / "eager_confirmation_guard.py"
EAGER_CONFIRMATION_ROOT = HERE / "eager-confirmation"
EAGER_CONFIRMATION_COMPARISON_PATH = HERE / "eager-confirmation-comparison.json"
EAGER_CONFIRMATION_REVIEW_NAMES = (
    "eager-confirmation-adverse-review.json",
    "eager-confirmation-review.json",
)
CHECKS_PATH = HERE / "checks.py"
CLOSURE_PATH = HERE / "closure_coverage.py"
CLOSURE_REPORT_PATH = HERE / "closure-coverage.json"

# These files are the frozen custody roots for this campaign.  The plan
# records the source revision, but that alone does not bind the executable
# driver or protect the evidence contract from an in-place script edit.
FROZEN_PLAN_SHA256 = "78d3fad228e4bd00476148044bbcffc8d48ec02f1e825b59c95385f92da40f97"
FROZEN_RUN_SHA256 = "1415937a8ee4602d5d8205c1b7fb98b84fd3d002d44758d3c5b8fe84a283b75d"
FROZEN_BASELINE_MANIFEST_SHA256 = "4838d157514eb9479d080d475f32921f97c30acc3ca59da43ceb1626cd9c5d03"
FROZEN_EMPTY_PATCH_SHA256 = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
FROZEN_CLOSURE_SCRIPT_SHA256 = "bc92a6956a3349db42bf98e583ba16e8a4a41d07608ede7934f1ba5e37be71cd"
FROZEN_CLOSURE_REPORT_SHA256 = "8ce112e24441e8d79ecc985a6139559320e998bdbf8aafb7671c572f99d7a7e9"

ROOTS = ("crates/litchi-xlsx/",)
SOURCE_SCOPE_EXACT = frozenset(
    {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
)
ADRS = HERE / "adr-manifest.json"
PRIOR_BINDING = HERE / "prior-source-binding.json"
VERIFICATION_REVIEW = HERE / "verification-review.md"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


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


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence bundle: {path}") from error


def is_digest(value: Any) -> bool:
    return isinstance(value, str) and bool(re.fullmatch(r"[0-9a-f]{64}", value))


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a relative path")
    path = Path(value)
    require(".." not in path.parts and path.as_posix() == value,
            f"{label} contains an unsafe path")
    return value


def plan_data() -> dict[str, Any]:
    require(sha(PLAN_PATH) == FROZEN_PLAN_SHA256,
            "plan digest differs from the frozen plan")
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("revision") == "f08daf3976714dffebe37e40bd895d87266aaf8d",
            "plan revision differs from the frozen base")
    require(plan.get("status") == "frozen-before-build-and-capture",
            "plan is not frozen before build/capture")
    require(plan.get("candidate_source_roots") == list(ROOTS),
            "candidate source roots differ from the frozen plan")
    require(plan.get("owned_paths") == [
        "/tmp/litchi-goal-0525", "/home/zhuhe/litchi-goal-0525-target"
    ], "owned temporary paths differ from the frozen plan")
    require(plan.get("cpu") == 2, "planned CPU differs")
    require(plan.get("native_order") == [
        "baseline native-r1 before candidate application",
        "candidate native-r1",
        "candidate native-r2",
        "retained baseline native-r2 under candidate source checkout",
    ], "native order differs from the frozen ABBA schedule")
    primary = plan.get("primary")
    require(isinstance(primary, dict), "primary plan is not an object")
    require(primary.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
            "primary case differs")
    require(primary.get("shapes") == ["medium", "dense-sparse"],
            "primary shape matrix differs")
    require(primary.get("repeats") == 2 and primary.get("warmup") == 20
            and primary.get("samples") == 100, "primary counts differ")
    allocation = plan.get("allocation")
    require(isinstance(allocation, dict), "allocation plan is not an object")
    require(allocation.get("shapes") == ["medium", "dense-sparse"]
            and allocation.get("repeats") == 2 and allocation.get("warmup") == 0
            and allocation.get("samples") == 5, "allocation plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "profile plan is not an object")
    require(profile.get("owner") ==
            "litchi_xlsx::cell_values::source::MultiSourceEdit::commit",
            "profile owner differs")
    require(profile.get("shapes") == ["medium", "dense-sparse"]
            and profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 1, "profile plan differs")
    hardware = plan.get("hardware")
    require(isinstance(hardware, dict), "hardware plan is not an object")
    require(hardware.get("shapes") == ["medium", "dense-sparse"]
            and hardware.get("repeats") == 2 and hardware.get("warmup") == 0
            and hardware.get("samples") == 100,
            "hardware plan differs")
    require(plan.get("guards") and plan.get("guard_repeats") == 2
            and plan.get("guard_warmup") == 10
            and plan.get("guard_samples") == 30,
            "guard plan is incomplete")
    admission = plan.get("admission")
    require(isinstance(admission, str)
            and re.search(r"at least\s+[0-9]+(?:\.[0-9]+)?%\s+p50 improvement",
                          admission, re.IGNORECASE)
            and re.search(r"at least\s+[0-9]+(?:\.[0-9]+)?%\s+commit Ir reduction",
                          admission, re.IGNORECASE),
            "admission text is malformed")
    return plan


def source_name(name: str) -> bool:
    # run.py binds the complete workspace source inventory.  The candidate
    # diff is restricted to ROOTS separately; unchanged crates remain in both
    # manifests and are part of the replay custody.
    return (name.startswith("crates/")
            or name.startswith("tools/perf-baseline/")
            or name.startswith(".cargo/")
            or name in SOURCE_SCOPE_EXACT)


def current_source_names() -> set[str]:
    tracked = subprocess.check_output([
        "git", "ls-files", "-z", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ], cwd=REPO).split(b"\0")
    names = {item.decode() for item in tracked if item}
    untracked = subprocess.check_output([
        "git", "ls-files", "--others", "--exclude-standard", "-z", "--",
        "crates", "tools/perf-baseline",
    ], cwd=REPO).split(b"\0")
    names.update(item.decode() for item in untracked if item and item.endswith(b".rs"))
    return {name for name in names if source_name(name) and (REPO / name).is_file()}


def manifest(path: Path) -> dict[str, str]:
    value = read_json(path)
    label = relative(path) if path.is_relative_to(HERE) else str(path)
    require(isinstance(value, dict) and value, f"{label} is not a nonempty manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label} entry")
        require(source_name(name), f"{label} contains out-of-scope {name}")
        require(is_digest(digest), f"{label} has malformed digest for {name}")
        require(name not in result, f"{label} repeats {name}")
        result[name] = digest
    return result


def source_names_at_revision(revision: str) -> set[str]:
    lines = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", revision], cwd=REPO, text=True
    ).splitlines()
    return {name for name in lines if source_name(name)}


def parse_index(index: Path) -> dict[str, tuple[str, int]]:
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    raw = subprocess.check_output(["git", "ls-files", "-s", "-z"], cwd=REPO, env=env)
    result: dict[str, tuple[str, int]] = {}
    for entry in raw.split(b"\0"):
        if not entry:
            continue
        metadata, encoded_name = entry.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3, "private Git index entry is malformed")
        name = encoded_name.decode()
        require(name not in result, f"private Git index repeats {name}")
        result[name] = (fields[1].decode(), int(fields[0]))
    return result


def index_blob_hashes(index: Path, oids: set[str]) -> dict[str, str]:
    if not oids:
        return {}
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    batch = subprocess.check_output(
        ["git", "cat-file", "--batch"], cwd=REPO,
        env=env, input=("\n".join(sorted(oids)) + "\n").encode(),
    )
    result: dict[str, str] = {}
    position = 0
    while position < len(batch):
        end = batch.find(b"\n", position)
        require(end >= 0, "Git batch response has no header terminator")
        oid, kind, size = batch[position:end].split()
        require(kind == b"blob", f"Git index object {oid!r} is not a blob")
        position = end + 1
        length = int(size)
        data = batch[position:position + length]
        require(len(data) == length, "Git batch response is truncated")
        result[oid.decode()] = hashlib.sha256(data).hexdigest()
        position += length
        require(position < len(batch) and batch[position:position + 1] == b"\n",
                "Git batch response lacks object separator")
        position += 1
    return result


def new_file_sidecar(stage_dir: Path) -> dict[str, dict[str, str]]:
    path = stage_dir / "new-files.json"
    if not path.exists():
        return {}
    value = read_json(path)
    raw = value.get("files") if isinstance(value, dict) else None
    require(isinstance(raw, dict), f"{relative(path)}.files is not an object")
    result: dict[str, dict[str, str]] = {}
    for name, item in raw.items():
        safe_relative(name, f"{relative(path)} file")
        require(any(name.startswith(root) for root in ROOTS),
                f"{relative(path)} file is outside candidate roots: {name}")
        if isinstance(item, str):
            digest, artifact = item, name
        else:
            require(isinstance(item, dict), f"{relative(path)} entry is not an object")
            digest = item.get("sha256")
            artifact = item.get("artifact", name)
        require(is_digest(digest), f"{relative(path)} has malformed digest for {name}")
        artifact = safe_relative(artifact, f"{relative(path)} artifact")
        artifact_path = stage_dir / artifact
        require(artifact_path.is_file() and not artifact_path.is_symlink(),
                f"{relative(path)} artifact is missing: {artifact}")
        require(sha(artifact_path) == digest,
                f"{relative(path)} artifact digest differs: {artifact}")
        result[name] = {"sha256": digest, "artifact": artifact}
    return result


def replay_source_dir(label: str, stage_dir: Path, plan: dict[str, Any],
                      candidate_like: bool) -> dict[str, Any]:
    manifest_value = manifest(stage_dir / "source-manifest.json")
    patch_path = stage_dir / "source.patch"
    require(patch_path.is_file() and not patch_path.is_symlink(),
            f"{label}/source.patch is missing")
    with tempfile.TemporaryDirectory(prefix="litchi-0525-source-replay-") as directory:
        index = Path(directory) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        subprocess.run(["git", "read-tree", plan["revision"]], cwd=REPO,
                       env=env, check=True, capture_output=True)
        patch = patch_path.read_bytes()
        if patch:
            subprocess.run(["git", "apply", "--cached", "--binary", str(patch_path)],
                           cwd=REPO, env=env, check=True, capture_output=True)
        indexed = parse_index(index)
        scope_index = {name: item for name, item in indexed.items() if source_name(name)}
        names_at_revision = source_names_at_revision(plan["revision"])
        sidecar = new_file_sidecar(stage_dir)
        missing_from_index = set(manifest_value) - set(scope_index)
        if missing_from_index:
            require(candidate_like,
                    f"{label} manifest has files absent from Git index")
            require(missing_from_index <= set(sidecar),
                    f"candidate new files lack custody sidecars: {sorted(missing_from_index)}")
        for name in set(scope_index) - set(manifest_value):
            require(name not in names_at_revision,
                    f"{label} manifest omits indexed source file {name}")
        hashes = index_blob_hashes(index, {oid for oid, _ in scope_index.values()})
        for name, digest in manifest_value.items():
            if name in scope_index:
                oid, mode = scope_index[name]
                require(mode in (100644, 100755, 120000),
                        f"{label} source file has unexpected Git mode: {name}")
                require(hashes.get(oid) == digest,
                        f"{label} replay blob differs for {name}")
            else:
                require(sidecar[name]["sha256"] == digest,
                        f"{label} new-file sidecar differs for {name}")
        # Inspect the complete cached patch.  A path-limited diff would allow
        # an out-of-scope hunk to ride along unnoticed while the manifest
        # comparison considered only the measured crates.
        changed = set(subprocess.check_output([
            "git", "diff", "--cached", "--name-only", plan["revision"]
        ], cwd=REPO, env=env, text=True).splitlines())
        require(all(source_name(name) for name in changed),
                f"{label} patch changes files outside the measured source scope")
        if candidate_like:
            require(all(any(name.startswith(root) for root in ROOTS)
                        for name in changed),
                    f"{label} patch changes files outside planned roots")
        if not candidate_like:
            require(not patch and not changed, "baseline source patch is not empty")
            require(set(manifest_value) == names_at_revision,
                    "baseline manifest file inventory differs from frozen revision")
        return {
            "stage": label,
            "manifest_sha256": sha(stage_dir / "source-manifest.json"),
            "manifest_entries": len(manifest_value),
            "patch_sha256": sha(patch_path),
            "replayed_changed_files": sorted(changed),
            "new_files": sidecar,
        }


def replay_source(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    return replay_source_dir(stage, HERE / stage, plan, stage != "baseline")


def verify_candidate_difference(plan: dict[str, Any], baseline: dict[str, Any],
                                candidate: dict[str, Any]) -> dict[str, Any]:
    before = manifest(HERE / "baseline/source-manifest.json")
    after = manifest(HERE / "candidate/source-manifest.json")
    differences = sorted(name for name in set(before) | set(after)
                         if before.get(name) != after.get(name))
    require(differences, "candidate source diff is empty")
    require(all(any(name.startswith(root) for root in ROOTS) for name in differences),
            f"candidate source diff escapes planned roots: {differences}")
    planned_files = plan.get("candidate_files", [])
    require(isinstance(planned_files, list), "plan candidate_files is malformed")
    for name in planned_files:
        safe_relative(name, "plan candidate file")
        require(any(name.startswith(root) for root in ROOTS),
                f"plan candidate file escapes planned roots: {name}")
    if planned_files:
        require(set(planned_files) == set(differences),
                "candidate source diff differs from the frozen candidate file list")
    source_diff_path = HERE / "candidate/source-diff.json"
    require(source_diff_path.is_file(), "candidate/source-diff.json is missing")
    source_diff = read_json(source_diff_path)
    require(source_diff.get("baseline_manifest_sha256") == baseline["manifest_sha256"],
            "candidate source-diff baseline binding differs")
    require(source_diff.get("candidate_manifest_sha256") == candidate["manifest_sha256"],
            "candidate source-diff candidate binding differs")
    require(source_diff.get("candidate_source_roots") == list(ROOTS),
            "candidate source-diff roots differ")
    changed = source_diff.get("changed_files")
    require(isinstance(changed, dict) and sorted(changed) == differences,
            "candidate source-diff file set differs from manifests")
    for name in differences:
        item = changed[name]
        require(isinstance(item, dict), f"candidate source-diff entry is not an object: {name}")
        require(item.get("baseline_sha256") == before.get(name)
                and item.get("candidate_sha256") == after.get(name),
                f"candidate source-diff digest differs: {name}")
    require(set(candidate["replayed_changed_files"]) == set(differences) -
            set(candidate["new_files"]),
            "candidate patch replay does not equal manifest difference")
    row_patch = HERE / "row-reuse-tests.patch"
    row_patch_info: dict[str, Any] = {"present": False}
    if row_patch.exists():
        text = read_text(row_patch)
        paths = set()
        for match in re.finditer(r"^diff --git a/(\S+) b/(\S+)$", text, re.MULTILINE):
            require(match.group(1) == match.group(2),
                    "row-reuse test patch contains a rename")
            name = match.group(1)
            require(any(name.startswith(root) for root in ROOTS),
                    f"row-reuse test patch escapes candidate roots: {name}")
            paths.add(name)
        require(paths, "row-reuse test patch has no file headers")
        row_patch_info = {"present": True, "sha256": sha(row_patch),
                          "files": sorted(paths)}
    return {"differences": differences, "source_diff": source_diff,
            "row_reuse_tests_patch": row_patch_info}


def validate_closure_coverage(plan: dict[str, Any]) -> dict[str, Any]:
    """Replay the source-derived closure count without running a build.

    The closure is a static explanation for why the candidate can omit cell
    spans.  It is deliberately kept separate from the runtime reports and
    cannot authorize a performance claim on its own.
    """

    require(CLOSURE_PATH.is_file(), "closure_coverage.py is missing")
    require(CLOSURE_REPORT_PATH.is_file(), "closure-coverage.json is missing")
    require(sha(CLOSURE_PATH) == FROZEN_CLOSURE_SCRIPT_SHA256,
            "closure coverage script differs from the frozen static translator")
    require(sha(CLOSURE_REPORT_PATH) == FROZEN_CLOSURE_REPORT_SHA256,
            "closure coverage report differs from the frozen static evidence")
    value = read_json(CLOSURE_REPORT_PATH)
    source_name_value = value.get("source") if isinstance(value, dict) else None
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0525-source-derived-closure-v1"
            and source_name_value == "tools/perf-baseline/src/lib.rs"
            and is_digest(value.get("source_sha256"))
            and value.get("source_sha256") == sha(REPO / source_name_value),
            "closure coverage source binding differs")
    baseline_value = manifest(HERE / "baseline/source-manifest.json")
    require(baseline_value.get(source_name_value) == value["source_sha256"],
            "closure coverage is not bound to the retained baseline source")
    require(value.get("owners") == ["xlsx_cell_crud_inventory", "xlsx_cell_crud_updates"],
            "closure coverage owners differ")
    results = value.get("results")
    require(isinstance(results, list) and len(results) == 2,
            "closure coverage shape matrix is incomplete")
    expected = {
        "medium": {
            "cells": 9216, "updates": 93, "touched_rows": 93,
            "cells_in_touched_rows": 4464, "row_only_skipped_cells": 4752,
            "cell_span_skipped_cells": 9123,
            "updates_by_sheet": [24, 23, 23, 23],
        },
        "dense-sparse": {
            "cells": 17792, "updates": 178, "touched_rows": 142,
            "cells_in_touched_rows": 16769, "row_only_skipped_cells": 1023,
            "cell_span_skipped_cells": 17614,
            "updates_by_sheet": [164, 11, 2, 1],
        },
    }
    seen: set[str] = set()
    for row in results:
        require(isinstance(row, dict) and row.get("shape") in expected,
                "closure coverage has an unexpected shape")
        shape = row["shape"]
        require(shape not in seen, f"closure coverage repeats shape: {shape}")
        seen.add(shape)
        for key, expected_value in expected[shape].items():
            require(row.get(key) == expected_value,
                    f"closure coverage {shape}.{key} differs")
        row_fraction = row.get("row_only_skipped_fraction")
        cell_fraction = row.get("cell_span_skipped_fraction")
        require(isinstance(row_fraction, (int, float))
                and not isinstance(row_fraction, bool)
                and isinstance(cell_fraction, (int, float))
                and not isinstance(cell_fraction, bool)
                and 0 <= row_fraction <= 1 and 0 <= cell_fraction <= 1,
                f"closure coverage {shape} fractions are invalid")
    require(seen == set(expected), "closure coverage shape set differs")
    result = subprocess.run(["python3", "-B", str(CLOSURE_PATH)], cwd=REPO,
                            text=True, capture_output=True)
    require(result.returncode == 0,
            f"closure coverage replay failed: {result.stderr[-2000:]}")
    return {
        "path": relative(CLOSURE_REPORT_PATH),
        "sha256": sha(CLOSURE_REPORT_PATH),
        "script_sha256": sha(CLOSURE_PATH),
        "source": value["source"],
        "source_sha256": value["source_sha256"],
        "shapes": sorted(seen),
        "replayed": True,
    }


def verify_static(plan: dict[str, Any]) -> dict[str, Any]:
    require(sha(RUN_PATH) == FROZEN_RUN_SHA256,
            "run.py digest differs from the frozen capture driver")
    require(ADRS.is_file(), "adr-manifest.json is missing")
    adr = read_json(ADRS)
    require(isinstance(adr, dict) and adr.get("revision") == plan["revision"],
            "ADR manifest revision differs")
    adr_files = adr.get("files")
    require(isinstance(adr_files, dict) and adr_files, "ADR manifest files are missing")
    for name, digest in adr_files.items():
        safe_relative(name, "ADR path")
        require(is_digest(digest) and (REPO / name).is_file(),
                f"ADR file is missing or malformed: {name}")
        require(sha(REPO / name) == digest, f"ADR digest differs: {name}")
    require(PRIOR_BINDING.is_file(), "prior-source-binding.json is missing")
    prior = read_json(PRIOR_BINDING)
    require(prior.get("revision") == plan["revision"]
            and prior.get("historical_manifest_sha256") ==
            sha(REPO / "docs/performance/results/change-0522/baseline/source-manifest.json")
            and prior.get("compared_xlsx_source_files") == 451
            and prior.get("mismatches") == [], "prior source binding is invalid")
    require((HERE / "mechanism-plan.md").is_file()
            and (HERE / "reconstruction-review.md").is_file()
            and (HERE / "adversarial-review.md").is_file(),
            "frozen design reviews are incomplete")
    require(VERIFICATION_REVIEW.is_file(), "verification-review.md is missing")
    review = read_text(VERIFICATION_REVIEW)
    require("0525" in review and "incomplete" in review.lower(),
            "verification review does not describe staged incomplete custody")
    baseline_dir = HERE / "baseline"
    require(baseline_dir.is_dir(), "baseline evidence directory is missing")
    baseline = replay_source("baseline", plan)
    require(baseline["manifest_sha256"] == FROZEN_BASELINE_MANIFEST_SHA256,
            "baseline source manifest differs from the frozen baseline")
    require(baseline["patch_sha256"] == FROZEN_EMPTY_PATCH_SHA256,
            "baseline source patch is not the frozen empty patch")
    closure = validate_closure_coverage(plan)
    historical = manifest(REPO / "docs/performance/results/change-0522/baseline/source-manifest.json")
    baseline_manifest = manifest(baseline_dir / "source-manifest.json")
    historical_xlsx = {name: digest for name, digest in historical.items()
                       if name.startswith("crates/litchi-xlsx/src/")}
    baseline_xlsx = {name: digest for name, digest in baseline_manifest.items()
                     if name.startswith("crates/litchi-xlsx/src/")}
    require(len(historical_xlsx) == 451 and len(baseline_xlsx) == 451
            and historical_xlsx == baseline_xlsx,
            "retained historical 451-file XLSX source binding differs")
    result: dict[str, Any] = {
        "plan_sha256": sha(PLAN_PATH),
        "run_script_sha256": sha(RUN_PATH),
        "adr_manifest_sha256": sha(ADRS),
        "prior_source_binding_sha256": sha(PRIOR_BINDING),
        "baseline": baseline,
        "closure_coverage": closure,
        "static_reviews": {
            "mechanism_plan": True,
            "reconstruction_review": True,
            "adversarial_review": True,
            "verification_review": True,
        },
    }
    candidate_dir = HERE / "candidate"
    if not candidate_dir.is_dir():
        result["candidate"] = None
        return result
    candidate = replay_source("candidate", plan)
    result["candidate"] = verify_candidate_difference(plan, baseline, candidate)
    result["candidate"]["replay"] = candidate
    return result


def primary_jobs(plan: dict[str, Any]) -> list[str]:
    primary = plan["primary"]
    names = []
    for repeat in range(1, int(primary["repeats"]) + 1):
        for shape in primary["shapes"]:
            names.append(f"native-r{repeat}-primary-{shape}")
    return names


def native_jobs(plan: dict[str, Any]) -> list[str]:
    names = primary_jobs(plan)
    for repeat in range(1, int(plan["guard_repeats"]) + 1):
        for index, guard in enumerate(plan["guards"]):
            for shape in guard["shapes"]:
                names.append(f"native-r{repeat}-guard{index}-{shape}")
    return names


def lane_jobs(plan: dict[str, Any], prefix: str, config: dict[str, Any]) -> list[str]:
    return [f"{prefix}-r{repeat}-{shape}"
            for repeat in range(1, int(config["repeats"]) + 1)
            for shape in config["shapes"]]


def capture_names(plan: dict[str, Any]) -> set[str]:
    return set(native_jobs(plan)
               + lane_jobs(plan, "alloc", plan["allocation"])
               + lane_jobs(plan, "profile", plan["profile"])
               + lane_jobs(plan, "hardware", plan["hardware"]))


def expected_receipt_names(plan: dict[str, Any]) -> set[str]:
    return {"build-normal", "build-alloc"} | capture_names(plan)


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} timestamp is not a string")
    try:
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} timestamp is invalid") from error
    require(result.tzinfo is not None, f"{label} timestamp has no timezone")
    return result


def expected_receipt_artifacts(name: str) -> set[str] | None:
    """Return the frozen run.py artifact set for a known receipt name."""

    if name.startswith("build-") or name.startswith("check-"):
        return {f"{name}.stdout", f"{name}.stderr"}
    if name.startswith("native-"):
        return {f"{name}.json", f"{name}.stdout", f"{name}.stderr",
                f"{name}.rss.json"}
    if name.startswith("alloc-"):
        return {f"{name}.json", f"{name}.stdout", f"{name}.stderr"}
    if name.startswith("profile-"):
        return ({f"{name}.json", f"{name}.stdout", f"{name}.stderr",
                 f"{name}.callgrind"}
                | {f"{name}.callgrind.{part}" for part in (1, 2, 3, 4)})
    if name.startswith("hardware-"):
        return {f"{name}.json", f"{name}.stdout", f"{name}.stderr",
                f"{name}.csv"}
    if name.startswith("eager-"):
        return {f"{name}.json", f"{name}.stdout", f"{name}.stderr",
                f"{name}.rss.json"}
    return None


def expected_receipt_binary(name: str, normal_sha: str, allocator_sha: str) -> str | None:
    if name.startswith(("native-", "profile-", "hardware-", "eager-")):
        return normal_sha
    if name.startswith("alloc-"):
        return allocator_sha
    # run.py passes no binary to build and quality-check commands.
    return None


def validate_receipt(path: Path, stage: str, plan: dict[str, Any],
                     manifest_sha: str, candidate_sha: str | None,
                     expected_binary: str | None = None,
                     check_binary: bool = False) -> dict[str, Any]:
    value = read_json(path)
    label = relative(path)
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(end > start and isinstance(seconds, (int, float))
            and not isinstance(seconds, bool) and seconds == seconds
            and seconds != float("inf") and seconds != float("-inf")
            and seconds > 0, f"{label} interval is invalid")
    exit_code = value.get("exit_code")
    require(isinstance(exit_code, int) and not isinstance(exit_code, bool),
            f"{label} exit code is invalid")
    if check_binary:
        require(value.get("binary_sha256") == expected_binary,
                f"{label} binary binding differs")
    require(value.get("source_manifest_sha256") == manifest_sha,
            f"{label} source manifest binding differs")
    require(value.get("plan_sha256") == sha(PLAN_PATH), f"{label} plan binding differs")
    require(value.get("script_sha256") == sha(RUN_PATH), f"{label} run script binding differs")
    expected_working = (candidate_sha if stage == "baseline"
                        and path.name.startswith("native-r2-") else manifest_sha)
    require(value.get("working_source_manifest_sha256") == expected_working,
            f"{label} working source binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict)
            and environment.get("TMPDIR") == str(Path(plan["owned_paths"][1]) / "test-tmp"),
            f"{label} TMPDIR binding differs")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict), f"{label} artifacts are not an object")
    for name, digest in artifacts.items():
        safe_relative(name, f"{label} artifact")
        require(Path(name).name == name,
                f"{label} artifact is not stage-local: {name}")
        artifact = path.parent / name
        require(artifact.is_file() and not artifact.is_symlink(),
                f"{label} artifact is missing: {name}")
        require(is_digest(digest) and sha(artifact) == digest,
                f"{label} artifact digest differs: {name}")
    receipt_name = path.name.removesuffix(".receipt.json")
    required_artifacts = expected_receipt_artifacts(receipt_name)
    minimum_artifacts = {f"{receipt_name}.stdout", f"{receipt_name}.stderr"}
    require(minimum_artifacts <= set(artifacts),
            f"{label} omits retained stdout/stderr artifacts")
    if exit_code == 0 and required_artifacts is not None:
        require(required_artifacts <= set(artifacts),
                f"{label} successful artifact inventory is incomplete")
    return {"path": relative(path), "start_utc": value["start_utc"],
            "end_utc": value["end_utc"], "exit_code": exit_code,
            "name": receipt_name,
            "stage": stage, "sha256": sha(path)}


def validate_build_identity(stage_dir: Path, kind: str, plan: dict[str, Any],
                            manifest_sha: str) -> dict[str, Any]:
    identity_path = stage_dir / f"binary-{kind}.json"
    value = read_json(identity_path)
    require(isinstance(value, dict) and is_digest(value.get("sha256")),
            f"{relative(identity_path)} binary identity is malformed")
    require(value.get("source_manifest_sha256") == manifest_sha,
            f"{relative(identity_path)} source binding differs")
    path = Path(value.get("path", ""))
    require(path == Path(plan["owned_paths"][0]) / f"{stage_dir.name}-{kind}",
            f"{relative(identity_path)} path differs")
    require(isinstance(value.get("bytes"), int) and value["bytes"] > 0,
            f"{relative(identity_path)} binary size is invalid")
    build_receipt = stage_dir / f"build-{kind}.receipt.json"
    require(build_receipt.is_file() and not build_receipt.is_symlink(),
            f"{relative(identity_path)} build receipt is missing")
    require(value.get("build_receipt_sha256") == sha(build_receipt),
            f"{relative(identity_path)} build receipt binding differs")
    if path.exists():
        require(path.is_file() and not path.is_symlink(),
                f"{relative(identity_path)} live binary is not regular")
        require(sha(path) == value["sha256"] and path.stat().st_size == value["bytes"],
                f"{relative(identity_path)} live binary differs")
    return {"kind": kind, "path": str(path), "sha256": value["sha256"],
            "bytes": value["bytes"],
            "build_receipt_sha256": value.get("build_receipt_sha256")}


def validate_build_command(path: Path, kind: str, plan: dict[str, Any]) -> None:
    value = read_json(path)
    label = relative(path)
    executable = ("litchi-perf-baseline-alloc" if kind == "alloc"
                  else "litchi-perf-baseline")
    expected = [
        "env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
        "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--bin", executable, "--target-dir", plan["owned_paths"][1],
    ]
    if kind == "alloc":
        expected += ["--features", "allocator-metrics"]
    require(value.get("command") == expected,
            f"{label} command differs from the frozen {kind} build")
    require(value.get("exit_code") == 0 and value.get("binary_sha256") is None,
            f"{label} build receipt is invalid")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {f"build-{kind}.stdout", f"build-{kind}.stderr"},
            f"{label} build artifact inventory differs")


def validate_stage_receipts(stage: str, plan: dict[str, Any],
                            candidate_sha: str | None) -> dict[str, Any]:
    stage_dir = HERE / stage
    stage_manifest = stage_dir / "source-manifest.json"
    manifest_sha = sha(stage_manifest)
    normal = validate_build_identity(stage_dir, "normal", plan, manifest_sha)
    allocator = validate_build_identity(stage_dir, "alloc", plan, manifest_sha)
    receipts = []
    for path in sorted(stage_dir.glob("*.receipt.json")):
        name = path.name.removesuffix(".receipt.json")
        expected_binary = expected_receipt_binary(
            name, normal["sha256"], allocator["sha256"])
        receipts.append(validate_receipt(
            path, stage, plan, manifest_sha, candidate_sha,
            expected_binary=expected_binary, check_binary=True))
    validate_build_command(stage_dir / "build-normal.receipt.json", "normal", plan)
    validate_build_command(stage_dir / "build-alloc.receipt.json", "alloc", plan)
    names = {row["name"] for row in receipts}
    required = expected_receipt_names(plan)
    require(required <= names, f"{stage} receipt set is incomplete: {sorted(required - names)}")
    require({name for name in names if name in required} == required,
            f"{stage} required receipt names differ")
    failures = [row for row in receipts if row["exit_code"] != 0]
    for row in receipts:
        if row["name"] in required and not row["name"].startswith("hardware-"):
            require(row["exit_code"] == 0,
                    f"required {stage} receipt failed: {row['name']}")
    return {"stage": stage, "manifest_sha256": manifest_sha,
            "binary_identities": {"normal": normal, "alloc": allocator},
            "receipts": receipts, "receipt_count": len(receipts),
            "failed_receipts": failures,
            "failed_extra_receipts": [row for row in failures
                                      if row["name"] not in required]}


def preflight_dirs() -> list[Path]:
    """Return retained preflight directories in deterministic order."""

    return sorted(
        path for path in HERE.iterdir()
        if path.is_dir() and (path.name == "preflight" or path.name.startswith("preflight-"))
    )


def validate_preflights(plan: dict[str, Any], candidate_sha: str,
                        static: dict[str, Any]) -> list[dict[str, Any]]:
    """Validate optional preflight source/receipt custody.

    A preflight is deliberately not a third performance stage.  It is a
    source-bound quality attempt.  A successful one may alias the canonical
    XLSX suite check only when its manifest is exactly the frozen candidate
    manifest; failed attempts remain in the serial receipt ledger.
    """

    reports: list[dict[str, Any]] = []
    for stage_dir in preflight_dirs():
        manifest_path = stage_dir / "source-manifest.json"
        patch_path = stage_dir / "source.patch"
        if not manifest_path.is_file() or not patch_path.is_file():
            raise EvidenceError(
                f"{relative(stage_dir)} preflight lacks source-manifest.json/source.patch"
            )
        replay = replay_source_dir(stage_dir.name, stage_dir, plan, True)
        manifest_sha = replay["manifest_sha256"]
        require(manifest_sha == candidate_sha or replay["replayed_changed_files"],
                f"{stage_dir.name} preflight is neither candidate-equivalent nor changed")
        receipts = []
        for path in sorted(stage_dir.glob("*.receipt.json")):
            receipts.append(validate_receipt(path, stage_dir.name, plan,
                                              manifest_sha, candidate_sha))
        reports.append({
            "stage": stage_dir.name,
            "manifest_sha256": manifest_sha,
            "candidate_manifest_equal": manifest_sha == candidate_sha,
            "replay": replay,
            "receipts": receipts,
            "receipt_count": len(receipts),
            "failed_receipts": [row["path"] for row in receipts
                                 if row["exit_code"] != 0],
        })
    return reports


def validate_global_intervals(receipts: list[dict[str, Any]]) -> None:
    ordered = sorted(receipts,
                     key=lambda row: parse_time(row["start_utc"], row["path"]))
    for left, right in zip(ordered, ordered[1:]):
        require(parse_time(left["end_utc"], left["path"])
                <= parse_time(right["start_utc"], right["path"]),
                f"serial receipt intervals overlap: {left['path']} and {right['path']}")


def validate_native_order(plan: dict[str, Any], receipts: list[dict[str, Any]]) -> dict[str, Any]:
    """Check the planned A1/B1/B2/A2 native blocks from receipt timestamps."""

    planned = set(native_jobs(plan))
    planned_by_repeat = {
        repeat: {name for name in planned if name.startswith(f"native-r{repeat}-")}
        for repeat in (1, 2)
    }
    rows = [row for row in receipts if row["name"] in planned
            and row["name"].startswith("native-")]
    require(len(rows) == 2 * len(planned),
            "native receipt matrix is incomplete for the serial order check")
    ordered = sorted(rows,
                     key=lambda row: parse_time(row["start_utc"], row["path"]))
    blocks: list[tuple[str, int]] = []
    block_names: dict[tuple[str, int], set[str]] = {}
    for row in ordered:
        match = re.match(r"^native-r([12])-", row["name"])
        require(match is not None, f"native receipt name is malformed: {row['name']}")
        block = (row["stage"], int(match.group(1)))
        if not blocks or blocks[-1] != block:
            blocks.append(block)
        block_names.setdefault(block, set()).add(row["name"])
    expected_blocks = [("baseline", 1), ("candidate", 1),
                       ("candidate", 2), ("baseline", 2)]
    require(blocks == expected_blocks,
            f"native receipt blocks differ from frozen ABBA order: {blocks}")
    for block in expected_blocks:
        require(block_names.get(block) == planned_by_repeat[block[1]],
                f"native block {block} does not contain the complete planned matrix")
    return {
        "blocks": [f"{stage}/native-r{repeat}" for stage, repeat in blocks],
        "planned_jobs_per_block": len(planned_by_repeat[1]),
        "retained_baseline_a2": True,
    }


def run_replay(script: Path, arguments: list[str], expected: Path,
               label: str) -> None:
    require(script.is_file(), f"missing analyzer {script}")
    with tempfile.TemporaryDirectory(prefix="litchi-0525-report-replay-") as directory:
        output = Path(directory) / expected.name
        result = subprocess.run(["python3", "-B", str(script), *arguments,
                                 str(output)], cwd=REPO, text=True,
                                capture_output=True)
        require(result.returncode == 0,
                f"{label} analyzer failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == expected.read_bytes(),
                f"{label} report does not replay byte-for-byte")


def run_eager_confirmation_replay(expected: Path) -> None:
    """Replay the frozen confirmation analyzer without replacing its output."""

    bootstrap = r'''
import importlib.util
import json
from pathlib import Path
import sys

script = Path(sys.argv[1]).resolve()
output = Path(sys.argv[2]).resolve()
spec = importlib.util.spec_from_file_location("litchi_0525_eager_confirmation_replay", script)
if spec is None or spec.loader is None:
    raise RuntimeError("cannot load confirmation wrapper")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

def retained_source_snapshot(plan, child):
    # The independent verifier already validates each retained build and
    # working manifest.  During a post-cleanup replay the checkout may be the
    # selected baseline or another retained final source, so replay the
    # wrapper against those retained manifests instead of requiring the live
    # checkout to still be the candidate.
    build_path = module.HERE / child["source_manifest"]
    working_path = module.HERE / child["working_source_manifest"]
    module.manifest_value(build_path, f"{child['label']} build")
    module.manifest_value(working_path, f"{child['label']} working")
    return {
        "build_source_manifest_sha256": module.sha(build_path),
        "working_source_manifest_sha256": module.sha(working_path),
    }

def redirected_write(path, value):
    path = Path(path)
    if path.name != "eager-confirmation-comparison.json":
        raise RuntimeError(f"unexpected confirmation analyzer write: {path}")
    output.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n",
                      encoding="utf-8")

module.source_manifest_snapshot = retained_source_snapshot
module.write_json = redirected_write
module.analyze()
'''
    with tempfile.TemporaryDirectory(prefix="litchi-0525-confirmation-replay-") as directory:
        output = Path(directory) / expected.name
        result = subprocess.run(
            ["python3", "-B", "-c", bootstrap,
             str(EAGER_CONFIRMATION_WRAPPER_PATH), str(output)],
            cwd=REPO, text=True, capture_output=True,
        )
        require(result.returncode == 0,
                f"eager confirmation analyzer failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == expected.read_bytes(),
                "eager confirmation comparison does not replay byte-for-byte")


def report_path(*names: str) -> Path | None:
    for name in names:
        path = HERE / name
        if path.is_file():
            return path
    return None


def quality_stage_and_name(row: dict[str, Any], final_stage: str) -> tuple[str, str]:
    name = row.get("name")
    safe_relative(name, "quality check name")
    stage = row.get("stage")
    if stage is None and "/" in name:
        stage = name.split("/", 1)[0]
    if stage is None:
        stage = final_stage
    require(isinstance(stage, str) and stage and "/" not in stage
            and ".." not in Path(stage).parts,
            "quality check stage is unsafe")
    # A path-qualified name is accepted for handoff records while retaining a
    # single stage-local receipt basename in the normal schema.
    if "/" in name:
        prefix, basename = name.split("/", 1)
        require(prefix == stage, "quality check path and stage disagree")
        name = basename
    require(name.endswith(".receipt.json") and name.startswith("check-"),
            f"quality check name is unexpected: {name}")
    safe_relative(name, "quality check receipt name")
    return stage, name


def validate_quality(plan: dict[str, Any], final_stage: str,
                     receipts: list[dict[str, Any]],
                     preflights: list[dict[str, Any]]) -> dict[str, Any]:
    path = HERE / "quality-summary.json"
    value = read_json(path)
    require(value.get("status") == "pass"
            and value.get("stage", final_stage) == final_stage,
            "quality summary status or stage differs")
    checks = value.get("checks")
    require(isinstance(checks, list) and len(checks) == 12,
            "quality summary does not contain all twelve frozen checks")
    receipt_map = {row["path"]: row for row in receipts}
    receipt_map.update({
        row["path"]: row
        for item in preflights for row in item["receipts"]
    })
    check_names = set()
    check_stages = set()
    test_count = 0
    for row in checks:
        require(isinstance(row, dict), "quality check row is not an object")
        stage, name = quality_stage_and_name(row, final_stage)
        key = f"{stage}/{name}"
        require(key not in check_names, f"quality check is repeated: {key}")
        check_names.add(key)
        check_stages.add(stage)
        receipt_path = HERE / stage / name
        require(receipt_path.is_file(), f"quality receipt is missing: {name}")
        receipt = read_json(receipt_path)
        require(sha(receipt_path) == row.get("receipt_sha256")
                and row.get("exit_code") == 0 and receipt.get("exit_code") == 0,
                f"quality receipt binding differs: {name}")
        log_path = receipt_path.with_name(name.replace(".receipt.json", ".stdout"))
        log = read_text(log_path)
        executed = sum(int(number) for number in
                       re.findall(r"test result: ok\. (\d+) passed;", log))
        require(isinstance(row.get("executed_tests"), int)
                and not isinstance(row.get("executed_tests"), bool)
                and row.get("executed_tests") == executed,
                f"quality test count differs: {name}")
        test_count += executed
        require(relative(receipt_path) in receipt_map,
                f"quality receipt is outside serialized custody: {name}")
    # A preflight alias is represented by the same command name under its
    # preflight stage.  All other checks must remain in the selected final
    # source stage.
    require({key.split("/", 1)[1] for key in check_names} ==
            {f"check-{name}.receipt.json" for name, _ in _quality_commands()},
            "quality check names differ from checks.py")
    for stage in check_stages:
        if stage != final_stage:
            require(any(item["stage"] == stage and item["candidate_manifest_equal"]
                        for item in preflights),
                    f"quality check aliases an unbound preflight: {stage}")
    require(value.get("executed_tests") == test_count,
            "quality summary aggregate test count differs")
    return {"path": relative(path), "sha256": sha(path),
            "checks": len(checks), "executed_tests": test_count,
            "stage": final_stage, "check_stages": sorted(check_stages)}


def _quality_commands() -> list[tuple[str, list[str]]]:
    require(CHECKS_PATH.is_file(), "checks.py is missing")
    try:
        tree = ast.parse(read_text(CHECKS_PATH), filename=str(CHECKS_PATH))
    except SyntaxError as error:
        raise EvidenceError(f"checks.py is not valid Python: {error}") from error
    value: Any = None
    for node in tree.body:
        targets: list[ast.expr] = []
        if isinstance(node, ast.Assign):
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            targets = [node.target]
        if any(isinstance(target, ast.Name) and target.id == "COMMANDS"
               for target in targets):
            try:
                value = ast.literal_eval(node.value)
            except (ValueError, SyntaxError) as error:
                raise EvidenceError("checks.py COMMANDS is not a literal matrix") from error
            break
    require(isinstance(value, list) and len(value) == 12,
            "checks.py command list differs from the frozen quality matrix")
    result: list[tuple[str, list[str]]] = []
    for item in value:
        require(isinstance(item, (tuple, list)) and len(item) == 2
                and isinstance(item[0], str) and item[0]
                and isinstance(item[1], list)
                and all(isinstance(token, str) and token for token in item[1]),
                "checks.py command row is malformed")
        result.append((item[0], item[1]))
    require(len({name for name, _ in result}) == len(result),
            "checks.py command names repeat")
    return result


def validate_check_receipts(plan: dict[str, Any], final_stage: str,
                            stage_report: dict[str, Any],
                            preflights: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    quality = read_json(HERE / "quality-summary.json")
    rows = quality.get("checks")
    require(isinstance(rows, list), "quality summary checks are not a list")
    preflight_by_stage = {item["stage"]: item for item in preflights}
    for name, command in _quality_commands():
        receipt_name = f"check-{name}.receipt.json"
        matches = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            stage, row_name = quality_stage_and_name(row, final_stage)
            if row_name == receipt_name:
                matches.append((stage, row, row_name))
        require(len(matches) == 1,
                f"quality summary has {len(matches)} rows for {receipt_name}")
        stage, row, _ = matches[0]
        if stage == final_stage:
            manifest_sha = stage_report["manifest_sha256"]
        else:
            require(stage in preflight_by_stage,
                    f"quality check stage is unknown: {stage}")
            item = preflight_by_stage[stage]
            require(item["candidate_manifest_equal"],
                    f"quality check stage is not candidate-equivalent: {stage}")
            manifest_sha = item["manifest_sha256"]
        path = HERE / stage / receipt_name
        value = read_json(path)
        require(value.get("exit_code") == 0, f"quality check failed: {relative(path)}")
        expected = ["env", "CARGO_TARGET_DIR=" + plan["owned_paths"][1],
                    "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0",
                    "RUSTDOCFLAGS=-D warnings"] + command
        require(value.get("command") == expected,
                f"quality command differs: {relative(path)}")
        require(value.get("source_manifest_sha256") == manifest_sha,
                f"quality source binding differs: {relative(path)}")
        result.append(validate_receipt(path, stage, plan, manifest_sha, None))
    return result


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    path = HERE / "cleanup.json"
    value = read_json(path)
    require(value.get("plan_sha256") == sha(PLAN_PATH),
            "cleanup receipt plan binding differs")
    require(value.get("removed") == plan["owned_paths"],
            "cleanup receipt does not enumerate the exact owned paths")
    require(value.get("accessible_process_references") == [],
            "cleanup receipt retains accessible process references")
    require(value.get("owned_paths_absent") is True
            and value.get("python_cache_absent") is True,
            "cleanup receipt does not prove owned cleanup")
    require(all(not Path(name).exists() for name in plan["owned_paths"]),
            "owned temporary path remains present")
    require(not list(HERE.rglob("__pycache__")),
            "evidence bundle retains a Python cache")
    return {"path": relative(path), "sha256": sha(path),
            "owned_paths_absent": True, "python_cache_absent": True}


def eager_adverse_review_source() -> tuple[Path, dict[str, Any]]:
    separate = HERE / "eager-adverse-review.json"
    if separate.is_file():
        value = read_json(separate)
        require(isinstance(value, dict),
                "eager-adverse-review.json is not an object")
        return separate, value
    main = report_path("adverse-review.json", "flag-review.json",
                       "adverse-flags-review.json")
    require(main is not None, "eager adverse flag review is missing")
    value = read_json(main)
    embedded = value.get("eager_adverse_review", value.get("eager"))
    require(isinstance(embedded, dict),
            "adverse review does not contain an eager review section")
    return main, embedded


def validate_eager_adverse_review(eager_comparison: dict[str, Any]) -> dict[str, Any]:
    require(isinstance(eager_comparison, dict),
            "eager guard comparison is not an object")
    path, value = eager_adverse_review_source()
    comparison_path = HERE / "eager-guard-comparison.json"
    require(comparison_path.is_file(), "eager guard comparison is missing")
    require(value.get("comparison_sha256") == sha(comparison_path),
            "eager adverse review comparison binding differs")
    comparison_obj = eager_comparison.get("comparison", eager_comparison)
    adverse_flags = comparison_obj.get("adverse_flags_over_five_percent", [])
    same_build_flags = comparison_obj.get("same_build_drift_over_five_percent", [])
    require(isinstance(adverse_flags, list) and isinstance(same_build_flags, list),
            "eager comparison adverse flag vectors are invalid")
    expected_count = len(adverse_flags) + len(same_build_flags)

    def rows_with_reviews(rows: Any, label: str) -> list[dict[str, Any]]:
        require(isinstance(rows, list), f"eager adverse review {label} rows are missing")
        result: list[dict[str, Any]] = []
        for row in rows:
            require(isinstance(row, dict)
                    and isinstance(row.get("review"), str)
                    and row["review"].strip(),
                    f"an eager {label} row has no review")
            result.append(row)
        return result

    if isinstance(value.get("matched"), list):
        reviewed_adverse = rows_with_reviews(value["matched"], "matched")
        reviewed_same_build = rows_with_reviews(value.get("same_build"), "same-build")
    elif isinstance(value.get("flags"), list):
        reviewed_adverse = rows_with_reviews(value["flags"], "flag")
        same_raw = value.get("same_build")
        if same_raw is None and len(reviewed_adverse) == expected_count:
            reviewed_adverse, reviewed_same_build = (
                reviewed_adverse[:len(adverse_flags)],
                reviewed_adverse[len(adverse_flags):],
            )
        else:
            reviewed_same_build = rows_with_reviews(same_raw, "same-build")
    else:
        combined = rows_with_reviews(value.get("reviews"), "combined")
        require(len(combined) == expected_count,
                "eager adverse review combined row count differs from comparison")
        reviewed_adverse = combined[:len(adverse_flags)]
        reviewed_same_build = combined[len(adverse_flags):]

    def match_expected(expected: list[Any], actual: list[dict[str, Any]], label: str) -> None:
        require(len(actual) == len(expected),
                f"eager adverse review {label} count differs from comparison")
        remaining = list(actual)
        for source_row in expected:
            require(isinstance(source_row, dict),
                    f"eager comparison {label} row is not an object")
            match_index = next(
                (index for index, reviewed in enumerate(remaining)
                 if all(reviewed.get(key) == item
                        for key, item in source_row.items())),
                None,
            )
            require(match_index is not None,
                    f"eager adverse review {label} row differs from comparison")
            remaining.pop(match_index)
        require(not remaining,
                f"eager adverse review {label} has an unbound row")

    match_expected(adverse_flags, reviewed_adverse, "matched")
    match_expected(same_build_flags, reviewed_same_build, "same-build")
    require(value.get("all_adverse_metrics_retained") is True,
            "eager adverse review does not retain all adverse metrics")
    require(value.get("no_primary_gain_claim") is True,
            "eager adverse review makes a primary gain claim")
    require(value.get("runtime_comparison_performed") is True,
            "eager adverse review lacks runtime comparison binding")
    phase_metrics = value.get("phase_metrics")
    require(isinstance(phase_metrics, dict)
            and phase_metrics.get("status") == "not_applicable",
            "eager adverse review phase scope is not not_applicable")
    require(isinstance(value.get("review"), str) and value["review"].strip(),
            "eager adverse review has no summary review")
    if "adverse_flag_count" in value:
        require(value["adverse_flag_count"] == len(adverse_flags),
                "eager adverse review count differs from comparison")
    if "same_build_drift_count" in value:
        require(value["same_build_drift_count"] == len(same_build_flags),
                "eager adverse review same-build count differs from comparison")
    if "complete" in value:
        require(value["complete"] is True, "eager adverse review is not complete")
    return {"path": relative(path), "sha256": sha(path),
            "flags": len(reviewed_adverse),
            "same_build_flags": len(reviewed_same_build),
            "comparison_flags": expected_count,
            "comparison_sha256": sha(comparison_path)}


def validate_adverse_review(comparison: dict[str, Any],
                            eager_comparison: dict[str, Any] | None = None) -> dict[str, Any]:
    path = report_path("adverse-review.json", "flag-review.json",
                       "adverse-flags-review.json")
    require(path is not None, "adverse flag review is missing")
    value = read_json(path)
    require(value.get("comparison_sha256") == sha(HERE / "comparison.json"),
            "adverse review comparison binding differs")
    comparison_obj = comparison.get("comparison", comparison)
    adverse_flags = comparison_obj.get("adverse_flags_over_five_percent", [])
    same_build_flags = comparison_obj.get("same_build_drift_over_five_percent", [])
    require(isinstance(adverse_flags, list) and isinstance(same_build_flags, list),
            "comparison adverse flag vectors are invalid")
    expected_count = len(adverse_flags) + len(same_build_flags)

    def rows_with_reviews(rows: Any, label: str) -> list[dict[str, Any]]:
        require(isinstance(rows, list), f"adverse review {label} rows are missing")
        result: list[dict[str, Any]] = []
        for row in rows:
            require(isinstance(row, dict)
                    and isinstance(row.get("review"), str)
                    and row["review"].strip(),
                    f"an adverse {label} row has no review")
            result.append(row)
        return result

    # Accept the two established layouts (matched/same_build or
    # flags/same_build), plus a single combined reviews vector.  In every
    # layout, compare every source row's complete payload before allowing
    # extra explanatory fields such as threshold or interpretation.
    if isinstance(value.get("matched"), list):
        reviewed_adverse = rows_with_reviews(value["matched"], "matched")
        reviewed_same_build = rows_with_reviews(value.get("same_build"), "same-build")
    elif isinstance(value.get("flags"), list):
        reviewed_adverse = rows_with_reviews(value["flags"], "flag")
        same_raw = value.get("same_build")
        if same_raw is None and len(reviewed_adverse) == expected_count:
            reviewed_adverse, reviewed_same_build = (
                reviewed_adverse[:len(adverse_flags)],
                reviewed_adverse[len(adverse_flags):],
            )
        else:
            reviewed_same_build = rows_with_reviews(same_raw, "same-build")
    else:
        combined = rows_with_reviews(value.get("reviews"), "combined")
        require(len(combined) == expected_count,
                "adverse review combined row count differs from comparison")
        reviewed_adverse = combined[:len(adverse_flags)]
        reviewed_same_build = combined[len(adverse_flags):]

    def match_expected(expected: list[Any], actual: list[dict[str, Any]], label: str) -> None:
        require(len(actual) == len(expected),
                f"adverse review {label} count differs from comparison")
        remaining = list(actual)
        for source_row in expected:
            require(isinstance(source_row, dict),
                    f"comparison {label} row is not an object")
            match_index = next(
                (index for index, reviewed in enumerate(remaining)
                 if all(reviewed.get(key) == item
                        for key, item in source_row.items())),
                None,
            )
            require(match_index is not None,
                    f"adverse review {label} row differs from comparison")
            remaining.pop(match_index)
        require(not remaining,
                f"adverse review {label} has an unbound row")

    match_expected(adverse_flags, reviewed_adverse, "matched")
    match_expected(same_build_flags, reviewed_same_build, "same-build")
    if "adverse_flag_count" in value:
        require(value["adverse_flag_count"] == len(adverse_flags),
                "adverse review count differs from comparison")
    if "same_build_drift_count" in value:
        require(value["same_build_drift_count"] == len(same_build_flags),
                "adverse review same-build count differs from comparison")
    if "complete" in value:
        require(value["complete"] is True, "adverse review is not complete")
    eager = validate_eager_adverse_review(eager_comparison)
    return {"path": relative(path), "sha256": sha(path),
            "flags": len(reviewed_adverse),
            "same_build_flags": len(reviewed_same_build),
            "comparison_flags": expected_count, "eager": eager}


def validate_disposition(plan: dict[str, Any], static: dict[str, Any],
                         comparison: dict[str, Any], profile: dict[str, Any],
                         adverse: dict[str, Any]) -> dict[str, Any]:
    paths = [HERE / "decision.json", HERE / "disposition.json"]
    existing = [path for path in paths if path.is_file()]
    require(len(existing) == 1, "exactly one decision/disposition record is required")
    path = existing[0]
    value = read_json(path)
    require(isinstance(value, dict), f"{relative(path)} is not an object")
    disposition = value.get("disposition")
    require(disposition in ("accepted", "rejected", "rejected_and_reverted"),
            "disposition is not accepted/rejected")
    final_source = value.get("final_source")
    require(final_source in ("baseline", "candidate", "final"),
            "final source stage is missing or unexpected")
    require(isinstance(value.get("production_change_retained"), bool),
            "production_change_retained must be an explicit boolean")
    candidate_manifest_sha = sha(HERE / "candidate/source-manifest.json")
    baseline_manifest_sha = sha(HERE / "baseline/source-manifest.json")
    final_manifest_path = HERE / "final/source-manifest.json"
    final_manifest_sha = sha(final_manifest_path) if final_manifest_path.is_file() else None
    requested_sha = value.get("final_source_manifest_sha256")
    require(is_digest(requested_sha), "final source manifest binding is missing")
    candidates = {"baseline": baseline_manifest_sha, "candidate": candidate_manifest_sha}
    if final_manifest_sha is not None:
        candidates["final"] = final_manifest_sha
    matching = [name for name, digest in candidates.items() if digest == requested_sha]
    require(matching, "final source manifest is not retained in the evidence bundle")
    require(final_source in matching, "decision final_source disagrees with manifest hash")
    selected = final_source
    final_replay = None
    if final_manifest_sha is not None:
        require(final_manifest_path.parent.joinpath("source.patch").is_file(),
                "final source manifest lacks source.patch custody")
        final_replay = replay_source_dir("final", final_manifest_path.parent,
                                         plan, True)
        require(final_replay["manifest_sha256"] == final_manifest_sha,
                "final source replay manifest differs")
    selected_manifest = (HERE / selected / "source-manifest.json")
    require(selected_manifest.is_file(), "selected final source manifest is missing")
    current = {name: sha(REPO / name) for name in current_source_names()}
    require(current == manifest(selected_manifest),
            "current checkout does not match the selected final source manifest")
    if value["production_change_retained"] is True:
        require(selected in ("candidate", "final"),
                "retained production decision points to a baseline source")
    if disposition == "accepted":
        require(value["production_change_retained"] is True,
                "accepted disposition says production change was not retained")
        require(comparison["native_admission"]["passed"],
                "accepted disposition lacks the frozen native gate")
        require(profile["comparison"]["profile_commit_ir_admission"]["passed"],
                "accepted disposition lacks the frozen profile gate")
        require(value.get("eager_confirmation_gate") is True,
                "accepted disposition lacks an explicit supplemental eager confirmation gate")
        confirmation = adverse.get("eager_confirmation")
        require(isinstance(confirmation, dict)
                and confirmation.get("gate_passed") is True,
                "accepted disposition lacks the supplemental eager confirmation gate")
    review_path = report_path("adverse-review.json", "flag-review.json",
                              "adverse-flags-review.json")
    require(review_path is not None, "adverse review is missing")
    eager_review = adverse.get("eager")
    require(isinstance(eager_review, dict),
            "eager adverse review is missing from the validated review set")
    eager_comparison_path = HERE / "eager-guard-comparison.json"
    require(eager_comparison_path.is_file(), "eager guard comparison is missing")
    eager_review_path = HERE / eager_review["path"]
    require(eager_review_path.is_file(), "eager adverse review artifact is missing")
    confirmation_review = adverse.get("eager_confirmation")
    require(isinstance(confirmation_review, dict),
            "eager confirmation adverse review is missing from the validated review set")
    if "eager_confirmation_gate" in value:
        require(isinstance(value["eager_confirmation_gate"], bool)
                and value["eager_confirmation_gate"] == confirmation_review.get("gate_passed"),
                f"{relative(path)} eager confirmation gate differs")
    if disposition == "accepted":
        require(confirmation_review.get("gate_passed") is True,
                "accepted disposition lacks a passing supplemental eager confirmation gate")
    confirmation_comparison_path = EAGER_CONFIRMATION_COMPARISON_PATH
    require(confirmation_comparison_path.is_file(),
            "eager confirmation comparison is missing")
    confirmation_review_path = HERE / confirmation_review["path"]
    require(confirmation_review_path.is_file(),
            "eager confirmation adverse review artifact is missing")
    for key, artifact in {
        "plan_sha256": PLAN_PATH,
        "comparison_sha256": HERE / "comparison.json",
        "profile_analysis_sha256": HERE / "profile-analysis.json",
        "quality_summary_sha256": HERE / "quality-summary.json",
        "adverse_review_sha256": review_path,
        "eager_comparison_sha256": eager_comparison_path,
        "eager_adverse_review_sha256": eager_review_path,
        "eager_confirmation_comparison_sha256": confirmation_comparison_path,
        "eager_confirmation_adverse_review_sha256": confirmation_review_path,
    }.items():
        require(value.get(key) == sha(artifact),
                f"{relative(path)} {key} binding differs")
    if "native_primary_gate" in value:
        require(value["native_primary_gate"] ==
                comparison["native_admission"]["passed"],
                f"{relative(path)} native gate differs")
    if "profile_gate" in value:
        require(value["profile_gate"] ==
                profile["comparison"]["profile_commit_ir_admission"]["passed"],
                f"{relative(path)} profile gate differs")
    if "profile_comparison_sha256" in value:
        artifact = HERE / "profile-comparison.json"
        require(artifact.is_file() and value["profile_comparison_sha256"] == sha(artifact),
                f"{relative(path)} profile_comparison_sha256 binding differs")
    profile_hashes = value.get("profiles_sha256")
    if profile_hashes is not None:
        require(isinstance(profile_hashes, dict),
                f"{relative(path)} profiles_sha256 is not an object")
        for stage, digest in profile_hashes.items():
            require(stage in ("baseline", "candidate"),
                    f"{relative(path)} profile stage is unexpected: {stage}")
            artifact = HERE / stage / "profile-analysis.json"
            if not artifact.is_file():
                artifact = HERE / "profile-analysis.json"
            require(is_digest(digest) and artifact.is_file() and sha(artifact) == digest,
                    f"{relative(path)} profile artifact binding differs: {stage}")
    return {"path": relative(path), "sha256": sha(path),
            "disposition": disposition, "final_source": selected,
            "final_source_manifest_sha256": requested_sha,
            "final_replay": final_replay}


def validate_reports(plan: dict[str, Any], stage_reports: dict[str, Any]) -> dict[str, Any]:
    # The numerical and profile wrappers are the frozen report authorities.
    run_replay(NUMERIC_PATH, [], HERE / "comparison.json", "numeric comparison")
    run_replay(PROFILE_PATH, [], HERE / "profile-analysis.json", "profile comparison")
    profile = read_json(HERE / "profile-analysis.json")
    require(profile.get("status") == "pass"
            and profile.get("comparison") is not None,
            "profile analysis is not a complete comparison")
    profile_gate = profile["comparison"].get("profile_commit_ir_admission")
    require(isinstance(profile_gate, dict) and isinstance(profile_gate.get("passed"), bool),
            "profile admission gate is missing")
    for stage in ("baseline", "candidate"):
        expected = HERE / stage / "hardware-analysis.json"
        run_replay(HARDWARE_PATH, ["--stage", stage], expected,
                   f"{stage} hardware diagnostic")
    comparison = read_json(HERE / "comparison.json")
    require(comparison.get("status") == "pass" and
            comparison.get("native_admission") is not None,
            "numeric comparison is not complete")
    native_gate = comparison["native_admission"]
    require(isinstance(native_gate.get("passed"), bool),
            "native admission gate is missing")
    rows = native_gate.get("rows")
    require(isinstance(rows, list) and len(rows) == 4,
            "primary native admission matrix is incomplete")
    for row in rows:
        require(row.get("passed") is
                (row["native_primary_total_p50"]["passed"] and
                 row["native_primary_commit_p50"]["passed"]),
                "native admission row does not reconcile")
    eager_compare = HERE / "eager-guard-comparison.json"
    eager_report: dict[str, Any] | None = None
    if EAGER_ANALYZER_PATH.is_file():
        # eager_guard.py remains the capture authority.  Its generic analyzer
        # expects source phase vectors that the eager report intentionally
        # omits; the companion analyzer owns this report schema and replay.
        run_replay(EAGER_ANALYZER_PATH, ["compare", "analyze", "--output"],
                   eager_compare, "eager guard comparison")
        eager_report = read_json(eager_compare)
        require(eager_report.get("comparison", {}).get("all_adverse_metrics_retained") is True
                and eager_report.get("comparison", {}).get("no_primary_gain_claim") is True,
                "eager guard comparison omits its diagnostic boundaries")
        for stage in ("baseline", "candidate"):
            expected = HERE / stage / "eager-guard-analysis.json"
            run_replay(EAGER_ANALYZER_PATH, [stage, "analyze", "--output"], expected,
                       f"{stage} eager guard analysis")
    hardware = {stage: read_json(HERE / stage / "hardware-analysis.json")
                for stage in ("baseline", "candidate")}
    for stage, report in hardware.items():
        require(report.get("no_operation_local_hardware_or_speedup_claim") is True,
                f"{stage} hardware report makes an operation-local claim")
    return {"comparison": comparison, "profile": profile,
            "eager": eager_report, "hardware": hardware,
            "native_gate": native_gate["passed"], "profile_gate": profile_gate["passed"]}


def validate_eager_confirmation_plan() -> dict[str, Any]:
    path = EAGER_CONFIRMATION_PLAN_PATH
    require(path.is_file() and not path.is_symlink(),
            "eager confirmation plan is missing")
    value = read_json(path)
    require(isinstance(value, dict), "eager confirmation plan is not an object")
    require(value.get("schema") == "litchi-0525-eager-confirmation-plan-v1",
            "eager confirmation plan schema differs")
    require(value.get("status") == "frozen-supplemental-before-capture",
            "eager confirmation plan is not frozen before capture")
    require(value.get("primary_plan_sha256") == sha(PLAN_PATH),
            "eager confirmation plan primary binding differs")
    primary_plan = read_json(PLAN_PATH)
    require(isinstance(primary_plan, dict)
            and value.get("revision") == primary_plan.get("revision"),
            "eager confirmation revision differs from the primary plan")
    require(value.get("capture_root") == EAGER_CONFIRMATION_ROOT.name,
            "eager confirmation capture root differs")
    require(value.get("capture_wrapper") == EAGER_CONFIRMATION_WRAPPER_PATH.name,
            "eager confirmation wrapper binding differs")
    require(value.get("case") == "xlsx_eager_cell_values_one_percent_edit_save",
            "eager confirmation case differs")
    require(value.get("shape") == "dense-sparse"
            and value.get("cpu") == 2
            and value.get("warmup") == 20
            and value.get("samples") == 100,
            "eager confirmation workload differs")
    require(value.get("order") == ["A1", "B1", "B2", "A2"],
            "eager confirmation ABBA order differs")
    require(value.get("owned_paths") == [
        "/tmp/litchi-goal-0525", "/home/zhuhe/litchi-goal-0525-target"
    ], "eager confirmation owned paths differ")
    require(value.get("tmpdir") == "/home/zhuhe/litchi-goal-0525-target/test-tmp",
            "eager confirmation TMPDIR differs")
    helpers = value.get("retained_helpers")
    require(isinstance(helpers, dict)
            and set(helpers) == {"frozen_run", "numeric", "eager_companion"},
            "eager confirmation retained helper bindings differ")
    expected_helpers = {
        "frozen_run": ("run.py", FROZEN_RUN_SHA256),
        "numeric": ("analyze.py", None),
        "eager_companion": ("analyze_eager_guard.py", None),
    }
    for name, (helper_path, expected_sha) in expected_helpers.items():
        helper = helpers[name]
        require(isinstance(helper, dict)
                and helper.get("path") == helper_path
                and is_digest(helper.get("sha256"))
                and sha(HERE / helper_path) == helper["sha256"],
                f"eager confirmation {name} helper binding differs")
        if expected_sha is not None:
            require(helper["sha256"] == expected_sha,
                    f"eager confirmation {name} frozen helper differs")
    custody = value.get("custody")
    require(isinstance(custody, dict)
            and custody.get("fresh_child_per_capture") is True
            and custody.get("fresh_child_per_sample") is False
            and custody.get("process_isolated") is True
            and custody.get("source_manifest_and_binary_hash_before_after_child") is True
            and custody.get("receipt_artifacts") == ["json", "rss.json", "stdout", "stderr"],
            "eager confirmation custody contract differs")
    gate = value.get("gate")
    require(isinstance(gate, dict)
            and gate.get("p50_and_mean_max_adverse_change_percent") == 5.0
            and gate.get("paired_comparisons") == ["A1_vs_B1", "A2_vs_B2"]
            and gate.get("primary_gain_claim") is False
            and gate.get("requires_both_pairs") is True
            and gate.get("review_every_over_five_percent_tail_or_rss") is True,
            "eager confirmation gate differs")
    metrics = value.get("metrics")
    require(isinstance(metrics, dict)
            and metrics.get("adverse_threshold_percent") == 5.0
            and metrics.get("elapsed") == "results[0].elapsed_ns"
            and metrics.get("elapsed_statistics") == ["p50", "p95", "p99", "mean"]
            and metrics.get("retain_all_adverse") is True
            and metrics.get("rss") == "whole_child_process.max_rss_kib",
            "eager confirmation metric contract differs")
    oracle = value.get("oracle_reference")
    safe_relative(oracle, "eager confirmation oracle reference")
    oracle_path = HERE / oracle
    require(oracle_path.is_file() and value.get("oracle_reference_sha256") == sha(oracle_path),
            "eager confirmation oracle binding differs")
    oracles = value.get("oracles")
    require(isinstance(oracles, dict)
            and oracles.get("corpus") == "exactly_equal_to_oracle_reference"
            and oracles.get("identity") == "exact_corpus_sink_output_and_source_absence_across_children"
            and oracles.get("operation_metrics_phase_vectors") == "not_applicable"
            and oracles.get("operation_metrics_source_status") == "not_applicable"
            and oracles.get("output_sha256") == "exactly_equal_to_oracle_reference"
            and oracles.get("result_source_field") == "absent"
            and oracles.get("sink") == "exactly_equal_to_oracle_reference",
            "eager confirmation oracle contract differs")
    expected_children = [
        {
            "binary_identity": "baseline/binary-normal.json",
            "binary_path": "/tmp/litchi-goal-0525/baseline-normal",
            "label": "A1", "repeat": 1,
            "source_manifest": "baseline/source-manifest.json",
            "working_source_manifest": "candidate/source-manifest.json",
            "stage": "baseline",
        },
        {
            "binary_identity": "candidate/binary-normal.json",
            "binary_path": "/tmp/litchi-goal-0525/candidate-normal",
            "label": "B1", "repeat": 1,
            "source_manifest": "candidate/source-manifest.json",
            "working_source_manifest": "candidate/source-manifest.json",
            "stage": "candidate",
        },
        {
            "binary_identity": "candidate/binary-normal.json",
            "binary_path": "/tmp/litchi-goal-0525/candidate-normal",
            "label": "B2", "repeat": 2,
            "source_manifest": "candidate/source-manifest.json",
            "working_source_manifest": "candidate/source-manifest.json",
            "stage": "candidate",
        },
        {
            "binary_identity": "baseline/binary-normal.json",
            "binary_path": "/tmp/litchi-goal-0525/baseline-normal",
            "label": "A2", "repeat": 2,
            "source_manifest": "baseline/source-manifest.json",
            "working_source_manifest": "candidate/source-manifest.json",
            "stage": "baseline",
        },
    ]
    require(value.get("children") == expected_children,
            "eager confirmation child matrix differs")
    require(value.get("comparison_output") == EAGER_CONFIRMATION_COMPARISON_PATH.name,
            "eager confirmation comparison output differs")
    parse_time(value.get("frozen_at_utc"), "eager confirmation plan")
    return {"path": relative(path), "sha256": sha(path), "plan": value,
            "oracle_path": relative(oracle_path)}


def confirmation_expected_command(plan: dict[str, Any], child: dict[str, Any],
                                  output: Path, rss: Path) -> list[str]:
    return [
        "taskset", "-c", str(plan["cpu"]), "/usr/bin/time", "-f",
        '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
        "-o", str(rss), child["binary_path"], "--warmup", str(plan["warmup"]),
        "--samples", str(plan["samples"]), "--case", plan["case"],
        "--xlsx-cell-crud-shape", plan["shape"], "--json", str(output),
    ]


def validate_confirmation_elapsed(value: Any, label: str, samples: int) -> dict[str, Any]:
    require(isinstance(value, dict) and value.get("unit") == "ns",
            f"{label} elapsed statistics are malformed")
    raw = value.get("samples")
    require(isinstance(raw, list) and len(raw) == samples
            and all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                    for item in raw),
            f"{label} elapsed samples differ")
    order = value.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(samples)),
            f"{label} elapsed sample order differs")
    for name in ("min", "p50", "p95", "p99", "max", "mean",
                 "standard_deviation"):
        item = value.get(name)
        require(isinstance(item, (int, float)) and not isinstance(item, bool)
                and math.isfinite(float(item)) and float(item) >= 0,
                f"{label} elapsed {name} is invalid")
    return value


def validate_confirmation_raw(path: Path, label: str, plan: dict[str, Any],
                              child: dict[str, Any], binary_sha: str,
                              oracle: dict[str, Any]) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict) and value.get("schema_version") == 1,
            f"{label} confirmation report schema differs")
    tool = value.get("tool")
    require(isinstance(tool, dict)
            and tool.get("binary") == "litchi-perf-baseline"
            and tool.get("profile") == "release"
            and tool.get("instrumentation") == "none",
            f"{label} confirmation tool identity differs")
    identity = value.get("binary_identity")
    require(isinstance(identity, dict)
            and identity.get("binary_sha256") == binary_sha
            and identity.get("profile") == "release",
            f"{label} confirmation binary identity differs")
    environment = value.get("environment")
    primary_plan = read_json(PLAN_PATH)
    require(isinstance(environment, dict)
            and isinstance(primary_plan, dict)
            and environment.get("git_revision") == primary_plan["revision"]
            and environment.get("cpu_affinity") == str(plan["cpu"]),
            f"{label} confirmation environment differs")
    configuration = value.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("samples_per_case") == plan["samples"]
            and configuration.get("warmup_iterations_per_case") == plan["warmup"]
            and configuration.get("cases") == [plan["case"]]
            and configuration.get("xlsx_cell_crud_shapes") == [plan["shape"]]
            and configuration.get("filesystem_fresh_child_per_sample") is True
            and configuration.get("filesystem_process_isolated") is True,
            f"{label} confirmation configuration differs")
    results = value.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} confirmation result matrix differs")
    result = results[0]
    require(isinstance(result, dict)
            and result.get("case") == plan["case"]
            and "source" not in result
            and result.get("corpus") == oracle.get("corpus")
            and result.get("sink") == oracle.get("sink")
            and result.get("output_sha256") == oracle.get("output_sha256"),
            f"{label} confirmation semantic oracle differs")
    validate_confirmation_elapsed(result.get("elapsed_ns"), label, plan["samples"])
    operation = result.get("operation_metrics")
    require(isinstance(operation, dict)
            and operation.get("sample_count") == plan["samples"]
            and operation.get("sample_indices") == list(range(plan["samples"]))
            and operation.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index"
            and operation.get("latency_claim") == "comparable_timed_operation",
            f"{label} confirmation operation metrics differ")
    for name in ("source", "process", "publication", "materialization", "cfb_phases"):
        section = operation.get(name)
        require(isinstance(section, dict) and section.get("status") == "not_applicable",
                f"{label} confirmation {name} scope differs")
    sink = operation.get("sink")
    require(isinstance(sink, dict) and sink.get("status") == "not_applicable"
            and sink.get("write_status") == "measured",
            f"{label} confirmation sink scope differs")
    return result


def validate_confirmation_child(plan: dict[str, Any], child: dict[str, Any],
                                oracle: dict[str, Any]) -> dict[str, Any]:
    label = child["label"]
    folder = EAGER_CONFIRMATION_ROOT / label
    require(folder.is_dir() and not folder.is_symlink(),
            f"eager confirmation {label} folder is missing")
    expected_names = {f"{label}.json", f"{label}.rss.json",
                      f"{label}.receipt.json", f"{label}.binding.json",
                      f"{label}.stdout", f"{label}.stderr"}
    require({path.name for path in folder.iterdir()} == expected_names,
            f"eager confirmation {label} artifact set differs")
    receipt_path = folder / f"{label}.receipt.json"
    binding_path = folder / f"{label}.binding.json"
    receipt = read_json(receipt_path)
    binding = read_json(binding_path)
    require(receipt.get("schema") == "litchi-0525-eager-confirmation-receipt-v1"
            and binding.get("schema") == "litchi-0525-eager-confirmation-binding-v1",
            f"eager confirmation {label} receipt/binding schema differs")
    require(receipt.get("label") == label and binding.get("label") == label
            and receipt.get("stage") == child["stage"]
            and receipt.get("repeat") == child["repeat"]
            and binding.get("stage") == child["stage"]
            and binding.get("repeat") == child["repeat"],
            f"eager confirmation {label} child identity differs")
    require(receipt.get("source_manifest") == child["source_manifest"]
            and receipt.get("working_source_manifest") == child["working_source_manifest"],
            f"eager confirmation {label} source manifest paths differ")
    require(receipt.get("exit_code") == 0 and receipt.get("validation") == "pass"
            and binding.get("validation") == "pass",
            f"eager confirmation {label} receipt is not a validated pass")
    plan_sha = sha(PLAN_PATH)
    wrapper_sha = sha(EAGER_CONFIRMATION_WRAPPER_PATH)
    for value, prefix in ((receipt, "receipt"), (binding, "binding")):
        require(value.get("primary_plan_sha256") == plan_sha
                and value.get("supplemental_plan_sha256") == sha(EAGER_CONFIRMATION_PLAN_PATH)
                and value.get("run_script_sha256") == sha(RUN_PATH)
                and value.get("wrapper_sha256") == wrapper_sha
                and value.get("retained_helpers") == plan["retained_helpers"],
                f"eager confirmation {label} {prefix} code/plan custody differs")
    source_path = child["source_manifest"]
    manifest_file = HERE / source_path
    require(manifest_file.is_file() and source_path in ("baseline/source-manifest.json",
                                                        "candidate/source-manifest.json"),
            f"eager confirmation {label} source manifest path differs")
    manifest_sha = sha(manifest_file)
    working_source_path = child["working_source_manifest"]
    working_manifest_file = HERE / working_source_path
    require(working_manifest_file.is_file()
            and working_source_path == "candidate/source-manifest.json",
            f"eager confirmation {label} working source manifest path differs")
    working_manifest_sha = sha(working_manifest_file)
    binary_identity_path = HERE / child["binary_identity"]
    identity = read_json(binary_identity_path)
    require(isinstance(identity, dict)
            and identity.get("path") == child["binary_path"]
            and is_digest(identity.get("sha256"))
            and identity.get("source_manifest_sha256") == manifest_sha,
            f"eager confirmation {label} binary identity differs")
    binary_sha = identity["sha256"]
    binary_path = Path(child["binary_path"])
    if binary_path.exists():
        require(binary_path.is_file() and not binary_path.is_symlink()
                and sha(binary_path) == binary_sha,
                f"eager confirmation {label} live binary differs")
    require(receipt.get("source_manifest") == source_path
            and receipt.get("source_manifest_sha256") == manifest_sha
            and receipt.get("working_source_manifest_sha256") == working_manifest_sha
            and receipt.get("binary_path") == child["binary_path"]
            and receipt.get("binary_sha256") == binary_sha
            and binding.get("source_manifest_sha256") == manifest_sha
            and binding.get("working_source_manifest_sha256") == working_manifest_sha
            and binding.get("binary_sha256") == binary_sha
            and binding.get("receipt_sha256") == sha(receipt_path),
            f"eager confirmation {label} source/binary binding differs")
    output_path = folder / f"{label}.json"
    rss_path = folder / f"{label}.rss.json"
    require(receipt.get("command") == confirmation_expected_command(
        plan, child, output_path, rss_path),
        f"eager confirmation {label} command differs")
    expected_artifacts = {
        relative(output_path), relative(rss_path),
        relative(folder / f"{label}.stdout"), relative(folder / f"{label}.stderr"),
    }
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts
            and binding.get("artifacts") == artifacts,
            f"eager confirmation {label} artifact inventory differs")
    for artifact_name, digest in artifacts.items():
        artifact_path = HERE / artifact_name
        require(is_digest(digest) and artifact_path.is_file()
                and not artifact_path.is_symlink() and sha(artifact_path) == digest,
                f"eager confirmation {label} artifact digest differs: {artifact_name}")
    environment = receipt.get("environment")
    require(isinstance(environment, dict) and environment.get("TMPDIR") == plan["tmpdir"],
            f"eager confirmation {label} TMPDIR binding differs")
    raw = validate_confirmation_raw(output_path, label, plan, child, binary_sha, oracle)
    rss = read_json(rss_path)
    require(isinstance(rss, dict)
            and set(rss) == {"max_rss_kib", "elapsed_seconds", "user_seconds", "system_seconds"}
            and isinstance(rss["max_rss_kib"], int) and rss["max_rss_kib"] >= 0
            and all(isinstance(rss[name], (int, float)) and not isinstance(rss[name], bool)
                    and math.isfinite(float(rss[name])) and rss[name] >= 0
                    for name in ("elapsed_seconds", "user_seconds", "system_seconds")),
            f"eager confirmation {label} RSS report differs")
    return {
        "label": label, "stage": child["stage"], "repeat": child["repeat"],
        "report_sha256": sha(output_path), "receipt_sha256": sha(receipt_path),
        "binding_sha256": sha(binding_path), "source_manifest_sha256": manifest_sha,
        "binary_sha256": binary_sha, "elapsed_ns": raw["elapsed_ns"], "rss": rss,
        "working_source_manifest_sha256": working_manifest_sha,
        "corpus": raw["corpus"], "sink": raw["sink"],
        "output_sha256": raw["output_sha256"],
        "start_utc": receipt["start_utc"], "end_utc": receipt["end_utc"],
    }


def validate_eager_confirmation_capture(plan_info: dict[str, Any]) -> dict[str, Any]:
    plan = plan_info["plan"]
    root = EAGER_CONFIRMATION_ROOT
    require(root.is_dir() and not root.is_symlink(),
            "eager confirmation capture root is missing")
    labels = plan["order"]
    actual_dirs = {path.name for path in root.iterdir() if path.is_dir()}
    require(actual_dirs == set(labels),
            "eager confirmation child directory set differs")
    state_path = root / "capture-state.json"
    state = read_json(state_path)
    require(isinstance(state, dict)
            and state.get("schema") == "litchi-0525-eager-confirmation-state-v1"
            and state.get("plan_sha256") == plan_info["sha256"]
            and state.get("order") == labels
            and state.get("completed") == labels
            and state.get("status") == "complete"
            and isinstance(state.get("entries"), dict)
            and set(state["entries"]) == set(labels)
            and "failed_label" not in state,
            "eager confirmation capture state differs")
    oracle = read_json(HERE / plan["oracle_reference"])["results"][0]
    rows = []
    for label in labels:
        child = next(item for item in plan["children"] if item["label"] == label)
        row = validate_confirmation_child(plan, child, oracle)
        entry = state["entries"][label]
        require(entry == {
            "receipt": f"{root.name}/{label}/{label}.receipt.json",
            "receipt_sha256": row["receipt_sha256"],
            "binding": f"{root.name}/{label}/{label}.binding.json",
            "binding_sha256": row["binding_sha256"],
            "stage": child["stage"], "repeat": child["repeat"],
        }, f"eager confirmation {label} state entry differs")
        rows.append(row)
    ordered = sorted(rows, key=lambda row: parse_time(row["start_utc"], row["label"]))
    require([row["label"] for row in ordered] == labels,
            "eager confirmation receipt order differs from ABBA plan")
    for left, right in zip(ordered, ordered[1:]):
        require(parse_time(left["end_utc"], left["label"]) <=
                parse_time(right["start_utc"], right["label"]),
                f"eager confirmation receipt intervals overlap: {left['label']} and {right['label']}")
    return {"path": root.name, "state_sha256": sha(state_path), "rows": rows,
            "order": labels}


def percent_change(left: Any, right: Any, label: str) -> float:
    require(isinstance(left, (int, float)) and not isinstance(left, bool)
            and isinstance(right, (int, float)) and not isinstance(right, bool)
            and math.isfinite(float(left)) and math.isfinite(float(right))
            and float(left) > 0,
            f"{label} comparison values are invalid")
    return (float(right) / float(left) - 1.0) * 100.0


def validate_eager_confirmation_comparison(plan_info: dict[str, Any],
                                           capture: dict[str, Any]) -> dict[str, Any]:
    plan = plan_info["plan"]
    path = EAGER_CONFIRMATION_COMPARISON_PATH
    require(path.is_file() and not path.is_symlink(),
            "eager confirmation comparison is missing")
    value = read_json(path)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0525-eager-confirmation-comparison-v1"
            and value.get("status") == "pass"
            and value.get("stage") == "compare"
            and value.get("supplemental_plan_sha256") == plan_info["sha256"]
            and value.get("primary_plan_sha256") == sha(PLAN_PATH),
            "eager confirmation comparison envelope differs")
    wrapper = value.get("capture_wrapper")
    require(isinstance(wrapper, dict)
            and wrapper.get("path") == EAGER_CONFIRMATION_WRAPPER_PATH.name
            and wrapper.get("sha256") == sha(EAGER_CONFIRMATION_WRAPPER_PATH),
            "eager confirmation comparison wrapper binding differs")
    require(value.get("order") == plan["order"],
            "eager confirmation comparison order differs")
    children = value.get("children")
    rows = capture["rows"]
    require(isinstance(children, list)
            and [row.get("label") for row in children] == plan["order"],
            "eager confirmation comparison child rows differ")
    by_label = {row["label"]: row for row in rows}
    for child in children:
        expected = by_label[child["label"]]
        for key in ("label", "stage", "repeat", "report_sha256", "receipt_sha256",
                    "binding_sha256", "source_manifest_sha256",
                    "working_source_manifest_sha256", "binary_sha256",
                    "output_sha256", "corpus", "sink"):
            require(child.get(key) == expected[key],
                    f"eager confirmation child {child['label']} {key} differs")
    pairs = value.get("matched_pairs")
    require(isinstance(pairs, list)
            and [row.get("pair") for row in pairs] == ["A1_vs_B1", "A2_vs_B2"],
            "eager confirmation paired matrix differs")
    by_pair = {"A1_vs_B1": (by_label["A1"], by_label["B1"]),
               "A2_vs_B2": (by_label["A2"], by_label["B2"])}
    for pair in pairs:
        name = pair["pair"]
        left, right = by_pair[name]
        require(pair.get("baseline_label") == left["label"]
                and pair.get("candidate_label") == right["label"]
                and pair.get("identity_equal") is True,
                f"eager confirmation {name} identity differs")
        metrics = pair.get("metrics")
        require(isinstance(metrics, dict)
                and isinstance(metrics.get("elapsed_ns"), dict)
                and isinstance(metrics.get("max_rss_kib"), dict),
                f"eager confirmation {name} metrics differ")
        for stat in ("p50", "p95", "p99", "mean"):
            record = metrics["elapsed_ns"].get(stat)
            require(isinstance(record, dict)
                    and record.get("baseline") == left["elapsed_ns"][stat]
                    and record.get("candidate") == right["elapsed_ns"][stat]
                    and math.isclose(record.get("change_percent"),
                                     percent_change(left["elapsed_ns"][stat], right["elapsed_ns"][stat],
                                                   f"{name} elapsed {stat}"),
                                     rel_tol=1e-12, abs_tol=1e-12)
                    and record.get("adverse_over_five_percent") is
                    (record["change_percent"] > 5.0),
                    f"eager confirmation {name} elapsed {stat} comparison differs")
        rss_record = metrics["max_rss_kib"]
        require(rss_record.get("baseline") == left["rss"]["max_rss_kib"]
                and rss_record.get("candidate") == right["rss"]["max_rss_kib"]
                and math.isclose(rss_record.get("change_percent"),
                                 percent_change(left["rss"]["max_rss_kib"],
                                                right["rss"]["max_rss_kib"],
                                                f"{name} RSS"),
                                 rel_tol=1e-12, abs_tol=1e-12)
                and rss_record.get("adverse_over_five_percent") is
                (rss_record["change_percent"] > 5.0),
                f"eager confirmation {name} RSS comparison differs")
    drift = value.get("same_build_drift")
    require(isinstance(drift, list)
            and [row.get("stage") for row in drift] == ["baseline", "candidate"],
            "eager confirmation same-build matrix differs")
    adverse = value.get("adverse_flags_over_five_percent")
    same_adverse = value.get("same_build_drift_over_five_percent")
    require(isinstance(adverse, list) and isinstance(same_adverse, list),
            "eager confirmation adverse vectors differ")
    gate = value.get("gate")
    require(isinstance(gate, dict)
            and isinstance(gate.get("passed"), bool)
            and isinstance(gate.get("paired_rows"), list)
            and [row.get("pair") for row in gate["paired_rows"]] ==
            plan["gate"]["paired_comparisons"],
            "eager confirmation gate is missing")
    for gate_row, pair_name in zip(gate["paired_rows"], plan["gate"]["paired_comparisons"]):
        require(isinstance(gate_row, dict)
                and gate_row.get("required_max_adverse_change_percent") == 5.0
                and isinstance(gate_row.get("passed"), bool),
                f"eager confirmation {pair_name} gate row is malformed")
        matched = next(row for row in pairs if row["pair"] == pair_name)
        expected_p50 = matched["metrics"]["elapsed_ns"]["p50"]
        expected_mean = matched["metrics"]["elapsed_ns"]["mean"]
        require(gate_row.get("p50") == expected_p50
                and gate_row.get("mean") == expected_mean
                and gate_row["passed"] is
                (expected_p50["change_percent"] <= 5.0
                 and expected_mean["change_percent"] <= 5.0),
                f"eager confirmation {pair_name} gate row does not reconcile")
    require(value.get("all_adverse_metrics_retained") is True
            and value.get("no_primary_gain_claim") is True,
            "eager confirmation diagnostic boundaries differ")
    scope = value.get("metrics_scope")
    require(isinstance(scope, dict)
            and scope.get("elapsed_statistics") == ["p50", "p95", "p99", "mean"]
            and scope.get("rss") == "whole_child_process.max_rss_kib"
            and scope.get("phase_metrics") == "not_applicable",
            "eager confirmation metric scope differs")
    return {"path": relative(path), "sha256": sha(path),
            "comparison": value, "adverse_flags": len(adverse),
            "same_build_flags": len(same_adverse)}


def eager_confirmation_review_source() -> tuple[Path, dict[str, Any]]:
    for name in EAGER_CONFIRMATION_REVIEW_NAMES:
        path = HERE / name
        if path.is_file():
            value = read_json(path)
            require(isinstance(value, dict),
                    f"{name} is not an object")
            return path, value
    for name in ("adverse-review.json", "eager-adverse-review.json"):
        path = HERE / name
        if not path.is_file():
            continue
        value = read_json(path)
        if not isinstance(value, dict):
            continue
        for key in ("eager_confirmation", "eager-confirmation",
                    "eager_confirmation_adverse_review"):
            section = value.get(key)
            if isinstance(section, dict):
                return path, section
    raise EvidenceError("eager confirmation adverse review is missing")


def validate_eager_confirmation_adverse_review(
        plan_info: dict[str, Any], comparison_info: dict[str, Any]) -> dict[str, Any]:
    path, value = eager_confirmation_review_source()
    comparison_path = EAGER_CONFIRMATION_COMPARISON_PATH
    require(value.get("comparison_sha256") == sha(comparison_path),
            "eager confirmation adverse review comparison binding differs")
    comparison = comparison_info["comparison"]
    adverse_flags = comparison.get("adverse_flags_over_five_percent", [])
    same_build_flags = comparison.get("same_build_drift_over_five_percent", [])
    require(isinstance(adverse_flags, list) and isinstance(same_build_flags, list),
            "eager confirmation adverse vectors are invalid")
    expected_count = len(adverse_flags) + len(same_build_flags)

    def rows_with_reviews(rows: Any, label: str) -> list[dict[str, Any]]:
        require(isinstance(rows, list),
                f"eager confirmation review {label} rows are missing")
        result: list[dict[str, Any]] = []
        for row in rows:
            require(isinstance(row, dict)
                    and isinstance(row.get("review"), str)
                    and row["review"].strip(),
                    f"an eager confirmation {label} row has no review")
            result.append(row)
        return result

    if isinstance(value.get("matched"), list):
        reviewed_adverse = rows_with_reviews(value["matched"], "matched")
        reviewed_same_build = rows_with_reviews(value.get("same_build"), "same-build")
    elif isinstance(value.get("flags"), list):
        reviewed_adverse = rows_with_reviews(value["flags"], "flag")
        same_raw = value.get("same_build")
        if same_raw is None and len(reviewed_adverse) == expected_count:
            reviewed_adverse, reviewed_same_build = (
                reviewed_adverse[:len(adverse_flags)],
                reviewed_adverse[len(adverse_flags):],
            )
        else:
            reviewed_same_build = rows_with_reviews(same_raw, "same-build")
    else:
        combined = rows_with_reviews(value.get("reviews"), "combined")
        require(len(combined) == expected_count,
                "eager confirmation review combined row count differs")
        reviewed_adverse = combined[:len(adverse_flags)]
        reviewed_same_build = combined[len(adverse_flags):]

    def match_expected(expected: list[Any], actual: list[dict[str, Any]], label: str) -> None:
        require(len(actual) == len(expected),
                f"eager confirmation review {label} count differs")
        remaining = list(actual)
        for source_row in expected:
            require(isinstance(source_row, dict),
                    f"eager confirmation comparison {label} row is not an object")
            match_index = next(
                (index for index, reviewed in enumerate(remaining)
                 if all(reviewed.get(key) == item
                        for key, item in source_row.items())),
                None,
            )
            require(match_index is not None,
                    f"eager confirmation review {label} row differs")
            remaining.pop(match_index)
        require(not remaining,
                f"eager confirmation review {label} has an unbound row")

    match_expected(adverse_flags, reviewed_adverse, "matched")
    match_expected(same_build_flags, reviewed_same_build, "same-build")
    plan = plan_info["plan"]
    require(value.get("order") == plan["order"],
            "eager confirmation adverse review order differs")
    require(value.get("paired_comparisons") == plan["gate"]["paired_comparisons"],
            "eager confirmation adverse review pair list differs")
    require(value.get("all_adverse_metrics_retained") is True
            and value.get("no_primary_gain_claim") is True,
            "eager confirmation adverse review diagnostic boundary differs")
    require(isinstance(value.get("review"), str) and value["review"].strip(),
            "eager confirmation adverse review has no summary review")
    phase_metrics = value.get("phase_metrics")
    if phase_metrics is not None:
        require(isinstance(phase_metrics, dict)
                and phase_metrics.get("status") == "not_applicable",
                "eager confirmation adverse review phase scope differs")
    if "runtime_comparison_performed" in value:
        require(value["runtime_comparison_performed"] is True,
                "eager confirmation adverse review lacks runtime comparison")
    if "adverse_flag_count" in value:
        require(value["adverse_flag_count"] == len(adverse_flags),
                "eager confirmation adverse review count differs")
    if "same_build_drift_count" in value:
        require(value["same_build_drift_count"] == len(same_build_flags),
                "eager confirmation adverse review same-build count differs")
    if "complete" in value:
        require(value["complete"] is True,
                "eager confirmation adverse review is not complete")
    return {"path": relative(path), "sha256": sha(path),
            "comparison_sha256": sha(comparison_path),
            "flags": len(reviewed_adverse),
            "same_build_flags": len(reviewed_same_build),
            "gate_passed": comparison_info["comparison"]["gate"]["passed"]}


def validate_eager_confirmation(plan_info: dict[str, Any]) -> dict[str, Any]:
    capture = validate_eager_confirmation_capture(plan_info)
    # The frozen wrapper has no output option and its analyzer writes the
    # canonical path.  Replay it in a subprocess with an in-memory write
    # redirect so retained evidence is never replaced.
    run_eager_confirmation_replay(EAGER_CONFIRMATION_COMPARISON_PATH)
    comparison = validate_eager_confirmation_comparison(plan_info, capture)
    review = validate_eager_confirmation_adverse_review(plan_info, comparison)
    return {"plan": plan_info, "capture": capture,
            "comparison": comparison, "review": review}


def validate_all(plan: dict[str, Any], static: dict[str, Any]) -> dict[str, Any]:
    require(static.get("candidate") is not None, "candidate source stage is incomplete")
    candidate_sha = sha(HERE / "candidate/source-manifest.json")
    stage_reports = {}
    all_receipts: list[dict[str, Any]] = []
    for stage in ("baseline", "candidate"):
        report = validate_stage_receipts(stage, plan, candidate_sha)
        stage_reports[stage] = report
        all_receipts.extend(report["receipts"])
    preflights = validate_preflights(plan, candidate_sha, static)
    for item in preflights:
        all_receipts.extend(item["receipts"])
    validate_global_intervals(all_receipts)
    native_order = validate_native_order(plan, all_receipts)
    # Profile and hardware analyzers validate the exact raw artifact sets and
    # native-result identities; numeric analysis validates all native/alloc
    # rows and their phase/allocation invariants.
    reports = validate_reports(plan, stage_reports)
    eager_confirmation_plan = validate_eager_confirmation_plan()
    eager_confirmation = validate_eager_confirmation(eager_confirmation_plan)
    adverse = validate_adverse_review(reports["comparison"], reports["eager"])
    adverse["eager_confirmation"] = eager_confirmation["review"]
    decision = validate_disposition(plan, static, reports["comparison"],
                                    reports["profile"], adverse)
    quality_document = read_json(HERE / "quality-summary.json")
    quality_stage = quality_document.get("stage", decision["final_source"])
    available_quality_stages = {"baseline", "candidate", "final"}
    available_quality_stages.update(item["stage"] for item in preflights)
    require(quality_stage in available_quality_stages,
            "quality summary stage is unknown")
    require((HERE / quality_stage / "source-manifest.json").is_file(),
            "quality summary stage manifest is missing")
    quality_stage_report = stage_reports.get(quality_stage)
    if quality_stage_report is None:
        quality_stage_report = next((item for item in preflights
                                     if item["stage"] == quality_stage), None)
    if quality_stage_report is None:
        quality_stage_report = {
            "manifest_sha256": sha(HERE / f"{quality_stage}/source-manifest.json")
        }
    # Stage and preflight receipt ledgers already include quality receipts.
    # Validate them against their command matrix here without appending the
    # same intervals twice to the global serial-order check.
    quality_receipts = validate_check_receipts(
        plan, quality_stage, quality_stage_report, preflights)
    existing_receipt_paths = {row["path"] for row in all_receipts}
    all_receipts.extend(row for row in quality_receipts
                        if row["path"] not in existing_receipt_paths)
    validate_global_intervals(all_receipts)
    quality = validate_quality(plan, quality_stage, all_receipts, preflights)
    cleanup = validate_cleanup(plan)
    return {
        "status": "pass",
        "performance_claim": "candidate-admission-only" if decision["disposition"] == "accepted" else "none",
        "disposition": decision,
        "source_replay": static,
        "preflights": preflights,
        "stage_receipts": {stage: report["receipt_count"] for stage, report in stage_reports.items()},
        "failed_receipts": {
            stage: [row["path"] for row in report["failed_receipts"]]
            for stage, report in stage_reports.items()
        } | {
            item["stage"]: item["failed_receipts"]
            for item in preflights if item["failed_receipts"]
        },
        "native_order": native_order,
        "serialized_intervals": len(all_receipts),
        "native_samples": len(stage_reports) * (
            len(primary_jobs(plan)) * int(plan["primary"]["samples"])
            + (len(native_jobs(plan)) - len(primary_jobs(plan)))
            * int(plan["guard_samples"])
        ),
        "allocator_samples": 2 * 4 * 5,
        "profile_children": 2 * 4,
        "hardware_children": 2 * 4,
        "eager_guard_children": 2 * 4,
        "quality": quality,
        "adverse_review": adverse,
        "eager_confirmation": eager_confirmation,
        "native_primary_gate": reports["native_gate"],
        "profile_commit_ir_gate": reports["profile_gate"],
        "exact_report_replay": True,
        "serial_receipts_non_overlapping": True,
        "owned_paths_absent": True,
        "cleanup": cleanup,
    }


def required_paths(plan: dict[str, Any]) -> list[str]:
    missing: list[str] = []
    for stage in ("baseline", "candidate"):
        root = HERE / stage
        for name in ("source-manifest.json", "source.patch", "build-normal.receipt.json",
                     "build-alloc.receipt.json", "binary-normal.json", "binary-alloc.json"):
            if not (root / name).is_file():
                missing.append(f"{stage}/{name}")
        for name in expected_receipt_names(plan):
            if not (root / f"{name}.receipt.json").is_file():
                missing.append(f"{stage}/{name}.receipt.json")
        if stage == "candidate" and not (root / "source-diff.json").is_file():
            missing.append("candidate/source-diff.json")
        for name in native_jobs(plan):
            for suffix in (".json", ".stdout", ".stderr", ".rss.json"):
                if not (root / f"{name}{suffix}").is_file():
                    missing.append(f"{stage}/{name}{suffix}")
        for name in lane_jobs(plan, "alloc", plan["allocation"]):
            for suffix in (".json", ".stdout", ".stderr"):
                if not (root / f"{name}{suffix}").is_file():
                    missing.append(f"{stage}/{name}{suffix}")
        for name in lane_jobs(plan, "profile", plan["profile"]):
            for suffix in (".json", ".stdout", ".stderr", ".callgrind"):
                if not (root / f"{name}{suffix}").is_file():
                    missing.append(f"{stage}/{name}{suffix}")
            for part in (1, 2, 3, 4):
                if not (root / f"{name}.callgrind.{part}").is_file():
                    missing.append(f"{stage}/{name}.callgrind.{part}")
            for suffix in (".inclusive.txt", ".self.txt"):
                if not (root / f"{name}{suffix}").is_file():
                    missing.append(f"{stage}/{name}{suffix}")
        for name in lane_jobs(plan, "hardware", plan["hardware"]):
            receipt = root / f"{name}.receipt.json"
            hardware_failed = False
            if receipt.is_file():
                try:
                    hardware_failed = read_json(receipt).get("exit_code") != 0
                except EvidenceError:
                    # The receipt itself is reported below; retain the full
                    # artifact requirements until its shape can be checked.
                    hardware_failed = False
            suffixes = (".stdout", ".stderr") if hardware_failed else (
                ".json", ".stdout", ".stderr", ".csv")
            for suffix in suffixes:
                if not (root / f"{name}{suffix}").is_file():
                    missing.append(f"{stage}/{name}{suffix}")
        for name in ("hardware-analysis.json",):
            if not (root / name).is_file():
                missing.append(f"{stage}/{name}")
        if EAGER_PATH.is_file():
            if not (root / "eager-guard-analysis.json").is_file():
                missing.append(f"{stage}/eager-guard-analysis.json")
            for name in [f"eager-r{repeat}-{shape}"
                         for repeat in (1, 2) for shape in ("medium", "dense-sparse")]:
                for suffix in (".json", ".stdout", ".stderr", ".rss.json", ".receipt.json", ".binding.json"):
                    if not (root / f"{name}{suffix}").is_file():
                        missing.append(f"{stage}/{name}{suffix}")
    for name in ("comparison.json", "profile-analysis.json", "quality-summary.json",
                 "cleanup.json"):
        if not (HERE / name).is_file():
            missing.append(name)
    if EAGER_ANALYZER_PATH.is_file() and not (HERE / "eager-guard-comparison.json").is_file():
        missing.append("eager-guard-comparison.json")
    if not (HERE / "decision.json").is_file() and not (HERE / "disposition.json").is_file():
        missing.append("decision.json or disposition.json")
    if not report_path("adverse-review.json", "flag-review.json",
                       "adverse-flags-review.json"):
        missing.append("adverse-review.json or flag-review.json")
    if not EAGER_CONFIRMATION_PLAN_PATH.is_file():
        missing.append(EAGER_CONFIRMATION_PLAN_PATH.name)
    else:
        if not EAGER_CONFIRMATION_WRAPPER_PATH.is_file():
            missing.append(EAGER_CONFIRMATION_WRAPPER_PATH.name)
        if not EAGER_CONFIRMATION_ROOT.is_dir():
            missing.append(EAGER_CONFIRMATION_ROOT.name)
        if not EAGER_CONFIRMATION_COMPARISON_PATH.is_file():
            missing.append(EAGER_CONFIRMATION_COMPARISON_PATH.name)
        confirmation_review_present = any(
            (HERE / name).is_file() for name in EAGER_CONFIRMATION_REVIEW_NAMES
        )
        if not confirmation_review_present:
            for name in ("adverse-review.json", "eager-adverse-review.json"):
                host = HERE / name
                if not host.is_file():
                    continue
                try:
                    host_value = read_json(host)
                except EvidenceError:
                    continue
                if isinstance(host_value, dict) and any(
                        isinstance(host_value.get(key), dict)
                        for key in ("eager_confirmation", "eager-confirmation",
                                    "eager_confirmation_adverse_review")):
                    confirmation_review_present = True
                    break
        if not confirmation_review_present:
            missing.append("eager-confirmation-adverse-review.json")
    if not HARDWARE_PATH.is_file():
        missing.append("analyze_hardware.py")
    if not EAGER_PATH.is_file():
        # The capture helper is part of the frozen supplemental guard contract.
        missing.append("eager_guard.py")
    if not EAGER_ANALYZER_PATH.is_file():
        missing.append("analyze_eager_guard.py")
    return sorted(set(missing))


def verify(sealed: bool = False) -> dict[str, Any]:
    plan = plan_data()
    static = verify_static(plan)
    missing = required_paths(plan)
    if static.get("candidate") is None:
        missing.append("candidate source manifest and patch")
    if missing:
        return {
            "status": "incomplete",
            "performance_claim": "none",
            "plan_sha256": sha(PLAN_PATH),
            "static": static,
            "missing": sorted(set(missing)),
            "reason": "The frozen campaign has not produced all required bound artifacts; no performance gate is evaluated.",
        }
    report = validate_all(plan, static)
    if sealed:
        sums = HERE / "SHA256SUMS"
        require(sums.is_file(), "SHA256SUMS is missing")
        expected: dict[str, str] = {}
        for line in read_text(sums).splitlines():
            digest, name = line.split("  ", 1)
            safe_relative(name, "seal entry")
            require(name != "SHA256SUMS" and name not in expected and is_digest(digest),
                    "invalid duplicate or digest in SHA256SUMS")
            expected[name] = digest
        actual = {relative(path): sha(path) for path in HERE.rglob("*")
                  if path.is_file() and path.name != "SHA256SUMS"}
        require(expected == actual, "SHA256SUMS does not match evidence inventory")
        require(not any(path.is_symlink() for path in HERE.rglob("*")),
                "sealed evidence contains a symlink")
        report["seal_entries"] = len(expected)
    else:
        report["seal_entries"] = None
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sealed", action="store_true")
    parser.add_argument("--strict", action="store_true",
                        help="return exit code 2 when required capture artifacts are incomplete")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        report = verify(args.sealed)
    except (EvidenceError, OSError, subprocess.CalledProcessError) as error:
        print(f"verify.py: error: {error}", file=os.sys.stderr)
        return 2
    if args.output:
        args.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, indent=2))
    if args.strict and report.get("status") != "pass":
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
