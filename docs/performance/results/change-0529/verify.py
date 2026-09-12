"""Bounded independent verifier for the 0529 XML attribute probe pilot.

The verifier deliberately separates incomplete evidence from failed evidence.
It can therefore be used while the serial campaign is running without turning
missing captures into a performance claim.  It validates the source replay,
the two allocator vectors (publication and commit), receipt custody, the
native ABBA order, analyzer replay, the frozen gates, final disposition,
quality receipts, cleanup, and the final SHA256 inventory.
"""

from __future__ import annotations

import argparse
import ast
import datetime as dt
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Any, Callable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
CHECKS = HERE / "checks.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
SOURCE_BINDING = HERE / "source-binding.json"
SOURCE_REVIEW = HERE / "source-review.md"
HARNESS_REVIEW = HERE / "harness-review.md"
HARNESS_PATCH = HERE / "harness.patch"
PRODUCTION_PATCH = HERE / "production.patch"
SEED_MANIFEST = HERE / "source-manifest.json"
ANALYZER = HERE / "analyze.py"
COMPARISON = HERE / "comparison.json"
BASELINE = HERE / "baseline"
CANDIDATE = HERE / "candidate"
FINAL = HERE / "final"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
SCRATCH = Path("/tmp/litchi-goal-0529")
TARGET = Path("/home/zhuhe/litchi-goal-0529-target")
HARNESS_FILE = "tools/perf-baseline/src/lib.rs"
SOURCE_ROOTS = ("crates/xml-minifier/",)
CONDITIONAL_LANES = ("profile", "hardware", "eager")
TIMING_STATS = ("p50", "p95", "p99", "mean")
ALLOC_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)


class EvidenceError(ValueError):
    """Malformed, contradictory, or out-of-scope evidence."""


class IncompleteError(EvidenceError):
    """An artifact needed by a selected component has not arrived."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def rel(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError:
        return str(path)


def repo_rel(path: Path) -> str:
    try:
        return path.relative_to(REPO).as_posix()
    except ValueError:
        return str(path)


def need(path: Path, label: str, *, symlink: bool = True) -> Path:
    if not path.exists():
        raise IncompleteError(f"{label} is missing: {rel(path)}")
    if symlink and path.is_symlink():
        raise EvidenceError(f"{label} is a symlink: {rel(path)}")
    return path


def read_json(path: Path, label: str | None = None) -> Any:
    label = label or rel(path)
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {label}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    need(path, label or rel(path))
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise EvidenceError(f"cannot read text {label or rel(path)}: {error}") from error


def sha(path: Path) -> str:
    need(path, f"artifact for hashing: {rel(path)}", symlink=False)
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {rel(path)}: {error}") from error


def digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(".." not in path.parts and path.as_posix() == value,
            f"{label} escapes its root")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} timestamp is invalid") from error
    require(result.tzinfo is not None, f"{label} timestamp has no timezone")
    return result


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and seconds > 0,
            f"{label}.seconds is not positive")
    require(end > start, f"{label} interval is inverted")
    return start, end


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def source_name(name: str) -> bool:
    exact = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
    return (name.startswith("crates/") or name.startswith("tools/perf-baseline/")
            or name.startswith(".cargo/") or name in exact)


def source_names_at_revision(revision: str) -> set[str]:
    try:
        raw = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", revision],
                                      cwd=REPO, text=True)
    except subprocess.CalledProcessError as error:
        raise EvidenceError(f"cannot list frozen revision: {revision}") from error
    return {name for name in raw.splitlines() if source_name(name)}


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


def source_manifest(path: Path) -> dict[str, str]:
    value = read_json(path, f"{rel(path)} source manifest")
    require(isinstance(value, dict) and value, f"{rel(path)} is not a nonempty manifest")
    result: dict[str, str] = {}
    for name, value in value.items():
        safe_relative(name, f"{rel(path)} source path")
        require(source_name(name) and digest(value),
                f"{rel(path)} has an invalid source entry: {name}")
        require(name not in result, f"{rel(path)} repeats {name}")
        result[name] = value
    return result


def run_checked(args: list[str], *, env: dict[str, str] | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, env=env, stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise EvidenceError(f"command failed ({' '.join(args)}): "
                            f"{detail.decode(errors='replace')[-2000:]}") from error


def parse_private_index(index: Path) -> dict[str, tuple[str, int]]:
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    try:
        raw = subprocess.check_output(["git", "ls-files", "-s", "-z"], cwd=REPO, env=env)
    except subprocess.CalledProcessError as error:
        raise EvidenceError("private source replay index cannot be read") from error
    result: dict[str, tuple[str, int]] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, name_bytes = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3, "private source replay entry is malformed")
        name = name_bytes.decode()
        require(name not in result, f"private source replay repeats {name}")
        result[name] = (fields[1].decode(), int(fields[0]))
    return result


def index_blob_hashes(index: Path, oids: set[str]) -> dict[str, str]:
    if not oids:
        return {}
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    # Use a direct, checked call with stdin so replay remains read-only and
    # does not touch HEAD.
    try:
        raw = subprocess.check_output(["git", "cat-file", "--batch"], cwd=REPO, env=env,
                                      input=("\n".join(sorted(oids)) + "\n").encode(),
                                      stderr=subprocess.PIPE)
    except subprocess.CalledProcessError as error:
        raise EvidenceError("private source replay objects cannot be read") from error
    result: dict[str, str] = {}
    position = 0
    while position < len(raw):
        end = raw.find(b"\n", position)
        require(end >= 0, "Git batch response is malformed")
        oid, kind, size = raw[position:end].split()
        require(kind == b"blob", f"source replay object is not a blob: {oid!r}")
        position = end + 1
        length = int(size)
        data = raw[position:position + length]
        require(len(data) == length, "Git batch response is truncated")
        result[oid.decode()] = hashlib.sha256(data).hexdigest()
        position += length
        require(raw[position:position + 1] == b"\n", "Git batch separator is missing")
        position += 1
    return result


def sidecars(folder: Path) -> dict[str, str]:
    path = folder / "new-files.json"
    if not path.exists():
        return {}
    value = read_json(path)
    rows = value.get("files") if isinstance(value, dict) else None
    require(isinstance(rows, dict), f"{rel(path)}.files is not an object")
    result: dict[str, str] = {}
    for name, item in rows.items():
        safe_relative(name, f"{rel(path)} new source")
        require(any(name.startswith(root) for root in SOURCE_ROOTS),
                f"{rel(path)} new source escapes candidate root: {name}")
        if isinstance(item, str):
            expected, artifact = item, name
        else:
            require(isinstance(item, dict), f"{rel(path)} entry is malformed: {name}")
            expected, artifact = item.get("sha256"), item.get("artifact", name)
        safe_relative(artifact, f"{rel(path)} artifact")
        require(digest(expected), f"{rel(path)} digest is malformed: {name}")
        artifact_path = folder / artifact
        require(artifact_path.is_file() and not artifact_path.is_symlink()
                and sha(artifact_path) == expected,
                f"{rel(path)} artifact custody differs: {artifact}")
        result[name] = expected
    return result


def replay_patch(patch_path: Path, revision: str) -> tuple[set[str], dict[str, tuple[str, int]], dict[str, str]]:
    need(patch_path, f"source patch {rel(patch_path)}")
    with tempfile.TemporaryDirectory(prefix="litchi-0529-replay-", dir="/home/zhuhe") as directory:
        index = Path(directory) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        run_checked(["git", "read-tree", revision], env=env)
        if patch_path.stat().st_size:
            run_checked(["git", "apply", "--cached", "--binary", str(patch_path)], env=env)
        indexed = parse_private_index(index)
        try:
            changed = set(subprocess.check_output(
                ["git", "diff", "--cached", "--name-only", revision],
                cwd=REPO, env=env, text=True).splitlines())
        except subprocess.CalledProcessError as error:
            raise EvidenceError(f"source replay diff cannot be read: {rel(patch_path)}") from error
        scoped = {name: row for name, row in indexed.items() if source_name(name)}
        hashes = index_blob_hashes(index, {oid for oid, _ in scoped.values()})
        return changed, scoped, hashes


def replay_stage(stage: str, plan: dict[str, Any], *, allow_empty: bool = False) -> dict[str, Any]:
    folder = HERE / stage
    manifest_path = need(folder / "source-manifest.json", f"{stage} source manifest")
    patch_path = need(folder / "source.patch", f"{stage} source patch")
    expected = source_manifest(manifest_path)
    changed, indexed, hashes = replay_patch(patch_path, plan["revision"])
    names_at_revision = source_names_at_revision(plan["revision"])
    extras = sidecars(folder)
    require(all(source_name(name) for name in changed),
            f"{stage} source patch changes an out-of-scope file")
    allowed = {HARNESS_FILE} | set(name for name in changed
                                   if any(name.startswith(root) for root in SOURCE_ROOTS))
    require(changed <= allowed, f"{stage} source patch has an unexpected file")
    require(set(expected) == (set(indexed) | set(extras)),
            f"{stage} source manifest/index inventory differs")
    for name, expected_sha in expected.items():
        if name in indexed:
            oid, mode = indexed[name]
            require(mode in (100644, 100755), f"{stage} source mode is unexpected: {name}")
            require(hashes.get(oid) == expected_sha,
                    f"{stage} replay blob differs: {name}")
        else:
            require(extras.get(name) == expected_sha,
                    f"{stage} sidecar source differs: {name}")
    if stage == "baseline":
        require(changed == {HARNESS_FILE}, "baseline diff must contain only the harness enabler")
        require(set(expected) == names_at_revision | {HARNESS_FILE},
                "baseline source inventory differs from the frozen revision")
    elif stage in ("candidate", "final"):
        if not allow_empty:
            require(changed, f"{stage} source patch is empty")
        if stage == "candidate":
            require(changed - {HARNESS_FILE}, "candidate patch has no production source change")
            require(HARNESS_FILE in changed,
                    "candidate patch omits the shared harness measurement enabler")
            require(all(any(name.startswith(root) for root in SOURCE_ROOTS)
                        for name in changed - {HARNESS_FILE}),
                    "candidate patch escapes the XML minifier root")
    return {"stage": stage, "manifest_sha256": sha(manifest_path),
            "manifest_entries": len(expected), "patch_sha256": sha(patch_path),
            "changed_files": sorted(changed), "new_files": sorted(extras)}


def load_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict) and set(frozen) == {
        "plan.json", "run.py", "checks.py", "harness.patch", "production.patch",
    }, "frozen input inventory differs")
    for name, path in (("plan.json", PLAN), ("run.py", RUN), ("checks.py", CHECKS),
                       ("harness.patch", HARNESS_PATCH),
                       ("production.patch", PRODUCTION_PATCH)):
        require(frozen.get(name) == sha(path), f"frozen {name} digest differs")
    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict) and plan.get("status") == "frozen-before-build-and-capture",
            "plan is not frozen before build/capture")
    revision = plan.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as error:
        raise EvidenceError("plan revision is not a Git commit") from error
    require(plan.get("candidate_source_roots") == list(SOURCE_ROOTS),
            "candidate source roots differ")
    require(plan.get("owned_paths") == [str(SCRATCH), str(TARGET),
                                         "/home/zhuhe/litchi-goal-0529-draft",
                                         "/home/zhuhe/litchi-goal-0529-tests",
                                         "/home/zhuhe/litchi-goal-0529-metrics"],
            "owned temporary paths differ")
    require(plan.get("cpu") == 2, "planned CPU differs")
    require(plan.get("native_order") == [
        "baseline native-r1 before candidate application", "candidate native-r1",
        "candidate native-r2", "retained baseline native-r2 under candidate source checkout",
    ], "native ABBA order differs")
    primary = plan.get("primary")
    require(isinstance(primary, dict) and primary.get("case") ==
            "xlsx_source_backed_cell_values_one_percent_edit_save"
            and primary.get("shapes") == ["medium", "dense-sparse"]
            and primary.get("repeats") == 2 and primary.get("warmup") == 20
            and primary.get("samples") == 200, "primary plan differs")
    require(plan.get("guard_repeats") == 2 and plan.get("guard_warmup") == 10
            and plan.get("guard_samples") == 30, "guard counts differ")
    require(isinstance(plan.get("guards"), list) and len(plan["guards"]) == 5,
            "guard plan differs")
    expected_guards = [
        {"case": "xlsx_source_backed_cell_values_one_edit_save",
         "shapes": ["medium", "dense-sparse"]},
        {"case": "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
         "shapes": ["medium", "dense-sparse"]},
        {"case": "xlsx_source_backed_cell_values_one_percent_edit_save",
         "shapes": ["vendor-extension"]},
        {"case": "xlsx_source_backed_cell_values_one_percent_edit_save",
         "shapes": ["noncompact"]},
        {"case": "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
         "shapes": ["noncompact"]},
    ]
    require(plan["guards"] == expected_guards, "guard cases differ")
    allocation = plan.get("allocation")
    require(isinstance(allocation, dict) and allocation.get("shapes") == primary["shapes"]
            and allocation.get("repeats") == 2 and allocation.get("warmup") == 0
            and allocation.get("samples") == 5
            and "publication" in str(allocation.get("scope", "")).lower(),
            "publication allocation plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict) and profile.get("shapes") == primary["shapes"]
            and profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 1
            and "publish_multi_commit_to_stream" in profile.get("owner", ""),
            "publication profile plan differs")
    gates = plan.get("gates")
    require(gates == {
        "total_p50_reduction_percent": 2.0,
        "total_mean_reduction_percent": 2.0,
        "allocation_calls_reduction_percent": 8.0,
        "require_every_shape_repeat": True,
        "publication_p50_reduction_percent": 5.0,
        "publication_ir_reduction_percent": 3.0,
    }, "0529 gates differ")
    require(isinstance(plan.get("draft_patch_sha256"), str)
            and plan["draft_patch_sha256"] == sha(PRODUCTION_PATCH),
            "production patch binding differs")
    require(isinstance(plan.get("harness_patch_sha256"), str)
            and plan["harness_patch_sha256"] == sha(HARNESS_PATCH),
            "harness patch binding differs")
    return plan


def validate_reviews_and_binding(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    files = value.get("files") if isinstance(value, dict) else None
    require(isinstance(files, dict) and files, "ADR manifest is malformed")
    for name, expected in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith("docs/adr/") and digest(expected)
                and (REPO / name).is_file() and sha(REPO / name) == expected,
                f"ADR binding differs: {name}")
    prior = value.get("prior_audit_source_binding_sha256")
    prior_path = HERE.parent / "change-0526" / "source-binding.json"
    require(digest(prior) and prior_path.is_file() and sha(prior_path) == prior,
            "prior 0526 source binding differs")
    for path, label in ((SOURCE_REVIEW, "source review"), (HARNESS_REVIEW, "harness review")):
        text = read_text(path, label)
        require(plan["revision"] in text and "0529" in text
                and "OLE2/OOXML" in text and "ODF" in text and "iWork" in text,
                f"{label} is not bound to the frozen priority/revision")
    source_text = read_text(SOURCE_REVIEW, "source review")
    require("quick-xml" in source_text and "0.41.0" in source_text,
            "source review omits pinned dependency scope")
    harness_text = read_text(HARNESS_REVIEW, "harness review")
    require("publication_allocation_metrics" in harness_text
            and "commit" in harness_text.lower(),
            "harness review omits the separate publication/commit channels")
    binding = read_json(SOURCE_BINDING, "source-binding.json")
    require(isinstance(binding, dict) and binding.get("revision") == plan["revision"],
            "source binding revision differs")
    # Bind every explicitly retained source hash, while allowing the binding
    # document to add useful labels without making its schema brittle.
    bound = binding.get("source_files")
    require(isinstance(bound, dict) and bound, "source binding source_files is missing")
    for name, expected in bound.items():
        safe_relative(name, "source binding path")
        require(digest(expected) and (REPO / name).is_file()
                and sha(REPO / name) == expected,
                f"source binding differs: {name}")
    return {"adr_sha256": sha(ADR), "prior_binding_sha256": sha(prior_path),
            "source_binding_sha256": sha(SOURCE_BINDING),
            "source_review_sha256": sha(SOURCE_REVIEW),
            "harness_review_sha256": sha(HARNESS_REVIEW)}


def validate_source() -> dict[str, Any]:
    plan = load_plan()
    reviews = validate_reviews_and_binding(plan)
    baseline = replay_stage("baseline", plan)
    candidate = replay_stage("candidate", plan)
    baseline_manifest = source_manifest(BASELINE / "source-manifest.json")
    candidate_manifest = source_manifest(CANDIDATE / "source-manifest.json")
    changed = sorted(name for name in set(baseline_manifest) | set(candidate_manifest)
                     if baseline_manifest.get(name) != candidate_manifest.get(name))
    require(changed and all(any(name.startswith(root) for root in SOURCE_ROOTS)
                            for name in changed),
            "candidate-baseline manifest diff escapes XML minifier")
    diff = read_json(CANDIDATE / "source-diff.json", "candidate source-diff.json")
    require(diff.get("baseline_manifest_sha256") == baseline["manifest_sha256"]
            and diff.get("candidate_manifest_sha256") == candidate["manifest_sha256"]
            and diff.get("candidate_source_roots") == list(SOURCE_ROOTS),
            "candidate source-diff binding differs")
    rows = diff.get("changed_files")
    require(isinstance(rows, dict) and sorted(rows) == changed,
            "candidate source-diff inventory differs")
    for name in changed:
        row = rows[name]
        require(isinstance(row, dict) and row.get("baseline_sha256") == baseline_manifest.get(name)
                and row.get("candidate_sha256") == candidate_manifest.get(name),
                f"candidate source-diff digest differs: {name}")
    # The frozen harness patch must replay to the exact baseline harness blob.
    harness_changed, harness_index, harness_hashes = replay_patch(HARNESS_PATCH, plan["revision"])
    require(harness_changed == {HARNESS_FILE}, "harness patch changes unexpected files")
    require(harness_hashes.get(harness_index[HARNESS_FILE][0]) == baseline_manifest[HARNESS_FILE],
            "baseline does not contain the reviewed harness patch")
    production_changed, _, _ = replay_patch(PRODUCTION_PATCH, plan["revision"])
    require(production_changed and all(any(name.startswith(root)
                                            for root in SOURCE_ROOTS)
                                        for name in production_changed),
            "production patch escapes the candidate root")
    return {"status": "pass", "reviews": reviews, "baseline": baseline,
            "candidate": candidate, "candidate_baseline_changes": changed,
            "baseline_harness_change": sorted(harness_changed),
            "production_patch_changes": sorted(production_changed)}


def expected_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    primary = plan["primary"]
    jobs: list[dict[str, Any]] = []
    if lane == "native":
        for repeat in range(1, plan["primary"]["repeats"] + 1):
            for shape in primary["shapes"]:
                jobs.append({"name": f"native-r{repeat}-primary-{shape}", "kind": "primary",
                             "guard": None, "repeat": repeat, "case": primary["case"],
                             "shape": shape, "warmup": primary["warmup"],
                             "samples": primary["samples"]})
        for repeat in range(1, plan["guard_repeats"] + 1):
            for guard, item in enumerate(plan["guards"]):
                for shape in item["shapes"]:
                    jobs.append({"name": f"native-r{repeat}-guard{guard}-{shape}",
                                 "kind": "guard", "guard": guard, "repeat": repeat,
                                 "case": item["case"], "shape": shape,
                                 "warmup": plan["guard_warmup"], "samples": plan["guard_samples"]})
    else:
        config = plan["allocation"]
        for repeat in range(1, config["repeats"] + 1):
            for shape in config["shapes"]:
                jobs.append({"name": f"alloc-r{repeat}-{shape}", "kind": "allocation",
                             "guard": None, "repeat": repeat, "case": primary["case"],
                             "shape": shape, "warmup": config["warmup"],
                             "samples": config["samples"]})
    return jobs


def expected_artifacts(name: str) -> set[str]:
    suffixes = (".json", ".stdout", ".stderr", ".rss.json") if name.startswith("native-") \
        else (".json", ".stdout", ".stderr")
    return {name + suffix for suffix in suffixes}


def expected_working_manifest(stage: str, name: str, stage_sha: str,
                              candidate_sha: str) -> str:
    return candidate_sha if stage == "candidate" or (stage == "baseline"
                                                       and name.startswith("native-r2-")) \
        else stage_sha


def validate_artifact_map(receipt: dict[str, Any], folder: Path,
                          expected: set[str], label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} artifact inventory differs")
    for name, expected_sha in artifacts.items():
        safe_relative(name, f"{label} artifact")
        path = folder / name
        require(Path(name).name == name and path.is_file() and not path.is_symlink()
                and digest(expected_sha) and sha(path) == expected_sha,
                f"{label} artifact custody differs: {name}")


def receipt_common(path: Path, stage: str, plan: dict[str, Any], stage_sha: str,
                   candidate_sha: str, *, binary_sha: str | None = None) -> dict[str, Any]:
    value = read_json(path, rel(path))
    start, end = interval(value, rel(path))
    require(value.get("exit_code") == 0, f"{rel(path)} did not exit successfully")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("script_sha256") == sha(RUN),
            f"{rel(path)} frozen plan/driver binding differs")
    name = path.name.removesuffix(".receipt.json")
    require(value.get("source_manifest_sha256") == stage_sha
            and value.get("working_source_manifest_sha256") ==
            expected_working_manifest(stage, name, stage_sha, candidate_sha),
            f"{rel(path)} source manifest binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict)
            and environment.get("TMPDIR") == str(TARGET / "test-tmp"),
            f"{rel(path)} TMPDIR binding differs")
    if binary_sha is not None:
        require(value.get("binary_sha256") == binary_sha,
                f"{rel(path)} binary binding differs")
    return {"path": rel(path), "name": name, "start_utc": value["start_utc"],
            "end_utc": value["end_utc"], "start": start.isoformat(),
            "end": end.isoformat(),
            "sha256": sha(path), "value": value}


def public_receipt(row: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in row.items() if key != "value"}


def build_command(kind: str, plan: dict[str, Any]) -> list[str]:
    executable = "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else "")
    command = ["env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
               "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
               "--bin", executable, "--target-dir", plan["owned_paths"][1]]
    if kind == "alloc":
        command += ["--features", "allocator-metrics"]
    return command


def validate_binary(stage: str, kind: str, stage_sha: str, plan: dict[str, Any]) -> dict[str, Any]:
    path = HERE / stage / f"binary-{kind}.json"
    value = read_json(path, rel(path))
    expected_path = SCRATCH / f"{stage}-{kind}"
    require(value.get("path") == str(expected_path) and digest(value.get("sha256")),
            f"{rel(path)} identity path/digest differs")
    nonnegative_int(value.get("bytes"), f"{rel(path)}.bytes")
    require(value["bytes"] > 0 and value.get("source_manifest_sha256") == stage_sha,
            f"{rel(path)} identity source binding differs")
    receipt = HERE / stage / f"build-{kind}.receipt.json"
    require(value.get("build_receipt_sha256") == sha(receipt),
            f"{rel(path)} build receipt binding differs")
    binary = Path(value["path"])
    if binary.exists():
        require(binary.is_file() and not binary.is_symlink() and sha(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                f"{rel(path)} retained binary differs")
    else:
        require(CLEANUP.is_file(), f"{rel(path)} binary disappeared without cleanup receipt")
        cleanup = read_json(CLEANUP, "cleanup.json")
        require(cleanup.get("owned_paths_absent") is True
                and cleanup.get("removed") == plan["owned_paths"],
                f"{rel(path)} absent binary lacks cleanup custody")
    return {"kind": kind, "path": str(binary), "sha256": value["sha256"],
            "bytes": value["bytes"], "identity_sha256": sha(path)}


def validate_builds() -> dict[str, Any]:
    plan = load_plan()
    source = validate_source()
    stages: dict[str, Any] = {}
    for stage in ("baseline", "candidate"):
        stage_sha = sha(HERE / stage / "source-manifest.json")
        rows = []
        for kind in ("normal", "alloc"):
            receipt_path = HERE / stage / f"build-{kind}.receipt.json"
            row = receipt_common(receipt_path, stage, plan, stage_sha,
                                 source["candidate"]["manifest_sha256"])
            value = row["value"]
            require(value.get("binary_sha256") is None
                    and value.get("command") == build_command(kind, plan),
                    f"{rel(receipt_path)} command/result differs")
            validate_artifact_map(value, receipt_path.parent,
                                  {f"build-{kind}.stdout", f"build-{kind}.stderr"}, rel(receipt_path))
            rows.append(public_receipt(row))
        identities = {kind: validate_binary(stage, kind, stage_sha, plan)
                      for kind in ("normal", "alloc")}
        stages[stage] = {"manifest_sha256": stage_sha, "builds": rows,
                         "binaries": identities}
    storage_path = HERE / "storage.json"
    storage = read_json(storage_path, "storage.json") if storage_path.exists() else None
    if storage is not None:
        require(storage.get("scratch_symlink") == str(SCRATCH)
                and storage.get("target") == str(TARGET / "retained-binaries"),
                "storage receipt differs")
    return {"status": "pass", "stages": stages, "storage": storage}


def validate_baseline_partial() -> dict[str, Any]:
    """Validate whatever baseline evidence exists before candidate A1.

    This intentionally does not require candidate custody or the final source
    binding.  Missing planned receipts are reported as ``incomplete`` by the
    component envelope, while completed build/child receipts are still checked
    with the same validators used by the final bundle.
    """
    plan = load_plan()
    replay = replay_stage("baseline", plan)
    folder = BASELINE
    stage_sha = replay["manifest_sha256"]
    pending: list[str] = []
    checked: list[dict[str, Any]] = []
    identities: dict[str, dict[str, Any]] = {}
    for kind in ("normal", "alloc"):
        receipt_path = folder / f"build-{kind}.receipt.json"
        identity_path = folder / f"binary-{kind}.json"
        if not receipt_path.exists() or not identity_path.exists():
            pending.append(f"build-{kind}")
            continue
        row = receipt_common(receipt_path, "baseline", plan, stage_sha, stage_sha)
        value = row["value"]
        require(value.get("binary_sha256") is None
                and value.get("command") == build_command(kind, plan),
                f"{rel(receipt_path)} command/result differs")
        validate_artifact_map(value, folder,
                              {f"build-{kind}.stdout", f"build-{kind}.stderr"}, rel(receipt_path))
        identity = validate_binary("baseline", kind, stage_sha, plan)
        identities[kind] = identity
        checked.append({"kind": kind, "receipt": row["path"],
                        "identity": identity["identity_sha256"]})
    jobs = expected_jobs(plan, "native") + expected_jobs(plan, "allocation")
    for job in jobs:
        receipt_path = folder / f"{job['name']}.receipt.json"
        if not receipt_path.exists():
            pending.append(job["name"])
            continue
        kind = "alloc" if job["kind"] == "allocation" else "normal"
        if kind not in identities:
            pending.append(f"{job['name']} (binary {kind})")
            continue
        checked.append(validate_capture_receipt(
            receipt_path, "baseline", job, identities[kind], plan, stage_sha, stage_sha))
    if checked:
        validate_serial([row for row in checked if "start" in row])
    return {"status": "pass" if not pending else "incomplete", "baseline": replay,
            "checked": checked, "pending": pending}


def allocation_sample(value: Any, label: str, measured: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    require(value.get("scope") == "operation_global_system_allocator",
            f"{label}.scope differs")
    if not measured:
        require(value.get("status") == "unavailable"
                and all(field not in value for field in ALLOC_FIELDS),
                f"{label} unavailable sample is malformed")
        return {"status": "unavailable", "scope": value["scope"]}
    require(value.get("status") == "measured", f"{label}.status is not measured")
    result: dict[str, Any] = {"status": "measured", "scope": value["scope"]}
    for field in ALLOC_FIELDS:
        nonnegative_int(value.get(field), f"{label}.{field}")
        result[field] = value[field]
    require(result["failed_allocation_calls"] == 0
            and result["live_bytes_before"] + result["allocated_bytes"] ==
            result["live_bytes_after"] + result["deallocated_bytes"],
            f"{label} allocation balance differs")
    require(result["peak_live_bytes_before"] >= result["live_bytes_before"]
            and result["peak_live_bytes_after"] >= result["peak_live_bytes_before"]
            and result["peak_live_bytes_after"] >= result["live_bytes_after"]
            and result["region_peak_live_bytes"] >= max(result["live_bytes_before"],
                                                          result["live_bytes_after"]),
            f"{label} allocation peak bounds differ")
    result["incremental_region_peak_live_bytes"] = (
        result["region_peak_live_bytes"] - result["live_bytes_before"])
    return result


def validate_raw_vectors(raw_path: Path, job: dict[str, Any], allocator: bool) -> None:
    raw = read_json(raw_path, rel(raw_path))
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{job['name']} raw result count differs")
    result = results[0]
    source = result.get("source") if isinstance(result, dict) else None
    xlsx = source.get("xlsx_cell_values") if isinstance(source, dict) else None
    require(isinstance(xlsx, dict), f"{job['name']} XLSX source evidence is missing")
    for field in ("commit_allocation_metrics", "publication_allocation_metrics"):
        values = xlsx.get(field)
        require(isinstance(values, list) and len(values) == job["samples"],
                f"{job['name']}.{field} vector length differs")
        for index, sample in enumerate(values):
            allocation_sample(sample, f"{job['name']}.{field}[{index}]", allocator)
    # The vectors are deliberately distinct evidence channels.  A report that
    # silently aliases publication to commit is not admissible.
    require("publication_allocation_metrics" in xlsx
            and "commit_allocation_metrics" in xlsx,
            f"{job['name']} omitted one allocation channel")


def capture_command(job: dict[str, Any], binary: dict[str, Any], folder: Path,
                    plan: dict[str, Any]) -> list[str]:
    command = ["taskset", "-c", str(plan["cpu"])]
    if job["kind"] != "allocation":
        command += ["/usr/bin/time", "-f",
                    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                    "-o", str(folder / (job["name"] + ".rss.json"))]
    command += [binary["path"], "--warmup", str(job["warmup"]), "--samples",
                str(job["samples"]), "--case", job["case"], "--xlsx-cell-crud-shape",
                job["shape"], "--json", str(folder / (job["name"] + ".json"))]
    return command


def validate_capture_receipt(path: Path, stage: str, job: dict[str, Any],
                            binary: dict[str, Any], plan: dict[str, Any],
                            stage_sha: str, candidate_sha: str) -> dict[str, Any]:
    row = receipt_common(path, stage, plan, stage_sha, candidate_sha,
                         binary_sha=binary["sha256"])
    value = row["value"]
    require(value.get("command") == capture_command(job, binary, path.parent, plan),
            f"{rel(path)} command differs")
    validate_artifact_map(value, path.parent, expected_artifacts(job["name"]), rel(path))
    validate_raw_vectors(path.parent / (job["name"] + ".json"), job,
                         job["kind"] == "allocation")
    row["job"] = job
    return public_receipt(row)


def validate_serial(rows: list[dict[str, Any]]) -> None:
    ordered = sorted(rows, key=lambda row: parse_time(row["start"], row["path"]))
    for left, right in zip(ordered, ordered[1:]):
        require(parse_time(left["end"], left["path"]) <=
                parse_time(right["start"], right["path"]),
                f"receipt intervals overlap: {left['path']} and {right['path']}")


def validate_abba(rows: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    native = [row for row in rows if row["job"]["kind"] in ("primary", "guard")]
    blocks: list[tuple[str, int]] = []
    block_rows: dict[tuple[str, int], set[str]] = {}
    for row in sorted(native, key=lambda item: parse_time(item["start"], item["path"])):
        match = re.match(r"^native-r([12])-", row["job"]["name"])
        require(match is not None, f"native receipt name is malformed: {row['path']}")
        key = (Path(row["path"]).parts[-2], int(match.group(1)))
        if not blocks or blocks[-1] != key:
            blocks.append(key)
        block_rows.setdefault(key, set()).add(row["job"]["name"])
    expected = [("baseline", 1), ("candidate", 1), ("candidate", 2), ("baseline", 2)]
    require(blocks == expected, f"native ABBA blocks differ: {blocks}")
    for stage, repeat in expected:
        names = {job["name"] for job in expected_jobs(plan, "native")
                 if job["name"].startswith(f"native-r{repeat}-")}
        require(block_rows.get((stage, repeat)) == names,
                f"{stage} native r{repeat} block is incomplete")
    return {"blocks": [f"{stage}/native-r{repeat}" for stage, repeat in blocks],
            "retained_baseline_a2": True}


def validate_captures() -> dict[str, Any]:
    plan = load_plan()
    source = validate_source()
    builds = validate_builds()
    candidate_sha = source["candidate"]["manifest_sha256"]
    rows: list[dict[str, Any]] = []
    stages: dict[str, Any] = {}
    for stage in ("baseline", "candidate"):
        folder = HERE / stage
        stage_sha = sha(folder / "source-manifest.json")
        identities = builds["stages"][stage]["binaries"]
        jobs = expected_jobs(plan, "native") + expected_jobs(plan, "allocation")
        expected = {job["name"] for job in jobs}
        actual = {path.name.removesuffix(".receipt.json")
                  for path in folder.glob("*.receipt.json")
                  if path.name.startswith(("native-", "alloc-"))}
        require(actual == expected, f"{stage} capture matrix differs: {sorted(actual ^ expected)}")
        stage_rows = []
        for job in jobs:
            binary = identities["alloc" if job["kind"] == "allocation" else "normal"]
            stage_rows.append(validate_capture_receipt(
                folder / (job["name"] + ".receipt.json"), stage, job, binary,
                plan, stage_sha, candidate_sha))
        rows.extend(stage_rows)
        stages[stage] = {"rows": stage_rows, "manifest_sha256": stage_sha}
    validate_serial(rows)
    abba = validate_abba(rows, plan)
    require(sum(row["job"]["samples"] for row in rows if row["job"]["kind"] != "allocation")
            == 2440, "native sample ledger differs")
    require(sum(row["job"]["samples"] for row in rows if row["job"]["kind"] == "allocation")
            == 40, "publication allocator sample ledger differs")
    return {"status": "pass", "stages": stages, "rows": rows,
            "abba": abba, "native_samples": 2440, "publication_allocator_samples": 40}


def load_analyzer() -> Any:
    need(ANALYZER, "canonical analyzer")
    spec = importlib.util.spec_from_file_location("litchi_0529_analyzer", ANALYZER)
    require(spec is not None and spec.loader is not None,
            "canonical analyzer cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def replay_analyzer(path: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0529-analysis-", dir="/home/zhuhe") as directory:
        output = Path(directory) / "comparison.json"
        result = subprocess.run([sys.executable, "-B", str(ANALYZER), "--output", str(output)],
                                cwd=REPO, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                check=False)
        require(result.returncode == 0,
                f"canonical analyzer failed: {result.stderr.decode(errors='replace')[-2000:]}")
        require(output.is_file() and output.read_bytes() == path.read_bytes(),
                "canonical analyzer report does not replay byte-for-byte")


def stage_evidence(report: dict[str, Any], stage: str) -> dict[str, Any]:
    value = report.get(stage)
    require(isinstance(value, dict), f"comparison {stage} evidence is missing")
    if isinstance(value.get("evidence"), dict):
        value = value["evidence"]
    require(isinstance(value.get("native"), dict)
            and isinstance(value.get("allocation"), dict),
            f"comparison {stage} lane evidence is malformed")
    return value


def primary_rows(evidence: dict[str, Any], plan: dict[str, Any], label: str) -> dict[tuple[int, str], dict[str, Any]]:
    rows = evidence["native"].get("rows")
    require(isinstance(rows, list), f"{label} native rows are missing")
    actual = [row for row in rows if isinstance(row, dict)
              and row.get("kind") == "primary" and row.get("guard") is None]
    expected = {(repeat, shape) for repeat in range(1, plan["primary"]["repeats"] + 1)
                for shape in plan["primary"]["shapes"]}
    keys = [(row.get("repeat"), row.get("shape")) for row in actual]
    require(set(keys) == expected and len(keys) == len(set(keys)) == len(expected),
            f"{label} primary matrix differs")
    return {(row["repeat"], row["shape"]): row for row in actual}


def stat(row: dict[str, Any], phase: str, name: str, label: str) -> float:
    timing = row.get("timing", {}).get(phase)
    require(isinstance(timing, dict), f"{label}.{phase} timing is missing")
    value = timing.get(name)
    finite(value, f"{label}.{phase}.{name}")
    require(float(value) >= 0, f"{label}.{phase}.{name} is negative")
    return float(value)


def reduction(left: float, right: float, label: str) -> float:
    require(left > 0, f"{label} baseline is not positive")
    return (left - right) / left * 100.0


def numeric_p50(values: list[Any], label: str) -> float:
    require(isinstance(values, list) and values, f"{label} is empty")
    for index, value in enumerate(values):
        nonnegative_int(value, f"{label}[{index}]")
    ordered = sorted(values)
    return (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2.0


def publication_alloc(row: dict[str, Any], label: str) -> list[dict[str, Any]]:
    # The numerical adapter retains the publication vector under an explicit
    # name while keeping the commit vector separately.  Never fall back to the
    # commit vector when the publication field is absent.
    for key in ("publication_allocation", "publication_allocation_metrics"):
        value = row.get(key)
        if isinstance(value, list):
            require(value and all(isinstance(item, dict) for item in value),
                    f"{label}.{key} is malformed")
            return value
    raise EvidenceError(f"{label} publication allocation lane is missing")


def check_gate_row(row: dict[str, Any], expected: dict[str, float], label: str) -> None:
    require(isinstance(row.get("passed"), bool), f"{label}.passed is missing")
    checks = row
    for name, value in expected.items():
        candidates = [item for key, item in checks.items()
                      if isinstance(key, str) and name in key and isinstance(item, dict)]
        require(candidates, f"{label} lacks {name} gate")
        selected = candidates[0]
        reduction_value = selected.get("reduction_percent")
        finite(reduction_value, f"{label}.{name}.reduction_percent")
        require(math.isclose(float(reduction_value), value, rel_tol=1e-9, abs_tol=1e-7),
                f"{label}.{name} gate arithmetic differs")
        require(selected.get("passed") is (value >= float(selected.get("required_reduction_percent"))),
                f"{label}.{name} gate status differs")


def independent_pilot(report: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    left = primary_rows(stage_evidence(report, "baseline"), plan, "baseline")
    right = primary_rows(stage_evidence(report, "candidate"), plan, "candidate")
    native_rows: list[dict[str, Any]] = []
    native_passed = True
    for key in sorted(left):
        lrow, rrow = left[key], right[key]
        total_p50 = reduction(stat(lrow, "elapsed_ns", "p50", str(key)),
                              stat(rrow, "elapsed_ns", "p50", str(key)), str(key))
        total_mean = reduction(stat(lrow, "elapsed_ns", "mean", str(key)),
                               stat(rrow, "elapsed_ns", "mean", str(key)), str(key))
        publication_p50 = reduction(stat(lrow, "publication_ns", "p50", str(key)),
                                    stat(rrow, "publication_ns", "p50", str(key)), str(key))
        passed = (total_p50 >= 2.0 and total_mean >= 2.0 and publication_p50 >= 5.0)
        native_passed = native_passed and passed
        native_rows.append({"repeat": key[0], "shape": key[1],
                            "total_p50": total_p50, "total_mean": total_mean,
                            "publication_p50": publication_p50, "passed": passed})
    lalloc = stage_evidence(report, "baseline")["allocation"].get("rows")
    ralloc = stage_evidence(report, "candidate")["allocation"].get("rows")
    require(isinstance(lalloc, list) and isinstance(ralloc, list),
            "allocation rows are missing")
    by_key = lambda rows: {(row.get("repeat"), row.get("shape")): row for row in rows
                           if isinstance(row, dict) and row.get("kind") == "allocation"}
    lb, rb = by_key(lalloc), by_key(ralloc)
    require(set(lb) == set(rb) and lb, "allocation matrix differs")
    allocation_rows: list[dict[str, Any]] = []
    allocation_passed = True
    for key in sorted(lb):
        ls, rs = publication_alloc(lb[key], str(key)), publication_alloc(rb[key], str(key))
        require(len(ls) == len(rs), f"allocation sample counts differ: {key}")
        lc, rc = [], []
        for index, (left_sample, right_sample) in enumerate(zip(ls, rs)):
            require(left_sample.get("status") == right_sample.get("status") == "measured",
                    f"publication allocator status differs: {key}/{index}")
            nonnegative_int(left_sample.get("allocation_calls"), f"{key} baseline calls")
            nonnegative_int(right_sample.get("allocation_calls"), f"{key} candidate calls")
            lc.append(left_sample["allocation_calls"])
            rc.append(right_sample["allocation_calls"])
        calls = reduction(numeric_p50(lc, str(key)), numeric_p50(rc, str(key)), str(key))
        passed = calls >= 8.0
        allocation_passed = allocation_passed and passed
        allocation_rows.append({"repeat": key[0], "shape": key[1],
                                "allocation_calls": calls, "passed": passed})
    passed = native_passed and allocation_passed
    return {"native": {"rows": native_rows, "passed": native_passed},
            "allocation": {"rows": allocation_rows, "passed": allocation_passed},
            "passed": passed, "decision": "eligible-for-conditional-lanes" if passed else "reject"}


def validate_conditional_profile(report: dict[str, Any], pilot_passed: bool, plan: dict[str, Any]) -> dict[str, Any]:
    if not pilot_passed:
        return {"status": "unmeasured", "reason": "pilot failed; conditional profile was not admitted"}
    # A profile text/JSON dump is not enough to establish callgrind symbol
    # ownership or the publication boundary.  The root coordinator must add
    # a dedicated raw-profile verifier before this conditional lane can pass.
    paths = [HERE / "raw-profile-verification.json", HERE / "profile-verification.json"]
    existing = next((path for path in paths if path.is_file()), None)
    if existing is None:
        raise IncompleteError("publication profile requires a raw-profile verifier")
    value = read_json(existing, rel(existing))
    require(value.get("status") == "pass" and value.get("plan_sha256") == sha(PLAN)
            and value.get("raw_profile_validated") is True,
            "raw publication profile verifier is not a bound passing report")
    gate = value.get("publication_ir_gate")
    require(isinstance(gate, dict) and gate.get("passed") is True,
            "raw publication profile verifier did not pass the Ir gate")
    reduction_value = gate.get("minimum_reduction_percent")
    finite(reduction_value, "raw publication profile minimum reduction")
    require(float(reduction_value) >= float(plan["gates"]["publication_ir_reduction_percent"]),
            "raw publication profile Ir reduction is below the frozen gate")
    return {"status": "pass", "path": rel(existing), "sha256": sha(existing),
            "raw_profile_validated": True, "minimum_reduction_percent": reduction_value}


def validate_analysis() -> dict[str, Any]:
    plan = load_plan()
    captures = validate_captures()
    path = need(COMPARISON, "comparison report")
    replay_analyzer(path)
    report = read_json(path, "comparison report")
    require(report.get("status") == "pass" and report.get("stage") == "compare"
            and report.get("plan_sha256") == sha(PLAN),
            "comparison envelope differs")
    require(report.get("structured_gates") == plan["gates"],
            "comparison structured gates differ")
    independent = independent_pilot(report, plan)
    pilot = report.get("pilot_admission")
    require(isinstance(pilot, dict) and pilot.get("passed") is independent["passed"],
            "comparison pilot decision does not replay independently")
    require(pilot.get("native", {}).get("passed") is independent["native"]["passed"]
            and pilot.get("allocation", {}).get("passed") is independent["allocation"]["passed"],
            "comparison lane gate decisions differ")
    profile = validate_conditional_profile(report, independent["passed"], plan)
    return {"status": "pass", "comparison_sha256": sha(path),
            "analyzer_sha256": sha(ANALYZER), "independent_pilot": independent,
            "pilot": pilot, "conditional_profile": profile, "captures": captures}


def review_path() -> Path:
    for name in ("adverse-review.json", "flag-review.json", "adverse-flags-review.json"):
        path = HERE / name
        if path.is_file():
            return path
    raise IncompleteError("adverse timing/allocation review is missing")


def validate_flags(analysis: dict[str, Any]) -> dict[str, Any]:
    path = review_path()
    value = read_json(path, rel(path))
    comparison = analysis["pilot"]
    raw = read_json(COMPARISON, "comparison report")
    body = raw.get("comparison")
    require(isinstance(body, dict), "comparison adverse rows are missing")
    adverse = body.get("adverse_flags_over_five_percent")
    drift = body.get("same_build_drift_over_five_percent")
    require(isinstance(adverse, list) and isinstance(drift, list),
            "comparison adverse vectors are malformed")
    require(value.get("comparison_sha256") == analysis["comparison_sha256"],
            "adverse review comparison binding differs")
    matched = value.get("matched", value.get("flags"))
    same = value.get("same_build", [])
    if matched is None:
        combined = value.get("reviews")
        require(isinstance(combined, list) and len(combined) == len(adverse) + len(drift),
                "combined adverse review inventory differs")
        matched, same = combined[:len(adverse)], combined[len(adverse):]
    require(isinstance(matched, list) and isinstance(same, list)
            and len(matched) == len(adverse) and len(same) == len(drift),
            "adverse review counts differ")
    def match(expected: list[Any], actual: list[Any], label: str) -> None:
        remaining = list(actual)
        for row in expected:
            require(isinstance(row, dict), f"{label} source row is malformed")
            index = next((i for i, item in enumerate(remaining)
                          if isinstance(item, dict)
                          and all(item.get(key) == val for key, val in row.items())), None)
            require(index is not None, f"{label} row is unbound")
            require(isinstance(remaining[index].get("review"), str)
                    and remaining[index]["review"].strip(), f"{label} row lacks review")
            remaining.pop(index)
        require(not remaining, f"{label} has unbound rows")
    match(adverse, matched, "adverse")
    match(drift, same, "same-build")
    require(value.get("complete") is True and value.get("all_adverse_metrics_retained") is True,
            "adverse review is incomplete")
    return {"path": rel(path), "sha256": sha(path), "adverse_flags": len(adverse),
            "same_build_flags": len(drift), "pilot": comparison.get("passed")}


def quality_commands() -> list[tuple[str, list[str]]]:
    text = read_text(CHECKS, "checks.py")
    try:
        tree = ast.parse(text, filename=str(CHECKS))
    except SyntaxError as error:
        raise EvidenceError(f"checks.py is invalid: {error}") from error
    value = None
    for node in tree.body:
        targets = node.targets if isinstance(node, ast.Assign) else ([node.target]
                  if isinstance(node, ast.AnnAssign) else [])
        if any(isinstance(target, ast.Name) and target.id == "COMMANDS" for target in targets):
            value = ast.literal_eval(node.value)
            break
    require(isinstance(value, list) and value, "checks.py COMMANDS is missing")
    result = []
    for row in value:
        require(isinstance(row, (list, tuple)) and len(row) == 2
                and isinstance(row[0], str) and isinstance(row[1], list)
                and all(isinstance(token, str) and token for token in row[1]),
                "checks.py command row is malformed")
        result.append((row[0], row[1]))
    require(len({name for name, _ in result}) == len(result), "quality command names repeat")
    return result


def validate_quality(final_source: str) -> dict[str, Any]:
    summary_path = need(HERE / "quality-summary.json", "quality summary")
    summary = read_json(summary_path, "quality-summary.json")
    require(summary.get("status") == "pass" and summary.get("stage") == final_source,
            "quality summary status/stage differs")
    rows = summary.get("checks")
    commands = quality_commands()
    require(isinstance(rows, list) and len(rows) == len(commands),
            "quality check count differs")
    expected_names = {f"check-{name}.receipt.json" for name, _ in commands}
    seen = set()
    total_tests = 0
    stage_folder = HERE / final_source
    stage_manifest_sha = sha(stage_folder / "source-manifest.json")
    for row in rows:
        require(isinstance(row, dict), "quality row is malformed")
        name = row.get("name")
        safe_relative(name, "quality receipt name")
        require(name in expected_names and name not in seen, "quality receipt name differs")
        seen.add(name)
        receipt_path = stage_folder / name
        receipt = receipt_common(receipt_path, final_source, load_plan(), stage_manifest_sha,
                                 stage_manifest_sha)
        value = receipt["value"]
        require(row.get("receipt_sha256") == receipt["sha256"]
                and row.get("exit_code") == value.get("exit_code") == 0,
                f"quality row binding differs: {name}")
        expected_command = ["env", "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
                            "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings"] \
            + dict(commands)[name.removeprefix("check-").removesuffix(".receipt.json")]
        require(value.get("command") == expected_command,
                f"quality command differs: {name}")
        log_path = stage_folder / name.replace(".receipt.json", ".stdout")
        log = read_text(log_path, rel(log_path))
        executed = sum(int(number) for number in re.findall(r"test result: ok\. (\d+) passed;", log))
        require(row.get("executed_tests") == executed, f"quality test count differs: {name}")
        total_tests += executed
    require(seen == expected_names and summary.get("executed_tests") == total_tests,
            "quality aggregate differs")
    return {"path": rel(summary_path), "sha256": sha(summary_path),
            "stage": final_source, "checks": len(rows), "executed_tests": total_tests}


def decision_path() -> Path:
    paths = [HERE / "decision.json", HERE / "disposition.json"]
    existing = [path for path in paths if path.is_file()]
    require(len(existing) == 1, "exactly one decision record is required")
    return existing[0]


def selected_manifest(stage: str) -> Path:
    require(stage in ("baseline", "candidate", "final"), "decision final source is invalid")
    return need(HERE / stage / "source-manifest.json", f"{stage} source manifest")


def validate_decision(analysis: dict[str, Any], flags: dict[str, Any],
                      quality: dict[str, Any]) -> dict[str, Any]:
    path = decision_path()
    value = read_json(path, rel(path))
    require(value.get("schema") == "litchi_0529_pilot_decision_v1",
            "decision schema differs")
    disposition = value.get("disposition")
    require(disposition in ("accepted", "rejected", "rejected_and_reverted"),
            "decision disposition differs")
    final_source = value.get("final_source")
    final_manifest = selected_manifest(final_source)
    if final_source == "final":
        replay_stage("final", load_plan(), allow_empty=True)
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("comparison_sha256") == analysis["comparison_sha256"]
            and value.get("adverse_review_sha256") == flags["sha256"]
            and value.get("quality_summary_sha256") == quality["sha256"]
            and value.get("final_source_manifest_sha256") == sha(final_manifest),
            "decision artifact binding differs")
    current = {name: sha(REPO / name) for name in current_source_names()}
    require(current == source_manifest(final_manifest),
            "current checkout does not match selected final source")
    pilot = analysis["independent_pilot"]["passed"]
    require(value.get("pilot_gates_passed") is pilot,
            "decision pilot gate binding differs")
    lanes = value.get("conditional_lanes")
    require(isinstance(lanes, dict) and set(lanes) == set(CONDITIONAL_LANES),
            "decision conditional lanes are incomplete")
    if not pilot:
        for lane in CONDITIONAL_LANES:
            item = lanes[lane]
            require(isinstance(item, dict) and item.get("status") in
                    ("unmeasured", "unmeasured_pilot_failed")
                    and "pilot" in str(item.get("reason", "")).lower(),
                    f"{lane} is not unmeasured after pilot failure")
    if disposition == "accepted":
        require(pilot and value.get("production_change_retained") is True
                and final_source in ("candidate", "final"),
                "accepted decision lacks retention custody")
        require(all(isinstance(lanes[lane], dict)
                    and lanes[lane].get("status") in ("pass", "measured")
                    for lane in ("profile", "eager")),
                "accepted decision lacks required conditional lanes")
    else:
        require(value.get("production_change_retained") is False
                and final_source == "final",
                "rejected decision selects an invalid source")
        baseline_manifest = source_manifest(BASELINE / "source-manifest.json")
        candidate_manifest = source_manifest(CANDIDATE / "source-manifest.json")
        final_values = source_manifest(final_manifest)
        retained_test = "crates/xml-minifier/tests/stream_audit.rs"
        differences = sorted(name for name in set(baseline_manifest) | set(final_values)
                             if baseline_manifest.get(name) != final_values.get(name))
        require(differences == [retained_test],
                "rejected final source must retain only stream_audit.rs beyond baseline")
        require(final_values.get("crates/xml-minifier/src/audit.rs") ==
                baseline_manifest.get("crates/xml-minifier/src/audit.rs")
                and final_values.get(HARNESS_FILE) == baseline_manifest.get(HARNESS_FILE),
                "rejected final source does not restore XML auditor/harness exactly")
        require(final_values.get(retained_test) == candidate_manifest.get(retained_test),
                "rejected final retained public test is not candidate-bound")
    require(quality["stage"] == final_source
            and isinstance(value.get("review"), str) and value["review"].strip(),
            "decision quality/review binding is incomplete")
    return {"path": rel(path), "sha256": sha(path), "disposition": disposition,
            "final_source": final_source, "pilot_gates_passed": pilot,
            "production_change_retained": value.get("production_change_retained")}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("removed") == plan["owned_paths"]
            and value.get("accessible_process_references") == []
            and value.get("owned_paths_absent") is True
            and value.get("python_cache_absent") is True,
            "cleanup receipt differs")
    require(all(not os.path.lexists(path) for path in plan["owned_paths"]),
            "owned path remains after cleanup")
    require(not list(HERE.rglob("__pycache__")), "Python cache remains in evidence bundle")
    return {"path": rel(CLEANUP), "sha256": sha(CLEANUP),
            "owned_paths_absent": True, "python_cache_absent": True}


def inventory() -> dict[str, str]:
    result = {}
    for path in HERE.rglob("*"):
        require(not path.is_symlink(), f"evidence bundle contains a symlink: {rel(path)}")
        if path.is_file() and path.name != "SHA256SUMS":
            result[rel(path)] = sha(path)
    return result


def validate_seal() -> dict[str, Any]:
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and digest(fields[0]), "SHA256SUMS line is malformed")
        value, name = fields
        safe_relative(name, "seal entry")
        require(name != "SHA256SUMS" and name not in expected,
                "SHA256SUMS entry is unsafe or duplicated")
        expected[name] = value
    require(expected == inventory(), "SHA256SUMS inventory differs")
    return {"path": rel(path), "entries": len(expected), "sha256": sha(path)}


def expect_reject(action: Callable[[], Any], label: str) -> str:
    try:
        action()
    except EvidenceError as error:
        return str(error)
    raise EvidenceError(f"negative probe was accepted: {label}")


def validate_negative_probes() -> dict[str, Any]:
    valid_sample = {
        "status": "measured",
        "scope": "operation_global_system_allocator",
        "allocation_calls": 3,
        "deallocation_calls": 1,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": 16,
        "deallocated_bytes": 8,
        "live_bytes_before": 4,
        "live_bytes_after": 12,
        "peak_live_bytes_before": 4,
        "peak_live_bytes_after": 16,
        "region_peak_live_bytes": 16,
    }

    def invalid_balance() -> None:
        # First prove that the complete fixture is valid, then mutate exactly
        # one counter.  This exercises the balance validator rather than the
        # missing-field/schema branch.
        allocation_sample(valid_sample, "probe valid allocation", True)
        mutated = dict(valid_sample)
        mutated["allocated_bytes"] += 1
        allocation_sample(mutated, "probe invalid allocation", True)

    probes = {
        "unsafe_relative_path": lambda: safe_relative("../escape", "probe path"),
        "inverted_receipt_interval": lambda: interval({
            "start_utc": "2026-01-01T00:00:02+00:00",
            "end_utc": "2026-01-01T00:00:01+00:00", "seconds": 1,
        }, "probe receipt"),
        "publication_vector_cannot_fallback_to_commit": lambda: publication_alloc(
            {"commit_allocation": [{"allocation_calls": 1}]}, "probe row"),
        "invalid_allocation_balance": invalid_balance,
        "incomplete_abba": lambda: validate_abba([{
            "path": "baseline/native-r1-primary-medium.receipt.json",
            "start": "2026-01-01T00:00:00+00:00",
            "end": "2026-01-01T00:00:01+00:00",
            "job": {"name": "native-r1-primary-medium", "kind": "primary"},
        }], {"primary": {"repeats": 1, "shapes": ["medium"]},
             "guard_repeats": 0, "guards": []}),
    }
    checks = [{"case": label, "rejected": True,
               "reason": expect_reject(action, label)} for label, action in probes.items()]
    return {"schema": "litchi-0529-verifier-negative-probes-v1", "status": "pass",
            "checks": checks,
            "scope": "In-memory custody, interval, allocation, publication-channel, and ABBA validator probes."}


def component(name: str) -> dict[str, Any]:
    plan = load_plan()
    if name in ("baseline", "partial"):
        return validate_baseline_partial()
    if name == "source":
        return validate_source()
    if name == "builds":
        return validate_builds()
    if name == "captures":
        return validate_captures()
    if name == "analysis":
        return validate_analysis()
    if name == "flags":
        analysis = validate_analysis()
        return {"analysis": analysis, "review": validate_flags(analysis)}
    if name == "quality":
        decision = read_json(decision_path(), "decision")
        source = decision.get("final_source")
        require(source in ("baseline", "candidate", "final"),
                "quality decision source is invalid")
        return validate_quality(source)
    if name == "decision":
        analysis = validate_analysis()
        flags = validate_flags(analysis)
        value = read_json(decision_path(), "decision")
        source = value.get("final_source")
        quality = validate_quality(source)
        return validate_decision(analysis, flags, quality)
    if name == "cleanup":
        return validate_cleanup(plan)
    if name == "seal":
        return validate_seal()
    if name == "probes":
        return validate_negative_probes()
    raise EvidenceError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    order = {
        "precleanup": ("source", "builds", "captures", "analysis", "flags",
                       "quality", "decision", "probes"),
        "all": ("source", "builds", "captures", "analysis", "flags", "quality",
                "decision", "cleanup", "seal", "probes"),
    }.get(selected, (selected,))
    result: dict[str, Any] = {}
    for name in order:
        try:
            value = component(name)
            component_status = value.get("status") if isinstance(value, dict) else None
            result[name] = {"status": component_status if component_status in
                            ("pass", "incomplete", "fail") else "pass",
                            "result": value}
        except IncompleteError as error:
            result[name] = {"status": "incomplete", "error": str(error)}
        except (EvidenceError, OSError, subprocess.CalledProcessError) as error:
            result[name] = {"status": "fail", "error": str(error)}
    statuses = [item["status"] for item in result.values()]
    status = "fail" if "fail" in statuses else ("incomplete" if "incomplete" in statuses else "pass")
    return {"schema": "litchi-0529-pilot-verification-v1", "status": status,
            "scope": selected, "performance_claim": "none", "components": result,
            "cleanup_seal_pending": result.get("cleanup", {}).get("status") != "pass"
            or result.get("seal", {}).get("status") != "pass"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=("baseline", "partial", "source", "builds", "captures", "analysis",
                                                 "flags", "quality", "decision", "cleanup",
                                                 "seal", "probes", "precleanup", "all"),
                        default="all")
    parser.add_argument("--stage", choices=("baseline", "partial"),
                        help="short form for the baseline/partial evidence view")
    parser.add_argument("--strict", action="store_true",
                        help="return 2 when a selected component is incomplete or failed")
    parser.add_argument("--output", type=Path,
                        help="write JSON to an external path; stdout is always retained")
    args = parser.parse_args()
    if args.output and SEAL.exists() and args.output.resolve().is_relative_to(HERE):
        raise SystemExit("verification output cannot overwrite a sealed bundle")
    selected = args.stage or args.component
    report = run_bundle(selected)
    text = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
    print(text, end="")
    return 2 if args.strict and report["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
