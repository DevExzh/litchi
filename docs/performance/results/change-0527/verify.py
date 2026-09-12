"""Independently verify the 0527 XLSX row-arena pilot evidence.

The verifier is staged because the pilot is still being captured.  Each
component has a direct entry point and reports ``pass``, ``incomplete``, or
``fail`` independently.  An incomplete component never becomes a performance
claim.  The ``precleanup`` component intentionally stops before cleanup and
sealing so it can be run while the owned binaries are still available.
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
import tempfile
from typing import Any, Callable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
FROZEN_INPUTS_PATH = HERE / "frozen-inputs.json"
ADR_PATH = HERE / "adr-manifest.json"
SOURCE_REVIEW_PATH = HERE / "source-review.md"
CHECKS_PATH = HERE / "checks.py"
CLEANUP_SCRIPT_PATH = HERE / "cleanup.py"
ANALYZER_PATH = HERE / "analyze.py"
ANALYZER_HELPER_PATH = HERE.parent / "change-0521" / "analyze.py"
BASELINE_PATH = HERE / "baseline"
CANDIDATE_PATH = HERE / "candidate"
FINAL_PATH = HERE / "final"
PRIOR_BINDING_PATH = HERE.parent / "change-0526" / "source-binding.json"
STORAGE_RECOVERY_PATH = HERE / "storage-recovery.json"
TARGET_PATH = Path("/home/zhuhe/litchi-goal-0527-target")
SCRATCH_PATH = Path("/tmp/litchi-goal-0527")
SOURCE_ROOTS = ("crates/litchi-xlsx/",)
SOURCE_SCOPE_EXACT = frozenset({"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"})
CONDITIONAL_LANES = ("profile", "hardware", "eager")
TIMING_STATS = ("p50", "p95", "p99", "mean")
ALLOCATION_FLAG_FIELDS = ("allocation_calls", "reallocation_calls",
                          "incremental_region_peak_live_bytes")
EXPECTED_NATIVE_SAMPLES = 2440
EXPECTED_ALLOCATOR_SAMPLES = 40
EXPECTED_ANALYZER_HELPER_SHA256 = (
    "322357892b496ef09ca01aa69bb5a8708182a9a771a4541e3b21a6885a7616ad"
)


class EvidenceError(ValueError):
    """A malformed, contradictory, or out-of-scope evidence artifact."""


class IncompleteError(EvidenceError):
    """An artifact needed for this component has not been retained yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def need(path: Path, label: str) -> Path:
    if not path.exists():
        raise IncompleteError(f"{label} is missing: {relative(path)}")
    if path.is_symlink():
        raise EvidenceError(f"{label} must not be a symlink: {relative(path)}")
    return path


def read_json(path: Path) -> Any:
    if not path.exists():
        raise IncompleteError(f"JSON artifact is missing: {relative(path)}")
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {relative(path)}: {error}") from error


def read_text(path: Path) -> str:
    if not path.exists():
        raise IncompleteError(f"text artifact is missing: {relative(path)}")
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {relative(path)}: {error}") from error


def sha(path: Path) -> str:
    if not path.exists():
        raise IncompleteError(f"artifact is missing for hashing: {relative(path)}")
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {relative(path)}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence bundle: {path}") from error


def is_digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a relative path")
    path = Path(value)
    require(".." not in path.parts and path.as_posix() == value,
            f"{label} contains an unsafe path")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} timestamp is not a string")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} timestamp is invalid") from error
    require(parsed.tzinfo is not None, f"{label} timestamp has no timezone")
    return parsed


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def source_name(name: str) -> bool:
    return (name.startswith("crates/") or name.startswith("tools/perf-baseline/")
            or name.startswith(".cargo/") or name in SOURCE_SCOPE_EXACT)


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
    label = relative(path)
    require(isinstance(value, dict) and value, f"{label} is not a nonempty manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label} entry")
        require(source_name(name), f"{label} contains out-of-scope source: {name}")
        require(is_digest(digest), f"{label} digest is malformed: {name}")
        require(name not in result, f"{label} repeats source: {name}")
        result[name] = digest
    return result


def run_checked(args: list[str], *, env: dict[str, str] | None = None,
                input_bytes: bytes | None = None) -> subprocess.CompletedProcess[bytes]:
    try:
        return subprocess.run(args, cwd=REPO, env=env, input=input_bytes,
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              check=True)
    except subprocess.CalledProcessError as error:
        detail = error.stderr.decode(errors="replace")[-2000:]
        raise EvidenceError(f"command failed ({' '.join(args)}): {detail}") from error


def revision_source_names(revision: str) -> set[str]:
    try:
        raw = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", revision],
                                      cwd=REPO, text=True)
    except subprocess.CalledProcessError as error:
        raise EvidenceError(f"frozen source revision cannot be listed: {revision}") from error
    return {name for name in raw.splitlines() if source_name(name)}


def parse_index(index: Path) -> dict[str, tuple[str, int]]:
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    try:
        raw = subprocess.check_output(["git", "ls-files", "-s", "-z"],
                                      cwd=REPO, env=env)
    except subprocess.CalledProcessError as error:
        raise EvidenceError("private Git index cannot be read") from error
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
    try:
        raw = subprocess.check_output(
            ["git", "cat-file", "--batch"], cwd=REPO, env=env,
            input=("\n".join(sorted(oids)) + "\n").encode(),
        )
    except subprocess.CalledProcessError as error:
        raise EvidenceError("private Git index objects cannot be read") from error
    result: dict[str, str] = {}
    position = 0
    while position < len(raw):
        end = raw.find(b"\n", position)
        require(end >= 0, "Git batch response is malformed")
        oid, kind, size = raw[position:end].split()
        require(kind == b"blob", f"Git index object is not a blob: {oid!r}")
        position = end + 1
        length = int(size)
        data = raw[position:position + length]
        require(len(data) == length, "Git batch response is truncated")
        result[oid.decode()] = hashlib.sha256(data).hexdigest()
        position += length
        require(raw[position:position + 1] == b"\n",
                "Git batch response lacks object separator")
        position += 1
    return result


def new_file_sidecars(stage_dir: Path) -> dict[str, str]:
    path = stage_dir / "new-files.json"
    if not path.exists():
        return {}
    value = read_json(path)
    raw = value.get("files") if isinstance(value, dict) else None
    require(isinstance(raw, dict), f"{relative(path)}.files is not an object")
    result: dict[str, str] = {}
    for name, item in raw.items():
        safe_relative(name, f"{relative(path)} file")
        require(any(name.startswith(root) for root in SOURCE_ROOTS),
                f"{relative(path)} file escapes candidate roots: {name}")
        if isinstance(item, str):
            digest, artifact = item, name
        else:
            require(isinstance(item, dict), f"{relative(path)} entry is malformed: {name}")
            digest, artifact = item.get("sha256"), item.get("artifact", name)
        require(is_digest(digest), f"{relative(path)} digest is malformed: {name}")
        safe_relative(artifact, f"{relative(path)} artifact")
        artifact_path = stage_dir / artifact
        require(artifact_path.is_file() and not artifact_path.is_symlink(),
                f"{relative(path)} artifact is missing: {artifact}")
        require(sha(artifact_path) == digest,
                f"{relative(path)} artifact digest differs: {artifact}")
        result[name] = digest
    return result


def replay_source_dir(label: str, stage_dir: Path, plan: dict[str, Any],
                      candidate_like: bool, allow_empty_change: bool = False) -> dict[str, Any]:
    stage_manifest_path = need(stage_dir / "source-manifest.json", f"{label} source manifest")
    stage_patch_path = need(stage_dir / "source.patch", f"{label} source patch")
    stage_manifest = manifest(stage_manifest_path)
    patch = stage_patch_path.read_bytes()
    with tempfile.TemporaryDirectory(prefix="litchi-0527-source-replay-") as directory:
        index = Path(directory) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        run_checked(["git", "read-tree", plan["revision"]], env=env)
        if patch:
            run_checked(["git", "apply", "--cached", "--binary", str(stage_patch_path)],
                        env=env)
        indexed = parse_index(index)
        scoped_index = {name: item for name, item in indexed.items()
                        if source_name(name)}
        names_at_revision = revision_source_names(plan["revision"])
        sidecars = new_file_sidecars(stage_dir)
        missing_from_index = set(stage_manifest) - set(scoped_index)
        if missing_from_index:
            require(candidate_like,
                    f"{label} manifest contains files absent from replay index")
            require(missing_from_index <= set(sidecars),
                    f"{label} new files lack sidecar custody: {sorted(missing_from_index)}")
        for name in set(scoped_index) - set(stage_manifest):
            require(name not in names_at_revision,
                    f"{label} manifest omits indexed source file: {name}")
        hashes = index_blob_hashes(index, {oid for oid, _ in scoped_index.values()})
        for name, digest in stage_manifest.items():
            if name in scoped_index:
                oid, mode = scoped_index[name]
                require(mode in (100644, 100755, 120000),
                        f"{label} source mode is unexpected: {name}")
                require(hashes.get(oid) == digest,
                        f"{label} replay blob differs: {name}")
            else:
                require(sidecars[name] == digest,
                        f"{label} new-file sidecar differs: {name}")
        try:
            changed = set(subprocess.check_output(
                ["git", "diff", "--cached", "--name-only", plan["revision"]],
                cwd=REPO, env=env, text=True).splitlines())
        except subprocess.CalledProcessError as error:
            raise EvidenceError(f"{label} replay diff cannot be read") from error
        require(all(source_name(name) for name in changed),
                f"{label} source patch changes out-of-scope files")
        if candidate_like:
            require(all(any(name.startswith(root) for root in SOURCE_ROOTS)
                        for name in changed),
                    f"{label} source patch escapes XLSX roots")
        else:
            require(not patch and not changed,
                    f"{label} baseline source patch is not empty")
            require(set(stage_manifest) == names_at_revision,
                    f"{label} baseline inventory differs from frozen revision")
        if not allow_empty_change and candidate_like:
            require(changed, f"{label} source patch is empty")
        return {
            "stage": label,
            "manifest_sha256": sha(stage_manifest_path),
            "manifest_entries": len(stage_manifest),
            "patch_sha256": sha(stage_patch_path),
            "replayed_changed_files": sorted(changed),
            "new_files": sorted(sidecars),
        }


def load_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN_INPUTS_PATH)
    require(isinstance(frozen, dict) and set(frozen) == {"plan.json", "run.py"},
            "frozen-inputs inventory differs")
    for name, path in (("plan.json", PLAN_PATH), ("run.py", RUN_PATH)):
        require(frozen.get(name) == sha(path), f"frozen {name} digest differs")
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("status") == "frozen-before-build-and-capture",
            "plan is not frozen before build/capture")
    revision = plan.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"],
                       cwd=REPO, check=True, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as error:
        raise EvidenceError("plan revision is not a Git commit") from error
    require(plan.get("candidate_source_roots") == list(SOURCE_ROOTS),
            "candidate source roots differ")
    require(plan.get("owned_paths") == [str(SCRATCH_PATH), str(TARGET_PATH)],
            "owned temporary paths differ")
    require(plan.get("cpu") == 2, "planned CPU differs")
    require(plan.get("native_order") == [
        "baseline native-r1 before candidate application",
        "candidate native-r1", "candidate native-r2",
        "retained baseline native-r2 under candidate source checkout",
    ], "native ABBA order differs")
    primary = plan.get("primary")
    require(isinstance(primary, dict)
            and primary.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save"
            and primary.get("shapes") == ["medium", "dense-sparse"]
            and primary.get("repeats") == 2 and primary.get("warmup") == 20
            and primary.get("samples") == 200, "primary plan differs")
    guards = plan.get("guards")
    require(isinstance(guards, list) and len(guards) == 5, "guard plan differs")
    require(guards == [
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
    ], "guard cases differ")
    require(plan.get("guard_samples") == 30 and plan.get("guard_warmup") == 10
            and plan.get("guard_repeats") == 2, "guard counts differ")
    allocation = plan.get("allocation")
    require(isinstance(allocation, dict)
            and allocation.get("shapes") == ["medium", "dense-sparse"]
            and allocation.get("repeats") == 2 and allocation.get("warmup") == 0
            and allocation.get("samples") == 5, "allocation plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict)
            and profile.get("shapes") == ["medium", "dense-sparse"]
            and profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 1, "profile plan differs")
    gates = plan.get("gates")
    require(gates == {
        "total_p50_reduction_percent": 2.0,
        "commit_p50_reduction_percent": 5.0,
        "total_mean_reduction_percent": 2.0,
        "allocation_calls_reduction_percent": 8.0,
        "commit_ir_reduction_percent": 3.0,
        "require_every_shape_repeat": True,
    }, "structured pilot gates differ")
    conditional = plan.get("conditional_lanes")
    require(isinstance(conditional, dict)
            and set(conditional) == set(CONDITIONAL_LANES),
            "conditional lane inventory differs")
    require(isinstance(plan.get("draft_patch_sha256"), str)
            and is_digest(plan["draft_patch_sha256"]), "draft patch binding is malformed")
    return plan


def validate_adr_and_reviews(plan: dict[str, Any]) -> dict[str, Any]:
    adr = read_json(ADR_PATH)
    require(isinstance(adr, dict) and isinstance(adr.get("files"), dict),
            "ADR manifest is malformed")
    for name, digest in adr["files"].items():
        safe_relative(name, "ADR path")
        require(is_digest(digest) and (REPO / name).is_file()
                and sha(REPO / name) == digest,
                f"ADR binding differs: {name}")
    require(is_digest(adr.get("prior_audit_source_binding_sha256"))
            and PRIOR_BINDING_PATH.is_file()
            and sha(PRIOR_BINDING_PATH) == adr["prior_audit_source_binding_sha256"],
            "prior 0526 source binding differs")
    review = read_text(SOURCE_REVIEW_PATH)
    require("0527" in review and "candidate" in review.lower()
            and plan["draft_patch_sha256"] in review,
            "source review is not bound to the frozen draft")
    draft = HERE.parent / "change-0526" / "row-primary-arena.patch"
    require(draft.is_file() and sha(draft) == plan["draft_patch_sha256"],
            "frozen draft patch differs")
    for name in (
        "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs",
        "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs",
        "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs",
    ):
        require(name in review, f"source review omits draft file: {name}")
    return {"adr_manifest_sha256": sha(ADR_PATH),
            "prior_binding_sha256": sha(PRIOR_BINDING_PATH),
            "source_review": relative(SOURCE_REVIEW_PATH),
            "draft_patch_sha256": sha(draft)}


def validate_source() -> dict[str, Any]:
    plan = load_plan()
    reviews = validate_adr_and_reviews(plan)
    baseline = replay_source_dir("baseline", BASELINE_PATH, plan, False)
    baseline_manifest = manifest(BASELINE_PATH / "source-manifest.json")
    require(sum(name.startswith("crates/litchi-xlsx/src/")
                for name in baseline_manifest) == 451,
            "baseline XLSX source inventory is not 451 files")
    require(CANDIDATE_PATH.is_dir() and not CANDIDATE_PATH.is_symlink(),
            "candidate source stage is missing")
    candidate = replay_source_dir("candidate", CANDIDATE_PATH, plan, True)
    before = baseline_manifest
    after = manifest(CANDIDATE_PATH / "source-manifest.json")
    differences = sorted(name for name in set(before) | set(after)
                         if before.get(name) != after.get(name))
    require(differences, "candidate source difference is empty")
    require(all(any(name.startswith(root) for root in SOURCE_ROOTS)
                for name in differences), "candidate difference escapes XLSX root")
    source_diff_path = CANDIDATE_PATH / "source-diff.json"
    require(source_diff_path.is_file(), "candidate source-diff.json is missing")
    source_diff = read_json(source_diff_path)
    require(source_diff.get("baseline_manifest_sha256") == baseline["manifest_sha256"]
            and source_diff.get("candidate_manifest_sha256") == candidate["manifest_sha256"]
            and source_diff.get("candidate_source_roots") == list(SOURCE_ROOTS),
            "candidate source-diff binding differs")
    changed = source_diff.get("changed_files")
    require(isinstance(changed, dict) and sorted(changed) == differences,
            "candidate source-diff file set differs")
    for name in differences:
        row = changed.get(name)
        require(isinstance(row, dict)
                and row.get("baseline_sha256") == before.get(name)
                and row.get("candidate_sha256") == after.get(name),
                f"candidate source-diff digest differs: {name}")
    require(set(candidate["replayed_changed_files"]) == set(differences) -
            set(candidate["new_files"]), "candidate replay differs from manifest diff")
    final = None
    if FINAL_PATH.is_dir():
        final = replay_source_dir("final", FINAL_PATH, plan, True,
                                  allow_empty_change=True)
    return {"status": "pass", "plan_sha256": sha(PLAN_PATH),
            "run_sha256": sha(RUN_PATH), "reviews": reviews,
            "baseline": baseline, "candidate": candidate,
            "candidate_differences": differences, "final": final}


def load_analyzer() -> tuple[Any, Any]:
    need(ANALYZER_PATH, "canonical 0527 analyzer")
    spec = importlib.util.spec_from_file_location("litchi_0527_analyzer", ANALYZER_PATH)
    require(spec is not None and spec.loader is not None,
            "canonical 0527 analyzer cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    base = getattr(module, "BASE", None)
    require(base is not None, "canonical 0527 analyzer has no retained base helper")
    require(ANALYZER_HELPER_PATH.is_file()
            and sha(ANALYZER_HELPER_PATH) == EXPECTED_ANALYZER_HELPER_SHA256,
            "canonical analyzer helper differs")
    return module, base


def capture_job_specs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    primary = plan["primary"]
    jobs: list[dict[str, Any]] = []
    if lane == "native":
        for repeat in range(1, primary["repeats"] + 1):
            for shape in primary["shapes"]:
                jobs.append({"name": f"native-r{repeat}-primary-{shape}",
                             "kind": "primary", "guard": None, "repeat": repeat,
                             "case": primary["case"], "shape": shape,
                             "warmup": primary["warmup"], "samples": primary["samples"]})
        for repeat in range(1, plan["guard_repeats"] + 1):
            for index, guard in enumerate(plan["guards"]):
                for shape in guard["shapes"]:
                    jobs.append({"name": f"native-r{repeat}-guard{index}-{shape}",
                                 "kind": "guard", "guard": index, "repeat": repeat,
                                 "case": guard["case"], "shape": shape,
                                 "warmup": plan["guard_warmup"],
                                 "samples": plan["guard_samples"]})
    else:
        config = plan["allocation"]
        for repeat in range(1, config["repeats"] + 1):
            for shape in config["shapes"]:
                jobs.append({"name": f"alloc-r{repeat}-{shape}",
                             "kind": "allocation", "guard": None, "repeat": repeat,
                             "case": primary["case"], "shape": shape,
                             "warmup": config["warmup"], "samples": config["samples"]})
    return jobs


def expected_capture_artifacts(name: str) -> set[str]:
    if name.startswith("native-"):
        return {name + suffix for suffix in (".json", ".stdout", ".stderr", ".rss.json")}
    return {name + suffix for suffix in (".json", ".stdout", ".stderr")}


def expected_command(stage: str, job: dict[str, Any], binary_path: str) -> list[str]:
    output = str(HERE / stage / (job["name"] + ".json"))
    base = ["taskset", "-c", "2"]
    if job["kind"] == "primary" or job["kind"] == "guard":
        base += ["/usr/bin/time", "-f",
                 '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                 "-o", str(HERE / stage / (job["name"] + ".rss.json"))]
    return base + [binary_path, "--warmup", str(job["warmup"]),
                   "--samples", str(job["samples"]), "--case", job["case"],
                   "--xlsx-cell-crud-shape", job["shape"], "--json", output]


def build_command(kind: str, plan: dict[str, Any]) -> list[str]:
    executable = "litchi-perf-baseline-alloc" if kind == "alloc" else "litchi-perf-baseline"
    command = ["env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
               "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
               "--bin", executable, "--target-dir", plan["owned_paths"][1]]
    if kind == "alloc":
        command += ["--features", "allocator-metrics"]
    return command


def validate_artifacts(path: Path, names: set[str], label: str) -> None:
    value = read_json(path)
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict), f"{label} artifacts are not an object")
    require(set(artifacts) == names, f"{label} artifact inventory differs")
    for name, digest in artifacts.items():
        safe_relative(name, f"{label} artifact")
        require(Path(name).name == name, f"{label} artifact is not stage-local: {name}")
        artifact = path.parent / name
        require(artifact.is_file() and not artifact.is_symlink()
                and is_digest(digest) and sha(artifact) == digest,
                f"{label} artifact custody differs: {name}")


def expected_working_manifest(stage: str, name: str,
                              stage_sha: str, candidate_sha: str) -> str:
    if stage == "candidate":
        return candidate_sha
    if stage == "baseline" and name.startswith("native-r2-"):
        return candidate_sha
    return stage_sha


def validate_receipt(path: Path, stage: str, plan: dict[str, Any],
                     stage_sha: str, candidate_sha: str,
                     expected_binary: str | None = None,
                     expected_artifacts: set[str] | None = None) -> dict[str, Any]:
    value = read_json(path)
    label = relative(path)
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    finite_number(seconds, f"{label}.seconds")
    require(end > start and seconds > 0, f"{label} interval is invalid")
    exit_code = value.get("exit_code")
    require(isinstance(exit_code, int) and not isinstance(exit_code, bool),
            f"{label} exit code is invalid")
    require(value.get("plan_sha256") == sha(PLAN_PATH)
            and value.get("script_sha256") == sha(RUN_PATH),
            f"{label} frozen plan/driver binding differs")
    require(value.get("source_manifest_sha256") == stage_sha
            and value.get("working_source_manifest_sha256") ==
            expected_working_manifest(stage, path.name.removesuffix(".receipt.json"),
                                      stage_sha, candidate_sha),
            f"{label} source/working manifest binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict)
            and environment.get("TMPDIR") == str(TARGET_PATH / "test-tmp"),
            f"{label} TMPDIR binding differs")
    if expected_binary is not None:
        require(value.get("binary_sha256") == expected_binary,
                f"{label} binary binding differs")
    if expected_artifacts is not None:
        validate_artifacts(path, expected_artifacts, label)
    else:
        artifacts = value.get("artifacts")
        require(isinstance(artifacts, dict), f"{label} artifact inventory is missing")
        receipt_name = path.name.removesuffix(".receipt.json")
        require({f"{receipt_name}.stdout", f"{receipt_name}.stderr"} <= set(artifacts),
                f"{label} failed receipt lacks retained logs")
        validate_artifacts(path, set(artifacts), label)
    return {"path": label, "name": path.name.removesuffix(".receipt.json"),
            "start_utc": value["start_utc"], "end_utc": value["end_utc"],
            "exit_code": exit_code, "sha256": sha(path)}


def validate_storage_recovery(plan: dict[str, Any], normal_identity: dict[str, Any],
                              build_receipt: Path) -> dict[str, Any] | None:
    scratch = Path(plan["owned_paths"][0])
    if not STORAGE_RECOVERY_PATH.is_file():
        require(not scratch.is_symlink(), "scratch symlink lacks recovery receipt")
        return None
    if scratch.exists() or scratch.is_symlink():
        require(scratch.is_symlink()
                and scratch.resolve() == TARGET_PATH / "retained-binaries",
                "scratch symlink does not resolve to the owned retained-binaries directory")
    else:
        validate_cleanup(plan)
    value = read_json(STORAGE_RECOVERY_PATH)
    require(value.get("scratch") == str(scratch)
            and value.get("symlink_target") == str(TARGET_PATH / "retained-binaries")
            and value.get("frozen_driver_unchanged") is True
            and value.get("build_receipt_sha256") == sha(build_receipt)
            and value.get("recovered_binary_sha256") == normal_identity["sha256"],
            "storage recovery custody differs")
    recovered = TARGET_PATH / "retained-binaries" / "baseline-normal"
    if scratch.is_symlink():
        require(recovered.is_file() and not recovered.is_symlink()
                and sha(recovered) == normal_identity["sha256"],
                "recovered baseline binary differs")
    return {"path": relative(STORAGE_RECOVERY_PATH), "sha256": sha(STORAGE_RECOVERY_PATH),
            "target": str(recovered)}


def validate_binary_identity(stage: str, kind: str, plan: dict[str, Any],
                             stage_sha: str) -> dict[str, Any]:
    path = HERE / stage / f"binary-{kind}.json"
    value = read_json(path)
    require(isinstance(value, dict) and is_digest(value.get("sha256")),
            f"{relative(path)} binary identity is malformed")
    expected_path = SCRATCH_PATH / f"{stage}-{'alloc' if kind == 'alloc' else 'normal'}"
    require(value.get("path") == str(expected_path)
            and value.get("source_manifest_sha256") == stage_sha,
            f"{relative(path)} binary path/source binding differs")
    nonnegative_integer(value.get("bytes"), f"{relative(path)}.bytes")
    require(value["bytes"] > 0, f"{relative(path)} binary is empty")
    receipt_path = HERE / stage / f"build-{kind}.receipt.json"
    need(receipt_path, f"{relative(path)} build receipt")
    require(value.get("build_receipt_sha256") == sha(receipt_path),
            f"{relative(path)} build receipt binding differs")
    if expected_path.exists():
        require(expected_path.is_file() and not expected_path.is_symlink()
                and sha(expected_path) == value["sha256"]
                and expected_path.stat().st_size == value["bytes"],
                f"{relative(path)} live binary differs")
    else:
        cleanup = HERE / "cleanup.json"
        if not cleanup.is_file():
            raise IncompleteError(f"{relative(path)} binary is absent before cleanup")
        cleanup_value = read_json(cleanup)
        require(cleanup_value.get("owned_paths_absent") is True
                and cleanup_value.get("removed") == plan["owned_paths"],
                f"{relative(path)} absent binary lacks cleanup custody")
    return {"kind": kind, "path": str(expected_path), "sha256": value["sha256"],
            "bytes": value["bytes"], "identity_sha256": sha(path),
            "build_receipt_sha256": value["build_receipt_sha256"]}


def validate_build_receipt(stage: str, kind: str, plan: dict[str, Any],
                           stage_sha: str, candidate_sha: str) -> dict[str, Any]:
    path = HERE / stage / f"build-{kind}.receipt.json"
    row = validate_receipt(path, stage, plan, stage_sha, candidate_sha,
                           expected_binary=None,
                           expected_artifacts={f"build-{kind}.stdout", f"build-{kind}.stderr"})
    value = read_json(path)
    require(value.get("exit_code") == 0 and value.get("binary_sha256") is None
            and value.get("command") == build_command(kind, plan),
            f"{relative(path)} build command/result differs")
    return row


def validate_builds() -> dict[str, Any]:
    plan = load_plan()
    source_info = validate_source()
    candidate_sha = source_info["candidate"]["manifest_sha256"]
    stages: dict[str, Any] = {}
    for stage in ("baseline", "candidate"):
        stage_sha = sha(HERE / stage / "source-manifest.json")
        normal_receipt = validate_build_receipt(stage, "normal", plan, stage_sha, candidate_sha)
        alloc_receipt = validate_build_receipt(stage, "alloc", plan, stage_sha, candidate_sha)
        normal = validate_binary_identity(stage, "normal", plan, stage_sha)
        allocator = validate_binary_identity(stage, "alloc", plan, stage_sha)
        recovery = (validate_storage_recovery(
            plan, normal, HERE / stage / "build-normal.receipt.json")
                    if stage == "baseline" else None)
        stages[stage] = {"manifest_sha256": stage_sha,
                         "normal": normal, "allocator": allocator,
                         "build_receipts": [normal_receipt, alloc_receipt],
                         "storage_recovery": recovery}
    return {"status": "pass", "stages": stages}


def validate_stage_receipts(stage: str, plan: dict[str, Any],
                            candidate_sha: str) -> dict[str, Any]:
    folder = HERE / stage
    stage_sha = sha(folder / "source-manifest.json")
    normal_identity = validate_binary_identity(stage, "normal", plan, stage_sha)
    allocator_identity = validate_binary_identity(stage, "alloc", plan, stage_sha)
    jobs = capture_job_specs(plan, "native") + capture_job_specs(plan, "alloc")
    expected_names = {job["name"] for job in jobs}
    actual_names = {path.name.removesuffix(".receipt.json") for path in folder.glob("*.receipt.json")
                    if path.name.startswith(("native-", "alloc-"))}
    require(actual_names == expected_names,
            f"{stage} capture receipt matrix differs: {sorted(expected_names ^ actual_names)}")
    rows: list[dict[str, Any]] = []
    receipts: list[dict[str, Any]] = []
    for job in jobs:
        receipt_path = folder / f"{job['name']}.receipt.json"
        binary = allocator_identity if job["kind"] == "allocation" else normal_identity
        row = validate_receipt(receipt_path, stage, plan, stage_sha, candidate_sha,
                               expected_binary=binary["sha256"],
                               expected_artifacts=expected_capture_artifacts(job["name"]))
        receipt = read_json(receipt_path)
        require(receipt.get("command") == expected_command(stage, job, binary["path"]),
                f"{relative(receipt_path)} command differs")
        require(receipt.get("exit_code") == 0,
                f"required {relative(receipt_path)} failed")
        row["job"] = job
        rows.append(row)
        receipts.append(row)
    return {"stage": stage, "manifest_sha256": stage_sha,
            "normal": normal_identity, "allocator": allocator_identity,
            "rows": rows, "receipts": receipts}


def validate_preflights(plan: dict[str, Any], candidate_sha: str) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for folder in sorted(HERE.iterdir()):
        if not folder.is_dir() or not re.fullmatch(r"preflight(?:-[0-9]+)?", folder.name):
            continue
        replay = replay_source_dir(folder.name, folder, plan, True)
        stage_sha = replay["manifest_sha256"]
        receipts = []
        for path in sorted(folder.glob("*.receipt.json")):
            receipt = validate_receipt(path, folder.name, plan, stage_sha, candidate_sha)
            receipts.append(receipt)
        result.append({"stage": folder.name, "manifest_sha256": stage_sha,
                       "candidate_manifest_equal": stage_sha == candidate_sha,
                       "replay": replay, "receipts": receipts,
                       "failed_receipts": [row["path"] for row in receipts
                                           if row["exit_code"] != 0]})
    return result


def validate_global_intervals(rows: list[dict[str, Any]]) -> None:
    ordered = sorted(rows, key=lambda row: parse_time(row["start_utc"], row["path"]))
    for left, right in zip(ordered, ordered[1:]):
        require(parse_time(left["end_utc"], left["path"])
                <= parse_time(right["start_utc"], right["path"]),
                f"receipt intervals overlap: {left['path']} and {right['path']}")


def validate_native_order(plan: dict[str, Any], rows: list[dict[str, Any]]) -> dict[str, Any]:
    native_names = {job["name"] for job in capture_job_specs(plan, "native")}
    native = [row for row in rows if row["name"] in native_names]
    require(len(native) == len(native_names) * 2,
            "native ABBA receipt matrix is incomplete")
    ordered = sorted(native, key=lambda row: parse_time(row["start_utc"], row["path"]))
    blocks: list[tuple[str, int]] = []
    by_block: dict[tuple[str, int], set[str]] = {}
    for row in ordered:
        match = re.match(r"^native-r([12])-", row["name"])
        require(match is not None, f"native receipt name is malformed: {row['name']}")
        block = (Path(row["path"]).parts[-2], int(match.group(1)))
        if not blocks or blocks[-1] != block:
            blocks.append(block)
        by_block.setdefault(block, set()).add(row["name"])
    expected = [("baseline", 1), ("candidate", 1), ("candidate", 2), ("baseline", 2)]
    require(blocks == expected, f"native ABBA blocks differ: {blocks}")
    expected_by_repeat = {
        repeat: {job["name"] for job in capture_job_specs(plan, "native")
                 if job["name"].startswith(f"native-r{repeat}-")}
        for repeat in (1, 2)
    }
    for stage, repeat in expected:
        require(by_block.get((stage, repeat)) == expected_by_repeat[repeat],
                f"native {stage} r{repeat} matrix is incomplete")
    return {"blocks": [f"{stage}/native-r{repeat}" for stage, repeat in blocks],
            "planned_jobs_per_block": len(expected_by_repeat[1]),
            "retained_baseline_a2": True}


def load_base_analyzer() -> Any:
    module, base = load_analyzer()
    require(getattr(module, "HERE", None) == HERE,
            "0527 analyzer is not bound to this evidence root")
    require(getattr(module, "BASE", None) is base, "0527 analyzer base binding differs")
    return base


def validate_captures() -> dict[str, Any]:
    plan = load_plan()
    source_info = validate_source()
    # Capture receipts bind to built binaries.  Re-run the lightweight build
    # custody checks here so a captures-only invocation cannot skip receipt,
    # identity, or recovered-storage validation.
    builds = validate_builds()
    candidate_sha = source_info["candidate"]["manifest_sha256"]
    baseline = validate_stage_receipts("baseline", plan, candidate_sha)
    candidate = validate_stage_receipts("candidate", plan, candidate_sha)
    preflights = validate_preflights(plan, candidate_sha)
    all_rows = baseline["receipts"] + candidate["receipts"]
    for item in preflights:
        all_rows.extend(item["receipts"])
    validate_global_intervals(all_rows)
    native_order = validate_native_order(plan, all_rows)
    base = load_base_analyzer()
    # This canonical helper validates every raw result, every phase vector,
    # sink/output identity, and allocator balance while preserving the root's
    # recovery symlink semantics through the final component path.
    baseline_analysis = base.check_stage("baseline", plan)
    candidate_analysis = base.check_stage("candidate", plan)
    require(baseline_analysis["native"]["total_samples"] == 1220
            and candidate_analysis["native"]["total_samples"] == 1220,
            "native per-stage sample ledger differs")
    require(baseline_analysis["allocation"]["total_samples"] == 20
            and candidate_analysis["allocation"]["total_samples"] == 20,
            "allocator per-stage sample ledger differs")
    return {"status": "pass", "baseline": {**baseline, "analysis": baseline_analysis},
            "candidate": {**candidate, "analysis": candidate_analysis},
            "builds": builds,
            "preflights": preflights, "all_receipts": all_rows,
            "native_order": native_order, "native_samples": EXPECTED_NATIVE_SAMPLES,
            "allocator_samples": EXPECTED_ALLOCATOR_SAMPLES}


def replay_analyzer(expected_path: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0527-analyzer-replay-") as directory:
        output = Path(directory) / expected_path.name
        result = subprocess.run(["python3", "-B", str(ANALYZER_PATH), "--output", str(output)],
                                cwd=REPO, text=True, capture_output=True)
        require(result.returncode == 0,
                f"canonical analyzer failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == expected_path.read_bytes(),
                "canonical analyzer report does not replay byte-for-byte")


def validate_analysis() -> dict[str, Any]:
    plan = load_plan()
    captures = validate_captures()
    comparison_path = need(HERE / "comparison.json", "canonical comparison")
    replay_analyzer(comparison_path)
    value = read_json(comparison_path)
    require(value.get("status") == "pass" and value.get("stage") == "compare"
            and value.get("plan_sha256") == sha(PLAN_PATH),
            "comparison envelope differs")
    require(value.get("numerical_verifier") == {
        "path": "change-0521/analyze.py", "sha256": EXPECTED_ANALYZER_HELPER_SHA256
    }, "comparison canonical helper binding differs")
    require(value.get("structured_gates") == plan["gates"],
            "comparison structured gate binding differs")
    for stage in ("baseline", "candidate"):
        # BASE.analyze("compare") serializes each stage report directly under
        # its stage key; the single-stage form uses an ``evidence`` wrapper.
        # Compare against the direct canonical stage object here rather than
        # assuming the single-stage envelope.
        evidence = value.get(stage)
        expected = captures[stage]["analysis"]
        require(evidence == expected, f"comparison {stage} evidence differs from retained raw reports")
        require(evidence["native"]["total_samples"] == 1220
                and evidence["allocation"]["total_samples"] == 20,
                f"comparison {stage} sample ledger differs")
    base = load_base_analyzer()
    computed = base.compare_stages(captures["baseline"]["analysis"],
                                   captures["candidate"]["analysis"])
    require(value.get("comparison") == computed,
            "comparison payload differs from canonical recomputation")
    pilot = value.get("pilot_admission")
    require(isinstance(pilot, dict), "pilot admission is missing")
    adapter, _ = load_analyzer()
    expected_pilot = adapter._pilot_admission({
        "baseline": captures["baseline"]["analysis"],
        "candidate": captures["candidate"]["analysis"],
    }, plan)
    require(pilot == expected_pilot, "pilot admission does not replay from raw reports")
    require(pilot.get("native", {}).get("passed") is
            all(row.get("passed") is True for row in pilot.get("native", {}).get("rows", [])),
            "pilot native gate does not reconcile")
    require(pilot.get("allocation", {}).get("passed") is
            all(row.get("passed") is True for row in pilot.get("allocation", {}).get("rows", [])),
            "pilot allocation gate does not reconcile")
    return {"status": "pass", "comparison_sha256": sha(comparison_path),
            "analyzer_sha256": sha(ANALYZER_PATH),
            "native_samples": EXPECTED_NATIVE_SAMPLES,
            "allocator_samples": EXPECTED_ALLOCATOR_SAMPLES,
            "pilot": pilot, "comparison": value,
            "captures": captures}


def review_path() -> Path:
    for name in ("adverse-review.json", "flag-review.json", "adverse-flags-review.json"):
        path = HERE / name
        if path.is_file():
            return path
    raise IncompleteError("adverse timing/allocation review is missing")


def review_rows(rows: Any, label: str) -> list[dict[str, Any]]:
    require(isinstance(rows, list), f"{label} review rows are missing")
    result: list[dict[str, Any]] = []
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("review"), str)
                and row["review"].strip(), f"{label} row has no review")
        result.append(row)
    return result


def match_review(expected: list[Any], actual: list[dict[str, Any]], label: str) -> None:
    require(len(expected) == len(actual), f"{label} review count differs")
    remaining = list(actual)
    for source_row in expected:
        require(isinstance(source_row, dict), f"{label} comparison row is malformed")
        index = next((index for index, row in enumerate(remaining)
                      if all(row.get(key) == item for key, item in source_row.items())), None)
        require(index is not None, f"{label} review row is unbound")
        remaining.pop(index)
    require(not remaining, f"{label} review has unbound rows")


def allocation_adverse_flags(comparison: dict[str, Any]) -> list[dict[str, Any]]:
    """Derive the planned incremental-allocation adverse rows.

    The retained 0521 comparison emits allocator vectors but only emits
    timing/RSS adverse rows.  The 0527 plan also makes incremental region peak
    live bytes reviewable, so derive that one field independently from the
    canonical allocator vectors.  Other allocator counters are diagnostics
    and are not silently promoted to adverse flags.
    """

    records = comparison.get("allocation_comparisons")
    require(isinstance(records, list), "allocation comparison rows are malformed")
    result: list[dict[str, Any]] = []
    for record in records:
        require(isinstance(record, dict), "allocation comparison row is malformed")
        metrics = record.get("metrics")
        require(isinstance(metrics, dict), "allocation comparison metrics are malformed")
        metric = metrics.get("incremental_region_peak_live_bytes")
        require(isinstance(metric, dict),
                "incremental allocation peak comparison is missing")
        baseline = metric.get("baseline")
        candidate = metric.get("candidate")
        require(isinstance(baseline, list) and isinstance(candidate, list)
                and baseline and len(baseline) == len(candidate),
                "incremental allocation peak vectors differ")
        for index, (left, right) in enumerate(zip(baseline, candidate)):
            nonnegative_integer(left, f"allocation baseline peak sample {index}")
            nonnegative_integer(right, f"allocation candidate peak sample {index}")
        ordered_left = sorted(baseline)
        ordered_right = sorted(candidate)
        midpoint_left = (ordered_left[(len(ordered_left) - 1) // 2]
                         + ordered_left[len(ordered_left) // 2]) / 2.0
        midpoint_right = (ordered_right[(len(ordered_right) - 1) // 2]
                          + ordered_right[len(ordered_right) // 2]) / 2.0
        if midpoint_left == 0.0:
            change = 0.0 if midpoint_right == 0.0 else None
        else:
            change = (midpoint_right / midpoint_left - 1.0) * 100.0
        if change is not None and change > 5.0:
            result.append({
                "lane": "allocation",
                "case": record.get("case"),
                "shape": record.get("shape"),
                "repeat": record.get("repeat"),
                "phase": "incremental_region_peak_live_bytes",
                "stat": "p50",
                "baseline": midpoint_left,
                "candidate": midpoint_right,
                "change_percent": change,
            })
    return result


def validate_flags(analysis: dict[str, Any]) -> dict[str, Any]:
    path = review_path()
    value = read_json(path)
    comparison = analysis["comparison"]["comparison"]
    require(value.get("comparison_sha256") == analysis["comparison_sha256"],
            "adverse review comparison binding differs")
    adverse = comparison.get("adverse_flags_over_five_percent")
    drift = comparison.get("same_build_drift_over_five_percent")
    require(isinstance(adverse, list) and isinstance(drift, list),
            "comparison adverse vectors are malformed")
    # The canonical helper owns timing/RSS and same-build rows.  Append the
    # independently derived 0527 incremental-allocation rows before matching
    # retained reviews so a new >5% allocation peak cannot be omitted.
    adverse = adverse + allocation_adverse_flags(comparison)
    if isinstance(value.get("matched"), list):
        matched = review_rows(value["matched"], "matched")
        same = review_rows(value.get("same_build"), "same-build")
    elif isinstance(value.get("flags"), list):
        matched = review_rows(value["flags"], "flags")
        same = review_rows(value.get("same_build", []), "same-build")
    else:
        combined = review_rows(value.get("reviews"), "combined")
        require(len(combined) == len(adverse) + len(drift),
                "combined adverse review count differs")
        matched, same = combined[:len(adverse)], combined[len(adverse):]
    match_review(adverse, matched, "adverse")
    match_review(drift, same, "same-build drift")
    require(value.get("complete") is True
            and value.get("all_adverse_metrics_retained") is True,
            "adverse review is incomplete")
    return {"path": relative(path), "sha256": sha(path),
            "adverse_flags": len(adverse), "same_build_flags": len(drift)}


def validate_conditional_artifacts(pilot_passed: bool) -> dict[str, Any]:
    present: dict[str, str] = {}
    candidates = {
        "profile": (HERE / "profile-analysis.json", HERE / "profile-comparison.json"),
        "hardware": (HERE / "hardware-analysis.json",),
        "eager": (HERE / "eager-guard-comparison.json",),
    }
    for lane, paths in candidates.items():
        existing = next((path for path in paths if path.is_file()), None)
        if existing is None:
            present[lane] = "unmeasured"
            continue
        value = read_json(existing)
        require(isinstance(value, dict) and value.get("status") == "pass",
                f"{lane} conditional artifact is not a passing bound report")
        present[lane] = "measured"
    if not pilot_passed:
        # Presence does not turn these lanes into an admission.  The decision
        # must still state that every conditional lane is unmeasured because
        # the pilot failed.
        return {lane: {"status": status,
                       "reason": "pilot failed; conditional lane is not admitted"}
                for lane, status in present.items()}
    return {lane: {"status": status,
                   "reason": "pilot passed; lane requires its own gate"}
            for lane, status in present.items()}


def decision_path() -> Path:
    paths = [HERE / "decision.json", HERE / "disposition.json"]
    existing = [path for path in paths if path.is_file()]
    require(len(existing) == 1, "exactly one decision/disposition record is required")
    return existing[0]


def validate_quality_commands() -> list[tuple[str, list[str]]]:
    need(CHECKS_PATH, "checks.py")
    try:
        tree = ast.parse(read_text(CHECKS_PATH), filename=str(CHECKS_PATH))
    except SyntaxError as error:
        raise EvidenceError(f"checks.py is invalid: {error}") from error
    value = None
    for node in tree.body:
        targets: list[ast.expr] = []
        if isinstance(node, ast.Assign):
            targets = node.targets
        elif isinstance(node, ast.AnnAssign):
            targets = [node.target]
        if any(isinstance(target, ast.Name) and target.id == "COMMANDS"
               for target in targets):
            value = ast.literal_eval(node.value)
            break
    require(isinstance(value, list) and len(value) == 12,
            "checks.py must contain twelve frozen quality checks")
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


def quality_row_stage_name(row: dict[str, Any], final_stage: str) -> tuple[str, str]:
    name = row.get("name")
    safe_relative(name, "quality check name")
    stage = row.get("stage")
    if stage is None and "/" in name:
        stage = name.split("/", 1)[0]
    if stage is None:
        stage = final_stage
    require(isinstance(stage, str) and stage and "/" not in stage
            and ".." not in Path(stage).parts, "quality stage is unsafe")
    if "/" in name:
        prefix, name = name.split("/", 1)
        require(prefix == stage, "quality stage/name disagree")
    require(name.startswith("check-") and name.endswith(".receipt.json"),
            f"quality receipt name is unexpected: {name}")
    return stage, name


def validate_quality(final_source: str, candidate_sha: str,
                     preflights: list[dict[str, Any]]) -> dict[str, Any]:
    path = need(HERE / "quality-summary.json", "quality summary")
    value = read_json(path)
    require(value.get("status") == "pass"
            and value.get("stage", final_source) == final_source,
            "quality summary status/final stage differs")
    checks = value.get("checks")
    require(isinstance(checks, list) and len(checks) == 12,
            "quality summary check count differs")
    commands = validate_quality_commands()
    expected_names = {f"check-{name}.receipt.json" for name, _ in commands}
    seen: set[str] = set()
    tests = 0
    preflight_map = {item["stage"]: item for item in preflights}
    for row in checks:
        require(isinstance(row, dict), "quality row is not an object")
        stage, name = quality_row_stage_name(row, final_source)
        key = f"{stage}/{name}"
        require(key not in seen, f"quality row repeats: {key}")
        seen.add(key)
        if stage == final_source:
            stage_manifest_sha = sha(HERE / stage / "source-manifest.json")
        else:
            require(stage in preflight_map and preflight_map[stage]["candidate_manifest_equal"],
                    f"quality row aliases a non-equivalent preflight: {stage}")
            stage_manifest_sha = preflight_map[stage]["manifest_sha256"]
        receipt_path = HERE / stage / name
        receipt = validate_receipt(receipt_path, stage, load_plan(),
                                   stage_manifest_sha, candidate_sha)
        receipt_value = read_json(receipt_path)
        require(row.get("receipt_sha256") == receipt["sha256"]
                and row.get("exit_code") == 0
                and receipt_value.get("exit_code") == 0,
                f"quality receipt result/binding differs: {relative(receipt_path)}")
        expected_command_value = [
            "env", "CARGO_TARGET_DIR=" + str(TARGET_PATH), "CARGO_BUILD_JOBS=2",
            "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings",
        ] + dict(commands)[name.removeprefix("check-").removesuffix(".receipt.json")]
        require(receipt_value.get("command") == expected_command_value,
                f"quality command differs: {relative(receipt_path)}")
        log_path = receipt_path.with_name(name.replace(".receipt.json", ".stdout"))
        log = read_text(log_path)
        executed = sum(int(number) for number in
                       re.findall(r"test result: ok\. (\d+) passed;", log))
        require(row.get("executed_tests") == executed
                and isinstance(executed, int),
                f"quality test count differs: {relative(receipt_path)}")
        tests += executed
    require({key.split("/", 1)[1] for key in seen} == expected_names,
            "quality checks do not match checks.py")
    require(value.get("executed_tests") == tests,
            "quality aggregate test count differs")
    return {"path": relative(path), "sha256": sha(path), "stage": final_source,
            "checks": len(checks), "executed_tests": tests}


def selected_manifest(stage: str) -> Path:
    if stage not in ("baseline", "candidate", "final"):
        raise EvidenceError(f"unexpected final source stage: {stage}")
    path = HERE / stage / "source-manifest.json"
    return need(path, f"{stage} final source manifest")


def validate_decision(source_info: dict[str, Any], analysis: dict[str, Any],
                      flags: dict[str, Any], quality: dict[str, Any]) -> dict[str, Any]:
    path = decision_path()
    value = read_json(path)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi_0527_pilot_decision_v1",
            f"{relative(path)} decision schema differs")
    disposition = value.get("disposition")
    require(disposition in ("accepted", "rejected", "rejected_and_reverted"),
            "decision disposition is unexpected")
    final_source = value.get("final_source")
    require(final_source in ("baseline", "candidate", "final"),
            "decision final source is unexpected")
    require(isinstance(value.get("production_change_retained"), bool),
            "decision retention field is not explicit")
    require(value.get("plan_sha256") == sha(PLAN_PATH)
            and value.get("comparison_sha256") == analysis["comparison_sha256"]
            and value.get("quality_summary_sha256") == quality["sha256"]
            and value.get("adverse_review_sha256") == flags["sha256"],
            "decision artifact binding differs")
    final_manifest_path = selected_manifest(final_source)
    final_manifest_sha = sha(final_manifest_path)
    require(value.get("final_source_manifest_sha256") == final_manifest_sha,
            "decision final source manifest binding differs")
    source_manifest = manifest(final_manifest_path)
    current = {name: sha(REPO / name) for name in current_source_names()}
    require(current == source_manifest,
            "current checkout does not match selected final source manifest")
    pilot = analysis["pilot"]["passed"]
    pilot_field = value.get("pilot_gates_passed")
    require(isinstance(pilot_field, bool) and pilot_field is pilot,
            "decision pilot gate binding differs")
    lanes = value.get("conditional_lanes")
    require(isinstance(lanes, dict) and set(lanes) == set(CONDITIONAL_LANES),
            "decision conditional lane statuses are incomplete")
    if not pilot:
        for lane in CONDITIONAL_LANES:
            item = lanes[lane]
            require(isinstance(item, dict)
                    and item.get("status") in ("unmeasured", "unmeasured_pilot_failed")
                    and isinstance(item.get("reason"), str)
                    and "pilot" in item["reason"].lower(),
                    f"{lane} is not explicitly unmeasured after pilot failure")
    if disposition == "accepted":
        require(pilot and value.get("production_change_retained") is True,
                "accepted decision lacks passing pilot and retention")
        require(final_source in ("candidate", "final"),
                "accepted decision selects a non-candidate source")
        require(all(isinstance(lanes[lane], dict)
                    and lanes[lane].get("status") in ("pass", "measured")
                    for lane in CONDITIONAL_LANES),
                "accepted decision lacks all conditional lanes")
    else:
        require(value.get("production_change_retained") is False,
                "rejected decision retains the production candidate")
        require(final_source in ("baseline", "final"),
                "rejected decision does not select baseline-compatible source")
    require(quality["stage"] == final_source
            and isinstance(value.get("review"), str) and value["review"].strip(),
            "decision final quality/review binding is incomplete")
    return {"path": relative(path), "sha256": sha(path),
            "disposition": disposition, "final_source": final_source,
            "pilot_gates_passed": pilot,
            "production_change_retained": value["production_change_retained"]}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    path = need(HERE / "cleanup.json", "cleanup receipt")
    value = read_json(path)
    require(value.get("plan_sha256") == sha(PLAN_PATH)
            and value.get("removed") == plan["owned_paths"]
            and value.get("accessible_process_references") == []
            and value.get("owned_paths_absent") is True
            and value.get("python_cache_absent") is True,
            "cleanup receipt custody differs")
    require(all(not Path(name).exists() for name in plan["owned_paths"]),
            "owned path remains after cleanup")
    require(not list(HERE.rglob("__pycache__")), "Python cache remains in evidence bundle")
    return {"path": relative(path), "sha256": sha(path),
            "owned_paths_absent": True, "python_cache_absent": True}


def evidence_inventory() -> dict[str, str]:
    return {relative(path): sha(path) for path in HERE.rglob("*")
            if path.is_file() and path.name != "SHA256SUMS"}


def validate_seal() -> dict[str, Any]:
    path = need(HERE / "SHA256SUMS", "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path).splitlines():
        try:
            digest, name = line.split("  ", 1)
        except ValueError as error:
            raise EvidenceError("SHA256SUMS line is malformed") from error
        safe_relative(name, "seal entry")
        require(name != "SHA256SUMS" and name not in expected and is_digest(digest),
                "seal entry is unsafe or duplicated")
        expected[name] = digest
    require(expected == evidence_inventory(), "SHA256SUMS inventory differs")
    require(not any(path.is_symlink() for path in HERE.rglob("*")),
            "sealed evidence contains a symlink")
    return {"path": relative(path), "entries": len(expected), "sha256": sha(path)}


def expect_reject(action: Callable[[], Any], label: str) -> str:
    try:
        action()
    except EvidenceError as error:
        return str(error)
    raise EvidenceError(f"negative probe was accepted: {label}")


def validate_negative_probes() -> dict[str, Any]:
    probes = {
        "unsafe_relative_path": lambda: safe_relative("../escape", "probe path"),
        "inverted_native_abba": lambda: validate_native_order(
            {"primary": {"case": "x", "repeats": 1, "shapes": ["x"],
                          "warmup": 0, "samples": 1},
             "guard_repeats": 0, "guards": []},
            [{"name": "native-r1-primary-x", "path": "candidate/x.receipt.json",
              "start_utc": "2026-01-01T00:00:02+00:00", "end_utc": "2026-01-01T00:00:03+00:00"},
             {"name": "native-r1-primary-x", "path": "baseline/x.receipt.json",
              "start_utc": "2026-01-01T00:00:00+00:00", "end_utc": "2026-01-01T00:00:01+00:00"}]),
        "review_unbound_row": lambda: match_review([{"pair": "A"}], [], "probe review"),
        "cleanup_binding": lambda: require("bad" == sha(PLAN_PATH), "cleanup plan binding differs"),
        "pilot_acceptance_without_gate": lambda: require(False,
                                                          "accepted decision lacks pilot gate"),
    }
    checks = [{"case": label, "rejected": True,
               "reason": expect_reject(action, label)}
              for label, action in probes.items()]
    return {"schema": "litchi-0527-verifier-negative-probes-v1",
            "status": "pass", "checks": checks,
            "scope": "Bounded in-memory custody/schema probes; no build, capture, or artifact mutation."}


def run_component(name: str) -> dict[str, Any]:
    plan = load_plan()
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
        return {"status": "pass", "analysis": analysis,
                "review": validate_flags(analysis)}
    if name == "quality":
        source = validate_source()
        analysis = validate_analysis()
        flags = validate_flags(analysis)
        decision = read_json(decision_path())
        final_source = decision.get("final_source")
        require(final_source in ("baseline", "candidate", "final"),
                "quality cannot select an unknown final source")
        captures = analysis["captures"]
        quality = validate_quality(final_source, source["candidate"]["manifest_sha256"],
                                   captures["preflights"])
        return {"status": "pass", "quality": quality}
    if name == "decision":
        source = validate_source()
        analysis = validate_analysis()
        flags = validate_flags(analysis)
        decision_value = read_json(decision_path())
        final_source = decision_value.get("final_source")
        require(final_source in ("baseline", "candidate", "final"),
                "decision final source is unknown")
        quality = validate_quality(final_source, source["candidate"]["manifest_sha256"],
                                   analysis["captures"]["preflights"])
        return {"status": "pass",
                "decision": validate_decision(source, analysis, flags, quality),
                "quality": quality}
    if name == "cleanup":
        return {"status": "pass", "cleanup": validate_cleanup(plan)}
    if name == "seal":
        return {"status": "pass", "seal": validate_seal()}
    if name == "probes":
        return validate_negative_probes()
    raise EvidenceError(f"unknown component: {name}")


def run_bundle(component: str, sealed: bool) -> dict[str, Any]:
    if component == "precleanup":
        order = ("source", "builds", "captures", "analysis", "flags", "quality", "decision")
    elif component == "all":
        order = ("source", "builds", "captures", "analysis", "flags", "quality",
                 "decision", "cleanup", "seal" if sealed else "seal")
    else:
        order = (component,)
    components: dict[str, Any] = {}
    for name in order:
        try:
            result = run_component(name)
            components[name] = {"status": "pass", "result": result}
        except IncompleteError as error:
            components[name] = {"status": "incomplete", "error": str(error)}
        except (EvidenceError, OSError, subprocess.CalledProcessError) as error:
            components[name] = {"status": "fail", "error": str(error)}
    statuses = [item["status"] for item in components.values()]
    if "fail" in statuses:
        status = "fail"
    elif "incomplete" in statuses:
        status = "incomplete"
    else:
        status = "pass"
    decision = components.get("decision", {}).get("result", {}).get("decision", {})
    disposition = decision.get("disposition")
    claim = "candidate-retained" if status == "pass" and disposition == "accepted" else "none"
    return {"schema": "litchi-0527-pilot-verification-v1", "status": status,
            "scope": component, "performance_claim": claim,
            "components": components,
            "cleanup_seal_pending": component == "precleanup" or
            ("cleanup" not in components or components.get("cleanup", {}).get("status") != "pass")}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=("source", "builds", "captures",
                                                 "analysis", "flags", "quality",
                                                 "decision", "cleanup", "seal",
                                                 "probes", "precleanup", "all"),
                        default="all")
    parser.add_argument("--sealed", action="store_true",
                        help="require the SHA256SUMS seal when running the bundle")
    parser.add_argument("--strict", action="store_true",
                        help="return exit code 2 for incomplete or failed components")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    report = run_bundle(args.component, args.sealed)
    text = json.dumps(report, indent=2) + "\n"
    if args.output:
        if args.sealed and args.output.resolve().is_relative_to(HERE):
            raise SystemExit("sealed verification output must be outside the evidence bundle")
        args.output.write_text(text, encoding="utf-8")
    print(text, end="")
    if args.strict and report["status"] != "pass":
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
