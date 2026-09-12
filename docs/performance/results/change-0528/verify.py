"""Independently verify the frozen 0528 XLSX attribution campaign.

This verifier is intentionally smaller than the native pilot verifiers.  The
0528 run has one baseline build and four Callgrind profile children; it does
not authorize a candidate, a latency claim, or a production change.  Checks
are staged so a live campaign is reported as ``incomplete`` and cannot be
mistaken for a passing attribution record.
"""

from __future__ import annotations

import argparse
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
FROZEN = HERE / "frozen-inputs.json"
SOURCE_BINDING = HERE / "source-binding.json"
SOURCE_REVIEW = HERE / "source-review.md"
PUBLICATION_REVIEW = HERE / "publication-source-review.md"
STORAGE = HERE / "storage.json"
SYMBOLS = HERE / "symbol-observation.json"
BASELINE = HERE / "baseline"
PUBLICATION_ANALYZER = HERE / "analyze_publication.py"
PUBLICATION_REPORT = HERE / "publication-analysis.json"
PROFILE_ANALYZER = HERE / "analyze.py"
PROFILE_REPORT = HERE / "profile-analysis.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
QUALITY_REUSE = HERE / "quality-reuse.json"
ANALYZER_TESTS = HERE / "analyzer-tests.json"
SCRATCH = Path("/tmp/litchi-goal-0528")
TARGET_ROOT = Path("/home/zhuhe/litchi-goal-0528-target")
RETAINED_BINARY_ROOT = TARGET_ROOT / "retained-binaries"
OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
SOURCE_ROOTS = ("crates/litchi-xlsx/",)
SOURCE_EXACT = frozenset({"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"})
PROFILE_JOBS = tuple(
    (repeat, shape)
    for repeat in range(1, 3)
    for shape in ("medium", "dense-sparse")
)
PROFILE_HELPER = HERE.parent / "change-0521" / "analyze_profiles.py"
RESULT_HELPER = HERE.parent / "change-0521" / "analyze.py"
RAW_HELPER = HERE.parent / "change-0519" / "analyze_profiles.py"
RAW_EDGES = HERE.parent / "change-0519" / "compare_profile_lanes.py"


class EvidenceError(ValueError):
    """A malformed, contradictory, or out-of-scope evidence artifact."""


class IncompleteError(EvidenceError):
    """An artifact needed by the selected component has not arrived yet."""


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


def digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def regular(path: Path) -> bool:
    return path.is_file() and not path.is_symlink()


def need(path: Path, label: str, *, reject_symlink: bool = True) -> Path:
    if not path.exists():
        raise IncompleteError(f"{label} is missing: {rel(path)}")
    if reject_symlink and path.is_symlink():
        raise EvidenceError(f"{label} is a symlink: {rel(path)}")
    return path


def read_json(path: Path, label: str | None = None) -> Any:
    label = label or rel(path)
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {label}: {error}") from error


def read_bytes(path: Path, label: str | None = None) -> bytes:
    need(path, label or rel(path))
    try:
        return path.read_bytes()
    except OSError as error:
        raise EvidenceError(f"cannot read {label or rel(path)}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    try:
        return read_bytes(path, label).decode("utf-8")
    except UnicodeDecodeError as error:
        raise EvidenceError(f"{label or rel(path)} is not UTF-8") from error


def sha(path: Path) -> str:
    if not path.exists():
        raise IncompleteError(f"artifact is missing for hashing: {rel(path)}")
    try:
        h = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                h.update(block)
        return h.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {rel(path)}: {error}") from error


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


def nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def source_name(name: str) -> bool:
    return (name.startswith("crates/") or name.startswith("tools/perf-baseline/")
            or name.startswith(".cargo/") or name in SOURCE_EXACT)


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
        safe_relative(name, f"{rel(path)} entry")
        require(source_name(name), f"{rel(path)} contains out-of-scope source {name}")
        require(digest(value), f"{rel(path)} has malformed digest for {name}")
        require(name not in result, f"{rel(path)} repeats {name}")
        result[name] = value
    return result


def run_checked(args: list[str], *, env: dict[str, str] | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, env=env, stderr=subprocess.PIPE)
    except subprocess.CalledProcessError as error:
        detail = error.stderr.decode(errors="replace")[-2000:]
        raise EvidenceError(f"command failed ({' '.join(args)}): {detail}") from error


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
        metadata, encoded_name = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3, "private source replay index entry is malformed")
        name = encoded_name.decode()
        require(name not in result, f"private source replay repeats {name}")
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
        raise EvidenceError("source replay objects cannot be read") from error
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


def replay_baseline_source() -> dict[str, Any]:
    manifest_path = need(BASELINE / "source-manifest.json", "baseline source manifest")
    patch_path = need(BASELINE / "source.patch", "baseline source patch")
    frozen_manifest = source_manifest(manifest_path)
    patch = read_bytes(patch_path)
    require_baseline_patch_empty(patch)
    plan = load_plan()
    with tempfile.TemporaryDirectory(prefix="litchi-0528-source-replay-") as directory:
        index = Path(directory) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        run_checked(["git", "read-tree", plan["revision"]], env=env)
        indexed = parse_private_index(index)
        scoped = {name: item for name, item in indexed.items() if source_name(name)}
        revision_names = {
            name for name in run_checked(["git", "ls-tree", "-r", "--name-only", plan["revision"]])
            .decode().splitlines() if source_name(name)
        }
        require(set(frozen_manifest) == revision_names,
                "baseline manifest differs from the frozen revision inventory")
        require(set(frozen_manifest) == current_source_names(),
                "baseline manifest differs from the current working source inventory")
        require(set(scoped) == set(frozen_manifest),
                "baseline manifest differs from the private replay index")
        blobs = index_blob_hashes(index, {oid for oid, _ in scoped.values()})
        for name, expected in frozen_manifest.items():
            oid, mode = scoped[name]
            require(mode in (100644, 100755), f"unexpected source mode for {name}")
            require(blobs.get(oid) == expected, f"source replay blob differs for {name}")
        changed = run_checked(
            ["git", "diff", "--cached", "--name-only", plan["revision"]], env=env
        ).decode().splitlines()
        require(not changed, "empty baseline patch produced a replay diff")
    return {
        "manifest_sha256": sha(manifest_path),
        "manifest_entries": len(frozen_manifest),
        "patch_sha256": sha(patch_path),
        "patch_empty": True,
        "replayed_changed_files": [],
        "xlsx_source_files": sum(name.startswith("crates/litchi-xlsx/src/")
                                  for name in frozen_manifest),
    }


def require_baseline_patch_empty(patch: bytes) -> None:
    """Apply the baseline custody rule to bytes, also used by probes."""

    require(not patch, "baseline source.patch must be empty for attribution-only 0528")


def load_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict) and set(frozen) == {"plan.json", "run.py"},
            "frozen-inputs inventory differs")
    require(frozen["plan.json"] == sha(PLAN), "frozen plan hash differs")
    require(frozen["run.py"] == sha(RUN), "frozen run.py hash differs")
    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("status") == "frozen-before-build-and-capture",
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
    require(plan.get("owned_paths") == [str(SCRATCH), str(TARGET_ROOT)],
            "owned temporary paths differ")
    require(plan.get("cpu") == 2, "planned CPU differs")
    require(plan.get("capture_lanes") == ["build-normal", "profile"],
            "capture lanes differ")
    require(plan.get("native_reference") ==
            "Sealed0527baseline; phase means only contextual, not a same-run timing/instruction equivalence.",
            "native reference scope differs")
    primary = plan.get("primary")
    require(isinstance(primary, dict)
            and primary.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save"
            and primary.get("shapes") == ["medium", "dense-sparse"]
            and primary.get("repeats") == 2 and primary.get("warmup") == 20
            and primary.get("samples") == 200, "primary plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict)
            and profile.get("shapes") == ["medium", "dense-sparse"]
            and profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 1 and profile.get("owner") == OWNER,
            "profile plan differs")
    scope = profile.get("scope", "")
    require(isinstance(scope, str) and all(term in scope for term in (
        "Method body only", "returned MultiSnapshot", "final call")),
            "profile scope does not bound the measured method")
    admission = plan.get("admission")
    require(isinstance(admission, str) and "Attribution only" in admission
            and "No production retention" in admission
            and "speedup gate" in admission.lower(),
            "attribution-only admission is missing")
    return plan


def validate_prior_seal(directory: str, expected_sha: str) -> dict[str, Any]:
    root = REPO / directory
    seal = need(root / "SHA256SUMS", f"prior seal {directory}")
    require(sha(seal) == expected_sha, f"prior seal digest differs: {directory}")
    entries: dict[str, str] = {}
    for line in read_text(seal, f"{directory}/SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and digest(fields[0]), f"malformed prior seal line: {directory}")
        value, name = fields
        safe_relative(name, f"{directory} seal entry")
        require(name != "SHA256SUMS" and name not in entries,
                f"unsafe or duplicate prior seal entry: {directory}/{name}")
        path = root / name
        require(regular(path), f"prior seal entry is not a regular file: {directory}/{name}")
        require(sha(path) == value, f"prior sealed file differs: {directory}/{name}")
        entries[name] = value
    actual = {
        p.relative_to(root).as_posix(): sha(p)
        for p in root.rglob("*") if p.is_file() and p.name != "SHA256SUMS"
    }
    require(entries == actual, f"prior seal inventory differs: {directory}")
    require(not any(p.is_symlink() for p in root.rglob("*")),
            f"prior seal contains a symlink: {directory}")
    return {"directory": directory, "seal_sha256": sha(seal), "entries": len(entries)}


def validate_bindings() -> dict[str, Any]:
    plan = load_plan()
    binding = read_json(SOURCE_BINDING, "source-binding.json")
    require(isinstance(binding, dict), "source-binding is not an object")
    require(binding.get("revision") == plan["revision"], "source-binding revision differs")
    require(binding.get("previous_turn_classification") == "progress",
            "source-binding classification differs")
    source_files = binding.get("source_files")
    require(isinstance(source_files, dict) and source_files, "source_files is missing")
    for name, expected in source_files.items():
        safe_relative(name, "source binding path")
        require(digest(expected) and regular(REPO / name), f"source binding differs: {name}")
        require(sha(REPO / name) == expected, f"current source differs: {name}")
    required_sources = {
        "Cargo.toml",
        "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs",
        "crates/litchi-xlsx/src/raw/worksheet/validation.rs",
        "crates/litchi-xlsx/src/cell_values/validation.rs",
        "crates/litchi-xlsx/src/cell_values/source.rs",
        "tools/perf-baseline/src/lib.rs",
        "tools/perf-baseline/Cargo.lock",
        "docs/GOAL.md",
        "docs/CRUD_Scenario_Checklist.md",
        "rust-toolchain.toml",
    }
    require(set(source_files) == required_sources, "source binding inventory differs")
    adr_files = binding.get("adr_files")
    require(isinstance(adr_files, dict) and len(adr_files) == 30,
            "ADR binding inventory differs")
    for name, expected in adr_files.items():
        safe_relative(name, "ADR binding path")
        require(name.startswith("docs/adr/") and digest(expected)
                and regular(REPO / name) and sha(REPO / name) == expected,
                f"ADR binding differs: {name}")
    dependencies = binding.get("dependency_inputs")
    require(isinstance(dependencies, dict) and len(dependencies) == 2,
            "dependency input inventory differs")
    dependency_rows = []
    for name, item in dependencies.items():
        require(Path(name).is_absolute() and isinstance(item, dict),
                "dependency binding path is malformed")
        source = Path(name)
        retained_name = item.get("retained")
        require(digest(item.get("sha256")) and regular(source),
                f"dependency source is missing or malformed: {name}")
        safe_relative(retained_name, f"retained dependency for {name}")
        retained = HERE / retained_name
        require(regular(retained), f"retained dependency is missing: {retained_name}")
        require(sha(source) == item["sha256"] == sha(retained),
                f"dependency hash differs: {name}")
        require(source.read_bytes() == retained.read_bytes(),
                f"retained dependency bytes differ: {name}")
        dependency_rows.append({"source": name, "retained": rel(retained),
                                "sha256": item["sha256"]})
    prior = binding.get("prior_seals")
    require(prior == {
        "docs/performance/results/change-0527": "16f51d610116ecdd7872d606802b9253fb7e2dcd15094f6722172db626202923",
        "docs/performance/results/change-0525": "ae8834847914ef936b6206f621d0e19c02960f407b27f3e2db797177e0a8b1ab",
    }, "prior seal binding differs")
    seals = [validate_prior_seal(name, value) for name, value in sorted(prior.items())]
    require(binding.get("priority") == "OLE2/OOXML first; ODF deferred; iWork excluded",
            "format priority differs")
    source_review = read_text(SOURCE_REVIEW, "source-review.md")
    publication_review = read_text(PUBLICATION_REVIEW, "publication-source-review.md")
    for review, label in ((source_review, "source-review"),
                          (publication_review, "publication-source-review")):
        require(plan["revision"] in review and "0528" in review,
                f"{label} is not bound to 0528 revision")
        require("production_change: none" in review.lower()
                or "no rust source was edited" in review.lower(),
                f"{label} does not state read-only scope")
        require("quick-xml" in review and "0.41.0" in review,
                f"{label} omits quick-xml dependency scope")
    require(sha(SOURCE_BINDING) in source_review,
            "source review does not bind source-binding.json")
    source_result = replay_baseline_source()
    return {"plan_sha256": sha(PLAN), "run_sha256": sha(RUN),
            "source_binding_sha256": sha(SOURCE_BINDING),
            "source_review_sha256": sha(SOURCE_REVIEW),
            "publication_review_sha256": sha(PUBLICATION_REVIEW),
            "dependency_inputs": dependency_rows, "prior_seals": seals,
            "baseline": source_result}


def validate_storage(*, post_cleanup: bool = False) -> dict[str, Any]:
    value = read_json(STORAGE, "storage.json")
    require(value == {
        "scratch": str(SCRATCH),
        "target": str(RETAINED_BINARY_ROOT),
        "reason": "Disk-backed retained binaries selected before builds because tmpfs user quota was exhausted in0527.",
    }, "storage selection differs")
    scratch_exists = os.path.lexists(SCRATCH)
    target_exists = TARGET_ROOT.exists()
    if scratch_exists:
        require(SCRATCH.is_symlink(), "scratch path must be the documented symlink")
        require(SCRATCH.resolve() == RETAINED_BINARY_ROOT,
                "scratch symlink resolves to an unexpected retained-binary directory")
        require(RETAINED_BINARY_ROOT.is_dir() and not RETAINED_BINARY_ROOT.is_symlink(),
                "retained-binary directory is not a regular directory")
    elif not post_cleanup:
        raise IncompleteError("scratch storage disappeared before cleanup receipt")
    if post_cleanup:
        require(not scratch_exists and not target_exists,
                "owned storage remains after cleanup")
    return {"scratch": str(SCRATCH), "target": str(RETAINED_BINARY_ROOT),
            "scratch_symlink": scratch_exists, "post_cleanup": post_cleanup}


def check_interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    finite_number(value.get("seconds"), f"{label}.seconds")
    require(end > start and value["seconds"] > 0, f"{label} interval is invalid")
    return start, end


def validate_artifacts(receipt: dict[str, Any], stage: Path, name: str) -> set[str]:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and artifacts, f"{name} artifacts are missing")
    for filename, expected in artifacts.items():
        safe_relative(filename, f"{name} artifact")
        require(Path(filename).name == filename and digest(expected),
                f"{name} artifact entry is malformed: {filename}")
        path = stage / filename
        require(regular(path) and sha(path) == expected,
                f"{name} artifact custody differs: {filename}")
    actual = {
        p.name for p in stage.glob(name + ".*")
        if p.is_file() and p.name != name + ".receipt.json"
        and not p.name.endswith((".inclusive.txt", ".self.txt"))
    }
    require(actual == set(artifacts), f"{name} receipt artifact inventory differs")
    return set(artifacts)


def validate_builds() -> dict[str, Any]:
    plan = load_plan()
    source = validate_bindings()
    if CLEANUP.exists():
        validate_cleanup()
    storage = validate_storage(post_cleanup=CLEANUP.exists())
    receipt_path = need(BASELINE / "build-normal.receipt.json", "baseline build receipt")
    receipt = read_json(receipt_path)
    start, end = check_interval(receipt, "build-normal")
    expected_command = [
        "env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
        "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--bin", "litchi-perf-baseline", "--target-dir", str(TARGET_ROOT),
    ]
    require(receipt.get("command") == expected_command, "build command differs")
    require(receipt.get("exit_code") == 0 and receipt.get("binary_sha256") is None,
            "baseline build receipt is not a successful source build")
    require(receipt.get("plan_sha256") == sha(PLAN)
            and receipt.get("script_sha256") == sha(RUN),
            "build receipt frozen bindings differ")
    require(receipt.get("source_manifest_sha256") == source["baseline"]["manifest_sha256"]
            and receipt.get("working_source_manifest_sha256") == source["baseline"]["manifest_sha256"],
            "build receipt source binding differs")
    environment = receipt.get("environment")
    require(isinstance(environment, dict)
            and environment.get("TMPDIR") == str(TARGET_ROOT / "test-tmp"),
            "build TMPDIR is not the planned target")
    require(validate_artifacts(receipt, BASELINE, "build-normal") ==
            {"build-normal.stdout", "build-normal.stderr"},
            "build output inventory differs")
    identity_path = need(BASELINE / "binary-normal.json", "baseline normal binary identity")
    identity = read_json(identity_path)
    binary_path_value = identity.get("path")
    binary_sha = identity.get("sha256")
    require(isinstance(binary_path_value, str) and Path(binary_path_value).is_absolute()
            and digest(binary_sha), "binary identity is malformed")
    binary_path = Path(binary_path_value)
    require(binary_path == SCRATCH / "baseline-normal",
            "binary identity path is outside retained storage")
    require(identity.get("bytes", 0) > 0 and identity.get("build_receipt_sha256") == sha(receipt_path)
            and identity.get("source_manifest_sha256") == source["baseline"]["manifest_sha256"],
            "binary identity bindings differ")
    if binary_path.exists():
        require(binary_path.resolve().parent == RETAINED_BINARY_ROOT, "live binary escaped owned storage")
        require(regular(binary_path) and sha(binary_path) == binary_sha
                and binary_path.stat().st_size == identity["bytes"],
                "retained baseline binary differs from identity")
    else:
        require(CLEANUP.exists(), "baseline binary disappeared without cleanup receipt")
    symbols = read_json(SYMBOLS, "symbol-observation.json")
    require(symbols.get("binary_sha256") == binary_sha
            and symbols.get("owner") == OWNER
            and symbols.get("command") == ["nm", "-C", "/tmp/litchi-goal-0528/baseline-normal"],
            "symbol observation binding differs")
    rows = symbols.get("symbols")
    require(isinstance(rows, list) and len(rows) == 5,
            "symbol observation does not retain five monomorphizations")
    require(all(isinstance(row, str) and OWNER in row for row in rows),
            "symbol observation owner rows differ")
    return {"source": source, "storage": storage, "interval": {
        "start_utc": start.isoformat(), "end_utc": end.isoformat()},
        "receipt": rel(receipt_path), "receipt_sha256": sha(receipt_path),
        "binary": {"path": binary_path_value, "sha256": binary_sha,
                    "bytes": identity["bytes"], "identity_sha256": sha(identity_path)},
        "symbols_sha256": sha(SYMBOLS)}


def profile_command(name: str, shape: str) -> list[str]:
    return [
        "taskset", "-c", "2", "valgrind", "--tool=callgrind", "--collect-atstart=no",
        "--toggle-collect=" + OWNER, "--zero-before=" + OWNER,
        "--dump-after=" + OWNER,
        "--callgrind-out-file=" + str(BASELINE / (name + ".callgrind")),
        str(SCRATCH / "baseline-normal"), "--warmup", "0", "--samples", "1",
        "--case", "xlsx_source_backed_cell_values_one_percent_edit_save",
        "--xlsx-cell-crud-shape", shape, "--json", str(BASELINE / (name + ".json")),
    ]


def load_0521_analyzer() -> Any:
    need(RESULT_HELPER, "retained 0521 result helper")
    spec = importlib.util.spec_from_file_location("litchi_0521_result_helper", RESULT_HELPER)
    require(spec is not None and spec.loader is not None, "0521 helper cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def raw_oracle(raw: dict[str, Any], name: str) -> dict[str, Any]:
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{name} result count differs")
    result = results[0]
    require(isinstance(result, dict), f"{name} result is not an object")
    source = result.get("source")
    xlsx = source.get("xlsx_cell_values") if isinstance(source, dict) else None
    require(isinstance(xlsx, dict), f"{name} has no XLSX source oracle")
    output = result.get("output_sha256")
    semantic = xlsx.get("semantic_sha256")
    require(digest(output), f"{name} output oracle is malformed")
    require(isinstance(semantic, list) and len(semantic) == 1 and digest(semantic[0]),
            f"{name} semantic oracle is malformed")
    require(isinstance(xlsx.get("output_sha256"), list)
            and xlsx["output_sha256"] == [output],
            f"{name} source/output oracle differs")
    require(result.get("corpus", {}).get("package_format") == "XLSX/OPC/ZIP",
            f"{name} corpus is not XLSX/OPC/ZIP")
    require(xlsx.get("timing_scope") ==
            "open, selector planning, commit, and stream publication; reopen/verification is separate and excluded",
            f"{name} timing scope includes an unplanned phase")
    return {"name": name, "output_sha256": output, "semantic_sha256": semantic[0],
            "archive_sha256": result["corpus"].get("archive_sha256"),
            "source_members": xlsx.get("source_members"),
            "corpus": result["corpus"]}


def require_profile_matrix(rows: list[dict[str, Any]], label: str) -> None:
    expected = set(PROFILE_JOBS)
    actual = {(row.get("repeat"), row.get("shape")) for row in rows}
    require(len(rows) == len(expected) and actual == expected,
            f"{label} profile matrix contains an omission or duplicate")


def require_serial_intervals(intervals: list[tuple[str, dt.datetime, dt.datetime]]) -> None:
    ordered = sorted(intervals, key=lambda row: row[1])
    require([row[0] for row in ordered]
            == [f"profile-r{repeat}-{shape}" for repeat, shape in PROFILE_JOBS],
            "profile receipts are not in frozen serial order")
    for previous, current in zip(ordered, ordered[1:]):
        require(previous[2] <= current[1],
                f"profile intervals overlap: {previous[0]} and {current[0]}")


def validate_profiles() -> dict[str, Any]:
    plan = load_plan()
    builds = validate_builds()
    binary_sha = builds["binary"]["sha256"]
    manifest_sha = builds["source"]["baseline"]["manifest_sha256"]
    helper = load_0521_analyzer()
    rows = []
    intervals = []
    for repeat, shape in PROFILE_JOBS:
        name = f"profile-r{repeat}-{shape}"
        receipt_path = need(BASELINE / (name + ".receipt.json"), f"{name} receipt")
        receipt = read_json(receipt_path)
        start, end = check_interval(receipt, name)
        intervals.append((name, start, end))
        require(receipt.get("command") == profile_command(name, shape),
                f"{name} command differs")
        require(receipt.get("exit_code") == 0
                and receipt.get("binary_sha256") == binary_sha
                and receipt.get("plan_sha256") == sha(PLAN)
                and receipt.get("script_sha256") == sha(RUN),
                f"{name} receipt binding differs")
        require(receipt.get("source_manifest_sha256") == manifest_sha
                and receipt.get("working_source_manifest_sha256") == manifest_sha,
                f"{name} source binding differs")
        environment = receipt.get("environment")
        require(isinstance(environment, dict)
                and environment.get("TMPDIR") == str(TARGET_ROOT / "test-tmp"),
                f"{name} TMPDIR differs")
        artifacts = validate_artifacts(receipt, BASELINE, name)
        raw_path = need(BASELINE / (name + ".json"), f"{name} profile output")
        raw = read_json(raw_path, f"{name}.json")
        require(sha(raw_path) == receipt["artifacts"].get(name + ".json"),
                f"{name} profile output is not receipt-bound")
        job = {"name": name, "kind": "profile", "guard": None,
               "repeat": repeat, "case": plan["primary"]["case"], "shape": shape,
               "warmup": 0, "samples": 1}
        try:
            normalized = helper.validate_result(raw, plan, job, {"sha256": binary_sha}, False)
        except Exception as error:
            raise EvidenceError(f"{name} raw profile oracle failed: {error}") from error
        oracle = raw_oracle(raw, name)
        numbered = sorted(
            BASELINE.glob(name + ".callgrind.[0-9]*"),
            key=lambda path: int(path.name.rsplit(".", 1)[1]),
        )
        require(numbered and [int(path.name.rsplit(".", 1)[1]) for path in numbered]
                == list(range(1, len(numbered) + 1)),
                f"{name} numbered Callgrind dumps are not contiguous")
        require(all((name + ".callgrind." + str(index)) in artifacts
                     for index in range(1, len(numbered) + 1)),
                f"{name} numbered dump is not receipt-bound")
        rows.append({"name": name, "repeat": repeat, "shape": shape,
                     "receipt_sha256": sha(receipt_path),
                     "profile_sha256": sha(raw_path), "callgrind_parts": len(numbered),
                     "oracle": oracle, "normalized_identity_sha256": normalized["identity_sha256"],
                     "start_utc": start.isoformat(), "end_utc": end.isoformat()})
    require_serial_intervals(intervals)
    return {"status": "pass", "binary_sha256": binary_sha,
            "source_manifest_sha256": manifest_sha, "rows": rows,
            "serial": True, "sample_count": len(rows)}


def load_profile_helpers() -> Any:
    need(PROFILE_HELPER, "0521 profile helper")
    spec = importlib.util.spec_from_file_location("litchi_0521_profile_helpers_0528", PROFILE_HELPER)
    require(spec is not None and spec.loader is not None, "profile helper cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.HERE = HERE
    # The old helper's annotation name splitter treats an earlier ` [T]` as
    # an object suffix.  The frozen publication analyzer supplies this exact
    # adapter; use the same adapter for the independent annotation replay.
    def display_name(text: str) -> str:
        text = text.rsplit(" [", 1)[0].strip()
        text = re.sub(r"\s+\([\d,]+x\)$", "", text)
        if text.startswith("???:"):
            return text[4:]
        return text.rsplit(":", 1)[-1] if text.startswith(("./", "/")) else text
    module.display_name = display_name
    return module


def edge_summary(helper: Any, path: Path, target: str, caller: str | None = None) -> dict[str, Any]:
    return helper.target_edge_summary(path, target, caller)


def publication_dump_parts(helper: Any, name: str) -> dict[str, Any]:
    stem = BASELINE / name
    raw = Path(str(stem) + ".callgrind")
    numbered = sorted(
        raw.parent.glob(raw.name + ".[0-9]*"),
        key=lambda path: int(path.suffix[1:]),
    )
    require(numbered and [int(path.suffix[1:]) for path in numbered]
            == list(range(1, len(numbered) + 1)), f"{name} dump sequence is not contiguous")
    parts = []
    selected: list[Path] = []
    for index, path in enumerate(numbered, 1):
        text = read_text(path, rel(path))
        require("events: Ir" in text.splitlines(), f"{rel(path)} event set differs")
        require(helper.part_number(text, rel(path)) == index,
                f"{rel(path)} part number differs")
        require(helper.trigger(text, rel(path)) == "--dump-after=" + OWNER,
                f"{rel(path)} trigger differs")
        total = helper.summary_ir(text, rel(path))
        measured = edge_summary(helper, path, OWNER, MEASURED_PARENT)
        lifecycle = edge_summary(helper, path, OWNER, LIFECYCLE_PARENT)
        if measured["inclusive_ir"] > 0:
            require(measured["positive_edge_count"] == 1 and measured["calls"] == 1
                    and measured["inclusive_ir"] == total,
                    f"{rel(path)} measured owner edge differs")
            selected.append(path)
        else:
            require(lifecycle["positive_edge_count"] == 1
                    and lifecycle["inclusive_ir"] == total,
                    f"{rel(path)} lifecycle owner edge differs")
        parts.append({"part": index, "path": rel(path), "sha256": sha(path),
                      "summary_ir": total, "measured_parent": measured,
                      "lifecycle_parent": lifecycle})
    require(len(selected) == 1 and selected[0] == numbered[-1],
            f"{name} measured owner call is not the unique final dump")
    terminal_text = read_text(raw, rel(raw))
    require(helper.trigger(terminal_text, rel(raw)) == "Program termination"
            and helper.summary_ir(terminal_text, rel(raw)) == 0,
            f"{name} terminal dump is not zero-ir Program termination")
    require(helper.part_number(terminal_text, rel(raw)) == len(numbered) + 1,
            f"{name} terminal part number is not after retained numbered dumps")
    return {"raw": raw, "numbered": numbered, "selected": selected[0],
            "parts": parts, "termination": {"path": rel(raw), "sha256": sha(raw),
                                              "summary_ir": 0}}


def validate_publication_analysis(profile_rows: dict[str, Any]) -> dict[str, Any]:
    helper = load_profile_helpers()
    report_path = need(PUBLICATION_REPORT, "publication-analysis.json")
    report = read_json(report_path)
    require(report.get("status") == "pass" and report.get("plan_sha256") == sha(PLAN)
            and report.get("owner") == OWNER
            and report.get("helper_sha256") == sha(PROFILE_HELPER),
            "publication analysis envelope differs")
    scope = report.get("scope", "")
    require(isinstance(scope, str) and "method Ir only" in scope
            and "Returned snapshot destruction" in scope
            and "No latency improvement" in scope,
            "publication analysis scope overclaims")
    rows = report.get("rows")
    require(isinstance(rows, list), "publication analysis rows are missing")
    require_profile_matrix(rows, "publication analysis")
    checked = []
    for row in rows:
        name = f"profile-r{row['repeat']}-{row['shape']}"
        computed = publication_dump_parts(helper, name)
        require(row.get("selected") == rel(computed["selected"]),
                f"{name} selected dump differs")
        require(row.get("termination") == computed["termination"],
                f"{name} terminal dump binding differs")
        parts = row.get("parts")
        require(isinstance(parts, list) and len(parts) == len(computed["parts"]),
                f"{name} retained dump list differs")
        for actual, expected in zip(parts, computed["parts"]):
            # The frozen publication analyzer records the part number inside
            # the dump and path, but omits a redundant JSON ``part`` field.
            expected_report_part = {key: value for key, value in expected.items()
                                    if key != "part"}
            require(actual == expected_report_part,
                    f"{name} raw dump attribution differs at part {expected['part']}")
        annotations = row.get("annotations")
        require(isinstance(annotations, dict), f"{name} annotation map is missing")
        annotation_paths = {
            rel(Path(str(BASELINE / (name + suffix))))
            for suffix in (".inclusive.txt", ".self.txt")
        }
        require(set(annotations) == annotation_paths, f"{name} annotation inventory differs")
        annotation_text: dict[str, str] = {}
        for annotation_name, expected_sha in annotations.items():
            path = HERE / safe_relative(annotation_name, f"{name} annotation")
            require(digest(expected_sha) and regular(path) and sha(path) == expected_sha,
                    f"{name} annotation custody differs")
            inclusive = annotation_name.endswith(".inclusive.txt")
            output, _ = helper.run_annotation(computed["selected"], inclusive)
            require(path.read_text() == output, f"{name} annotation replay differs")
            annotation_text[annotation_name] = output
        inclusive_name = rel(BASELINE / (name + ".inclusive.txt"))
        self_name = rel(BASELINE / (name + ".self.txt"))
        inc = helper.parse_annotation(annotation_text[inclusive_name], OWNER, inclusive_name)
        own = helper.parse_annotation(annotation_text[self_name], OWNER, self_name)
        direct = helper.direct_map(inc["direct"])
        require(direct == helper.direct_map(own["direct"])
                and inc["selected_ir"] == own["selected_ir"] + sum(direct.values())
                and inc["selected_ir"] == computed["parts"][-1]["summary_ir"],
                f"{name} owner self/direct equation differs")
        owners = row.get("owners")
        require(isinstance(owners, dict) and OWNER in owners,
                f"{name} owner decomposition is missing")
        for owner_name, owner_data in owners.items():
            require(isinstance(owner_name, str) and isinstance(owner_data, dict),
                    f"{name} owner decomposition is malformed")
            for field in ("inclusive_ir", "self_ir"):
                nonnegative_integer(owner_data.get(field), f"{name}.{owner_name}.{field}")
            owner_direct = owner_data.get("direct")
            require(isinstance(owner_direct, dict), f"{name}.{owner_name}.direct is missing")
            total = 0
            for child, cost in owner_direct.items():
                require(isinstance(child, str), f"{name} child name is malformed")
                nonnegative_integer(cost, f"{name}.{owner_name}->{child}")
                total += cost
                edge = edge_summary(helper, computed["selected"], child, owner_name)
                require(edge["positive_edge_count"] > 0 and edge["inclusive_ir"] == cost,
                        f"{name} raw edge differs: {owner_name} -> {child}")
            require(owner_data["inclusive_ir"] == owner_data["self_ir"] + total,
                    f"{name} nested owner equation differs: {owner_name}")
        checked.append({"name": name, "parts": len(computed["parts"]),
                        "selected": rel(computed["selected"]),
                        "owner_ir": inc["selected_ir"],
                        "owners": len(owners)})
    return {"report": rel(report_path), "report_sha256": sha(report_path),
            "rows": checked}


def replay_report(script: Path, report_path: Path, args: list[str]) -> None:
    need(script, f"analyzer {rel(script)}")
    need(report_path, f"analyzer report {rel(report_path)}")
    with tempfile.TemporaryDirectory(prefix="litchi-0528-analysis-") as directory:
        output = Path(directory) / "report.json"
        command = [sys.executable, "-B", str(script), *args, "--output", str(output)]
        try:
            result = subprocess.run(command, cwd=REPO, stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, check=False)
        except OSError as error:
            raise EvidenceError(f"cannot replay {rel(script)}: {error}") from error
        require(result.returncode == 0,
                f"{rel(script)} replay failed: {result.stderr.decode(errors='replace')[-2000:]}")
        require(output.is_file() and output.read_bytes() == report_path.read_bytes(),
                f"{rel(script)} replay differs from retained report")


def validate_raw_profile_analysis() -> dict[str, Any]:
    report = read_json(PROFILE_REPORT, "profile-analysis.json")
    replay_report(PROFILE_ANALYZER, PROFILE_REPORT, [])
    require(report.get("schema") == "litchi-0528-scanner-ceiling-audit-v1"
            and report.get("status") == "pass"
            and report.get("selected_owner") ==
            "litchi_xlsx::cell_values::source::MultiSourceEdit::commit",
            "raw profile analysis envelope differs")
    scope = report.get("scope", "")
    require(isinstance(scope, str) and "read-only" in scope.lower()
            and "attribution" in scope.lower() and "no build" in scope.lower()
            and "production edit" in scope.lower(),
            "raw profile analysis scope overclaims")
    inputs = report.get("inputs")
    require(isinstance(inputs, dict), "raw profile analysis inputs are missing")
    for key, expected_path, expected_sha in (
        ("plan", "docs/performance/results/change-0528/plan.json", sha(PLAN)),
        ("run", "docs/performance/results/change-0528/run.py", sha(RUN)),
        ("frozen_inputs", "docs/performance/results/change-0528/frozen-inputs.json", sha(FROZEN)),
    ):
        row = inputs.get(key)
        require(isinstance(row, dict) and row.get("path") == expected_path
                and row.get("sha256") == expected_sha,
                f"raw profile analysis {key} binding differs")
    helpers = inputs.get("helpers")
    require(isinstance(helpers, dict)
            and helpers.get(repo_rel(PROFILE_HELPER)) == sha(PROFILE_HELPER)
            and helpers.get(repo_rel(RAW_HELPER)) == sha(RAW_HELPER)
            and helpers.get(repo_rel(RAW_EDGES)) == sha(RAW_EDGES),
            "raw profile analysis helper bindings differ")
    require(report.get("prior_0525_seal_entries") == 550,
            "raw profile analysis prior seal count differs")
    rows = report.get("rows")
    require(isinstance(rows, list) and {(r.get("repeat"), r.get("shape")) for r in rows}
            == set(PROFILE_JOBS), "raw profile analysis row matrix differs")
    aggregate = report.get("aggregate")
    require(isinstance(aggregate, dict)
            and aggregate.get("scanner_to_resolve_event", {}).get("upper_bound_only") is True
            and aggregate.get("scanner_to_resolve_event", {}).get("call_counts_used_for_estimate") is False,
            "resolver ceiling is not bounded as an upper bound")
    for row in rows:
        require(row.get("validation", {}).get("raw_direct_edges_match_annotations") is True
                and row.get("validation", {}).get("self_plus_direct_equations") is True
                and row.get("resolver_edge", {}).get("raw_edge_matches_annotation") is True,
                f"raw profile analysis row validation differs: {row.get('repeat')}/{row.get('shape')}")
    limitations = report.get("limitations")
    require(isinstance(limitations, list)
            and any("upper bound" in item.lower() and "start" in item.lower()
                    for item in limitations if isinstance(item, str))
            and any("native latency" in item.lower() for item in limitations if isinstance(item, str)),
            "raw profile analysis limitations are incomplete")
    return {"report": rel(PROFILE_REPORT), "report_sha256": sha(PROFILE_REPORT),
            "rows": len(rows), "prior_seal_entries": report["prior_0525_seal_entries"],
            "resolver_upper_bound_only": True}


def decision_path() -> Path:
    paths = [HERE / "decision.json", HERE / "disposition.json"]
    existing = [path for path in paths if path.exists()]
    require(len(existing) == 1, "exactly one 0528 decision record is required")
    return need(existing[0], "0528 decision record")


def validate_decision() -> dict[str, Any]:
    path = decision_path()
    value = read_json(path, "decision")
    require(value.get("schema") == "litchi_0528_attribution_decision_v1",
            "decision schema differs")
    require(value.get("disposition") in ("attribution_complete", "recorded", "rejected"),
            "decision disposition is not attribution-only")
    require(value.get("production_change_retained") is False,
            "0528 decision retains an unauthorized production change")
    require(value.get("performance_claim") == "none" and value.get("speedup_claim") is False,
            "0528 decision contains a performance claim")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("source_binding_sha256") == sha(SOURCE_BINDING),
            "decision frozen source bindings differ")
    require(value.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json"),
            "decision source manifest differs")
    require(value.get("binary_sha256") == read_json(BASELINE / "binary-normal.json")["sha256"],
            "decision binary binding differs")
    require(value.get("publication_analysis_sha256") == sha(PUBLICATION_REPORT)
            and value.get("profile_analysis_sha256") == sha(PROFILE_REPORT),
            "decision analysis bindings differ")
    require(value.get("profile_oracle_sha256") == sha(BASELINE / "profile-r1-medium.json")
            or value.get("profile_oracle_sha256") == sha(BASELINE / "profile-r1-dense-sparse.json"),
            "decision profile oracle binding differs")
    review = value.get("review")
    require(isinstance(review, str) and review.strip()
            and "attribution" in review.lower()
            and "no" in review.lower(), "decision review is incomplete")
    return {"path": rel(path), "sha256": sha(path),
            "disposition": value["disposition"], "production_change_retained": False,
            "cleanup_and_seal_checked_separately": True}


def validate_cleanup() -> dict[str, Any]:
    path = need(CLEANUP, "cleanup receipt")
    value = read_json(path, "cleanup receipt")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("removed") == [str(SCRATCH), str(TARGET_ROOT)]
            and value.get("accessible_process_references") == []
            and value.get("owned_paths_absent") is True
            and value.get("python_cache_absent") is True,
            "cleanup receipt custody differs")
    storage = validate_storage(post_cleanup=True)
    require(not list(HERE.rglob("__pycache__")), "Python cache remains in evidence bundle")
    return {"path": rel(path), "sha256": sha(path), "storage": storage,
            "owned_paths_absent": True, "python_cache_absent": True}


def evidence_inventory() -> dict[str, str]:
    result: dict[str, str] = {}
    for path in HERE.rglob("*"):
        if path.is_symlink():
            raise EvidenceError(f"evidence bundle contains a symlink: {rel(path)}")
        if path.is_file() and path.name != "SHA256SUMS":
            result[rel(path)] = sha(path)
    return result


def validate_seal() -> dict[str, Any]:
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and digest(fields[0]), "malformed SHA256SUMS line")
        value, name = fields
        safe_relative(name, "seal entry")
        require(name != "SHA256SUMS" and name not in expected,
                "unsafe or duplicate SHA256SUMS entry")
        expected[name] = value
    require(expected == evidence_inventory(), "SHA256SUMS inventory differs")
    return {"path": rel(path), "entries": len(expected), "sha256": sha(path)}


def expect_reject(action: Callable[[], Any], label: str) -> str:
    try:
        action()
    except EvidenceError as error:
        return str(error)
    raise EvidenceError(f"negative probe was accepted: {label}")


def validate_quality_reuse() -> dict[str, Any]:
    value = read_json(HERE / "quality-reuse.json")
    prior = HERE.parent / "change-0527"
    require((BASELINE / "source-manifest.json").read_bytes() ==
            (prior / "final/source-manifest.json").read_bytes(), "quality source differs")
    require(value["source_manifest_sha256"] == sha(BASELINE / "source-manifest.json")
            and value["prior_quality_summary_sha256"] == sha(prior / "quality-summary.json")
            and value["prior_seal_sha256"] == sha(prior / "SHA256SUMS"), "quality bindings differ")
    quality = read_json(prior / "quality-summary.json")
    require(quality["status"] == "pass" and quality["stage"] == "final"
            and len(quality["checks"]) == 12, "prior quality is incomplete")
    tests = 0
    names = set()
    for row in quality["checks"]:
        name = row["name"]
        require(Path(name).name == name and name.startswith("check-")
                and name.endswith(".receipt.json") and name not in names,
                "unsafe or duplicate quality receipt")
        names.add(name)
        receipt_path = prior / "final" / name
        receipt = read_json(receipt_path)
        require(row["stage"] == "final" and row["receipt_sha256"] == sha(receipt_path)
                and receipt["exit_code"] == row["exit_code"] == 0
                and receipt["source_manifest_sha256"] == value["source_manifest_sha256"],
                "prior quality receipt differs")
        for artifact, digest in receipt["artifacts"].items():
            require(Path(artifact).name == artifact and sha(prior / "final" / artifact) == digest,
                    "quality log differs")
        log = (prior / "final" / name.replace(".receipt.json", ".stdout")).read_text()
        count = sum(int(n) for n in re.findall(r"test result: ok\. (\d+) passed;", log))
        require(count == row["executed_tests"], "test count differs")
        tests += count
    require(tests == quality["executed_tests"] == value["successful_test_executions"] == 1297,
            "aggregate quality count differs")
    receipt = read_json(HERE / "analyzer-tests.json")
    require(receipt["exit_code"] == 0 and receipt["script_sha256"] == sha(HERE / "analyzer_test.py")
            and receipt["analyzer_sha256"] == sha(HERE / "analyze_publication.py"),
            "analyzer test custody differs")
    result = subprocess.run(["python3", "-B", str(HERE / "analyzer_test.py")], capture_output=True, text=True)
    require(result.returncode == 0 and "Ran 2 tests" in result.stderr and "OK" in result.stderr,
            "analyzer tests failed")
    return {"status": "pass", "prior_checks": 12, "prior_tests": tests,
            "exact_source_reuse": True, "analyzer_tests": 2}


def run_component(name: str) -> dict[str, Any]:
    if name == "source":
        return validate_bindings()
    if name == "builds":
        return validate_builds()
    if name == "profiles":
        return validate_profiles()
    if name == "publication":
        profiles = validate_profiles()
        return {"profiles": profiles,
                "publication": validate_publication_analysis(profiles)}
    if name == "profile-analysis":
        return validate_raw_profile_analysis()
    if name == "decision":
        profiles = validate_profiles()
        validate_publication_analysis(profiles)
        validate_raw_profile_analysis()
        return validate_decision()
    if name == "quality":
        return validate_quality_reuse()
    if name == "cleanup":
        return validate_cleanup()
    if name == "seal":
        return validate_seal()
    raise EvidenceError(f"unknown component: {name}")


def run_bundle(component: str) -> dict[str, Any]:
    if component == "precleanup":
        order = ("source", "builds", "profiles", "publication", "profile-analysis", "quality", "decision")
    elif component == "all":
        order = ("source", "builds", "profiles", "publication", "profile-analysis", "quality", "decision", "cleanup", "seal")
    else:
        order = (component,)
    components: dict[str, Any] = {}
    for name in order:
        try:
            components[name] = {"status": "pass", "result": run_component(name)}
        except IncompleteError as error:
            components[name] = {"status": "incomplete", "error": str(error)}
        except (EvidenceError, OSError, subprocess.CalledProcessError) as error:
            components[name] = {"status": "fail", "error": str(error)}
    statuses = [value["status"] for value in components.values()]
    status = "fail" if "fail" in statuses else ("incomplete" if "incomplete" in statuses else "pass")
    return {
        "schema": "litchi-0528-attribution-verification-v1",
        "status": status,
        "scope": component,
        "performance_claim": "none",
        "components": components,
        "cleanup_seal_pending": component == "precleanup"
        or components.get("cleanup", {}).get("status") != "pass"
        or components.get("seal", {}).get("status") != "pass",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=(
        "source", "builds", "profiles", "publication", "profile-analysis",
        "decision", "cleanup", "seal", "quality", "precleanup", "all"),
        default="all")
    parser.add_argument("--strict", action="store_true",
                        help="return exit code 2 for incomplete or failed components")
    parser.add_argument("--output", type=Path,
                        help="write the verification result to this path")
    args = parser.parse_args()
    report = run_bundle(args.component)
    if args.output:
        require(not (SEAL.exists() and args.output.resolve().is_relative_to(HERE)),
                "verification output must be outside the sealed bundle")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps(report, indent=2, sort_keys=True))
    return 2 if args.strict and report["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
