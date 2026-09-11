#!/usr/bin/env python3
"""Fail-closed replay verifier for the 0514 XLSX fusion experiment.

The verifier authenticates the frozen source epochs, builds, corpus reports,
operation allocation vectors, pilot gates, and exact save profiles.  It also
understands the deliberately separate semantic no-op guard bundle.  A
rejected candidate is a valid outcome: the control evidence and candidate
pilot/guard evidence remain mandatory, while candidate full native captures
are optional in that branch.  No unavailable counter is interpreted as zero.

This program only reads retained evidence and the live source tree.  It never
builds the harness or runs a benchmark.
"""

from __future__ import annotations

import copy
import datetime as _datetime
import hashlib
import json
import math
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from tools.summarize_crud_baseline import _validate_elapsed  # noqa: E402
from tools.validate_perf_corpus_binding import validate_binding  # noqa: E402


BASE_REVISION = "33a21e0f087b4ead80ca63187c7cf30d0584b1f6"
PREVIOUS_MANIFEST = HERE.parent / "change-0513" / "source-manifest.json"
CASES = (
    "xlsx_one_cell_commit",
    "xlsx_one_percent_commit",
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
)
SHAPES = ("tiny", "medium", "dense-wide")
SAVED_CASES = {"xlsx_one_cell_commit_save", "xlsx_one_percent_commit_save"}
ABBA = (("before", "r1"), ("after", "r1"), ("after", "r2"), ("before", "r2"))
REPEATS = ("r1", "r2")
STATISTICS = ("p50", "mean", "p95", "p99")
DRIFT_LIMITS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
NATIVE_SAMPLES = 500
NATIVE_WARMUPS = 5
GUARD_SAMPLES = 100
GUARD_WARMUPS = 3
PILOT_SAMPLES = 20
PILOT_WARMUPS = 2
PREFLIGHT_SAMPLES = 1
PREFLIGHT_WARMUPS = 0
ALLOC_SAMPLES = 10
ALLOC_WARMUPS = 1
PROFILE_SAMPLES = 3
PROFILE_WARMUPS = 0
ALLOC_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
CALLGRIND_SUMMARY_RE = re.compile(r"^summary:\s*(\d+)\s*$", re.MULTILINE)
CALLGRIND_FN_RE = re.compile(r"^fn=\(\d+\)\s*(?P<name>.*)$")
CALLGRIND_CFN_RE = re.compile(r"^cfn=\(\d+\)\s*(?P<name>.*)$")
CALLGRIND_CALLS_RE = re.compile(r"^calls=(?P<count>[\d,]+)")
CALLGRIND_FUNCTION_RE = re.compile(r"^\s*[\d,]+\s+\([^)]*\)\s+\*\s+(?P<function>.+?)\s*$")
CALLGRIND_EDGE_RE = re.compile(
    r"^\s*[\d,]+\s+\([^)]*\)\s+>\s+(?P<function>.+?)\s+\((?P<count>[\d,]+)x\)(?:\s+\[[^]]*\])?\s*$"
)
REPORT_PREFIX = ("/usr/bin/time", "-v", "taskset", "-c", "2")
PROFILE_TOGGLE = "--toggle-collect=*litchi_perf_baseline::xlsx_commit_save_operation"
NATIVE_SCOPE = "Existing native commit or commit+save clock; setup, expected output, sink reservation, oracles and drop excluded"
ALLOC_SCOPE = "Allocation region begins before Instant and ends after elapsed; operation only, no setup/oracles/drop; instrumented elapsed and RSS excluded"
PROFILE_SCOPE = "Three xlsx_commit_save_operation helper calls only; fixture and expected-output commits/writes excluded; helper returns Commit before caller drop"
GUARD_PROBE = HERE / "guard-probe"
GUARD_PLAN = HERE / "guard-plan.json"
GUARD_SOURCE_MANIFEST = HERE / "guard-source-manifest.json"
GUARD_CANONICAL_SOURCES = (
    "tools/perf-baseline/src/allocation_metrics.rs",
    "tools/perf-baseline/src/bin/support/counting_allocator.rs",
)
GUARD_SCENARIOS = (
    "cold-first-cell-read",
    "cold-same-one-cell",
    "cold-same-one-percent",
    "warm-same-one-cell",
    "warm-same-one-percent",
    "warm-changed-one-cell",
    "warm-changed-one-percent",
)
GUARD_SHAPE_SIDES = {"tiny": 8, "medium": 32, "dense-wide": 256}
GUARD_SCOPE = (
    "Separate public guard; per-scenario commit or first-cell clock excludes setup, "
    "warming, oracles and result drop; allocator timing/RSS excluded."
)
GUARD_TIMER_SCOPE = (
    "Edit::commit or first public Worksheet::cell Store load only; fresh Workbook open, "
    "Store warming, edit preparation, output/readback oracles, and Commit/View drop are outside the clock"
)
GUARD_PROBE_NAME = "litchi-xlsx-public-commit-guard-v1"
GUARD_NORMAL_BINARY = "litchi-xlsx-commit-guard"
GUARD_ALLOCATOR_BINARY = "litchi-xlsx-commit-guard-alloc"
GUARD_NORMAL_COMMAND_BINARY = "xlsx-guard"
GUARD_ALLOCATOR_COMMAND_BINARY = "xlsx-guard-alloc"


class VerificationError(ValueError):
    """A missing, malformed, or inconsistent evidence item."""


def fail(message: str) -> None:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def _reject_constant(value: str) -> None:
    fail(f"non-finite JSON number {value!r}")


def load(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except VerificationError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot load {path}: {error}")


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")


def check_hash(value: Any, context: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{context} is not a lowercase SHA-256")
    return value


def canonical(value: Any, context: str = "value") -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        fail(f"{context} is not canonical JSON: {error}")


def finite(value: Any, context: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{context} must be numeric")
    result = float(value)
    require(math.isfinite(result), f"{context} must be finite")
    return result


def regular(path: Path, context: str) -> Path:
    require(path.is_file() and not path.is_symlink(),
            f"{context} is missing, symlinked, or not a regular file")
    return path


def stage_dir(stage: str) -> Path:
    require(stage in {"before", "after"}, f"unknown stage {stage!r}")
    return HERE / stage


def bundle_file(path: Path, context: str) -> Path:
    root = HERE.resolve()
    resolved = path.resolve()
    require(resolved.is_relative_to(root), f"{context} escapes the evidence bundle")
    return regular(path, context)


def safe_relative(value: Any, context: str) -> Path:
    require(isinstance(value, str) and value, f"{context} is not a relative path")
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts and path.as_posix() == value,
            f"{context} is unsafe")
    return path


def recorded_path(value: Any, expected_relative: str, context: str) -> Path:
    """Validate a path recorded by a command without requiring its origin root.

    Captures can be replayed from a relocated checkout.  Relative command
    paths therefore remain exact, while absolute paths are accepted only when
    their canonical, traversal-free suffix is the expected repository or
    role-relative path.
    """
    require(isinstance(value, str) and value, f"{context} path is missing")
    expected = Path(expected_relative)
    require(not expected.is_absolute() and ".." not in expected.parts
            and expected.as_posix() == expected_relative,
            f"{context} expected path is unsafe")
    path = Path(value)
    require(".." not in path.parts, f"{context} path contains traversal")
    if path.is_absolute():
        require(len(path.parts) >= len(expected.parts)
                and tuple(path.parts[-len(expected.parts):]) == expected.parts,
                f"{context} absolute path has the wrong role or basename")
    else:
        require(path.as_posix() == expected_relative,
                f"{context} relative path differs")
    return path


def repository_file(relative: str, context: str) -> Path:
    path = Path(relative)
    require(relative and not path.is_absolute() and ".." not in path.parts
            and path.as_posix() == relative, f"{context} is not a safe repository path")
    resolved = (REPO / path).resolve()
    require(resolved.is_relative_to(REPO.resolve()), f"{context} escapes the repository")
    return regular(resolved, context)


def git_blob_sha(revision: str, relative: str) -> str:
    try:
        data = subprocess.check_output(["git", "show", f"{revision}:{relative}"], cwd=REPO)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {revision}:{relative} from git: {error}")
    return hashlib.sha256(data).hexdigest()


def current_sources() -> dict[str, str]:
    try:
        names = subprocess.check_output(
            [
                "git", "ls-files", "--cached", "--others", "--exclude-standard", "-z",
                "crates", "tools/perf-baseline", "Cargo.toml", "Cargo.lock",
                "rust-toolchain.toml", ".cargo",
            ],
            cwd=REPO,
        ).decode("utf-8").split("\0")
    except (OSError, UnicodeError, subprocess.CalledProcessError) as error:
        fail(f"cannot enumerate source files: {error}")
    result: dict[str, str] = {}
    for name in sorted(item for item in names if item):
        if Path(name).suffix in {".rs", ".toml", ".lock"}:
            result[name] = sha(repository_file(name, f"source {name}"))
    return result


def manifest_path(stage: str) -> Path:
    if stage == "before":
        return HERE / "before" / "source-manifest.json"
    root = HERE / "source-manifest.json"
    if root.is_file():
        return root
    return HERE / "after" / "source-manifest.json"


def _manifest_candidates() -> list[Path]:
    return [
        HERE / "candidate-source-manifest.json",
        HERE / "candidate" / "source-manifest.json",
        HERE / "after" / "candidate-source-manifest.json",
        HERE / "after" / "source-manifest.json",
        HERE / "source-manifest.json",
    ]


def _manifest_matching_digest(digest: str) -> list[Path]:
    matches = []
    for path in _manifest_candidates():
        if path.is_file() and not path.is_symlink() and sha(path) == digest:
            matches.append(path)
    return matches


def validate_manifest(value: Any, context: str) -> dict[str, str]:
    require(isinstance(value, dict) and value, f"{context} is empty")
    result: dict[str, str] = {}
    for name, digest in value.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and ".." not in Path(name).parts, f"{context} contains an unsafe path")
        result[name] = check_hash(digest, f"{context}.{name}")
    return result


def load_manifest(path: Path, context: str) -> dict[str, str]:
    return validate_manifest(load(path), context)


def _patch_paths(text: str, context: str) -> list[str]:
    """Return the exact repository paths named by unified diff sections."""
    paths: list[str] = []
    for line in text.splitlines():
        if not line.startswith("diff --git "):
            continue
        fields = line.split()
        require(len(fields) == 4, f"{context} has a malformed diff header")
        left, right = fields[2], fields[3]
        require(left.startswith("a/") and right.startswith("b/"),
                f"{context} has an unsafe diff header")
        left = left[2:]
        right = right[2:]
        require(left == right, f"{context} contains a rename or copy")
        path = safe_relative(left, f"{context} path")
        require(path.as_posix() == left, f"{context} path is not canonical")
        paths.append(left)
    require(paths and len(paths) == len(set(paths)),
            f"{context} does not contain a unique diff-section inventory")
    return paths


def _base_blob(relative: str) -> bytes | None:
    try:
        return subprocess.check_output(
            ["git", "show", f"{BASE_REVISION}:{relative}"],
            cwd=REPO, stderr=subprocess.DEVNULL,
        )
    except (OSError, subprocess.CalledProcessError):
        return None


def _replay_rejected_candidate(before: dict[str, str], after: dict[str, str],
                               after_manifest_path: Path) -> dict[str, Any]:
    """Authenticate and replay a rejected candidate patch in an ephemeral tree.

    Only the changed paths are materialized from ``git show``.  The patch is
    applied there with no repository checkout, and the resulting file
    inventory and hashes are compared to the retained candidate manifest.
    The temporary tree is removed by ``TemporaryDirectory`` on every exit.
    """
    patch_meta = load(HERE / "candidate-patch.json")
    require(isinstance(patch_meta, dict), "candidate-patch.json must be an object")
    required_meta = {"base_revision", "path", "sha256", "source_manifest_sha256",
                     "changed", "exact_replay_passed"}
    require(required_meta <= set(patch_meta), "candidate-patch.json is incomplete")
    require(patch_meta.get("base_revision") == BASE_REVISION,
            "candidate patch base revision differs")
    patch_relative = safe_relative(patch_meta.get("path"), "candidate patch path")
    require(patch_relative.as_posix() == "candidate.patch",
            "candidate patch path is not the retained candidate.patch")
    patch_path = bundle_file(HERE / patch_relative, "candidate.patch")
    patch_digest = check_hash(patch_meta.get("sha256"), "candidate patch sha256")
    require(sha(patch_path) == patch_digest, "candidate.patch hash differs")
    require(patch_meta.get("source_manifest_sha256") == sha(after_manifest_path),
            "candidate patch manifest hash differs")
    require(patch_meta.get("exact_replay_passed") is True,
            "candidate patch metadata does not record exact replay")
    changed_value = patch_meta.get("changed")
    require(isinstance(changed_value, list) and changed_value
            and all(isinstance(item, str) for item in changed_value)
            and len(set(changed_value)) == len(changed_value),
            "candidate patch changed inventory is malformed")
    changed = [safe_relative(item, "candidate patch changed path").as_posix()
               for item in changed_value]
    require(sorted(changed) == sorted(
        name for name in set(before) | set(after) if before.get(name) != after.get(name)
    ), "candidate patch changed inventory differs from source manifests")
    patch_text = patch_path.read_text(encoding="utf-8", errors="strict")
    require(sorted(_patch_paths(patch_text, "candidate.patch")) == sorted(changed),
            "candidate patch diff sections differ from its changed inventory")

    restoration = load(HERE / "restoration.json")
    require(isinstance(restoration, dict), "restoration.json must be an object")
    require(restoration.get("base_revision") == BASE_REVISION,
            "restoration base revision differs")
    parse_time(restoration.get("restored_utc"), "restoration restored_utc")
    require(restoration.get("production_and_tests_restored") is True,
            "restoration does not attest production/test restoration")
    require(restoration.get("exact_source_manifest_match") is True,
            "restoration does not attest exact source restoration")
    before_manifest_digest = sha(manifest_path("before"))
    require(restoration.get("source_manifest_sha256") == before_manifest_digest,
            "restoration source manifest hash differs from control")
    require(restoration.get("candidate_manifest_sha256") == sha(after_manifest_path),
            "restoration candidate manifest hash differs")
    require(restoration.get("candidate_patch_sha256") == patch_digest,
            "restoration candidate patch hash differs")
    restored = restoration.get("restored_paths")
    require(isinstance(restored, list) and all(isinstance(item, str) for item in restored)
            and len(set(restored)) == len(restored),
            "restoration path inventory is malformed")
    restored_paths = [safe_relative(item, "restoration path").as_posix() for item in restored]
    require(sorted(restored_paths) == sorted(changed),
            "restoration path inventory differs from candidate patch")

    # Populate only base blobs touched by the candidate.  New-file paths are
    # intentionally left absent so git-apply must create them itself.
    with tempfile.TemporaryDirectory(prefix="litchi-0514-rejected-replay-") as temporary:
        root = Path(temporary)
        require(root.is_dir() and not root.is_symlink(),
                "candidate replay temporary directory is invalid")
        for name in changed:
            blob = _base_blob(name)
            if name in before:
                require(blob is not None, f"base blob is missing for {name}")
                require(hashlib.sha256(blob).hexdigest() == before[name],
                        f"base blob hash differs for {name}")
                target = root / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(blob)
            else:
                require(blob is None, f"new candidate path already exists at base: {name}")

        check = subprocess.run(
            ["git", "apply", "--check", str(patch_path)],
            cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        require(check.returncode == 0, "candidate patch does not apply cleanly to base blobs")
        applied = subprocess.run(
            ["git", "apply", "--whitespace=nowarn", str(patch_path)],
            cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        require(applied.returncode == 0, "candidate patch replay failed")

        actual_files: set[str] = set()
        for path in root.rglob("*"):
            require(not path.is_symlink(), "candidate replay created a symlink")
            if path.is_file():
                actual_files.add(path.relative_to(root).as_posix())
        expected_files = {name for name in changed if name in after}
        require(actual_files == expected_files,
                f"candidate replay created an unexpected file inventory: {sorted(actual_files)!r}")
        for name in expected_files:
            target = root / name
            require(sha(target) == after[name],
                    f"candidate replay hash differs from candidate manifest: {name}")

    return {
        "candidate_patch_sha256": patch_digest,
        "candidate_manifest_sha256": sha(after_manifest_path),
        "changed": changed,
        "restored_paths": restored_paths,
        "exact_replay_passed": True,
        "temporary_tree_cleaned": True,
    }


def load_decision() -> tuple[dict[str, Any], bool]:
    decision = load(HERE / "decision.json")
    require(isinstance(decision, dict), "decision.json must be an object")
    status = decision.get("status", decision.get("outcome", decision.get("decision")))
    require(isinstance(status, str), "decision must disclose status/outcome")
    normalized = status.lower().replace("-", "_").replace(" ", "_")
    if normalized in {"retained", "retain", "accepted", "accept", "kept", "keep", "candidate"}:
        kept = True
    elif normalized in {"rejected", "reject", "reverted", "revert", "baseline", "baseline_only", "retain_baseline"}:
        kept = False
    else:
        fail(f"unrecognized candidate decision {status!r}")
    for key in ("candidate_kept", "keep_candidate", "production_change_retained"):
        if key in decision:
            require(isinstance(decision[key], bool) and decision[key] == kept,
                    f"decision.{key} disagrees with status")
    if not kept:
        reason = decision.get("reason", decision.get("disposition"))
        require(isinstance(reason, (str, dict)) and reason, "rejected decision needs a reason")
    return decision, kept


def verify_sources(decision: dict[str, Any], kept: bool) -> dict[str, Any]:
    plan = load(HERE / "plan.json")
    require(isinstance(plan, dict) and plan.get("base_revision") == BASE_REVISION,
            "plan base revision does not match the frozen 0513 candidate")
    require(plan.get("native_protocol") and plan.get("allocator_protocol")
            and plan.get("profile_protocol") and plan.get("semantic_noop_guard"),
            "plan omits a required 0514 evidence protocol")
    before = load_manifest(manifest_path("before"), "before source manifest")
    previous = load_manifest(PREVIOUS_MANIFEST, "previous source manifest")
    require(before == previous, "control manifest does not preserve the 0513 source epoch")
    after_path = manifest_path("after")
    if not kept:
        # A rejection may restore the live tree and leave the candidate
        # manifest under an explicit candidate name.  Prefer the retained
        # manifest whose contents differ from the control epoch so the
        # rejected production source set remains authenticated.
        candidates = []
        for candidate in _manifest_candidates():
            if candidate.is_file() and not candidate.is_symlink():
                value = load_manifest(candidate, f"candidate source manifest {candidate.name}")
                if value != before:
                    candidates.append((candidate, value))
        if candidates:
            require(len({canonical(value, "candidate source manifest")
                          for _, value in candidates}) == 1,
                    "rejected decision retains multiple distinct candidate manifests")
            after_path = candidates[0][0]
    after = load_manifest(after_path, "candidate source manifest")
    current = current_sources()
    if kept:
        require(after == current, "candidate source tree differs from its manifest")
    else:
        # A rejection may restore the live tree.  The candidate manifest is
        # still retained and is authenticated by the candidate build receipts.
        require(current in (before, after),
                "rejected decision has neither restored baseline nor live candidate sources")
    for name, digest in before.items():
        require(git_blob_sha(BASE_REVISION, name) == digest,
                f"control source is not from the declared base revision: {name}")
    if kept or current == after:
        for name, digest in after.items():
            require(sha(repository_file(name, f"candidate source {name}")) == digest,
                    f"candidate source changed after build: {name}")
    changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
    require(changed, "candidate source manifest does not describe a production change")
    production = "crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs"
    require(production in changed, "candidate does not change XLSX transaction production code")
    require(all(name.startswith("crates/litchi-xlsx/") for name in changed),
            f"source change set leaves the XLSX crate: {changed!r}")
    require(any("tests" in Path(name).parts for name in changed),
            "candidate source change set omits focused XLSX tests")
    # No harness source is allowed to move in this production-only batch.
    require(not any(name.startswith("tools/perf-baseline/") for name in changed),
            "0514 candidate unexpectedly changes the performance harness")
    rejection_patch = None
    if not kept:
        rejection_patch = _replay_rejected_candidate(before, after, after_path)

    adr = load(HERE / "adr-manifest.json")
    require(isinstance(adr, dict) and adr.get("revision") == BASE_REVISION,
            "ADR manifest revision differs from the frozen base")
    adr_files = adr.get("files")
    require(isinstance(adr_files, dict) and len(adr_files) == 30,
            "ADR manifest must contain 30 files")
    for name, digest in adr_files.items():
        check_hash(digest, f"ADR {name}")
        require(sha(repository_file(name, f"ADR {name}")) == digest, f"ADR changed: {name}")

    verifier_sources = load(HERE / "verifier-sources.json")
    require(isinstance(verifier_sources, dict) and verifier_sources,
            "verifier-sources.json must be a non-empty object")
    for name, digest in verifier_sources.items():
        check_hash(digest, f"verifier source {name}")
        require(sha(repository_file(name, f"verifier source {name}")) == digest,
                f"verifier source changed: {name}")

    fixtures = load(HERE / "compile-fixtures.json")
    expected_fixtures = {
        "test-data/rtf/watermark.rtf",
        "test-data/poi/test-data/spreadsheet/54016.xls",
    }
    require(isinstance(fixtures, dict) and set(fixtures) == expected_fixtures,
            "compile-fixtures.json does not bind the two compiled fixtures")
    for name, digest in fixtures.items():
        check_hash(digest, f"fixture {name}")
        require(sha(repository_file(name, f"fixture {name}")) == digest,
                f"fixture changed: {name}")
        require(git_blob_sha(BASE_REVISION, name) == digest,
                f"fixture is not from the base revision: {name}")
    return {
        "base_revision": BASE_REVISION,
        "before_manifest_sha256": sha(manifest_path("before")),
        "after_manifest_sha256": sha(after_path),
        "source_files_before": len(before),
        "source_files_after": len(after),
        "live_source_epoch": "candidate" if current == after else "restored_baseline",
        "changed": changed,
        "adr_files": len(adr_files),
        "rejection_patch": rejection_patch,
    }


def build_receipt_path(stage: str, allocator: bool = False) -> Path:
    if allocator:
        return stage_dir(stage) / "allocator-build-receipt.json"
    return HERE / "before" / "build-receipt.json" if stage == "before" else HERE / "build-receipt.json"


def build_log_path(stage: str, allocator: bool = False) -> Path:
    return build_receipt_path(stage, allocator).with_name(
        "allocator-build.log" if allocator else "build.log"
    )


def verify_build(stage: str, allocator: bool = False) -> dict[str, Any]:
    receipt_path = build_receipt_path(stage, allocator)
    receipt = load(receipt_path)
    label = f"{stage} {'allocator ' if allocator else ''}build"
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{label} did not exit successfully")
    require(receipt.get("source_unchanged") is True, f"{label} source changed during build")
    binary = check_hash(receipt.get("binary_sha256"), f"{label}.binary_sha256")
    manifest_digest = check_hash(receipt.get("source_manifest_sha256"),
                                  f"{label}.source_manifest_sha256")
    manifest = manifest_path(stage)
    # A rejected candidate can retain a candidate-only manifest at an
    # explicitly recorded path after restoring the live source tree.
    recorded_path = receipt.get("source_manifest_path")
    if stage == "after" and isinstance(recorded_path, str):
        candidate = bundle_file(HERE / safe_relative(recorded_path, f"{label}.source_manifest_path"),
                                f"{label}.source_manifest_path")
        require(sha(candidate) == manifest_digest, f"{label} recorded manifest hash differs")
        manifest = candidate
    else:
        if stage == "after" and (not manifest.is_file() or sha(manifest) != manifest_digest):
            matches = _manifest_matching_digest(manifest_digest)
            require(matches, f"{label} source manifest is not retained")
            manifest = matches[0]
        require(sha(manifest) == manifest_digest, f"{label} source manifest is not bound")
    log_path = build_log_path(stage, allocator)
    require(receipt.get("log_sha256") == sha(log_path), f"{label} log hash is not bound")
    elapsed = finite(receipt.get("elapsed_seconds"), f"{label}.elapsed_seconds")
    require(elapsed > 0, f"{label} elapsed time is not positive")
    command = receipt.get("command")
    require(isinstance(command, list) and command and all(isinstance(item, str) for item in command),
            f"{label}.command is not an argv list")
    expected_normal = [
        "cargo", "build", "--release", "--bin", "litchi-perf-baseline",
        "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
    ]
    expected_allocator = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics",
        "--bin", "litchi-perf-baseline-alloc",
    ]
    require(command == (expected_allocator if allocator else expected_normal),
            f"{label} command differs from the frozen build")
    return {
        "binary_sha256": binary,
        "source_manifest_sha256": manifest_digest,
        "log_sha256": sha(log_path),
        "allocator": allocator,
    }


def option(command: list[str], flag: str, context: str) -> str:
    positions = [index for index, value in enumerate(command) if value == flag]
    require(len(positions) == 1 and positions[0] + 1 < len(command),
            f"{context} command must contain one {flag} value")
    value = command[positions[0] + 1]
    require(isinstance(value, str) and value, f"{context} {flag} value is empty")
    return value


def equals_option(command: list[str], flag: str, context: str) -> str:
    prefix = flag + "="
    values = [item[len(prefix):] for item in command if item.startswith(prefix)]
    require(flag not in command and len(values) == 1 and values[0],
            f"{context} command must contain one {prefix} value")
    return values[0]


def verify_artifacts(stage: str, name: str, receipt: dict[str, Any], build: dict[str, Any],
                     required: set[str]) -> None:
    context = f"{stage}/{name}"
    require(receipt.get("exit_code") == 0, f"{context} did not exit successfully")
    require(receipt.get("source_unchanged") is True, f"{context} source changed during capture")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{context} binary hash differs from build")
    check_hash(receipt.get("binary_sha256"), f"{context}.binary_sha256")
    require(receipt.get("source_manifest_sha256") == build["source_manifest_sha256"],
            f"{context} source manifest differs from build")
    check_hash(receipt.get("source_manifest_sha256"), f"{context}.source_manifest_sha256")
    elapsed = finite(receipt.get("elapsed_seconds"), f"{context}.elapsed_seconds")
    require(elapsed > 0, f"{context} receipt elapsed is not positive")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and required <= set(artifacts),
            f"{context} receipt does not bind all required artifacts")
    for relative, digest in artifacts.items():
        path = safe_relative(relative, f"{context} artifact path")
        artifact = bundle_file(stage_dir(stage) / path, f"{context} artifact {relative}")
        check_hash(digest, f"{context} artifact {relative}")
        require(sha(artifact) == digest, f"{context} artifact hash changed: {relative}")


def verify_command(command: Any, name: str, samples: int, warmup: int,
                   allocator: bool = False, profile: bool = False) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{name}.command is not an argv list")
    require(tuple(command[:5]) == REPORT_PREFIX, f"{name} is not pinned to CPU 2")
    basename = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    binaries = [item for item in command if Path(item).name == basename]
    require(len(binaries) == 1, f"{name} command does not identify one {basename}")
    require(option(command, "--samples", name) == str(samples), f"{name} sample count differs")
    require(option(command, "--warmup", name) == str(warmup), f"{name} warmup count differs")
    report = option(command, "--json", name)
    catalog = option(command, "--corpus-manifest", name)
    require(Path(report).name == f"{name}-report.json", f"{name} report path is not bound")
    require(Path(catalog).name == f"{name}-catalog.json", f"{name} catalog path is not bound")
    if profile:
        require("valgrind" in command and "--tool=callgrind" in command,
                f"{name} is not a Callgrind command")
        require("--collect-atstart=no" in command and PROFILE_TOGGLE in command,
                f"{name} does not bind the exact save-helper profile boundary")
        output = equals_option(command, "--callgrind-out-file", name)
        require(Path(output).name == f"{name}.out", f"{name} Callgrind output path is not bound")
        require(option(command, "--case", name) == "xlsx_one_percent_commit_save",
                f"{name} profile case differs")
        require(option(command, "--xlsx-shape", name) == "dense-wide",
                f"{name} profile shape differs")
    else:
        require(option(command, "--case", name) == ",".join(CASES),
                f"{name} case selection differs")
        require(option(command, "--xlsx-shape", name) == ",".join(SHAPES),
                f"{name} shape selection differs")
    return command


def verify_report_identity(report: dict[str, Any], build: dict[str, Any], samples: int,
                           warmup: int, cases: list[str], shapes: list[str],
                           context: str, allocator: bool = False) -> None:
    require(report.get("schema_version") == 1, f"{context} schema version is not 1")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{context}.tool is missing")
    expected = {
        "name": "litchi-perf-baseline", "version": "0.1.0",
        "binary": "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline",
        "profile": "release", "target_os": "linux", "target_arch": "x86_64",
        "instrumentation": "system_allocator_operation_scoped" if allocator else "none",
    }
    for key, value in expected.items():
        require(tool.get(key) == value, f"{context}.tool.{key} identity differs")
    if allocator:
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{context} allocator counter revision differs")
    else:
        require("allocator_counter_revision" not in tool,
                f"{context} normal report exposes allocator counter revision")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{context}.binary_identity is missing")
    require(identity.get("binary_sha256") == build["binary_sha256"],
            f"{context} report binary does not match build")
    check_hash(identity.get("binary_sha256"), f"{context}.binary_identity.binary_sha256")
    require(Path(str(identity.get("path"))).name == expected["binary"],
            f"{context}.binary_identity path does not name selected binary")
    require(identity.get("executable") is True and identity.get("profile") == "release",
            f"{context}.binary_identity executable/profile is invalid")
    require(isinstance(identity.get("binary_bytes"), int) and identity["binary_bytes"] > 0,
            f"{context}.binary_identity.binary_bytes is invalid")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{context}.environment is missing")
    require(environment.get("git_revision") == BASE_REVISION,
            f"{context}.environment.git_revision differs")
    require(isinstance(environment.get("git_worktree_dirty"), bool),
            f"{context}.environment.git_worktree_dirty is not boolean")
    require(environment.get("cpu_affinity") == "2", f"{context} is not pinned to CPU 2")
    expected_allocator = "CountingSystemAllocator(std::alloc::System)" if allocator else "Rust system allocator"
    require(environment.get("allocator") == expected_allocator,
            f"{context}.environment allocator identity differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{context}.configuration is missing")
    require(configuration.get("samples_per_case") == samples,
            f"{context} report sample count differs")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{context} report warmup count differs")
    require(configuration.get("cases") == cases, f"{context} case selection differs")
    require(configuration.get("xlsx_shapes") == shapes, f"{context} shape selection differs")
    require(configuration.get("filesystem_process_isolated") is True,
            f"{context} is not process isolated")
    require(configuration.get("filesystem_fresh_child_per_sample") is True,
            f"{context} does not use a fresh child per sample")


def verify_sink(value: Any, context: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{context}.sink is missing")
    required = {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets"}
    require(set(value) == required, f"{context}.sink schema differs")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        require(isinstance(value[key], int) and not isinstance(value[key], bool) and value[key] >= 0,
                f"{context}.sink.{key} is invalid")
    buckets = value["write_size_buckets"]
    expected = {"bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384",
                "bytes_16385_to_65536", "bytes_over_65536"}
    require(isinstance(buckets, dict) and set(buckets) == expected,
            f"{context}.sink.write_size_buckets schema differs")
    for key, count in buckets.items():
        require(isinstance(count, int) and not isinstance(count, bool) and count >= 0,
                f"{context}.sink.write_size_buckets.{key} is invalid")
    require(sum(buckets.values()) == value["write_calls"],
            f"{context}.sink write buckets do not sum to write_calls")
    return value


def verify_metric_vector(value: Any, samples: int, context: str, allow_pattern: bool = False) -> None:
    require(isinstance(value, dict), f"{context} must be a metric vector")
    status = value.get("status")
    require(status in {"measured", "not_applicable", "unavailable", "overflow"},
            f"{context}.status is invalid")
    require(isinstance(value.get("scope"), str) and value["scope"], f"{context}.scope is missing")
    values = value.get("values")
    if status != "measured":
        require(values is None, f"{context} publishes values with status {status}")
        return
    require(isinstance(values, list) and len(values) == samples,
            f"{context}.values must contain {samples} values")
    for index, item in enumerate(values):
        if allow_pattern:
            require(item in {"sequential", "random", "unknown"},
                    f"{context}.values[{index}] is not a read pattern")
        else:
            require(isinstance(item, int) and not isinstance(item, bool) and item >= 0,
                    f"{context}.values[{index}] is not a non-negative integer")


def verify_metric_tree(value: Any, samples: int, context: str) -> None:
    if isinstance(value, dict):
        if "status" in value and "scope" in value and set(value) <= {"status", "scope", "values"}:
            verify_metric_vector(value, samples, context, allow_pattern=("pattern" in context))
            return
        for key, child in value.items():
            verify_metric_tree(child, samples, f"{context}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            verify_metric_tree(child, samples, f"{context}[{index}]")


def _measured_values(vector: Any, samples: int, context: str) -> list[int]:
    require(isinstance(vector, dict) and vector.get("status") == "measured",
            f"{context} is not measured")
    values = vector.get("values")
    require(isinstance(values, list) and len(values) == samples,
            f"{context} has wrong cardinality")
    require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0 for item in values),
            f"{context} has invalid values")
    return values


def verify_operation_metrics(row: dict[str, Any], samples: int, context: str,
                             allocator: bool, elapsed_order: list[int]) -> dict[str, Any]:
    operation = row.get("operation_metrics")
    require(isinstance(operation, dict), f"{context}.operation_metrics is missing")
    require(operation.get("sample_count") == samples, f"{context} operation sample count differs")
    indices = operation.get("sample_indices")
    require(isinstance(indices, list) and sorted(indices) == list(range(samples))
            and len(set(indices)) == samples, f"{context} operation sample_indices is not a permutation")
    require(indices == elapsed_order, f"{context} operation vectors are not aligned to elapsed samples")
    require(operation.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{context} operation alignment is not explicit")
    verify_metric_tree(operation, samples, f"{context}.operation_metrics")
    allocation = operation.get("allocation")
    require(isinstance(allocation, dict), f"{context}.operation allocation envelope is missing")
    if allocator:
        require(allocation.get("status") == "measured",
                f"{context} allocator metrics are not measured")
        require(allocation.get("scope") == "operation_global_system_allocator",
                f"{context} allocator scope differs")
        values = {field: _measured_values(allocation.get(field), samples,
                                           f"{context}.allocation.{field}")
                  for field in ALLOC_FIELDS}
        increments = [peak - base for peak, base in zip(
            values["region_peak_live_bytes"], values["live_bytes_before"])]
        require(min(increments) >= 0, f"{context} region peak precedes live-byte baseline")
        require(all(region >= after >= 0 and region <= peak_after
                    for region, after, peak_after in zip(
                        values["region_peak_live_bytes"], values["live_bytes_after"],
                        values["peak_live_bytes_after"])),
                f"{context} region peak is outside live/high-water bounds")
        return {
            "status": "measured", "scope": allocation["scope"],
            "absolute": values, "incremental_region_peak": increments,
        }
    require(all(isinstance(allocation.get(field), dict)
                and allocation[field].get("status") == "unavailable"
                and allocation[field].get("scope") == "operation_global_system_allocator"
                and allocation[field].get("values") is None for field in ALLOC_FIELDS),
            f"{context} normal allocation must be unavailable without fabricated values")
    return {"status": "unavailable"}


def verify_sink_vectors(row: dict[str, Any], operation: dict[str, Any], samples: int,
                        context: str) -> None:
    if row["case"] not in SAVED_CASES:
        return
    sink = verify_sink(row.get("sink"), context)
    op_sink = operation.get("sink")
    require(isinstance(op_sink, dict) and op_sink.get("write_status") == "measured",
            f"{context} save operation sink vectors are not measured")
    mapping = {
        "accepted_bytes": "accepted_bytes", "write_calls": "write_calls",
        "largest_write": "largest_write",
    }
    for output_name, vector_name in mapping.items():
        values = _measured_values(op_sink.get(vector_name), samples,
                                  f"{context}.operation.sink.{vector_name}")
        require(set(values) == {sink[output_name]},
                f"{context} operation sink {output_name} differs from retained sink")
    buckets = op_sink.get("write_size_buckets")
    require(isinstance(buckets, dict) and buckets.get("status") == "measured",
            f"{context} operation sink buckets are not measured")
    for name, expected in sink["write_size_buckets"].items():
        values = _measured_values(buckets.get(name), samples,
                                  f"{context}.operation.sink.write_size_buckets.{name}")
        require(set(values) == {expected},
                f"{context} operation sink bucket {name} differs from retained sink")


def verify_row(row: Any, case: str, shape: str, samples: int, context: str,
               allocator: bool) -> dict[str, Any]:
    require(isinstance(row, dict), f"{context} is not an object")
    require(row.get("case") == case, f"{context}.case differs")
    corpus = row.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == shape,
            f"{context}.corpus shape differs")
    require(corpus.get("name") == f"xlsx-{shape}", f"{context}.corpus.name differs")
    require(corpus.get("generator") == "litchi-xlsx-synthetic-v1",
            f"{context}.corpus.generator differs")
    require(corpus.get("package_format") == "XLSX/OPC/ZIP",
            f"{context}.corpus.package_format differs")
    check_hash(corpus.get("archive_sha256"), f"{context}.corpus.archive_sha256")
    check_hash(corpus.get("target_payload_sha256"), f"{context}.corpus.target_payload_sha256")
    try:
        elapsed = _validate_elapsed(row, samples, context)
    except Exception as error:
        fail(f"{context} elapsed vector is invalid: {error}")
    if case in SAVED_CASES:
        verify_sink(row.get("sink"), context)
    else:
        require(row.get("sink") is None, f"{context} non-save row has a sink")
    require(row.get("source") is None, f"{context} unexpectedly publishes source evidence")
    operation = row.get("operation_metrics")
    require(isinstance(operation, dict), f"{context} operation metrics are missing")
    allocation = verify_operation_metrics(row, samples, context, allocator,
                                          elapsed["sample_order"])
    verify_sink_vectors(row, operation, samples, context)
    output = row.get("output_sha256")
    if output is not None:
        check_hash(output, f"{context}.output_sha256")
    return {"row": row, "elapsed": elapsed, "allocation": allocation}


def verify_capture(stage: str, lane: str, build: dict[str, Any], samples: int,
                   warmup: int, allocator: bool = False, profile: bool = False) -> tuple[dict[str, Any], dict[str, Any]]:
    name = "allocator-" + lane if allocator else lane
    directory = stage_dir(stage)
    receipt = load(directory / f"{name}-receipt.json")
    required = {f"{name}-report.json", f"{name}-catalog.json", f"{name}.log"}
    if profile:
        required.update({f"{name}.out", f"{name}-inclusive.txt", f"{name}-exclusive.txt"})
    verify_artifacts(stage, name, receipt, build, required)
    command = verify_command(receipt.get("command"), name, samples, warmup, allocator, profile)
    expected_scope = ALLOC_SCOPE if allocator else (PROFILE_SCOPE if profile else NATIVE_SCOPE)
    require(receipt.get("scope") == expected_scope, f"{stage}/{name} scope is not explicit")
    report = load(directory / f"{name}-report.json")
    catalog = load(directory / f"{name}-catalog.json")
    require(isinstance(report, dict) and isinstance(catalog, dict),
            f"{stage}/{name} report/catalog must be objects")
    try:
        validate_binding(report, catalog)
    except Exception as error:
        fail(f"{stage}/{name} report/catalog binding failed: {error}")
    report_cases = ["xlsx_one_percent_commit_save"] if profile else list(CASES)
    report_shapes = ["dense-wide"] if profile else list(SHAPES)
    verify_report_identity(report, build, samples, warmup, report_cases, report_shapes,
                           f"{stage}/{name}", allocator)
    rows = report.get("results")
    expected_keys = [(shape, case) for shape in report_shapes for case in report_cases]
    require(isinstance(rows, list) and len(rows) == len(expected_keys),
            f"{stage}/{name} has the wrong report row count")
    parsed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, ((shape, case), row) in enumerate(zip(expected_keys, rows)):
        parsed[case, shape] = verify_row(
            row, case, shape, samples, f"{stage}/{name}.results[{index}]", allocator
        )
    profile_info = verify_profile_artifacts(stage, name, receipt) if profile else None
    return receipt, {
        "report": report, "catalog": catalog, "rows": parsed,
        "annotations": profile_info, "command": command,
    }


def _profile_edges_from_raw(text: str) -> list[tuple[str, str, int]]:
    edges: list[tuple[str, str, int]] = []
    parent: str | None = None
    child: str | None = None
    for line in text.splitlines():
        match = CALLGRIND_FN_RE.match(line)
        if match:
            parent = match.group("name").strip()
            child = None
            continue
        match = CALLGRIND_CFN_RE.match(line)
        if match:
            child = match.group("name").strip()
            continue
        match = CALLGRIND_CALLS_RE.match(line)
        if match and parent is not None and child is not None:
            edges.append((parent, child, int(match.group("count").replace(",", ""))))
            child = None
    return edges


def _profile_edges_from_annotation(text: str) -> list[tuple[str, str, int]]:
    edges: list[tuple[str, str, int]] = []
    parent: str | None = None
    for line in text.splitlines():
        function = CALLGRIND_FUNCTION_RE.match(line)
        if function:
            parent = function.group("function").strip()
            continue
        edge = CALLGRIND_EDGE_RE.match(line)
        if edge and parent is not None:
            edges.append((parent, edge.group("function").strip(),
                          int(edge.group("count").replace(",", ""))))
    return edges


def verify_profile_artifacts(stage: str, name: str, receipt: dict[str, Any]) -> dict[str, Any]:
    raw = bundle_file(stage_dir(stage) / f"{name}.out", f"{stage}/{name} raw Callgrind")
    text = raw.read_text(encoding="utf-8", errors="strict")
    summaries = CALLGRIND_SUMMARY_RE.findall(text)
    require(len(summaries) == 1 and int(summaries[0]) > 0,
            f"{stage}/{name} raw Callgrind summary is missing")
    edges = _profile_edges_from_raw(text)
    annotations: dict[str, list[tuple[str, str, int]]] = {}
    annotation_hashes = {}
    for suffix in ("inclusive", "exclusive"):
        path = bundle_file(stage_dir(stage) / f"{name}-{suffix}.txt",
                          f"{stage}/{name} {suffix} annotation")
        annotation_text = path.read_text(encoding="utf-8", errors="strict")
        require(path.stat().st_size > 0, f"{stage}/{name} {suffix} annotation is empty")
        annotations[suffix] = _profile_edges_from_annotation(annotation_text)
        annotation_hashes[suffix] = sha(path)
        artifacts = receipt.get("artifacts", {})
        require(isinstance(artifacts, dict)
                and artifacts.get(f"{name}-{suffix}.txt") == annotation_hashes[suffix],
                f"{stage}/{name} {suffix} annotation hash is not bound")
    # Direct-call evidence comes from the inclusive function annotation.  The
    # raw Callgrind format may split function names across fn/cfn mapping
    # records, so accepting a raw edge alone would lose the parent identity.
    annotated_edges = annotations["inclusive"]
    helper_to_commit = [count for parent, child, count in annotated_edges
                        if "xlsx_commit_save_operation" in parent and "Edit::commit" in child]
    runner_to_helper = [count for parent, child, count in annotated_edges
                        if "run_xlsx_update_commit_save" in parent
                        and "xlsx_commit_save_operation" in child]
    writer_tokens = ("PackageWriter", "CountingSink", "Workbook::write_to", "Writer<")
    helper_to_writer = [count for parent, child, count in annotated_edges
                        if "xlsx_commit_save_operation" in parent
                        and any(token in child for token in writer_tokens)]
    require(helper_to_commit == [3],
            f"{stage}/{name} profile does not prove three helper-to-commit calls: {helper_to_commit!r}")
    require(runner_to_helper == [3],
            f"{stage}/{name} profile does not prove three runner-to-helper calls: {runner_to_helper!r}")
    require(3 in helper_to_writer,
            f"{stage}/{name} profile has no three-call retained writer subtree: {helper_to_writer!r}")
    return {
        "summary_ir": int(summaries[0]),
        "commit_direct_calls": 3,
        "runner_to_helper_calls": 3,
        "helper_to_writer_calls": 3,
        "raw_sha256": sha(raw),
        "annotation_sha256": annotation_hashes,
        "scope": PROFILE_SCOPE,
    }


def verify_row_identity(current: dict[str, Any], expected: dict[str, Any], context: str) -> None:
    require(canonical(current.get("corpus"), context + ".corpus") ==
            canonical(expected.get("corpus"), context + ".corpus"),
            f"{context} corpus identity differs")
    require(current.get("sink") == expected.get("sink"), f"{context} sink identity differs")
    require(current.get("output_sha256") == expected.get("output_sha256"),
            f"{context} output identity differs")


def capture_identity(captures: dict[tuple[str, str], dict[str, Any]], context: str) -> None:
    reference = captures["before", "r1"]["rows"]
    for (stage, lane), capture in captures.items():
        for key, current in capture["rows"].items():
            if key in reference:
                verify_row_identity(current["row"], reference[key]["row"],
                                    f"{context}/{stage}/{lane}/{key[0]}/{key[1]}")


def percent(after: float, before: float) -> float:
    require(before > 0, "cannot compute a percent change from a non-positive value")
    return (after / before - 1.0) * 100.0


def native_comparisons(captures: dict[tuple[str, str], dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    pairs: list[dict[str, Any]] = []
    drift: list[dict[str, Any]] = []
    for repeat in REPEATS:
        before = captures["before", repeat]["rows"]
        after = captures["after", repeat]["rows"]
        for shape in SHAPES:
            for case in CASES:
                key = (case, shape)
                b = before[key]["elapsed"]["statistics"]
                a = after[key]["elapsed"]["statistics"]
                changes = {metric: percent(float(a[metric]), float(b[metric])) for metric in STATISTICS}
                throughput = percent(float(b["mean"]), float(a["mean"]))
                pairs.append({
                    "repeat": repeat, "case": case, "shape": shape,
                    "before_statistics_ns": b, "after_statistics_ns": a,
                    "latency_change_percent": changes,
                    "throughput_change_percent": throughput,
                    "adverse_over_5_percent": any(changes[metric] > 5.0 for metric in STATISTICS)
                    or throughput < -5.0,
                })
    for stage in ("before", "after"):
        first = captures[stage, "r1"]["rows"]
        second = captures[stage, "r2"]["rows"]
        for shape in SHAPES:
            for case in CASES:
                key = (case, shape)
                left = first[key]["elapsed"]["statistics"]
                right = second[key]["elapsed"]["statistics"]
                changes = {metric: percent(float(right[metric]), float(left[metric])) for metric in STATISTICS}
                exceeds = {metric: abs(changes[metric]) > DRIFT_LIMITS[metric] for metric in STATISTICS}
                drift.append({
                    "stage": stage, "case": case, "shape": shape,
                    "repeat_change_percent": changes,
                    "drift_limits_percent": dict(DRIFT_LIMITS),
                    "exceeds_drift_ceiling": exceeds,
                    "any_exceeds_drift_ceiling": any(exceeds.values()),
                })
    return pairs, drift


def control_drift(capture_set: dict[tuple[str, str], dict[str, Any]],
                  stage: str = "before") -> list[dict[str, Any]]:
    """Report same-build repeat drift without fabricating candidate deltas."""
    first = capture_set[stage, "r1"]["rows"]
    second = capture_set[stage, "r2"]["rows"]
    result = []
    for shape in SHAPES:
        for case in CASES:
            key = (case, shape)
            left = first[key]["elapsed"]["statistics"]
            right = second[key]["elapsed"]["statistics"]
            changes = {metric: percent(float(right[metric]), float(left[metric]))
                       for metric in STATISTICS}
            exceeds = {metric: abs(changes[metric]) > DRIFT_LIMITS[metric]
                       for metric in STATISTICS}
            result.append({
                "stage": stage, "case": case, "shape": shape,
                "first_statistics_ns": left, "second_statistics_ns": right,
                "repeat_change_percent": changes,
                "drift_limits_percent": dict(DRIFT_LIMITS),
                "exceeds_drift_ceiling": exceeds,
                "any_exceeds_drift_ceiling": any(exceeds.values()),
                "scope": "same-build control repeat drift; no candidate comparison",
            })
    return result


def control_rss(stage: str = "before", lanes: tuple[str, str] = ("r1", "r2")) -> dict[str, Any]:
    first = rss(stage_dir(stage) / f"{lanes[0]}.log")
    second = rss(stage_dir(stage) / f"{lanes[1]}.log")
    change = percent(float(second), float(first))
    return {
        "stage": stage, "first_lane": lanes[0], "second_lane": lanes[1],
        "scope": "same-build whole-child RSS; no candidate comparison",
        "first_peak_rss_kib": first, "second_peak_rss_kib": second,
        "change_percent": change, "adverse_over_5_percent": change > 5.0,
    }


def rss(path: Path) -> int:
    text = path.read_text(encoding="utf-8", errors="strict")
    values = RSS_RE.findall(text)
    require(len(values) == 1, f"{path} must contain one whole-child RSS line")
    result = int(values[0])
    require(result > 0, f"{path} reports non-positive RSS")
    return result


def native_rss(capture_set: dict[tuple[str, str], dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for repeat in REPEATS:
        before = rss(stage_dir("before") / f"{repeat}.log")
        after = rss(stage_dir("after") / f"{repeat}.log")
        change = percent(float(after), float(before))
        result.append({
            "repeat": repeat,
            "scope": "whole child: all 12 rows, fixture generation, setup, oracles and drop",
            "before_peak_rss_kib": before, "after_peak_rss_kib": after,
            "change_percent": change, "adverse_over_5_percent": change > 5.0,
        })
    del capture_set
    return result


def parse_time(value: Any, context: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{context} has no start time")
    try:
        return _datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{context} has an invalid start time: {error}")


def verify_capture_order(receipts: dict[tuple[str, str], dict[str, Any]], context: str = "native") -> None:
    starts = [parse_time(receipts[stage, repeat].get("started_utc"), f"{context} {stage}/{repeat}")
              for stage, repeat in ABBA]
    require(starts == sorted(starts) and len(set(starts)) == len(starts),
            f"{context} captures are not serial ABBA")


def verify_pilot_precedes_candidate_full(receipts: dict[tuple[str, str], dict[str, Any]],
                                         context: str = "native") -> None:
    pilot = parse_time(receipts["after", "pilot"].get("started_utc"),
                       f"{context} after/pilot")
    for repeat in REPEATS:
        full = parse_time(receipts["after", repeat].get("started_utc"),
                          f"{context} after/{repeat}")
        require(pilot < full, f"{context} candidate pilot does not precede after/{repeat}")


def verify_negative_vectors(capture: dict[str, Any], samples: int) -> dict[str, bool]:
    row = capture["rows"][("xlsx_one_percent_commit_save", "dense-wide")]["row"]
    short = copy.deepcopy(row)
    short["elapsed_ns"]["samples"].pop()
    short_rejected = False
    try:
        _validate_elapsed(short, samples, "negative short vector")
    except Exception:
        short_rejected = True
    require(short_rejected, "negative short elapsed vector was accepted")
    corrupt = copy.deepcopy(row)
    corrupt["elapsed_ns"]["samples"][0] += 1
    corrupt_rejected = False
    try:
        _validate_elapsed(corrupt, samples, "negative corrupt vector")
    except Exception:
        corrupt_rejected = True
    require(corrupt_rejected, "negative corrupt elapsed vector was accepted")
    op_short = copy.deepcopy(row)
    op_short["operation_metrics"]["sample_indices"].pop()
    op_rejected = False
    try:
        verify_operation_metrics(op_short, samples, "negative operation vector", False,
                                 row["operation_metrics"]["sample_indices"])
    except Exception:
        op_rejected = True
    require(op_rejected, "negative short operation vector was accepted")
    return {
        "short_vector_rejected": short_rejected,
        "corrupt_vector_rejected": corrupt_rejected,
        "short_operation_vector_rejected": op_rejected,
    }


def _vector_summary(values: list[int]) -> dict[str, float]:
    ordered = sorted(values)
    middle = (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2
    return {"min": float(min(values)), "p50": middle, "max": float(max(values)),
            "mean": sum(values) / len(values)}


def allocation_review(allocator: dict[tuple[str, str], dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []
    for repeat in REPEATS:
        before = allocator["before", repeat]["rows"]
        after = allocator["after", repeat]["rows"]
        for shape in SHAPES:
            for case in CASES:
                key = (case, shape)
                b = before[key]["allocation"]
                a = after[key]["allocation"]
                require(b["status"] == a["status"] == "measured", "allocator review has unavailable data")
                absolute = {}
                for field in ("region_peak_live_bytes", "live_bytes_before", "live_bytes_after",
                              "peak_live_bytes_before", "peak_live_bytes_after"):
                    absolute[field] = {
                        "before": _vector_summary(b["absolute"][field]),
                        "after": _vector_summary(a["absolute"][field]),
                    }
                bi = _vector_summary(b["incremental_region_peak"])
                ai = _vector_summary(a["incremental_region_peak"])
                absolute["incremental_region_peak_bytes"] = {"before": bi, "after": ai}
                flags = {}
                for name in ("region_peak_live_bytes", "incremental_region_peak_bytes"):
                    left = absolute[name]["before"]["p50"]
                    right = absolute[name]["after"]["p50"]
                    flags[name] = percent(right, left) > 5.0 if left > 0 else None
                rows.append({"repeat": repeat, "case": case, "shape": shape,
                             "memory": absolute, "adverse_over_5_percent": flags,
                             "scope": "absolute region peak plus region peak minus live_bytes_before"})
    return rows


def pilot_comparison(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    pairs = []
    for shape in SHAPES:
        for case in CASES:
            key = (case, shape)
            verify_row_identity(after["rows"][key]["row"], before["rows"][key]["row"],
                                f"matched-pilot/{case}/{shape}")
            b = before["rows"][key]["elapsed"]["statistics"]
            a = after["rows"][key]["elapsed"]["statistics"]
            changes = {metric: percent(float(a[metric]), float(b[metric])) for metric in STATISTICS}
            pairs.append({"case": case, "shape": shape,
                          "before_statistics_ns": b, "candidate_statistics_ns": a,
                          "latency_change_percent": changes,
                          "throughput_change_percent": percent(float(b["mean"]), float(a["mean"])),
                          "scope": "candidate 20/2 pilot against one matched control pilot; admission evidence only"})
    return {
        "scope": "matched candidate/control pilot; historical control full captures remain separate",
        "pairs": pairs,
    }


def _guard_plan() -> dict[str, Any]:
    plan = load(GUARD_PLAN)
    require(isinstance(plan, dict), "guard-plan.json must be an object")
    require(isinstance(plan.get("frozen_utc"), str), "guard plan has no frozen timestamp")
    parse_time(plan["frozen_utc"], "guard plan frozen_utc")
    if "final_probe_freeze_utc" in plan:
        final_freeze = parse_time(plan["final_probe_freeze_utc"],
                                  "guard plan final_probe_freeze_utc")
        require(final_freeze >= parse_time(plan["frozen_utc"], "guard plan frozen_utc"),
                "guard plan final freeze precedes its initial freeze")
    require(plan.get("shapes") == list(SHAPES), "guard plan shape set differs")
    require(plan.get("scenarios") == list(GUARD_SCENARIOS),
            "guard plan scenario order differs")
    for key in ("scenario_scope", "native", "allocator", "fidelity", "source_binding"):
        require(isinstance(plan.get(key), str) and plan[key],
                f"guard plan omits {key}")
    native_protocol = re.sub(r"\s+", "", plan["native"])
    require("100samplesand3warmups" in native_protocol and "serialABBAbefore/r1,after/r1,after/r2,before/r2" in native_protocol,
            "guard plan does not declare native serial ABBA")
    require("pilot20/2" in native_protocol,
            "guard plan does not declare the candidate pilot")
    allocator_protocol = re.sub(r"\s+", "", plan["allocator"])
    require("10samplesand1warmup" in allocator_protocol,
            "guard plan does not declare the allocator sample protocol")
    binding_text = plan["source_binding"].lower()
    require("source" in binding_text and "manifest" in binding_text
            and "lock" in binding_text
            and ("binary" in binding_text or "executable" in binding_text),
            "guard plan does not declare separate custody bindings")
    return {"sha256": sha(GUARD_PLAN), "frozen_utc": plan["frozen_utc"]}


def _guard_source_manifest() -> dict[str, Any]:
    manifest = load(GUARD_SOURCE_MANIFEST)
    require(isinstance(manifest, dict) and manifest, "guard-source-manifest.json is empty")
    probe_relative = GUARD_PROBE.relative_to(REPO).as_posix()
    expected = {
        f"{probe_relative}/Cargo.toml",
        f"{probe_relative}/Cargo.lock",
        f"{probe_relative}/src/main.rs",
        *GUARD_CANONICAL_SOURCES,
    }
    require(set(manifest) == expected,
            f"guard source manifest files differ: {sorted(manifest)!r}")
    for name, digest in manifest.items():
        check_hash(digest, f"guard source {name}")
        require(sha(repository_file(name, f"guard source {name}")) == digest,
                f"guard source changed: {name}")
    return {"sha256": sha(GUARD_SOURCE_MANIFEST), "files": len(manifest)}


def _guard_lock() -> dict[str, Any]:
    receipt_path = HERE / "before" / "guard-lock-receipt.json"
    receipt = load(receipt_path)
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            "guard lock generation did not exit successfully")
    lock_digest = check_hash(receipt.get("lock_sha256"), "guard lock_sha256")
    require(lock_digest == sha(repository_file(
        f"{GUARD_PROBE.relative_to(REPO).as_posix()}/Cargo.lock", "guard Cargo.lock"
    )),
            "guard lock receipt is not bound to guard-probe/Cargo.lock")
    log_path = HERE / "before" / "guard-lock.log"
    log_digest = check_hash(receipt.get("log_sha256"), "guard lock log_sha256")
    require(log_digest == sha(log_path), "guard lock log hash is not bound")
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            "guard lock command is not an argv list")
    require(command[:2] == ["cargo", "generate-lockfile"] and "--offline" in command,
            "guard lock command is not offline cargo generate-lockfile")
    manifest = option(command, "--manifest-path", "guard lock")
    recorded_path(manifest,
                  f"{GUARD_PROBE.relative_to(REPO).as_posix()}/Cargo.toml",
                  "guard lock manifest")
    require(command.count("--locked") == 0, "guard lock generation cannot use --locked")
    return {"sha256": lock_digest, "log_sha256": log_digest}


def _guard_build(stage: str, allocator: bool, main_build: dict[str, Any],
                 guard_manifest_sha256: str) -> dict[str, Any]:
    directory = stage_dir(stage)
    name = "guard-build-allocator" if allocator else "guard-build"
    receipt_path = directory / f"{name}-receipt.json"
    receipt = load(receipt_path)
    context = f"{stage}/{name}"
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{context} did not exit successfully")
    require(receipt.get("source_unchanged") is True,
            f"{context} source changed during build")
    binary = check_hash(receipt.get("binary_sha256"), f"{context}.binary_sha256")
    source_digest = check_hash(receipt.get("source_manifest_sha256"),
                               f"{context}.source_manifest_sha256")
    require(source_digest == main_build["source_manifest_sha256"],
            f"{context} production source role differs from main build")
    require(receipt.get("guard_manifest_sha256") == guard_manifest_sha256,
            f"{context} guard source manifest is not bound")
    elapsed = finite(receipt.get("elapsed_seconds"), f"{context}.elapsed_seconds")
    require(elapsed > 0, f"{context} elapsed time is not positive")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and f"{name}.log" in artifacts,
            f"{context} does not bind its build log")
    log_path = directory / f"{name}.log"
    log_digest = check_hash(artifacts[f"{name}.log"], f"{context} artifact {name}.log")
    require(log_digest == sha(log_path), f"{context} log hash is not bound")
    if "log_sha256" in receipt:
        require(check_hash(receipt["log_sha256"], f"{context}.log_sha256") == log_digest,
                f"{context}.log_sha256 disagrees with its artifact")
    for relative, digest in artifacts.items():
        path = safe_relative(relative, f"{context} artifact path")
        artifact = bundle_file(directory / path, f"{context} artifact {relative}")
        check_hash(digest, f"{context} artifact {relative}")
        require(sha(artifact) == digest, f"{context} artifact changed: {relative}")
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{context}.command is not an argv list")
    expected_binary = GUARD_ALLOCATOR_BINARY if allocator else GUARD_NORMAL_BINARY
    require(command[:6] == ["/usr/bin/time", "-v", "cargo", "build", "--release", "--locked"],
            f"{context} command is not a locked release build")
    require("--offline" in command, f"{context} command is not offline")
    manifest = option(command, "--manifest-path", context)
    recorded_path(manifest,
                  f"{GUARD_PROBE.relative_to(REPO).as_posix()}/Cargo.toml",
                  f"{context} manifest")
    require(option(command, "--bin", context) == GUARD_NORMAL_BINARY,
            f"{context} command does not select the guard binary")
    if allocator:
        require(command.count("--features") == 1
                and option(command, "--features", context) == "allocator-metrics",
                f"{context} command does not select allocator-metrics")
    else:
        require("--features" not in command, f"{context} normal guard build is instrumented")
    return {
        "binary_sha256": binary,
        "source_manifest_sha256": source_digest,
        "guard_manifest_sha256": guard_manifest_sha256,
        "log_sha256": log_digest,
        "allocator": allocator,
        "binary_name": expected_binary,
    }


def _guard_command(command: Any, stage: str, name: str, samples: int, warmups: int,
                   allocator: bool) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{name}.command is not an argv list")
    require(tuple(command[:5]) == REPORT_PREFIX, f"{name} is not pinned to CPU 2")
    binary = (GUARD_ALLOCATOR_COMMAND_BINARY if allocator
              else GUARD_NORMAL_COMMAND_BINARY)
    binary_paths = [Path(item) for item in command if Path(item).name == binary]
    require(len(binary_paths) == 1,
            f"{name} command does not identify one {binary}")
    recorded_path(binary_paths[0].as_posix(), f"{stage}/{binary}",
                  f"{name} binary")
    require(option(command, "--samples", name) == str(samples),
            f"{name} sample count differs")
    require(option(command, "--warmup", name) == str(warmups),
            f"{name} warmup count differs")
    require(option(command, "--shape", name) == ",".join(SHAPES),
            f"{name} shape selection differs")
    report = option(command, "--json", name)
    require(Path(report).name == f"{name}-report.json",
            f"{name} report path is not bound")
    recorded_path(
        report,
        f"{stage_dir(stage).relative_to(REPO).as_posix()}/{name}-report.json",
        f"{name} report",
    )
    scenario_positions = [index for index, value in enumerate(command) if value == "--scenario"]
    require(len(scenario_positions) <= 1, f"{name} repeats --scenario")
    if scenario_positions:
        require(option(command, "--scenario", name) == ",".join(GUARD_SCENARIOS),
                f"{name} scenario selection differs")
    require("--corpus-manifest" not in command,
            f"{name} unexpectedly uses the main harness corpus catalog")
    return command


def _guard_allocation_sample(sample: Any, samples_context: str, allocator: bool) -> dict[str, Any]:
    require(isinstance(sample, dict), f"{samples_context} is not an allocation sample")
    expected_keys = {"status", "scope"}
    if allocator:
        expected_keys |= set(ALLOC_FIELDS)
        require(set(sample) == expected_keys,
                f"{samples_context} measured allocation sample schema differs")
        require(sample.get("status") == "measured",
                f"{samples_context} allocator sample is not measured")
        require(sample.get("scope") == "operation_global_system_allocator",
                f"{samples_context} allocator sample scope differs")
        values: dict[str, int] = {}
        for field in ALLOC_FIELDS:
            value = sample.get(field)
            require(isinstance(value, int) and not isinstance(value, bool)
                    and 0 <= value <= (1 << 64) - 1,
                    f"{samples_context}.{field} is invalid")
            values[field] = value
        require(values["region_peak_live_bytes"] >= values["live_bytes_before"]
                and values["region_peak_live_bytes"] >= values["live_bytes_after"]
                and values["region_peak_live_bytes"] <= values["peak_live_bytes_after"],
                f"{samples_context} region peak is outside live/high-water bounds")
        require(values["peak_live_bytes_after"] >= values["peak_live_bytes_before"],
                f"{samples_context} peak live bytes decreased")
        return {"status": "measured", "absolute": values,
                "incremental_region_peak": values["region_peak_live_bytes"]
                - values["live_bytes_before"]}
    require(set(sample) == expected_keys,
            f"{samples_context} unavailable allocation sample publishes values")
    require(sample.get("status") == "unavailable"
            and sample.get("scope") == "operation_global_system_allocator",
            f"{samples_context} unavailable allocation sample identity differs")
    return {"status": "unavailable"}


def _guard_percentile(values: list[int], percentile: int) -> int:
    return sorted(values)[(len(values) - 1) * percentile // 100]


def _guard_scenario(row: Any, shape: str, samples: int, warmups: int,
                    context: str, allocator: bool) -> dict[str, Any]:
    require(isinstance(row, dict), f"{context} is not an object")
    required = {"scenario", "warm_store", "changed", "update_count",
                "warmup_iterations", "sample_count", "sample_indices", "elapsed_ns",
                "stats", "allocation_samples", "oracle"}
    require(set(row) == required, f"{context} scenario schema differs")
    scenario = row.get("scenario")
    require(scenario in GUARD_SCENARIOS, f"{context}.scenario is unknown")
    expected_one_cell = scenario.endswith("one-cell")
    expected_percent = scenario.endswith("one-percent")
    expected_first_read = scenario == "cold-first-cell-read"
    expected_changed = scenario.startswith("warm-changed")
    expected_warm = scenario.startswith("warm-")
    require(row.get("warm_store") is expected_warm, f"{context}.warm_store differs")
    require(row.get("changed") is expected_changed, f"{context}.changed differs")
    expected_updates = 0 if expected_first_read else (1 if expected_one_cell else None)
    if expected_percent:
        expected_updates = (2 * GUARD_SHAPE_SIDES[shape] * GUARD_SHAPE_SIDES[shape] + 99) // 100
    require(isinstance(row.get("update_count"), int)
            and not isinstance(row["update_count"], bool) and row["update_count"] >= 0,
            f"{context}.update_count is invalid")
    require(row["update_count"] == expected_updates, f"{context}.update_count differs")
    require(isinstance(row.get("warmup_iterations"), int)
            and not isinstance(row["warmup_iterations"], bool)
            and row["warmup_iterations"] == warmups,
            f"{context}.warmup_iterations differs")
    require(isinstance(row.get("sample_count"), int)
            and not isinstance(row["sample_count"], bool)
            and row["sample_count"] == samples,
            f"{context}.sample_count differs")
    indices = row.get("sample_indices")
    require(indices == list(range(samples)), f"{context}.sample_indices are not chronological")
    elapsed = row.get("elapsed_ns")
    require(isinstance(elapsed, list) and len(elapsed) == samples,
            f"{context}.elapsed_ns has wrong cardinality")
    require(all(isinstance(item, int) and not isinstance(item, bool)
                and 0 < item <= (1 << 64) - 1 for item in elapsed),
            f"{context}.elapsed_ns contains an invalid value")
    stats = row.get("stats")
    required_stats = {"min_ns", "p50_ns", "p95_ns", "p99_ns", "max_ns", "mean_ns"}
    require(isinstance(stats, dict) and set(stats) == required_stats,
            f"{context}.stats schema differs")
    ordered = sorted(elapsed)
    expected_stats: dict[str, Any] = {
        "min_ns": ordered[0], "p50_ns": _guard_percentile(elapsed, 50),
        "p95_ns": _guard_percentile(elapsed, 95), "p99_ns": _guard_percentile(elapsed, 99),
        "max_ns": ordered[-1], "mean_ns": float(sum(elapsed)) / float(samples),
    }
    for key, expected in expected_stats.items():
        if key == "mean_ns":
            require(isinstance(stats[key], (int, float)) and not isinstance(stats[key], bool)
                    and math.isfinite(float(stats[key])) and float(stats[key]) == expected,
                    f"{context}.stats.{key} differs")
        else:
            require(isinstance(stats[key], int) and not isinstance(stats[key], bool)
                    and stats[key] == expected, f"{context}.stats.{key} differs")
    allocation_samples = row.get("allocation_samples")
    require(isinstance(allocation_samples, list) and len(allocation_samples) == samples,
            f"{context}.allocation_samples cardinality differs")
    allocation = [_guard_allocation_sample(sample, f"{context}.allocation_samples[{index}]",
                                           allocator)
                  for index, sample in enumerate(allocation_samples)]
    oracle = row.get("oracle")
    oracle_keys = {"iterations_checked", "patch_empty", "source_bytes_equal",
                   "changed_readback", "first_cell_readback"}
    require(isinstance(oracle, dict) and set(oracle) == oracle_keys,
            f"{context}.oracle schema differs")
    require(isinstance(oracle.get("iterations_checked"), int)
            and not isinstance(oracle["iterations_checked"], bool)
            and oracle["iterations_checked"] == samples + warmups,
            f"{context}.oracle iteration count differs")
    for key in oracle_keys - {"iterations_checked"}:
        require(oracle[key] is None or isinstance(oracle[key], bool),
                f"{context}.oracle.{key} is not bool/null")
    if expected_first_read:
        expected_oracle = {"patch_empty": None, "source_bytes_equal": None,
                           "changed_readback": None, "first_cell_readback": True}
    elif expected_changed:
        expected_oracle = {"patch_empty": False, "source_bytes_equal": None,
                           "changed_readback": True, "first_cell_readback": None}
    else:
        expected_oracle = {"patch_empty": True, "source_bytes_equal": True,
                           "changed_readback": None, "first_cell_readback": None}
    for key, expected in expected_oracle.items():
        require(oracle[key] is expected, f"{context}.oracle.{key} does not prove expected result")
    return {"scenario": scenario, "warm_store": expected_warm, "changed": expected_changed,
            "update_count": row["update_count"], "elapsed_ns": elapsed,
            "stats": stats, "allocation": allocation,
            "identity": {"scenario": scenario, "warm_store": expected_warm,
                         "changed": expected_changed, "update_count": row["update_count"]}}


def _guard_report(report: Any, samples: int, warmups: int, allocator: bool,
                  context: str) -> dict[str, Any]:
    require(isinstance(report, dict), f"{context} report is not an object")
    required = {"schema_version", "probe", "timer_scope", "allocation", "samples",
                "warmups", "shapes"}
    require(set(report) == required, f"{context} report schema differs")
    require(report.get("schema_version") == 1, f"{context} schema version differs")
    require(report.get("probe") == GUARD_PROBE_NAME, f"{context} probe identity differs")
    require(report.get("timer_scope") == GUARD_TIMER_SCOPE,
            f"{context} timer scope differs")
    require(report.get("samples") == samples and report.get("warmups") == warmups,
            f"{context} sample protocol differs")
    allocation = report.get("allocation")
    require(isinstance(allocation, dict)
            and set(allocation) == {"binary", "allocator", "instrumentation", "counter_revision"},
            f"{context} allocation identity schema differs")
    expected_binary = GUARD_ALLOCATOR_BINARY if allocator else GUARD_NORMAL_BINARY
    expected_allocator = ("CountingSystemAllocator(std::alloc::System)" if allocator
                          else "Rust system allocator")
    expected_instrumentation = ("system_allocator_operation_scoped" if allocator else "none")
    require(allocation.get("binary") == expected_binary,
            f"{context} allocation binary identity differs")
    require(allocation.get("allocator") == expected_allocator,
            f"{context} allocator identity differs")
    require(allocation.get("instrumentation") == expected_instrumentation,
            f"{context} instrumentation identity differs")
    require(allocation.get("counter_revision") ==
            ("serialized_region_peak_v3" if allocator else None),
            f"{context} counter revision differs")
    shapes = report.get("shapes")
    require(isinstance(shapes, list) and len(shapes) == len(SHAPES)
            and all(isinstance(item, dict) for item in shapes),
            f"{context} shape list is malformed")
    require([item.get("shape") for item in shapes] == list(SHAPES),
            f"{context} shape order differs")
    parsed_shapes: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(shapes):
        shape_context = f"{context}.shapes[{index}]"
        require(isinstance(item, dict), f"{shape_context} is not an object")
        required_shape = {"shape", "sheet_count", "rows_per_sheet", "columns_per_sheet",
                          "cells", "one_percent_update_count", "corpus_bytes",
                          "corpus_sha256", "scenarios"}
        require(set(item) == required_shape, f"{shape_context} schema differs")
        shape = item["shape"]
        require(shape in GUARD_SHAPE_SIDES, f"{shape_context}.shape is unknown")
        side = GUARD_SHAPE_SIDES[shape]
        cells = 2 * side * side
        require(item.get("sheet_count") == 2 and item.get("rows_per_sheet") == side
                and item.get("columns_per_sheet") == side and item.get("cells") == cells,
                f"{shape_context} dimensions differ")
        require(item.get("one_percent_update_count") == (cells + 99) // 100,
                f"{shape_context} one-percent update count differs")
        require(isinstance(item.get("corpus_bytes"), int) and item["corpus_bytes"] > 0,
                f"{shape_context}.corpus_bytes is invalid")
        check_hash(item.get("corpus_sha256"), f"{shape_context}.corpus_sha256")
        scenarios = item.get("scenarios")
        require(isinstance(scenarios, list) and len(scenarios) == len(GUARD_SCENARIOS)
                and all(isinstance(scenario, dict) for scenario in scenarios),
                f"{shape_context} scenario list is malformed")
        require([scenario.get("scenario") for scenario in scenarios] == list(GUARD_SCENARIOS),
                f"{shape_context} scenario order differs")
        parsed_scenarios: dict[str, dict[str, Any]] = {}
        for scenario_index, scenario in enumerate(scenarios):
            parsed_scenarios[scenario["scenario"]] = _guard_scenario(
                scenario, shape, samples, warmups,
                f"{shape_context}.scenarios[{scenario_index}]", allocator
            )
        parsed_shapes[shape] = {
            "shape": shape, "sheet_count": item["sheet_count"],
            "rows_per_sheet": item["rows_per_sheet"],
            "columns_per_sheet": item["columns_per_sheet"], "cells": item["cells"],
            "one_percent_update_count": item["one_percent_update_count"],
            "corpus_bytes": item["corpus_bytes"], "corpus_sha256": item["corpus_sha256"],
            "scenarios": parsed_scenarios,
        }
    return {"report": report, "shapes": parsed_shapes}


def _guard_capture(stage: str, lane: str, build: dict[str, Any], guard_build: dict[str, Any],
                   guard_manifest_sha256: str, samples: int, warmups: int,
                   allocator: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    name = f"guard-{lane}"
    directory = stage_dir(stage)
    context = f"{stage}/{name}"
    receipt = load(directory / f"{name}-receipt.json")
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{context} did not exit successfully")
    require(receipt.get("source_unchanged") is True,
            f"{context} source changed during capture")
    require(receipt.get("binary_sha256") == guard_build["binary_sha256"],
            f"{context} binary hash differs from guard build")
    check_hash(receipt.get("binary_sha256"), f"{context}.binary_sha256")
    require(receipt.get("source_manifest_sha256") == guard_build["source_manifest_sha256"],
            f"{context} production source manifest differs from guard build")
    require(receipt.get("guard_manifest_sha256") == guard_manifest_sha256,
            f"{context} guard source manifest differs from guard build")
    require(receipt.get("scope") == GUARD_SCOPE, f"{context} scope is not explicit")
    elapsed = finite(receipt.get("elapsed_seconds"), f"{context}.elapsed_seconds")
    require(elapsed > 0, f"{context} elapsed time is not positive")
    report_name = f"{name}-report.json"
    log_name = f"{name}.log"
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and {report_name, log_name} <= set(artifacts),
            f"{context} artifact map is incomplete")
    for relative, digest in artifacts.items():
        path = safe_relative(relative, f"{context} artifact path")
        artifact = bundle_file(directory / path, f"{context} artifact {relative}")
        check_hash(digest, f"{context} artifact {relative}")
        require(sha(artifact) == digest, f"{context} artifact changed: {relative}")
    _guard_command(receipt.get("command"), stage, name, samples, warmups, allocator)
    report = load(directory / report_name)
    parsed = _guard_report(report, samples, warmups, allocator, context)
    return receipt, parsed


def _guard_vector_summary(values: list[int]) -> dict[str, float]:
    require(values, "cannot summarize an empty guard allocation vector")
    ordered = sorted(values)
    return {
        "min": float(ordered[0]),
        "p50": float(ordered[(len(ordered) - 1) // 2]),
        "max": float(ordered[-1]),
        "mean": float(sum(values)) / float(len(values)),
    }


def _guard_memory_review(captures: dict[tuple[str, str], dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for repeat in REPEATS:
        before = captures["before", f"allocator-{repeat}"]["shapes"]
        after = captures["after", f"allocator-{repeat}"]["shapes"]
        for shape in SHAPES:
            for scenario in GUARD_SCENARIOS:
                b_samples = before[shape]["scenarios"][scenario]["allocation"]
                a_samples = after[shape]["scenarios"][scenario]["allocation"]
                require(all(sample["status"] == "measured" for sample in b_samples + a_samples),
                        "guard allocator review has unavailable data")
                memory: dict[str, Any] = {}
                for field in ("region_peak_live_bytes", "live_bytes_before", "live_bytes_after",
                              "peak_live_bytes_before", "peak_live_bytes_after"):
                    memory[field] = {
                        "before": _guard_vector_summary(
                            [sample["absolute"][field] for sample in b_samples]
                        ),
                        "after": _guard_vector_summary(
                            [sample["absolute"][field] for sample in a_samples]
                        ),
                    }
                memory["incremental_region_peak_bytes"] = {
                    "before": _guard_vector_summary(
                        [sample["incremental_region_peak"] for sample in b_samples]
                    ),
                    "after": _guard_vector_summary(
                        [sample["incremental_region_peak"] for sample in a_samples]
                    ),
                }
                flags = {}
                for field in ("region_peak_live_bytes", "incremental_region_peak_bytes"):
                    left = memory[field]["before"]["p50"]
                    right = memory[field]["after"]["p50"]
                    flags[field] = percent(right, left) > 5.0 if left > 0 else None
                result.append({"repeat": repeat, "shape": shape, "scenario": scenario,
                               "memory": memory, "adverse_over_5_percent": flags,
                               "scope": "absolute region peak plus region peak minus live_bytes_before"})
    return result


def _guard_control_drift(captures: dict[tuple[str, str], dict[str, Any]]) -> list[dict[str, Any]]:
    first = captures["before", "r1"]["shapes"]
    second = captures["before", "r2"]["shapes"]
    result = []
    for shape in SHAPES:
        for scenario in GUARD_SCENARIOS:
            left = first[shape]["scenarios"][scenario]["stats"]
            right = second[shape]["scenarios"][scenario]["stats"]
            changes = {
                metric: percent(float(right[f"{metric}_ns"]), float(left[f"{metric}_ns"]))
                for metric in ("p50", "mean", "p95", "p99")
            }
            exceeds = {metric: abs(changes[metric]) > DRIFT_LIMITS[metric]
                       for metric in changes}
            result.append({
                "stage": "before", "shape": shape, "scenario": scenario,
                "first_statistics_ns": left, "second_statistics_ns": right,
                "repeat_change_percent": changes,
                "drift_limits_percent": dict(DRIFT_LIMITS),
                "exceeds_drift_ceiling": exceeds,
                "any_exceeds_drift_ceiling": any(exceeds.values()),
                "scope": "same-build guard control repeat drift; no candidate comparison",
            })
    return result


def _guard_control_rss() -> dict[str, Any]:
    first = rss(stage_dir("before") / "guard-r1.log")
    second = rss(stage_dir("before") / "guard-r2.log")
    change = percent(float(second), float(first))
    return {
        "stage": "before", "first_lane": "r1", "second_lane": "r2",
        "scope": "same-build guard whole-child RSS; no candidate comparison",
        "first_peak_rss_kib": first, "second_peak_rss_kib": second,
        "change_percent": change, "adverse_over_5_percent": change > 5.0,
    }


def verify_guard(builds: dict[str, Any], kept: bool) -> dict[str, Any]:
    plan = _guard_plan()
    source = _guard_source_manifest()
    lock = _guard_lock()
    guard_builds = {
        (stage, allocator): _guard_build(stage, allocator, builds[stage], source["sha256"])
        for stage in ("before", "after")
        for allocator in (False, True)
    }
    for allocator in (False, True):
        require(guard_builds["before", allocator]["source_manifest_sha256"]
                != guard_builds["after", allocator]["source_manifest_sha256"],
                "guard control and candidate builds share one production source epoch")
    normal: dict[tuple[str, str], dict[str, Any]] = {}
    receipts: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in ("before", "after"):
        for lane, samples, warmups in (("preflight", PREFLIGHT_SAMPLES, PREFLIGHT_WARMUPS),
                                        ("pilot", PILOT_SAMPLES, PILOT_WARMUPS)):
            receipt, capture = _guard_capture(
                stage, lane, builds[stage], guard_builds[stage, False], source["sha256"],
                samples, warmups, False
            )
            normal[stage, lane] = capture
            receipts[stage, lane] = receipt
    full_lanes = (("before", "r1"), ("before", "r2"))
    if kept:
        full_lanes = (("before", "r1"), ("after", "r1"),
                      ("after", "r2"), ("before", "r2"))
    for stage, repeat in full_lanes:
        receipt, capture = _guard_capture(
            stage, repeat, builds[stage], guard_builds[stage, False], source["sha256"],
            GUARD_SAMPLES, GUARD_WARMUPS, False
        )
        normal[stage, repeat] = capture
        receipts[stage, repeat] = receipt
    if kept:
        verify_capture_order({key: receipts[key] for key in ABBA}, "guard native")
        verify_pilot_precedes_candidate_full(receipts, "guard native")
    allocator: dict[tuple[str, str], dict[str, Any]] = {}
    allocator_receipts: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in ("before", "after"):
        for repeat in REPEATS:
            lane = f"allocator-{repeat}"
            receipt, capture = _guard_capture(
                stage, lane, builds[stage], guard_builds[stage, True], source["sha256"],
                ALLOC_SAMPLES, ALLOC_WARMUPS, True
            )
            allocator[stage, lane] = capture
            allocator_receipts[stage, lane] = receipt
    all_captures = {**normal, **allocator}
    reference: dict[str, Any] | None = None
    for key, capture in all_captures.items():
        del key
        if reference is None:
            reference = capture
            continue
        for shape in SHAPES:
            expected = reference["shapes"][shape]
            current = capture["shapes"][shape]
            for field in ("shape", "sheet_count", "rows_per_sheet", "columns_per_sheet",
                          "cells", "one_percent_update_count", "corpus_bytes", "corpus_sha256"):
                require(current[field] == expected[field],
                        f"guard corpus identity differs for {shape}.{field}")
            for scenario in GUARD_SCENARIOS:
                left = expected["scenarios"][scenario]["identity"]
                right = current["scenarios"][scenario]["identity"]
                require(right == left, f"guard scenario identity differs: {shape}/{scenario}")
    return {
        "plan_sha256": plan["sha256"], "source_manifest_sha256": source["sha256"],
        "source_manifest_files": source["files"], "lock_sha256": lock["sha256"],
        "guard_builds": {
            f"{stage}/{'allocator' if allocator else 'normal'}": value
            for (stage, allocator), value in guard_builds.items()
        },
        "source_roles_bound_to_main_builds": True,
        "lock_bound_to_probe": True,
        "binary_roles_bound": True,
        "native": {
            "control_and_candidate": kept,
            "samples_per_scenario": GUARD_SAMPLES,
            "warmups_per_scenario": GUARD_WARMUPS,
            "formal_after_captures": "required" if kept else "not_required_after_candidate_rejection",
            "abba_order": [f"{stage}/{repeat}" for stage, repeat in ABBA] if kept else None,
            "control_same_build_drift": _guard_control_drift(normal),
            "control_before_only_rss": _guard_control_rss(),
        },
        "pilot": {"samples": PILOT_SAMPLES, "warmups": PILOT_WARMUPS,
                  "both_roles": True, "matched_corpus_identity": True},
        "allocator": {"samples": ALLOC_SAMPLES, "warmups": ALLOC_WARMUPS,
                      "both_roles": True, "two_repeats_each_role": True,
                      "instrumented_elapsed_excluded": True,
                      "memory_review": _guard_memory_review(allocator)},
        "scenario_order": list(GUARD_SCENARIOS),
        "shapes": list(SHAPES),
        "no_op_and_changed_oracles_verified": True,
        "first_cell_readback_oracle_verified": True,
        "reports": {
            **{f"{stage}/{lane}": len(capture["shapes"]) * len(GUARD_SCENARIOS)
               for (stage, lane), capture in normal.items()},
            **{f"{stage}/{lane}": len(capture["shapes"]) * len(GUARD_SCENARIOS)
               for (stage, lane), capture in allocator.items()},
        },
    }


def _capture_if_present(stage: str, lane: str, build: dict[str, Any], samples: int,
                        warmup: int, allocator: bool = False, profile: bool = False) -> tuple[dict[str, Any], dict[str, Any]]:
    return verify_capture(stage, lane, build, samples, warmup, allocator, profile)


def verify() -> dict[str, Any]:
    decision, kept = load_decision()
    custody = verify_sources(decision, kept)
    builds = {
        "before": verify_build("before"),
        "after": verify_build("after"),
        "before_allocator": verify_build("before", allocator=True),
        "after_allocator": verify_build("after", allocator=True),
    }
    normal: dict[tuple[str, str], dict[str, Any]] = {}
    receipts: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in ("before", "after"):
        receipt, capture = verify_capture(stage, "preflight", builds[stage],
                                          PREFLIGHT_SAMPLES, PREFLIGHT_WARMUPS)
        normal[stage, "preflight"] = capture
        receipts[stage, "preflight"] = receipt
        receipt, capture = verify_capture(stage, "pilot", builds[stage],
                                          PILOT_SAMPLES, PILOT_WARMUPS)
        normal[stage, "pilot"] = capture
        receipts[stage, "pilot"] = receipt
    # Control full native observations are retained for every outcome.
    for stage, repeat in (("before", "r1"), ("before", "r2")):
        receipt, capture = verify_capture(stage, repeat, builds[stage],
                                          NATIVE_SAMPLES, NATIVE_WARMUPS)
        normal[stage, repeat] = capture
        receipts[stage, repeat] = receipt
    if kept:
        for stage, repeat in (("after", "r1"), ("after", "r2")):
            receipt, capture = verify_capture(stage, repeat, builds[stage],
                                              NATIVE_SAMPLES, NATIVE_WARMUPS)
            normal[stage, repeat] = capture
            receipts[stage, repeat] = receipt
        verify_capture_order({key: receipts[key] for key in ABBA})
        verify_pilot_precedes_candidate_full(receipts)
    # Candidate pilot is always required before admission; this comparison is
    # kept separate from historical control and from the formal native delta.
    pilot = pilot_comparison(normal["before", "pilot"], normal["after", "pilot"])

    allocator: dict[tuple[str, str], dict[str, Any]] = {}
    allocator_receipts: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in ("before", "after"):
        build = builds[f"{stage}_allocator"]
        for repeat in REPEATS:
            receipt, capture = verify_capture(stage, "r" + repeat[-1], build,
                                              ALLOC_SAMPLES, ALLOC_WARMUPS, allocator=True)
            allocator[stage, repeat] = capture
            allocator_receipts[stage, repeat] = receipt
            normal_reference = normal[stage, "pilot"]
            for key, row in capture["rows"].items():
                reference = normal_reference["rows"][key]["row"]
                verify_row_identity(row["row"], reference,
                                    f"allocator/{stage}/{repeat}/{key[0]}/{key[1]}")

    profiles: dict[str, dict[str, Any]] = {}
    profile_receipts: dict[str, dict[str, Any]] = {}
    for stage in ("before", "after"):
        receipt, profile = verify_capture(stage, "profile", builds[stage],
                                          PROFILE_SAMPLES, PROFILE_WARMUPS, profile=True)
        profiles[stage] = profile
        profile_receipts[stage] = receipt
        profile_key = ("xlsx_one_percent_commit_save", "dense-wide")
        verify_row_identity(profile["rows"][profile_key]["row"],
                            normal[stage, "pilot"]["rows"][profile_key]["row"],
                            f"profile/{stage}/dense-wide/xlsx_one_percent_commit_save")

    # Every available normal report has deterministic corpus/sink identity.
    capture_identity(normal, "native-and-pilot")
    guard = verify_guard(builds, kept)
    memory = allocation_review(allocator)
    negative = verify_negative_vectors(normal["before", "r1"], NATIVE_SAMPLES)
    result: dict[str, Any] = {
        "schema": "litchi-0514-verification-v1",
        "decision": {"status": decision.get("status", decision.get("outcome", decision.get("decision"))),
                     "candidate_kept": kept},
        "comparison_scope": "matched 0514 XLSX candidate/control pilot and formal native evidence; guard and allocator/profile timings excluded from native deltas",
        "custody": custody,
        "builds": builds,
        "pilot": pilot,
        "allocator": {
            "both_roles": True, "rows_per_capture": 12,
            "samples_per_row": ALLOC_SAMPLES, "warmups_per_row": ALLOC_WARMUPS,
            "total_samples": 4 * 12 * ALLOC_SAMPLES,
            "instrumented_elapsed_and_rss_excluded": True,
            "normal_allocation_status": "unavailable; no zero inferred",
            "memory_review": memory,
            "absolute_and_incremental_peak_scope": "region_peak_live_bytes and region_peak_live_bytes minus live_bytes_before",
        },
        "profiles": {stage: profiles[stage]["annotations"] for stage in ("before", "after")},
        "guard": guard,
        "negative_vectors": negative,
        "report_catalog_bindings_verified": True,
        "artifact_and_identity_bindings_verified": True,
    }
    if kept:
        pairs, drift = native_comparisons({key: normal[key] for key in normal if key[1] in REPEATS})
        result["native"] = {
            "rows_per_capture": 12, "cases": list(CASES), "shapes": list(SHAPES),
            "preflight_samples": PREFLIGHT_SAMPLES, "preflight_warmups": PREFLIGHT_WARMUPS,
            "pilot_samples": PILOT_SAMPLES, "pilot_warmups": PILOT_WARMUPS,
            "samples_per_capture": NATIVE_SAMPLES, "warmups_per_capture": NATIVE_WARMUPS,
            "total_elapsed_samples": 4 * 12 * NATIVE_SAMPLES,
            "abba_order": [f"{stage}/{repeat}" for stage, repeat in ABBA],
            "pairs": pairs, "same_role_drift": drift,
            "whole_child_rss": native_rss({key: normal[key] for key in ABBA}),
            "adverse_flags_are_observations": True,
            "corpus_sink_equivalence_verified": True,
        }
    else:
        result["native"] = {
            "formal_after_captures": "not_required_after_candidate_rejection",
            "control_rows_per_capture": 12,
            "control_samples_per_capture": NATIVE_SAMPLES,
            "control_same_build_drift": control_drift(normal),
            "control_before_only_rss": control_rss(),
            "candidate_full_native_comparison": None,
            "scope": "rejection branch preserves control context and matched pilot evidence without inventing unavailable after-native data",
        }
    return result


def main() -> int:
    try:
        print(json.dumps(verify(), indent=2, sort_keys=True))
    except VerificationError as error:
        print(f"verification failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
