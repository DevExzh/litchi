#!/usr/bin/env python3
"""Fail-closed, read-only custody verifier for the 0553 XLSX campaign.

The capture drivers and analyzers own measurement and report generation.  This
module checks their frozen inputs, source manifests, receipt matrices, report
replay, conditional profile rule, quality custody, and final disposition.  It
never builds, captures, changes the checkout, removes the owned target, or
writes an evidence artifact.  Missing later evidence is reported as
``incomplete``; it is never represented by fabricated rows or zeroes.

The verifier intentionally has no 0552 public-test, draft-attempt, or analyzer
amendment path.  The one explicit historical reuse allowed in this bundle is
the separately recorded 0553 baseline-correctness artifact; it proves only
the restored baseline and cannot satisfy candidate custody.
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

sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
CAPTURE = HERE / "capture.py"
GUARDED_CAPTURE = HERE / "guarded_capture.py"
CHECK_ATTEMPT = HERE / "check_attempt.py"
QUALITY = HERE / "quality.py"
METRICS = HERE / "analyze_metrics.py"
GUARDS = HERE / "analyze_guards.py"
FROZEN = HERE / "frozen-inputs.json"
SUPPLEMENTAL = HERE / "supplemental-inputs.json"
ANALYSIS_INPUTS = HERE / "analysis-inputs.json"
ADR = HERE / "adr-manifest.json"
HOST = HERE / "host.json"
QUALITY_PLAN = HERE / "quality-plan.json"
LOCK_BINDING = HERE / "workspace-lock.json"
LOCK_COPY = HERE / "workspace-Cargo.lock"
BASELINE_CORRECTNESS = HERE / "baseline-correctness.json"
CANDIDATE_BINDING = HERE / "candidate-binding.json"
CANDIDATE_CORRECTNESS = HERE / "candidate-correctness.json"
ADVERSE_REVIEW = HERE / "adverse-review.json"
DOCUMENTATION_MANIFEST = HERE / "documentation-manifest.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
PRECLEANUP = HERE / "precleanup-verification.json"
BASELINE = HERE / "baseline"
CANDIDATE = HERE / "candidate"
FINAL = HERE / "final"
TARGET = Path("/home/zhuhe/litchi-goal-0553-target")
SCRATCH_ROOT = TARGET / "retained"

METRICS_REPORT_NAMES = (
    "metrics-analysis.json",
    "main-analysis.json",
    "metrics-comparison.json",
)
GUARDS_REPORT_NAMES = (
    "guards-analysis.json",
    "guard-cap-analysis.json",
)
PROFILE_DECISION_NAMES = (
    "profile-decision.json",
    "profile-analysis.json",
    "profile-gate.json",
    "profile-pilot.json",
)
DECISION_NAMES = ("decision.json", "disposition.json", "admission.json")

STAGES = ("baseline", "candidate")
BINARY_KINDS = ("normal", "alloc", "guard-normal", "guard-alloc", "cap")
SHAPES = ("medium", "dense-sparse", "noncompact", "vendor-extension")
CASES = (
    "xlsx_source_backed_cell_values_one_edit_save",
    "xlsx_source_backed_cell_values_one_percent_edit_save",
    "xlsx_source_backed_managed_cell_values_one_edit_save",
    "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
)
LANES = ("preflight", "native", "alloc")
GUARD_LANES = ("normal", "alloc")
GUARD_SHAPES = ("medium", "dense-sparse")
GUARD_CASES = ("valid", "late-validator", "late-raw")
CAP_SIZES = (1, 2, 160, 164, 256)
PROFILE_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
PROFILE_OWNER = "litchi_xlsx::cell_values::source::MultiSourceEdit::commit"
CPU = 2
PROFILE_DECISION_SCHEMA = "xlsx_0553_profile_decision_v1"
CANDIDATE_BINDING_SCHEMA = "xlsx_0553_candidate_binding_v1"
CANDIDATE_CORRECTNESS_SCHEMA = "xlsx_0553_candidate_correctness_v1"
ADVERSE_REVIEW_SCHEMA = "xlsx_0553_adverse_review_v1"
DOCUMENTATION_MANIFEST_SCHEMA = "xlsx_0553_documentation_manifest_v1"
CLEANUP_SCHEMA = "xlsx_0553_cleanup_v1"
WORKSPACE_LOCK_SHA256 = (
    "9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a"
)
HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
PRIORITY = "OLE2/OOXML first; ODF deferred until that goal completes; iWork excluded"
TARGET_STRING = "/home/zhuhe/litchi-goal-0553-target"
PROCESS_REFERENCE_SCOPE = (
    "Accessible /proc cwd, executable and open file descriptors; "
    "cleanup process ancestors excluded."
)
CLEANUP_SCOPE = (
    "Remove only owned change0553 build/retained-binary target after passing "
    "precleanup and exact binary hash checks."
)
DOCUMENTATION_SCOPE = "Files outside this evidence bundle included in completed change0553"
DOCUMENTATION_REQUIRED = "docs/performance/0553-xlsx-commit-local-compact-proof.md"
REVISION = "8aa0c5baf0616d16c79eba0c6c28dc1716338ad6"
PLAN_SHA256 = "17d810a24912065fde8c71de6109be983f7c5b3bdcea5fd6584f4725c6499ea4"
RUN_SHA256 = "f1398fcc87dc7bccfc05930276e26a4a8a21106480613238eab49310ae1c23f0"
CAPTURE_SHA256 = "21b0153923e87a331906552457a5f77bb444686d4ce9c9c99e3efcb2dbe5d2c1"
ANALYSIS_INPUTS_SHA256 = (
    "2d95091698893cf402ce85003f407e0352a80b3fd70fa5c4c5e3110f8332a38a"
)

CANDIDATE_RESOURCES = {
    "max_proof_logical_heap_bytes": 2097152,
    "cell_slot_bytes": 8,
    "source_byte_cap": 8388608,
    "source_event_cap": 131072,
    "record_initial_capacity": 8,
    "record_growth": (
        "checked geometric doubling; charged before fallible reserve_exact; "
        "oversized reported capacity declines proof"
    ),
    "source_identity": (
        "source allocation pointer/length and parser Store entries allocation "
        "pointer/length; same immutable Snapshot owns both; replaced worksheet "
        "payloads clear proof"
    ),
}
CANDIDATE_LIMITS = {
    "source_bytes": 8388608,
    "events": 131072,
    "logical_proof_bytes": 2097152,
}
REVIEW_SOURCES = {
    "metrics_adverse": "metrics-analysis.json:comparisons.adverse",
    "metrics_drift": "metrics-analysis.json:repeat_drift_over_five_percent",
    "guard_adverse": "guard-cap-analysis.json:comparison.guard.adverse_flags_over_five_percent",
    "guard_drift": "guard-cap-analysis.json:comparison.guard.same_build_drift_over_five_percent",
    "cap_adverse": "guard-cap-analysis.json:comparison.cap.adverse_flags_over_five_percent",
    "cap_drift": "guard-cap-analysis.json:comparison.cap.same_build_drift_over_five_percent",
}

ANALYSIS_FILES = {
    "docs/performance/results/change-0553/analyze_metrics.py":
        "69042472302568c42241fbfcbdad8a8c683f3134e7e7c0ebf9c8ed5ec8f42426",
    "docs/performance/results/change-0553/analyze_guards.py":
        "89c9ed81e66ff3ea35505cfa9aaf0c9ffefdfa4292ccff3757f88c2b64edb2d0",
    "docs/performance/results/change-0553/run.py":
        "f1398fcc87dc7bccfc05930276e26a4a8a21106480613238eab49310ae1c23f0",
    "docs/performance/results/change-0553/capture.py":
        "21b0153923e87a331906552457a5f77bb444686d4ce9c9c99e3efcb2dbe5d2c1",
    "docs/performance/results/change-0553/plan.json":
        "17d810a24912065fde8c71de6109be983f7c5b3bdcea5fd6584f4725c6499ea4",
    "docs/performance/results/change-0546/integration/analyze.py":
        "c380298762c64f104c4aaaf1bf53fae7e613bc7aa9148fe1d6f1da19b2eb9614",
    "docs/performance/results/change-0521/analyze.py":
        "322357892b496ef09ca01aa69bb5a8708182a9a771a4541e3b21a6885a7616ad",
    "tools/validate_perf_corpus_binding.py":
        "e20abbd1220623f387283ff5cbb0cec95ea57edad2e320bd22d93b268cfa9558",
}
SUPPLEMENTAL_FILES = {
    "docs/performance/results/change-0553/guarded_capture.py":
        "ba3a2dc8c8f43531dfdd099310e19aeb9ce6ce160b5ce89dc577e21fc85f4119",
    "docs/performance/results/change-0553/workspace-lock.json":
        "894384b4d9fad830960c1348cd4f5f2ae3d59f8beb908b12c954d2dfb0ed5147",
    "docs/performance/results/change-0553/workspace-Cargo.lock":
        "9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a",
    "docs/performance/results/change-0553/quality-plan.json":
        "5000e7b9b4d3ab7d59ab91ec97839c9aca11ef783acb8bd89250752f2847af91",
}
FROZEN_FILES = {
    "docs/performance/results/change-0553/plan.json": ANALYSIS_FILES[
        "docs/performance/results/change-0553/plan.json"],
    "docs/performance/results/change-0553/run.py": ANALYSIS_FILES[
        "docs/performance/results/change-0553/run.py"],
    "docs/performance/results/change-0553/capture.py": ANALYSIS_FILES[
        "docs/performance/results/change-0553/capture.py"],
    "docs/performance/results/change-0553/adr-manifest.json":
        "c4e7331b91816d3752917c3cb4b320ed56c35efede61baa9170d215546e10ad2",
}

RECEIPT_KEYS = {
    "command", "start_utc", "end_utc", "seconds", "exit_code",
    "execution_stage", "execution_manifest_sha256", "binary_sha256",
    "source_manifest_sha256", "script_sha256", "plan_sha256", "environment",
    "artifacts",
}
RECEIPT_ENVIRONMENT = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
ATTEMPT_KEYS = {
    "command", "cwd", "start_utc", "end_utc", "seconds", "exit_code",
    "source_stable", "source_manifest_sha256", "script_sha256", "run_sha256",
    "plan_sha256", "environment", "artifacts",
}
ATTEMPT_ENVIRONMENT = {
    "TMPDIR", "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
    "RUSTDOCFLAGS", "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD",
}
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
EXPECTED_ERRORS = {
    "late-validator": "value-only edits refuse attribute 'future' on 'c'",
    "late-raw": "invalid worksheet boolean 'maybe'",
}


class VerificationError(ValueError):
    """Malformed, contradictory, or out-of-scope evidence."""


class IncompleteError(VerificationError):
    """Evidence required by the selected component is not retained yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def rel(path: Path) -> str:
    for root in (HERE, REPO):
        try:
            return path.relative_to(root).as_posix()
        except ValueError:
            pass
    return path.as_posix()


def need(path: Path, label: str | None = None, *, directory: bool = False) -> Path:
    label = label or rel(path)
    if not path.exists() or path.is_symlink():
        raise IncompleteError(f"{label} is missing or is a symlink")
    if directory:
        require(path.is_dir(), f"{label} is not a directory")
    else:
        require(path.is_file(), f"{label} is not a regular file")
    return path


def read_bytes(path: Path, label: str | None = None) -> bytes:
    need(path, label)
    try:
        return path.read_bytes()
    except OSError as error:
        raise VerificationError(f"cannot read {label or rel(path)}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    try:
        return read_bytes(path, label).decode("utf-8")
    except UnicodeDecodeError as error:
        raise VerificationError(f"{label or rel(path)} is not UTF-8") from error


def read_json(path: Path, label: str | None = None) -> Any:
    label = label or rel(path)
    try:
        return json.loads(read_text(path, label))
    except json.JSONDecodeError as error:
        raise VerificationError(f"cannot parse {label}: {error}") from error


def sha(path: Path) -> str:
    need(path, f"artifact for hashing: {rel(path)}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise VerificationError(f"cannot hash {rel(path)}: {error}") from error
    return digest.hexdigest()


def is_hash(value: Any) -> bool:
    return isinstance(value, str) and SHA256_RE.fullmatch(value) is not None


def check_hash(value: Any, label: str) -> str:
    require(is_hash(value), f"{label} is not a lowercase SHA-256")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str) and value, f"{label} is missing")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise VerificationError(f"{label} is not ISO-8601") from error
    require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    require(end > start, f"{label} interval is inverted")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and seconds > 0,
            f"{label}.seconds is not positive")
    wall = (end - start).total_seconds()
    require(abs(float(seconds) - wall) <= max(0.25, wall * 0.02 + 0.05),
            f"{label}.seconds does not match UTC interval")
    return start, end


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and "." not in path.parts and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def safe_repo_path(value: Any, label: str) -> Path:
    safe_relative(value, label)
    path = REPO / value
    require(path.resolve().is_relative_to(REPO.resolve()), f"{label} escapes repository")
    return path


def source_name(name: str) -> bool:
    return (
        name in {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
        or name.startswith((".cargo/", "crates/", "tools/perf-baseline/"))
    )


def git(args: list[str], *, env: dict[str, str] | None = None,
        input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(
            args, cwd=REPO, env=env, input=input_data, stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise VerificationError(
            f"Git command failed ({' '.join(args)}): "
            f"{detail.decode(errors='replace')[-2000:]}"
        ) from error


def git_tree_manifest(revision: str) -> dict[str, str]:
    """Hash the frozen Git tree contents without touching the worktree."""
    raw = git([
        "git", "ls-tree", "-r", "-z", revision, "--", "crates",
        "tools/perf-baseline", "Cargo.toml", "Cargo.lock", ".cargo",
        "rust-toolchain.toml",
    ])
    entries: list[tuple[str, str]] = []
    for item in raw.split(b"\0"):
        if not item:
            continue
        try:
            metadata, encoded = item.split(b"\t", 1)
            fields = metadata.split()
            name = encoded.decode("utf-8")
        except (UnicodeDecodeError, ValueError) as error:
            raise VerificationError("Git source tree entry is malformed") from error
        require(len(fields) == 3 and fields[0] in (b"100644", b"100755")
                and fields[1] == b"blob" and source_name(name),
                f"Git source tree entry is out of scope: {name}")
        entries.append((name, fields[2].decode()))
    require(entries, "Git source tree has no source entries")
    response = git(
        ["git", "cat-file", "--batch"],
        input_data=("\n".join(oid for _, oid in entries) + "\n").encode(),
    )
    result: dict[str, str] = {}
    position = 0
    for name, oid in entries:
        end = response.find(b"\n", position)
        require(end >= 0, "Git source object response is truncated")
        fields = response[position:end].split()
        require(len(fields) == 3 and fields[0].decode() == oid and fields[1] == b"blob",
                "Git source object response is malformed")
        try:
            length = int(fields[2])
        except ValueError as error:
            raise VerificationError("Git source object length is malformed") from error
        position = end + 1
        data = response[position:position + length]
        require(len(data) == length, "Git source object is truncated")
        result[name] = hashlib.sha256(data).hexdigest()
        position += length
        require(response[position:position + 1] == b"\n",
                "Git source object separator is missing")
        position += 1
    require(position == len(response) and len(result) == len(entries),
            "Git source tree response has trailing data or duplicate paths")
    return dict(sorted(result.items()))


def source_manifest(path: Path, label: str) -> dict[str, str]:
    value = read_json(path, label)
    require(isinstance(value, dict) and value, f"{label} is empty")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label} path")
        require(source_name(name), f"{label} has out-of-scope source path {name}")
        check_hash(digest, f"{label} {name}")
        require(name not in result, f"{label} repeats {name}")
        result[name] = digest
    require("tools/perf-baseline/src/xlsx_planning_guard.rs" in result,
            f"{label} omits planning-guard source")
    require("crates/litchi-xlsx/examples/perf_cap_boundary.rs" in result,
            f"{label} omits cap-example source")
    return dict(sorted(result.items()))


def current_source_manifest() -> dict[str, str]:
    """Hash the live source set with the same inclusion rule as run.freeze."""
    tracked = git([
        "git", "ls-files", "-z", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ]).split(b"\0")
    untracked = git([
        "git", "ls-files", "--others", "--exclude-standard", "-z", "--",
        "crates", "tools/perf-baseline",
    ]).split(b"\0")
    names = {item.decode() for item in tracked if item}
    names |= {item.decode() for item in untracked if item.endswith(b".rs")}
    result = {}
    for name in sorted(names):
        path = REPO / name
        if path.is_file() and not path.is_symlink():
            result[name] = sha(path)
    return dict(sorted(result.items()))


def patch_paths(path: Path, label: str) -> set[str]:
    data = read_bytes(path, label)
    if not data:
        return set()
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        raise VerificationError(f"{label} is not UTF-8") from error
    names: set[str] = set()
    diff_headers = [line for line in text.splitlines()
                    if line.startswith("diff --git a/")]
    if diff_headers:
        for line in diff_headers:
            fields = line.split()
            require(len(fields) == 4 and fields[2].startswith("a/")
                    and fields[3].startswith("b/"), f"{label} diff header is malformed")
            left, right = fields[2][2:], fields[3][2:]
            require(left == right, f"{label} renames are not admissible")
            safe_relative(left, f"{label} path")
            require(source_name(left), f"{label} contains an out-of-scope path {left}")
            names.add(left)
    else:
        # Candidate draft patches are intentionally stored as ordinary unified
        # diffs.  Pair each file header before admitting its path; accepting a
        # lone header would let a truncated patch masquerade as custody.
        pending: str | None = None
        saw_header = False
        for line in text.splitlines():
            if line.startswith("--- "):
                require(pending is None, f"{label} unified diff has an unpaired header")
                pending = line[4:].split("\t", 1)[0].strip()
                saw_header = True
                continue
            if not line.startswith("+++ "):
                continue
            require(pending is not None, f"{label} unified diff lacks a source header")
            right = line[4:].split("\t", 1)[0].strip()
            left = pending
            pending = None
            if left == "/dev/null":
                require(right.startswith("b/"), f"{label} new-file header is malformed")
                name = right[2:]
            elif right == "/dev/null":
                require(left.startswith("a/"), f"{label} delete header is malformed")
                name = left[2:]
            else:
                require(left.startswith("a/") and right.startswith("b/"),
                        f"{label} unified diff header is malformed")
                require(left[2:] == right[2:], f"{label} renames are not admissible")
                name = left[2:]
            safe_relative(name, f"{label} path")
            require(source_name(name), f"{label} contains an out-of-scope path {name}")
            names.add(name)
        require(saw_header and pending is None,
                f"{label} unified diff headers are incomplete")
    require(names, f"{label} has no file diff headers")
    return names


def witness_for(name: str, expected_digest: str) -> Path | None:
    """Find a retained source witness matching the expected candidate bytes."""
    # The final checkout is intentionally restored to baseline after a reject.
    # Retained candidate snapshots therefore take precedence over the live
    # checkout, whose bytes may no longer be the candidate bytes.
    choices: list[Path] = [HERE / "candidate-sources" / name]
    attempts = HERE / "candidate-attempts"
    if attempts.is_dir() and not attempts.is_symlink():
        choices.extend(
            folder / "sources" / name
            for folder in sorted(attempts.glob("draft-[0-9][0-9]"),
                                 reverse=True)
            if folder.is_dir() and not folder.is_symlink()
        )
    choices.append(REPO / name)
    for path in choices:
        if path.is_file() and not path.is_symlink() and sha(path) == expected_digest:
            return path
    return None


def validate_stage_source(
    stage: str, plan: dict[str, Any], baseline: dict[str, str] | None = None,
    candidate: dict[str, str] | None = None,
) -> dict[str, Any]:
    require(stage in ("baseline", "candidate", "final"), f"unknown source stage {stage}")
    folder = need(HERE / stage, f"{stage} stage", directory=True)
    manifest_path = need(folder / "source-manifest.json", f"{stage}/source-manifest.json")
    patch_path = need(folder / "source.patch", f"{stage}/source.patch")
    manifest = source_manifest(manifest_path, f"{stage}/source-manifest.json")
    patch = patch_paths(patch_path, f"{stage}/source.patch") if patch_path.stat().st_size else set()
    revision = plan["revision"]
    base = baseline or git_tree_manifest(revision)
    if stage == "baseline":
        require(manifest == base, "baseline source manifest differs from frozen Git tree")
        require(not patch, "baseline source.patch is not empty")
    elif stage == "candidate":
        require(manifest != base, "candidate source manifest has no candidate delta")
        delta = {name for name in set(base) | set(manifest)
                 if base.get(name) != manifest.get(name)}
        require(patch == delta, "candidate source.patch changed-path inventory differs")
        for name, digest in manifest.items():
            witness = witness_for(name, digest)
            if witness is not None:
                require(sha(witness) == digest,
                        f"candidate source witness differs: {name}")
            elif name in delta:
                raise IncompleteError(f"candidate source witness is missing: {name}")
        require(patch, "candidate source.patch is empty")
    else:
        allowed = (base,) if candidate is None else (base, candidate)
        require(manifest in allowed,
                "final source is neither the restored baseline nor the bound candidate")
        if manifest == base:
            require(not patch, "restored final source.patch is not empty")
        else:
            delta = {name for name in set(base) | set(manifest)
                     if base.get(name) != manifest.get(name)}
            require(patch == delta, "accepted final source.patch changed-path inventory differs")
            for name, digest in manifest.items():
                witness = witness_for(name, digest)
                if witness is not None:
                    require(sha(witness) == digest,
                            f"final candidate source witness differs: {name}")
                elif name in delta:
                    raise IncompleteError(f"final candidate source witness is missing: {name}")
    return {
        "status": "pass", "stage": stage, "manifest": manifest,
        "manifest_sha256": sha(manifest_path), "entries": len(manifest),
        "patch_sha256": sha(patch_path), "patch_paths": sorted(patch),
    }


def stage_manifest(stage: str) -> tuple[dict[str, str], str]:
    """Return one sealed stage manifest for the decision consumer.

    The decision script uses this small API after the full report validators
    have established the stage relationship.  Keep it source-only so it never
    manufactures a candidate or treats a missing stage as an empty map.
    """
    require(stage in (*STAGES, "final"), f"unknown source stage {stage}")
    path = need(HERE / stage / "source-manifest.json",
                f"{stage}/source-manifest.json")
    manifest = source_manifest(path, f"{stage}/source-manifest.json")
    return manifest, sha(path)


def validate_frozen_inputs() -> dict[str, Any]:
    value = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(value, dict) and set(value) == {"frozen_utc", "files"},
            "frozen-inputs envelope differs")
    parse_time(value["frozen_utc"], "frozen-inputs.frozen_utc")
    require(value["files"] == FROZEN_FILES, "frozen-inputs file inventory differs")
    for name, digest in value["files"].items():
        check_hash(digest, f"frozen-inputs.files.{name}")
        path = safe_repo_path(name, f"frozen-inputs.files.{name}")
        require(sha(path) == digest, f"frozen input changed: {name}")
    return {"status": "pass", "sha256": sha(FROZEN), "files": value["files"]}


def validate_adr() -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    require(isinstance(value, dict) and set(value) == {"files", "checked_utc", "status"},
            "ADR manifest envelope differs")
    parse_time(value["checked_utc"], "adr-manifest.checked_utc")
    require(value["status"] == "all previously read ADRs unchanged",
            "ADR manifest status differs")
    files = value["files"]
    require(isinstance(files, dict) and files, "ADR file map is empty")
    for name, digest in files.items():
        safe_relative(name, f"ADR path {name}")
        require(name.startswith("docs/adr/"), f"ADR path is out of scope: {name}")
        check_hash(digest, f"ADR hash {name}")
        require(sha(REPO / name) == digest, f"ADR changed: {name}")
    return {"status": "pass", "sha256": sha(ADR), "entries": len(files)}


def validate_supplemental_inputs() -> dict[str, Any]:
    value = read_json(SUPPLEMENTAL, "supplemental-inputs.json")
    require(isinstance(value, dict)
            and set(value) == {"frozen_utc", "scope", "files"},
            "supplemental-inputs envelope differs")
    parse_time(value["frozen_utc"], "supplemental-inputs.frozen_utc")
    require(value["scope"] == (
        "Workspace lock and quality commands frozen before all builds; guarded captures "
        "check the ignored workspace lock per child from the first baseline pipeline."
    ), "supplemental-inputs scope differs")
    require(value["files"] == SUPPLEMENTAL_FILES,
            "supplemental-inputs file inventory differs")
    for name, digest in value["files"].items():
        require(sha(safe_repo_path(name, f"supplemental-inputs.{name}")) == digest,
                f"supplemental input changed: {name}")
    return {"status": "pass", "sha256": sha(SUPPLEMENTAL), "files": value["files"]}


def validate_analysis_inputs() -> dict[str, Any]:
    value = read_json(ANALYSIS_INPUTS, "analysis-inputs.json")
    require(isinstance(value, dict)
            and set(value) == {"schema", "frozen_utc", "scope", "files"},
            "analysis-inputs envelope differs")
    require(value["schema"] == "xlsx_0553_analysis_inputs_v1",
            "analysis-inputs schema differs")
    parse_time(value["frozen_utc"], "analysis-inputs.frozen_utc")
    require(value["scope"] == (
        "Corrected inherited main and guard/cap analyzers frozen before the first baseline capture. "
        "All0552 gates retained."
    ), "analysis-inputs scope differs")
    require(value["files"] == ANALYSIS_FILES, "analysis-inputs file inventory differs")
    for name, digest in value["files"].items():
        require(sha(safe_repo_path(name, f"analysis-inputs.{name}")) == digest,
                f"analysis input changed: {name}")
    amendment = HERE / "analyzer-amendment.json"
    require(not amendment.exists(), "0553 has no analyzer-amendment custody path")
    return {"status": "pass", "sha256": sha(ANALYSIS_INPUTS), "files": value["files"]}


def validate_workspace_lock() -> dict[str, Any]:
    value = read_json(LOCK_BINDING, "workspace-lock.json")
    require(isinstance(value, dict) and set(value) == {
        "frozen_utc", "path", "sha256", "retained", "reason", "baseline_cap_not_started",
    }, "workspace-lock envelope differs")
    parse_time(value["frozen_utc"], "workspace-lock.frozen_utc")
    require(value["path"] == "Cargo.lock"
            and value["retained"] == "workspace-Cargo.lock"
            and value["sha256"] == WORKSPACE_LOCK_SHA256
            and value["reason"] == (
                "Ignored workspace lock frozen before every0553 build. guarded_capture.py checks it "
                "per child, starting with baseline."
            ) and value["baseline_cap_not_started"] is True,
            "workspace-lock identity differs")
    require(sha(REPO / value["path"]) == value["sha256"],
            "live workspace Cargo.lock differs")
    require(sha(HERE / value["retained"]) == value["sha256"],
            "retained workspace Cargo.lock differs")
    return {"status": "pass", "sha256": sha(LOCK_BINDING), "lock_sha256": value["sha256"]}


def validate_host() -> dict[str, Any]:
    value = read_json(HOST, "host.json")
    require(isinstance(value, dict) and set(value) == {
        "utc", "uname", "cpu_affinity", "rustc", "cargo",
    }, "host envelope differs")
    parse_time(value["utc"], "host.utc")
    require(isinstance(value["uname"], str) and value["uname"].strip()
            and isinstance(value["rustc"], str) and value["rustc"].strip()
            and isinstance(value["cargo"], str) and value["cargo"].strip(),
            "host tool identity is missing")
    affinity = value["cpu_affinity"]
    require(isinstance(affinity, list) and affinity == sorted(set(affinity))
            and all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                    for item in affinity)
            and CPU in affinity, "host CPU affinity differs")
    return {"status": "pass", "sha256": sha(HOST), "cpu_affinity": affinity}


PLAN_ADMISSION = {
    "primary": "Every one-percent case/shape/repeat, including managed, workflow p50 and mean improve at least 3%.",
    "latency_guards": "Every one-cell case/shape/repeat and valid-noop planning/cap control p50 and mean <= 1.05x matched baseline.",
    "refusals": "Exact errors/retry identities unchanged; each invalid guard p50 and mean <= max(1.05x corresponding baseline invalid,2x baseline valid); each invalid attributable peak <= 1.10x baseline valid attributable peak.",
    "workflow_memory": "For each main case/shape/repeat, max planning/commit/publication absolute region peak minus planning live_bytes_before <= 1.05x baseline; process peakRSS <= 1.05x baseline. Use allocation-instrumented comparisons only for allocation metrics.",
    "noop_memory": "Valid planning guard region peak minus live_bytes_before <= 1.05x baseline. Decline if durable metadata cannot fit representative no-op memory envelope.",
    "allocation": "Workflow sum allocated bytes <= 1.05x baseline; allocation/reallocation calls and per-phase retained-after values reported individually. Any >5% adverse/drift row requires explicit review; no hidden geometric mean.",
    "proof_resources": "Explicit checked byte/event caps and fallible reservation before growth required; immutable original-source binding and complete fallback. Concrete candidate cap/source hash frozen before application.",
    "correctness": "Exact scanner/writer differential tests, source fallback/error-order, no-op, managed execution/cancellation, source identity, inverse/clone/re-edit, preservation and quality gates required.",
    "profile": "Exact commit Ir must decrease for every shape/repeat if pilot native/memory gates pass; absence of scan symbol alone is insufficient.",
    "disposition": "All mandatory gates required; reject and restore baseline on failure. No exception for a necessary enabler without measured justification.",
}


def validate_plan() -> dict[str, Any]:
    value = read_json(PLAN, "plan.json")
    require(isinstance(value, dict), "plan is not an object")
    expected_keys = {
        "revision", "created_utc", "previous_turn", "priority", "scope", "hypothesis",
        "cpu", "owned_paths", "cases", "shapes", "native", "alloc", "profile", "guard",
        "cap", "native_order", "allocator_order", "admission", "limitations",
    }
    require(set(value) == expected_keys, "plan field inventory differs")
    require(value["revision"] == REVISION and REVISION_RE.fullmatch(value["revision"]),
            "plan revision differs")
    try:
        subprocess.run(["git", "cat-file", "-e", value["revision"] + "^{commit}"],
                       cwd=REPO, check=True, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError("plan revision is not a Git commit") from error
    parse_time(value["created_utc"], "plan.created_utc")
    require(value["previous_turn"] == (
        "progress: committed sealed0552 rejected planning-time proof experiment; baseline production "
        "restored, seven publicguards retained, fullquality passed and ownedtarget cleaned"
    ), "plan previous_turn differs")
    require(value["priority"] == PRIORITY and value["scope"] == (
        "Matched commit-local compact source-cell proof experiment for source-backed XLSX MultiSourceEdit"
    ), "plan identity differs")
    require(value["hypothesis"] == (
        "Collect an ephemeral compact source-cell layout only at commit for effective existing-cell edits, "
        "using a direct borrowed XML walk to avoid complete unchanged-tag layout materialization. Leave "
        "planning and no-op paths unchanged; no second semantic Store and no proof retained in Snapshot."
    ), "plan hypothesis differs")
    require(value["cpu"] == CPU and value["owned_paths"] == [TARGET_STRING],
            "plan execution envelope differs")
    require(value["cases"] == list(CASES) and value["shapes"] == list(SHAPES),
            "plan main matrix differs")
    require(value["native"] == {"repeats": 2, "samples": 200, "warmup": 20}
            and value["alloc"] == {"repeats": 2, "samples": 20, "warmup": 3},
            "plan sample matrix differs")
    require(value["profile"] == {
        "repeats": 2, "samples": 1, "warmup": 0, "case": PROFILE_CASE,
        "owner": PROFILE_OWNER, "parent": "litchi_perf_baseline::run_xlsx_cell_values_edit_save",
    }, "plan profile matrix differs")
    require(value["guard"] == {
        "shapes": list(GUARD_SHAPES), "cases": list(GUARD_CASES), "repeats": 2,
        "native_samples": 200, "native_warmup": 20, "alloc_samples": 20,
        "alloc_warmup": 3,
    } and value["cap"] == {
        "sizes": list(CAP_SIZES), "repeats": 2, "samples": 200, "warmup": 20,
    }, "plan guard/cap matrix differs")
    expected_order = [
        "baseline r1", "candidate r1", "candidate r2",
        "retained baseline r2 under candidate source manifest",
    ]
    require(value["native_order"] == expected_order
            and value["allocator_order"] == expected_order, "plan ABBA order differs")
    require(value["admission"] == PLAN_ADMISSION, "plan admission text differs")
    require(value["limitations"] == [
        "No hardware/cold/provider/scaling claim from this campaign.",
        "Current single-sheet SourceEdit API is outside measured owner.",
        "Profiles conditional on native/memory pilot; all failures retained.",
        "Historical guards use no-op oracle after planning interval; full no-op publication time is not measured.",
    ], "plan limitations differ")
    return {"status": "pass", "sha256": sha(PLAN), "revision": value["revision"]}


def validate_quality_plan() -> dict[str, Any]:
    value = read_json(QUALITY_PLAN, "quality-plan.json")
    require(isinstance(value, dict) and set(value) == {"scope", "commands"},
            "quality-plan envelope differs")
    require(value["scope"] == (
        "Commit-local compact XLSX source-cell proof candidate or restored baseline plus public guards; "
        "warning-denied owner checks and workspace feature check; no iWork optimization."
    ), "quality-plan scope differs")
    commands = value["commands"]
    require(isinstance(commands, dict) and len(commands) == 11,
            "quality-plan command count differs")
    for name, command in commands.items():
        require(isinstance(name, str) and re.fullmatch(r"[A-Za-z0-9_-]+", name)
                and isinstance(command, list) and command
                and all(isinstance(item, str) and item for item in command),
                f"quality-plan command is malformed: {name}")
    return {"status": "pass", "sha256": sha(QUALITY_PLAN), "commands": commands}


def validate_receipt_common(value: Any, label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS,
            f"{label} receipt schema differs")
    start, end = interval(value, label)
    environment = value["environment"]
    require(isinstance(environment, dict) and set(environment) == RECEIPT_ENVIRONMENT
            and all(item is None for item in environment.values()),
            f"{label}.environment differs")
    require(isinstance(value["command"], list) and value["command"]
            and all(isinstance(item, str) and item for item in value["command"]),
            f"{label}.command is malformed")
    require(isinstance(value["exit_code"], int) and not isinstance(value["exit_code"], bool),
            f"{label}.exit_code is malformed")
    for field in ("execution_manifest_sha256", "source_manifest_sha256",
                  "script_sha256", "plan_sha256"):
        check_hash(value[field], f"{label}.{field}")
    if value["binary_sha256"] is not None:
        check_hash(value["binary_sha256"], f"{label}.binary_sha256")
    artifacts = value["artifacts"]
    require(isinstance(artifacts, dict), f"{label}.artifacts is missing")
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).name == name,
                f"{label}.artifacts has a non-local name")
        check_hash(digest, f"{label}.artifacts.{name}")
    return start, end


def validate_host_sidecar(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict) and set(value) == {
        "observed_utc", "compiler_processes", "scope",
    }, f"{label} host sidecar differs")
    parse_time(value["observed_utc"], f"{label}.observed_utc")
    require(value["scope"] == HOST_SCOPE and isinstance(value["compiler_processes"], list),
            f"{label} host scope differs")
    for index, process in enumerate(value["compiler_processes"]):
        require(isinstance(process, dict) and set(process) == {"pid", "comm", "cwd"},
                f"{label}.compiler_processes[{index}] differs")
        require(isinstance(process["pid"], int) and not isinstance(process["pid"], bool)
                and process["pid"] > 0 and process["comm"] in ("cargo", "rustc")
                and isinstance(process["cwd"], str) and process["cwd"],
                f"{label}.compiler_processes[{index}] is malformed")


def validate_artifacts(folder: Path, stem: str, receipt: dict[str, Any],
                       expected: set[str], label: str) -> None:
    require(set(receipt["artifacts"]) == expected,
            f"{label}.artifacts inventory differs")
    actual = {item.name for item in folder.iterdir()
              if item.is_file() and not item.is_symlink()
              and item.name.startswith(f"{stem}.")
              and item.name != f"{stem}.receipt.json"}
    require(actual == expected, f"{label} artifact inventory differs")
    for name, digest in receipt["artifacts"].items():
        path = folder / name
        require(sha(path) == digest, f"{label}/{name} digest differs")
        if name.endswith(".host.json"):
            validate_host_sidecar(path, f"{label}/{name}")


def expected_build_command(kind: str) -> list[str]:
    prefix = [
        "env", f"TMPDIR={TARGET / 'tmp'}", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0",
        "cargo", "build", "--release", "--locked",
    ]
    if kind in ("normal", "alloc"):
        command = prefix + ["--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
                            "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else ""),
                            "--target-dir", TARGET_STRING]
        return command + (["--features", "allocator-metrics"] if kind == "alloc" else [])
    if kind in ("guard-normal", "guard-alloc"):
        command = prefix + ["--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
                            "xlsx_planning_guard", "--target-dir", TARGET_STRING]
        return command + (["--features", "allocator-metrics"] if kind == "guard-alloc" else [])
    require(kind == "cap", f"unknown build kind {kind}")
    return prefix + ["-p", "litchi-xlsx", "--example", "perf_cap_boundary",
                     "--target-dir", TARGET_STRING]


def descriptor_spec(kind: str) -> tuple[str, str]:
    if kind in ("normal", "alloc"):
        return f"binary-{kind}.json", kind
    if kind in ("guard-normal", "guard-alloc"):
        return f"binary-{kind}.json", kind
    require(kind == "cap", f"unknown binary kind {kind}")
    return "binary-cap.json", "cap"


def validate_binary_descriptor(stage: str, kind: str,
                               manifest_sha: str) -> dict[str, Any]:
    """Validate one retained descriptor without assuming its binary still exists.

    Descriptors remain the source of truth after the owned build target is
    removed.  The caller separately proves either that the binary is still
    present or that a validated cleanup record carries the same digest.
    """
    require(stage in STAGES and kind in BINARY_KINDS,
            f"invalid binary descriptor dimension: {stage}/{kind}")
    folder = HERE / stage
    descriptor_name, scratch_name = descriptor_spec(kind)
    descriptor_path = need(folder / descriptor_name, f"{stage}/{descriptor_name}")
    descriptor = read_json(descriptor_path, rel(descriptor_path))
    require(isinstance(descriptor, dict) and set(descriptor) == {
        "path", "sha256", "bytes", "build_receipt_sha256", "source_manifest_sha256",
    }, f"{rel(descriptor_path)} inventory differs")
    expected_path = SCRATCH_ROOT / stage / scratch_name
    require(isinstance(descriptor["path"], str)
            and Path(descriptor["path"]) == expected_path,
            f"{rel(descriptor_path)} path differs")
    digest = check_hash(descriptor["sha256"], f"{rel(descriptor_path)}.sha256")
    require(isinstance(descriptor["bytes"], int) and not isinstance(descriptor["bytes"], bool)
            and descriptor["bytes"] > 0, f"{rel(descriptor_path)}.bytes differs")
    require(descriptor["source_manifest_sha256"] == manifest_sha,
            f"{rel(descriptor_path)} source binding differs")
    receipt_path = need(folder / f"build-{kind}.receipt.json",
                        f"{stage}/build-{kind}.receipt.json")
    require(descriptor["build_receipt_sha256"] == sha(receipt_path),
            f"{rel(descriptor_path)} build receipt hash differs")
    return {
        "stage": stage, "kind": kind, "path": str(expected_path),
        "sha256": digest, "bytes": descriptor["bytes"],
        "descriptor_path": rel(descriptor_path),
        "descriptor_sha256": sha(descriptor_path),
        "build_receipt_path": rel(receipt_path),
        "build_receipt_sha256": sha(receipt_path),
    }


def expected_binary_keys() -> set[str]:
    return {f"{stage}/{kind}" for stage in STAGES for kind in BINARY_KINDS}


def cleanup_binary_digest(cleanup: dict[str, Any], stage: str, kind: str) -> str:
    hashes = cleanup.get("binary_sha256_by_kind")
    require(isinstance(hashes, dict) and set(hashes) == expected_binary_keys(),
            "cleanup binary_sha256_by_kind inventory differs")
    return check_hash(hashes[f"{stage}/{kind}"],
                      f"cleanup.binary_sha256_by_kind.{stage}/{kind}")


def validate_build(stage: str, kind: str, manifest_sha: str,
                   cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    descriptor = validate_binary_descriptor(stage, kind, manifest_sha)
    binary = Path(descriptor["path"])
    if os.path.lexists(binary):
        require(not binary.is_symlink() and binary.is_file(),
                f"{descriptor['descriptor_path']} retained binary is not a regular file")
        require(sha(binary) == descriptor["sha256"]
                and binary.stat().st_size == descriptor["bytes"],
                f"{descriptor['descriptor_path']} binary custody differs")
    else:
        require(cleanup is not None and cleanup.get("owned_paths_absent") is True
                and cleanup.get("accessible_process_references") == [],
                f"{descriptor['descriptor_path']} retained binary is missing without cleanup custody")
        require(cleanup_binary_digest(cleanup, stage, kind) == descriptor["sha256"],
                f"{descriptor['descriptor_path']} cleanup binary custody differs")
    folder = HERE / stage
    receipt_path = need(folder / f"build-{kind}.receipt.json",
                        f"{stage}/build-{kind}.receipt.json")
    receipt = read_json(receipt_path, rel(receipt_path))
    start, end = validate_receipt_common(receipt, rel(receipt_path))
    require(receipt["exit_code"] == 0 and receipt["binary_sha256"] is None
            and receipt["execution_stage"] == stage
            and receipt["execution_manifest_sha256"] == manifest_sha
            and receipt["source_manifest_sha256"] == manifest_sha
            and receipt["script_sha256"] == sha(RUN)
            and receipt["plan_sha256"] == sha(PLAN)
            and receipt["command"] == expected_build_command(kind),
            f"{rel(receipt_path)} binding differs")
    stem = f"build-{kind}"
    validate_artifacts(folder, stem, receipt,
                       {f"{stem}.host.json", f"{stem}.stdout", f"{stem}.stderr"},
                       rel(receipt_path))
    descriptor["start"] = start
    descriptor["end"] = end
    descriptor["receipt_sha256"] = sha(receipt_path)
    return descriptor


def validate_owned_binaries(cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate exactly the ten baseline/candidate owned binary descriptors."""
    plan = validate_plan()
    baseline = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, baseline["manifest"])
    manifests = {"baseline": baseline, "candidate": candidate}
    binaries: dict[str, dict[str, Any]] = {stage: {} for stage in STAGES}
    for stage in STAGES:
        manifest_sha = manifests[stage]["manifest_sha256"]
        for kind in BINARY_KINDS:
            binaries[stage][kind] = validate_build(
                stage, kind, manifest_sha, cleanup=cleanup
            )
    require(set(binaries) == set(STAGES)
            and all(set(binaries[stage]) == set(BINARY_KINDS) for stage in STAGES),
            "owned binary custody is not exactly ten descriptors")
    return {
        "status": "pass", "count": len(STAGES) * len(BINARY_KINDS),
        "binaries": binaries,
        "source": {stage: {"manifest_sha256": value["manifest_sha256"]}
                    for stage, value in manifests.items()},
    }


def expected_main_jobs(stage: str, lane: str) -> list[dict[str, Any]]:
    require(stage in STAGES and lane in LANES, "invalid main matrix dimension")
    repeats = (1,) if lane == "preflight" else (1, 2)
    result = []
    for repeat in repeats:
        shapes = SHAPES if repeat == 1 else tuple(reversed(SHAPES))
        for shape in shapes:
            for index, case in enumerate(CASES):
                result.append({
                    "name": f"{lane}-r{repeat}-{shape}-c{index}", "stage": stage,
                    "lane": lane, "repeat": repeat, "shape": shape, "case": case,
                    "samples": 1 if lane == "preflight" else (200 if lane == "native" else 20),
                    "warmup": 0 if lane == "preflight" else (20 if lane == "native" else 3),
                    "execution_stage": "candidate" if stage == "candidate" or repeat == 2 else "baseline",
                })
    return result


def expected_guard_jobs(lane: str) -> list[dict[str, Any]]:
    require(lane in GUARD_LANES, "invalid guard lane")
    result = []
    samples, warmup = (200, 20) if lane == "normal" else (20, 3)
    name_lane = "native" if lane == "normal" else "alloc"
    for repeat in (1, 2):
        shapes = GUARD_SHAPES if repeat == 1 else tuple(reversed(GUARD_SHAPES))
        for shape in shapes:
            for case in GUARD_CASES:
                result.append({
                    "name": f"guard-{name_lane}-r{repeat}-{shape}-{case}",
                    "lane": lane, "repeat": repeat, "shape": shape, "case": case,
                    "samples": samples, "warmup": warmup,
                })
    return result


def expected_cap_jobs() -> list[dict[str, Any]]:
    result = []
    for repeat in (1, 2):
        sizes = CAP_SIZES if repeat == 1 else tuple(reversed(CAP_SIZES))
        for size in sizes:
            result.append({"name": f"cap-r{repeat}-{size}", "repeat": repeat,
                           "size": size, "samples": 200, "warmup": 20})
    return result


def expected_profile_jobs(stage: str) -> list[dict[str, Any]]:
    result = []
    for repeat in (1, 2):
        for shape in SHAPES if repeat == 1 else tuple(reversed(SHAPES)):
            result.append({"name": f"profile-r{repeat}-{shape}-c0", "stage": stage,
                           "repeat": repeat, "shape": shape})
    return result


def expected_main_command(job: dict[str, Any], binary: dict[str, Any]) -> list[str]:
    folder = HERE / job["stage"]
    return ["taskset", "-c", str(CPU), "/usr/bin/time", "-v", binary["path"],
            "--case", job["case"], "--xlsx-cell-crud-shape", job["shape"],
            "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
            "--json", str(folder / (job["name"] + ".json")),
            "--corpus-manifest", str(folder / (job["name"] + ".catalog.json"))]


def validate_capture_receipt(stage: str, job: dict[str, Any], binary: dict[str, Any],
                             manifest_sha: str, baseline_sha: str, candidate_sha: str,
                             kind: str) -> tuple[dt.datetime, dt.datetime]:
    folder = HERE / stage
    receipt_path = need(folder / f"{job['name']}.receipt.json", rel(folder / f"{job['name']}.receipt.json"))
    receipt = read_json(receipt_path, rel(receipt_path))
    start, end = validate_receipt_common(receipt, rel(receipt_path))
    expected_execution = "candidate" if stage == "candidate" or job["repeat"] == 2 else "baseline"
    expected_manifest = candidate_sha if expected_execution == "candidate" else baseline_sha
    if kind == "main":
        command = expected_main_command(job, binary)
        artifacts = {f"{job['name']}.json", f"{job['name']}.catalog.json",
                     f"{job['name']}.stdout", f"{job['name']}.stderr",
                     f"{job['name']}.host.json"}
    elif kind == "guard":
        command = ["taskset", "-c", str(CPU), binary["path"], "--shape", job["shape"],
                   "--case", job["case"], "--samples", str(job["samples"]),
                   "--warmup", str(job["warmup"]), "--json",
                   str(folder / (job["name"] + ".json"))]
        artifacts = {f"{job['name']}.json", f"{job['name']}.stdout",
                     f"{job['name']}.stderr", f"{job['name']}.host.json"}
    else:
        command = ["taskset", "-c", str(CPU), binary["path"], "--size", str(job["size"]),
                   "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
                   "--json", str(folder / (job["name"] + ".json")), "--fixture-out",
                   str(folder / (job["name"] + ".zip"))]
        artifacts = {f"{job['name']}.json", f"{job['name']}.stdout",
                     f"{job['name']}.stderr", f"{job['name']}.host.json",
                     f"{job['name']}.zip"}
    require(receipt["exit_code"] == 0 and receipt["binary_sha256"] == binary["sha256"]
            and receipt["execution_stage"] == expected_execution
            and receipt["execution_manifest_sha256"] == expected_manifest
            and receipt["source_manifest_sha256"] == manifest_sha
            and receipt["script_sha256"] == sha(RUN)
            and receipt["plan_sha256"] == sha(PLAN)
            and receipt["command"] == command,
            f"{rel(receipt_path)} binding differs")
    validate_artifacts(folder, job["name"], receipt, artifacts, rel(receipt_path))
    return start, end


def stage_inventory(stage: str, profiles_present: bool) -> set[str]:
    expected = {"source-manifest.json", "source.patch"}
    for kind in ("normal", "alloc", "guard-normal", "guard-alloc", "cap"):
        stem = f"build-{kind}"
        expected |= {f"{stem}.receipt.json", f"{stem}.host.json",
                     f"{stem}.stdout", f"{stem}.stderr", f"binary-{kind}.json"}
    for lane in LANES:
        for job in expected_main_jobs(stage, lane):
            stem = job["name"]
            expected |= {f"{stem}.receipt.json", f"{stem}.host.json", f"{stem}.json",
                         f"{stem}.catalog.json", f"{stem}.stdout", f"{stem}.stderr"}
    for lane in GUARD_LANES:
        for job in expected_guard_jobs(lane):
            stem = job["name"]
            expected |= {f"{stem}.receipt.json", f"{stem}.host.json", f"{stem}.json",
                         f"{stem}.stdout", f"{stem}.stderr"}
    for job in expected_cap_jobs():
        stem = job["name"]
        expected |= {f"{stem}.receipt.json", f"{stem}.host.json", f"{stem}.json",
                     f"{stem}.stdout", f"{stem}.stderr", f"{stem}.zip"}
    if profiles_present:
        for job in expected_profile_jobs(stage):
            stem = job["name"]
            expected |= {f"{stem}.receipt.json", f"{stem}.host.json", f"{stem}.json",
                         f"{stem}.catalog.json", f"{stem}.stdout", f"{stem}.stderr",
                         f"{stem}.callgrind"}
    folder = need(HERE / stage, f"{stage} stage", directory=True)
    actual = {item.name for item in folder.iterdir()
              if item.is_file() and not item.is_symlink()}
    require(actual == expected, f"{stage} raw artifact inventory differs")
    return expected


def validate_stage_capture(stage: str, baseline_sha: str, candidate_sha: str,
                           *, require_complete: bool = True,
                           cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    folder = HERE / stage
    if not folder.exists():
        raise IncompleteError(f"{stage} capture directory is missing")
    manifest = source_manifest(folder / "source-manifest.json", f"{stage}/source-manifest.json")
    manifest_sha = sha(folder / "source-manifest.json")
    expected_sha = baseline_sha if stage == "baseline" else candidate_sha
    require(manifest_sha == expected_sha, f"{stage} capture source binding differs")
    builds: dict[str, dict[str, Any]] = {}
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for kind in ("normal", "alloc", "guard-normal", "guard-alloc", "cap"):
        item = validate_build(stage, kind, manifest_sha, cleanup=cleanup)
        builds[kind] = item
        intervals.append((item["start"], item["end"], f"build-{kind}"))
    rows = {"main": [], "guard": [], "cap": []}
    for lane in LANES:
        jobs = expected_main_jobs(stage, lane)
        for job in jobs:
            kind = "alloc" if lane == "alloc" else "normal"
            times = validate_capture_receipt(stage, job, builds[kind], manifest_sha,
                                             baseline_sha, candidate_sha, "main")
            rows["main"].append(job["name"])
            intervals.append((times[0], times[1], job["name"]))
    for lane in GUARD_LANES:
        binary = builds["guard-normal" if lane == "normal" else "guard-alloc"]
        for job in expected_guard_jobs(lane):
            times = validate_capture_receipt(stage, job, binary, manifest_sha,
                                             baseline_sha, candidate_sha, "guard")
            rows["guard"].append(job["name"])
            intervals.append((times[0], times[1], job["name"]))
    for job in expected_cap_jobs():
        times = validate_capture_receipt(stage, job, builds["cap"], manifest_sha,
                                         baseline_sha, candidate_sha, "cap")
        rows["cap"].append(job["name"])
        intervals.append((times[0], times[1], job["name"]))
    profile_paths = list(folder.glob("profile-*.receipt.json"))
    stage_inventory(stage, bool(profile_paths))
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{stage} build/capture receipts overlap")
    return {"status": "pass", "stage": stage, "manifest_sha256": manifest_sha,
            "builds": {kind: {k: v for k, v in item.items() if k not in ("start", "end")}
                       for kind, item in builds.items()},
            "counts": {key: len(value) for key, value in rows.items()},
            "intervals": len(intervals), "profile_receipts": len(profile_paths)}


def bundle_snapshot() -> dict[str, str]:
    result: dict[str, str] = {}
    for path in HERE.rglob("*"):
        if path.is_file() and not path.is_symlink():
            result[rel(path)] = sha(path)
    return result


def import_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load {rel(path)}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        raise VerificationError(f"cannot import {rel(path)}: {error}") from error
    return module


def deterministic_document(value: dict[str, Any]) -> bytes:
    try:
        return (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    except (TypeError, ValueError) as error:
        raise VerificationError(f"analysis document is not deterministic JSON: {error}") from error


def cli_json_default(value: Any) -> str:
    """Serialize verifier-only timestamps without weakening JSON custody."""
    if isinstance(value, dt.datetime):
        return value.isoformat()
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def cli_json(value: dict[str, Any]) -> str:
    """Format a CLI envelope while rejecting every unknown object type."""
    return json.dumps(value, indent=2, sort_keys=True, default=cli_json_default)


def analyzer_output(names: tuple[str, ...], label: str) -> tuple[Path, list[Path]]:
    paths = [HERE / name for name in names if (HERE / name).exists()]
    for path in paths:
        require(path.is_file() and not path.is_symlink(),
                f"{label} output is not a regular file: {rel(path)}")
    if not paths:
        raise IncompleteError(f"{label} canonical report is missing")
    return paths[0], paths


def replay_analyzer(path: Path, module_name: str, names: tuple[str, ...],
                    schema: str, label: str) -> dict[str, Any]:
    before = bundle_snapshot()
    module = import_module(path, module_name)
    require(hasattr(module, "analyze") and callable(module.analyze),
            f"{label} has no callable analyze()")
    try:
        value = module.analyze()
    except FileNotFoundError as error:
        raise IncompleteError(f"{label} is missing {error.filename}") from error
    except (KeyError, TypeError, ValueError, OSError, AssertionError) as error:
        raise VerificationError(f"{label} failed: {error}") from error
    after = bundle_snapshot()
    require(before == after, f"{label} replay mutated retained evidence")
    require(isinstance(value, dict) and value.get("schema") == schema,
            f"{label} schema differs")
    encoded = deterministic_document(value)
    output, paths = analyzer_output(names, label)
    for candidate in paths:
        require(candidate.read_bytes() == encoded,
                f"{label} canonical replay differs: {rel(candidate)}")
    return {"status": value.get("status"), "schema": schema, "path": rel(output),
            "sha256": sha(output), "document": value}


def require_analyzer_pass(result: dict[str, Any], label: str) -> None:
    require(result.get("status") == "pass", f"{label} did not produce a passing analysis")


def bundle_path(value: Any, label: str) -> Path:
    """Resolve a path declared inside this evidence bundle."""
    safe_relative(value, label)
    path = HERE / value
    require(path.resolve().is_relative_to(HERE.resolve()),
            f"{label} escapes the 0553 evidence bundle")
    return path


def hash_map(value: Any, label: str) -> dict[str, str]:
    """Validate a compact path-to-content-hash map without accepting aliases."""
    require(isinstance(value, dict) and value, f"{label} is empty")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label} path")
        require(source_name(name), f"{label} has an out-of-scope path: {name}")
        check_hash(digest, f"{label}.{name}")
        require(name not in result, f"{label} repeats {name}")
        result[name] = digest
    return dict(sorted(result.items()))


def git_index_manifest(env: dict[str, str]) -> dict[str, str]:
    """Hash the source files in an isolated temporary Git index."""
    raw = git(["git", "ls-files", "--stage", "-z"], env=env)
    entries: list[tuple[str, str]] = []
    for item in raw.split(b"\0"):
        if not item:
            continue
        try:
            metadata, encoded = item.split(b"\t", 1)
            fields = metadata.split()
            name = encoded.decode("utf-8")
        except (UnicodeDecodeError, ValueError) as error:
            raise VerificationError("temporary Git index entry is malformed") from error
        if not source_name(name):
            continue
        require(len(fields) == 3 and fields[0] in (b"100644", b"100755")
                and fields[2] == b"0",
                f"temporary Git index source entry is malformed: {name}")
        entries.append((name, fields[1].decode()))
    require(entries, "temporary Git index has no source entries")
    response = git(
        ["git", "cat-file", "--batch"], env=env,
        input_data=("\n".join(oid for _, oid in entries) + "\n").encode(),
    )
    result: dict[str, str] = {}
    position = 0
    for name, oid in entries:
        end = response.find(b"\n", position)
        require(end >= 0, "temporary Git object response is truncated")
        fields = response[position:end].split()
        require(len(fields) == 3 and fields[0].decode() == oid
                and fields[1] == b"blob", "temporary Git object response is malformed")
        try:
            length = int(fields[2])
        except ValueError as error:
            raise VerificationError("temporary Git object length is malformed") from error
        position = end + 1
        data = response[position:position + length]
        require(len(data) == length, "temporary Git object is truncated")
        result[name] = hashlib.sha256(data).hexdigest()
        position += length
        require(response[position:position + 1] == b"\n",
                "temporary Git object separator is missing")
        position += 1
    require(position == len(response) and len(result) == len(entries),
            "temporary Git object response has trailing data")
    return dict(sorted(result.items()))


def replay_candidate_patches(paths: list[Path], revision: str) -> dict[str, str]:
    """Replay candidate patches into a temporary index, never the checkout."""
    with tempfile.TemporaryDirectory(prefix=".litchi-0553-candidate-",
                                      dir=REPO.parent) as folder:
        index = Path(folder) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", revision], env=env)
        for path in paths:
            git(["git", "apply", "--cached", "--binary", str(path)], env=env)
        return git_index_manifest(env)


def replay_candidate_patch(path: Path, revision: str) -> dict[str, str]:
    """Replay one candidate patch into a temporary index, never the checkout."""
    return replay_candidate_patches([path], revision)


def candidate_attempt_paths() -> list[Path]:
    root = HERE / "candidate-attempts"
    if not root.exists():
        raise IncompleteError("candidate-attempts directory is missing")
    require(root.is_dir() and not root.is_symlink(),
            "candidate-attempts is not a directory")
    paths = sorted((path for path in root.iterdir() if not path.is_symlink()),
                   key=lambda item: item.name)
    require(paths, "candidate-attempts is empty")
    numbers: list[int] = []
    for path in paths:
        require(path.is_dir() and re.fullmatch(r"draft-[0-9]{2}", path.name),
                f"candidate attempt label is unsafe: {path.name}")
        numbers.append(int(path.name[-2:]))
    require(numbers == list(range(1, len(numbers) + 1)),
            "candidate attempt sequence is not consecutive from draft-01")
    return paths


def validate_candidate_attempts(plan: dict[str, Any],
                                baseline: dict[str, str]) -> dict[str, Any]:
    """Validate every immutable candidate draft and its first application record.

    Draft custody is deliberately separate from the selected final binding.
    A later iteration may replace the draft, but it cannot erase or silently
    reinterpret an earlier patch/application pair.
    """
    attempts: list[dict[str, Any]] = []
    previous_replayed: dict[str, str] | None = None
    previous_source_hashes: dict[str, str] | None = None
    previous_binding_sha256: str | None = None
    previous_patch: Path | None = None
    for number, folder in enumerate(candidate_attempt_paths(), start=1):
        binding_path = need(folder / "binding.json", f"{rel(folder)}/binding.json")
        application_path = need(folder / "application.json",
                                f"{rel(folder)}/application.json")
        patch_path = need(folder / "candidate.patch", f"{rel(folder)}/candidate.patch")

        binding = read_json(binding_path, rel(binding_path))
        binding_keys = {
            "schema", "frozen_utc", "scope", "base_revision", "plan_sha256",
            "baseline_manifest_sha256", "workspace_lock_sha256",
            "original_isolated_source_hashes", "source_hashes", "patch_sha256",
            "limits", "root_changes", "lifetime",
        }
        if number > 1:
            binding_keys |= {"parent_binding_sha256", "transition_sha256"}
        if number >= 3:
            binding_keys.add("previous_check_receipt_sha256")
        require(isinstance(binding, dict) and set(binding) == binding_keys,
                f"{rel(binding_path)} envelope differs")
        frozen = parse_time(binding["frozen_utc"], f"{rel(binding_path)}.frozen_utc")
        require(binding["schema"] == "xlsx_0553_candidate_draft_binding_v1"
                and isinstance(binding["scope"], str) and binding["scope"].strip()
                and binding["base_revision"] == plan["revision"] == REVISION
                and binding["plan_sha256"] == sha(PLAN) == PLAN_SHA256,
                f"{rel(binding_path)} identity differs")
        if number > 1:
            require(binding["parent_binding_sha256"] == previous_binding_sha256,
                    f"{rel(binding_path)} parent binding differs")
            transition_path = need(folder / "transition.patch",
                                   f"{rel(folder)}/transition.patch")
            require(binding["transition_sha256"] == sha(transition_path),
                    f"{rel(binding_path)} transition digest differs")
        else:
            transition_path = None
        require(binding["baseline_manifest_sha256"] == sha(BASELINE / "source-manifest.json"),
                f"{rel(binding_path)} baseline binding differs")
        require(binding["workspace_lock_sha256"] == WORKSPACE_LOCK_SHA256,
                f"{rel(binding_path)} workspace lock differs")
        require(binding["limits"] == CANDIDATE_LIMITS,
                f"{rel(binding_path)} candidate limits differ")
        for field in ("root_changes", "lifetime"):
            require(isinstance(binding[field], str) and binding[field].strip(),
                    f"{rel(binding_path)}.{field} is empty")
        original = hash_map(binding["original_isolated_source_hashes"],
                            f"{rel(binding_path)}.original_isolated_source_hashes")
        source_hashes = hash_map(binding["source_hashes"],
                                 f"{rel(binding_path)}.source_hashes")
        require(binding["patch_sha256"] == sha(patch_path),
                f"{rel(binding_path)} patch digest differs")
        patch = patch_paths(patch_path, rel(patch_path))
        require(patch == set(source_hashes),
                f"{rel(binding_path)} patch/source inventory differs")
        expected_files = {"binding.json", "application.json", "candidate.patch"}
        if transition_path is not None:
            expected_files.add("transition.patch")
        preapplication_paths = sorted(folder.glob("preapplication-attempt-*.json"),
                                      key=lambda item: item.name)
        for item in preapplication_paths:
            require(item.is_file() and not item.is_symlink(),
                    f"{rel(item)} is not a regular file")
            expected_files.add(item.name)
        actual = {item.name for item in folder.iterdir()
                  if item.is_file() and not item.is_symlink()}
        require(actual == expected_files, f"{rel(folder)} raw inventory differs")
        directories = {item.name for item in folder.iterdir() if item.is_dir()}
        require(directories <= {"sources"},
                f"{rel(folder)} directory inventory differs")
        source_root = folder / "sources"
        if source_root.exists():
            require(source_root.is_dir() and not source_root.is_symlink(),
                    f"{rel(source_root)} is not a directory")
            witnesses = {
                item.relative_to(source_root).as_posix(): item
                for item in source_root.rglob("*")
                if item.is_file() and not item.is_symlink()
            }
            require(set(witnesses) == set(source_hashes),
                    f"{rel(source_root)} source inventory differs")
            for name, item in witnesses.items():
                require(sha(item) == source_hashes[name],
                        f"{rel(source_root)} source hash differs: {name}")

        replayed = replay_candidate_patch(patch_path, binding["base_revision"])
        changed = {name for name in set(baseline) | set(replayed)
                   if baseline.get(name) != replayed.get(name)}
        require(changed == set(source_hashes),
                f"{rel(binding_path)} patch replay changed-path inventory differs")
        for name, digest in source_hashes.items():
            require(replayed.get(name) == digest,
                    f"{rel(binding_path)} replay source hash differs: {name}")
        if transition_path is not None:
            require(previous_replayed is not None and previous_patch is not None,
                    f"{rel(binding_path)} has no prior patch for transition replay")
            transitioned = replay_candidate_patches(
                [previous_patch, transition_path], binding["base_revision"]
            )
            transition_names = patch_paths(transition_path, rel(transition_path))
            transition_changed = {
                name for name in set(previous_replayed) | set(transitioned)
                if previous_replayed.get(name) != transitioned.get(name)
            }
            require(transition_names == transition_changed
                    and transitioned == replayed,
                    f"{rel(binding_path)} transition replay differs")

        application = read_json(application_path, rel(application_path))
        application_keys = {
            "schema", "applied_utc", "binding_sha256", "patch_sha256",
            "source_verified",
        }
        if number == 1:
            application_keys.add("new_paths")
        elif number == 2:
            application_keys |= {"transition_sha256", "removed_owned_duplicate_tests"}
        else:
            application_keys.add("transition_sha256")
            # Later iterations may add a new source file while retaining the
            # same transition/application record shape.
            application_keys_with_new = application_keys | {"new_paths"}
        if number < 3:
            require(isinstance(application, dict) and set(application) == application_keys,
                    f"{rel(application_path)} envelope differs")
        else:
            require(isinstance(application, dict)
                    and set(application) in (application_keys, application_keys_with_new),
                    f"{rel(application_path)} envelope differs")
        applied = parse_time(application["applied_utc"],
                             f"{rel(application_path)}.applied_utc")
        require(application["schema"] == "xlsx_0553_candidate_application_v1"
                and application["binding_sha256"] == sha(binding_path)
                and application["patch_sha256"] == sha(patch_path)
                and application["source_verified"] is True
                and applied > frozen,
                f"{rel(application_path)} identity/timing differs")
        if number == 1:
            new_paths = application["new_paths"]
            require(isinstance(new_paths, list)
                    and len(new_paths) == len(set(new_paths)),
                    f"{rel(application_path)} new_paths differs")
            for name in new_paths:
                safe_relative(name, f"{rel(application_path)} new path")
                require(name in changed and name not in baseline,
                        f"{rel(application_path)} new path is not a patch addition: {name}")
            require(set(new_paths) == {name for name in changed if name not in baseline},
                    f"{rel(application_path)} new path inventory differs")
        elif number == 2:
            require(transition_path is not None
                    and application["transition_sha256"] == sha(transition_path),
                    f"{rel(application_path)} transition binding differs")
            removed = application["removed_owned_duplicate_tests"]
            require(isinstance(removed, list)
                    and removed == sorted(set(removed)),
                    f"{rel(application_path)} removed path inventory differs")
            for name in removed:
                safe_relative(name, f"{rel(application_path)} removed path")
            require(set(removed) == set(previous_source_hashes or ()) - set(source_hashes),
                    f"{rel(application_path)} removed path binding differs")
        else:
            require(transition_path is not None
                    and application["transition_sha256"] == sha(transition_path),
                    f"{rel(application_path)} transition binding differs")
            if "new_paths" in application:
                new_paths = application["new_paths"]
                require(isinstance(new_paths, list)
                        and len(new_paths) == len(set(new_paths)),
                        f"{rel(application_path)} new_paths differs")
                for name in new_paths:
                    safe_relative(name, f"{rel(application_path)} new path")
                require(set(new_paths) == set(source_hashes)
                        - set(previous_source_hashes or ()),
                        f"{rel(application_path)} new path binding differs")

        if number >= 3:
            previous_check_hash = binding["previous_check_receipt_sha256"]
            check_hash(previous_check_hash,
                       f"{rel(binding_path)}.previous_check_receipt_sha256")
            matches = [item for item in (HERE / "check-attempts").glob("*/receipt.json")
                       if sha(item) == previous_check_hash]
            require(len(matches) == 1,
                    f"{rel(binding_path)} previous check receipt is not unique")
            previous_expected = dict(previous_replayed or {})
            previous_expected["Cargo.lock"] = WORKSPACE_LOCK_SHA256
            previous_expected = dict(sorted(previous_expected.items()))
            previous_check = validate_attempt(matches[0].parent, previous_expected)
            require(previous_check["end"] <= frozen,
                    f"{rel(binding_path)} previous check postdates draft freeze")
        for preapplication_path in preapplication_paths:
            preapplication = read_json(preapplication_path, rel(preapplication_path))
            legacy_keys = {
                "recorded_utc", "status", "reason",
                "unchanged_source_binding_sha256", "old_source_check_receipt_sha256",
                "old_source_check_manifest_matches_draft02", "correction",
            }
            observation_keys = {
                "recorded_utc", "status", "observed_commands", "observation",
                "checks_started",
            }
            require(isinstance(preapplication, dict)
                    and set(preapplication) in (legacy_keys, observation_keys),
                    f"{rel(preapplication_path)} envelope differs")
            recorded = parse_time(preapplication["recorded_utc"],
                                  f"{rel(preapplication_path)}.recorded_utc")
            if set(preapplication) == legacy_keys:
                old_check_hash = preapplication["old_source_check_receipt_sha256"]
                check_hash(old_check_hash,
                           f"{rel(preapplication_path)}.old_source_check_receipt_sha256")
                old_matches = [
                    item for item in (HERE / "check-attempts").glob("*/receipt.json")
                    if sha(item) == old_check_hash
                ]
                require(len(old_matches) == 1,
                        f"{rel(preapplication_path)} old source check is not unique")
                old_check = validate_attempt(old_matches[0].parent, previous_expected)
                require(preapplication["status"] == "application_not_performed"
                        and isinstance(preapplication["reason"], str)
                        and preapplication["reason"].strip()
                        and preapplication["unchanged_source_binding_sha256"] == previous_binding_sha256
                        and preapplication["old_source_check_manifest_matches_draft02"] is True
                        and isinstance(preapplication["correction"], str)
                        and preapplication["correction"].strip()
                        and old_check["source_manifest_sha256"] == sha(
                            old_matches[0].parent / "source-manifest.json")
                        and frozen <= old_check["end"] <= recorded <= applied,
                        f"{rel(preapplication_path)} application-abort custody differs")
            else:
                commands = preapplication["observed_commands"]
                require(preapplication["status"] ==
                        "previous global compiler guard stopped before application"
                        and isinstance(commands, list) and commands
                        and all(isinstance(command, str) and command.strip()
                                for command in commands)
                        and isinstance(preapplication["observation"], str)
                        and preapplication["observation"].strip()
                        and preapplication["checks_started"] is False
                        and frozen <= recorded <= applied,
                        f"{rel(preapplication_path)} no-check observation custody differs")
        attempts.append({
            "name": folder.name,
            "path": rel(folder),
            "binding_path": rel(binding_path),
            "binding_sha256": sha(binding_path),
            "application_path": rel(application_path),
            "application_sha256": sha(application_path),
            "frozen_utc": frozen,
            "applied_utc": applied,
            "source_hashes": source_hashes,
            "original_source_hashes": original,
            "patch_sha256": sha(patch_path),
            "limits": dict(binding["limits"]),
        })
        previous_replayed = replayed
        previous_source_hashes = source_hashes
        previous_binding_sha256 = sha(binding_path)
        previous_patch = patch_path
    for previous, current in zip(attempts, attempts[1:]):
        require(current["frozen_utc"] > previous["frozen_utc"],
                "candidate attempt timestamps are not strictly increasing")
    return {"status": "pass", "count": len(attempts), "attempts": attempts,
            "latest": attempts[-1]}


def validate_candidate_binding(plan: dict[str, Any], baseline: dict[str, str],
                               candidate: dict[str, Any],
                               attempts: dict[str, Any]) -> dict[str, Any]:
    """Require an explicit final binding that selects one complete draft."""
    path = need(CANDIDATE_BINDING, "candidate-binding.json")
    value = read_json(path, "candidate-binding.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "frozen_utc", "recorded_utc", "scope", "plan_sha256", "base_revision",
        "baseline_manifest_sha256", "selected_attempt", "selected_binding_sha256",
        "application_sha256", "candidate_manifest_sha256", "candidate_patch_sha256",
        "source_hashes", "limits", "first_application", "check_manifest", "status",
    }, "candidate-binding envelope differs")
    frozen = parse_time(value["frozen_utc"], "candidate-binding.frozen_utc")
    recorded = parse_time(value["recorded_utc"], "candidate-binding.recorded_utc")
    require(value["schema"] == CANDIDATE_BINDING_SCHEMA
            and value["status"] == "frozen for measurement; admission pending"
            and isinstance(value["scope"], str) and value["scope"].strip()
            and value["plan_sha256"] == sha(PLAN) == PLAN_SHA256
            and value["base_revision"] == plan["revision"] == REVISION
            and value["baseline_manifest_sha256"] == sha(BASELINE / "source-manifest.json"),
            "candidate-binding identity differs")
    bundle_path(value["selected_attempt"], "candidate-binding.selected_attempt")
    selected = next((item for item in attempts["attempts"]
                     if item["path"] == value["selected_attempt"]), None)
    require(selected is not None,
            "candidate-binding selected attempt is not retained")
    require(value["selected_binding_sha256"] == selected["binding_sha256"]
            and value["application_sha256"] == selected["application_sha256"],
            "candidate-binding selected draft digest differs")
    require(frozen == selected["frozen_utc"],
            "candidate-binding was not frozen with the selected draft")
    candidate_manifest = candidate["manifest"]
    baseline_manifest = baseline
    delta = {name: candidate_manifest[name]
             for name in sorted(set(baseline_manifest) | set(candidate_manifest))
             if baseline_manifest.get(name) != candidate_manifest.get(name)}
    require(value["candidate_manifest_sha256"] == candidate["manifest_sha256"]
            and value["candidate_patch_sha256"] == candidate["patch_sha256"]
            and value["source_hashes"] == delta,
            "candidate-binding source custody differs")
    require(value["limits"] == CANDIDATE_LIMITS,
            "candidate-binding limits differ")
    first = value["first_application"]
    require(isinstance(first, dict) and set(first) == {
        "path", "sha256", "applied_utc", "source_verified",
    }, "candidate-binding first_application envelope differs")
    application_path = bundle_path(first["path"],
                                   "candidate-binding.first_application.path")
    require(first["path"] == selected["application_path"]
            and first["sha256"] == selected["application_sha256"]
            and application_path.is_file() and not application_path.is_symlink()
            and first["source_verified"] is True,
            "candidate-binding first application differs")
    applied = parse_time(first["applied_utc"],
                         "candidate-binding.first_application.applied_utc")
    require(applied == selected["applied_utc"] and applied > frozen,
            "candidate-binding first application timing differs")
    require(recorded >= applied,
            "candidate-binding recorded_utc predates first application")
    check_manifest = value["check_manifest"]
    require(isinstance(check_manifest, dict)
            and set(check_manifest) == {"path", "sha256"},
            "candidate-binding check_manifest envelope differs")
    check_manifest_path = bundle_path(
        check_manifest["path"], "candidate-binding.check_manifest.path"
    )
    require(check_manifest["path"].startswith("check-attempts/")
            and check_manifest_path.name == "source-manifest.json",
            "candidate-binding check_manifest path differs")
    check_manifest_value = source_manifest(check_manifest_path,
                                          "candidate-binding check_manifest")
    expected_check_manifest = dict(candidate_manifest)
    expected_check_manifest["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    expected_check_manifest = dict(sorted(expected_check_manifest.items()))
    require(check_manifest_value == expected_check_manifest
            and check_manifest["sha256"] == sha(check_manifest_path),
            "candidate-binding check manifest mapping differs")
    return {
        "status": "pass", "path": rel(path), "sha256": sha(path),
        "selected_attempt": selected["path"],
        "selected_binding_sha256": selected["binding_sha256"],
        "application_sha256": selected["application_sha256"],
        "candidate_manifest_sha256": candidate["manifest_sha256"],
        "candidate_patch_sha256": candidate["patch_sha256"],
        "source_hashes": delta, "limits": dict(value["limits"]),
        "first_application": {"path": first["path"], "applied_utc": applied},
        "check_manifest": {"path": check_manifest["path"],
                           "sha256": check_manifest["sha256"]},
    }


def validate_candidate_correctness(binding: dict[str, Any],
                                  candidate: dict[str, Any]) -> dict[str, Any]:
    """Require source-bound compact success, fallback, resource, and review rows."""
    path = need(CANDIDATE_CORRECTNESS, "candidate-correctness.json")
    value = read_json(path, "candidate-correctness.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "status", "created_utc", "candidate_binding_sha256",
        "candidate_manifest_sha256", "independent_review", "checks",
    }, "candidate-correctness envelope differs")
    created = parse_time(value["created_utc"], "candidate-correctness.created_utc")
    require(value["schema"] == CANDIDATE_CORRECTNESS_SCHEMA
            and value["status"] == "pass"
            and value["candidate_binding_sha256"] == binding["sha256"]
            and value["candidate_manifest_sha256"] == candidate["manifest_sha256"],
            "candidate-correctness identity differs")
    review = value["independent_review"]
    require(isinstance(review, dict) and set(review) == {
        "path", "sha256", "status", "reviewer",
    } and review["status"] == "pass"
            and isinstance(review["reviewer"], str) and review["reviewer"].strip(),
            "candidate-correctness independent review differs")
    review_path = bundle_path(review["path"],
                              "candidate-correctness.independent_review.path")
    require(review_path.is_file() and not review_path.is_symlink()
            and sha(review_path) == review["sha256"],
            "candidate-correctness independent review custody differs")
    rows = value["checks"]
    require(isinstance(rows, list) and rows,
            "candidate-correctness checks are missing")
    expected_manifest = dict(candidate["manifest"])
    expected_manifest["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    expected_manifest = dict(sorted(expected_manifest.items()))
    kinds: set[str] = set()
    names: set[str] = set()
    checked: list[dict[str, Any]] = []
    latest_end: dt.datetime | None = None
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and set(row) == {
            "kind", "name", "attempt", "receipt_sha256",
            "source_manifest_sha256", "exit_code", "source_stable",
        }, f"candidate-correctness check row {index} differs")
        kind = row["kind"]
        require(kind in {"compact-success", "compact-fallback", "resource"},
                f"candidate-correctness check kind is out of scope: {kind}")
        require(kind not in kinds and isinstance(row["name"], str)
                and row["name"].strip() and row["name"] not in names,
                f"candidate-correctness check identity repeats: {index}")
        kinds.add(kind)
        names.add(row["name"])
        attempt_path = bundle_path(row["attempt"],
                                   f"candidate-correctness check {index}.attempt")
        require(row["attempt"].startswith("check-attempts/"),
                f"candidate-correctness check {index} is not a check attempt")
        attempt = validate_attempt(attempt_path, expected_manifest)
        require(row["receipt_sha256"] == attempt["receipt_sha256"]
                and row["source_manifest_sha256"] == attempt["source_manifest_sha256"]
                and row["exit_code"] == 0 and attempt["exit_code"] == 0
                and row["source_stable"] is True,
                f"candidate-correctness check {index} custody differs")
        require(attempt["end"] <= created,
                f"candidate-correctness check {index} postdates correctness record")
        latest_end = attempt["end"] if latest_end is None else max(latest_end, attempt["end"])
        checked.append({"kind": kind, "name": row["name"],
                        "attempt": row["attempt"],
                        "receipt_sha256": attempt["receipt_sha256"],
                        "source_manifest_sha256": attempt["source_manifest_sha256"]})
    require({"compact-success", "compact-fallback", "resource"} <= kinds,
            "candidate-correctness is missing a direct success, fallback, or resource check")
    require(created >= binding["first_application"]["applied_utc"],
            "candidate-correctness predates first application")
    return {"status": "pass", "path": rel(path), "sha256": sha(path),
            "candidate_binding_sha256": binding["sha256"],
            "candidate_manifest_sha256": candidate["manifest_sha256"],
            "independent_review": {"path": review["path"], "sha256": review["sha256"],
                                    "reviewer": review["reviewer"]},
            "checks": checked}


def validate_candidate_custody(plan: dict[str, Any], baseline: dict[str, str],
                               candidate: dict[str, Any]) -> dict[str, Any]:
    """Gate all performance reports on explicit candidate custody evidence."""
    attempts = validate_candidate_attempts(plan, baseline)
    binding = validate_candidate_binding(plan, baseline, candidate, attempts)
    correctness = validate_candidate_correctness(binding, candidate)
    return {"status": "pass", "attempts": attempts, "binding": binding,
            "correctness": correctness}


def complete_capture_prerequisites() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    plan = validate_plan()
    validate_frozen_inputs()
    validate_supplemental_inputs()
    validate_analysis_inputs()
    validate_workspace_lock()
    base = validate_stage_source("baseline", read_json(PLAN, "plan.json"))
    candidate = validate_stage_source("candidate", read_json(PLAN, "plan.json"), base["manifest"])
    validate_candidate_custody(plan, base["manifest"], candidate)
    return plan, base, candidate


def validate_baseline_correctness() -> dict[str, Any]:
    """Validate the explicit 0553 record reusing prior final-source quality."""
    path = need(BASELINE_CORRECTNESS, "baseline-correctness.json")
    value = read_json(path, "baseline-correctness.json")
    expected_keys = {
        "schema", "status", "recorded_utc", "scope", "baseline_manifest_sha256",
        "prior_final_manifest", "prior_final_manifest_sha256", "prior_quality",
        "prior_quality_sha256", "workspace_lock_sha256", "manifest_relation", "checks",
        "test_groups", "passed", "failed", "ignored",
    }
    require(isinstance(value, dict) and set(value) == expected_keys,
            "baseline-correctness envelope differs")
    require(value["schema"] == "xlsx_0553_baseline_correctness_v1"
            and value["status"] == "pass"
            and value["scope"] == (
                "Reuse of committed 0552 final-source quality on identical 0553 baseline source; "
                "these commands were not rerun for 0553 and do not validate candidate source."
            ), "baseline-correctness identity differs")
    parse_time(value["recorded_utc"], "baseline-correctness.recorded_utc")
    baseline_path = need(BASELINE / "source-manifest.json", "baseline/source-manifest.json")
    baseline = source_manifest(baseline_path, "baseline/source-manifest.json")
    baseline_sha = sha(baseline_path)
    require(value["baseline_manifest_sha256"] == baseline_sha,
            "baseline-correctness baseline hash differs")
    prior_manifest_path = safe_repo_path(value["prior_final_manifest"],
                                         "baseline-correctness.prior_final_manifest")
    require(value["prior_final_manifest"] == "docs/performance/results/change-0552/final/source-manifest.json",
            "baseline-correctness prior final path differs")
    prior_manifest = source_manifest(prior_manifest_path, "prior final source-manifest.json")
    require(sha(prior_manifest_path) == value["prior_final_manifest_sha256"]
            and prior_manifest == baseline,
            "baseline-correctness prior final mapping differs")
    prior_quality_path = safe_repo_path(value["prior_quality"],
                                        "baseline-correctness.prior_quality")
    require(value["prior_quality"] == "docs/performance/results/change-0552/quality.json"
            and sha(prior_quality_path) == value["prior_quality_sha256"],
            "baseline-correctness prior quality binding differs")
    quality = read_json(prior_quality_path, "prior quality.json")
    require(isinstance(quality, dict) and quality.get("schema") == "xlsx_0552_quality_v1"
            and quality.get("status") == "pass" and quality.get("source_stage") == "final"
            and quality.get("source_manifest_sha256") == value["prior_final_manifest_sha256"]
            and quality.get("workspace_lock_sha256") == WORKSPACE_LOCK_SHA256,
            "baseline-correctness prior quality identity differs")
    require(value["workspace_lock_sha256"] == WORKSPACE_LOCK_SHA256
            and value["manifest_relation"] == (
                "Prior final and current baseline mappings are identical; each prior check mapping "
                "adds only the separately bound ignored Cargo.lock."
            ), "baseline-correctness lock/relation differs")
    require(value["test_groups"] == 59 and value["passed"] == 1313
            and value["failed"] == 0 and value["ignored"] == 0,
            "baseline-correctness aggregate counts differ")
    check_rows = value["checks"]
    require(isinstance(check_rows, list) and len(check_rows) == 11,
            "baseline-correctness check count differs")
    require(isinstance(quality.get("commands"), list) and len(quality["commands"]) == 11,
            "prior quality command count differs")
    expected_manifest = dict(baseline)
    expected_manifest["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    expected_manifest = dict(sorted(expected_manifest.items()))
    seen: set[str] = set()
    validated: list[dict[str, Any]] = []
    for index, row in enumerate(check_rows):
        require(isinstance(row, dict) and set(row) == {
            "name", "receipt", "receipt_sha256", "command", "exit_code",
        }, f"baseline-correctness check row {index} differs")
        name = row["name"]
        require(isinstance(name, str) and name not in seen
                and name == quality["commands"][index]["name"],
                f"baseline-correctness check name {index} differs")
        seen.add(name)
        receipt_file = safe_repo_path(row["receipt"],
                                      f"baseline-correctness check {name} receipt")
        expected_file = REPO / "docs/performance/results/change-0552/check-attempts" / (
            "final-quality-01-" + name) / "receipt.json"
        require(receipt_file == expected_file and receipt_file.name == "receipt.json",
                f"baseline-correctness {name} receipt path differs")
        receipt_dir = receipt_file.parent
        check = validate_prior_attempt(receipt_dir, expected_manifest)
        quality_row = quality["commands"][index]
        require(row["command"] == check["command"] == quality_row["command"]
                and row["receipt_sha256"] == check["receipt_sha256"] == quality_row["receipt_sha256"]
                and row["exit_code"] == check["exit_code"] == 0
                and quality_row["runner_exit_code"] == 0
                and quality_row["source_stable"] is True
                and quality_row["attempt"] == receipt_dir.relative_to(REPO).as_posix(),
                f"baseline-correctness {name} custody differs")
        validated.append({"name": name, "receipt_sha256": check["receipt_sha256"],
                          "source_manifest_sha256": check["source_manifest_sha256"]})
    require(seen == {row["name"] for row in quality["commands"]},
            "baseline-correctness check inventory is incomplete")
    return {"status": "pass", "sha256": sha(path), "baseline_manifest_sha256": baseline_sha,
            "checks": validated, "test_groups": value["test_groups"],
            "passed": value["passed"], "failed": value["failed"], "ignored": value["ignored"]}


def validate_prior_attempt(path: Path, expected_manifest: dict[str, str]) -> dict[str, Any]:
    label = rel(path)
    require(path.is_dir() and not path.is_symlink(), f"{label} is not a directory")
    receipt_path = need(path / "receipt.json", f"{label}/receipt.json")
    value = read_json(receipt_path, f"{label}/receipt.json")
    require(isinstance(value, dict) and set(value) == ATTEMPT_KEYS,
            f"{label} receipt schema differs")
    start, end = interval(value, label)
    require(value["cwd"] == str(REPO) and value["exit_code"] == 0
            and value["source_stable"] is True
            and isinstance(value["command"], list) and value["command"],
            f"{label} status/cwd differs")
    environment = value["environment"]
    require(isinstance(environment, dict) and set(environment) == ATTEMPT_ENVIRONMENT
            and environment == {
                "TMPDIR": "/home/zhuhe/litchi-goal-0552-target/tmp",
                "CARGO_TARGET_DIR": "/home/zhuhe/litchi-goal-0552-target",
                "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
                "RUSTDOCFLAGS": "-D warnings", "RUSTFLAGS": None,
                "CARGO_ENCODED_RUSTFLAGS": None, "LD_PRELOAD": None,
            }, f"{label} environment differs")
    manifest_path = need(path / "source-manifest.json", f"{label}/source-manifest.json")
    manifest = source_manifest(manifest_path, f"{label}/source-manifest.json")
    require(manifest == expected_manifest and value["source_manifest_sha256"] == sha(manifest_path),
            f"{label} source manifest differs")
    for field, expected_path in (
        ("script_sha256", REPO / "docs/performance/results/change-0552/check_attempt.py"),
        ("run_sha256", REPO / "docs/performance/results/change-0552/run.py"),
        ("plan_sha256", REPO / "docs/performance/results/change-0552/plan.json"),
    ):
        check_hash(value[field], f"{label}.{field}")
        require(value[field] == sha(expected_path), f"{label}.{field} differs")
    artifacts = value["artifacts"]
    expected_artifacts = {"source-manifest.json", "stderr", "stdout", "tracked-source.patch"}
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
            f"{label} artifact inventory differs")
    actual = {item.name for item in path.iterdir()
              if item.is_file() and not item.is_symlink()}
    require(actual == expected_artifacts | {"receipt.json"}, f"{label} raw inventory differs")
    for name, digest in artifacts.items():
        require(sha(path / name) == digest, f"{label}/{name} digest differs")
    return {"receipt_sha256": sha(receipt_path), "command": value["command"],
            "exit_code": value["exit_code"], "source_manifest_sha256": sha(manifest_path),
            "start": start, "end": end}


def validate_metrics_analysis(cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    plan, base, candidate = complete_capture_prerequisites()
    baseline_sha, candidate_sha = base["manifest_sha256"], candidate["manifest_sha256"]
    validate_stage_capture("baseline", baseline_sha, candidate_sha, cleanup=cleanup)
    validate_stage_capture("candidate", baseline_sha, candidate_sha, cleanup=cleanup)
    replay = replay_analyzer(METRICS, "xlsx_0553_metrics_analyzer",
                             METRICS_REPORT_NAMES, "xlsx_multisource_edit_metrics_0553_v1",
                             "main metrics analyzer")
    require_analyzer_pass(replay, "main metrics analyzer")
    document = replay["document"]
    require(sha(METRICS) == ANALYSIS_FILES["docs/performance/results/change-0553/analyze_metrics.py"],
            "main metrics analyzer is not the frozen revision")
    require(document.get("plan_sha256") == sha(PLAN)
            and document.get("capture_sha256") == sha(CAPTURE)
            and document.get("run_sha256") == sha(RUN)
            and document.get("stage") == "matched baseline/candidate",
            "main metrics analyzer driver binding differs")
    manifests = document.get("source_manifests")
    require(isinstance(manifests, dict) and set(manifests) == set(STAGES),
            "main metrics source-manifest inventory differs")
    for stage, expected in (("baseline", base), ("candidate", candidate)):
        require(isinstance(manifests[stage], dict)
                and manifests[stage].get("sha256") == expected["manifest_sha256"],
                f"main metrics {stage} manifest binding differs")
    binaries = document.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == set(STAGES),
            "main metrics binary stage inventory differs")
    for stage in STAGES:
        require(isinstance(binaries[stage], dict)
                and set(binaries[stage]) == {"normal", "alloc"},
                f"main metrics {stage} binary inventory differs")
    gates = document.get("main_gates")
    require(isinstance(gates, dict) and set(gates) == {
        "primary_one_percent", "one_cell_latency", "workflow_memory", "allocation",
        "correctness_identity", "all_frozen_main_gates_pass", "external_controls_required",
    } and isinstance(gates["all_frozen_main_gates_pass"], bool),
            "main metrics gate inventory differs")
    controls = gates["external_controls_required"]
    require(controls == {
        "status": "pending", "validated_here": False,
        "required": ["guard", "cap", "quality", "profile"],
        "reason": "main metrics analyzer does not own guard, cap, quality, or profile evidence",
    }, "main metrics external gate differs")
    require(isinstance(document.get("comparisons"), dict)
            and all(isinstance(document["comparisons"].get(name), list)
                    for name in ("numeric", "exact_source", "adverse"))
            and isinstance(document.get("repeat_drift"), list)
            and isinstance(document.get("repeat_drift_over_five_percent"), list),
            "main metrics diagnostic inventory differs")
    return {"status": "pass", "report": {"path": replay["path"], "sha256": replay["sha256"],
                                             "schema": replay["schema"]},
            "main_gates": gates, "source": {"baseline": base, "candidate": candidate},
            "stage_counts": {"baseline": 80, "candidate": 80}}


def validate_guard_cap_analysis(cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    plan, base, candidate = complete_capture_prerequisites()
    baseline_sha, candidate_sha = base["manifest_sha256"], candidate["manifest_sha256"]
    validate_stage_capture("baseline", baseline_sha, candidate_sha, cleanup=cleanup)
    validate_stage_capture("candidate", baseline_sha, candidate_sha, cleanup=cleanup)
    replay = replay_analyzer(GUARDS, "xlsx_0553_guard_analyzer", GUARDS_REPORT_NAMES,
                             "litchi.xlsx.guard-cap-analysis.v1", "guard/cap analyzer")
    require_analyzer_pass(replay, "guard/cap analyzer")
    document = replay["document"]
    require(sha(GUARDS) == ANALYSIS_FILES["docs/performance/results/change-0553/analyze_guards.py"],
            "guard/cap analyzer is not the frozen revision")
    require(document.get("plan_sha256") == sha(PLAN)
            and document.get("capture_sha256") == sha(CAPTURE)
            and document.get("run_sha256") == sha(RUN)
            and document.get("guarded_capture_sha256") == sha(GUARDED_CAPTURE),
            "guard/cap analyzer driver binding differs")
    expected = document.get("expected")
    require(isinstance(expected, dict)
            and expected.get("stages") == list(STAGES)
            and expected.get("guard_lanes") == list(GUARD_LANES)
            and expected.get("guard_shapes") == list(GUARD_SHAPES)
            and expected.get("guard_cases") == list(GUARD_CASES)
            and expected.get("repeats") == [1, 2]
            and expected.get("cap_sizes") == list(CAP_SIZES),
            "guard/cap expected matrix differs")
    comparison = document.get("comparison")
    require(isinstance(comparison, dict)
            and isinstance(comparison.get("guard"), dict)
            and isinstance(comparison.get("cap"), dict),
            "guard/cap comparison inventory differs")
    guard_pass = comparison["guard"].get("admission_passed")
    cap_pass = comparison["cap"].get("admission_passed")
    require(isinstance(guard_pass, bool) and isinstance(cap_pass, bool),
            "guard/cap admission fields are missing")
    require(document.get("admission_status") == ("pass" if guard_pass and cap_pass else "reject"),
            "guard/cap admission status differs")
    return {"status": "pass", "report": {"path": replay["path"], "sha256": replay["sha256"],
                                             "schema": replay["schema"]},
            "guard_admission_passed": guard_pass, "cap_admission_passed": cap_pass}


def report_field(value: Any, path: str) -> Any:
    current = value
    for part in path.split("."):
        require(isinstance(current, dict) and part in current,
                f"canonical report field is missing: {path}")
        current = current[part]
    return current


def validate_review_rows(raw: Any, reviewed: Any, source: str) -> list[dict[str, Any]]:
    require(isinstance(raw, list) and isinstance(reviewed, list)
            and len(raw) == len(reviewed),
            f"{source} review length differs")
    remaining = list(reviewed)
    result: list[dict[str, Any]] = []
    for index, original in enumerate(raw):
        require(isinstance(original, dict), f"{source} raw row {index} is malformed")
        matches = [
            (position, row) for position, row in enumerate(remaining)
            if isinstance(row, dict) and row.get("original") == original
        ]
        require(len(matches) == 1,
                f"{source} raw row {index} is not individually reviewed")
        position, row = matches[0]
        if "source" in row:
            require(row["source"] == source,
                    f"{source} row {index} source label differs")
        for field in ("id", "classification", "interpretation", "disposition"):
            require(isinstance(row.get(field), str) and row[field].strip(),
                    f"{source} row {index}.{field} is empty")
        result.append(row)
        remaining.pop(position)
    require(not remaining, f"{source} review has extra rows")
    return result


def validate_adverse_review(metrics: dict[str, Any], metrics_sha: str,
                            guards: dict[str, Any], guards_sha: str) -> dict[str, Any]:
    """Require one-for-one disposition of every adverse and drift diagnostic."""
    # CLI components pass validated report references; the decision API passes
    # complete documents. Resolve references with the same digest custody.
    documents = []
    for report, digest in ((metrics, metrics_sha), (guards, guards_sha)):
        if "report" in report:
            reference = report["report"]
            source = need(HERE / reference["path"], "adverse review source report")
            require(reference["sha256"] == digest and sha(source) == digest,
                    "adverse review source report digest differs")
            report = read_json(source, rel(source))
        documents.append(report)
    metrics, guards = documents
    path = need(ADVERSE_REVIEW, "adverse-review.json")
    value = read_json(path, "adverse-review.json")
    required = {
        "schema", "status", "complete", "all_diagnostic_rows_retained",
        "adoption_allowed", "comparison_sha256", "guard_analysis_sha256", "groups",
    }
    optional = {"counts", "source_coverage"}
    require(isinstance(value, dict) and required <= set(value)
            and set(value) <= required | optional,
            "adverse-review envelope differs")
    require(value["schema"] == ADVERSE_REVIEW_SCHEMA
            and value["status"] == "complete"
            and value["complete"] is True
            and value["all_diagnostic_rows_retained"] is True
            and isinstance(value["adoption_allowed"], bool)
            and value["comparison_sha256"] == metrics_sha
            and value["guard_analysis_sha256"] == guards_sha,
            "adverse-review identity differs")
    raw = {
        "metrics_adverse": report_field(metrics, "comparisons.adverse"),
        "metrics_drift": report_field(metrics, "repeat_drift_over_five_percent"),
        "guard_adverse": report_field(
            guards, "comparison.guard.adverse_flags_over_five_percent"),
        "guard_drift": report_field(
            guards, "comparison.guard.same_build_drift_over_five_percent"),
        "cap_adverse": report_field(
            guards, "comparison.cap.adverse_flags_over_five_percent"),
        "cap_drift": report_field(
            guards, "comparison.cap.same_build_drift_over_five_percent"),
    }
    require(all(isinstance(rows, list) for rows in raw.values()),
            "adverse-review source arrays are malformed")
    groups = value["groups"]
    require(isinstance(groups, list) and len(groups) == len(REVIEW_SOURCES),
            "adverse-review group inventory differs")
    by_source: dict[str, Any] = {}
    for index, group in enumerate(groups):
        require(isinstance(group, dict) and set(group) == {"source", "rows"},
                f"adverse-review group {index} differs")
        source = group["source"]
        require(isinstance(source, str) and source.strip()
                and source not in by_source,
                f"adverse-review group {index}.source differs")
        require(isinstance(group["rows"], list),
                f"adverse-review group {index}.rows differs")
        by_source[source] = group["rows"]
    require(set(by_source) == set(REVIEW_SOURCES.values()),
            "adverse-review source inventory differs")
    retained: dict[str, list[dict[str, Any]]] = {}
    for key, source in REVIEW_SOURCES.items():
        retained[key] = validate_review_rows(raw[key], by_source[source], source)
    expected_count = sum(len(rows) for rows in raw.values())
    if "counts" in value:
        counts = value["counts"]
        require(isinstance(counts, dict)
                and counts.get("reviewed_flags") == expected_count,
                "adverse-review counts differ")
    if "source_coverage" in value:
        coverage = value["source_coverage"]
        require(isinstance(coverage, list)
                and all(isinstance(row, dict) and set(row) == {"source", "count"}
                        for row in coverage),
                "adverse-review source_coverage differs")
        expected_coverage = [
            {"source": REVIEW_SOURCES[key], "count": len(raw[key])}
            for key in REVIEW_SOURCES
        ]
        require(sorted(coverage, key=lambda row: (row.get("source", ""),
                                                   row.get("count", -1)))
                == sorted(expected_coverage,
                          key=lambda row: (row["source"], row["count"])),
                "adverse-review source_coverage differs")
    return {"status": "pass", "path": rel(path), "sha256": sha(path),
            "schema": ADVERSE_REVIEW_SCHEMA, "adoption_allowed": value["adoption_allowed"],
            "reviewed_flags": expected_count,
            "groups": [{"source": REVIEW_SOURCES[key], "rows": retained[key]}
                       for key in REVIEW_SOURCES]}


def profile_pilot_gates(metrics: dict[str, Any], guards: dict[str, Any]) -> dict[str, bool]:
    main = metrics["main_gates"].get("all_frozen_main_gates_pass")
    guard = guards.get("guard_admission_passed")
    cap = guards.get("cap_admission_passed")
    require(isinstance(main, bool) and isinstance(guard, bool) and isinstance(cap, bool),
            "profile pilot gates are missing")
    return {"main": main, "guard": guard, "cap": cap}


def validate_profiles(metrics: dict[str, Any] | None = None,
                      guards: dict[str, Any] | None = None,
                      *, pilot_expected: bool | None = None,
                      cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    if metrics is None:
        metrics = validate_metrics_analysis(cleanup=cleanup)
    if guards is None:
        guards = validate_guard_cap_analysis(cleanup=cleanup)
    paths = [HERE / name for name in PROFILE_DECISION_NAMES if (HERE / name).exists()]
    require(len(paths) <= 1, "multiple profile decisions are retained")
    decision_path = paths[0] if paths else None
    gates = profile_pilot_gates(metrics, guards)
    pilot = all(gates.values())
    if pilot_expected is not None:
        require(isinstance(pilot_expected, bool) and pilot_expected is pilot,
                "profile pilot expectation differs from frozen gates")
    if decision_path is None:
        raise IncompleteError("profile decision is missing")
    require(decision_path.is_file() and not decision_path.is_symlink(),
            f"{rel(decision_path)} is not a regular file")
    value = read_json(decision_path, rel(decision_path))
    require(isinstance(value, dict) and set(value) == {
        "schema", "status", "scope", "pilot_passed", "profile_required",
        "profile_gate_passed", "main_analysis_sha256", "pilot_gates",
        "profile_rows", "reason",
    }, f"{rel(decision_path)} profile decision envelope differs")
    require(value["schema"] == PROFILE_DECISION_SCHEMA
            and isinstance(value["scope"], str) and value["scope"].strip()
            and isinstance(value["reason"], str) and value["reason"].strip(),
            f"{rel(decision_path)} profile decision identity differs")
    require(isinstance(value["pilot_passed"], bool)
            and isinstance(value["profile_required"], bool)
            and isinstance(value["profile_gate_passed"], bool)
            and value["pilot_gates"] == gates
            and value["pilot_passed"] is pilot
            and value["profile_required"] is pilot,
            f"{rel(decision_path)} profile pilot binding differs")
    metrics_sha = metrics["report"]["sha256"]
    require(value["main_analysis_sha256"] == metrics_sha,
            f"{rel(decision_path)} main analysis binding differs")
    require(isinstance(value["profile_rows"], list),
            f"{rel(decision_path)} profile_rows differs")
    if not pilot:
        require(value["status"] == "skipped" and value["profile_gate_passed"] is True
                and value["profile_rows"] == [],
                f"{rel(decision_path)} failed pilot is not an explicit skip")
        receipts = [path for stage in STAGES for path in (HERE / stage).glob("profile-*.receipt.json")]
        require(not receipts, "profile captures exist although the pilot did not pass")
        return {"status": "pass", "required": False, "pilot_passed": False,
                "gate_passed": True, "decision": rel(decision_path),
                "decision_sha256": sha(decision_path), "pilot_gates": gates,
                "profile_rows": []}
    require(value["status"] in {"pass", "failed", "reject", "rejected"}
            and value["profile_rows"], f"{rel(decision_path)} required profiles are missing")
    # A passing pilot makes all sixteen profile receipts mandatory.  Detailed
    # callgrind/result identity is checked here; the main analyzer owns the
    # benchmark report schema and is replayed above.
    owned = validate_owned_binaries(cleanup=cleanup)
    rows: list[dict[str, Any]] = []
    by_key: dict[tuple[str, int, str], int] = {}
    for stage in STAGES:
        binary = owned["binaries"][stage]["normal"]
        binary_path = Path(binary["path"])
        binary_sha = binary["sha256"]
        for job in expected_profile_jobs(stage):
            folder = HERE / stage
            stem = job["name"]
            receipt_path = need(folder / f"{stem}.receipt.json", f"{stage}/{stem}.receipt.json")
            receipt = read_json(receipt_path, rel(receipt_path))
            start, end = validate_receipt_common(receipt, rel(receipt_path))
            callgrind = folder / f"{stem}.callgrind"
            command = ["taskset", "-c", str(CPU), "valgrind", "--tool=callgrind",
                       "--vgdb=no", "--vgdb-prefix=" + str(TARGET / "tmp/vgdb"),
                       "--collect-atstart=no", "--toggle-collect=" + PROFILE_OWNER,
                       "--zero-before=" + PROFILE_OWNER, "--dump-after=" + PROFILE_OWNER,
                       "--callgrind-out-file=" + str(callgrind), str(binary_path),
                       "--case", PROFILE_CASE, "--xlsx-cell-crud-shape", job["shape"],
                       "--samples", "1", "--warmup", "0", "--json",
                       str(folder / f"{stem}.json"), "--corpus-manifest",
                       str(folder / f"{stem}.catalog.json")]
            require(receipt["exit_code"] == 0 and receipt["binary_sha256"] == binary_sha
                    and receipt["execution_stage"] == "candidate"
                    and receipt["execution_manifest_sha256"] == metrics["source"]["candidate"]["manifest_sha256"]
                    and receipt["source_manifest_sha256"] == metrics["source"][stage]["manifest_sha256"]
                    and receipt["script_sha256"] == sha(RUN)
                    and receipt["plan_sha256"] == sha(PLAN)
                    and receipt["command"] == command,
                    f"{rel(receipt_path)} binding differs")
            validate_artifacts(folder, stem, receipt,
                               {f"{stem}.json", f"{stem}.catalog.json", f"{stem}.stdout",
                                f"{stem}.stderr", f"{stem}.host.json", f"{stem}.callgrind"},
                               rel(receipt_path))
            text = read_text(callgrind, rel(callgrind))
            require("events: Ir" in text and PROFILE_OWNER in text,
                    f"{rel(callgrind)} lacks the profiled owner")
            summaries = [int(match.group(1)) for line in text.splitlines()
                         if (match := re.fullmatch(r"\s*summary:\s*([0-9]+)\s*", line))]
            require(len(summaries) == 1 and summaries[0] > 0,
                    f"{rel(callgrind)} Ir summary differs")
            by_key[(stage, job["repeat"], job["shape"])] = summaries[0]
            rows.append({"stage": stage, "repeat": job["repeat"], "shape": job["shape"],
                         "instruction_references": summaries[0],
                         "receipt_sha256": sha(receipt_path), "callgrind_sha256": sha(callgrind),
                         "start": start, "end": end})
    ir_rows = []
    for repeat in (1, 2):
        for shape in SHAPES:
            before = by_key[("baseline", repeat, shape)]
            after = by_key[("candidate", repeat, shape)]
            ir_rows.append({"repeat": repeat, "shape": shape, "baseline": before,
                            "candidate": after, "passed": after < before})
    passed = all(row["passed"] for row in ir_rows)
    require(value["profile_gate_passed"] is passed,
            f"{rel(decision_path)} profile gate differs from Ir replay")
    return {"status": "pass", "required": True, "pilot_passed": True,
            "gate_passed": passed, "decision": rel(decision_path),
            "decision_sha256": sha(decision_path), "pilot_gates": gates,
            "profile_rows": ir_rows, "captures": rows}


def validate_attempt(path: Path, expected_manifest: dict[str, str]) -> dict[str, Any]:
    """Validate one 0553 quality check-attempt directory."""
    label = rel(path)
    require(path.is_dir() and not path.is_symlink(), f"{label} is not a directory")
    receipt_path = need(path / "receipt.json", f"{label}/receipt.json")
    value = read_json(receipt_path, f"{label}/receipt.json")
    require(isinstance(value, dict) and set(value) == ATTEMPT_KEYS,
            f"{label} receipt schema differs")
    start, end = interval(value, label)
    require(value["cwd"] == str(REPO) and value["source_stable"] is True
            and isinstance(value["exit_code"], int) and not isinstance(value["exit_code"], bool),
            f"{label} receipt status differs")
    environment = value["environment"]
    require(isinstance(environment, dict) and set(environment) == ATTEMPT_ENVIRONMENT
            and environment["TMPDIR"] == str(TARGET / "tmp")
            and environment["CARGO_TARGET_DIR"] == str(TARGET)
            and environment["CARGO_BUILD_JOBS"] == "2"
            and environment["CARGO_INCREMENTAL"] == "0"
            and environment["RUSTDOCFLAGS"] == "-D warnings"
            and all(environment[name] is None for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD")),
            f"{label} environment differs")
    manifest_path = need(path / "source-manifest.json", f"{label}/source-manifest.json")
    manifest = source_manifest(manifest_path, f"{label}/source-manifest.json")
    require(manifest == expected_manifest and value["source_manifest_sha256"] == sha(manifest_path)
            and value["script_sha256"] == sha(CHECK_ATTEMPT)
            and value["run_sha256"] == sha(RUN)
            and value["plan_sha256"] == sha(PLAN), f"{label} source binding differs")
    expected_artifacts = {"source-manifest.json", "stderr", "stdout", "tracked-source.patch"}
    require(set(value["artifacts"]) == expected_artifacts, f"{label} artifacts differ")
    actual = {item.name for item in path.iterdir()
              if item.is_file() and not item.is_symlink()}
    require(actual == expected_artifacts | {"receipt.json"}, f"{label} raw inventory differs")
    for name, digest in value["artifacts"].items():
        require(sha(path / name) == digest, f"{label}/{name} digest differs")
    return {"path": label, "receipt_sha256": sha(receipt_path), "command": value["command"],
            "exit_code": value["exit_code"], "source_manifest_sha256": sha(manifest_path),
            "start": start, "end": end}


def validate_quality() -> dict[str, Any]:
    plan = validate_quality_plan()
    root = HERE / "quality-attempts"
    if not root.exists():
        raise IncompleteError("quality-attempts directory is missing")
    require(root.is_dir() and not root.is_symlink(), "quality-attempts is not a directory")
    attempts: list[dict[str, Any]] = []
    for path in sorted(root.iterdir(), key=lambda item: item.name):
        require(re.fullmatch(r"[A-Za-z0-9_-]+", path.name) is not None,
                f"quality attempt label is unsafe: {path.name}")
        inputs_path = need(path / "inputs.json", f"{rel(path)}/inputs.json")
        result_path = need(path / "result.json", f"{rel(path)}/result.json")
        inputs = read_json(inputs_path, rel(inputs_path))
        require(isinstance(inputs, dict) and set(inputs) == {
            "schema", "created_utc", "source_stage", "source_manifest_sha256",
            "workspace_lock_sha256", "quality_plan_sha256", "supplemental_inputs_sha256",
            "scripts", "commands",
        }, f"{rel(path)}/inputs.json differs")
        require(inputs["schema"] == "xlsx_0553_quality_inputs_v1"
                and inputs["source_stage"] in ("candidate", "final"),
                f"{rel(path)}/inputs identity differs")
        created = parse_time(inputs["created_utc"], f"{rel(path)}/inputs.created_utc")
        source_stage = inputs["source_stage"]
        baseline_for_quality = source_manifest(
            BASELINE / "source-manifest.json", "baseline/source-manifest.json"
        )
        candidate_for_quality = None
        if source_stage == "final":
            candidate_for_quality = source_manifest(
                CANDIDATE / "source-manifest.json", "candidate/source-manifest.json"
            )
        stage = validate_stage_source(source_stage, validate_plan(), baseline_for_quality,
                                      candidate_for_quality)
        expected_manifest = dict(stage["manifest"])
        expected_manifest["Cargo.lock"] = WORKSPACE_LOCK_SHA256
        expected_manifest = dict(sorted(expected_manifest.items()))
        require(inputs["source_manifest_sha256"] == stage["manifest_sha256"]
                and inputs["workspace_lock_sha256"] == WORKSPACE_LOCK_SHA256
                and inputs["quality_plan_sha256"] == sha(QUALITY_PLAN)
                and inputs["supplemental_inputs_sha256"] == sha(SUPPLEMENTAL),
                f"{rel(path)}/inputs binding differs")
        require(inputs["scripts"] == {"quality.py": sha(QUALITY),
                                       "check_attempt.py": sha(CHECK_ATTEMPT),
                                       "run.py": sha(RUN)}
                and inputs["commands"] == plan["commands"],
                f"{rel(path)}/inputs script/command binding differs")
        value = read_json(result_path, rel(result_path))
        require(isinstance(value, dict) and set(value) == {
            "schema", "status", "source_stage", "source_manifest_sha256",
            "workspace_lock_sha256", "quality_plan_sha256", "inputs_path", "inputs_sha256",
            "completed_utc", "commands",
        }, f"{rel(path)}/result.json differs")
        completed = parse_time(value["completed_utc"], f"{rel(path)}/result.completed_utc")
        require(value["schema"] == "xlsx_0553_quality_v1"
                and value["status"] in ("pass", "failed")
                and value["source_stage"] == source_stage
                and value["source_manifest_sha256"] == inputs["source_manifest_sha256"]
                and value["workspace_lock_sha256"] == WORKSPACE_LOCK_SHA256
                and value["quality_plan_sha256"] == sha(QUALITY_PLAN)
                and value["inputs_path"] == inputs_path.relative_to(REPO).as_posix()
                and value["inputs_sha256"] == sha(inputs_path)
                and completed >= created, f"{rel(path)}/result binding differs")
        rows = value["commands"]
        require(isinstance(rows, list) and len(rows) <= len(plan["commands"]),
                f"{rel(path)}/result command rows differ")
        checked_rows = []
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and set(row) == {
                "name", "command", "attempt", "receipt_sha256", "exit_code",
                "runner_exit_code", "source_stable",
            }, f"{rel(path)} command row {index} differs")
            name = list(plan["commands"])[index]
            require(row["name"] == name and row["command"] == plan["commands"][name]
                    and row["source_stable"] is True,
                    f"{rel(path)} command row {index} binding differs")
            attempt_path = safe_repo_path(row["attempt"], f"{rel(path)} attempt {index}")
            expected_attempt_manifest = expected_manifest
            attempt = validate_attempt(attempt_path, expected_attempt_manifest)
            require(row["receipt_sha256"] == attempt["receipt_sha256"]
                    and row["command"] == attempt["command"]
                    and row["exit_code"] == attempt["exit_code"]
                    and row["runner_exit_code"] == row["exit_code"]
                    and created <= attempt["start"] <= attempt["end"] <= completed,
                    f"{rel(path)} command row {index} custody differs")
            checked_rows.append(row)
        if value["status"] == "pass":
            require(len(rows) == len(plan["commands"])
                    and all(row["exit_code"] == 0 and row["runner_exit_code"] == 0
                            and row["source_stable"] for row in rows),
                    f"{rel(path)} claims pass with a failed command")
        require({item.name for item in path.iterdir()
                 if item.is_file() and not item.is_symlink()} == {"inputs.json", "result.json"},
                f"{rel(path)} raw inventory differs")
        attempts.append({"path": rel(path), "result_sha256": sha(result_path),
                         "status": value["status"], "source_stage": source_stage,
                         "source_manifest_sha256": stage["manifest_sha256"],
                         "commands": checked_rows})
    passing = [item for item in attempts if item["status"] == "pass" and item["source_stage"] == "final"]
    if not passing:
        raise IncompleteError("no passing final quality attempt is retained")
    canonical = need(HERE / "quality.json", "quality.json")
    selected = [item for item in passing
                if canonical.read_bytes() == (HERE / item["path"] / "result.json").read_bytes()]
    require(len(selected) == 1, "quality.json is not an exact passing final attempt")
    return {"status": "pass", "canonical_sha256": sha(canonical),
            "selected_attempt": selected[0]["path"], "attempts": attempts}


def decision_path() -> Path:
    paths = [HERE / name for name in DECISION_NAMES if (HERE / name).exists()]
    if not paths:
        raise IncompleteError("final disposition document is missing")
    require(len(paths) == 1, "multiple final disposition documents are retained")
    return need(paths[0], rel(paths[0]))


def validate_final_source(disposition: str, candidate: dict[str, str],
                          candidate_sha: str) -> dict[str, Any]:
    """Bind the final checkout to the candidate or an exact baseline restore."""
    require(disposition in ("accepted", "rejected"),
            "final source disposition is invalid")
    plan = validate_plan()
    baseline = source_manifest(BASELINE / "source-manifest.json",
                               "baseline/source-manifest.json")
    candidate_path = need(CANDIDATE / "source-manifest.json",
                          "candidate/source-manifest.json")
    candidate_observed = source_manifest(candidate_path,
                                         "candidate/source-manifest.json")
    require(candidate == candidate_observed and candidate_sha == sha(candidate_path),
            "candidate source argument is not the retained stage manifest")
    final = validate_stage_source("final", plan, baseline, candidate)
    expected = candidate if disposition == "accepted" else baseline
    require(final["manifest"] == expected, "final source does not match disposition")
    require(current_source_manifest() == final["manifest"],
            "live source does not match final disposition")
    return {"status": "pass", "stage": "final",
            "manifest_sha256": final["manifest_sha256"],
            "entries": final["entries"], "patch_sha256": final["patch_sha256"]}


def validate_disposition(metrics: dict[str, Any], guards: dict[str, Any],
                         profiles: dict[str, Any], quality: dict[str, Any],
                         review: dict[str, Any] | None = None) -> dict[str, Any]:
    if review is None:
        review = validate_adverse_review(
            metrics, metrics["report"]["sha256"], guards, guards["report"]["sha256"]
        )
    path = decision_path()
    value = read_json(path, rel(path))
    require(isinstance(value, dict), f"{rel(path)} is not an object")
    disposition = value.get("disposition")
    require(disposition in ("accepted", "rejected"), f"{rel(path)} disposition differs")
    observed = value.get("observed_utc", value.get("completed_utc", value.get("utc")))
    parse_time(observed, f"{rel(path)} timestamp")
    require(isinstance(value.get("scope"), str) and value["scope"].strip(),
            f"{rel(path)} scope is missing")
    main_gate = bool(metrics["main_gates"]["all_frozen_main_gates_pass"])
    guard_gate = bool(guards["guard_admission_passed"])
    cap_gate = bool(guards["cap_admission_passed"])
    profile_gate = bool(profiles["gate_passed"])
    quality_gate = quality["status"] == "pass"
    review_gate = bool(review["adoption_allowed"])
    adoption = (main_gate and guard_gate and cap_gate and profile_gate
                and quality_gate and review_gate)
    require(isinstance(value.get("adoption_allowed"), bool)
            and value["adoption_allowed"] is adoption
            and disposition == ("accepted" if adoption else "rejected"),
            f"{rel(path)} adoption differs from independent gates")
    for name, expected in (("main_gate", main_gate), ("native_primary_gate", main_gate),
                           ("guard_gate", guard_gate), ("cap_gate", cap_gate),
                           ("profile_gate", profile_gate), ("quality_gate", quality_gate)):
        if name in value:
            require(isinstance(value[name], bool) and value[name] is expected,
                    f"{rel(path)} {name} differs")
    for name, expected in (("metrics_analysis_sha256", metrics["report"]["sha256"]),
                           ("main_analysis_sha256", metrics["report"]["sha256"]),
                           ("guard_analysis_sha256", guards["report"]["sha256"]),
                           ("quality_sha256", quality["canonical_sha256"]),
                           ("quality_summary_sha256", quality["canonical_sha256"])):
        if name in value:
            require(value[name] == expected, f"{rel(path)} {name} differs")
    if "adverse_review_gate" in value:
        require(isinstance(value["adverse_review_gate"], bool)
                and value["adverse_review_gate"] is review_gate,
                f"{rel(path)} adverse_review_gate differs")
    if "adverse_review_sha256" in value:
        require(value["adverse_review_sha256"] == review["sha256"],
                f"{rel(path)} adverse_review_sha256 differs")
    candidate_stage = metrics["source"]["candidate"]
    final = validate_final_source(disposition, candidate_stage["manifest"],
                                  candidate_stage["manifest_sha256"])
    require(value.get("source_manifest_sha256") == final["manifest_sha256"],
            f"{rel(path)} final source hash differs")
    return {"status": "pass", "decision": rel(path), "decision_sha256": sha(path),
            "disposition": disposition, "adoption_allowed": adoption,
            "gates": {"main": main_gate, "guard": guard_gate, "cap": cap_gate,
                      "profile": profile_gate, "quality": quality_gate,
                      "adverse_review": review_gate},
            "final_source": final, "adverse_review": review}


def validate_documentation() -> dict[str, Any]:
    """Validate the manifest for completed-work documentation outside the bundle."""
    path = need(DOCUMENTATION_MANIFEST, "documentation-manifest.json")
    value = read_json(path, "documentation-manifest.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "scope", "files",
    }, "documentation-manifest envelope differs")
    require(value["schema"] == DOCUMENTATION_MANIFEST_SCHEMA
            and value["scope"] == DOCUMENTATION_SCOPE,
            "documentation-manifest identity differs")
    files = value["files"]
    require(isinstance(files, dict) and files,
            "documentation-manifest file map is empty")
    require(list(files) == sorted(files),
            "documentation-manifest file map is not sorted")
    bundle_prefix = HERE.relative_to(REPO).as_posix()
    for name, digest in files.items():
        safe_relative(name, "documentation-manifest path")
        require(name.startswith("docs/")
                and name != bundle_prefix
                and not name.startswith(bundle_prefix + "/"),
                f"documentation-manifest path is not external: {name}")
        check_hash(digest, f"documentation-manifest hash {name}")
        item = safe_repo_path(name, f"documentation-manifest path {name}")
        require(not item.is_symlink() and item.is_file(),
                f"documentation-manifest file is missing or unsafe: {name}")
        require(sha(item) == digest,
                f"documentation-manifest digest differs: {name}")
    require(DOCUMENTATION_REQUIRED in files,
            f"documentation-manifest omits {DOCUMENTATION_REQUIRED}")
    return {"status": "pass", "path": rel(path), "sha256": sha(path),
            "files": dict(sorted(files.items()))}


def validate_precleanup_record(value: dict[str, Any],
                              cleanup: dict[str, Any]) -> dict[str, Any]:
    """Validate the retained pre-deletion report against cleanup hashes."""
    require(isinstance(value, dict) and set(value) == {
        "schema", "status", "scope", "result",
    }, "precleanup-verification envelope differs")
    require(value["schema"] == "litchi.xlsx.verification.0553.v1"
            and value["status"] == "pass" and value["scope"] == "precleanup",
            "precleanup-verification identity differs")
    result = value["result"]
    require(isinstance(result, dict) and set(result) == {
        "status", "phase", "target_present", "observed_utc", "campaign",
    } and result["status"] == "pass" and result["phase"] == "precleanup"
            and result["target_present"] is True,
            "precleanup-verification result differs")
    precleanup_time = parse_time(result["observed_utc"],
                                 "precleanup-verification.observed_utc")
    cleanup_time = parse_time(cleanup["observed_utc"], "cleanup.observed_utc")
    require(precleanup_time <= cleanup_time,
            "precleanup verification postdates cleanup")
    campaign = result["campaign"]
    require(isinstance(campaign, dict) and set(campaign) == {
        "metrics", "guards", "profiles", "quality", "adverse_review",
        "decision", "owned_binaries", "documentation",
    }, "precleanup-verification campaign inventory differs")
    owned = campaign["owned_binaries"]
    require(isinstance(owned, dict) and owned.get("status") == "pass"
            and owned.get("count") == len(STAGES) * len(BINARY_KINDS),
            "precleanup-verification binary summary differs")
    binaries = owned.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == set(STAGES),
            "precleanup-verification stage binary inventory differs")
    descriptor_hashes = cleanup["binary_descriptor_sha256_by_kind"]
    for stage in STAGES:
        require(isinstance(binaries[stage], dict)
                and set(binaries[stage]) == set(BINARY_KINDS),
                f"precleanup-verification {stage} binary inventory differs")
        for kind in BINARY_KINDS:
            key = f"{stage}/{kind}"
            row = binaries[stage][kind]
            require(isinstance(row, dict)
                    and row.get("path") == str(SCRATCH_ROOT / stage / descriptor_spec(kind)[1])
                    and row.get("sha256") == cleanup["binary_sha256_by_kind"][key]
                    and row.get("descriptor_sha256") == descriptor_hashes[key],
                    f"precleanup-verification binary custody differs: {key}")
            check_hash(row.get("sha256"), f"precleanup-verification {key}.sha256")
            check_hash(row.get("descriptor_sha256"),
                       f"precleanup-verification {key}.descriptor_sha256")
            start = parse_time(row.get("start"), f"precleanup-verification {key}.start")
            end = parse_time(row.get("end"), f"precleanup-verification {key}.end")
            require(end > start, f"precleanup-verification {key} interval is inverted")
    return {"status": "pass", "sha256": sha(PRECLEANUP),
            "observed_utc": result["observed_utc"], "count": owned["count"]}


def validate_cleanup() -> dict[str, Any]:
    """Validate post-deletion custody for the owned build target."""
    path = need(CLEANUP, "cleanup.json")
    value = read_json(path, "cleanup.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "observed_utc", "plan_sha256", "target", "removed",
        "owned_paths_absent", "accessible_process_references",
        "process_reference_scope", "python_cache_absent",
        "binary_sha256_by_kind", "binary_descriptor_sha256_by_kind",
        "precleanup_verification_sha256", "scope",
    }, "cleanup.json envelope differs")
    require(value["schema"] == CLEANUP_SCHEMA
            and value["plan_sha256"] == sha(PLAN)
            and value["target"] == TARGET_STRING
            and value["removed"] == [TARGET_STRING]
            and value["owned_paths_absent"] is True
            and value["accessible_process_references"] == []
            and value["process_reference_scope"] == PROCESS_REFERENCE_SCOPE
            and value["python_cache_absent"] is True
            and value["scope"] == CLEANUP_SCOPE,
            "cleanup target custody differs")
    parse_time(value["observed_utc"], "cleanup.observed_utc")
    precleanup_path = need(PRECLEANUP, "precleanup-verification.json")
    require(value["precleanup_verification_sha256"] == sha(precleanup_path),
            "cleanup precleanup-verification binding differs")
    precleanup = read_json(precleanup_path, rel(precleanup_path))
    require(not os.path.lexists(TARGET),
            "cleanup claims target absent but owned path remains")
    require(not any("__pycache__" in item.parts for item in HERE.rglob("*")),
            "Python bytecode cache remains after cleanup")

    hashes = value["binary_sha256_by_kind"]
    descriptor_hashes = value["binary_descriptor_sha256_by_kind"]
    require(isinstance(hashes, dict) and set(hashes) == expected_binary_keys()
            and isinstance(descriptor_hashes, dict)
            and set(descriptor_hashes) == expected_binary_keys(),
            "cleanup binary custody inventory differs")
    validate_precleanup_record(precleanup, value)
    checked: list[dict[str, str]] = []
    for stage in STAGES:
        manifest, manifest_sha = stage_manifest(stage)
        for kind in BINARY_KINDS:
            descriptor = validate_binary_descriptor(stage, kind, manifest_sha)
            key = f"{stage}/{kind}"
            require(cleanup_binary_digest(value, stage, kind) == descriptor["sha256"],
                    f"cleanup binary digest differs: {key}")
            descriptor_digest = check_hash(
                descriptor_hashes[key],
                f"cleanup.binary_descriptor_sha256_by_kind.{key}",
            )
            require(descriptor_digest == descriptor["descriptor_sha256"],
                    f"cleanup descriptor digest differs: {key}")
            checked.append({"key": key, "sha256": descriptor["sha256"],
                            "descriptor_sha256": descriptor_digest})
    require(len(checked) == len(STAGES) * len(BINARY_KINDS),
            "cleanup binary custody is not exactly ten entries")
    return {"status": "pass", "path": rel(path), "sha256": sha(path),
            "owned_paths_absent": True, "binary_custody": checked}


def sealed_inventory() -> dict[str, str]:
    """Return every regular bundle file except the seal, rejecting symlinks."""
    result: dict[str, str] = {}
    for item in HERE.rglob("*"):
        require(not item.is_symlink(),
                f"sealed bundle contains a symlink: {rel(item)}")
        require("__pycache__" not in item.parts,
                f"sealed bundle contains Python cache: {rel(item)}")
        if item.is_dir():
            continue
        require(item.is_file(), f"sealed bundle contains a non-file: {rel(item)}")
        if item == SEAL:
            continue
        name = item.relative_to(HERE).as_posix()
        result[name] = sha(item)
    require("verify.py" in result and "verifier-schema.md" in result,
            "sealed inventory omits the verifier or schema")
    return dict(sorted(result.items()))


def validate_seal() -> dict[str, Any]:
    """Validate the recursive SHA256SUMS inventory after terminal cleanup."""
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    text = read_text(path, "SHA256SUMS")
    lines = text.splitlines()
    require(lines and text.endswith("\n"), "SHA256SUMS is empty or lacks a final newline")
    for line in lines:
        fields = line.split("  ", 1)
        require(len(fields) == 2 and fields[0] and fields[1],
                "SHA256SUMS line differs")
        check_hash(fields[0], "SHA256SUMS digest")
        safe_relative(fields[1], "SHA256SUMS path")
        require(fields[1] != "SHA256SUMS" and fields[1] not in expected,
                "SHA256SUMS inventory is unsafe or duplicated")
        expected[fields[1]] = fields[0]
    require(list(expected) == sorted(expected),
            "SHA256SUMS paths are not sorted")
    actual = sealed_inventory()
    require(expected == actual, "SHA256SUMS inventory differs")
    return {"status": "pass", "path": rel(path), "sha256": sha(path),
            "entries": len(expected)}


def validate_campaign(cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate the complete campaign with optional post-cleanup binary custody."""
    validate_adr()
    binaries = validate_owned_binaries(cleanup=cleanup)
    metrics = validate_metrics_analysis(cleanup=cleanup)
    guards = validate_guard_cap_analysis(cleanup=cleanup)
    profiles = validate_profiles(metrics, guards, cleanup=cleanup)
    quality = validate_quality()
    review = validate_adverse_review(metrics, metrics["report"]["sha256"],
                                     guards, guards["report"]["sha256"])
    decision = validate_disposition(metrics, guards, profiles, quality, review)
    documentation = validate_documentation()
    return {
        "metrics": metrics["report"], "guards": guards["report"],
        "profiles": {key: item for key, item in profiles.items()
                     if key != "captures"},
        "quality": {"sha256": quality["canonical_sha256"],
                    "attempt": quality["selected_attempt"]},
        "adverse_review": {"sha256": review["sha256"],
                           "reviewed_flags": review["reviewed_flags"]},
        "decision": decision, "owned_binaries": binaries,
        "documentation": documentation,
    }


def validate_precleanup() -> dict[str, Any]:
    """Validate the complete campaign while the owned target is still present."""
    require(os.path.lexists(TARGET) and TARGET.is_dir() and not TARGET.is_symlink(),
            "precleanup requires the owned target directory to remain present")
    campaign = validate_campaign()
    return {"status": "pass", "phase": "precleanup", "target_present": True,
            "observed_utc": dt.datetime.now(dt.timezone.utc),
            "campaign": campaign}


def validate_terminal_all() -> dict[str, Any]:
    """Validate the complete campaign after cleanup and recursive sealing."""
    cleanup = validate_cleanup()
    campaign = validate_campaign(cleanup=cleanup)
    seal = validate_seal()
    return {"status": "pass", "phase": "sealed", "target_present": False,
            "campaign": campaign, "cleanup": cleanup, "seal": seal}


def validate_inputs() -> dict[str, Any]:
    plan = validate_plan()
    return {"status": "pass", "plan": plan, "frozen": validate_frozen_inputs(),
            "supplemental": validate_supplemental_inputs(),
            "analysis_inputs": validate_analysis_inputs(),
            "workspace_lock": validate_workspace_lock(), "ADR": validate_adr(),
            "host": validate_host(), "quality_plan": validate_quality_plan()}


def run_component(selected: str) -> dict[str, Any]:
    if selected in ("inputs", "plan"):
        result = validate_inputs()
        return result if selected == "inputs" else {"status": "pass", "plan": result["plan"]}
    if selected in ("baseline-correctness", "correctness"):
        validate_inputs()
        return validate_baseline_correctness()
    if selected in ("candidate-custody", "candidate-binding", "candidate-correctness"):
        plan = validate_plan()
        base = validate_stage_source("baseline", plan)
        candidate = validate_stage_source("candidate", plan, base["manifest"])
        custody = validate_candidate_custody(plan, base["manifest"], candidate)
        if selected == "candidate-binding":
            return custody["binding"]
        if selected == "candidate-correctness":
            return custody["correctness"]
        return custody
    if selected == "source":
        plan = validate_plan()
        base = validate_stage_source("baseline", plan)
        if not CANDIDATE.exists():
            raise IncompleteError("candidate source stage is missing")
        candidate = validate_stage_source("candidate", plan, base["manifest"])
        return {"status": "pass", "baseline": base, "candidate": candidate}
    if selected == "quality":
        validate_inputs()
        return validate_quality()
    if selected in ("metrics", "main"):
        validate_adr()
        return validate_metrics_analysis()
    if selected in ("guards", "guard-cap"):
        validate_adr()
        return validate_guard_cap_analysis()
    if selected in ("adverse-review", "review"):
        metrics = validate_metrics_analysis()
        guards = validate_guard_cap_analysis()
        return validate_adverse_review(metrics, metrics["report"]["sha256"],
                                       guards, guards["report"]["sha256"])
    if selected in ("profiles", "profile"):
        metrics = validate_metrics_analysis()
        guards = validate_guard_cap_analysis()
        return validate_profiles(metrics, guards)
    if selected in ("disposition", "decision"):
        metrics = validate_metrics_analysis()
        guards = validate_guard_cap_analysis()
        profiles = validate_profiles(metrics, guards)
        quality = validate_quality()
        review = validate_adverse_review(metrics, metrics["report"]["sha256"],
                                         guards, guards["report"]["sha256"])
        return validate_disposition(metrics, guards, profiles, quality, review)
    if selected in ("owned-binaries", "binaries"):
        require(os.path.lexists(TARGET),
                "owned-binaries requires the owned target to remain present")
        return validate_owned_binaries()
    if selected == "documentation":
        return validate_documentation()
    if selected == "cleanup":
        return validate_cleanup()
    if selected == "seal":
        return validate_seal()
    if selected == "precleanup":
        return validate_precleanup()
    if selected == "all":
        return validate_terminal_all()
    raise VerificationError(f"unknown component: {selected}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", "-c", default="all", choices=(
        "inputs", "plan", "source", "baseline-correctness", "correctness", "metrics", "main",
        "candidate-custody", "candidate-binding", "candidate-correctness",
        "guards", "guard-cap", "adverse-review", "review", "profiles", "profile",
        "quality", "disposition", "decision",
        "owned-binaries", "binaries", "documentation", "cleanup", "seal",
        "precleanup", "all",
    ))
    parser.add_argument("--strict", action="store_true",
                        help="return nonzero when the selected evidence is incomplete")
    args = parser.parse_args(argv)
    try:
        result = run_component(args.component)
        envelope = {"schema": "litchi.xlsx.verification.0553.v1", "status": "pass",
                    "scope": args.component, "result": result}
        print(cli_json(envelope))
        return 0
    except IncompleteError as error:
        envelope = {"schema": "litchi.xlsx.verification.0553.v1", "status": "incomplete",
                    "scope": args.component, "error": str(error)}
        print(cli_json(envelope))
        return 2 if args.strict else 0
    except (VerificationError, OSError, KeyError, TypeError, AttributeError, IndexError) as error:
        envelope = {"schema": "litchi.xlsx.verification.0553.v1", "status": "fail",
                    "scope": args.component, "error": str(error)}
        print(cli_json(envelope))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
