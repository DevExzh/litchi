#!/usr/bin/env python3
"""Read-only, fail-closed custody verifier for the 0555 OLE2 campaign.

The measurement and quality drivers own execution and artifact creation.  This
    program only checks retained evidence.  It never builds, captures, edits the
checkout, removes the owned target, or writes an evidence file.  An absent
artifact is reported as ``incomplete``; it is never represented by a synthetic
row or zero.

The verifier independently checks custody, then replays the bound pure analysis
functions against retained captures and compares canonical reports exactly.
It also checks the final decision and terminal cleanup.  Reports are referenced by
path and digest in verifier output rather than copied into it.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys

sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TARGET = Path("/home/zhuhe/litchi-goal-0555-target")
SCRATCH_ROOT = TARGET / "retained"

PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
HOST = HERE / "host.json"
LOCK_BINDING = HERE / "workspace-lock.json"
LOCK_COPY = HERE / "workspace-Cargo.lock"
ANALYSIS_INPUTS = HERE / "analysis-inputs.json"
QUALITY_PLAN = HERE / "quality-plan.json"
QUALITY = HERE / "quality.json"
QUALITY_ATTEMPTS = HERE / "quality-attempts"
BASELINE = HERE / "baseline"
CANDIDATE = HERE / "candidate"
FINAL = HERE / "final"
METRICS_NAMES = ("metrics-analysis.json", "metrics.json", "analysis.json")
PROFILE_NAMES = ("profile-analysis.json", "profile-comparison.json", "profile-decision.json")
REVIEW_NAMES = ("adverse-review.json", "review.json")
DECISION_NAMES = ("decision.json", "disposition.json")
DOCUMENTATION = HERE / "documentation-manifest.json"
PRECLEANUP = HERE / "precleanup-verification.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"

SCHEMA = "litchi.ole2.verification.0555.v1"
CLEANUP_SCHEMA = "ole2_0555_cleanup_v1"
DOCUMENTATION_SCHEMA = "ole2_0555_documentation_manifest_v1"
STAGES = ("baseline", "candidate")
KINDS = ("normal", "alloc")
LANES = ("native", "alloc")
PROFILE_LANE = "profile"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
CPU = 2
HOST_SCOPE = "Recorded host/compiler; individual receipts retain accessible compiler observations without claiming machine quiescence"
RECEIPT_HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
PROCESS_REFERENCE_SCOPE = (
    "Accessible /proc cwd, executable and open file descriptors; "
    "cleanup process ancestors excluded."
)
ASSEMBLY_OUTPUT_SCOPE = (
    "Static instructions of emitted matching symbols; absent helpers may be inlined. "
    "Dynamic Callgrind edges and positions remain required for work attribution."
)
TARGET_STRING = str(TARGET)
DOCUMENTATION_REQUIRED = "docs/performance/0555-ole2-physical-accounting.md"
METRICS_SCHEMA = "ole2_physical_marker_metrics_0555_v1"
PROFILE_STAGE_SCHEMA = "ole2_physical_marker_0555_profile_analysis_v1"
PROFILE_COMPARISON_SCHEMA = "ole2_physical_marker_0555_profile_comparison_v1"
QUALITY_PLAN_SCHEMA = "ole2_0555_quality_plan_v1"
QUALITY_INPUT_SCHEMA = "ole2_0555_quality_inputs_v1"
QUALITY_RESULT_SCHEMA = "ole2_0555_quality_v1"
QUALITY_RECEIPT_SCHEMA = "ole2_0555_quality_receipt_v1"
REVIEW_SCHEMA = "ole2_0555_adverse_review_v1"
DECISION_SCHEMA = "ole2_0555_decision_v1"
CANDIDATE_BINDING = HERE / "candidate-binding.json"
CANDIDATE_APPLICATION = HERE / "candidate-application.json"
CANDIDATE_CORRECTION = HERE / "candidate-correction.json"
CANDIDATE_CORRECTNESS = HERE / "candidate-correctness.json"


class VerificationError(ValueError):
    """Malformed, contradictory, or out-of-scope retained evidence."""


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


def read_json(path: Path, label: str | None = None) -> object:
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


def check_hash(value: object, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase SHA-256")
    return value


def parse_time(value: object, label: str) -> dt.datetime:
    require(isinstance(value, str) and value, f"{label} is missing")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise VerificationError(f"{label} is not ISO-8601") from error
    require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed


def interval(value: object, label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    require(end > start, f"{label} interval is inverted")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and float(seconds) > 0,
            f"{label}.seconds is not positive")
    wall = (end - start).total_seconds()
    require(abs(float(seconds) - wall) <= max(0.25, wall * 0.02 + 0.05),
            f"{label}.seconds does not match UTC interval")
    return start, end


def safe_relative(value: object, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and "." not in path.parts and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def safe_repo_path(value: object, label: str) -> Path:
    safe_relative(value, label)
    path = REPO / str(value)
    require(path.resolve().is_relative_to(REPO.resolve()), f"{label} escapes repository")
    return path


def bundle_path(value: object, label: str) -> Path:
    safe_relative(value, label)
    path = HERE / str(value)
    require(path.resolve().is_relative_to(HERE.resolve()),
            f"{label} escapes the 0555 evidence bundle")
    return path


def source_name(name: str) -> bool:
    return (
        name in {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
        or name.startswith((".cargo/", "crates/", "tools/perf-baseline/"))
    )


def git(args: list[str], *, input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(
            args, cwd=REPO, input=input_data, stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise VerificationError(
            f"Git command failed ({' '.join(args)}): "
            f"{detail.decode(errors='replace')[-2000:]}"
        ) from error


def git_tree_manifest(revision: str) -> dict[str, str]:
    """Hash the frozen Git source tree without touching the worktree."""
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
            oid = fields[2].decode()
        except (UnicodeDecodeError, ValueError, IndexError) as error:
            raise VerificationError("Git source tree entry is malformed") from error
        require(len(fields) == 3 and fields[0] in (b"100644", b"100755")
                and fields[1] == b"blob" and source_name(name),
                f"Git source tree entry is out of scope: {name}")
        entries.append((name, oid))
    require(entries, "Git source tree has no source entries")
    response = git(["git", "cat-file", "--batch"],
                   input_data=("\n".join(oid for _, oid in entries) + "\n").encode())
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


def current_source_manifest() -> dict[str, str]:
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
    result: dict[str, str] = {}
    for name in sorted(names):
        path = REPO / name
        if path.is_file() and not path.is_symlink():
            result[name] = sha(path)
    return dict(sorted(result.items()))


def source_manifest(path: Path, label: str) -> dict[str, str]:
    value = read_json(path, label)
    require(isinstance(value, dict) and value, f"{label} is empty")
    require(list(value) == sorted(value), f"{label} paths are not sorted")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label} path")
        require(source_name(name), f"{label} has out-of-scope source path {name}")
        check_hash(digest, f"{label} {name}")
        require(name not in result, f"{label} repeats {name}")
        result[name] = digest
    return result


def patch_paths(path: Path, label: str) -> set[str]:
    data = read_bytes(path, label)
    if not data:
        return set()
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        raise VerificationError(f"{label} is not UTF-8") from error
    names: set[str] = set()
    headers = [line for line in text.splitlines() if line.startswith("diff --git a/")]
    if headers:
        for line in headers:
            fields = line.split()
            require(len(fields) == 4 and fields[2].startswith("a/")
                    and fields[3].startswith("b/"), f"{label} diff header is malformed")
            left, right = fields[2][2:], fields[3][2:]
            require(left == right, f"{label} renames are not admissible")
            safe_relative(left, f"{label} path")
            require(source_name(left), f"{label} contains an out-of-scope path {left}")
            names.add(left)
    else:
        pending: str | None = None
        saw = False
        for line in text.splitlines():
            if line.startswith("--- "):
                require(pending is None, f"{label} unified diff has an unpaired header")
                pending = line[4:].split("\t", 1)[0].strip()
                saw = True
            elif line.startswith("+++ "):
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
                    require(left.startswith("a/") and right.startswith("b/"
                            ) and left[2:] == right[2:],
                            f"{label} unified diff header is malformed")
                    name = left[2:]
                safe_relative(name, f"{label} path")
                require(source_name(name), f"{label} contains an out-of-scope path {name}")
                names.add(name)
        require(saw and pending is None, f"{label} unified diff headers are incomplete")
    require(names, f"{label} has no file diff headers")
    return names


def witness_for(name: str, digest: str) -> Path | None:
    """Find candidate bytes retained independently of a restored checkout."""
    choices: list[Path] = []
    direct = HERE / "candidate-preparation" / Path(name).name
    if name == "crates/litchi-cfb/src/file.rs":
        choices.append(direct)
        # A qualification attempt is a separate, immutable source witness.
        # Keep it discoverable after a test-only correction replaces the
        # top-level candidate preparation bytes.
        attempts = HERE / "candidate-attempts"
        if attempts.is_dir():
            choices.extend(
                path for path in sorted(attempts.glob("*/candidate-preparation"))
                if path.is_dir() for path in (path / Path(name).name,)
            )
    choices.extend((HERE / "candidate-sources" / name,
                    HERE / "candidate-attempts" / "sources" / name))
    choices.append(REPO / name)
    for path in choices:
        if path.is_file() and not path.is_symlink() and sha(path) == digest:
            return path
    return None


def validate_stage_source(stage: str, plan: dict[str, object],
                           baseline: dict[str, str] | None = None,
                           candidate: dict[str, str] | None = None) -> dict[str, object]:
    require(stage in (*STAGES, "final"), f"unknown source stage {stage}")
    folder = need(HERE / stage, f"{stage} stage", directory=True)
    manifest_path = need(folder / "source-manifest.json", f"{stage}/source-manifest.json")
    patch_path = need(folder / "source.patch", f"{stage}/source.patch")
    manifest = source_manifest(manifest_path, f"{stage}/source-manifest.json")
    patch = patch_paths(patch_path, f"{stage}/source.patch") if patch_path.stat().st_size else set()
    base = baseline or git_tree_manifest(str(plan["revision"]))
    if stage == "baseline":
        require(manifest == base, "baseline source manifest differs from frozen Git tree")
        require(not patch, "baseline source.patch is not empty")
    elif stage == "candidate":
        require(manifest != base, "candidate source manifest has no candidate delta")
        delta = {name for name in set(base) | set(manifest)
                 if base.get(name) != manifest.get(name)}
        require(patch == delta, "candidate source.patch changed-path inventory differs")
        for name in delta:
            require(witness_for(name, manifest[name]) is not None,
                    f"candidate source witness is missing: {name}")
    else:
        allowed = (base,) if candidate is None else (base, candidate)
        require(manifest in allowed,
                "final source is neither restored baseline nor bound candidate")
        if manifest == base:
            require(not patch, "restored final source.patch is not empty")
        else:
            delta = {name for name in set(base) | set(manifest)
                     if base.get(name) != manifest.get(name)}
            require(patch == delta, "accepted final source.patch changed-path inventory differs")
            for name in delta:
                require(witness_for(name, manifest[name]) is not None,
                        f"final candidate source witness is missing: {name}")
    return {"stage": stage, "manifest": manifest,
            "manifest_sha256": sha(manifest_path), "entries": len(manifest),
            "patch_sha256": sha(patch_path), "patch_paths": sorted(patch)}


def validate_frozen_inputs() -> dict[str, object]:
    value = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "frozen_utc", "files",
    }, "frozen-inputs envelope differs")
    require(value["schema"] == "ole2_0555_frozen_inputs_v1",
            "frozen-inputs schema differs")
    parse_time(value["frozen_utc"], "frozen-inputs.frozen_utc")
    files = value["files"]
    require(isinstance(files, dict) and set(files) == {
        "plan.json", "run.py", "workspace-lock.json", "adr-manifest.json",
    }, "frozen-inputs inventory differs")
    expected = {
        "plan.json": PLAN,
        "run.py": RUN,
        "workspace-lock.json": LOCK_BINDING,
        "adr-manifest.json": ADR,
    }
    for name, path in expected.items():
        check_hash(files[name], f"frozen-inputs.files.{name}")
        require(files[name] == sha(path), f"frozen input changed: {name}")
    return {"path": rel(FROZEN), "sha256": sha(FROZEN), "files": dict(files),
            "frozen_utc": value["frozen_utc"]}


def validate_adr() -> dict[str, object]:
    value = read_json(ADR, "adr-manifest.json")
    require(isinstance(value, dict) and value, "ADR manifest is empty")
    require(list(value) == sorted(value), "ADR manifest paths are not sorted")
    for name, digest in value.items():
        safe_relative(name, f"ADR path {name}")
        require(name.startswith("docs/adr/"), f"ADR path is out of scope: {name}")
        check_hash(digest, f"ADR hash {name}")
        require(sha(REPO / name) == digest, f"ADR changed: {name}")
    return {"path": rel(ADR), "sha256": sha(ADR), "entries": len(value)}


def validate_workspace_lock() -> dict[str, object]:
    value = read_json(LOCK_BINDING, "workspace-lock.json")
    require(isinstance(value, dict) and set(value) == {"path", "sha256", "scope"},
            "workspace-lock envelope differs")
    require(value["path"] == "Cargo.lock"
            and value["scope"] == (
                "Ignored workspace lock additionally bound on every child; perf harness lock is in source manifest"
            ), "workspace-lock identity differs")
    digest = check_hash(value["sha256"], "workspace-lock.sha256")
    require(sha(REPO / "Cargo.lock") == digest, "live workspace Cargo.lock differs")
    require(sha(LOCK_COPY) == digest, "retained workspace Cargo.lock differs")
    return {"path": rel(LOCK_BINDING), "sha256": sha(LOCK_BINDING),
            "lock_sha256": digest}


def retained_bundle_file(value: object, label: str) -> Path:
    """Resolve a retained repository path and reject outside/symlink custody."""
    path = safe_repo_path(value, label)
    require(path.resolve().is_relative_to(HERE.resolve())
            and path.is_file() and not path.is_symlink(),
            f"{label} is not a retained bundle file")
    return path


def validate_analysis_amendment(original_digest: str) -> dict[str, object]:
    """Validate the one recorded 0555 consumer-only type correction.

    The amendment is a same-turn repair to the metrics reader after a
    preflight exercised the frozen input.  It may change only the two tuple /
    set membership expressions; no plan, gate, measurement, or custody input
    is mutable through this path.
    """
    path = need(HERE / "analysis-amendment.json", "analysis-amendment.json")
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == {
        "after_path", "amended_sha256", "before_path", "changes", "created_utc",
        "failed_receipt_path", "failed_receipt_reason", "failed_receipt_sha256",
        "original_analysis_inputs_path", "original_analysis_inputs_sha256",
        "original_sha256", "path", "preserved_analyzer_path",
        "preserved_analyzer_sha256", "preserved_failed_receipt_path",
        "preserved_failed_receipt_sha256", "preserved_failed_stderr_path",
        "preserved_failed_stderr_sha256", "preserved_failed_stdout_path",
        "preserved_failed_stdout_sha256", "reason", "schema", "scope",
    }, "analysis-amendment envelope differs")
    require(value["schema"] == "ole2_0555_analysis_amendment_v1"
            and value["path"] == value["after_path"]
            and value["after_path"] == "docs/performance/results/change-0555/analyze_metrics.py"
            and value["before_path"] == value["preserved_analyzer_path"]
            and value["original_sha256"] == original_digest
            and value["preserved_analyzer_sha256"] == original_digest
            and value["original_analysis_inputs_path"] == ANALYSIS_INPUTS.relative_to(REPO).as_posix()
            and value["original_analysis_inputs_sha256"] == sha(ANALYSIS_INPUTS),
            "analysis-amendment identity differs")
    created = parse_time(value["created_utc"], "analysis-amendment.created_utc")
    for key in ("reason", "failed_receipt_reason", "scope"):
        require(isinstance(value[key], str) and value[key].strip(),
                f"analysis-amendment.{key} is missing")
    require("tuple" in value["reason"] and "set" in value["reason"]
            and "tuple" in value["failed_receipt_reason"]
            and "set" in value["failed_receipt_reason"]
            and "measurement" in value["scope"],
            "analysis-amendment explanation changes the admitted boundary")
    before = retained_bundle_file(value["before_path"], "analysis-amendment.before_path")
    after = retained_bundle_file(value["after_path"], "analysis-amendment.after_path")
    require(sha(before) == original_digest and sha(after) == value["amended_sha256"],
            "analysis-amendment analyzer digest differs")
    amended = check_hash(value["amended_sha256"], "analysis-amendment.amended_sha256")
    require(amended == sha(HERE / "analyze_metrics.py"),
            "analysis-amendment current analyzer differs")
    changes = value["changes"]
    require(isinstance(changes, dict) and set(changes) == {"analyze_metrics.py"},
            "analysis-amendment change inventory differs")
    change = changes["analyze_metrics.py"]
    require(isinstance(change, dict) and set(change) == {
        "amended", "description", "final_sha256", "initial_sha256", "original",
    } and change["amended"] == "analyze_metrics.py"
            and change["original"] == "analyzer-attempts/metrics-before-stage-set/analyze_metrics.py"
            and change["initial_sha256"] == original_digest
            and change["final_sha256"] == amended
            and isinstance(change["description"], str)
            and change["description"] == (
                "Replace tuple/set membership with set(STAGES) | {'final'} for stage and "
                "execution-stage validation; no measurement, gate, matrix, or custody rule changes."
            ), "analysis-amendment change differs")
    before_text = read_text(before, "analysis-amendment.before analyzer")
    after_text = read_text(after, "analysis-amendment.after analyzer")
    require(before_text.count("in STAGES | {\"final\"}") == 2
            and after_text.count("in set(STAGES) | {\"final\"}") == 2
            and before_text.replace("in STAGES | {\"final\"}",
                                    "in set(STAGES) | {\"final\"}") == after_text,
            "analysis-amendment contains changes outside the two type fixes")
    failed_path = retained_bundle_file(value["failed_receipt_path"],
                                       "analysis-amendment.failed_receipt_path")
    preserved_receipt = retained_bundle_file(value["preserved_failed_receipt_path"],
                                             "analysis-amendment.preserved_failed_receipt_path")
    preserved_stdout = retained_bundle_file(value["preserved_failed_stdout_path"],
                                            "analysis-amendment.preserved_failed_stdout_path")
    preserved_stderr = retained_bundle_file(value["preserved_failed_stderr_path"],
                                            "analysis-amendment.preserved_failed_stderr_path")
    check_hash(value["failed_receipt_sha256"], "analysis-amendment.failed_receipt_sha256")
    check_hash(value["preserved_failed_receipt_sha256"],
               "analysis-amendment.preserved_failed_receipt_sha256")
    check_hash(value["preserved_failed_stdout_sha256"],
               "analysis-amendment.preserved_failed_stdout_sha256")
    check_hash(value["preserved_failed_stderr_sha256"],
               "analysis-amendment.preserved_failed_stderr_sha256")
    require(value["failed_receipt_sha256"] == sha(failed_path)
            and value["preserved_failed_receipt_sha256"] == sha(preserved_receipt)
            and value["preserved_failed_stdout_sha256"] == sha(preserved_stdout)
            and value["preserved_failed_stderr_sha256"] == sha(preserved_stderr)
            and failed_path.read_bytes() == preserved_receipt.read_bytes(),
            "analysis-amendment failed preflight custody differs")
    failed = read_json(failed_path, rel(failed_path))
    require(isinstance(failed, dict) and set(failed) == {
        "command", "exit_code", "observed_utc", "script_sha256", "scope",
    } and failed["exit_code"] == 1
            and failed["command"] == [
                "python3", "-B",
                "docs/performance/results/change-0555/analysis-attempts/baseline-r1-preflight-01/script.py",
            ] and failed["scope"] == (
                "Read-only baseline R1 parser and build custody preflight; no matched performance conclusion"
            ), "analysis-amendment failed receipt differs")
    observed = parse_time(failed["observed_utc"], "failed analysis observed_utc")
    check_hash(failed["script_sha256"], "failed analysis script_sha256")
    script = retained_bundle_file(failed["command"][2], "failed analysis script")
    require(failed["script_sha256"] == sha(script)
            and observed <= created,
            "analysis-amendment failed preflight binding differs")
    require("tuple" in read_text(preserved_stderr, "preserved failed analyzer stderr")
            and "set" in read_text(preserved_stderr, "preserved failed analyzer stderr"),
            "analysis-amendment does not retain the recorded type failure")
    return {"path": rel(path), "sha256": sha(path), "original_sha256": original_digest,
            "amended_sha256": amended, "failed_receipt_sha256": value["failed_receipt_sha256"]}


def validate_profile_amendment(original_digest: str) -> dict[str, object]:
    """Validate the one recorded 0555 profile-consumer state correction."""
    path = need(HERE / "profile-amendment.json", "profile-amendment.json")
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == {
        "after_path", "amended_sha256", "before_path", "changes", "created_utc",
        "failed_receipt_path", "failed_receipt_reason", "failed_receipt_sha256",
        "original_analysis_inputs_path", "original_analysis_inputs_sha256",
        "original_sha256", "path", "preserved_analyzer_path",
        "preserved_analyzer_sha256", "preserved_failed_receipt_path",
        "preserved_failed_receipt_sha256", "preserved_failed_stderr_path",
        "preserved_failed_stderr_sha256", "preserved_failed_stdout_path",
        "preserved_failed_stdout_sha256", "reason", "schema", "scope",
    }, "profile-amendment envelope differs")
    require(value["schema"] == "ole2_0555_profile_amendment_v1"
            and value["path"] == value["after_path"]
            and value["after_path"] == "docs/performance/results/change-0555/analyze_profiles.py"
            and value["before_path"] == value["preserved_analyzer_path"]
            and value["original_sha256"] == original_digest
            and value["preserved_analyzer_sha256"] == original_digest
            and value["original_analysis_inputs_path"] == ANALYSIS_INPUTS.relative_to(REPO).as_posix()
            and value["original_analysis_inputs_sha256"] == sha(ANALYSIS_INPUTS),
            "profile-amendment identity differs")
    created = parse_time(value["created_utc"], "profile-amendment.created_utc")
    for key in ("reason", "failed_receipt_reason", "scope"):
        require(isinstance(value[key], str) and value[key].strip(),
                f"profile-amendment.{key} is missing")
    require("state" in value["reason"] and "per-dump" in value["reason"]
            and "KeyError" in value["failed_receipt_reason"]
            and "measurement" in value["scope"]
            and "mechanism thresholds" in value["scope"]
            and "adoption rules" in value["scope"],
            "profile-amendment explanation changes the admitted boundary")
    before = retained_bundle_file(value["before_path"], "profile-amendment.before_path")
    after = retained_bundle_file(value["after_path"], "profile-amendment.after_path")
    require(sha(before) == original_digest and sha(after) == value["amended_sha256"]
            and sha(HERE / "analyze_profiles.py") == value["amended_sha256"],
            "profile-amendment analyzer digest differs")
    amended = check_hash(value["amended_sha256"], "profile-amendment.amended_sha256")
    changes = value["changes"]
    require(isinstance(changes, dict) and set(changes) == {"analyze_profiles.py"},
            "profile-amendment change inventory differs")
    change = changes["analyze_profiles.py"]
    require(isinstance(change, dict) and set(change) == {
        "amended", "description", "final_sha256", "initial_sha256", "original",
    } and change["amended"] == "analyze_profiles.py"
            and change["original"] == "analyzer-attempts/profile-before-comparison-state/analyze_profiles.py"
            and change["initial_sha256"] == original_digest
            and change["final_sha256"] == amended
            and isinstance(change["description"], str)
            and change["description"] == (
                "Use the actual per-dump claim_sector attribution states when evaluating the "
                "inline/out-of-line evidence flag; no capture, source, profile target, "
                "mechanism threshold, or adoption rule changes."
            ), "profile-amendment change differs")
    before_text = read_text(before, "profile-amendment.before analyzer")
    after_text = read_text(after, "profile-amendment.after analyzer")
    insert_block = (
        "    claim_states = [\n"
        "        dump[\"attribution\"][\"claim_sector\"][\"state\"]\n"
        "        for profile in [*left.values(), *right.values()]\n"
        "        for dump in profile[\"timed_dumps\"]\n"
        "    ]\n"
    )
    anchor = (
        "    physical_decreases = (\n"
        "        physical_present and all(\n"
        "            row[\"targets\"][\"physical_reconciliation\"][\"self_ir\"][\"delta_ir\"] < 0\n"
        "            for row in physical_rows\n"
        "        )\n"
        "    )\n"
    )
    old_gate = (
        "                row[\"targets\"][\"claim_sector\"][\"state\"] in {\n"
        "                    \"out_of_line\", \"inlined_or_absent\",\n"
        "                    \"present_without_positive_incoming_edge\",\n"
        "                }\n"
        "                for row in rows\n"
    )
    new_gate = (
        "                state in {\n"
        "                    \"out_of_line\", \"inlined_or_absent\",\n"
        "                    \"present_without_positive_incoming_edge\",\n"
        "                }\n"
        "                for state in claim_states\n"
    )
    require(before_text.count(anchor) == 1 and before_text.count(old_gate) == 1,
            "profile-amendment source anchors differ")
    expected_after = before_text.replace(anchor, anchor + insert_block, 1)
    expected_after = expected_after.replace(old_gate, new_gate, 1)
    require(after_text == expected_after, "profile-amendment contains extra source changes")

    failed_path = retained_bundle_file(value["failed_receipt_path"],
                                       "profile-amendment.failed_receipt_path")
    preserved_receipt = retained_bundle_file(value["preserved_failed_receipt_path"],
                                             "profile-amendment.preserved_failed_receipt_path")
    preserved_stdout = retained_bundle_file(value["preserved_failed_stdout_path"],
                                            "profile-amendment.preserved_failed_stdout_path")
    preserved_stderr = retained_bundle_file(value["preserved_failed_stderr_path"],
                                            "profile-amendment.preserved_failed_stderr_path")
    for key in ("failed_receipt_sha256", "preserved_failed_receipt_sha256",
                "preserved_failed_stdout_sha256", "preserved_failed_stderr_sha256"):
        check_hash(value[key], f"profile-amendment.{key}")
    require(value["failed_receipt_sha256"] == sha(failed_path)
            and value["preserved_failed_receipt_sha256"] == sha(preserved_receipt)
            and value["preserved_failed_stdout_sha256"] == sha(preserved_stdout)
            and value["preserved_failed_stderr_sha256"] == sha(preserved_stderr)
            and failed_path.read_bytes() == preserved_receipt.read_bytes(),
            "profile-amendment failed comparison custody differs")
    failed = read_json(failed_path, rel(failed_path))
    require(isinstance(failed, dict) and set(failed) == {
        "command", "start_utc", "end_utc", "exit_code", "script_sha256",
        "plan_sha256", "stdout_sha256", "stderr_sha256",
    } and failed["command"] == [
        "python3", "-B", "docs/performance/results/change-0555/analyze_profiles.py",
        "--compare", "--output", "docs/performance/results/change-0555/profile-comparison.json",
    ] and failed["exit_code"] == 2
            and failed["script_sha256"] == original_digest
            and failed["plan_sha256"] == sha(PLAN)
            and failed["stdout_sha256"] == sha(preserved_stdout)
            and failed["stderr_sha256"] == sha(preserved_stderr),
            "profile-amendment failed receipt differs")
    failed_start = parse_time(failed["start_utc"], "failed profile start_utc")
    failed_end = parse_time(failed["end_utc"], "failed profile end_utc")
    require(failed_end > failed_start and failed_end <= created,
            "profile-amendment failed receipt timing differs")
    return {"path": rel(path), "sha256": sha(path), "original_sha256": original_digest,
            "amended_sha256": amended, "failed_receipt_sha256": value["failed_receipt_sha256"]}


def validate_analysis_inputs() -> dict[str, object]:
    value = read_json(ANALYSIS_INPUTS, "analysis-inputs.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "frozen_utc", "scope", "files",
    }, "analysis-inputs envelope differs")
    require(value["schema"] == "ole2_0555_analysis_inputs_v1"
            and isinstance(value["scope"], str) and value["scope"].strip(),
            "analysis-inputs identity differs")
    frozen = parse_time(value["frozen_utc"], "analysis-inputs.frozen_utc")
    files = value["files"]
    require(isinstance(files, dict) and files and list(files) == sorted(files),
            "analysis-inputs file map differs")
    amendments: dict[str, object] = {}
    for name, digest in files.items():
        check_hash(digest, f"analysis-inputs.files.{name}")
        actual = sha(safe_repo_path(name, f"analysis-inputs.files.{name}"))
        if actual != digest and name == "docs/performance/results/change-0555/analyze_metrics.py":
            amendments["analyze_metrics.py"] = validate_analysis_amendment(digest)
        elif actual != digest and name == "docs/performance/results/change-0555/analyze_profiles.py":
            amendments["analyze_profiles.py"] = validate_profile_amendment(digest)
        else:
            require(actual == digest, f"analysis input changed: {name}")
    # 0555 has a fresh consumer freeze.  A prior-turn admission supplement is
    # never admissible.  A same-turn consumer correction, when retained, is
    # checked by validate_analysis_amendment() rather than silently ignored.
    require(not (HERE / "admission-supplement.json").exists(),
            "admission supplement is not admissible in 0555")
    if (HERE / "analysis-amendment.json").exists() and "analyze_metrics.py" not in amendments:
        amendments["analyze_metrics.py"] = validate_analysis_amendment(
            files["docs/performance/results/change-0555/analyze_metrics.py"]
        )
    if (HERE / "profile-amendment.json").exists() and "analyze_profiles.py" not in amendments:
        amendments["analyze_profiles.py"] = validate_profile_amendment(
            files["docs/performance/results/change-0555/analyze_profiles.py"]
        )
    # This freeze is meaningful only if it predates the first candidate
    # capture.  During the preparatory baseline phase no candidate receipt may
    # exist yet, so defer the ordering assertion until one is retained.
    candidate_receipts = sorted(CANDIDATE.glob("*.receipt.json")) \
        if CANDIDATE.is_dir() else []
    if candidate_receipts:
        first = min(parse_time(read_json(path, rel(path))["start_utc"], rel(path))
                    for path in candidate_receipts)
        require(frozen < first, "analysis-inputs freeze postdates first candidate capture")
    return {"path": rel(ANALYSIS_INPUTS), "sha256": sha(ANALYSIS_INPUTS),
            "frozen_utc": value["frozen_utc"], "files": dict(files),
            "amendments": amendments}


def validate_candidate_variant(manifest_path: Path, patch_path: Path,
                               base: dict[str, str], label: str) -> dict[str, object]:
    """Validate an isolated candidate source snapshot and its exact patch."""
    manifest = source_manifest(manifest_path, f"{label}/source-manifest.json")
    patch = patch_paths(patch_path, f"{label}/source.patch")
    require(manifest != base, f"{label} has no candidate delta")
    delta = {name for name in set(base) | set(manifest)
             if base.get(name) != manifest.get(name)}
    require(patch == delta, f"{label} patch changed-path inventory differs")
    for name in delta:
        require(witness_for(name, manifest[name]) is not None,
                f"{label} source witness is missing: {name}")
    return {"manifest": manifest, "manifest_sha256": sha(manifest_path),
            "patch_sha256": sha(patch_path), "patch_paths": sorted(patch)}


def validate_candidate_history(plan: dict[str, object]) -> dict[str, object]:
    """Bind the candidate review, correction, and preserved failed attempt.

    0555 had a qualification-only correction before candidate capture.  The
    failed source and quality result remain separate evidence; this validator
    prevents the corrected source from erasing that history.
    """
    base = git_tree_manifest(str(plan["revision"]))
    candidate_manifest_path = need(CANDIDATE / "source-manifest.json",
                                   "candidate/source-manifest.json")
    candidate_patch_path = need(CANDIDATE / "source.patch", "candidate/source.patch")
    candidate = validate_candidate_variant(candidate_manifest_path,
                                           candidate_patch_path, base, "candidate")
    candidate_file = "crates/litchi-cfb/src/file.rs"
    require(candidate["manifest"].get(candidate_file),
            "candidate source file is absent")

    binding_path = need(CANDIDATE_BINDING, "candidate-binding.json")
    binding = read_json(binding_path, rel(binding_path))
    require(isinstance(binding, dict) and set(binding) == {
        "schema", "frozen_utc", "baseline_revision", "source_path",
        "baseline_source_sha256", "candidate_source_sha256", "candidate_patch_sha256",
        "plan_sha256", "representation_selection_sha256", "candidate_review_sha256",
        "resource_design_review_sha256", "scope",
    }, "candidate-binding envelope differs")
    require(binding["schema"] == "ole2_0555_candidate_binding_v1"
            and binding["baseline_revision"] == plan["revision"]
            and binding["source_path"] == candidate_file
            and binding["baseline_source_sha256"] == base[candidate_file]
            and binding["candidate_source_sha256"] == candidate["manifest"][candidate_file]
            and binding["candidate_patch_sha256"] == candidate["patch_sha256"]
            and binding["plan_sha256"] == sha(PLAN)
            and binding["representation_selection_sha256"] == sha(HERE / "representation-selection.json")
            and binding["candidate_review_sha256"] == sha(HERE / "candidate-review.md")
            and binding["resource_design_review_sha256"] == sha(HERE / "resource-design-review.md")
            and isinstance(binding["scope"], str) and binding["scope"].strip(),
            "candidate-binding identity differs")
    parse_time(binding["frozen_utc"], "candidate-binding.frozen_utc")
    for key in ("baseline_source_sha256", "candidate_source_sha256",
                "candidate_patch_sha256", "plan_sha256",
                "representation_selection_sha256", "candidate_review_sha256",
                "resource_design_review_sha256"):
        check_hash(binding[key], f"candidate-binding.{key}")

    application_path = need(CANDIDATE_APPLICATION, "candidate-application.json")
    application = read_json(application_path, rel(application_path))
    require(isinstance(application, dict) and set(application) == {
        "schema", "applied_utc", "binding_sha256", "source_verified", "scope",
    } and application["schema"] == "ole2_0555_candidate_application_v1"
            and application["binding_sha256"] == sha(binding_path)
            and application["source_verified"] is True
            and isinstance(application["scope"], str) and application["scope"].strip(),
            "candidate-application identity differs")
    parse_time(application["applied_utc"], "candidate-application.applied_utc")

    correction_path = need(CANDIDATE_CORRECTION, "candidate-correction.json")
    correction = read_json(correction_path, rel(correction_path))
    require(isinstance(correction, dict) and set(correction) == {
        "schema", "observed_utc", "reason", "failed_quality_result_sha256",
        "original_source_sha256", "corrected_source_sha256", "corrected_patch_sha256",
        "scope",
    }, "candidate-correction envelope differs")
    correction_time = parse_time(correction["observed_utc"],
                                 "candidate-correction.observed_utc")
    require(correction["schema"] == "ole2_0555_candidate_correction_v1"
            and correction["corrected_source_sha256"] == candidate["manifest"][candidate_file]
            and correction["corrected_patch_sha256"] == candidate["patch_sha256"]
            and isinstance(correction["reason"], str) and correction["reason"].strip()
            and isinstance(correction["scope"], str) and correction["scope"] == (
                "Test-only qualification; production candidate implementation identical; no candidate captures existed"
            ), "candidate-correction identity differs")
    for key in ("failed_quality_result_sha256", "original_source_sha256",
                "corrected_source_sha256", "corrected_patch_sha256"):
        check_hash(correction[key], f"candidate-correction.{key}")

    # Locate the immutable pre-correction candidate snapshot and bind its
    # source/patch pair to the correction record.
    attempt_root = HERE / "candidate-attempts" / "qualification-01"
    attempt_stage = need(attempt_root / "stage", "candidate-attempts/qualification-01/stage",
                          directory=True)
    attempt_manifest_path = need(attempt_stage / "source-manifest.json",
                                 "candidate-attempts/qualification-01/stage/source-manifest.json")
    attempt_patch_path = need(attempt_stage / "source.patch",
                              "candidate-attempts/qualification-01/stage/source.patch")
    attempt = validate_candidate_variant(attempt_manifest_path, attempt_patch_path,
                                         base, "candidate-attempts/qualification-01/stage")
    require(attempt["manifest"][candidate_file] == correction["original_source_sha256"],
            "candidate-correction original source differs")

    attempt_binding_path = need(attempt_root / "candidate-binding.json",
                                "candidate-attempts/qualification-01/candidate-binding.json")
    attempt_binding = read_json(attempt_binding_path, rel(attempt_binding_path))
    require(isinstance(attempt_binding, dict) and set(attempt_binding) == {
        "schema", "frozen_utc", "baseline_revision", "source_path",
        "baseline_source_sha256", "candidate_source_sha256", "candidate_patch_sha256",
        "plan_sha256", "representation_selection_sha256", "candidate_review_sha256",
        "resource_design_review_sha256", "scope",
    } and attempt_binding["schema"] == "ole2_0555_candidate_binding_v1"
            and attempt_binding["baseline_revision"] == plan["revision"]
            and attempt_binding["source_path"] == candidate_file
            and attempt_binding["baseline_source_sha256"] == base[candidate_file]
            and attempt_binding["candidate_source_sha256"] == attempt["manifest"][candidate_file]
            and attempt_binding["plan_sha256"] == sha(PLAN)
            and attempt_binding["representation_selection_sha256"] == sha(HERE / "representation-selection.json")
            and attempt_binding["candidate_review_sha256"] == sha(attempt_root / "candidate-review.md")
            and attempt_binding["resource_design_review_sha256"] == sha(HERE / "resource-design-review.md")
            and isinstance(attempt_binding["scope"], str)
            and attempt_binding["scope"].strip(),
            "preserved candidate-binding identity differs")
    parse_time(attempt_binding["frozen_utc"], "preserved candidate-binding.frozen_utc")
    for key in ("baseline_source_sha256", "candidate_source_sha256",
                "candidate_patch_sha256", "plan_sha256",
                "representation_selection_sha256", "candidate_review_sha256",
                "resource_design_review_sha256"):
        check_hash(attempt_binding[key], f"preserved candidate-binding.{key}")
    attempt_preparation_patch = need(
        attempt_root / "candidate-preparation" / "candidate.patch",
        "preserved candidate patch witness")
    require(attempt_binding["candidate_patch_sha256"] == sha(attempt_preparation_patch),
            "preserved candidate patch witness differs")
    attempt_application_path = need(attempt_root / "candidate-application.json",
                                    "candidate-attempts/qualification-01/candidate-application.json")
    attempt_application = read_json(attempt_application_path, rel(attempt_application_path))
    require(isinstance(attempt_application, dict) and set(attempt_application) == {
        "schema", "applied_utc", "binding_sha256", "source_verified",
    } and attempt_application["schema"] == "ole2_0555_candidate_application_v1"
            and attempt_application["binding_sha256"] == sha(attempt_binding_path)
            and attempt_application["source_verified"] is True,
            "preserved candidate-application identity differs")
    parse_time(attempt_application["applied_utc"],
               "preserved candidate-application.applied_utc")

    failed_quality_path = need(HERE / "quality-attempts" / "candidate-targeted-01" / "result.json",
                               "quality-attempts/candidate-targeted-01/result.json")
    require(correction["failed_quality_result_sha256"] == sha(failed_quality_path),
            "candidate-correction failed quality result differs")
    failed_quality = read_json(failed_quality_path, rel(failed_quality_path))
    require(isinstance(failed_quality, dict)
            and failed_quality.get("schema") == QUALITY_RESULT_SCHEMA
            and failed_quality.get("status") == "failed"
            and failed_quality.get("stage") == "candidate"
            and failed_quality.get("source_manifest_sha256") == attempt["manifest_sha256"],
            "candidate-correction quality failure differs")
    # The correction record says no candidate capture existed yet.  Any
    # candidate receipt retained later must therefore begin after it.
    later_receipts = sorted(CANDIDATE.glob("*.receipt.json")) \
        if CANDIDATE.is_dir() else []
    require(all(correction_time < parse_time(read_json(item, rel(item))["start_utc"],
                                              rel(item)) for item in later_receipts),
            "candidate-correction predates no candidate capture")
    original_source = need(attempt_root / "candidate-preparation" / "file.rs",
                           "preserved candidate source witness")
    corrected_source = need(HERE / "candidate-preparation" / "file.rs",
                            "corrected candidate source witness")
    original_text = read_text(original_source, "preserved candidate source witness")
    corrected_text = read_text(corrected_source, "corrected candidate source witness")
    old_line = "assert_eq!(std::mem::size_of::<PhysicalSectorRole>(), 1);"
    new_line = "assert_eq!(size_of::<PhysicalSectorRole>(), 1);"
    require(original_text.count(old_line) == 1 and corrected_text.count(new_line) == 1
            and original_text.replace(old_line, new_line) == corrected_text,
            "candidate-correction changes more than the qualification spelling")

    correctness_path = need(CANDIDATE_CORRECTNESS, "candidate-correctness.json")
    correctness = read_json(correctness_path, rel(correctness_path))
    require(isinstance(correctness, dict) and set(correctness) == {
        "schema", "status", "source_manifest_sha256", "candidate_binding_sha256",
        "review_sha256", "quality_result_path", "quality_result_sha256",
        "passed_tests", "failed_tests", "ignored_tests", "groups", "scope",
    } and correctness["schema"] == "ole2_0555_candidate_correctness_v1"
            and correctness["status"] == "pass"
            and correctness["source_manifest_sha256"] == candidate["manifest_sha256"]
            and correctness["candidate_binding_sha256"] == sha(binding_path)
            and correctness["review_sha256"] == sha(HERE / "candidate-review.md")
            and correctness["quality_result_path"] == "quality-attempts/candidate-targeted-02/result.json"
            and correctness["quality_result_sha256"] == sha(
                HERE / "quality-attempts/candidate-targeted-02/result.json")
            and correctness["passed_tests"] == 313
            and correctness["failed_tests"] == 0
            and correctness["ignored_tests"] == 1
            and correctness["groups"] == 4
            and isinstance(correctness["scope"], str) and correctness["scope"].strip(),
            "candidate-correctness identity differs")
    for key in ("source_manifest_sha256", "candidate_binding_sha256", "review_sha256",
                "quality_result_sha256"):
        check_hash(correctness[key], f"candidate-correctness.{key}")
    return {
        "binding": {"path": rel(binding_path), "sha256": sha(binding_path)},
        "application": {"path": rel(application_path), "sha256": sha(application_path)},
        "correction": {"path": rel(correction_path), "sha256": sha(correction_path)},
        "correctness": {"path": rel(correctness_path), "sha256": sha(correctness_path)},
        "preserved_attempt": {
            "binding_sha256": sha(attempt_binding_path),
            "source_manifest_sha256": attempt["manifest_sha256"],
            "quality_result_sha256": correction["failed_quality_result_sha256"],
        },
    }


def validate_host() -> dict[str, object]:
    value = read_json(HOST, "host.json")
    require(isinstance(value, dict) and set(value) == {
        "observed_utc", "rustc", "cargo", "uname", "cpu_allowed", "scope",
    }, "host envelope differs")
    parse_time(value["observed_utc"], "host.observed_utc")
    for field in ("rustc", "cargo", "uname"):
        require(isinstance(value[field], str) and value[field].strip(),
                f"host.{field} is missing")
    affinity = value["cpu_allowed"]
    require(isinstance(affinity, list) and affinity == sorted(set(affinity))
            and all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                    for item in affinity) and CPU in affinity,
            "host CPU affinity differs")
    require(value["scope"] == HOST_SCOPE, "host scope differs")
    return {"path": rel(HOST), "sha256": sha(HOST), "cpu_allowed": affinity}


def validate_plan() -> dict[str, object]:
    value = read_json(PLAN, "plan.json")
    require(isinstance(value, dict), "plan is not an object")
    required = {
        "schema", "revision", "created_utc", "previous_turn", "priority", "scope",
        "hypothesis", "candidate_files", "candidate_scope", "cpu", "groups", "native",
        "allocation", "profile", "assembly", "admission", "prior_scope_note",
        "source_binding", "owned_paths", "temporary_storage", "status", "limitations",
    }
    require(set(value) == required, "plan field inventory differs")
    require(value["schema"] == "ole2_physical_marker_0555_plan_v1"
            and isinstance(value["revision"], str)
            and REVISION_RE.fullmatch(value["revision"]) is not None,
            "plan identity differs")
    try:
        subprocess.run(["git", "cat-file", "-e", value["revision"] + "^{commit}"],
                       cwd=REPO, check=True, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError("plan revision is not a Git commit") from error
    parse_time(value["created_utc"], "plan.created_utc")
    require(value["priority"] == (
                "OLE2/OOXML first; ODF deferred until the OLE2/OOXML optimization goal "
                "completes; iWork excluded"
            )
            and value["scope"] == (
                "Matched OLE2 physical-sector marker accounting experiment across the FAT, "
                "DIFAT, Directory, MiniFAT, MiniStream, and RegularStream roles; no public "
                "API or semantic format change"
            )
            and value["cpu"] == CPU and value["owned_paths"] == [TARGET_STRING]
            and value["candidate_files"] == ["crates/litchi-cfb/src/file.rs"]
            and value["status"] == (
                "frozen before build, capture, candidate application, and result observation"
            ),
            "plan execution identity differs")
    for key in ("previous_turn", "hypothesis", "prior_scope_note", "temporary_storage"):
        require(isinstance(value[key], str) and value[key].strip(),
                f"plan.{key} is missing")
    candidate_scope = value["candidate_scope"]
    require(isinstance(candidate_scope, dict) and set(candidate_scope) == {
        "implementation", "physical_roles", "semantic_invariants", "allocation_constraint",
    }, "plan candidate scope differs")
    require(isinstance(candidate_scope["implementation"], str)
            and candidate_scope["implementation"].strip()
            and isinstance(candidate_scope["allocation_constraint"], str)
            and candidate_scope["allocation_constraint"].strip(),
            "plan candidate scope text differs")
    invariants = candidate_scope["semantic_invariants"]
    require(isinstance(invariants, list) and invariants
            and all(isinstance(item, str) and item.strip() for item in invariants),
            "plan semantic invariants differ")
    require(candidate_scope["physical_roles"] == [
        "FAT", "DIFAT", "Directory", "MiniFAT", "MiniStream", "RegularStream",
    ], "plan physical role inventory differs")
    require(value["groups"] == {
        "xls": {"cases": [
            "xls_semantic_open", "xls_eager_open_list_worksheets",
            "xls_eager_open_one_cell", "xls_source_backed_open",
            "xls_source_backed_open_list_worksheets", "xls_source_backed_open_one_cell",
            "xls_owned_source_open", "xls_owned_source_open_list_worksheets",
            "xls_owned_source_open_one_cell",
        ], "primary_cases": [
            "xls_source_backed_open", "xls_source_backed_open_one_cell",
            "xls_owned_source_open", "xls_owned_source_open_one_cell",
        ]},
        "cfb": {"cases": ["cfb_open"],
                "shapes": ["tiny", "many-small", "few-large"],
                "payload": "incompressible"},
    }, "plan case matrix differs")
    order = [
        "baseline r1 xls/cfb", "candidate r1 xls/cfb",
        "candidate r2 cfb/xls", "baseline r2 cfb/xls under candidate execution-stage binding",
    ]
    require(value["native"] == {
        "repeats": 2, "warmup": 20, "samples": 1000,
        "order": order,
        "source_identity": (
            "Every receipt binds both output stage and live execution stage; retained "
            "baseline-r2 uses the baseline binary/output folder with candidate source "
            "execution identity"
        ),
    }, "plan native lane differs")
    require(value["allocation"] == {
        "repeats": 2, "warmup": 3, "samples": 30,
        "order": order,
        "scope": (
            "Existing constructor/operation clock with the canonical operation-global "
            "System allocator region; excludes fixture construction, correctness oracles, "
            "report construction, and object drop"
        ),
        "vector": (
            "Every nine XLS case and three CFB shape in both repeats; calls, allocated "
            "bytes, and incremental region peak are retained per operation"
        ),
    }, "plan allocation lane differs")
    profile = value["profile"]
    require(isinstance(profile, dict) and profile.get("repeats") == 2
            and profile.get("warmup") == 0 and profile.get("samples") == 5
            and profile.get("jobs") == ["xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large"]
            and isinstance(profile.get("positive_timed_dumps"), str)
            and profile.get("positive_timed_dumps").strip()
            and profile.get("xls_owner") == (
                "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
            )
            and profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open"
            and isinstance(profile.get("scope"), str) and profile["scope"].strip()
            and profile.get("instruction_flags") == [
                "--dump-instr=yes", "--dump-line=no", "--compress-pos=no",
                "--collect-jumps=yes",
            ]
            and profile.get("runtime_flags") == [
                "--vgdb=no", "--collect-atstart=no", "--toggle-collect=<owner>",
                "--zero-before=<owner>", "--dump-after=<owner>",
            ],
            "plan profile matrix differs")
    assembly = value["assembly"]
    require(isinstance(assembly, dict) and set(assembly) == {
        "scope", "requested_owners", "both_stages",
    } and isinstance(assembly["scope"], str) and assembly["scope"].strip()
            and isinstance(assembly["requested_owners"], list)
            and assembly["requested_owners"]
            and all(isinstance(owner, str) and owner.strip()
                    for owner in assembly["requested_owners"])
            and len(set(assembly["requested_owners"])) == len(assembly["requested_owners"])
            and assembly["both_stages"] is True,
            "plan assembly matrix differs")
    admission = value["admission"]
    require(isinstance(admission, dict) and admission, "plan admission is empty")
    require(set(admission) == {"primary_xls_p50", "primary_xls_mean",
                               "native_xls_controls", "native_cfb_controls", "native_rss",
                               "allocation", "correctness", "mechanism", "review", "disposition"},
            "plan admission field inventory differs")
    for key in admission:
        require(isinstance(admission.get(key), str) and admission[key].strip(),
                f"plan admission.{key} is missing")
    source_binding = value["source_binding"]
    require(isinstance(source_binding, dict) and set(source_binding) == {
        "baseline_revision", "source_manifest", "workspace_lock_binding",
        "workspace_lock_copy", "adr_manifest", "receipt_requirements",
    }, "plan source binding differs")
    require(source_binding["baseline_revision"] == value["revision"]
            and source_binding["source_manifest"] == "{stage}/source-manifest.json"
            and source_binding["workspace_lock_binding"] == "workspace-lock.json"
            and source_binding["workspace_lock_copy"] == "workspace-Cargo.lock"
            and source_binding["adr_manifest"] == "adr-manifest.json"
            and source_binding["receipt_requirements"] == [
                "script_sha256", "plan_sha256", "source_manifest_sha256",
                "execution_manifest_sha256", "workspace_lock_sha256",
                "workspace_lock_binding_sha256",
                "binary_sha256 when a binary is used",
            ], "plan source binding identity differs")
    limitations = value["limitations"]
    require(isinstance(limitations, list) and limitations
            and all(isinstance(item, str) and item.strip() for item in limitations),
            "plan limitations differ")
    return {"path": rel(PLAN), "sha256": sha(PLAN), "revision": value["revision"],
            "value": value}


def validate_inputs() -> dict[str, object]:
    plan = validate_plan()
    plan_sha = sha(PLAN)
    frozen = validate_frozen_inputs()
    require(frozen["files"]["plan.json"] == plan_sha, "frozen plan binding differs")
    return {
        "plan": {"sha256": plan_sha, "revision": plan["revision"]},
        "frozen": frozen, "lock": validate_workspace_lock(),
        "analysis_inputs": validate_analysis_inputs(),
        "candidate": validate_candidate_history(plan),
        "ADR": validate_adr(), "host": validate_host(),
        "quality_plan": validate_quality_plan(),
    }


def validate_quality_plan() -> dict[str, object]:
    value = read_json(QUALITY_PLAN, "quality-plan.json")
    require(isinstance(value, dict) and set(value) == {"schema", "frozen_utc", "commands", "targeted"},
            "quality-plan envelope differs")
    require(value["schema"] == "ole2_0555_quality_plan_v1", "quality-plan schema differs")
    parse_time(value["frozen_utc"], "quality-plan.frozen_utc")
    for lane in ("commands", "targeted"):
        commands = value[lane]
        require(isinstance(commands, dict) and commands, f"quality-plan.{lane} is empty")
        require(list(commands) == list(commands), f"quality-plan.{lane} is malformed")
        for name, command in commands.items():
            require(isinstance(name, str) and name and isinstance(command, list)
                    and command and all(isinstance(item, str) and item for item in command),
                    f"quality-plan.{lane}.{name} command is malformed")
    return {"path": rel(QUALITY_PLAN), "sha256": sha(QUALITY_PLAN),
            "commands": {lane: list(value[lane]) for lane in ("commands", "targeted")}}


def validate_host_sidecar(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict) and set(value) == {
        "observed_utc", "compiler_processes", "scope",
    }, f"{label} host sidecar differs")
    parse_time(value["observed_utc"], f"{label}.observed_utc")
    require(value["scope"] == RECEIPT_HOST_SCOPE
            and isinstance(value["compiler_processes"], list),
            f"{label} host scope differs")
    for index, process in enumerate(value["compiler_processes"]):
        require(isinstance(process, dict) and set(process) == {"pid", "comm", "cwd"},
                f"{label}.compiler_processes[{index}] differs")
        require(isinstance(process["pid"], int) and process["pid"] > 0
                and process["comm"] in ("cargo", "rustc")
                and isinstance(process["cwd"], str) and process["cwd"],
                f"{label}.compiler_processes[{index}] is malformed")


def validate_receipt(value: object, label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict) and set(value) == {
        "schema", "stage", "command", "start_utc", "end_utc", "seconds", "exit_code",
        "execution_stage", "execution_manifest_sha256", "binary_sha256",
        "source_manifest_sha256", "workspace_lock_sha256", "workspace_lock_binding_sha256",
        "script_sha256", "plan_sha256", "environment", "artifacts",
    }, f"{label} receipt schema differs")
    start, end = interval(value, label)
    require(isinstance(value["command"], list) and value["command"]
            and all(isinstance(item, str) and item for item in value["command"]),
            f"{label}.command is malformed")
    require(isinstance(value["exit_code"], int) and not isinstance(value["exit_code"], bool),
            f"{label}.exit_code is malformed")
    require(value["schema"] == "ole2_0555_run_receipt_v1"
            and value["stage"] in (*STAGES, "final")
            and value["execution_stage"] in (*STAGES, "final"),
            f"{label}.execution_stage differs")
    for field in ("execution_manifest_sha256", "source_manifest_sha256",
                  "workspace_lock_sha256", "workspace_lock_binding_sha256",
                  "script_sha256", "plan_sha256"):
        check_hash(value[field], f"{label}.{field}")
    if value["binary_sha256"] is not None:
        check_hash(value["binary_sha256"], f"{label}.binary_sha256")
    environment = value["environment"]
    require(isinstance(environment, dict) and set(environment) == {
        "TMPDIR", "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
        "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
        "GLIBC_TUNABLES",
    }, f"{label}.environment differs")
    require(all(item is None or isinstance(item, str) for item in environment.values()),
            f"{label}.environment values differ")
    artifacts = value["artifacts"]
    require(isinstance(artifacts, dict), f"{label}.artifacts is missing")
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).name == name,
                f"{label}.artifacts has a non-local name")
        check_hash(digest, f"{label}.artifacts.{name}")
    return start, end


def validate_run_binding(value: dict[str, object], label: str, *, stage: str,
                         execution_stage: str, execution_manifest_sha: str,
                         source_manifest_sha: str, script: Path,
                         binary_sha: str | None = None) -> None:
    """Bind the output and live execution identities carried by run.py."""
    require(value["stage"] == stage
            and value["execution_stage"] == execution_stage
            and value["execution_manifest_sha256"] == execution_manifest_sha
            and value["source_manifest_sha256"] == source_manifest_sha
            and value["workspace_lock_sha256"] == sha(REPO / "Cargo.lock")
            and value["workspace_lock_binding_sha256"] == sha(LOCK_BINDING)
            and value["script_sha256"] == sha(script)
            and value["plan_sha256"] == sha(PLAN),
            f"{label} source/lock binding differs")
    if binary_sha is None:
        require(value["binary_sha256"] is None, f"{label} binary binding differs")
    else:
        require(value["binary_sha256"] == binary_sha, f"{label} binary binding differs")


def validate_artifacts(folder: Path, stem: str, receipt: dict[str, object],
                       expected: set[str], label: str) -> None:
    artifacts = receipt["artifacts"]
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label}.artifacts inventory differs")
    actual = {
        item.name for item in folder.iterdir()
        if item.is_file() and not item.is_symlink()
        and item.name in expected
    }
    require(actual == expected, f"{label} artifact inventory differs")
    for name, digest in artifacts.items():
        path = folder / name
        require(not path.is_symlink() and path.is_file(),
                f"{label}/{name} is not a regular file")
        require(sha(path) == digest, f"{label}/{name} digest differs")
        if name.endswith(".host.json"):
            validate_host_sidecar(path, f"{label}/{name}")


def expected_build_command(kind: str) -> list[str]:
    require(kind in KINDS, f"unknown build kind {kind}")
    command = ["env", f"TMPDIR={TARGET_STRING}/tmp", "CARGO_BUILD_JOBS=2",
               "CARGO_INCREMENTAL=0", "cargo", "build", "--release", "--locked",
               "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
               "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else ""),
               "--target-dir", TARGET_STRING]
    if kind == "alloc":
        command += ["--features", "allocator-metrics"]
    return command


def descriptor_spec(kind: str) -> tuple[str, str]:
    require(kind in KINDS, f"unknown descriptor kind {kind}")
    return f"binary-{kind}.json", kind


def validate_binary_descriptor(stage: str, kind: str, manifest_sha: str,
                               cleanup: dict[str, object] | None = None) -> dict[str, object]:
    name, scratch_name = descriptor_spec(kind)
    path = need(HERE / stage / name, f"{stage}/{name}")
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == {
        "path", "sha256", "bytes", "build_receipt_sha256", "source_manifest_sha256",
        "workspace_lock_sha256",
    }, f"{rel(path)} inventory differs")
    expected_path = SCRATCH_ROOT / stage / scratch_name
    require(value["path"] == str(expected_path), f"{rel(path)} path differs")
    digest = check_hash(value["sha256"], f"{rel(path)}.sha256")
    require(isinstance(value["bytes"], int) and value["bytes"] > 0,
            f"{rel(path)}.bytes differs")
    require(value["source_manifest_sha256"] == manifest_sha,
            f"{rel(path)} source binding differs")
    require(value["workspace_lock_sha256"] == sha(REPO / "Cargo.lock"),
            f"{rel(path)} workspace lock binding differs")
    receipt_path = need(HERE / stage / f"build-{kind}.receipt.json",
                        f"{stage}/build-{kind}.receipt.json")
    require(value["build_receipt_sha256"] == sha(receipt_path),
            f"{rel(path)} build receipt hash differs")
    if os.path.lexists(expected_path):
        require(expected_path.is_file() and not expected_path.is_symlink(),
                f"{rel(path)} retained binary is unsafe")
        require(sha(expected_path) == digest and expected_path.stat().st_size == value["bytes"],
                f"{rel(path)} binary custody differs")
    else:
        require(cleanup is not None and cleanup.get("owned_paths_absent") is True
                and cleanup.get("accessible_process_references") == [],
                f"{rel(path)} retained binary is missing without cleanup custody")
        hashes = cleanup.get("binary_sha256_by_kind")
        require(isinstance(hashes, dict) and hashes.get(f"{stage}/{kind}") == digest,
                f"{rel(path)} cleanup binary custody differs")
    return {"stage": stage, "kind": kind, "path": str(expected_path),
            "sha256": digest, "bytes": value["bytes"],
            "descriptor_path": rel(path), "descriptor_sha256": sha(path),
            "build_receipt_sha256": value["build_receipt_sha256"]}


def validate_build(stage: str, kind: str, manifest_sha: str,
                   cleanup: dict[str, object] | None = None) -> dict[str, object]:
    descriptor = validate_binary_descriptor(stage, kind, manifest_sha, cleanup)
    folder = HERE / stage
    path = need(folder / f"build-{kind}.receipt.json", rel(folder / f"build-{kind}.receipt.json"))
    value = read_json(path, rel(path))
    start, end = validate_receipt(value, rel(path))
    validate_run_binding(value, rel(path), stage=stage, execution_stage=stage,
                         execution_manifest_sha=manifest_sha,
                         source_manifest_sha=manifest_sha, script=RUN)
    require(value["exit_code"] == 0 and value["command"] == expected_build_command(kind),
            f"{rel(path)} binding differs")
    stem = f"build-{kind}"
    validate_artifacts(folder, stem, value,
                       {f"{stem}.host.json", f"{stem}.stdout", f"{stem}.stderr"},
                       rel(path))
    descriptor.update({"start": start, "end": end, "receipt_sha256": sha(path)})
    return descriptor


def expected_capture_jobs(lane: str) -> list[dict[str, object]]:
    require(lane in LANES, f"invalid capture lane {lane}")
    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    groups = plan["groups"]
    names = ("xls", "cfb")
    output: list[dict[str, object]] = []
    for repeat in (1, 2):
        order = names if repeat == 1 else tuple(reversed(names))
        for group in order:
            selection = groups[group]
            config = plan["native"] if lane == "native" else plan["allocation"]
            output.append({"name": f"{lane}-r{repeat}-{group}", "repeat": repeat,
                           "group": group, "selection": selection,
                           "samples": config["samples"], "warmup": config["warmup"]})
    return output


def expected_profile_jobs() -> list[dict[str, object]]:
    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    jobs = plan["profile"]["jobs"]
    output: list[dict[str, object]] = []
    for repeat in (1, 2):
        for job in jobs:
            output.append({"name": f"profile-r{repeat}-{job}", "repeat": repeat,
                           "job": job})
    return output


def expected_capture_command(stage: str, job: dict[str, object],
                             binary: dict[str, object], lane: str) -> list[str]:
    folder = HERE / stage
    selection = job["selection"]
    command = ["taskset", "-c", str(CPU)]
    if lane == "native":
        command += ["/usr/bin/time", "-f",
                    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                    "-o", str(folder / (job["name"] + ".rss.json"))]
    command += [str(binary["path"]), "--case", ",".join(selection["cases"]),
                "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
                "--json", str(folder / (job["name"] + ".json")),
                "--corpus-manifest", str(folder / (job["name"] + ".catalog.json"))]
    if "shapes" in selection:
        command += ["--shape", ",".join(selection["shapes"]),
                    "--payload", selection["payload"]]
    return command


def expected_profile_command(stage: str, job: dict[str, object],
                             binary: dict[str, object]) -> list[str]:
    plan = read_json(PLAN, "plan.json")
    folder = HERE / stage
    selection = ({"cases": ["xls_owned_source_open_one_cell"]}
                 if job["job"] == "xls-owned"
                 else {"cases": ["cfb_open"], "shapes": [job["job"][4:]],
                       "payload": "incompressible"})
    owner = plan["profile"]["cfb_owner" if job["job"].startswith("cfb-") else "xls_owner"]
    command = ["taskset", "-c", str(CPU), "valgrind", "--vgdb=no", "--tool=callgrind",
               "--collect-atstart=no", "--toggle-collect=" + owner,
               "--zero-before=" + owner, "--dump-after=" + owner,
               "--callgrind-out-file=" + str(folder / (job["name"] + ".callgrind"))]
    command += plan["profile"]["instruction_flags"]
    command += [str(binary["path"]), "--case", ",".join(selection["cases"]),
                "--warmup", "0", "--samples", str(plan["profile"]["samples"]),
                "--json", str(folder / (job["name"] + ".json")),
                "--corpus-manifest", str(folder / (job["name"] + ".catalog.json"))]
    if "shapes" in selection:
        command += ["--shape", ",".join(selection["shapes"]),
                    "--payload", selection["payload"]]
    return command


def validate_capture_receipt(stage: str, job: dict[str, object],
                             binary: dict[str, object], manifest_sha: str,
                             baseline_sha: str, candidate_sha: str,
                             lane: str) -> tuple[dt.datetime, dt.datetime]:
    folder = HERE / stage
    path = need(folder / f"{job['name']}.receipt.json", rel(folder / f"{job['name']}.receipt.json"))
    value = read_json(path, rel(path))
    start, end = validate_receipt(value, rel(path))
    execution = "candidate" if stage == "candidate" or job["repeat"] == 2 else "baseline"
    expected_manifest = candidate_sha if execution == "candidate" else baseline_sha
    command = expected_capture_command(stage, job, binary, lane)
    artifacts = {f"{job['name']}.json", f"{job['name']}.catalog.json",
                 f"{job['name']}.stdout", f"{job['name']}.stderr",
                 f"{job['name']}.host.json"}
    if lane == "native":
        artifacts.add(f"{job['name']}.rss.json")
    validate_run_binding(value, rel(path), stage=stage, execution_stage=execution,
                         execution_manifest_sha=expected_manifest,
                         source_manifest_sha=manifest_sha, script=RUN,
                         binary_sha=binary["sha256"])
    require(value["exit_code"] == 0 and value["command"] == command,
            f"{rel(path)} binding differs")
    validate_artifacts(folder, job["name"], value, artifacts, rel(path))
    return start, end


def validate_profile_receipt(stage: str, job: dict[str, object],
                             binary: dict[str, object], manifest_sha: str,
                             baseline_sha: str, candidate_sha: str
                             ) -> tuple[dt.datetime, dt.datetime]:
    folder = HERE / stage
    path = need(folder / f"{job['name']}.receipt.json", rel(folder / f"{job['name']}.receipt.json"))
    value = read_json(path, rel(path))
    start, end = validate_receipt(value, rel(path))
    # Profiles follow the same frozen ABBA order as the native and allocator
    # lanes.  In particular, baseline-r2 remains in the baseline output
    # folder and uses the retained baseline binary, while the live checkout
    # and execution manifest are candidate.  Folder names alone must not be
    # used to infer this binding.
    execution = "candidate" if stage == "baseline" and job["repeat"] == 2 else stage
    expected_manifest = candidate_sha if execution == "candidate" else baseline_sha
    artifacts = {f"{job['name']}.json", f"{job['name']}.catalog.json",
                 f"{job['name']}.callgrind", f"{job['name']}.stdout",
                 f"{job['name']}.stderr", f"{job['name']}.host.json"}
    count = 5 if job["job"] == "xls-owned" else 6
    artifacts.update(f"{job['name']}.callgrind.{number}"
                     for number in range(1, count + 1))
    validate_run_binding(value, rel(path), stage=stage, execution_stage=execution,
                         execution_manifest_sha=expected_manifest,
                         source_manifest_sha=manifest_sha, script=RUN,
                         binary_sha=binary["sha256"])
    require(value["exit_code"] == 0
            and value["command"] == expected_profile_command(stage, job, binary),
            f"{rel(path)} binding differs")
    validate_artifacts(folder, job["name"], value, artifacts, rel(path))
    return start, end


def validate_assembly(stage: str, manifest_sha: str, binary: dict[str, object]
                      ) -> tuple[list[tuple[dt.datetime, dt.datetime, str]], int]:
    """Validate the dynamic symbol/disassembly evidence produced by inspect_assembly.py."""
    folder = HERE / stage
    index_path = need(folder / "assembly-index.json", f"{stage}/assembly-index.json")
    index = read_json(index_path, rel(index_path))
    require(isinstance(index, dict) and set(index) == {
        "schema", "plan_sha256", "binary_sha256", "source_manifest_sha256",
        "script_sha256", "rows", "requested_owners", "scope",
    }, f"{rel(index_path)} envelope differs")
    script = need(HERE / "inspect_assembly.py", "inspect_assembly.py")
    plan = validate_plan()["value"]
    assembly_plan = plan["assembly"]
    owners = assembly_plan["requested_owners"]
    require(index["schema"] == "ole2_0555_assembly_v1"
            and index["plan_sha256"] == sha(PLAN)
            and index["binary_sha256"] == binary["sha256"]
            and index["source_manifest_sha256"] == manifest_sha
            and index["script_sha256"] == sha(script)
            and index["requested_owners"] == owners
            and index["scope"] == ASSEMBLY_OUTPUT_SCOPE,
            f"{rel(index_path)} binding differs")
    rows = index["rows"]
    require(isinstance(rows, list) and rows, f"{rel(index_path)} rows are empty")
    expected_receipts = {
        f"assembly-{number}.receipt.json" for number in range(len(rows))
    }
    actual_receipts = {
        item.name for item in folder.glob("assembly-*.receipt.json")
        if item.is_file() and not item.is_symlink()
    }
    require(actual_receipts == expected_receipts,
            f"{rel(index_path)} disassembly receipt inventory differs")
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    symbols_path = need(folder / "symbols.receipt.json", f"{stage}/symbols.receipt.json")
    symbols = read_json(symbols_path, rel(symbols_path))
    start, end = validate_receipt(symbols, rel(symbols_path))
    validate_run_binding(symbols, rel(symbols_path), stage=stage,
                         execution_stage=stage,
                         execution_manifest_sha=manifest_sha,
                         source_manifest_sha=manifest_sha, script=RUN,
                         binary_sha=binary["sha256"])
    require(symbols["exit_code"] == 0
            and symbols["command"] == ["nm", "-S", "--defined-only", binary["path"]],
            f"{rel(symbols_path)} binding differs")
    validate_artifacts(folder, "symbols", symbols,
                       {"symbols.host.json", "symbols.stdout", "symbols.stderr"},
                       rel(symbols_path))
    intervals.append((start, end, "symbols"))
    names: list[str] = []
    symbols_seen: set[str] = set()
    for index_number, row in enumerate(rows):
        require(isinstance(row, dict) and set(row) == {
            "name", "symbol", "address_hex", "size_bytes", "receipt_sha256",
        }, f"{rel(index_path)} row {index_number} differs")
        name = row["name"]
        require(name == f"assembly-{index_number}" and name not in names,
                f"{rel(index_path)} row name order differs")
        symbol = row["symbol"]
        require(isinstance(symbol, str) and symbol and symbol not in symbols_seen,
                f"{rel(index_path)} row symbol differs")
        require("litchi_cfb" in symbol
                and any(owner in symbol for owner in owners),
                f"{rel(index_path)} row symbol is outside requested owners")
        require(isinstance(row["address_hex"], str)
                and re.fullmatch(r"[0-9A-Fa-f]+", row["address_hex"]) is not None,
                f"{rel(index_path)} row address differs")
        require(isinstance(row["size_bytes"], int) and row["size_bytes"] > 0,
                f"{rel(index_path)} row size differs")
        check_hash(row["receipt_sha256"], f"{rel(index_path)} row receipt hash")
        receipt_path = need(folder / f"{name}.receipt.json", rel(folder / f"{name}.receipt.json"))
        require(sha(receipt_path) == row["receipt_sha256"],
                f"{rel(index_path)} row receipt hash differs")
        receipt = read_json(receipt_path, rel(receipt_path))
        start, end = validate_receipt(receipt, rel(receipt_path))
        validate_run_binding(receipt, rel(receipt_path), stage=stage,
                             execution_stage=stage,
                             execution_manifest_sha=manifest_sha,
                             source_manifest_sha=manifest_sha, script=RUN,
                             binary_sha=binary["sha256"])
        require(receipt["exit_code"] == 0 and receipt["command"] == [
                    "objdump", "-d", "--disassemble=" + symbol, binary["path"]
                ], f"{rel(receipt_path)} binding differs")
        validate_artifacts(folder, name, receipt,
                           {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"},
                           rel(receipt_path))
        intervals.append((start, end, name))
        names.append(name)
        symbols_seen.add(symbol)
    return intervals, len(rows)


def stage_inventory(stage: str, profiles: bool, assembly: bool) -> set[str]:
    expected = {"source-manifest.json", "source.patch"}
    for kind in KINDS:
        stem = f"build-{kind}"
        expected |= {f"{stem}.receipt.json", f"{stem}.host.json",
                     f"{stem}.stdout", f"{stem}.stderr", f"binary-{kind}.json"}
    for lane in LANES:
        for job in expected_capture_jobs(lane):
            stem = job["name"]
            expected |= {f"{stem}.receipt.json", f"{stem}.host.json",
                         f"{stem}.json", f"{stem}.catalog.json",
                         f"{stem}.stdout", f"{stem}.stderr"}
            if lane == "native":
                expected.add(f"{stem}.rss.json")
    if profiles:
        for job in expected_profile_jobs():
            stem = job["name"]
            expected |= {f"{stem}.receipt.json", f"{stem}.host.json",
                         f"{stem}.json", f"{stem}.catalog.json",
                         f"{stem}.stdout", f"{stem}.stderr", f"{stem}.callgrind"}
            count = 5 if job["job"] == "xls-owned" else 6
            expected |= {f"{stem}.callgrind.{number}"
                         for number in range(1, count + 1)}
        # The profile analyzer retains its stage report beside the raw
        # Callgrind receipts.  Instruction attribution may be incorporated in
        # that report or supplied by a later consumer; it is not a second
        # mandatory artifact for this fresh 0555 freeze.
        expected.add("profile-analysis.json")
    if assembly:
        expected.add("assembly-index.json")
        expected |= {"symbols.receipt.json", "symbols.host.json",
                     "symbols.stdout", "symbols.stderr"}
        for item in (HERE / stage).glob("assembly-*.receipt.json"):
            stem = item.name.removesuffix(".receipt.json")
            expected |= {f"{stem}.receipt.json", f"{stem}.host.json",
                         f"{stem}.stdout", f"{stem}.stderr"}
    folder = need(HERE / stage, f"{stage} stage", directory=True)
    actual = {item.name for item in folder.iterdir()
              if item.is_file() and not item.is_symlink()}
    require(actual == expected, f"{stage} raw artifact inventory differs")
    return expected


def validate_stage_capture(stage: str, baseline_sha: str,
                           candidate_sha: str,
                           cleanup: dict[str, object] | None = None,
                           profiles: bool | None = None) -> dict[str, object]:
    plan = validate_plan()["value"]
    source = validate_stage_source(stage, plan, git_tree_manifest(plan["revision"]) if stage == "baseline" else None)
    manifest_sha = source["manifest_sha256"]
    builds = {kind: validate_build(stage, kind, manifest_sha, cleanup) for kind in KINDS}
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = [
        (row["start"], row["end"], f"build-{kind}") for kind, row in builds.items()
    ]
    rows: dict[str, int] = {}
    for lane in LANES:
        entries = expected_capture_jobs(lane)
        for job in entries:
            times = validate_capture_receipt(
                stage, job, builds["alloc" if lane == "alloc" else "normal"],
                manifest_sha, baseline_sha, candidate_sha, lane,
            )
            intervals.append((times[0], times[1], job["name"]))
        rows[lane] = len(entries)
    # 0555 profile evidence is a mandatory matched owner-attribution lane.  A
    # partial profile directory cannot silently turn into a valid native-only
    # campaign; missing receipts remain an incomplete stage.
    profile_paths = list((HERE / stage).glob("profile-*.receipt.json"))
    has_profiles = True
    require(profile_paths, f"{stage} profile evidence is missing")
    for job in expected_profile_jobs():
        times = validate_profile_receipt(
            stage, job, builds["normal"], manifest_sha, baseline_sha, candidate_sha,
        )
        intervals.append((times[0], times[1], job["name"]))
    rows[PROFILE_LANE] = len(expected_profile_jobs())
    assembly_paths = list((HERE / stage).glob("assembly-*.receipt.json"))
    assembly_index = (HERE / stage / "assembly-index.json").exists()
    require(assembly_index and assembly_paths,
            f"{stage} assembly evidence is missing")
    require(bool(assembly_paths) == assembly_index,
            f"{stage} assembly evidence is only partially retained")
    assembly_intervals, assembly_count = validate_assembly(stage, manifest_sha, builds["normal"])
    intervals.extend(assembly_intervals)
    stage_inventory(stage, has_profiles, True)
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{stage} build/capture receipts overlap")
    return {"stage": stage, "manifest_sha256": manifest_sha,
            "builds": {kind: {key: value for key, value in row.items()
                               if key not in ("start", "end")}
                       for kind, row in builds.items()},
            "counts": rows, "intervals": len(intervals), "profiles": has_profiles,
            "assembly_rows": assembly_count}


def validate_captures(cleanup: dict[str, object] | None = None,
                      stage: str | None = None) -> dict[str, object]:
    plan = validate_plan()["value"]
    base = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, base["manifest"])
    stages = (stage,) if stage is not None else STAGES
    for selected in stages:
        require(selected in STAGES, f"unknown capture stage {selected}")
    rows = {selected: validate_stage_capture(
        selected, base["manifest_sha256"], candidate["manifest_sha256"], cleanup
    ) for selected in stages}
    if stage is None:
        intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
        for selected in STAGES:
            folder = HERE / selected
            for path in folder.glob("*.receipt.json"):
                value = read_json(path, rel(path))
                start, end = interval(value, rel(path))
                intervals.append((start, end, rel(path)))
        ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
        require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
                "cross-stage build/capture receipts overlap")
    return {"status": "pass", "stages": rows,
            "source": {"baseline": base["manifest_sha256"],
                       "candidate": candidate["manifest_sha256"]}}


def find_one(names: tuple[str, ...], label: str) -> Path:
    found = [HERE / name for name in names if (HERE / name).exists()]
    require(all(path.is_file() and not path.is_symlink() for path in found),
            f"{label} output is unsafe")
    if not found:
        raise IncompleteError(f"{label} document is missing")
    require(len(found) == 1, f"multiple {label} documents are retained")
    return found[0]


def find_profile_comparison() -> Path:
    """Find the canonical comparison, allowing only a byte-identical alias."""
    found = [HERE / name for name in PROFILE_NAMES if (HERE / name).exists()]
    require(all(path.is_file() and not path.is_symlink() for path in found),
            "profile output is unsafe")
    if not found:
        raise IncompleteError("profile comparison document is missing")
    canonical = HERE / "profile-analysis.json"
    if len(found) > 1:
        data = found[0].read_bytes()
        require(all(path.read_bytes() == data for path in found[1:]),
                "multiple non-identical profile comparison documents are retained")
    return canonical if canonical in found else found[0]


def manifest_hashes() -> dict[str, str]:
    plan = validate_plan()["value"]
    base = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, base["manifest"])
    return {"baseline": base["manifest_sha256"], "candidate": candidate["manifest_sha256"]}


def find_hash(value: object, names: tuple[str, ...]) -> str | None:
    if isinstance(value, dict):
        for name in names:
            if name in value and isinstance(value[name], str) and SHA256_RE.fullmatch(value[name]):
                return value[name]
        for child in value.values():
            found = find_hash(child, names)
            if found:
                return found
    elif isinstance(value, list):
        for child in value:
            found = find_hash(child, names)
            if found:
                return found
    return None


def replay_report(script: Path, function: str, expected: dict[str, object]) -> None:
    """Replay a pure reader in an isolated interpreter; never invoke a writing CLI."""
    code = (
        "import importlib.util,json,sys; from pathlib import Path; "
        "p=Path(sys.argv[1]); sys.path.insert(0,str(p.parent)); "
        "spec=importlib.util.spec_from_file_location('verified_analysis',p); "
        "m=importlib.util.module_from_spec(spec); sys.modules[spec.name]=m; "
        "spec.loader.exec_module(m); "
        "v=m.analyze() if sys.argv[2]=='analyze' else "
        "m.compare(m.load_plan() if hasattr(m,'load_plan') else m.plan_data()); "
        "print(json.dumps(v,sort_keys=True))"
    )
    result = subprocess.run([sys.executable, "-B", "-c", code, str(script), function],
                            cwd=REPO, capture_output=True, text=True)
    require(result.returncode == 0, f"{rel(script)} replay failed: {result.stderr[-2000:]}")
    replayed = json.loads(result.stdout)
    if script.name == "analyze_metrics.py" and not os.path.lexists(TARGET):
        validate_cleanup()
        for stage in STAGES:
            for kind in KINDS:
                require(expected["binaries"][stage][kind]["present"] is True
                        and replayed["binaries"][stage][kind]["present"] is False,
                        "postcleanup metrics binary presence differs")
                replayed["binaries"][stage][kind]["present"] = True
    require(replayed == expected, f"{rel(script)} deterministic replay differs")


def validate_metrics() -> dict[str, object]:
    path = find_one(METRICS_NAMES, "metrics")
    value = read_json(path, rel(path))
    require(isinstance(value, dict), f"{rel(path)} is not an object")
    require(set(value) == {
        "schema", "status", "stage", "scope", "priority", "performance_claim",
        "disposition", "plan_sha256", "run_sha256", "frozen_inputs",
        "prior_scope_note", "helpers", "source_manifests", "host_scope", "binaries",
        "job_counts", "row_counts", "sample_counts", "identities", "native",
        "allocation", "matched_identity", "comparisons", "repeat_drift",
        "repeat_drift_over_five_percent", "main_gates", "limits",
    }, f"{rel(path)} field inventory differs")
    require(value.get("schema") == METRICS_SCHEMA,
            f"{rel(path)} schema differs")
    require(value.get("status") == "pass", f"{rel(path)} is not a passing analysis")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("run_sha256") == sha(RUN),
            f"{rel(path)} plan/driver bindings differ")
    plan = validate_plan()["value"]
    base = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, base["manifest"])
    source_manifests = value.get("source_manifests")
    require(isinstance(source_manifests, dict)
            and source_manifests.get("baseline") == base["manifest"]
            and source_manifests.get("candidate") == candidate["manifest"]
            and set(source_manifests) == set(STAGES),
            f"{rel(path)} source manifest bindings differ")
    require(value.get("host_scope") == RECEIPT_HOST_SCOPE,
            f"{rel(path)} host scope differs")
    frozen = value.get("frozen_inputs")
    require(isinstance(frozen, dict)
            and set(frozen) == {"adr-manifest.json", "plan.json", "run.py",
                                "workspace-lock.json"},
            f"{rel(path)} frozen input bindings differ")
    frozen_record = validate_frozen_inputs()
    require(frozen == frozen_record["files"],
            f"{rel(path)} frozen input values differ")
    helpers = value.get("helpers")
    require(isinstance(helpers, dict)
            and set(helpers) == {"tools/summarize_crud_baseline.py",
                                 "tools/validate_perf_corpus_binding.py"},
            f"{rel(path)} helper inventory differs")
    for name, digest in helpers.items():
        check_hash(digest, f"{rel(path)} helper {name}")
        helper = safe_repo_path(name, f"{rel(path)} helper {name}")
        require(sha(helper) == digest, f"{rel(path)} helper changed: {name}")
    gates = value.get("main_gates")
    gate_names = {
        "primary_xls_p50", "primary_xls_mean", "native_xls_controls",
        "native_cfb_controls", "native_rss", "allocation",
    }
    require(isinstance(gates, dict)
            and set(gates) == gate_names | {
                "all_frozen_main_gates_pass", "external_controls_required",
            }, f"{rel(path)} main gate inventory differs")
    for name in gate_names:
        group = gates[name]
        require(isinstance(group, dict)
                and set(group) == {"name", "description", "pass", "check_count", "checks"}
                and group["name"] == name
                and isinstance(group["description"], str) and group["description"].strip()
                and isinstance(group["pass"], bool)
                and isinstance(group["check_count"], int) and group["check_count"] >= 0
                and isinstance(group["checks"], list)
                and group["check_count"] == len(group["checks"]),
                f"{rel(path)} {name} gate differs")
        for index, check in enumerate(group["checks"]):
            require(isinstance(check, dict) and set(check) == {
                "lane", "group", "case", "shape", "repeat", "metric", "criterion",
                "pass", "baseline", "candidate", "change_percent",
            } and isinstance(check["pass"], bool)
                    and isinstance(check["criterion"], str) and check["criterion"].strip()
                    and isinstance(check["metric"], str) and check["metric"].strip(),
                    f"{rel(path)} {name} check {index} differs")
    external = gates["external_controls_required"]
    require(isinstance(external, dict)
            and set(external) == {"status", "validated_here", "required", "reason"}
            and external["status"] == "pending"
            and external["validated_here"] is False
            and external["required"] == ["profile", "correctness", "quality"]
            and isinstance(external["reason"], str) and external["reason"].strip(),
            f"{rel(path)} external gate declaration differs")
    overall = gates.get("all_frozen_main_gates_pass")
    require(isinstance(overall, bool), f"{rel(path)} overall gate is missing")
    replay_report(HERE / "analyze_metrics.py", "analyze", value)
    return {"path": rel(path), "sha256": sha(path), "schema": value["schema"],
            "gate_passed": overall}


def validate_profile_stage_report(stage: str, manifest: dict[str, str]) -> dict[str, object]:
    path = need(HERE / stage / "profile-analysis.json",
                f"{stage}/profile-analysis.json")
    value = read_json(path, rel(path))
    require(isinstance(value, dict)
            and value.get("schema") == PROFILE_STAGE_SCHEMA
            and value.get("status") == "pass" and value.get("stage") == stage
            and value.get("plan") == "plan.json"
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("performance_claim") == "diagnostic-only"
            and value.get("scope") == validate_plan()["value"]["scope"],
            f"{rel(path)} envelope differs")
    profiles = value.get("profiles")
    jobs = expected_profile_jobs()
    require(isinstance(profiles, list) and len(profiles) == len(jobs)
            and [item.get("name") for item in profiles] == [item["name"] for item in jobs],
            f"{rel(path)} profile matrix differs")
    for item, job in zip(profiles, jobs):
        require(isinstance(item, dict) and item.get("repeat") == job["repeat"]
                and item.get("stage") == stage
                and item.get("group") == ("xls-owned" if job["job"] == "xls-owned"
                                            else job["job"])
                and item.get("kind") in ("xls", "cfb"),
                f"{rel(path)} profile identity differs")
        receipt = item.get("receipt")
        execution = "candidate" if stage == "baseline" and job["repeat"] == 2 else stage
        receipt_path = need(HERE / stage / f"{job['name']}.receipt.json",
                            f"{stage}/{job['name']}.receipt.json")
        raw_receipt = read_json(receipt_path, rel(receipt_path))
        require(isinstance(receipt, dict)
                and receipt.get("path") == rel(receipt_path)
                and receipt.get("sha256") == sha(receipt_path)
                and receipt.get("source_manifest_sha256") == sha(HERE / stage / "source-manifest.json")
                and receipt.get("execution_stage") == execution
                and receipt.get("execution_manifest_sha256") == raw_receipt.get("execution_manifest_sha256")
                and receipt.get("binary_sha256") == raw_receipt.get("binary_sha256")
                and receipt.get("command") == raw_receipt.get("command")
                and receipt.get("artifacts") == raw_receipt.get("artifacts"),
                f"{rel(path)} receipt binding differs")
    require(value.get("profile_count") == len(jobs)
            and value.get("timed_constructor_dump_count") == 40
            and value.get("setup_dump_count") == 6,
            f"{rel(path)} dump counts differ")
    helpers = value.get("helpers")
    require(isinstance(helpers, dict)
            and helpers.get("run.py") == sha(RUN)
            and helpers.get("workspace-lock.json") == sha(LOCK_BINDING)
            and all(isinstance(key, str) and isinstance(digest, str)
                    and SHA256_RE.fullmatch(digest) is not None
                    for key, digest in helpers.items()),
            f"{rel(path)} driver helper binding differs")
    mechanism = value.get("mechanism_gate")
    required_mechanism = {
        "all_timed_owner_edges_classified",
        "all_cfb_setup_roles_classified_from_positive_ancestry",
        "selected_owner_self_and_inclusive_separate",
        "physical_target_absence_is_indeterminate",
        "claim_sector_inline_fallback_retained",
        "moved_work_targets_retained_separately",
        "no_native_or_adoption_claim",
    }
    require(isinstance(mechanism, dict) and set(mechanism) == required_mechanism
            and all(mechanism[key] is True for key in required_mechanism),
            f"{rel(path)} mechanism evidence differs")
    validation = value.get("validation")
    required_validation = {
        "plan_receipt_binary_and_execution_bindings",
        "profile_reports_and_artifacts_hashed",
        "positive_owner_ancestry_used_for_scope",
        "setup_not_in_timed_aggregates",
        "raw_target_self_direct_inclusive_fields_separate",
        "missing_or_inline_values_are_not_zero",
        "no_native_allocation_or_adoption_claim",
    }
    require(isinstance(validation, dict) and set(validation) == required_validation
            and all(validation[key] is True for key in required_validation),
            f"{rel(path)} validation differs")
    return {"path": rel(path), "sha256": sha(path), "profiles": len(profiles)}


def validate_profile(metrics: dict[str, object] | None = None) -> dict[str, object]:
    plan = validate_plan()["value"]
    base = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, base["manifest"])
    stages = {
        "baseline": validate_profile_stage_report("baseline", base["manifest"]),
        "candidate": validate_profile_stage_report("candidate", candidate["manifest"]),
    }
    path = find_one(PROFILE_NAMES, "profile")
    value = read_json(path, rel(path))
    require(isinstance(value, dict), f"{rel(path)} is not an object")
    require(value.get("schema") == PROFILE_COMPARISON_SCHEMA
            and value.get("status") == "pass"
            and value.get("plan") == "plan.json"
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("scope") == plan["scope"]
            and value.get("performance_claim") == "diagnostic-only"
            and value.get("stage_selection") == ["baseline", "candidate"],
            f"{rel(path)} envelope differs")
    embedded = value.get("stages")
    require(isinstance(embedded, dict) and set(embedded) == set(STAGES),
            f"{rel(path)} stage reports are missing")
    for stage in STAGES:
        require(embedded[stage] == read_json(HERE / stage / "profile-analysis.json", stage),
                f"{stage} embedded profile differs from canonical report")
        require(embedded[stage].get("schema") == PROFILE_STAGE_SCHEMA
                and embedded[stage].get("status") == "pass"
                and embedded[stage].get("stage") == stage
                and embedded[stage].get("plan_sha256") == sha(PLAN),
                f"{rel(path)} embedded {stage} report differs")
    validation = value.get("validation")
    require(isinstance(validation, dict)
            and validation.get("both_stages_valid") is True
            and validation.get("matched_owner_edges_compared") is True
            and validation.get("physical_marker_targets_retained") is True
            and validation.get("inline_or_absent_not_zero_work") is True
            and validation.get("no_native_or_adoption_claim") is True,
            f"{rel(path)} validation differs")
    comparison = value.get("comparison")
    require(isinstance(comparison, dict), f"{rel(path)} comparison is missing")
    require(set(comparison) == {
        "rows", "moved_work", "mechanism_gate", "validation", "interpretation",
    }, f"{rel(path)} comparison field inventory differs")
    mechanism_gate = comparison.get("mechanism_gate")
    require(isinstance(mechanism_gate, dict),
            f"{rel(path)} comparison mechanism gate is missing")
    gate_names = (
        "xls_owner_inclusive_ir_decreases_each_repeat",
        "physical_reconciliation_self_ir_decreases_each_repeat",
        "physical_reconciliation_present_for_all_xls_rows",
        "all_target_rows_kept_separate",
        "claim_sector_inline_or_out_of_line_explicit",
        "moved_work_retained_without_elimination_claim",
        "no_native_or_adoption_claim",
    )
    require(set(mechanism_gate) == set(gate_names) | {
                "physical_reconciliation_self_ir_status",
            }
            and all(isinstance(mechanism_gate.get(name), bool) for name in gate_names)
            and mechanism_gate.get("physical_reconciliation_self_ir_status") in {
                "proven", "indeterminate_assembly_required"
            },
            f"{rel(path)} comparison mechanism gate values differ")
    profile_gate = all(mechanism_gate[name] for name in gate_names)
    replay_report(HERE / "analyze_profiles.py", "compare", value)
    # Assembly is independently mandatory for the physical attribution
    # boundary, while a separate instruction consumer remains optional.  A
    # missing or inlined profile target therefore stays indeterminate and can
    # never satisfy the physical mechanism gate by itself.
    require(all(stages[stage]["profiles"] == len(expected_profile_jobs())
                for stage in STAGES), "profile stage coverage differs")
    return {"path": rel(path), "sha256": sha(path),
            "required": True, "gate_passed": profile_gate,
            "stages": stages}


def quality_source_manifest(stage: str, digest: str, label: str) -> Path:
    """Find the immutable source snapshot named by a quality attempt."""
    require(stage in (*STAGES, "final"), f"{label} stage differs")
    candidates = [HERE / stage / "source-manifest.json"]
    if stage == "candidate":
        candidates.extend(
            path for path in sorted((HERE / "candidate-attempts").glob("*/stage/source-manifest.json"))
            if path.is_file() and not path.is_symlink()
        )
    matches = [path for path in candidates if path.is_file() and not path.is_symlink()
               and sha(path) == digest]
    require(len(matches) == 1, f"{label} source snapshot is not uniquely retained")
    selected = matches[0]
    if selected != HERE / stage / "source-manifest.json":
        plan = validate_plan()["value"]
        base = git_tree_manifest(str(plan["revision"]))
        validate_candidate_variant(selected, selected.parent / "source.patch", base,
                                   rel(selected.parent))
    return selected


def validate_quality_attempt(path: Path, final_sha: str) -> dict[str, object]:
    need(path, rel(path), directory=True)
    inputs_path = need(path / "inputs.json", rel(path / "inputs.json"))
    inputs = read_json(inputs_path, rel(inputs_path))
    require(isinstance(inputs, dict) and set(inputs) == {
        "schema", "created_utc", "stage", "mode", "source_manifest_sha256",
        "plan_sha256", "script_sha256", "commands", "workspace_lock_sha256",
    }, f"{rel(inputs_path)} envelope differs")
    require(inputs["schema"] == QUALITY_INPUT_SCHEMA
            and inputs["stage"] in (*STAGES, "final")
            and inputs["mode"] in ("targeted", "commands")
            and inputs["plan_sha256"] == sha(QUALITY_PLAN)
            and inputs["script_sha256"] == sha(HERE / "quality.py")
            and inputs["workspace_lock_sha256"] == json.loads(LOCK_BINDING.read_text())["sha256"],
            f"{rel(inputs_path)} binding differs")
    parse_time(inputs["created_utc"], f"{rel(inputs_path)}.created_utc")
    plan = read_json(QUALITY_PLAN, "quality-plan.json")
    require(inputs["commands"] == plan[inputs["mode"]], f"{rel(inputs_path)} commands differ")
    input_stage_sha = check_hash(inputs["source_manifest_sha256"],
                                 f"{rel(inputs_path)}.source_manifest_sha256")
    input_stage_path = quality_source_manifest(inputs["stage"], input_stage_sha,
                                               rel(inputs_path))
    result_path = need(path / "result.json", rel(path / "result.json"))
    result = read_json(result_path, rel(result_path))
    require(isinstance(result, dict) and set(result) == {
        "schema", "status", "stage", "mode", "source_manifest_sha256",
        "inputs_sha256", "completed_utc", "rows",
    }, f"{rel(result_path)} envelope differs")
    require(result["schema"] == QUALITY_RESULT_SCHEMA
            and result["status"] in ("pass", "failed")
            and result["stage"] == inputs["stage"] and result["mode"] == inputs["mode"]
            and result["source_manifest_sha256"] == input_stage_sha
            and result["inputs_sha256"] == sha(inputs_path),
            f"{rel(result_path)} binding differs")
    parse_time(result["completed_utc"], f"{rel(result_path)}.completed_utc")
    commands = plan["commands"]
    if result["mode"] == "targeted":
        commands = plan["targeted"]
    rows = result["rows"]
    require(isinstance(rows, list), f"{rel(result_path)} rows are malformed")
    failed = result["status"] == "failed"
    if failed:
        # Failed attempts are retained history.  Validate their completed
        # prefix and preserve the non-zero receipt that stopped the run.
        require(0 < len(rows) < len(commands) + 1,
                f"{rel(result_path)} failed row count differs")
        expected_names = list(commands)[:len(rows)]
        require([row.get("name") for row in rows] == expected_names,
                f"{rel(result_path)} failed row order differs")
    else:
        require(len(rows) == len(commands), f"{rel(result_path)} row count differs")
    is_final = result["stage"] == "final" and result["mode"] == "commands"
    if is_final:
        require(result["source_manifest_sha256"] == final_sha,
                f"{rel(result_path)} final source binding differs")
    expected_names = list(commands) if not failed else list(commands)[:len(rows)]
    require([row.get("name") for row in rows] == expected_names,
            f"{rel(result_path)} row order differs")
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for row in rows:
        expected_exit = 0 if not failed else None
        require(isinstance(row, dict) and set(row) == {"name", "path", "receipt_sha256", "exit_code"}
                and isinstance(row["exit_code"], int)
                and not isinstance(row["exit_code"], bool)
                and (row["exit_code"] == expected_exit if expected_exit is not None
                     else row["exit_code"] != 0),
                f"{rel(result_path)} row differs")
        relative = safe_relative(row["path"], f"{rel(result_path)} row.path")
        folder = bundle_path(relative, f"{rel(result_path)} row.path")
        require(folder == path / row["name"], f"{rel(result_path)} row path differs")
        receipt_path = need(folder / "receipt.json", rel(folder / "receipt.json"))
        require(row["receipt_sha256"] == sha(receipt_path),
                f"{rel(result_path)} receipt hash differs")
        receipt = read_json(receipt_path, rel(receipt_path))
        start, end = interval(receipt, rel(receipt_path))
        intervals.append((start, end, rel(receipt_path)))
        require(isinstance(receipt, dict) and set(receipt) == {
            "schema", "command", "start_utc", "end_utc", "seconds", "exit_code",
            "source_stable", "source_manifest_sha256", "inputs_sha256", "environment",
            "artifacts",
        }, f"{rel(receipt_path)} envelope differs")
        require(receipt["schema"] == QUALITY_RECEIPT_SCHEMA
                and receipt["command"] == commands[row["name"]]
                and receipt["exit_code"] == row["exit_code"]
                and (receipt["exit_code"] == 0 if not failed else receipt["exit_code"] != 0)
                and receipt["source_stable"] is True
                and receipt["source_manifest_sha256"] == input_stage_sha
                and receipt["inputs_sha256"] == sha(inputs_path),
                f"{rel(receipt_path)} binding differs")
        environment = receipt["environment"]
        require(isinstance(environment, dict) and set(environment) == {
            "TMPDIR", "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
            "RUSTDOCFLAGS", "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD",
        }, f"{rel(receipt_path)} environment differs")
        artifacts = receipt["artifacts"]
        require(isinstance(artifacts, dict) and set(artifacts) == {"stdout", "stderr"},
                f"{rel(receipt_path)} artifacts differ")
        validate_artifacts(folder, "", {"artifacts": artifacts}, {"stdout", "stderr"},
                           rel(receipt_path))
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{rel(result_path)} quality command receipts overlap")
    raw = {item.name for item in path.iterdir() if item.is_file() and not item.is_symlink()}
    require(raw == {"inputs.json", "result.json"}, f"{rel(path)} raw inventory differs")
    for row in rows:
        folder = path / row["name"]
        actual = {item.name for item in folder.iterdir()
                  if item.is_file() and not item.is_symlink()}
        require(actual == {"stdout", "stderr", "receipt.json"},
                f"{rel(folder)} raw inventory differs")
    return {"path": rel(result_path), "sha256": sha(result_path),
            "inputs_sha256": sha(inputs_path), "rows": len(rows),
            "status": result["status"], "final": is_final}


def validate_quality(final_source_sha: str | None = None) -> dict[str, object]:
    if final_source_sha is None:
        final_source_sha = sha(need(FINAL / "source-manifest.json", "final/source-manifest.json"))
    canonical = need(QUALITY, "quality.json")
    canonical_value = read_json(canonical, "quality.json")
    require(isinstance(canonical_value, dict) and canonical_value.get("status") == "pass",
            "quality.json is not a passing result")
    attempts = need(QUALITY_ATTEMPTS, "quality-attempts", directory=True)
    passing: list[dict[str, object]] = []
    for item in sorted(attempts.iterdir()):
        if item.is_dir() and not item.is_symlink():
            result = validate_quality_attempt(item, final_source_sha)
            if result.get("status") == "pass" and result.get("final") is True \
                    and result["sha256"] == sha(canonical):
                passing.append(result)
    require(len(passing) == 1, "quality.json is not an exact passing final attempt")
    result_path = bundle_path(passing[0]["path"], "quality selected path")
    require(canonical.read_bytes() == result_path.read_bytes(),
            "quality.json is not byte-identical to its selected final attempt")
    return {"path": rel(canonical), "sha256": sha(canonical),
            "attempt": passing[0]["path"], "rows": passing[0]["rows"]}


def validate_review(metrics: dict[str, object], profile: dict[str, object]) -> dict[str, object]:
    path = find_one(REVIEW_NAMES, "adverse review")
    value = read_json(path, rel(path))
    require(isinstance(value, dict)
            and value.get("schema") == REVIEW_SCHEMA and value.get("status") == "pass"
            and value.get("complete") is True
            and value.get("all_diagnostic_rows_retained") is True
            and isinstance(value.get("adoption_allowed"), bool), "review envelope differs")
    for key, expected in (("metrics_sha256", metrics["sha256"]),
                          ("profile_sha256", profile["sha256"]),
                          ("quality_sha256", sha(HERE / "quality.json"))):
        require(value.get(key) == expected, f"review {key} differs")
    document = read_json(HERE / "metrics-analysis.json", "review metrics")
    comparisons = document.get("comparisons")
    require(isinstance(comparisons, dict)
            and isinstance(comparisons.get("adverse_over_five_percent"), list)
            and isinstance(document.get("repeat_drift_over_five_percent"), list),
            "review metrics comparison rows are missing")
    expected = [comparisons["adverse_over_five_percent"],
                document["repeat_drift_over_five_percent"]]
    require(len(value["groups"]) == 2, "review group inventory differs")
    ids = set()
    for group, rows in zip(value["groups"], expected):
        require(isinstance(group, dict) and isinstance(group.get("rows"), list),
                "review group rows are missing")
        require(len(group["rows"]) == len(rows), "review row count differs")
        originals = []
        for row in group["rows"]:
            require(isinstance(row, dict), "review row is not an object")
            for key in ("id", "classification", "interpretation", "disposition"):
                require(isinstance(row.get(key), str) and row[key].strip(), "review interpretation missing")
            require(row["id"] not in ids, "duplicate review id")
            ids.add(row["id"])
            require(isinstance(row.get("original"), dict), "review original row is missing")
            originals.append(json.dumps(row["original"], sort_keys=True))
        require(sorted(originals) == sorted(json.dumps(row, sort_keys=True) for row in rows),
                "review is not one-for-one canonical rows")
    require(value.get("counts") == {"metrics_adverse": len(expected[0]),
            "metrics_drift": len(expected[1]), "reviewed_flags": len(ids)},
            "review counts differ")
    return {"path": rel(path), "sha256": sha(path),
            "adoption_allowed": value["adoption_allowed"], "reviewed_flags": len(ids)}


def validate_documentation() -> dict[str, object]:
    value = read_json(DOCUMENTATION, "documentation-manifest.json")
    require(isinstance(value, dict) and set(value) == {"schema", "scope", "files"},
            "documentation-manifest envelope differs")
    require(value["schema"] == DOCUMENTATION_SCHEMA
            and isinstance(value["scope"], str) and value["scope"].strip(),
            "documentation-manifest identity differs")
    files = value["files"]
    require(isinstance(files, dict) and files and list(files) == sorted(files),
            "documentation-manifest file map differs")
    prefix = HERE.relative_to(REPO).as_posix()
    for name, digest in files.items():
        safe_relative(name, "documentation-manifest path")
        require(name.startswith("docs/") and name != prefix
                and not name.startswith(prefix + "/"),
                f"documentation-manifest path is not external: {name}")
        check_hash(digest, f"documentation-manifest hash {name}")
        path = safe_repo_path(name, f"documentation-manifest path {name}")
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"documentation-manifest file differs: {name}")
    require(DOCUMENTATION_REQUIRED in files,
            f"documentation-manifest omits {DOCUMENTATION_REQUIRED}")
    return {"path": rel(DOCUMENTATION), "sha256": sha(DOCUMENTATION),
            "files": dict(files)}


def validate_decision(metrics: dict[str, object], profile: dict[str, object],
                      quality: dict[str, object], review: dict[str, object]) -> dict[str, object]:
    path = find_one(DECISION_NAMES, "decision")
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == {
        "schema", "status", "observed_utc", "scope", "disposition",
        "adoption_allowed", "plan_sha256", "metrics_sha256", "profile_sha256",
        "quality_sha256", "review_sha256", "source_manifest_sha256",
    }, f"{rel(path)} field inventory differs")
    require(value.get("schema") == DECISION_SCHEMA
            and value.get("status") == "pass"
            and isinstance(value.get("scope"), str) and value["scope"].strip(),
            f"{rel(path)} schema/status differs")
    disposition = value.get("disposition")
    require(disposition in ("accepted", "rejected"), f"{rel(path)} disposition differs")
    require(isinstance(value.get("adoption_allowed"), bool),
            f"{rel(path)} adoption_allowed is missing")
    parse_time(value.get("observed_utc"), f"{rel(path)}.observed_utc")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("metrics_sha256") == metrics["sha256"]
            and value.get("profile_sha256") == profile["sha256"]
            and value.get("quality_sha256") == quality["sha256"]
            and value.get("review_sha256") == review["sha256"]
            and check_hash(value.get("source_manifest_sha256"),
                           f"{rel(path)}.source_manifest_sha256"),
            f"{rel(path)} evidence bindings differ")
    adoption = bool(metrics["gate_passed"]) and bool(profile["required"])
    adoption = adoption and bool(profile["gate_passed"]) \
        and bool(review["adoption_allowed"])
    require(value["adoption_allowed"] is adoption
            and disposition == ("accepted" if adoption else "rejected"),
            f"{rel(path)} disposition does not follow gates")
    return {"path": rel(path), "sha256": sha(path),
            "disposition": disposition, "adoption_allowed": adoption,
            "source_manifest_sha256": value["source_manifest_sha256"]}


def validate_final_source(disposition: str) -> dict[str, object]:
    plan = validate_plan()["value"]
    base = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, base["manifest"])
    final = validate_stage_source("final", plan, base["manifest"], candidate["manifest"])
    expected = candidate["manifest"] if disposition == "accepted" else base["manifest"]
    require(final["manifest"] == expected, "final source does not match disposition")
    require(current_source_manifest() == expected, "live source does not match disposition")
    return {"stage": "final", "manifest_sha256": final["manifest_sha256"],
            "entries": final["entries"], "patch_sha256": final["patch_sha256"]}


def validate_campaign(cleanup: dict[str, object] | None = None) -> dict[str, object]:
    inputs = validate_inputs()
    plan = validate_plan()["value"]
    base = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, base["manifest"])
    captures = validate_captures(cleanup=cleanup)
    metrics = validate_metrics()
    profile = validate_profile(metrics)
    quality = validate_quality()
    review = validate_review(metrics, profile)
    decision = validate_decision(metrics, profile, quality, review)
    final = validate_final_source(decision["disposition"])
    require(final["manifest_sha256"] == decision["source_manifest_sha256"],
            "decision final source binding differs")
    binaries = {stage: captures["stages"][stage]["builds"] for stage in STAGES}
    return {"inputs": {"plan_sha256": inputs["plan"]["sha256"]},
            "source": {"baseline": base["manifest_sha256"],
                       "candidate": candidate["manifest_sha256"],
                       "final": final["manifest_sha256"]},
            "captures": {"stages": list(captures["stages"])},
            "metrics": {"sha256": metrics["sha256"], "gate_passed": metrics["gate_passed"]},
            "profile": {"sha256": profile["sha256"], "required": profile["required"],
                        "gate_passed": profile["gate_passed"]},
            "quality": {"sha256": quality["sha256"], "rows": quality["rows"]},
            "review": {"sha256": review["sha256"],
                       "adoption_allowed": review["adoption_allowed"]},
            "decision": decision, "owned_binaries": binaries,
            "documentation": validate_documentation()}


def validate_precleanup() -> dict[str, object]:
    require(os.path.lexists(TARGET) and TARGET.is_dir() and not TARGET.is_symlink(),
            "precleanup requires the owned target directory to remain present")
    campaign = validate_campaign()
    return {"status": "pass", "phase": "precleanup", "target_present": True,
            "observed_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
            "campaign": campaign}


def expected_binary_keys() -> set[str]:
    return {f"{stage}/{kind}" for stage in STAGES for kind in KINDS}


def validate_precleanup_record(value: object, cleanup: dict[str, object]) -> dict[str, object]:
    require(isinstance(value, dict) and set(value) == {
        "schema", "status", "scope", "result",
    }, "precleanup-verification envelope differs")
    require(value["schema"] == SCHEMA and value["status"] == "pass"
            and value["scope"] == "precleanup", "precleanup-verification identity differs")
    result = value["result"]
    require(isinstance(result, dict) and result.get("status") == "pass"
            and result.get("phase") == "precleanup"
            and result.get("target_present") is True,
            "precleanup-verification result differs")
    parse_time(result.get("observed_utc"), "precleanup-verification.observed_utc")
    parse_time(cleanup["observed_utc"], "cleanup.observed_utc")
    require(result["observed_utc"] <= cleanup["observed_utc"],
            "precleanup verification postdates cleanup")
    campaign = result.get("campaign")
    require(isinstance(campaign, dict) and set(campaign) == {
        "inputs", "source", "captures", "metrics", "profile", "quality", "review",
        "decision", "owned_binaries", "documentation",
    }, "precleanup campaign inventory differs")
    owned = campaign["owned_binaries"]
    require(isinstance(owned, dict) and set(owned) == set(STAGES),
            "precleanup binary inventory differs")
    for stage in STAGES:
        require(set(owned[stage]) == set(KINDS), "precleanup binary kinds differ")
        for kind in KINDS:
            row = owned[stage][kind]
            key = f"{stage}/{kind}"
            require(row["stage"] == stage and row["kind"] == kind
                    and row["sha256"] == cleanup["binary_sha256_by_kind"][key]
                    and row["descriptor_sha256"] == cleanup["binary_descriptor_sha256_by_kind"][key],
                    f"precleanup binary custody differs: {key}")
    return {"status": "pass", "sha256": sha(PRECLEANUP),
            "observed_utc": result["observed_utc"]}


def validate_cleanup() -> dict[str, object]:
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "observed_utc", "plan_sha256", "target", "removed",
        "owned_paths_absent", "accessible_process_references",
        "process_reference_scope", "python_cache_absent",
        "binary_sha256_by_kind", "binary_descriptor_sha256_by_kind",
        "precleanup_verification_sha256", "scope",
    }, "cleanup.json envelope differs")
    require(value["schema"] == CLEANUP_SCHEMA and value["plan_sha256"] == sha(PLAN)
            and value["target"] == TARGET_STRING and value["removed"] == [TARGET_STRING]
            and value["owned_paths_absent"] is True
            and value["accessible_process_references"] == []
            and value["process_reference_scope"] == PROCESS_REFERENCE_SCOPE
            and value["python_cache_absent"] is True
            and value["scope"] == (
                "Remove only owned change0555 build/retained-binary target after passing precleanup and exact binary hash checks."
            ), "cleanup target custody differs")
    parse_time(value["observed_utc"], "cleanup.observed_utc")
    require(not os.path.lexists(TARGET), "cleanup claims target absent but owned path remains")
    require(not any("__pycache__" in item.parts for item in HERE.rglob("*")),
            "Python bytecode cache remains after cleanup")
    hashes = value["binary_sha256_by_kind"]
    descriptor_hashes = value["binary_descriptor_sha256_by_kind"]
    require(isinstance(hashes, dict) and set(hashes) == expected_binary_keys()
            and isinstance(descriptor_hashes, dict)
            and set(descriptor_hashes) == expected_binary_keys(),
            "cleanup binary custody inventory differs")
    pre_path = need(PRECLEANUP, "precleanup-verification.json")
    require(value["precleanup_verification_sha256"] == sha(pre_path),
            "cleanup precleanup-verification binding differs")
    validate_precleanup_record(read_json(pre_path, rel(pre_path)), value)
    plan = validate_plan()["value"]
    base = validate_stage_source("baseline", plan)
    candidate = validate_stage_source("candidate", plan, base["manifest"])
    checked = []
    for stage, source in (("baseline", base), ("candidate", candidate)):
        for kind in KINDS:
            row = validate_binary_descriptor(stage, kind, source["manifest_sha256"], value)
            key = f"{stage}/{kind}"
            require(hashes[key] == row["sha256"] and descriptor_hashes[key] == row["descriptor_sha256"],
                    f"cleanup binary digest differs: {key}")
            checked.append({"key": key, "sha256": row["sha256"],
                            "descriptor_sha256": row["descriptor_sha256"]})
    return {"path": rel(CLEANUP), "sha256": sha(CLEANUP),
            "owned_paths_absent": True, "binary_custody": checked,
            "precleanup_verification_sha256": value["precleanup_verification_sha256"]}


def sealed_inventory() -> dict[str, str]:
    result: dict[str, str] = {}
    for item in HERE.rglob("*"):
        require(not item.is_symlink(), f"sealed bundle contains a symlink: {rel(item)}")
        require("__pycache__" not in item.parts,
                f"sealed bundle contains Python cache: {rel(item)}")
        if item.is_dir():
            continue
        require(item.is_file(), f"sealed bundle contains a non-file: {rel(item)}")
        if item == SEAL:
            continue
        result[item.relative_to(HERE).as_posix()] = sha(item)
    require("verify.py" in result and "verifier-schema.md" in result,
            "sealed inventory omits verifier or schema")
    return dict(sorted(result.items()))


def validate_seal() -> dict[str, object]:
    text = read_text(SEAL, "SHA256SUMS")
    require(text.endswith("\n") and text.strip(), "SHA256SUMS is empty or lacks a final newline")
    expected: dict[str, str] = {}
    for line in text.splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and fields[1] and fields[1] != "SHA256SUMS",
                "SHA256SUMS line differs")
        check_hash(fields[0], "SHA256SUMS digest")
        safe_relative(fields[1], "SHA256SUMS path")
        require(fields[1] not in expected, "SHA256SUMS inventory is duplicated")
        expected[fields[1]] = fields[0]
    require(list(expected) == sorted(expected), "SHA256SUMS paths are not sorted")
    actual = sealed_inventory()
    require(expected == actual, "SHA256SUMS inventory differs")
    return {"path": rel(SEAL), "sha256": sha(SEAL), "entries": len(expected)}


def validate_all() -> dict[str, object]:
    cleanup = validate_cleanup()
    campaign = validate_campaign(cleanup=read_json(CLEANUP, "cleanup.json"))
    seal = validate_seal()
    return {"status": "pass", "phase": "sealed", "target_present": False,
            "campaign": campaign, "cleanup": cleanup, "seal": seal}


def run_component(component: str, stage: str | None = None) -> dict[str, object]:
    if component == "inputs":
        return validate_inputs()
    if component == "plan":
        plan = validate_plan()
        return {key: value for key, value in plan.items() if key != "value"}
    if component == "source":
        plan = validate_plan()["value"]
        base = validate_stage_source("baseline", plan)
        candidate = validate_stage_source("candidate", plan, base["manifest"])
        return {"baseline": {key: value for key, value in base.items() if key != "manifest"},
                "candidate": {key: value for key, value in candidate.items() if key != "manifest"}}
    if component in ("captures", "capture"):
        return validate_captures(stage=stage)
    if component == "metrics":
        return validate_metrics()
    if component in ("profile", "profiles"):
        return validate_profile()
    if component == "quality":
        return validate_quality()
    if component in ("review", "adverse-review"):
        metrics = validate_metrics()
        profile = validate_profile(metrics)
        return validate_review(metrics, profile)
    if component in ("decision", "disposition"):
        metrics = validate_metrics()
        profile = validate_profile(metrics)
        quality = validate_quality()
        review = validate_review(metrics, profile)
        return validate_decision(metrics, profile, quality, review)
    if component == "documentation":
        return validate_documentation()
    if component == "precleanup":
        return validate_precleanup()
    if component == "cleanup":
        return validate_cleanup()
    if component == "seal":
        return validate_seal()
    if component == "all":
        return validate_all()
    raise VerificationError(f"unknown component: {component}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", "-c", default="all", choices=(
        "inputs", "plan", "source", "captures", "capture", "metrics", "profile",
        "profiles", "quality", "review", "adverse-review", "decision", "disposition",
        "documentation", "precleanup", "cleanup", "seal", "all",
    ))
    parser.add_argument("--stage", choices=STAGES)
    parser.add_argument("--strict", action="store_true",
                        help="return nonzero when selected evidence is incomplete")
    args = parser.parse_args(argv)
    try:
        result = run_component(args.component, args.stage)
        print(json.dumps({"schema": SCHEMA, "status": "pass", "scope": args.component,
                          "result": result}, indent=2, sort_keys=True))
        return 0
    except IncompleteError as error:
        print(json.dumps({"schema": SCHEMA, "status": "incomplete", "scope": args.component,
                          "error": str(error)}, indent=2, sort_keys=True))
        return 2 if args.strict else 0
    except (VerificationError, OSError, KeyError, TypeError, AttributeError, IndexError) as error:
        print(json.dumps({"schema": SCHEMA, "status": "fail", "scope": args.component,
                          "error": str(error)}, indent=2, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
