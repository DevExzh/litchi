#!/usr/bin/env python3
"""Fail-closed, read-only custody verifier for the 0552 XLSX experiment.

The capture drivers and the two analyzers are deliberately independent of this
module.  This verifier checks their frozen inputs, exact serial matrix,
source/binary/corpus bindings, quality receipts, and final disposition.  It
never builds, captures, edits the checkout, or writes an evidence report
inside this bundle.  Missing later-stage evidence is reported as
incomplete; it is never converted into a passing or zero-filled result.

The component selectors are useful while the campaign is being assembled.
precleanup checks all evidence expected before removing the owned target.
all additionally requires an explicit final disposition, quality run,
cleanup record, and recursive seal.
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
from typing import Any

sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
CAPTURE = HERE / "capture.py"
GUARDED = HERE / "guarded_capture.py"
CHECK_ATTEMPT = HERE / "check_attempt.py"
FROZEN = HERE / "frozen-inputs.json"
SUPPLEMENTAL = HERE / "supplemental-inputs.json"
ADR = HERE / "adr-manifest.json"
HOST = HERE / "host.json"
QUALITY_PLAN = HERE / "quality-plan.json"
ANALYSIS_INPUTS = HERE / "analysis-inputs.json"
BASELINE_CORRECTNESS = HERE / "baseline-correctness.json"
LOCK_BINDING = HERE / "workspace-lock.json"
LOCK_COPY = HERE / "workspace-Cargo.lock"
LOCK_COMPLETION = HERE / "baseline-lock-completion.json"
BASELINE = HERE / "baseline"
CANDIDATE = HERE / "candidate"
FINAL = HERE / "final"
PUBLIC = HERE / "public-test-sources"
PUBLIC_EXACT = HERE / "public-exact-test-sources"
CANDIDATE_ATTEMPTS = HERE / "candidate-attempts"
BASELINE_RESTORE = HERE / "baseline-restore-for-public.json"
CANDIDATE_BINDING = HERE / "candidate-source-binding.json"
PREFLIGHT_SUMMARY = HERE / "preflight-summary-draft07.json"
TARGET = Path("/home/zhuhe/litchi-goal-0552-target")
SCRATCH_ROOT = TARGET / "retained"
SEAL = HERE / "SHA256SUMS"

METRICS = HERE / "analyze_metrics.py"
GUARDS = HERE / "analyze_guards.py"
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
DECISION_NAMES = (
    "decision.json",
    "disposition.json",
    "admission.json",
)
ANALYZER_METRICS_SHA256 = (
    "af6dc11cd31160ac267f695d76ff2422b19cce80a5376b8e704fdff9addca5cc"
)
ANALYZER_GUARDS_SHA256 = (
    "c30dbe68e0d0db4ff15ab214f75eeda0a4e5d3ca042f646c6e46d79f5d7bf817"
)
ANALYZER_GUARDS_AMENDED_SHA256 = (
    "6714a465703be541c969a1256ad3c1d5d21f81387f94aa65d613b03890a79dbc"
)
ANALYZER_AMENDMENT = HERE / "analyzer-amendment.json"
ANALYZER_AMENDMENT_SCHEMA = "xlsx_0552_analyzer_serialization_amendment_v1"
ANALYZER_AMENDMENT_PATCH_SHA256 = (
    "c3f920980e937b56454f21d1291accef68c303710424fad4b97a6cbc829e9e5f"
)
ANALYSIS_HELPER_HASHES = {
    "docs/performance/results/change-0546/integration/analyze.py":
        "c380298762c64f104c4aaaf1bf53fae7e613bc7aa9148fe1d6f1da19b2eb9614",
    "docs/performance/results/change-0521/analyze.py":
        "322357892b496ef09ca01aa69bb5a8708182a9a771a4541e3b21a6885a7616ad",
    "tools/validate_perf_corpus_binding.py":
        "e20abbd1220623f387283ff5cbb0cec95ea57edad2e320bd22d93b268cfa9558",
}

STAGES = ("baseline", "candidate")
SHAPES = ("medium", "dense-sparse", "noncompact", "vendor-extension")
CASES = (
    "xlsx_source_backed_cell_values_one_edit_save",
    "xlsx_source_backed_cell_values_one_percent_edit_save",
    "xlsx_source_backed_managed_cell_values_one_edit_save",
    "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
)
GUARD_SHAPES = ("medium", "dense-sparse")
GUARD_CASES = ("valid", "late-validator", "late-raw")
CAP_SIZES = (1, 2, 160, 164, 256)
PROFILE_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
PROFILE_OWNER = "litchi_xlsx::cell_values::source::MultiSourceEdit::commit"
PROFILE_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
CPU = 2
WORKSPACE_LOCK_SHA256 = (
    "9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a"
)
HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
PREFLIGHT_PRIORITY = "OLE2/OOXML first; ODF deferred; iWork excluded"
PROFILE_DECISION_SCHEMA = "xlsx_0552_profile_decision_v1"
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
DRIFT_PERCENT = 5.0
EXPECTED_ERRORS = {
    "late-validator": "value-only edits refuse attribute 'future' on 'c'",
    "late-raw": "invalid worksheet boolean 'maybe'",
}


class VerificationError(ValueError):
    """Malformed, contradictory, or out-of-scope retained evidence."""


class IncompleteError(VerificationError):
    """Evidence needed by the selected component has not arrived."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def rel(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError:
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
        raise VerificationError(f"{label} is not an ISO-8601 timestamp") from error
    require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    require(end > start, f"{label} interval is inverted")
    seconds = value.get("seconds")
    require(
        isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
        and math.isfinite(float(seconds)) and seconds > 0,
        f"{label}.seconds is not positive",
    )
    wall = (end - start).total_seconds()
    require(
        abs(float(seconds) - wall) <= max(0.25, wall * 0.02 + 0.05),
        f"{label}.seconds does not match UTC interval",
    )
    return start, end


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and "." not in path.parts and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def source_name(name: str) -> bool:
    return (
        name in {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
        or name.startswith((".cargo/", "crates/", "tools/perf-baseline/"))
    )


def git(args: list[str], *, env: dict[str, str] | None = None,
        input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(
            args, cwd=REPO, env=env, input=input_data, stderr=subprocess.PIPE
        )
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise VerificationError(
            f"Git command failed ({' '.join(args)}): "
            f"{detail.decode(errors='replace')[-2000:]}"
        ) from error


def git_tree_manifest(revision: str) -> dict[str, str]:
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
        require(
            len(fields) == 3 and fields[0].decode() == oid and fields[1] == b"blob",
            "Git source object response is malformed",
        )
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
    return result


def parse_index(index: Path) -> dict[str, str]:
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    raw = git(["git", "ls-files", "-s", "-z"], env=env)
    result: dict[str, str] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        try:
            metadata, encoded = item.split(b"\t", 1)
            fields = metadata.split()
            name = encoded.decode("utf-8")
        except (UnicodeDecodeError, ValueError) as error:
            raise VerificationError("private source index entry is malformed") from error
        require(len(fields) == 3 and fields[0] in (b"100644", b"100755"),
                "private source index mode is unexpected")
        require(name not in result, f"private source index repeats {name}")
        result[name] = fields[1].decode()
    return result


def index_hashes(index: Path, oids: set[str]) -> dict[str, str]:
    if not oids:
        return {}
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    response = git(
        ["git", "cat-file", "--batch"],
        env=env,
        input_data=("\n".join(sorted(oids)) + "\n").encode(),
    )
    result: dict[str, str] = {}
    position = 0
    while position < len(response):
        end = response.find(b"\n", position)
        require(end >= 0, "private source object response is truncated")
        fields = response[position:end].split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "private source object response is malformed")
        position = end + 1
        length = int(fields[2])
        data = response[position:position + length]
        require(len(data) == length, "private source object is truncated")
        result[fields[0].decode()] = hashlib.sha256(data).hexdigest()
        position += length
        require(response[position:position + 1] == b"\n",
                "private source object separator is missing")
        position += 1
    return result


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
    return dict(sorted(result.items()))


def validate_frozen_inputs() -> dict[str, Any]:
    value = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(value, dict) and set(value) == {"frozen_utc", "files"},
            "frozen-inputs envelope differs")
    parse_time(value["frozen_utc"], "frozen-inputs.frozen_utc")
    expected_names = {
        "docs/performance/results/change-0552/plan.json",
        "docs/performance/results/change-0552/run.py",
        "docs/performance/results/change-0552/capture.py",
        "docs/performance/results/change-0552/adr-manifest.json",
    }
    files = value["files"]
    require(isinstance(files, dict) and set(files) == expected_names,
            "frozen-inputs file inventory differs")
    for name, digest in files.items():
        check_hash(digest, f"frozen-inputs.files.{name}")
        path = REPO / name
        need(path, name)
        require(sha(path) == digest, f"frozen input changed: {name}")
    return {"sha256": sha(FROZEN), "files": dict(sorted(files.items()))}


def validate_supplemental_inputs() -> dict[str, Any]:
    value = read_json(SUPPLEMENTAL, "supplemental-inputs.json")
    require(isinstance(value, dict)
            and set(value) == {"frozen_utc", "scope", "files"},
            "supplemental-inputs envelope differs")
    parse_time(value["frozen_utc"], "supplemental-inputs.frozen_utc")
    expected_scope = (
        "Workspace-lock custody before first workspace build; per-child lock checks "
        "on later pipelines; quality commands frozen before use. Original capture inputs unchanged."
    )
    require(value["scope"] == expected_scope, "supplemental-inputs scope differs")
    expected_names = {
        "docs/performance/results/change-0552/guarded_capture.py",
        "docs/performance/results/change-0552/workspace-lock.json",
        "docs/performance/results/change-0552/workspace-Cargo.lock",
        "docs/performance/results/change-0552/quality-plan.json",
    }
    files = value["files"]
    require(isinstance(files, dict) and set(files) == expected_names,
            "supplemental-inputs file inventory differs")
    for name, digest in files.items():
        check_hash(digest, f"supplemental-inputs.files.{name}")
        path = REPO / name
        need(path, name)
        require(sha(path) == digest, f"supplemental input changed: {name}")
    return {"sha256": sha(SUPPLEMENTAL), "scope": value["scope"],
            "files": dict(sorted(files.items()))}


def validate_guard_analyzer_amendment() -> dict[str, Any] | None:
    """Validate the one-line post-capture guard-analyzer serialization fix.

    ``analysis-inputs.json`` intentionally retains the pre-capture analyzer
    digest.  A post-capture serialization-only correction is admissible only
    through this exact custody record, with the original and amended source
    witnesses, patch, and the failed pre-fix analyzer receipt all bound.
    """
    current_sha = sha(GUARDS)
    if current_sha == ANALYZER_GUARDS_SHA256 and not ANALYZER_AMENDMENT.exists():
        return None
    if current_sha == ANALYZER_GUARDS_SHA256:
        raise VerificationError(
            "guard analyzer amendment record exists but the live analyzer is still the original"
        )
    require(current_sha == ANALYZER_GUARDS_AMENDED_SHA256,
            "guard analyzer is neither the frozen nor the authorized amended revision")
    if not ANALYZER_AMENDMENT.exists():
        raise IncompleteError("guard analyzer amendment custody record is missing")

    value = read_json(ANALYZER_AMENDMENT, "analyzer-amendment.json")
    expected_keys = {
        "schema", "created_utc", "reason", "path", "original_sha256",
        "amended_sha256", "before_path", "after_path", "patch_path",
        "patch_sha256", "original_analysis_inputs_sha256",
        "failed_receipt_path", "failed_receipt_sha256", "scope",
    }
    require(isinstance(value, dict) and set(value) == expected_keys,
            "analyzer amendment envelope differs")
    require(value["schema"] == ANALYZER_AMENDMENT_SCHEMA,
            "analyzer amendment schema differs")
    created = parse_time(value["created_utc"], "analyzer-amendment.created_utc")
    require(value["reason"] == (
        "Completed guard analysis cannot serialize datetime objects in build interval metadata. "
        "Convert only those two timestamps to ISO strings; no numerical, gate, receipt-validation "
        "or ordering logic changes."
    ), "analyzer amendment reason differs")
    require(value["scope"] == (
        "Post-capture output serialization correction; original frozen analysis inputs and all "
        "captures retained unchanged."
    ), "analyzer amendment scope differs")
    expected_paths = {
        "path": "docs/performance/results/change-0552/analyze_guards.py",
        "before_path": (
            "docs/performance/results/change-0552/"
            "analyzer-amendments/guard-serialization-01/before.py"
        ),
        "after_path": (
            "docs/performance/results/change-0552/"
            "analyzer-amendments/guard-serialization-01/after.py"
        ),
        "patch_path": (
            "docs/performance/results/change-0552/"
            "analyzer-amendments/guard-serialization-01/change.patch"
        ),
        "failed_receipt_path": (
            "docs/performance/results/change-0552/"
            "check-attempts/matched-guards-analysis-01/receipt.json"
        ),
    }
    for name, expected in expected_paths.items():
        require(value[name] == expected,
                f"analyzer amendment {name} differs")
    for name in (
        "original_sha256", "amended_sha256", "patch_sha256",
        "original_analysis_inputs_sha256", "failed_receipt_sha256",
    ):
        check_hash(value[name], f"analyzer-amendment.{name}")
    require(value["original_sha256"] == ANALYZER_GUARDS_SHA256,
            "analyzer amendment original digest differs from frozen inputs")
    require(value["amended_sha256"] == ANALYZER_GUARDS_AMENDED_SHA256,
            "analyzer amendment amended digest is not the authorized revision")
    require(value["patch_sha256"] == ANALYZER_AMENDMENT_PATCH_SHA256,
            "analyzer amendment patch digest is not the authorized patch")
    require(value["original_analysis_inputs_sha256"] == sha(ANALYSIS_INPUTS),
            "analyzer amendment does not bind the original analysis-inputs record")

    before_path = REPO / value["before_path"]
    after_path = REPO / value["after_path"]
    patch_path = REPO / value["patch_path"]
    require(sha(before_path) == value["original_sha256"],
            "analyzer amendment before witness differs")
    require(sha(after_path) == value["amended_sha256"],
            "analyzer amendment after witness differs")
    require(sha(patch_path) == value["patch_sha256"],
            "analyzer amendment patch witness differs")
    require(sha(GUARDS) == value["amended_sha256"],
            "live guard analyzer does not equal the amended witness")

    before_lines = read_text(before_path, rel(before_path)).splitlines()
    after_lines = read_text(after_path, rel(after_path)).splitlines()
    require(len(before_lines) == len(after_lines),
            "analyzer amendment changed the source line count")
    changed = [
        index for index, (before_line, after_line)
        in enumerate(zip(before_lines, after_lines))
        if before_line != after_line
    ]
    require(changed == [632],
            "analyzer amendment is not a single interval serialization change")
    require(before_lines[632] == '            "interval": (start, end)}'
            and after_lines[632] == (
                '            "interval": (start.isoformat(), end.isoformat())}'
            ), "analyzer amendment changed the wrong expression")

    failed_path = REPO / value["failed_receipt_path"]
    expected_manifest, _ = quality_expected_manifest("candidate")
    failed = validate_check_attempt(failed_path.parent, expected_manifest)
    require(failed["receipt_sha256"] == value["failed_receipt_sha256"]
            and failed["exit_code"] == 1
            and failed["command"] == [
                "python3", "-B",
                "docs/performance/results/change-0552/analyze_guards.py",
                "--output", "docs/performance/results/change-0552/guards-analysis.json",
            ], "analyzer amendment failed receipt differs")
    require(parse_time(failed["end_utc"], "analyzer amendment failed receipt end") <= created,
            "analyzer amendment predates its failed receipt")
    return {
        "status": "pass",
        "schema": value["schema"],
        "path": rel(ANALYZER_AMENDMENT),
        "sha256": sha(ANALYZER_AMENDMENT),
        "original_sha256": value["original_sha256"],
        "amended_sha256": value["amended_sha256"],
        "patch_sha256": value["patch_sha256"],
        "original_analysis_inputs_sha256": value["original_analysis_inputs_sha256"],
        "failed_receipt": {
            "path": value["failed_receipt_path"],
            "sha256": failed["receipt_sha256"],
            "exit_code": failed["exit_code"],
            "source_manifest_sha256": failed["source_manifest_sha256"],
        },
    }


def validate_analysis_inputs() -> dict[str, Any]:
    """Validate the separately frozen analyzer and helper dependency set."""
    value = read_json(ANALYSIS_INPUTS, "analysis-inputs.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "frozen_utc", "scope", "files",
    }, "analysis-inputs envelope differs")
    require(value["schema"] == "xlsx_0552_analysis_inputs_v1",
            "analysis-inputs schema differs")
    parse_time(value["frozen_utc"], "analysis-inputs.frozen_utc")
    require(value["scope"] == (
        "Main workflow and guard/cap analyzers frozen before candidate capture; "
        "profile and final decision evidence remain conditional/pending."
    ), "analysis-inputs scope differs")
    expected = {
        "docs/performance/results/change-0552/analyze_metrics.py":
            ANALYZER_METRICS_SHA256,
        "docs/performance/results/change-0552/analyze_guards.py":
            ANALYZER_GUARDS_SHA256,
        "docs/performance/results/change-0552/run.py": sha(RUN),
        "docs/performance/results/change-0552/capture.py": sha(CAPTURE),
        "docs/performance/results/change-0552/plan.json": sha(PLAN),
        **ANALYSIS_HELPER_HASHES,
    }
    files = value["files"]
    require(isinstance(files, dict) and set(files) == set(expected),
            "analysis-inputs file inventory differs")
    guard_amendment = validate_guard_analyzer_amendment()
    for name, digest in expected.items():
        check_hash(files.get(name), f"analysis-inputs.files.{name}")
        require(files[name] == digest,
                f"analysis-inputs frozen digest differs: {name}")
        path = REPO / name
        need(path, name)
        actual = sha(path)
        if name == "docs/performance/results/change-0552/analyze_guards.py":
            expected_actual = (
                guard_amendment["amended_sha256"]
                if guard_amendment is not None else digest
            )
            require(actual == expected_actual,
                    f"analysis input changed: {name}")
        else:
            require(actual == digest, f"analysis input changed: {name}")
    return {"status": "pass", "sha256": sha(ANALYSIS_INPUTS),
            "schema": value["schema"], "files": dict(sorted(files.items())),
            "guard_analyzer_amendment": guard_amendment}


def validate_plan() -> dict[str, Any]:
    frozen = validate_frozen_inputs()
    value = read_json(PLAN, "plan.json")
    require(isinstance(value, dict), "plan is not an object")
    expected_keys = {
        "revision", "created_utc", "previous_turn", "priority", "scope",
        "hypothesis", "cpu", "owned_paths", "cases", "shapes", "native",
        "alloc", "profile", "guard", "cap", "native_order", "allocator_order",
        "admission", "limitations",
    }
    require(set(value) == expected_keys, "plan field inventory differs")
    revision = value["revision"]
    require(isinstance(revision, str) and REVISION_RE.fullmatch(revision),
            "plan revision is malformed")
    try:
        subprocess.run(
            ["git", "cat-file", "-e", revision + "^{commit}"],
            cwd=REPO, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError("plan revision is not a Git commit") from error
    parse_time(value["created_utc"], "plan.created_utc")
    require(value["priority"] ==
            "OLE2/OOXML first; ODF deferred until that goal completes; iWork excluded",
            "plan priority differs")
    require(value["scope"] ==
            "Matched compact source-cell proof experiment for source-backed XLSX MultiSourceEdit",
            "plan scope differs")
    require(value["cpu"] == CPU and value["owned_paths"] == [str(TARGET)],
            "plan execution envelope differs")
    require(value["cases"] == list(CASES) and value["shapes"] == list(SHAPES),
            "plan main matrix differs")
    require(value["native"] == {"repeats": 2, "samples": 200, "warmup": 20},
            "plan native matrix differs")
    require(value["alloc"] == {"repeats": 2, "samples": 20, "warmup": 3},
            "plan allocator matrix differs")
    require(value["profile"] == {
        "repeats": 2, "samples": 1, "warmup": 0, "case": PROFILE_CASE,
        "owner": PROFILE_OWNER, "parent": PROFILE_PARENT,
    }, "plan profile envelope differs")
    require(value["guard"] == {
        "shapes": list(GUARD_SHAPES), "cases": list(GUARD_CASES), "repeats": 2,
        "native_samples": 200, "native_warmup": 20,
        "alloc_samples": 20, "alloc_warmup": 3,
    }, "plan guard matrix differs")
    require(value["cap"] == {
        "sizes": list(CAP_SIZES), "repeats": 2, "samples": 200, "warmup": 20,
    }, "plan cap matrix differs")
    expected_order = [
        "baseline r1", "candidate r1", "candidate r2",
        "retained baseline r2 under candidate source manifest",
    ]
    require(value["native_order"] == expected_order
            and value["allocator_order"] == expected_order,
            "plan ABBA order differs")
    admission = value["admission"]
    require(isinstance(admission, dict) and set(admission) == {
        "primary", "latency_guards", "refusals", "workflow_memory", "noop_memory",
        "allocation", "proof_resources", "correctness", "profile", "disposition",
    }, "plan admission inventory differs")
    expected_admission = {
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
    require(admission == expected_admission, "plan admission text differs")
    require(value["limitations"] == [
        "No hardware/cold/provider/scaling claim from this campaign.",
        "Current single-sheet SourceEdit API is outside measured owner.",
        "Profiles conditional on native/memory pilot; all failures retained.",
        "Historical guards use no-op oracle after planning interval; full no-op publication time is not measured.",
    ], "plan limitations differ")
    return {"status": "pass", "sha256": sha(PLAN), "revision": revision,
            "scope": value["scope"], "priority": value["priority"], "frozen": frozen}


def validate_adr() -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    require(isinstance(value, dict) and set(value) == {"files", "checked_utc", "status"},
            "ADR manifest inventory differs")
    parse_time(value["checked_utc"], "adr-manifest.checked_utc")
    require(value["status"] == "all previously read ADRs unchanged",
            "ADR manifest status differs")
    files = value["files"]
    require(isinstance(files, dict) and files, "ADR manifest files are empty")
    for name, digest in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith("docs/adr/"), f"ADR path is out of scope: {name}")
        check_hash(digest, f"ADR {name}")
        path = REPO / name
        need(path, name)
        require(sha(path) == digest, f"ADR changed: {name}")
    return {"status": "pass", "sha256": sha(ADR), "entries": len(files)}


def validate_host() -> dict[str, Any]:
    value = read_json(HOST, "host.json")
    require(isinstance(value, dict) and set(value) == {
        "utc", "uname", "cpu_affinity", "rustc", "cargo",
    }, "host envelope differs")
    parse_time(value["utc"], "host.utc")
    require(isinstance(value["uname"], str) and value["uname"].strip(),
            "host.uname is empty")
    require(value["cpu_affinity"] == list(range(32)), "host CPU affinity differs")
    for field in ("rustc", "cargo"):
        require(isinstance(value[field], str) and value[field].strip(),
                f"host.{field} is empty")
    return {"status": "pass", "sha256": sha(HOST), "cpu_affinity": value["cpu_affinity"]}


def validate_workspace_lock() -> dict[str, Any]:
    supplemental = validate_supplemental_inputs()
    value = read_json(LOCK_BINDING, "workspace-lock.json")
    require(isinstance(value, dict) and set(value) == {
        "frozen_utc", "path", "sha256", "retained", "reason",
        "baseline_cap_not_started",
    }, "workspace-lock envelope differs")
    frozen_utc = parse_time(value["frozen_utc"], "workspace-lock.frozen_utc")
    require(value["path"] == "Cargo.lock" and value["retained"] == "workspace-Cargo.lock",
            "workspace-lock paths differ")
    check_hash(value["sha256"], "workspace-lock.sha256")
    require(value["sha256"] == WORKSPACE_LOCK_SHA256,
            "workspace-lock digest differs")
    require(value["baseline_cap_not_started"] is True,
            "workspace-lock baseline marker differs")
    require(isinstance(value["reason"], str) and value["reason"].strip(),
            "workspace-lock reason is empty")
    need(REPO / value["path"], "workspace Cargo.lock")
    need(HERE / value["retained"], "workspace-Cargo.lock")
    require(sha(REPO / value["path"]) == value["sha256"],
            "workspace Cargo.lock changed")
    require(sha(HERE / value["retained"]) == value["sha256"],
            "retained workspace Cargo.lock changed")
    completion = read_json(LOCK_COMPLETION, "baseline-lock-completion.json")
    require(isinstance(completion, dict) and set(completion) == {
        "checked_utc", "status", "workspace_lock_sha256",
        "baseline_capture_exit_code", "baseline_session", "scope",
    }, "baseline-lock-completion envelope differs")
    completed_utc = parse_time(completion["checked_utc"],
                               "baseline-lock-completion.checked_utc")
    require(completion["status"] == "pass"
            and completion["workspace_lock_sha256"] == value["sha256"]
            and completion["baseline_capture_exit_code"] == 0
            and isinstance(completion["baseline_session"], int)
            and completion["baseline_session"] > 0
            and completion["scope"] == (
                "Separately frozen before workspace build; unchanged at complete baseline pipeline. "
                "No claim of per-child workspace-lock checks in this original pipeline."
            ), "baseline lock completion differs")
    require(frozen_utc <= completed_utc,
            "workspace lock was frozen after baseline lock completion")
    cap_receipt = BASELINE / "build-cap.receipt.json"
    if cap_receipt.exists():
        cap_value = read_json(cap_receipt, "baseline/build-cap.receipt.json")
        cap_start, _ = interval(cap_value, "baseline/build-cap.receipt.json")
        require(frozen_utc <= cap_start,
                "workspace lock was not frozen before the baseline cap build")
    baseline_receipts = [path for path in BASELINE.glob("*.receipt.json")
                         if "-r2-" not in path.name]
    for receipt_path in baseline_receipts:
        _, receipt_end = interval(
            read_json(receipt_path, rel(receipt_path)), rel(receipt_path)
        )
        require(receipt_end <= completed_utc,
                "baseline lock completion predates a retained baseline receipt")
    return {"status": "pass", "sha256": sha(LOCK_BINDING),
            "workspace_lock_sha256": value["sha256"],
            "supplemental_sha256": supplemental["sha256"],
            "completion_sha256": sha(LOCK_COMPLETION)}


def validate_public_test_bundle() -> dict[str, Any]:
    need(PUBLIC, "public-test-sources", directory=True)
    hashes = read_json(PUBLIC / "source-hashes.json", "public-test-sources/source-hashes.json")
    require(isinstance(hashes, dict) and set(hashes) == {
        "crates/litchi-xlsx/tests/source_backed_cell_values/compact_source_proof.rs",
        "crates/litchi-xlsx/tests/source_backed_cell_values.rs",
    }, "public test source inventory differs")
    plan = validate_plan()
    baseline = git_tree_manifest(plan["revision"])
    for name, row in hashes.items():
        require(isinstance(row, dict) and set(row) == {"baseline_sha256", "test_sha256"},
                f"public test hash row differs: {name}")
        base_digest = row["baseline_sha256"]
        if base_digest is None:
            require(name.endswith("compact_source_proof.rs")
                    and name not in baseline, f"new public test baseline row differs: {name}")
        else:
            check_hash(base_digest, f"public test baseline hash {name}")
            require(baseline.get(name) == base_digest,
                    f"public test baseline hash is not frozen source: {name}")
        check_hash(row["test_sha256"], f"public test hash {name}")
    patch = need(PUBLIC / "public-tests.patch", "public-test-sources/public-tests.patch")
    require(patch.stat().st_size > 0, "public test patch is empty")
    need(PUBLIC / "README.md", "public-test-sources/README.md")
    with tempfile.TemporaryDirectory(prefix=".litchi-0552-public-", dir="/home/zhuhe") as folder:
        index = Path(folder) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", plan["revision"]], env=env)
        git(["git", "apply", "--cached", "--binary", str(patch)], env=env)
        changed = set(git(["git", "diff", "--cached", "--name-only", plan["revision"]],
                          env=env).decode().splitlines())
        require(changed == set(hashes),
                "public test patch tracked path inventory differs")
        indexed = parse_index(index)
        oids = {indexed[name] for name in hashes}
        blobs = index_hashes(index, oids)
        for name, row in hashes.items():
            require(blobs[indexed[name]] == row["test_sha256"],
                    f"public test patch hash differs: {name}")
    return {"status": "pass", "sha256": sha(PUBLIC / "source-hashes.json"),
            "files": sorted(hashes), "patch_sha256": sha(patch)}


def regular_files(root: Path, label: str) -> dict[str, Path]:
    """Return a symlink-free recursive file inventory under an evidence root."""
    need(root, label, directory=True)
    result: dict[str, Path] = {}
    for path in root.rglob("*"):
        require(not path.is_symlink(), f"{label} contains a symlink: {rel(path)}")
        if path.is_file():
            name = path.relative_to(root).as_posix()
            require(name not in result, f"{label} repeats {name}")
            result[name] = path
        else:
            require(path.is_dir(), f"{label} contains a non-directory entry")
    return result


def public_exact_manifest(child_sha256: str,
                          *, include_lock: bool = False) -> dict[str, str]:
    """Build the source manifest for one exact-public oracle revision."""
    check_hash(child_sha256, "public exact child hash")
    baseline = git_tree_manifest(validate_plan()["revision"])
    public_hashes = read_json(PUBLIC / "source-hashes.json",
                              "public-test-sources/source-hashes.json")
    exact_inputs = read_json(PUBLIC_EXACT / "inputs.json",
                             "public-exact-test-sources/inputs.json")
    require(isinstance(public_hashes, dict), "public test hashes are malformed")
    require(isinstance(exact_inputs, dict)
            and isinstance(exact_inputs.get("source_hashes"), dict),
            "public exact source hashes are malformed")
    result = dict(baseline)
    for name, row in public_hashes.items():
        require(isinstance(row, dict) and is_hash(row.get("test_sha256")),
                f"public test hash row is malformed: {name}")
        result[name] = row["test_sha256"]
    for name, digest in exact_inputs["source_hashes"].items():
        check_hash(digest, f"public exact source hash {name}")
        result[name] = child_sha256 if name.endswith("public_exact_output.rs") else digest
    if include_lock:
        result["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    return dict(sorted(result.items()))


def test_result_summaries(path: Path) -> list[dict[str, int | str]]:
    """Parse cargo's compact test-result rows without accepting omissions."""
    pattern = re.compile(
        r"test result: (ok|FAILED)\.\s+(\d+) passed;\s+(\d+) failed;\s+"
        r"(\d+) ignored;\s+(\d+) measured;\s+(\d+) filtered out"
    )
    rows: list[dict[str, int | str]] = []
    for match in pattern.finditer(read_text(path, rel(path))):
        rows.append({
            "status": match.group(1), "passed": int(match.group(2)),
            "failed": int(match.group(3)), "ignored": int(match.group(4)),
            "measured": int(match.group(5)), "filtered": int(match.group(6)),
        })
    return rows


def validate_public_exact_tests() -> dict[str, Any]:
    """Validate the retained exact-output oracle and every correction attempt."""
    root = need(PUBLIC_EXACT, "public-exact-test-sources", directory=True)
    inputs = read_json(root / "inputs.json", "public-exact-test-sources/inputs.json")
    require(isinstance(inputs, dict) and set(inputs) == {
        "frozen_utc", "scope", "source_hashes", "parent_before_sha256",
        "patch_sha256",
    }, "public exact inputs envelope differs")
    frozen = parse_time(inputs["frozen_utc"],
                        "public-exact-test-sources.frozen_utc")
    require(inputs["scope"] == (
        "Explicit public whole-worksheet oracle; validate restored baseline first, then candidate."
    ), "public exact scope differs")
    source_hashes = inputs["source_hashes"]
    exact_names = {
        "crates/litchi-xlsx/tests/source_backed_cell_values.rs",
        "crates/litchi-xlsx/tests/source_backed_cell_values/public_exact_output.rs",
    }
    require(isinstance(source_hashes, dict) and set(source_hashes) == exact_names,
            "public exact source inventory differs")
    for name, digest in source_hashes.items():
        check_hash(digest, f"public exact source hash {name}")
    public_hashes = read_json(PUBLIC / "source-hashes.json",
                              "public-test-sources/source-hashes.json")
    require(inputs["parent_before_sha256"] ==
            public_hashes["crates/litchi-xlsx/tests/source_backed_cell_values.rs"][
                "test_sha256"],
            "public exact parent-before hash differs")
    patch = need(root / "public-tests.patch",
                 "public-exact-test-sources/public-tests.patch")
    require(patch.stat().st_size > 0
            and sha(patch) == check_hash(inputs["patch_sha256"],
                                         "public exact patch hash"),
            "public exact patch hash differs")
    witness = need(root / "public_exact_output.rs",
                   "public-exact-test-sources/public_exact_output.rs")
    child_name = "crates/litchi-xlsx/tests/source_backed_cell_values/public_exact_output.rs"
    require(sha(witness) == source_hashes[child_name],
            "public exact top-level source witness differs")
    files = regular_files(root, "public-exact-test-sources")
    expected_files = {
        "inputs.json", "public-tests.patch", "public_exact_output.rs",
        *(f"sources/{name}" for name in exact_names),
    }
    attempt_names = [f"attempt-{number:02d}" for number in range(2, 6)]
    for attempt_name in attempt_names:
        expected_files |= {
            f"{attempt_name}/inputs.json", f"{attempt_name}/public-tests.patch",
            f"{attempt_name}/sources/{child_name}",
        }
    require(set(files) == expected_files,
            "public exact recursive file inventory differs")
    for name, digest in source_hashes.items():
        path = root / "sources" / name
        require(sha(path) == digest, f"public exact source snapshot differs: {name}")

    plan = validate_plan()
    with tempfile.TemporaryDirectory(prefix=".litchi-0552-exact-", dir="/home/zhuhe") as folder:
        index = Path(folder) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", plan["revision"]], env=env)
        git(["git", "apply", "--cached", "--binary", str(PUBLIC / "public-tests.patch")],
            env=env)
        git(["git", "apply", "--cached", "--binary", str(patch)], env=env)

        def index_manifest(names: set[str]) -> dict[str, str]:
            indexed = parse_index(index)
            require(names <= set(indexed), "public exact replay omits a source path")
            oids = {indexed[name] for name in names}
            blobs = index_hashes(index, oids)
            return dict(sorted({name: blobs[indexed[name]] for name in names}.items()))

        root_expected = public_exact_manifest(source_hashes[child_name])
        root_names = set(root_expected)
        require(index_manifest(root_names) == root_expected,
                "public exact root patch replay differs")

        attempt_specs = {
            "attempt-02": {
                "receipt": "baseline-public-exact-01", "receipt_field": "failed_check_receipt_sha256",
                "exit_code": 101,
                "command": [
                    "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
                    "--all-features", "--test", "source_backed_cell_values",
                    "public_multi_edit_matches_baseline_whole_worksheet_output", "--",
                    "--test-threads=2",
                ],
            },
            "attempt-03": {
                "receipt": "baseline-public-exact-02", "receipt_field": "failed_check_receipt_sha256",
                "exit_code": 101,
                "command": [
                    "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
                    "--all-features", "--test", "source_backed_cell_values",
                    "public_multi_edit_matches_baseline_whole_worksheet_output", "--",
                    "--test-threads=2",
                ],
            },
            "attempt-04": {
                "receipt": "baseline-public-exact-03", "receipt_field": "prior_check_receipt_sha256",
                "exit_code": 0,
                "command": [
                    "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
                    "--all-features", "--test", "source_backed_cell_values",
                    "public_multi_edit_matches_baseline_whole_worksheet_output", "--",
                    "--test-threads=2",
                ],
            },
            "attempt-05": {
                "receipt": "baseline-public-exact-04", "receipt_field": "prior_check_receipt_sha256",
                "exit_code": 101,
                "command": [
                    "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
                    "--all-features", "--test", "source_backed_cell_values",
                    "public_exact_output", "--", "--test-threads=2",
                ],
            },
        }
        child = source_hashes[child_name]
        rows: list[dict[str, Any]] = []
        for attempt_name in attempt_names:
            attempt = root / attempt_name
            value = read_json(attempt / "inputs.json", rel(attempt / "inputs.json"))
            expected_keys = {
                "frozen_utc", "scope", "path", "before_sha256", "after_sha256",
                "patch_sha256", attempt_specs[attempt_name]["receipt_field"],
            }
            require(isinstance(value, dict) and set(value) == expected_keys,
                    f"{attempt_name} inputs envelope differs")
            attempted = parse_time(value["frozen_utc"], f"{attempt_name}.frozen_utc")
            require(attempted >= frozen, f"{attempt_name} predates the exact root")
            require(value["path"] == child_name,
                    f"{attempt_name} source path differs")
            check_hash(value["before_sha256"], f"{attempt_name}.before_sha256")
            check_hash(value["after_sha256"], f"{attempt_name}.after_sha256")
            check_hash(value["patch_sha256"], f"{attempt_name}.patch_sha256")
            check_hash(value[attempt_specs[attempt_name]["receipt_field"]],
                       f"{attempt_name} receipt hash")
            require(value["before_sha256"] == child,
                    f"{attempt_name} before hash does not continue the oracle")
            source = attempt / "sources" / child_name
            require(sha(source) == value["after_sha256"],
                    f"{attempt_name} source witness differs")
            attempt_patch = need(attempt / "public-tests.patch",
                                 f"{attempt_name}/public-tests.patch")
            require(sha(attempt_patch) == value["patch_sha256"],
                    f"{attempt_name} patch hash differs")
            receipt_path = HERE / "check-attempts" / attempt_specs[attempt_name]["receipt"]
            expected_manifest = public_exact_manifest(child, include_lock=True)
            receipt = validate_check_attempt(receipt_path, expected_manifest)
            require(receipt["receipt_sha256"] ==
                    value[attempt_specs[attempt_name]["receipt_field"]]
                    and receipt["exit_code"] == attempt_specs[attempt_name]["exit_code"]
                    and receipt["command"] == attempt_specs[attempt_name]["command"],
                    f"{attempt_name} failed-check custody differs")
            require(parse_time(receipt["end_utc"], f"{attempt_name} receipt end") <= attempted,
                    f"{attempt_name} receipt postdates its frozen metadata")
            if attempt_name == "attempt-04":
                summaries = test_result_summaries(receipt_path / "stdout")
                require(len(summaries) == 1 and summaries[0]["status"] == "ok"
                        and summaries[0]["passed"] == 1
                        and summaries[0]["failed"] == 0,
                        "baseline-public-exact-03 test summary differs")
            child = value["after_sha256"]
            git(["git", "apply", "--cached", "--binary", str(attempt_patch)], env=env)
            expected = public_exact_manifest(child)
            require(index_manifest(set(expected)) == expected,
                    f"{attempt_name} patch replay differs")
            rows.append({"attempt": attempt_name, "receipt": {
                         key: value for key, value in receipt.items() if key != "times"
                     },
                         "before_sha256": value["before_sha256"],
                         "after_sha256": child})

        final_path = HERE / "check-attempts" / "baseline-public-exact-05"
        final_command = [
            "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
            "--all-features", "--test", "source_backed_cell_values",
            "public_exact_output", "--", "--test-threads=2",
        ]
        final_receipt = validate_check_attempt(
            final_path, public_exact_manifest(child, include_lock=True)
        )
        require(final_receipt["exit_code"] == 0
                and final_receipt["command"] == final_command,
                "baseline-public-exact-05 custody differs")
        final_summaries = test_result_summaries(final_path / "stdout")
        require(len(final_summaries) == 1 and final_summaries[0]["status"] == "ok"
                and final_summaries[0]["passed"] == 2
                and final_summaries[0]["failed"] == 0
                and final_summaries[0]["ignored"] == 0,
                "baseline-public-exact-05 test summary differs")
    return {
        "status": "pass", "inputs_sha256": sha(root / "inputs.json"),
        "patch_sha256": sha(patch), "source_hashes": dict(sorted(source_hashes.items())),
        "latest_child_sha256": child, "attempts": rows,
        "final_receipt": {key: value for key, value in final_receipt.items()
                          if key != "times"},
    }


def validate_public_tests() -> dict[str, Any]:
    """Validate both the original public guards and the exact-output oracle."""
    return {"status": "pass", "bundle": validate_public_test_bundle(),
            "exact": validate_public_exact_tests()}


def stage_manifest(stage: str) -> tuple[dict[str, str], str]:
    require(stage in STAGES, f"unknown stage {stage}")
    path = HERE / stage / "source-manifest.json"
    manifest = source_manifest(path, f"{stage}/source-manifest.json")
    require("tools/perf-baseline/src/xlsx_planning_guard.rs" in manifest,
            f"{stage} source manifest omits planning guard source")
    require("crates/litchi-xlsx/examples/perf_cap_boundary.rs" in manifest,
            f"{stage} source manifest omits cap example")
    return manifest, sha(path)


def validate_source_patch(stage: str, manifest: dict[str, str]) -> dict[str, Any]:
    patch = need(HERE / stage / "source.patch", f"{stage}/source.patch")
    revision = validate_plan()["revision"]
    tree = git_tree_manifest(revision)
    if stage == "baseline":
        require(patch.read_bytes() == b"", "baseline source.patch is not empty")
        require(manifest == tree, "baseline source manifest differs from frozen Git tree")
        return {"status": "pass", "patch_sha256": sha(patch), "changed_tracked": []}
    require(patch.stat().st_size > 0, "candidate source.patch is empty")
    with tempfile.TemporaryDirectory(prefix=".litchi-0552-source-", dir="/home/zhuhe") as folder:
        index = Path(folder) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", revision], env=env)
        git(["git", "apply", "--cached", "--binary", str(patch)], env=env)
        indexed = parse_index(index)
        tracked_names = set(tree) & set(manifest)
        require(tracked_names <= set(indexed),
                f"{stage} private replay omits tracked source")
        blobs = index_hashes(index, {indexed[name] for name in tracked_names})
        for name in sorted(tracked_names):
            require(blobs[indexed[name]] == manifest[name],
                    f"{stage} replay hash differs: {name}")
        changed = set(git(["git", "diff", "--cached", "--name-only", revision],
                          env=env).decode().splitlines())
    require(changed <= set(manifest),
            f"{stage} source.patch changes paths absent from source manifest")
    require(changed, f"{stage} source.patch has no tracked changes")
    for name, digest in manifest.items():
        if name in tree:
            continue
        candidates = (
            REPO / name,
            HERE / "candidate-sources" / name,
            HERE / "candidate-attempts" / "draft-07" / "sources" / name,
            HERE / stage / "candidate-sources" / name,
        )
        matches = [path for path in candidates if path.is_file() and not path.is_symlink()]
        require(matches and any(sha(path) == digest for path in matches),
                f"{stage} untracked source has no retained byte witness: {name}")
    require(set(manifest) >= set(tree), f"{stage} source manifest dropped frozen paths")
    return {"status": "pass", "patch_sha256": sha(patch),
            "changed_tracked": sorted(changed)}


def private_source_replay(stage: str) -> dict[str, Any]:
    manifest, manifest_sha = stage_manifest(stage)
    patch_result = validate_source_patch(stage, manifest)
    return {"status": "pass", "manifest_sha256": manifest_sha,
            "entries": len(manifest), "patch": patch_result}


def validate_source(stage: str, *, require_stage: bool = True) -> dict[str, Any]:
    if not (HERE / stage).is_dir():
        if require_stage:
            raise IncompleteError(f"{stage} stage directory is missing")
        return {"status": "not-present"}
    return private_source_replay(stage)


def expected_main_jobs(lane: str, stage: str) -> list[dict[str, Any]]:
    require(lane in ("preflight", "native", "alloc"), f"unknown main lane {lane}")
    require(stage in STAGES, f"unknown stage {stage}")
    if lane == "preflight":
        repeats = (1,)
        samples, warmup = 1, 0
    elif lane == "native":
        repeats = (1, 2)
        samples, warmup = 200, 20
    else:
        repeats = (1, 2)
        samples, warmup = 20, 3
    jobs: list[dict[str, Any]] = []
    for repeat in repeats:
        shapes = SHAPES if repeat == 1 else tuple(reversed(SHAPES))
        for shape in shapes:
            for index, case in enumerate(CASES):
                jobs.append({
                    "name": f"{lane}-r{repeat}-{shape}-c{index}",
                    "lane": lane, "repeat": repeat, "shape": shape, "case": case,
                    "samples": samples, "warmup": warmup,
                })
    return jobs


def expected_guard_jobs(lane: str) -> list[dict[str, Any]]:
    require(lane in ("normal", "alloc"), f"unknown guard lane {lane}")
    samples, warmup = (200, 20) if lane == "normal" else (20, 3)
    name_lane = "native" if lane == "normal" else "alloc"
    jobs: list[dict[str, Any]] = []
    for repeat in (1, 2):
        shapes = GUARD_SHAPES if repeat == 1 else tuple(reversed(GUARD_SHAPES))
        for shape in shapes:
            for case in GUARD_CASES:
                jobs.append({
                    "name": f"guard-{name_lane}-r{repeat}-{shape}-{case}",
                    "lane": lane, "repeat": repeat, "shape": shape, "case": case,
                    "samples": samples, "warmup": warmup,
                })
    return jobs


def expected_cap_jobs() -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    for repeat in (1, 2):
        sizes = CAP_SIZES if repeat == 1 else tuple(reversed(CAP_SIZES))
        for size in sizes:
            jobs.append({
                "name": f"cap-r{repeat}-{size}", "repeat": repeat, "size": size,
                "samples": 200, "warmup": 20,
            })
    return jobs


def expected_profile_jobs() -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    for repeat in (1, 2):
        shapes = SHAPES if repeat == 1 else tuple(reversed(SHAPES))
        for shape in shapes:
            jobs.append({
                "name": f"profile-r{repeat}-{shape}-c0",
                "repeat": repeat, "shape": shape, "case": PROFILE_CASE,
                "samples": 1, "warmup": 0,
            })
    return jobs


def validate_host_sidecar(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict) and set(value) == {
        "observed_utc", "compiler_processes", "scope",
    }, f"{label} host sidecar inventory differs")
    parse_time(value["observed_utc"], f"{label}.observed_utc")
    require(value["scope"] == HOST_SCOPE, f"{label} host scope differs")
    require(isinstance(value["compiler_processes"], list),
            f"{label} compiler_processes is not a list")
    for index, process in enumerate(value["compiler_processes"]):
        require(isinstance(process, dict) and set(process) == {"pid", "comm", "cwd"},
                f"{label} compiler process {index} differs")
        require(isinstance(process["pid"], int) and process["pid"] > 0,
                f"{label} compiler process pid differs")
        require(process["comm"] in ("cargo", "rustc"),
                f"{label} compiler process comm differs")
        require(isinstance(process["cwd"], str) and process["cwd"],
                f"{label} compiler process cwd is empty")


def validate_receipt_common(value: Any, label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS,
            f"{label} receipt inventory differs")
    start, end = interval(value, label)
    require(isinstance(value["command"], list)
            and all(isinstance(item, str) and item for item in value["command"]),
            f"{label}.command is malformed")
    require(isinstance(value["exit_code"], int) and not isinstance(value["exit_code"], bool),
            f"{label}.exit_code is malformed")
    env = value["environment"]
    require(isinstance(env, dict) and set(env) == RECEIPT_ENVIRONMENT
            and all(item is None for item in env.values()),
            f"{label}.environment is not controlled")
    for field in ("execution_manifest_sha256", "source_manifest_sha256",
                  "script_sha256", "plan_sha256"):
        check_hash(value[field], f"{label}.{field}")
    artifacts = value["artifacts"]
    require(isinstance(artifacts, dict), f"{label}.artifacts is not an object")
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).name == name,
                f"{label}.artifacts has unsafe name")
        check_hash(digest, f"{label}.artifacts.{name}")
    return start, end


def validate_artifacts(folder: Path, name: str, receipt: dict[str, Any],
                       expected: set[str]) -> None:
    artifacts = receipt["artifacts"]
    require(set(artifacts) == expected, f"{folder.name}/{name} artifact inventory differs")
    require(folder.is_dir() and not folder.is_symlink(),
            f"{folder.name} artifact folder is not a directory")
    actual = {
        path.name for path in folder.iterdir()
        if path.is_file() and not path.is_symlink()
        and (path.name == name + ".receipt.json" or path.name.startswith(name + "."))
    }
    require(actual == expected | {name + ".receipt.json"},
            f"{folder.name}/{name} raw artifact inventory differs")
    for filename, digest in artifacts.items():
        path = folder / filename
        need(path, f"{folder.name}/{filename}")
        require(sha(path) == digest, f"{folder.name}/{filename} digest differs")
        if filename.endswith(".host.json"):
            validate_host_sidecar(path, f"{folder.name}/{filename}")


def expected_build_command(kind: str) -> list[str]:
    prefix = [
        "env", f"TMPDIR={TARGET / 'tmp'}", "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0", "cargo", "build", "--release", "--locked",
    ]
    if kind in ("normal", "alloc"):
        command = prefix + [
            "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
            "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else ""),
            "--target-dir", str(TARGET),
        ]
        if kind == "alloc":
            command += ["--features", "allocator-metrics"]
        return command
    if kind in ("guard-normal", "guard-alloc"):
        command = prefix + [
            "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
            "xlsx_planning_guard", "--target-dir", str(TARGET),
        ]
        if kind == "guard-alloc":
            command += ["--features", "allocator-metrics"]
        return command
    require(kind == "cap", f"unknown build kind {kind}")
    return prefix + ["-p", "litchi-xlsx", "--example", "perf_cap_boundary",
                      "--target-dir", str(TARGET)]


def descriptor_name(kind: str) -> str:
    return {
        "normal": "binary-normal.json", "alloc": "binary-alloc.json",
        "guard-normal": "binary-guard-normal.json",
        "guard-alloc": "binary-guard-alloc.json", "cap": "binary-cap.json",
    }[kind]


def retained_binary_name(kind: str) -> str:
    return {
        "normal": "normal", "alloc": "alloc",
        "guard-normal": "guard-normal", "guard-alloc": "guard-alloc",
        "cap": "cap",
    }[kind]


def cleanup_binary_hash(stage: str, kind: str, path: Path) -> str | None:
    for candidate in (HERE / "cleanup.json", HERE / stage / "cleanup.json"):
        if not candidate.is_file() or candidate.is_symlink():
            continue
        value = read_json(candidate, rel(candidate))
        if not isinstance(value, dict) or value.get("owned_paths_absent") is not True:
            continue
        if value.get("accessible_process_references") != []:
            continue
        hashes = value.get("binary_sha256_by_kind")
        if not isinstance(hashes, dict):
            continue
        for key in (kind, path.name, str(path), f"{stage}/{kind}"):
            if is_hash(hashes.get(key)):
                return hashes[key]
    return None


def validate_binary(stage: str, kind: str, manifest_sha: str) -> dict[str, Any]:
    folder = HERE / stage
    descriptor_path = folder / descriptor_name(kind)
    descriptor = read_json(descriptor_path, rel(descriptor_path))
    require(isinstance(descriptor, dict) and set(descriptor) == {
        "path", "sha256", "bytes", "build_receipt_sha256",
        "source_manifest_sha256",
    }, f"{rel(descriptor_path)} inventory differs")
    expected_path = SCRATCH_ROOT / stage / retained_binary_name(kind)
    path = Path(descriptor["path"])
    require(path == expected_path, f"{rel(descriptor_path)} path differs")
    digest = check_hash(descriptor["sha256"], f"{rel(descriptor_path)}.sha256")
    require(isinstance(descriptor["bytes"], int) and descriptor["bytes"] > 0,
            f"{rel(descriptor_path)} byte count differs")
    require(descriptor["source_manifest_sha256"] == manifest_sha,
            f"{rel(descriptor_path)} source manifest differs")
    build_path = folder / f"build-{kind}.receipt.json"
    require(descriptor["build_receipt_sha256"] == sha(build_path),
            f"{rel(descriptor_path)} build receipt hash differs")
    if path.exists():
        require(path.is_file() and not path.is_symlink(), f"{rel(path)} is not a binary")
        require(sha(path) == digest and path.stat().st_size == descriptor["bytes"],
                f"{rel(descriptor_path)} binary custody differs")
    else:
        require(cleanup_binary_hash(stage, kind, path) == digest,
                f"{rel(descriptor_path)} binary was removed without cleanup custody")
    return {
        "kind": kind, "path": str(path), "sha256": digest,
        "bytes": descriptor["bytes"], "stage": stage,
        "descriptor_sha256": sha(descriptor_path),
    }


def validate_build(stage: str, kind: str, manifest_sha: str) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    folder = HERE / stage
    path = folder / f"build-{kind}.receipt.json"
    value = read_json(path, rel(path))
    start, end = validate_receipt_common(value, rel(path))
    require(value["exit_code"] == 0 and value["binary_sha256"] is None,
            f"{rel(path)} is not a successful build")
    require(value["execution_stage"] == stage
            and value["execution_manifest_sha256"] == manifest_sha
            and value["source_manifest_sha256"] == manifest_sha,
            f"{rel(path)} stage/source binding differs")
    require(value["script_sha256"] == sha(RUN)
            and value["plan_sha256"] == sha(PLAN)
            and value["command"] == expected_build_command(kind),
            f"{rel(path)} command or driver binding differs")
    stem = path.name[:-len(".receipt.json")]
    expected = {stem + suffix for suffix in (".host.json", ".stdout", ".stderr")}
    validate_artifacts(folder, stem, value, expected)
    return {"name": stem, "receipt_sha256": sha(path), "start_utc": value["start_utc"],
            "end_utc": value["end_utc"], "seconds": value["seconds"]}, (start, end)


def expected_main_command(stage: str, job: dict[str, Any], binary: dict[str, Any]) -> list[str]:
    folder = HERE / stage
    report = folder / f"{job['name']}.json"
    catalog = folder / f"{job['name']}.catalog.json"
    return [
        "taskset", "-c", str(CPU), "/usr/bin/time", "-v", binary["path"],
        "--case", job["case"], "--xlsx-cell-crud-shape", job["shape"],
        "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
        "--json", str(report), "--corpus-manifest", str(catalog),
    ]


def expected_execution(stage: str, repeat: int,
                       baseline_manifest_sha: str,
                       candidate_manifest_sha: str) -> tuple[str, str]:
    if stage == "candidate" or repeat == 2:
        return "candidate", candidate_manifest_sha
    return "baseline", baseline_manifest_sha


def validate_time_stderr(path: Path, label: str) -> dict[str, int]:
    text = read_text(path, label)
    match = re.search(r"Maximum resident set size \(kbytes\):\s+([0-9]+)", text)
    require(match is not None, f"{label} lacks /usr/bin/time maximum RSS")
    rss = int(match.group(1))
    require(rss > 0, f"{label} maximum RSS is zero")
    return {"maximum_resident_set_size_kib": rss}


def corpus_binding(report: dict[str, Any], catalog_path: Path, label: str) -> dict[str, Any]:
    catalog = read_json(catalog_path, label + ".catalog")
    require(isinstance(catalog, dict), f"{label}.catalog is not an object")
    helper_path = REPO / "tools/validate_perf_corpus_binding.py"
    need(helper_path, "tools/validate_perf_corpus_binding.py")
    spec = importlib.util.spec_from_file_location("corpus_binding_0552", helper_path)
    require(spec is not None and spec.loader is not None,
            "cannot load corpus binding helper")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    try:
        module.validate_binding(report, catalog)
    except Exception as error:
        raise VerificationError(f"{label} corpus binding failed: {error}") from error
    reference = report.get("corpus_catalog")
    require(isinstance(reference, dict), f"{label}.corpus_catalog is missing")
    for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256"):
        require(key in reference, f"{label}.corpus_catalog.{key} is missing")
    require(reference["manifest_version"] == catalog.get("manifest_version") == 2
            and reference["catalog_id"] == catalog.get("catalog_id") == "litchi-perf-corpus-v2",
            f"{label} corpus catalog identity differs")
    check_hash(catalog.get("catalog_sha256"), f"{label}.catalog_sha256")
    check_hash(catalog.get("content_set_sha256"), f"{label}.content_set_sha256")
    require(reference["catalog_sha256"] == catalog["catalog_sha256"]
            and reference["content_set_sha256"] == catalog["content_set_sha256"],
            f"{label} corpus catalog reference differs")
    return {
        "manifest_version": catalog["manifest_version"],
        "catalog_id": catalog["catalog_id"],
        "catalog_sha256": catalog["catalog_sha256"],
        "content_set_sha256": catalog["content_set_sha256"],
        "sha256": sha(catalog_path),
    }


def main_report_identity(report: dict[str, Any], job: dict[str, Any],
                         binary: dict[str, Any], plan: dict[str, Any],
                         catalog_path: Path, label: str) -> dict[str, Any]:
    require(isinstance(report, dict) and report.get("schema_version") == 1,
            f"{label} report schema differs")
    tool = report.get("tool")
    expected_binary_name = (
        "litchi-perf-baseline-alloc" if job["lane"] == "alloc"
        else "litchi-perf-baseline"
    )
    expected_instrumentation = (
        "system_allocator_operation_scoped" if job["lane"] == "alloc" else "none"
    )
    require(isinstance(tool, dict)
            and tool.get("name") == "litchi-perf-baseline"
            and tool.get("binary") == expected_binary_name
            and tool.get("profile") == "release"
            and tool.get("instrumentation") == expected_instrumentation,
            f"{label} tool identity differs")
    if job["lane"] == "alloc":
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{label} allocator counter revision differs")
    binary_identity = report.get("binary_identity")
    require(isinstance(binary_identity, dict)
            and binary_identity.get("path") == binary["path"]
            and binary_identity.get("binary_sha256") == binary["sha256"]
            and binary_identity.get("binary_bytes") == binary["bytes"]
            and binary_identity.get("profile") == "release",
            f"{label} binary identity differs")
    environment = report.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == plan["revision"]
            and environment.get("cpu_affinity") == str(CPU),
            f"{label} benchmark environment differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("cases") == [job["case"]]
            and configuration.get("xlsx_cell_crud_shapes") == [job["shape"]]
            and configuration.get("samples_per_case") == job["samples"]
            and configuration.get("warmup_iterations_per_case") == job["warmup"],
            f"{label} benchmark configuration differs")
    parallel = report.get("parallel_metrics")
    require(isinstance(parallel, dict)
            and parallel.get("schema_version") == 1
            and parallel.get("scope") == "explicit_local_execution_only"
            and parallel.get("claim") == "descriptive",
            f"{label} parallel metrics envelope differs")
    workers = parallel.get("configured_worker_budget")
    require(isinstance(workers, dict)
            and workers.get("status") == "measured"
            and workers.get("value") == [1],
            f"{label} worker budget differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1
            and isinstance(results[0], dict)
            and results[0].get("case") == job["case"],
            f"{label} result matrix differs")
    result = results[0]
    corpus = result.get("corpus")
    require(isinstance(corpus, dict)
            and corpus.get("generator") ==
            "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1"
            and corpus.get("shape") == job["shape"]
            and corpus.get("package_format") == "XLSX/OPC/ZIP",
            f"{label} corpus identity differs")
    for field in ("archive_sha256", "target_payload_sha256"):
        check_hash(corpus.get(field), f"{label}.corpus.{field}")
    for field in ("archive_bytes", "archive_member_count", "entry_count",
                  "entry_bytes", "uncompressed_payload_bytes", "target_payload_bytes"):
        require(isinstance(corpus.get(field), int) and corpus[field] >= 0,
                f"{label}.corpus.{field} differs")
    corpus_xlsx = corpus.get("xlsx")
    require(isinstance(corpus_xlsx, dict)
            and isinstance(corpus_xlsx.get("source_members"), dict)
            and len(corpus_xlsx["source_members"].get("worksheets", [])) ==
            corpus_xlsx.get("sheet_count"),
            f"{label}.corpus.xlsx identity differs")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns",
            f"{label}.elapsed_ns is missing")
    values = elapsed.get("samples")
    require(isinstance(values, list) and len(values) == job["samples"]
            and all(isinstance(value, int) and not isinstance(value, bool) and value >= 0
                    for value in values)
            and values == sorted(values),
            f"{label}.elapsed_ns sample vector differs")
    order = elapsed.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(job["samples"])),
            f"{label}.elapsed_ns sample order differs")
    for field in ("min", "p50", "p95", "p99", "max", "mean",
                  "standard_deviation"):
        require(field in elapsed and isinstance(elapsed[field], (int, float))
                and not isinstance(elapsed[field], bool)
                and math.isfinite(float(elapsed[field])),
                f"{label}.elapsed_ns.{field} differs")
    source = result.get("source")
    require(isinstance(source, dict)
            and isinstance(source.get("xlsx_cell_values"), dict),
            f"{label}.source cell-values metrics are missing")
    xlsx = source["xlsx_cell_values"]
    managed = "managed" in job["case"]
    expected_implementation = "managed-source-backed" if managed else "source-backed"
    expected_cache_mode = "managed-budget" if managed else "unmanaged-control"
    expected_sheet_count = 4 if "one_percent" in job["case"] else 1
    require(xlsx.get("implementation") == expected_implementation
            and xlsx.get("cache_mode") == expected_cache_mode
            and isinstance(xlsx.get("update_count"), int)
            and xlsx["update_count"] > 0
            and xlsx.get("selected_worksheet_count") == expected_sheet_count,
            f"{label}.source cell-values identity differs")
    output = check_hash(result.get("output_sha256"), f"{label}.output_sha256")
    check_hash(xlsx.get("output_sha256", [None])[0]
               if isinstance(xlsx.get("output_sha256"), list)
               and xlsx["output_sha256"] else None, f"{label}.source output")
    source_output = xlsx["output_sha256"][0]
    require(source_output == output, f"{label} source/output identity differs")
    semantic = xlsx.get("semantic_sha256")
    untouched_count = xlsx.get("untouched_member_count")
    untouched_hashes = xlsx.get("untouched_member_sha256")
    require(isinstance(semantic, list) and semantic and all(is_hash(item) for item in semantic)
            and isinstance(untouched_count, int) and untouched_count >= 0
            and isinstance(untouched_hashes, list) and untouched_hashes
            and all(is_hash(item) for item in untouched_hashes),
            f"{label} source provenance identity is incomplete")
    catalog_identity = corpus_binding(report, catalog_path, label)
    return {
        "case": job["case"], "shape": job["shape"],
        "corpus": corpus, "corpus_catalog": catalog_identity,
        "output_sha256": output, "semantic_sha256": semantic[0],
        "untouched_member_count": untouched_count,
        "untouched_member_sha256": untouched_hashes[0],
        "source_output_sha256": source_output,
    }


def validate_main_job(stage: str, job: dict[str, Any], binary: dict[str, Any],
                      manifest_sha: str, baseline_manifest_sha: str,
                      candidate_manifest_sha: str, plan: dict[str, Any]) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    folder = HERE / stage
    receipt_path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(receipt_path, rel(receipt_path))
    start, end = validate_receipt_common(receipt, rel(receipt_path))
    expected_stage, expected_execution_manifest = expected_execution(
        stage, job["repeat"], baseline_manifest_sha, candidate_manifest_sha
    )
    require(receipt["exit_code"] == 0
            and receipt["binary_sha256"] == binary["sha256"]
            and receipt["execution_stage"] == expected_stage
            and receipt["execution_manifest_sha256"] == expected_execution_manifest
            and receipt["source_manifest_sha256"] == manifest_sha
            and receipt["script_sha256"] == sha(RUN)
            and receipt["plan_sha256"] == sha(PLAN),
            f"{rel(receipt_path)} source/driver/binary binding differs")
    require(receipt["command"] == expected_main_command(stage, job, binary),
            f"{rel(receipt_path)} command differs from frozen capture")
    stem = job["name"]
    expected_artifacts = {
        f"{stem}.json", f"{stem}.catalog.json", f"{stem}.stdout",
        f"{stem}.stderr", f"{stem}.host.json",
    }
    validate_artifacts(folder, stem, receipt, expected_artifacts)
    rss = validate_time_stderr(folder / f"{stem}.stderr", rel(folder / f"{stem}.stderr"))
    report_path = folder / f"{stem}.json"
    catalog_path = folder / f"{stem}.catalog.json"
    identity = main_report_identity(read_json(report_path), job, binary, plan,
                                    catalog_path, rel(report_path))
    return {
        "name": stem, "stage": stage, "execution_stage": expected_stage,
        "lane": job["lane"], "repeat": job["repeat"], "shape": job["shape"],
        "case": job["case"], "samples": job["samples"], "warmup": job["warmup"],
        "report_sha256": sha(report_path), "catalog_sha256": sha(catalog_path),
        "receipt_sha256": sha(receipt_path), "binary_sha256": binary["sha256"],
        "identity": identity, "rss": rss,
    }, (start, end)


def validate_main_stage(stage: str, plan: dict[str, Any],
                        baseline_manifest_sha: str,
                        candidate_manifest_sha: str,
                        include_repeat_two: bool) -> tuple[list[dict[str, Any]], list[tuple[dt.datetime, dt.datetime, str]]]:
    manifest, manifest_sha = stage_manifest(stage)
    require(manifest_sha == (baseline_manifest_sha if stage == "baseline"
                             else candidate_manifest_sha),
            f"{stage} manifest binding differs")
    binaries = {
        "native": validate_binary(stage, "normal", manifest_sha),
        "alloc": validate_binary(stage, "alloc", manifest_sha),
    }
    jobs: list[dict[str, Any]] = []
    for lane in ("preflight", "native", "alloc"):
        expected = expected_main_jobs(lane, stage)
        if stage == "baseline" and not include_repeat_two:
            expected = [job for job in expected if job["repeat"] == 1]
        expected_names = {job["name"] for job in expected}
        prefix = f"{lane}-"
        actual_names = {
            path.name[:-len(".receipt.json")]
            for path in (HERE / stage).glob(prefix + "*.receipt.json")
        }
        require(actual_names == expected_names,
                f"{stage}/{lane} receipt matrix differs")
        jobs.extend(expected)
    rows: list[dict[str, Any]] = []
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for job in jobs:
        binary = binaries["alloc" if job["lane"] == "alloc" else "native"]
        row, times = validate_main_job(
            stage, job, binary, manifest_sha, baseline_manifest_sha,
            candidate_manifest_sha, plan,
        )
        rows.append(row)
        intervals.append((times[0], times[1], job["name"]))
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{stage} main receipts overlap")
    return rows, intervals


def stage_repeats(stage: str, prefixes: tuple[str, ...]) -> bool:
    """Return whether a stage has entered its second ABBA repeat.

    Candidate is always a complete two-repeat stage.  Baseline starts with
    repeat one and receives repeat two only after the candidate run.  A
    partially-created repeat is treated as malformed by the caller rather
    than being silently interpreted as a one-repeat preliminary result.
    """
    require(stage in STAGES, f"unknown stage {stage}")
    if stage == "candidate":
        return True
    for prefix in prefixes:
        if any("-r2-" in path.name for path in
               (HERE / stage).glob(prefix + "*.receipt.json")):
            return True
    return False


def validate_stage_builds(
    stage: str, manifest_sha: str
) -> tuple[dict[str, dict[str, Any]], list[tuple[dt.datetime, dt.datetime, str]]]:
    """Validate every retained executable and its source-bound build receipt."""
    folder = HERE / stage
    expected_kinds = ("normal", "guard-normal", "cap", "alloc", "guard-alloc")
    actual = {
        path.name[:-len(".receipt.json")]
        for path in folder.glob("build-*.receipt.json")
    }
    require(actual == {f"build-{kind}" for kind in expected_kinds},
            f"{stage} build receipt inventory differs")
    binaries: dict[str, dict[str, Any]] = {}
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for kind in expected_kinds:
        row, times = validate_build(stage, kind, manifest_sha)
        binaries[kind] = validate_binary(stage, kind, manifest_sha)
        intervals.append((times[0], times[1], row["name"]))
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{stage} build receipts overlap")
    return binaries, intervals


def expected_guard_command(stage: str, job: dict[str, Any],
                           binary: dict[str, Any]) -> list[str]:
    return [
        "taskset", "-c", str(CPU), binary["path"],
        "--shape", job["shape"], "--case", job["case"],
        "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
        "--json", str(HERE / stage / f"{job['name']}.json"),
    ]


def expected_cap_command(stage: str, job: dict[str, Any],
                         binary: dict[str, Any]) -> list[str]:
    return [
        "taskset", "-c", str(CPU), binary["path"],
        "--size", str(job["size"]), "--samples", str(job["samples"]),
        "--warmup", str(job["warmup"]),
        "--json", str(HERE / stage / f"{job['name']}.json"),
        "--fixture-out", str(HERE / stage / f"{job['name']}.zip"),
    ]


def validate_guard_capture(
    stage: str, job: dict[str, Any], binary: dict[str, Any],
    stage_manifest_sha: str, baseline_manifest_sha: str,
    candidate_manifest_sha: str,
) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime, str]]:
    folder = HERE / stage
    path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(path, rel(path))
    start, end = validate_receipt_common(receipt, rel(path))
    execution_stage, execution_manifest = expected_execution(
        stage, job["repeat"], baseline_manifest_sha, candidate_manifest_sha
    )
    require(
        receipt["exit_code"] == 0
        and receipt["binary_sha256"] == binary["sha256"]
        and receipt["execution_stage"] == execution_stage
        and receipt["execution_manifest_sha256"] == execution_manifest
        and receipt["source_manifest_sha256"] == stage_manifest_sha
        and receipt["script_sha256"] == sha(RUN)
        and receipt["plan_sha256"] == sha(PLAN)
        and receipt["command"] == expected_guard_command(stage, job, binary),
        f"{rel(path)} guard receipt binding differs",
    )
    stem = job["name"]
    expected = {
        f"{stem}.host.json", f"{stem}.json", f"{stem}.stdout",
        f"{stem}.stderr",
    }
    validate_artifacts(folder, stem, receipt, expected)
    return {
        "name": stem, "stage": stage, "repeat": job["repeat"],
        "lane": job["lane"], "shape": job["shape"], "case": job["case"],
        "receipt_sha256": sha(path), "report_sha256": sha(folder / f"{stem}.json"),
        "binary_sha256": binary["sha256"],
    }, (start, end, stem)


def validate_cap_capture(
    stage: str, job: dict[str, Any], binary: dict[str, Any],
    stage_manifest_sha: str, baseline_manifest_sha: str,
    candidate_manifest_sha: str,
) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime, str]]:
    folder = HERE / stage
    path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(path, rel(path))
    start, end = validate_receipt_common(receipt, rel(path))
    execution_stage, execution_manifest = expected_execution(
        stage, job["repeat"], baseline_manifest_sha, candidate_manifest_sha
    )
    require(
        receipt["exit_code"] == 0
        and receipt["binary_sha256"] == binary["sha256"]
        and receipt["execution_stage"] == execution_stage
        and receipt["execution_manifest_sha256"] == execution_manifest
        and receipt["source_manifest_sha256"] == stage_manifest_sha
        and receipt["script_sha256"] == sha(RUN)
        and receipt["plan_sha256"] == sha(PLAN)
        and receipt["command"] == expected_cap_command(stage, job, binary),
        f"{rel(path)} cap receipt binding differs",
    )
    stem = job["name"]
    expected = {
        f"{stem}.host.json", f"{stem}.json", f"{stem}.stdout",
        f"{stem}.stderr", f"{stem}.zip",
    }
    validate_artifacts(folder, stem, receipt, expected)
    return {
        "name": stem, "stage": stage, "repeat": job["repeat"],
        "size": job["size"], "receipt_sha256": sha(path),
        "report_sha256": sha(folder / f"{stem}.json"),
        "fixture_sha256": sha(folder / f"{stem}.zip"),
        "binary_sha256": binary["sha256"],
    }, (start, end, stem)


def validate_guard_cap_stage(
    stage: str, baseline_manifest_sha: str, candidate_manifest_sha: str,
    include_repeat_two: bool,
) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime, str]]]:
    """Validate supplemental receipt matrices before invoking their analyzer."""
    manifest, stage_manifest_sha = stage_manifest(stage)
    require(stage_manifest_sha == (
        baseline_manifest_sha if stage == "baseline" else candidate_manifest_sha
    ), f"{stage} supplemental manifest binding differs")
    binaries = {
        "normal": validate_binary(stage, "guard-normal", stage_manifest_sha),
        "alloc": validate_binary(stage, "guard-alloc", stage_manifest_sha),
        "cap": validate_binary(stage, "cap", stage_manifest_sha),
    }
    actual_r2 = stage_repeats(stage, ("guard-native-", "guard-alloc-", "cap-"))
    require(actual_r2 == include_repeat_two,
            f"{stage} supplemental repeat inventory differs")
    guard_rows: dict[str, list[dict[str, Any]]] = {}
    cap_rows: list[dict[str, Any]] = []
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for lane in ("normal", "alloc"):
        jobs = expected_guard_jobs(lane)
        if not include_repeat_two:
            jobs = [job for job in jobs if job["repeat"] == 1]
        prefix = "guard-native-" if lane == "normal" else "guard-alloc-"
        expected_names = {job["name"] for job in jobs}
        actual_names = {
            path.name[:-len(".receipt.json")]
            for path in (HERE / stage).glob(prefix + "*.receipt.json")
        }
        require(actual_names == expected_names,
                f"{stage}/{prefix[:-1]} receipt matrix differs")
        rows: list[dict[str, Any]] = []
        for job in jobs:
            row, times = validate_guard_capture(
                stage, job, binaries["normal" if lane == "normal" else "alloc"],
                stage_manifest_sha, baseline_manifest_sha, candidate_manifest_sha,
            )
            rows.append(row)
            intervals.append(times)
        guard_rows[lane] = rows
    jobs = expected_cap_jobs()
    if not include_repeat_two:
        jobs = [job for job in jobs if job["repeat"] == 1]
    expected_names = {job["name"] for job in jobs}
    actual_names = {
        path.name[:-len(".receipt.json")]
        for path in (HERE / stage).glob("cap-*.receipt.json")
    }
    require(actual_names == expected_names, f"{stage}/cap receipt matrix differs")
    for job in jobs:
        row, times = validate_cap_capture(
            stage, job, binaries["cap"], stage_manifest_sha,
            baseline_manifest_sha, candidate_manifest_sha,
        )
        cap_rows.append(row)
        intervals.append(times)
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{stage} guard/cap receipts overlap")
    return {"stage": stage, "manifest_sha256": stage_manifest_sha,
            "guard": guard_rows, "cap": cap_rows,
            "repeat_two": include_repeat_two}, intervals


def validate_stage_raw_inventory(stage: str, include_repeat_two: bool,
                                 profile_required: bool) -> set[str]:
    """Require that a completed stage contains exactly its declared files."""
    require(stage in STAGES, f"unknown stage {stage}")
    folder = HERE / stage
    require(folder.is_dir() and not folder.is_symlink(),
            f"{stage} stage is not a directory")
    expected = {"source-manifest.json", "source.patch"}
    for kind in ("normal", "alloc", "guard-normal", "guard-alloc", "cap"):
        build = f"build-{kind}"
        expected |= {
            f"{build}.receipt.json", f"{build}.host.json",
            f"{build}.stdout", f"{build}.stderr",
            f"binary-{kind}.json",
        }
    for lane in ("preflight", "native", "alloc"):
        jobs = expected_main_jobs(lane, stage)
        if not include_repeat_two:
            jobs = [job for job in jobs if job["repeat"] == 1]
        for job in jobs:
            stem = job["name"]
            expected |= {
                f"{stem}.receipt.json", f"{stem}.host.json",
                f"{stem}.json", f"{stem}.catalog.json",
                f"{stem}.stdout", f"{stem}.stderr",
            }
    for lane in ("normal", "alloc"):
        jobs = expected_guard_jobs(lane)
        if not include_repeat_two:
            jobs = [job for job in jobs if job["repeat"] == 1]
        for job in jobs:
            stem = job["name"]
            expected |= {
                f"{stem}.receipt.json", f"{stem}.host.json",
                f"{stem}.json", f"{stem}.stdout", f"{stem}.stderr",
            }
    jobs = expected_cap_jobs()
    if not include_repeat_two:
        jobs = [job for job in jobs if job["repeat"] == 1]
    for job in jobs:
        stem = job["name"]
        expected |= {
            f"{stem}.receipt.json", f"{stem}.host.json", f"{stem}.json",
            f"{stem}.stdout", f"{stem}.stderr", f"{stem}.zip",
        }
    if profile_required:
        for job in expected_profile_jobs():
            stem = job["name"]
            expected |= {
                f"{stem}.receipt.json", f"{stem}.host.json",
                f"{stem}.json", f"{stem}.catalog.json",
                f"{stem}.stdout", f"{stem}.stderr", f"{stem}.callgrind",
            }
    actual_entries = list(folder.iterdir())
    require(all(item.is_file() and not item.is_symlink()
                for item in actual_entries),
            f"{stage} stage contains a non-file artifact")
    actual = {item.name for item in actual_entries}
    require(actual == expected, f"{stage} raw artifact inventory differs")
    return expected


def bundle_snapshot() -> dict[str, str]:
    """Hash retained files so analyzer replay cannot mutate evidence."""
    snapshot: dict[str, str] = {}
    for path in HERE.rglob("*"):
        if path.is_symlink():
            continue
        if path.is_file():
            snapshot[rel(path)] = sha(path)
    return snapshot


def import_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None,
            f"cannot load {rel(path)}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, ImportError, SyntaxError, ValueError) as error:
        raise VerificationError(f"cannot load {rel(path)}: {error}") from error
    return module


def deterministic_document(value: Any) -> bytes:
    try:
        return (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    except (TypeError, ValueError) as error:
        raise VerificationError(f"analysis document is not deterministic JSON: {error}") from error


def analyzer_output(names: tuple[str, ...], label: str) -> tuple[Path, list[Path]]:
    paths = [HERE / name for name in names if (HERE / name).exists()]
    for path in paths:
        require(path.is_file() and not path.is_symlink(),
                f"{label} output is not a regular file: {rel(path)}")
    if not paths:
        raise IncompleteError(f"{label} canonical report is missing")
    return paths[0], paths


def replay_analyzer(path: Path, module_name: str, output_names: tuple[str, ...],
                    expected_schema: str, label: str) -> dict[str, Any]:
    before = bundle_snapshot()
    module = import_module(path, module_name)
    require(hasattr(module, "analyze") and callable(module.analyze),
            f"{label} has no callable analyze()")
    try:
        value = module.analyze()
    except FileNotFoundError as error:
        raise IncompleteError(f"{label} is missing {error.filename}") from error
    except (KeyError, TypeError, ValueError, OSError, AssertionError) as error:
        text = str(error)
        if "missing" in text.lower() or "not found" in text.lower():
            raise IncompleteError(f"{label} is incomplete: {text}") from error
        raise VerificationError(f"{label} failed: {text}") from error
    after = bundle_snapshot()
    require(before == after, f"{label} replay mutated retained evidence")
    require(isinstance(value, dict) and value.get("schema") == expected_schema,
            f"{label} schema differs")
    encoded = deterministic_document(value)
    output, paths = analyzer_output(output_names, label)
    for candidate in paths:
        require(candidate.read_bytes() == encoded,
                f"{label} canonical replay differs: {rel(candidate)}")
    return {"status": value.get("status"), "schema": value["schema"],
            "path": rel(output), "sha256": sha(output),
            "document": value}


def require_analyzer_pass(replay: dict[str, Any], label: str) -> None:
    """Keep an analyzer's explicit pending state distinguishable from failure."""
    status = replay.get("status")
    if status in ("pending", "incomplete", "not-present"):
        raise IncompleteError(f"{label} remains {status}")
    require(status == "pass", f"{label} did not produce a passing analysis")


def compact_main_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {key: row[key] for key in (
            "name", "stage", "execution_stage", "lane", "repeat",
            "shape", "case", "samples", "warmup", "report_sha256",
            "catalog_sha256", "receipt_sha256", "binary_sha256",
        ) if key in row}
        for row in rows
    ]


def validate_main_bundle() -> dict[str, Any]:
    """Validate both complete main matrices and their retained binaries."""
    plan = validate_plan()
    validate_supplemental_inputs()
    validate_analysis_inputs()
    lock = validate_workspace_lock()
    validate_public_tests()
    baseline = validate_source("baseline")
    candidate = validate_source("candidate")
    candidate_correctness = validate_candidate_correctness()
    baseline_manifest_sha = baseline["manifest_sha256"]
    candidate_manifest_sha = candidate["manifest_sha256"]
    builds: dict[str, Any] = {}
    captures: dict[str, Any] = {}
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for stage in STAGES:
        builds[stage], build_intervals = validate_stage_builds(
            stage, baseline_manifest_sha if stage == "baseline"
            else candidate_manifest_sha,
        )
        intervals.extend(build_intervals)
        rows, capture_intervals = validate_main_stage(
            stage, plan, baseline_manifest_sha, candidate_manifest_sha, True
        )
        captures[stage] = rows
        intervals.extend(capture_intervals)
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "main build/capture receipts overlap")
    return {
        "status": "pass", "plan_sha256": plan["sha256"],
        "workspace_lock": lock, "source": {
            "baseline": baseline, "candidate": candidate,
        }, "baseline_correctness": validate_baseline_correctness(),
        "candidate_correctness": candidate_correctness,
        "builds": builds, "captures": {
            stage: {"rows": compact_main_rows(rows), "count": len(rows)}
            for stage, rows in captures.items()
        }, "intervals": len(intervals),
    }


def validate_main_preliminary() -> dict[str, Any]:
    """Validate the completed baseline while the candidate is still pending."""
    plan = validate_plan()
    validate_supplemental_inputs()
    validate_analysis_inputs()
    lock = validate_workspace_lock()
    validate_public_tests()
    correctness = validate_baseline_correctness()
    baseline = validate_source("baseline")
    builds, build_intervals = validate_stage_builds(
        "baseline", baseline["manifest_sha256"]
    )
    rows, capture_intervals = validate_main_stage(
        "baseline", plan, baseline["manifest_sha256"],
        baseline["manifest_sha256"], False,
    )
    guard_rows, guard_intervals = validate_guard_cap_stage(
        "baseline", baseline["manifest_sha256"],
        baseline["manifest_sha256"], False,
    )
    intervals = build_intervals + capture_intervals + guard_intervals
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "baseline build/capture receipts overlap")
    validate_stage_raw_inventory("baseline", False, False)
    return {
        "status": "preliminary", "complete": False,
        "baseline_complete": True,
        "candidate": {"status": "pending"},
        "workspace_lock": lock, "source": baseline,
        "baseline_correctness": correctness,
        "builds": builds, "main_rows": compact_main_rows(rows),
        "guard_cap": guard_rows, "intervals": len(intervals),
    }


def validate_metrics_analysis() -> dict[str, Any]:
    """Replay the frozen full main analyzer and bind its corrected gates."""
    bundle = validate_main_bundle()
    validate_adr()
    replay = replay_analyzer(
        METRICS, "xlsx_0552_metrics_analyzer", METRICS_REPORT_NAMES,
        "xlsx_multisource_edit_metrics_0552_v1", "main metrics analyzer",
    )
    require_analyzer_pass(replay, "main metrics analyzer")
    document = replay["document"]
    require(replay["sha256"] and sha(METRICS) == ANALYZER_METRICS_SHA256,
            "main metrics analyzer is not the frozen corrected revision")
    require(document.get("plan_sha256") == sha(PLAN)
            and document.get("capture_sha256") == sha(CAPTURE)
            and document.get("run_sha256") == sha(RUN),
            "main metrics analyzer driver binding differs")
    require(document.get("stage") == "matched baseline/candidate",
            "main metrics analyzer stage differs")
    source_manifests = document.get("source_manifests")
    require(isinstance(source_manifests, dict)
            and set(source_manifests) == set(STAGES),
            "main metrics source manifest inventory differs")
    for stage in STAGES:
        item = source_manifests[stage]
        require(isinstance(item, dict)
                and item.get("stage") == stage
                and item.get("sha256") == bundle["source"][stage]["manifest_sha256"],
                f"main metrics {stage} manifest binding differs")
    binaries = document.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == set(STAGES),
            "main metrics binary stage inventory differs")
    for stage in STAGES:
        require(isinstance(binaries[stage], dict)
                and set(binaries[stage]) == {"normal", "alloc"},
                f"main metrics {stage} binary inventory differs")
        for kind, binary in binaries[stage].items():
            expected = bundle["builds"][stage][kind]
            require(isinstance(binary, dict)
                    and binary.get("sha256") == expected["sha256"]
                    and binary.get("path") == expected["path"]
                    and binary.get("bytes") == expected["bytes"],
                    f"main metrics {stage}/{kind} binary binding differs")
    gates = document.get("main_gates")
    require(isinstance(gates, dict)
            and set(gates) == {
                "primary_one_percent", "one_cell_latency", "workflow_memory",
                "allocation", "correctness_identity",
                "all_frozen_main_gates_pass", "external_controls_required",
            }
            and isinstance(gates.get("all_frozen_main_gates_pass"), bool),
            "main metrics frozen gate result is missing")
    for name in ("primary_one_percent", "one_cell_latency", "workflow_memory",
                 "allocation"):
        group = gates[name]
        require(isinstance(group, dict)
                and isinstance(group.get("pass"), bool)
                and isinstance(group.get("checks"), list)
                and isinstance(group.get("check_count"), int)
                and not isinstance(group["check_count"], bool)
                and group["check_count"] == len(group["checks"])
                and group["check_count"] > 0,
                f"main metrics {name} gate shape differs")
        require(all(isinstance(item, dict)
                    and isinstance(item.get("pass"), bool)
                    for item in group["checks"]),
                f"main metrics {name} gate checks are malformed")
    correctness_gate = gates["correctness_identity"]
    require(isinstance(correctness_gate, dict)
            and isinstance(correctness_gate.get("pass"), bool)
            and isinstance(correctness_gate.get("identity_equal"), bool)
            and isinstance(correctness_gate.get("exact_checks"), list)
            and isinstance(correctness_gate.get("exact_check_count"), int)
            and not isinstance(correctness_gate["exact_check_count"], bool)
            and correctness_gate["exact_check_count"] ==
            len(correctness_gate["exact_checks"]),
            "main metrics correctness gate shape differs")
    controls = gates["external_controls_required"]
    require(controls == {
        "status": "pending", "validated_here": False,
        "required": ["guard", "cap", "quality", "profile"],
        "reason": "main metrics analyzer does not own guard, cap, quality, or profile evidence",
    }, "main metrics external gate shape differs")
    main_pass = gates["all_frozen_main_gates_pass"]
    disposition = document.get("disposition")
    require(isinstance(disposition, str),
            "main metrics disposition is missing")
    if main_pass:
        require(disposition == "pending external guard/cap/quality/profile evidence",
                "main metrics passing disposition differs")
    else:
        require(disposition == "reject: one or more frozen main metrics gates failed",
                "main metrics rejected disposition differs")
    comparisons = document.get("comparisons")
    require(isinstance(comparisons, dict)
            and isinstance(comparisons.get("numeric"), list)
            and isinstance(comparisons.get("exact_source"), list)
            and isinstance(comparisons.get("adverse"), list),
            "main metrics comparison inventory differs")
    require(isinstance(document.get("repeat_drift"), list)
            and isinstance(document.get("repeat_drift_over_five_percent"), list),
            "main metrics repeat drift inventory differs")
    return {
        "status": "pass", "report": {
            "path": replay["path"], "sha256": replay["sha256"],
            "schema": replay["schema"],
        }, "main_gates": gates,
        "source": bundle["source"], "builds": bundle["builds"],
        "capture_intervals": bundle["intervals"],
    }


def validate_guard_cap_analysis() -> dict[str, Any]:
    """Replay the finalized independent guard/cap analyzer."""
    plan = validate_plan()
    validate_supplemental_inputs()
    analysis_inputs = validate_analysis_inputs()
    validate_workspace_lock()
    validate_public_tests()
    correctness = validate_baseline_correctness()
    baseline = validate_source("baseline")
    candidate = validate_source("candidate")
    base_sha, candidate_sha = baseline["manifest_sha256"], candidate["manifest_sha256"]
    stages: dict[str, Any] = {}
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for stage in STAGES:
        rows, stage_intervals = validate_guard_cap_stage(
            stage, base_sha, candidate_sha, True
        )
        stages[stage] = rows
        intervals.extend(stage_intervals)
    replay = replay_analyzer(
        GUARDS, "xlsx_0552_guard_analyzer", GUARDS_REPORT_NAMES,
        "litchi.xlsx.guard-cap-analysis.v1", "guard/cap analyzer",
    )
    require_analyzer_pass(replay, "guard/cap analyzer")
    guard_analyzer_sha = sha(GUARDS)
    guard_amendment = analysis_inputs.get("guard_analyzer_amendment")
    expected_guard_analyzer_sha = (
        guard_amendment["amended_sha256"]
        if guard_amendment is not None else ANALYZER_GUARDS_SHA256
    )
    require(guard_analyzer_sha == expected_guard_analyzer_sha,
            "guard/cap analyzer is not the frozen finalized revision")
    document = replay["document"]
    require(document.get("analyzer_sha256") == expected_guard_analyzer_sha
            and document.get("plan_sha256") == sha(PLAN)
            and document.get("capture_sha256") == sha(CAPTURE)
            and document.get("run_sha256") == sha(RUN)
            and document.get("guarded_capture_sha256") == sha(GUARDED),
            "guard/cap analyzer driver binding differs")
    comparison = document.get("comparison")
    require(isinstance(comparison, dict)
            and isinstance(comparison.get("guard"), dict)
            and isinstance(comparison.get("cap"), dict),
            "guard/cap comparison inventory differs")
    guard_pass = comparison["guard"].get("admission_passed")
    cap_pass = comparison["cap"].get("admission_passed")
    require(isinstance(guard_pass, bool) and isinstance(cap_pass, bool),
            "guard/cap admission fields are missing")
    expected_admission_status = "pass" if guard_pass and cap_pass else "reject"
    require(document.get("admission_status") == expected_admission_status,
            "guard/cap admission status does not match both gates")
    require(isinstance(comparison["guard"].get("adverse_flags_over_five_percent"), list)
            and isinstance(comparison["guard"].get("same_build_drift"), list)
            and isinstance(comparison["cap"].get("adverse_flags_over_five_percent"), list)
            and isinstance(comparison["cap"].get("same_build_drift"), list),
            "guard/cap adverse or drift arrays are missing")
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "guard/cap receipt intervals overlap")
    return {
        "status": "pass", "report": {
            "path": replay["path"], "sha256": replay["sha256"],
            "schema": replay["schema"],
        }, "admission_status": document["admission_status"],
        "guard_admission_passed": guard_pass,
        "cap_admission_passed": cap_pass, "stages": stages,
        "intervals": len(intervals),
        "analyzer_amendment": guard_amendment,
    }


def profile_bool(value: dict[str, Any], names: tuple[str, ...],
                label: str) -> bool | None:
    observed: list[bool] = []
    for name in names:
        if isinstance(value.get(name), bool):
            observed.append(value[name])
    nested = value.get("pilot")
    if isinstance(nested, dict):
        for name in names:
            if isinstance(nested.get(name), bool):
                observed.append(nested[name])
    require(not observed or all(item is observed[0] for item in observed),
            f"{label} has contradictory boolean aliases")
    return observed[0] if observed else None


def validate_profile_capture(
    stage: str, job: dict[str, Any], binary: dict[str, Any],
    stage_manifest_sha: str,
) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime, str]]:
    folder = HERE / stage
    path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(path, rel(path))
    start, end = validate_receipt_common(receipt, rel(path))
    callgrind = folder / f"{job['name']}.callgrind"
    command = [
        "taskset", "-c", str(CPU), "valgrind", "--tool=callgrind",
        "--vgdb=no", "--vgdb-prefix=" + str(TARGET / "tmp/vgdb"),
        "--collect-atstart=no", "--toggle-collect=" + PROFILE_OWNER,
        "--zero-before=" + PROFILE_OWNER, "--dump-after=" + PROFILE_OWNER,
        "--callgrind-out-file=" + str(callgrind),
        binary["path"], "--case", PROFILE_CASE,
        "--xlsx-cell-crud-shape", job["shape"], "--samples", "1",
        "--warmup", "0", "--json", str(folder / f"{job['name']}.json"),
        "--corpus-manifest", str(folder / f"{job['name']}.catalog.json"),
    ]
    require(
        receipt["exit_code"] == 0
        and receipt["binary_sha256"] == binary["sha256"]
        and receipt["execution_stage"] == "candidate"
        and receipt["execution_manifest_sha256"] == sha(CANDIDATE / "source-manifest.json")
        and receipt["source_manifest_sha256"] == stage_manifest_sha
        and receipt["script_sha256"] == sha(RUN)
        and receipt["plan_sha256"] == sha(PLAN)
        and receipt["command"] == command,
        f"{rel(path)} profile receipt binding differs",
    )
    stem = job["name"]
    expected = {
        f"{stem}.host.json", f"{stem}.json", f"{stem}.catalog.json",
        f"{stem}.stdout", f"{stem}.stderr", f"{stem}.callgrind",
    }
    validate_artifacts(folder, stem, receipt, expected)
    profile_text = read_text(callgrind, rel(callgrind))
    require(any(line.strip() == "events: Ir" for line in profile_text.splitlines()),
            f"{rel(callgrind)} lacks the Ir event declaration")
    require(PROFILE_OWNER in profile_text,
            f"{rel(callgrind)} lacks the exact profiled owner")
    summaries = [
        int(match.group(1))
        for line in profile_text.splitlines()
        if (match := re.fullmatch(r"\s*summary:\s*([0-9]+)\s*", line))
    ]
    require(len(summaries) == 1 and summaries[0] > 0,
            f"{rel(callgrind)} has no positive unique Ir summary")
    job_for_report = {
        **job, "lane": "native", "case": PROFILE_CASE,
        "samples": 1, "warmup": 0,
    }
    identity = main_report_identity(
        read_json(folder / f"{stem}.json"), job_for_report, binary,
        validate_plan(), folder / f"{stem}.catalog.json", rel(folder / f"{stem}.json"),
    )
    return {
        "name": stem, "stage": stage, "repeat": job["repeat"],
        "shape": job["shape"], "report_sha256": sha(folder / f"{stem}.json"),
        "catalog_sha256": sha(folder / f"{stem}.catalog.json"),
        "callgrind_sha256": sha(callgrind), "receipt_sha256": sha(path),
        "identity": identity, "binary_sha256": binary["sha256"],
        "instruction_references": summaries[0],
    }, (start, end, stem)


def profile_pilot_gates(
    metrics: dict[str, Any], guards: dict[str, Any] | None = None,
) -> dict[str, bool]:
    """Return the three independent gates that make profiling conditional."""
    gates = metrics.get("main_gates")
    require(isinstance(gates, dict), "main metrics pilot gates are missing")
    main_gate = gates.get("all_frozen_main_gates_pass")
    require(isinstance(main_gate, bool),
            "main metrics aggregate pilot gate is malformed")
    for name in ("one_cell_latency", "workflow_memory"):
        group = gates.get(name)
        require(isinstance(group, dict) and isinstance(group.get("pass"), bool),
                f"main metrics {name} pilot gate is malformed")
    if guards is None:
        guards = validate_guard_cap_analysis()
    guard_gate = guards.get("guard_admission_passed")
    cap_gate = guards.get("cap_admission_passed")
    require(isinstance(guard_gate, bool) and isinstance(cap_gate, bool),
            "guard/cap pilot gates are missing")
    return {"main": main_gate, "guard": guard_gate, "cap": cap_gate}


def profile_pilot_from_metrics(
    metrics: dict[str, Any], guards: dict[str, Any] | None = None,
) -> bool:
    """Return whether the complete conditional-profile pilot passed."""
    return all(profile_pilot_gates(metrics, guards).values())


def validate_profiles(*, require_decision: bool = True,
                      pilot_expected: bool | None = None,
                      metrics_result: dict[str, Any] | None = None,
                      guards_result: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate conditional Callgrind evidence after the complete pilot."""
    decision_paths = [HERE / name for name in PROFILE_DECISION_NAMES
                      if (HERE / name).exists()]
    require(len(decision_paths) <= 1,
            "multiple profile pilot decisions are retained")
    decision_path: Path | None = decision_paths[0] if decision_paths else None

    # The profile decision records the main report digest and all three pilot
    # gates.  Resolve both reports before accepting a decision so a narrow
    # native/memory helper cannot make a profile lane mandatory or vacuous.
    if metrics_result is None:
        metrics_result = validate_metrics_analysis()
    if guards_result is None:
        guards_result = validate_guard_cap_analysis()
    pilot_gates = profile_pilot_gates(metrics_result, guards_result)
    computed_pilot = all(pilot_gates.values())
    if pilot_expected is None:
        pilot_expected = computed_pilot
    require(pilot_expected is computed_pilot,
            "profile pilot expectation differs from the complete pilot gates")
    report = metrics_result.get("report")
    require(isinstance(report, dict) and is_hash(report.get("sha256")),
            "main metrics report digest is missing from profile pilot")
    metrics_sha = report["sha256"]

    if decision_path is None:
        if not pilot_expected:
            profile_receipts = [
                path for stage in STAGES
                for path in (HERE / stage).glob("profile-*.receipt.json")
            ]
            require(not profile_receipts,
                    "profile captures exist although the pilot did not pass")
            return {
                "status": "skipped", "required": False,
                "pilot_passed": False, "gate_passed": True,
                "reported_gate_passed": None, "decision": None,
                "decision_sha256": None, "profile_rows": [], "intervals": 0,
                "pilot_gates": pilot_gates,
                "main_analysis_sha256": metrics_sha,
            }
        if require_decision:
            raise IncompleteError("profile pilot decision is missing")
        return {
            "status": "pending", "pilot_passed": None, "required": True,
            "pilot_gates": pilot_gates, "main_analysis_sha256": metrics_sha,
        }

    require(decision_path.is_file() and not decision_path.is_symlink(),
            f"{rel(decision_path)} is not a regular file")
    decision = read_json(decision_path, rel(decision_path))
    expected_keys = {
        "schema", "status", "scope", "pilot_passed", "profile_required",
        "profile_gate_passed", "main_analysis_sha256", "pilot_gates",
        "profile_rows", "reason",
    }
    require(isinstance(decision, dict) and set(decision) == expected_keys,
            f"{rel(decision_path)} profile decision envelope differs")
    require(decision["schema"] == PROFILE_DECISION_SCHEMA,
            f"{rel(decision_path)} profile decision schema differs")
    require(decision["status"] in {"pass", "failed", "reject", "rejected", "skipped"},
            f"{rel(decision_path)} profile decision status differs")
    require(isinstance(decision["scope"], str) and decision["scope"].strip()
            and isinstance(decision["reason"], str) and decision["reason"].strip(),
            f"{rel(decision_path)} profile decision scope/reason is missing")
    require(isinstance(decision["pilot_passed"], bool)
            and isinstance(decision["profile_required"], bool)
            and isinstance(decision["profile_gate_passed"], bool),
            f"{rel(decision_path)} profile decision booleans are malformed")
    check_hash(decision["main_analysis_sha256"],
               f"{rel(decision_path)}.main_analysis_sha256")
    require(decision["main_analysis_sha256"] == metrics_sha,
            f"{rel(decision_path)} main analysis binding differs")
    require(isinstance(decision["pilot_gates"], dict)
            and set(decision["pilot_gates"]) == {"main", "guard", "cap"}
            and all(isinstance(item, bool)
                    for item in decision["pilot_gates"].values())
            and decision["pilot_gates"] == pilot_gates,
            f"{rel(decision_path)} pilot gates differ from independent evidence")
    require(isinstance(decision["profile_rows"], list),
            f"{rel(decision_path)} profile_rows is not an array")
    require(decision["pilot_passed"] is pilot_expected
            and decision["profile_required"] is pilot_expected,
            f"{rel(decision_path)} profile pilot/required binding differs")
    reported_gate_passed = decision["profile_gate_passed"]

    if not pilot_expected:
        require(decision["status"] == "skipped"
                and reported_gate_passed is True
                and decision["profile_rows"] == [],
                f"{rel(decision_path)} failed pilot does not have an explicit profile skip")
        profile_receipts = [
            path for stage in STAGES
            for path in (HERE / stage).glob("profile-*.receipt.json")
        ]
        require(not profile_receipts,
                "profile captures exist although the pilot did not pass")
        return {
            "status": "pass", "required": False, "pilot_passed": False,
            "gate_passed": True, "reported_gate_passed": reported_gate_passed,
            "decision": rel(decision_path),
            "decision_sha256": sha(decision_path), "profile_rows": [],
            "intervals": 0, "pilot_gates": pilot_gates,
            "main_analysis_sha256": metrics_sha,
        }

    require(decision["status"] in {"pass", "failed", "reject", "rejected"}
            and decision["profile_rows"]
            and all(isinstance(row, dict)
                    and isinstance(row.get("passed"), bool)
                    for row in decision["profile_rows"]),
            f"{rel(decision_path)} required profile rows are incomplete")
    baseline = validate_source("baseline")
    candidate = validate_source("candidate")
    rows: dict[str, list[dict[str, Any]]] = {}
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for stage, manifest in (
        ("baseline", baseline["manifest_sha256"]),
        ("candidate", candidate["manifest_sha256"]),
    ):
        binary = validate_binary(stage, "normal", manifest)
        expected = expected_profile_jobs()
        actual = {
            path.name[:-len(".receipt.json")]
            for path in (HERE / stage).glob("profile-*.receipt.json")
        }
        require(actual == {job["name"] for job in expected},
                f"{stage} profile receipt matrix differs")
        stage_rows: list[dict[str, Any]] = []
        stage_times: list[tuple[dt.datetime, dt.datetime, str]] = []
        for job in expected:
            row, times = validate_profile_capture(stage, job, binary, manifest)
            stage_rows.append(row)
            stage_times.append(times)
            intervals.append(times)
        ordered_stage = sorted(stage_times, key=lambda item: (item[0], item[1], item[2]))
        require([item[2] for item in ordered_stage] == [
            job["name"] for job in expected
        ], f"{stage} profile serial order differs")
        rows[stage] = stage_rows
    by_key = {
        (row["stage"], row["repeat"], row["shape"]): row
        for stage_rows in rows.values() for row in stage_rows
    }
    ir_rows: list[dict[str, Any]] = []
    for repeat in (1, 2):
        for shape in SHAPES:
            before = by_key[("baseline", repeat, shape)]["instruction_references"]
            after = by_key[("candidate", repeat, shape)]["instruction_references"]
            ir_rows.append({
                "repeat": repeat, "shape": shape,
                "baseline": before, "candidate": after,
                "reduction_percent": (before - after) * 100.0 / before,
                "passed": after < before,
            })
    ir_gate_passed = all(row["passed"] for row in ir_rows)
    require(reported_gate_passed is ir_gate_passed,
            f"{rel(decision_path)} profile gate differs from Ir replay")
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "profile receipts overlap")
    return {
        "status": "pass", "required": True, "pilot_passed": True,
        "gate_passed": ir_gate_passed,
        "reported_gate_passed": reported_gate_passed,
        "decision": rel(decision_path),
        "decision_sha256": sha(decision_path), "profile_rows": rows,
        "instruction_references": ir_rows, "intervals": len(intervals),
        "pilot_gates": pilot_gates, "main_analysis_sha256": metrics_sha,
    }


def validate_quality_plan() -> dict[str, Any]:
    value = read_json(QUALITY_PLAN, "quality-plan.json")
    require(isinstance(value, dict)
            and set(value) == {"scope", "commands"},
            "quality-plan envelope differs")
    require(value["scope"] == (
        "Compact XLSX source-cell proof candidate or restored baseline plus public guards; "
        "warning-denied owner checks and workspace feature check; no iWork optimization."
    ), "quality-plan scope differs")
    commands = value["commands"]
    require(isinstance(commands, dict) and len(commands) == 11,
            "quality-plan command count differs")
    for name, command in commands.items():
        require(isinstance(name, str) and re.fullmatch(r"[A-Za-z0-9_-]+", name),
                "quality-plan command name differs")
        require(isinstance(command, list) and command
                and all(isinstance(item, str) and item for item in command),
                f"quality-plan command is malformed: {name}")
    return {"status": "pass", "sha256": sha(QUALITY_PLAN),
            "scope": value["scope"], "commands": commands}


def quality_expected_manifest(stage: str) -> tuple[dict[str, str], str]:
    manifest, digest = stage_manifest(stage)
    expected = dict(manifest)
    expected["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    return dict(sorted(expected.items())), digest


def validate_check_attempt(path: Path, expected_manifest: dict[str, str] | None = None
                           ) -> dict[str, Any]:
    """Validate one source-bound check_attempt receipt, including failures."""
    require(path.parent == HERE / "check-attempts"
            and re.fullmatch(r"[A-Za-z0-9_-]+", path.name) is not None,
            "check_attempt path escapes its evidence root")
    label = rel(path)
    require(path.is_dir() and not path.is_symlink(), f"{label} is not a directory")
    receipt_path = path / "receipt.json"
    receipt = read_json(receipt_path, label + "/receipt.json")
    require(isinstance(receipt, dict) and set(receipt) == ATTEMPT_KEYS,
            f"{label} receipt inventory differs")
    start, end = interval(receipt, label)
    require(receipt["cwd"] == str(REPO)
            and isinstance(receipt["command"], list) and receipt["command"]
            and all(isinstance(item, str) and item for item in receipt["command"]),
            f"{label} command/cwd differs")
    require(isinstance(receipt["exit_code"], int)
            and not isinstance(receipt["exit_code"], bool)
            and isinstance(receipt["source_stable"], bool),
            f"{label} status fields differ")
    require(receipt["source_stable"] is True,
            f"{label} source was unstable during check")
    env = receipt["environment"]
    require(isinstance(env, dict) and set(env) == ATTEMPT_ENVIRONMENT,
            f"{label} environment inventory differs")
    require(env["TMPDIR"] == str(TARGET / "tmp")
            and env["CARGO_TARGET_DIR"] == str(TARGET)
            and env["CARGO_BUILD_JOBS"] == "2"
            and env["CARGO_INCREMENTAL"] == "0"
            and env["RUSTDOCFLAGS"] == "-D warnings",
            f"{label} controlled quality environment differs")
    for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD"):
        require(env[name] is None, f"{label} uncontrolled environment differs: {name}")
    source_path = path / "source-manifest.json"
    manifest_value = source_manifest(source_path, label + "/source-manifest.json")
    require(manifest_value.get("Cargo.lock") == WORKSPACE_LOCK_SHA256,
            f"{label} workspace Cargo.lock binding differs")
    if expected_manifest is not None:
        require(manifest_value == expected_manifest,
                f"{label} source manifest differs from selected stage")
    require(receipt["source_manifest_sha256"] == sha(source_path)
            and receipt["script_sha256"] == sha(CHECK_ATTEMPT)
            and receipt["run_sha256"] == sha(RUN)
            and receipt["plan_sha256"] == sha(PLAN),
            f"{label} driver/source hash binding differs")
    artifacts = receipt["artifacts"]
    expected_artifacts = {"source-manifest.json", "stderr", "stdout",
                          "tracked-source.patch"}
    require(isinstance(artifacts, dict)
            and set(artifacts) == expected_artifacts,
            f"{label} artifact inventory differs")
    actual = {
        item.name for item in path.iterdir()
        if item.is_file() and not item.is_symlink()
    }
    require(actual == expected_artifacts | {"receipt.json"},
            f"{label} raw artifact inventory differs")
    for name, digest in artifacts.items():
        check_hash(digest, f"{label}/{name}")
        require(sha(path / name) == digest, f"{label}/{name} digest differs")
    return {
        "path": label, "receipt_sha256": sha(receipt_path),
        "exit_code": receipt["exit_code"], "start_utc": receipt["start_utc"],
        "end_utc": receipt["end_utc"], "seconds": receipt["seconds"],
        "source_manifest_sha256": receipt["source_manifest_sha256"],
        "command": receipt["command"], "times": (start, end),
    }


def validate_all_check_attempts() -> dict[str, Any]:
    root = HERE / "check-attempts"
    if not root.exists():
        raise IncompleteError("check-attempts directory is missing")
    require(root.is_dir() and not root.is_symlink(),
            "check-attempts is not a directory")
    rows: list[dict[str, Any]] = []
    for path in sorted(root.iterdir(), key=lambda item: item.name):
        require(path.name and re.fullmatch(r"[A-Za-z0-9_-]+", path.name),
                f"check-attempt label is unsafe: {path.name}")
        rows.append(validate_check_attempt(path))
    require(rows, "no check_attempt receipts are retained")
    return {"status": "pass", "count": len(rows), "rows": rows}


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
CANDIDATE_PRIOR_CHECKS = {
    2: ("candidate-compile-01", 101, [
        "cargo", "check", "--locked", "-p", "litchi-xlsx", "--all-features",
    ]),
    3: ("candidate-compile-02", 101, [
        "cargo", "check", "--locked", "-p", "litchi-xlsx", "--all-features",
    ]),
    4: ("candidate-compile-03", 101, [
        "cargo", "check", "--locked", "-p", "litchi-xlsx", "--all-features",
    ]),
    5: ("candidate-proof-tests-01", 101, [
        "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
        "--all-features", "--lib", "source_proof", "--", "--test-threads=2",
    ]),
    6: ("baseline-public-exact-05", 0, [
        "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
        "--all-features", "--test", "source_backed_cell_values",
        "public_exact_output", "--", "--test-threads=2",
    ]),
    7: ("candidate-clippy-preflight-01", 101, [
        "cargo", "clippy", "--locked", "-p", "litchi-xlsx", "--all-features",
        "--lib", "--", "-D", "warnings",
    ]),
}
CANDIDATE_REQUIRED_CHECKS = {
    "candidate-proof-tests-03": (16, [
        "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
        "--all-features", "--lib", "source_proof", "--", "--test-threads=2",
    ]),
    "candidate-source-integration-02": (74, [
        "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
        "--all-features", "--test", "source_backed_cell_values", "--",
        "--test-threads=2",
    ]),
    "candidate-clippy-preflight-02": (None, [
        "cargo", "clippy", "--locked", "-p", "litchi-xlsx", "--all-features",
        "--lib", "--", "-D", "warnings",
    ]),
    "candidate-no-default-preflight-01": (None, [
        "cargo", "check", "--locked", "-p", "litchi-xlsx",
        "--no-default-features",
    ]),
    "candidate-fmt-preflight-01": (None, ["cargo", "fmt", "--all", "--check"]),
    "candidate-boundaries-preflight-01": (None, [
        "python3", "-B", "tools/check_crate_boundaries.py",
    ]),
}
CANDIDATE_CHECK_SNAPSHOT = {
    "candidate-proof-tests-03": "draft-06",
    "candidate-source-integration-02": "draft-06",
    "candidate-clippy-preflight-02": "candidate",
    "candidate-no-default-preflight-01": "candidate",
    "candidate-fmt-preflight-01": "candidate",
    "candidate-boundaries-preflight-01": "candidate",
}


def candidate_index_manifest(index: Path) -> dict[str, str]:
    indexed = parse_index(index)
    names = {name for name in indexed if source_name(name)}
    require(names, "candidate replay source index is empty")
    oids = {indexed[name] for name in names}
    blobs = index_hashes(index, oids)
    return dict(sorted({name: blobs[indexed[name]] for name in names}.items()))


def candidate_attempt_expected_manifest(snapshot: dict[str, str],
                                        public: str, *,
                                        include_lock: bool = True) -> dict[str, str]:
    """Construct the complete check-attempt manifest for a candidate snapshot."""
    result = dict(git_tree_manifest(validate_plan()["revision"]))
    result.update(snapshot)
    public_hashes = read_json(PUBLIC / "source-hashes.json",
                              "public-test-sources/source-hashes.json")
    if public in ("public", "all"):
        for name, row in public_hashes.items():
            result[name] = row["test_sha256"]
    if public == "all":
        exact_inputs = read_json(PUBLIC_EXACT / "inputs.json",
                                 "public-exact-test-sources/inputs.json")
        for name, digest in exact_inputs["source_hashes"].items():
            result[name] = digest
        exact = read_json(PUBLIC_EXACT / "attempt-05/inputs.json",
                          "public-exact-test-sources/attempt-05/inputs.json")
        result[exact["path"]] = exact["after_sha256"]
    if include_lock:
        result["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    return dict(sorted(result.items()))


def validate_baseline_restore(draft05: dict[str, Any],
                              plan: dict[str, Any]) -> dict[str, Any]:
    path = need(BASELINE_RESTORE, "baseline-restore-for-public.json")
    value = read_json(path, "baseline-restore-for-public.json")
    require(isinstance(value, dict) and set(value) == {
        "recorded_utc", "purpose", "candidate_inputs_sha256", "paths",
    }, "baseline restore envelope differs")
    recorded = parse_time(value["recorded_utc"], "baseline-restore.recorded_utc")
    require(value["purpose"] == (
        "Validate added public oracle against baseline; candidate draft05 retained for later restoration."
    ), "baseline restore purpose differs")
    require(value["candidate_inputs_sha256"] == draft05["inputs_sha256"],
            "baseline restore candidate input binding differs")
    base = git_tree_manifest(plan["revision"])
    paths = value["paths"]
    require(isinstance(paths, dict) and set(paths) == set(draft05["snapshot_hashes"]),
            "baseline restore path inventory differs")
    for name, row in paths.items():
        require(isinstance(row, dict) and set(row) == {"before", "after"},
                f"baseline restore row differs: {name}")
        require(row["before"] == draft05["snapshot_hashes"][name],
                f"baseline restore before hash differs: {name}")
        require(row["after"] == base.get(name),
                f"baseline restore after hash differs: {name}")
        if row["after"] is not None:
            check_hash(row["after"], f"baseline restore after hash {name}")
    return {"status": "pass", "sha256": sha(path),
            "recorded_utc": recorded.isoformat(),
            "paths": paths}


def validate_candidate_attempts() -> dict[str, Any]:
    """Replay every retained candidate draft, including the restored-base fork."""
    root = need(CANDIDATE_ATTEMPTS, "candidate-attempts", directory=True)
    plan = validate_plan()
    base = git_tree_manifest(plan["revision"])
    paths = sorted((path for path in root.iterdir() if not path.is_symlink()),
                   key=lambda item: item.name)
    require(paths, "candidate-attempts is empty")
    numbers: list[int] = []
    for path in paths:
        require(path.is_dir() and re.fullmatch(r"draft-[0-9]{2}", path.name),
                f"candidate attempt label is unsafe: {path.name}")
        numbers.append(int(path.name[-2:]))
    require(numbers == list(range(1, len(numbers) + 1)) and numbers[-1] >= 7,
            "candidate attempt sequence is not consecutive from draft-01")
    restore: dict[str, Any] | None = None
    attempts: list[dict[str, Any]] = []
    previous_time: dt.datetime | None = None
    previous_snapshots: dict[str, str] | None = None
    latest_value: dict[str, Any] | None = None

    with tempfile.TemporaryDirectory(prefix=".litchi-0552-candidate-", dir="/home/zhuhe") as folder:
        index = Path(folder) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", plan["revision"]], env=env)
        for number, attempt in zip(numbers, paths):
            if number == 6:
                git(["git", "read-tree", plan["revision"]], env=env)
            inputs_path = need(attempt / "inputs.json", rel(attempt / "inputs.json"))
            patch_path = need(attempt / "candidate.patch", rel(attempt / "candidate.patch"))
            value = read_json(inputs_path, rel(inputs_path))
            base_keys = {
                "schema", "frozen_utc", "scope", "revision", "patch_sha256",
                "changes", "snapshot_hashes", "resources", "adr_manifest_sha256",
            }
            expected_keys = set(base_keys)
            if number >= 2:
                expected_keys |= {"application_base_attempt", "prior_check_receipt_sha256"}
            if number >= 6:
                expected_keys |= {"logical_parent_attempt", "baseline_restore_sha256"}
            require(isinstance(value, dict) and set(value) == expected_keys,
                    f"{attempt.name}/inputs.json inventory differs")
            frozen = parse_time(value["frozen_utc"], f"{attempt.name}.frozen_utc")
            require(previous_time is None or frozen > previous_time,
                    "candidate attempt timestamps are not strictly increasing")
            previous_time = frozen
            require(value["schema"] == "xlsx_0552_candidate_attempt_v1"
                    and value["revision"] == plan["revision"],
                    f"{attempt.name} identity differs")
            require(isinstance(value["scope"], str) and value["scope"].strip(),
                    f"{attempt.name} scope is empty")
            require(value["patch_sha256"] == sha(patch_path),
                    f"{attempt.name} candidate patch hash differs")
            require(value["adr_manifest_sha256"] == sha(ADR),
                    f"{attempt.name} ADR binding differs")
            require(value["resources"] == CANDIDATE_RESOURCES,
                    f"{attempt.name} proof resources differ")
            changes = value["changes"]
            snapshots = value["snapshot_hashes"]
            require(isinstance(changes, dict) and changes,
                    f"{attempt.name} changes are empty")
            require(isinstance(snapshots, dict) and snapshots,
                    f"{attempt.name} snapshots are empty")
            require(set(changes) <= set(snapshots),
                    f"{attempt.name} changes omit snapshots")
            for name, digest in snapshots.items():
                safe_relative(name, f"{attempt.name} snapshot path")
                require(name.startswith("crates/litchi-xlsx/src/"),
                        f"{attempt.name} snapshot path is outside XLSX production")
                check_hash(digest, f"{attempt.name} snapshot hash {name}")
            for name, row in changes.items():
                require(isinstance(row, dict) and set(row) == {"baseline", "candidate"},
                        f"{attempt.name} change row differs: {name}")
                safe_relative(name, f"{attempt.name} change path")
                require(name.startswith("crates/litchi-xlsx/src/"),
                        f"{attempt.name} change path is outside XLSX production")
                if row["baseline"] is not None:
                    check_hash(row["baseline"], f"{attempt.name} baseline hash {name}")
                check_hash(row["candidate"], f"{attempt.name} candidate hash {name}")
                require(row["candidate"] == snapshots[name],
                        f"{attempt.name} candidate snapshot differs: {name}")

            source_root = attempt / "sources"
            source_files = regular_files(source_root, f"{attempt.name}/sources")
            require(set(source_files) == set(snapshots),
                    f"{attempt.name} source snapshot inventory differs")
            for name, digest in snapshots.items():
                require(sha(source_files[name]) == digest,
                        f"{attempt.name} source snapshot hash differs: {name}")
            expected_files = {"inputs.json", "candidate.patch"}
            if number == 1:
                expected_files.add("design.md")
                require(read_text(attempt / "design.md", rel(attempt / "design.md")).strip(),
                        "draft-01 design is empty")
            require(set(item.relative_to(attempt).as_posix()
                        for item in attempt.rglob("*")
                        if item.is_file() and not item.is_symlink()) ==
                    expected_files | {f"sources/{name}" for name in snapshots},
                    f"{attempt.name} recursive file inventory differs")
            top_dirs = {item.name for item in attempt.iterdir() if item.is_dir()}
            require(top_dirs == {"sources"},
                    f"{attempt.name} directory inventory differs")

            before = candidate_index_manifest(index)
            if number == 6:
                require(value["application_base_attempt"] == "restored HEAD baseline"
                        and value["logical_parent_attempt"] == "draft-05",
                        "draft-06 restored-base metadata differs")
                if restore is None:
                    draft05 = attempts[4]
                    restore = validate_baseline_restore(draft05, plan)
                require(value["baseline_restore_sha256"] == restore["sha256"],
                        "draft-06 restore binding differs")
                require(parse_time(restore["recorded_utc"],
                                   "baseline restore recorded_utc") <= frozen,
                        "draft-06 predates baseline restore record")
            elif number >= 2:
                expected_base = f"draft-{number - 1:02d}"
                require(value["application_base_attempt"] == expected_base,
                        f"{attempt.name} application base differs")
                if number >= 6:
                    require(value["logical_parent_attempt"] == expected_base,
                            f"{attempt.name} logical parent differs")
                    require(value["baseline_restore_sha256"] == restore["sha256"],
                            f"{attempt.name} restore binding differs")
            if number == 1:
                require(previous_snapshots is None, "draft-01 has an unexpected parent")
            else:
                require(previous_snapshots is not None,
                        f"{attempt.name} has no previous candidate snapshot")

            expected_prior = CANDIDATE_PRIOR_CHECKS.get(number)
            if number >= 2:
                require(expected_prior is not None, f"{attempt.name} prior check is unknown")
                prior_name, prior_exit, prior_command = expected_prior
                prior_hash = value["prior_check_receipt_sha256"]
                matches = [
                    item for item in (HERE / "check-attempts").glob("*/receipt.json")
                    if sha(item) == prior_hash
                ]
                require(len(matches) == 1
                        and matches[0].parent.name == prior_name,
                        f"{attempt.name} prior receipt is not unique")
                if number == 6:
                    exact = validate_public_exact_tests()
                    expected_prior_manifest = public_exact_manifest(
                        exact["latest_child_sha256"], include_lock=True
                    )
                elif number == 7:
                    expected_prior_manifest = candidate_attempt_expected_manifest(
                        json.loads((CANDIDATE_ATTEMPTS / "draft-06/inputs.json").read_text())[
                            "snapshot_hashes"], "all"
                    )
                else:
                    parent_snapshots = json.loads(
                        (CANDIDATE_ATTEMPTS / f"draft-{number - 1:02d}/inputs.json").read_text()
                    )["snapshot_hashes"]
                    expected_prior_manifest = candidate_attempt_expected_manifest(
                        parent_snapshots, "public"
                    )
                prior = validate_check_attempt(
                    HERE / "check-attempts" / prior_name, expected_prior_manifest
                )
                require(prior["receipt_sha256"] == prior_hash
                        and prior["exit_code"] == prior_exit
                        and prior["command"] == prior_command,
                        f"{attempt.name} prior check binding differs")
                require(parse_time(prior["end_utc"], f"{attempt.name} prior end") <= frozen,
                        f"{attempt.name} prior check postdates the draft")

            if number >= 6 and restore is None:
                restore = validate_baseline_restore(attempts[4], plan)
            # Apply exactly this draft to the isolated index.  The index is
            # reset at draft-06 to model the explicit restored-baseline fork.
            git(["git", "apply", "--cached", "--binary", str(patch_path)], env=env)
            after = candidate_index_manifest(index)
            changed = {name for name in set(before) | set(after)
                       if before.get(name) != after.get(name)}
            require(changed == set(changes),
                    f"{attempt.name} patch changed-path inventory differs")
            for name, row in changes.items():
                require(row["baseline"] == before.get(name)
                        and row["candidate"] == after.get(name),
                        f"{attempt.name} patch replay hash differs: {name}")
            changed_from_plan = {name for name in set(base) | set(after)
                                 if base.get(name) != after.get(name)}
            require(changed_from_plan == set(snapshots),
                    f"{attempt.name} snapshot changed-path inventory differs")
            for name, digest in snapshots.items():
                require(after.get(name) == digest,
                        f"{attempt.name} replay snapshot differs: {name}")
            previous_snapshots = dict(snapshots)
            latest_value = {**value, "inputs_sha256": sha(inputs_path)}
            attempts.append({
                "attempt": attempt.name, "inputs_sha256": sha(inputs_path),
                "frozen_utc": frozen.isoformat(), "snapshot_hashes": dict(snapshots),
                "changes": dict(changes), "patch_sha256": value["patch_sha256"],
            })
    require(latest_value is not None and restore is not None,
            "candidate attempt replay did not reach a restored-base-aware latest draft")
    return {"status": "pass", "count": len(attempts), "latest": attempts[-1],
            "attempts": attempts, "baseline_restore": restore}


def validate_candidate_source_binding(
    attempts_result: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Bind the frozen candidate source to its latest draft and preflight rows."""
    attempts = attempts_result or validate_candidate_attempts()
    latest = attempts["latest"]
    latest_name = latest["attempt"]
    latest_snapshots = latest["snapshot_hashes"]
    candidate_manifest, candidate_manifest_sha = stage_manifest("candidate")
    binding = read_json(CANDIDATE_BINDING, "candidate-source-binding.json")
    require(isinstance(binding, dict) and set(binding) == {
        "schema", "created_utc", "attempt", "inputs_sha256",
        "source_manifest_sha256", "production_source_hashes",
        "public_test_hashes", "preflight_summary_sha256", "status",
    }, "candidate source binding envelope differs")
    binding_time = parse_time(binding["created_utc"],
                              "candidate-source-binding.created_utc")
    require(binding["schema"] == "xlsx_0552_candidate_source_binding_v1"
            and binding["status"] == "frozen for measurement; admission pending",
            "candidate source binding identity differs")
    require(binding["attempt"] == latest_name
            and binding["inputs_sha256"] == latest["inputs_sha256"],
            "candidate source binding draft differs")
    require(binding["source_manifest_sha256"] == candidate_manifest_sha
            and binding_time >= parse_time(latest["frozen_utc"],
                                          "latest candidate frozen_utc"),
            "candidate source binding manifest/timestamp differs")
    production = binding["production_source_hashes"]
    require(production == latest_snapshots,
            "candidate source binding production inventory differs")
    for name, digest in production.items():
        require(candidate_manifest.get(name) == digest,
                f"candidate source binding production hash differs: {name}")

    exact = validate_public_exact_tests()
    public_hashes = read_json(PUBLIC / "source-hashes.json",
                              "public-test-sources/source-hashes.json")
    parent_name = "crates/litchi-xlsx/tests/source_backed_cell_values.rs"
    compact_name = "crates/litchi-xlsx/tests/source_backed_cell_values/compact_source_proof.rs"
    exact_name = "crates/litchi-xlsx/tests/source_backed_cell_values/public_exact_output.rs"
    expected_public = {
        parent_name: exact["source_hashes"][parent_name],
        compact_name: public_hashes[compact_name]["test_sha256"],
        exact_name: exact["latest_child_sha256"],
    }
    require(binding["public_test_hashes"] == expected_public,
            "candidate source binding public inventory differs")
    for name, digest in expected_public.items():
        require(candidate_manifest.get(name) == digest,
                f"candidate source binding public hash differs: {name}")

    summary = read_json(PREFLIGHT_SUMMARY, "preflight-summary-draft07.json")
    require(isinstance(summary, dict) and set(summary) == {
        "schema", "created_utc", "candidate_attempt", "candidate_inputs_sha256",
        "scope", "checks", "performance_measured", "admission", "priority",
    }, "preflight summary envelope differs")
    summary_time = parse_time(summary["created_utc"],
                              "preflight-summary.created_utc")
    require(summary["schema"] == "xlsx_0552_preflight_summary_v1"
            and summary["candidate_attempt"] == latest_name
            and summary["candidate_inputs_sha256"] == latest["inputs_sha256"]
            and summary["scope"] == (
                "Pre-capture checks only. Baseline exact public oracle:2 passed. "
                "Draft06 private16 and integration74 passed. Draft07 replaces a "
                "never-loop with equivalent if; Clippy, no-default, fmt and crate "
                "boundaries pass. Full final quality and representative performance "
                "are still required."
            )
            and summary["performance_measured"] is False
            and summary["admission"] == "pending"
            and summary["priority"] == PREFLIGHT_PRIORITY
            and summary_time <= binding_time,
            "preflight summary identity differs")
    summary_specs = [
        ("baseline-public-exact-05", "baseline", None, [
            "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
            "--all-features", "--test", "source_backed_cell_values",
            "public_exact_output", "--", "--test-threads=2",
        ]),
        ("candidate-proof-tests-03", "draft-06", 16,
         CANDIDATE_REQUIRED_CHECKS["candidate-proof-tests-03"][1]),
        ("candidate-source-integration-02", "draft-06", 74,
         CANDIDATE_REQUIRED_CHECKS["candidate-source-integration-02"][1]),
        ("candidate-clippy-preflight-02", "candidate", None,
         CANDIDATE_REQUIRED_CHECKS["candidate-clippy-preflight-02"][1]),
        ("candidate-no-default-preflight-01", "candidate", None,
         CANDIDATE_REQUIRED_CHECKS["candidate-no-default-preflight-01"][1]),
        ("candidate-fmt-preflight-01", "candidate", None,
         CANDIDATE_REQUIRED_CHECKS["candidate-fmt-preflight-01"][1]),
        ("candidate-boundaries-preflight-01", "candidate", None,
         CANDIDATE_REQUIRED_CHECKS["candidate-boundaries-preflight-01"][1]),
    ]
    rows = summary["checks"]
    require(isinstance(rows, list) and len(rows) == len(summary_specs),
            "preflight summary check count differs")
    draft6_snapshots = next(
        row["snapshot_hashes"] for row in attempts["attempts"]
        if row["attempt"] == "draft-06"
    )
    validated_rows: list[dict[str, Any]] = []
    for row, (name, mode, count, command) in zip(rows, summary_specs):
        require(isinstance(row, dict) and set(row) == {
            "attempt", "receipt_sha256", "command", "source_manifest_sha256",
            "exit_code", "source_stable",
        }, f"preflight summary row differs: {name}")
        expected_path = f"docs/performance/results/change-0552/check-attempts/{name}"
        require(row["attempt"] == expected_path
                and row["command"] == command
                and row["exit_code"] == 0
                and row["source_stable"] is True,
                f"preflight summary binding differs: {name}")
        safe_relative(row["attempt"], f"preflight summary attempt path {name}")
        check_hash(row["receipt_sha256"], f"preflight summary receipt {name}")
        if mode == "baseline":
            expected_manifest = public_exact_manifest(
                exact["latest_child_sha256"], include_lock=True
            )
        elif mode == "draft-06":
            expected_manifest = candidate_attempt_expected_manifest(
                draft6_snapshots, "all"
            )
        else:
            expected_manifest = dict(candidate_manifest)
            expected_manifest["Cargo.lock"] = WORKSPACE_LOCK_SHA256
            expected_manifest = dict(sorted(expected_manifest.items()))
        receipt = validate_check_attempt(HERE / "check-attempts" / name,
                                         expected_manifest)
        require(row["receipt_sha256"] == receipt["receipt_sha256"]
                and row["source_manifest_sha256"] == receipt["source_manifest_sha256"]
                and row["command"] == receipt["command"]
                and receipt["exit_code"] == 0,
                f"preflight receipt custody differs: {name}")
        require(parse_time(receipt["end_utc"], f"preflight receipt end {name}")
                <= summary_time,
                f"preflight receipt postdates summary: {name}")
        if count is not None:
            summaries = test_result_summaries(HERE / "check-attempts" / name / "stdout")
            require(len(summaries) == 1 and summaries[0]["status"] == "ok"
                    and summaries[0]["passed"] == count
                    and summaries[0]["failed"] == 0 and summaries[0]["ignored"] == 0,
                    f"preflight test count differs: {name}")
        validated_rows.append({"name": name, "receipt_sha256": receipt["receipt_sha256"],
                               "source_manifest_sha256": receipt["source_manifest_sha256"],
                               "mode": mode})
    return {
        "status": "pass", "binding_sha256": sha(CANDIDATE_BINDING),
        "source_manifest_sha256": candidate_manifest_sha,
        "attempt": latest_name, "inputs_sha256": latest["inputs_sha256"],
        "preflight_summary_sha256": sha(PREFLIGHT_SUMMARY),
        "checks": validated_rows,
    }


def validate_candidate_correctness() -> dict[str, Any]:
    """Require semantic candidate checks before performance evidence is admitted."""
    attempts = validate_candidate_attempts()
    binding = validate_candidate_source_binding(attempts)
    candidate_manifest, candidate_sha = stage_manifest("candidate")
    expected_manifest = dict(candidate_manifest)
    expected_manifest["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    expected_manifest = dict(sorted(expected_manifest.items()))
    draft6_snapshots = next(
        row["snapshot_hashes"] for row in attempts["attempts"]
        if row["attempt"] == "draft-06"
    )
    rows: list[dict[str, Any]] = []
    for name, (count, command) in CANDIDATE_REQUIRED_CHECKS.items():
        mode = CANDIDATE_CHECK_SNAPSHOT[name]
        check_manifest = expected_manifest if mode == "candidate" else \
            candidate_attempt_expected_manifest(draft6_snapshots, "all")
        receipt = validate_check_attempt(HERE / "check-attempts" / name,
                                         check_manifest)
        require(receipt["exit_code"] == 0 and receipt["command"] == command,
                f"{name} does not prove candidate correctness")
        if count is not None:
            summaries = test_result_summaries(HERE / "check-attempts" / name / "stdout")
            require(len(summaries) == 1 and summaries[0]["status"] == "ok"
                    and summaries[0]["passed"] == count
                    and summaries[0]["failed"] == 0 and summaries[0]["ignored"] == 0,
                    f"{name} test summary differs")
        rows.append({"name": name, "receipt_sha256": receipt["receipt_sha256"],
                     "source_manifest_sha256": receipt["source_manifest_sha256"],
                     "manifest_mode": mode, "count": count})
    return {"status": "pass", "attempts": attempts, "binding": binding,
            "candidate_manifest_sha256": candidate_sha, "checks": rows}


def validate_baseline_correctness() -> dict[str, Any]:
    """Replay the baseline test aggregation and its exact source deltas."""
    public_exact = validate_public_exact_tests()
    path = need(BASELINE_CORRECTNESS, "baseline-correctness.json")
    value = read_json(path, "baseline-correctness.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "status", "recorded_utc", "scope",
        "baseline_manifest_sha256", "source_differences_from_capture_baseline",
        "checks",
    }, "baseline-correctness envelope differs")
    require(value["schema"] == "xlsx_0552_baseline_correctness_v1"
            and value["status"] == "pass"
            and value["scope"] == (
                "Baseline production with baseline-compatible public guards; "
                "does not validate candidate production."
            ), "baseline-correctness identity differs")
    parse_time(value["recorded_utc"], "baseline-correctness.recorded_utc")
    baseline, baseline_sha = stage_manifest("baseline")
    require(value["baseline_manifest_sha256"] == baseline_sha,
            "baseline-correctness baseline hash differs")
    hashes = read_json(PUBLIC / "source-hashes.json",
                       "public-test-sources/source-hashes.json")
    expected_differences = {
        "Cargo.lock": {"capture_baseline": None,
                       "test_source": WORKSPACE_LOCK_SHA256},
    }
    for name, row in hashes.items():
        expected_differences[name] = {
            "capture_baseline": row["baseline_sha256"],
            "test_source": row["test_sha256"],
        }
    require(value["source_differences_from_capture_baseline"] == expected_differences,
            "baseline-correctness source differences differ")
    expected_manifest = dict(public_augmented_manifest())
    expected_manifest["Cargo.lock"] = WORKSPACE_LOCK_SHA256
    expected_manifest = dict(sorted(expected_manifest.items()))
    expected_checks = {
        "baseline-public-01": {
            "command": [
                "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
                "--test", "source_backed_cell_values", "compact_public_guard",
                "--", "--test-threads=2",
            ], "test_groups": 1, "passed": 5,
        },
        "baseline-owner-tests-01": {
            "command": [
                "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
                "--all-features", "--", "--test-threads=2",
            ], "test_groups": 59, "passed": 1311,
        },
    }
    checks = value["checks"]
    require(isinstance(checks, list) and len(checks) == 2,
            "baseline-correctness checks differ")
    seen: set[str] = set()
    rows: list[dict[str, Any]] = []
    for row in checks:
        require(isinstance(row, dict) and set(row) == {
            "attempt", "receipt_sha256", "command",
            "source_manifest_sha256", "test_groups", "passed",
            "failed", "ignored",
        }, "baseline-correctness check row differs")
        attempt = row["attempt"]
        require(isinstance(attempt, str)
                and attempt.startswith("docs/performance/results/change-0552/check-attempts/"),
                "baseline-correctness attempt path differs")
        safe_relative(attempt, "baseline-correctness attempt path")
        attempt_name = Path(attempt).name
        require(attempt_name in expected_checks and attempt_name not in seen,
                "baseline-correctness attempt inventory differs")
        seen.add(attempt_name)
        expected = expected_checks[attempt_name]
        attempt_row = validate_check_attempt(REPO / attempt, expected_manifest)
        require(row["command"] == expected["command"]
                and row["command"] == attempt_row["command"]
                and row["receipt_sha256"] == attempt_row["receipt_sha256"]
                and row["source_manifest_sha256"] == attempt_row["source_manifest_sha256"]
                and row["test_groups"] == expected["test_groups"]
                and row["passed"] == expected["passed"]
                and row["failed"] == 0 and row["ignored"] == 0
                and attempt_row["exit_code"] == 0,
                f"baseline-correctness {attempt_name} differs")
        rows.append({"name": attempt_name, "receipt_sha256": row["receipt_sha256"],
                     "passed": row["passed"]})
    require(seen == set(expected_checks),
            "baseline-correctness check names are incomplete")
    return {"status": "pass", "sha256": sha(path),
            "checks": rows, "source_manifest_sha256": baseline_sha,
            "public_exact": public_exact}


def validate_quality_attempt(path: Path, plan_data: dict[str, Any]
                             ) -> dict[str, Any]:
    label = rel(path)
    require(path.is_dir() and not path.is_symlink(), f"{label} is not a directory")
    inputs_path = path / "inputs.json"
    result_path = path / "result.json"
    inputs = read_json(inputs_path, label + "/inputs.json")
    require(isinstance(inputs, dict) and set(inputs) == {
        "schema", "created_utc", "source_stage", "source_manifest_sha256",
        "workspace_lock_sha256", "quality_plan_sha256",
        "supplemental_inputs_sha256", "scripts", "commands",
    }, f"{label}/inputs.json inventory differs")
    require(inputs["schema"] == "xlsx_0552_quality_inputs_v1"
            and inputs["source_stage"] in ("candidate", "final"),
            f"{label}/inputs.json identity differs")
    created = parse_time(inputs["created_utc"], label + "/inputs.created_utc")
    expected_manifest, stage_sha = quality_expected_manifest(inputs["source_stage"])
    require(inputs["source_manifest_sha256"] == stage_sha
            and inputs["workspace_lock_sha256"] == WORKSPACE_LOCK_SHA256
            and inputs["quality_plan_sha256"] == sha(QUALITY_PLAN)
            and inputs["supplemental_inputs_sha256"] == sha(SUPPLEMENTAL),
            f"{label}/inputs source/lock/plan binding differs")
    scripts = inputs["scripts"]
    require(isinstance(scripts, dict)
            and set(scripts) == {"quality.py", "check_attempt.py", "run.py"},
            f"{label}/inputs script inventory differs")
    require(scripts["quality.py"] == sha(HERE / "quality.py")
            and scripts["check_attempt.py"] == sha(CHECK_ATTEMPT)
            and scripts["run.py"] == sha(RUN),
            f"{label}/inputs script binding differs")
    commands = plan_data["commands"]
    require(inputs["commands"] == commands,
            f"{label}/inputs command matrix differs")
    value = read_json(result_path, label + "/result.json")
    require(isinstance(value, dict) and set(value) == {
        "schema", "status", "source_stage", "source_manifest_sha256",
        "workspace_lock_sha256", "quality_plan_sha256", "inputs_path",
        "inputs_sha256", "completed_utc", "commands",
    }, f"{label}/result.json inventory differs")
    completed = parse_time(value["completed_utc"], label + "/result.completed_utc")
    require(completed >= created,
            f"{label}/result completed before quality attempt creation")
    require(value["schema"] == "xlsx_0552_quality_v1"
            and value["status"] in ("pass", "failed")
            and value["source_stage"] == inputs["source_stage"]
            and value["source_manifest_sha256"] == stage_sha
            and value["workspace_lock_sha256"] == WORKSPACE_LOCK_SHA256
            and value["quality_plan_sha256"] == sha(QUALITY_PLAN)
            and value["inputs_path"] == rel(inputs_path)
            and value["inputs_sha256"] == sha(inputs_path),
            f"{label}/result source/plan binding differs")
    require(isinstance(value["commands"], list),
            f"{label}/result command rows are not a list")
    rows = value["commands"]
    require(len(rows) <= len(commands), f"{label}/result has too many commands")
    expected_names = list(commands)
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and set(row) == {
            "name", "command", "attempt", "receipt_sha256",
            "exit_code", "runner_exit_code", "source_stable",
        }, f"{label}/result command row {index} differs")
        require(row["name"] == expected_names[index]
                and row["command"] == commands[row["name"]]
                and isinstance(row["attempt"], str)
                and row["attempt"].startswith("docs/performance/results/change-0552/check-attempts/"),
                f"{label}/result command row {index} binding differs")
        safe_relative(row["attempt"], f"{label}/result attempt path {index}")
        check_hash(row["receipt_sha256"], f"{label}/result receipt hash {index}")
        attempt = REPO / row["attempt"]
        attempt_row = validate_check_attempt(attempt, expected_manifest)
        require(row["receipt_sha256"] == attempt_row["receipt_sha256"]
                and row["exit_code"] == attempt_row["exit_code"]
                and row["source_stable"] is True
                and isinstance(row["runner_exit_code"], int)
                and not isinstance(row["runner_exit_code"], bool)
                and row["runner_exit_code"] == row["exit_code"],
                f"{label}/result attempt row {index} differs")
        require(row["command"] == attempt_row["command"],
                f"{label}/result command differs from attempt {index}")
        attempt_start, attempt_end = attempt_row["times"]
        require(created <= attempt_start <= attempt_end <= completed,
                f"{label}/result attempt interval is outside quality attempt")
    if value["status"] == "pass":
        require(len(rows) == len(commands)
                and all(row["exit_code"] == 0
                        and row["runner_exit_code"] == 0
                        and row["source_stable"] for row in rows),
                f"{label} claims pass with a failed command")
    actual = {
        item.name for item in path.iterdir()
        if item.is_file() and not item.is_symlink()
    }
    require(actual == {"inputs.json", "result.json"},
            f"{label} raw attempt inventory differs")
    return {
        "path": label, "result_path": rel(result_path),
        "result_sha256": sha(result_path), "status": value["status"],
        "source_stage": value["source_stage"], "source_manifest_sha256": stage_sha,
        "completed_utc": value["completed_utc"], "commands": rows,
    }


def validate_quality() -> dict[str, Any]:
    plan_data = validate_quality_plan()
    attempts_root = HERE / "quality-attempts"
    if not attempts_root.exists():
        raise IncompleteError("quality-attempts directory is missing")
    require(attempts_root.is_dir() and not attempts_root.is_symlink(),
            "quality-attempts is not a directory")
    attempts: list[dict[str, Any]] = []
    for path in sorted(attempts_root.iterdir(), key=lambda item: item.name):
        require(re.fullmatch(r"[A-Za-z0-9_-]+", path.name) is not None,
                f"quality attempt label is unsafe: {path.name}")
        attempts.append(validate_quality_attempt(path, plan_data))
    passing = [item for item in attempts
               if item["status"] == "pass" and item["source_stage"] == "final"]
    if not passing:
        raise IncompleteError("no passing final quality attempt is retained")
    canonical = need(HERE / "quality.json", "quality.json")
    canonical_bytes = read_bytes(canonical, "quality.json")
    selected = [item for item in passing
                if canonical_bytes == read_bytes(REPO / item["result_path"],
                                                 item["result_path"])]
    require(len(selected) == 1,
            "quality.json is not the exact copy of one passing attempt")
    value = read_json(canonical, "quality.json")
    require(value.get("schema") == "xlsx_0552_quality_v1"
            and value.get("status") == "pass",
            "quality.json does not contain a passing quality result")
    return {
        "status": "pass", "canonical_sha256": sha(canonical),
        "selected_attempt": selected[0]["result_path"],
        "attempts": attempts, "commands": plan_data["commands"],
    }


def current_source_manifest() -> dict[str, str]:
    """Hash the live source set using the same inclusion rule as run.freeze."""
    try:
        tracked = subprocess.check_output(
            ["git", "ls-files", "-z", "crates", "tools/perf-baseline",
             "Cargo.toml", "Cargo.lock", ".cargo", "rust-toolchain.toml"],
            cwd=REPO,
        ).split(b"\0")
        untracked = subprocess.check_output(
            ["git", "ls-files", "--others", "--exclude-standard", "-z", "--",
             "crates", "tools/perf-baseline"],
            cwd=REPO,
        ).split(b"\0")
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError(f"cannot enumerate live source: {error}") from error
    names = {item.decode("utf-8") for item in tracked if item}
    names |= {
        item.decode("utf-8") for item in untracked
        if item and item.endswith(b".rs")
    }
    result: dict[str, str] = {}
    for name in sorted(names):
        path = REPO / name
        require(not path.is_symlink(), f"live source path is a symlink: {name}")
        if path.is_file():
            require(source_name(name), f"live source path is out of scope: {name}")
            result[name] = sha(path)
    require(result, "live source manifest is empty")
    return result


def public_augmented_manifest() -> dict[str, str]:
    baseline, _ = stage_manifest("baseline")
    hashes = read_json(PUBLIC / "source-hashes.json",
                       "public-test-sources/source-hashes.json")
    result = dict(baseline)
    for name, row in hashes.items():
        require(isinstance(row, dict) and is_hash(row.get("test_sha256")),
                f"public test hash row is malformed: {name}")
        result[name] = row["test_sha256"]
    return dict(sorted(result.items()))


def public_oracle_augmented_manifest() -> dict[str, str]:
    """Return the restored-baseline manifest retaining the validated exact oracle."""
    exact = validate_public_exact_tests()
    return public_exact_manifest(exact["latest_child_sha256"])


def validate_final_source(disposition: str, candidate_manifest: dict[str, str],
                          candidate_sha: str) -> dict[str, Any]:
    """Bind the final source stage to acceptance or baseline restoration.

    Public guard tests are retained after a rejection, so a restored final
    source may equal the frozen baseline plus the exact public-test bundle.
    The frozen baseline itself remains an accepted restoration form for a
    coordinator that removes those tests before sealing.
    """
    require(disposition in ("accepted", "rejected"),
            "final source disposition is malformed")
    manifest, manifest_sha = stage_manifest("final")
    final_patch = need(FINAL / "source.patch", "final/source.patch")
    baseline, baseline_sha = stage_manifest("baseline")
    if disposition == "accepted":
        require(manifest == candidate_manifest and manifest_sha == candidate_sha,
                "accepted final source does not equal candidate")
        validate_source_patch("final", manifest)
        expected = "candidate"
    else:
        allowed = (baseline, public_augmented_manifest(),
                   public_oracle_augmented_manifest())
        require(manifest in allowed,
                "rejected final source is not the restored baseline")
        if manifest == baseline:
            require(final_patch.read_bytes() == b"",
                    "frozen-baseline final source patch is not empty")
        elif manifest == public_augmented_manifest():
            require(final_patch.stat().st_size > 0,
                    "public-test-retaining final source patch is empty")
            validate_source_patch("final", manifest)
        else:
            require(final_patch.stat().st_size > 0,
                    "public-oracle-retaining final source patch is empty")
            validate_source_patch("final", manifest)
        expected = (
            "baseline" if manifest == baseline else
            "baseline-plus-public-tests" if manifest == public_augmented_manifest()
            else "baseline-plus-public-oracle"
        )
        require(manifest_sha in (baseline_sha, sha(FINAL / "source-manifest.json")),
                "rejected final source hash is unstable")
    live = current_source_manifest()
    require(live == manifest,
            "live checkout does not match final source disposition")
    return {
        "status": "pass", "stage": "final",
        "manifest_sha256": manifest_sha, "manifest_entries": len(manifest),
        "expected": expected, "live_manifest_sha256": sha(
            FINAL / "source-manifest.json"
        ),
    }


def decision_path() -> Path:
    paths = [HERE / name for name in DECISION_NAMES if (HERE / name).exists()]
    if not paths:
        raise IncompleteError("final disposition document is missing")
    require(len(paths) == 1,
            "multiple final disposition documents are retained")
    path = paths[0]
    require(path.is_file() and not path.is_symlink(),
            f"{rel(path)} is not a regular file")
    return path


def decision_gate(value: dict[str, Any], names: tuple[str, ...],
                  expected: bool, label: str) -> None:
    for name in names:
        if name in value:
            require(isinstance(value[name], bool) and value[name] is expected,
                    f"{label} differs from independent evidence")


def validate_disposition() -> dict[str, Any]:
    """Recompute adoption from every independent gate and bind final source."""
    metrics = validate_metrics_analysis()
    guards = validate_guard_cap_analysis()
    profiles = validate_profiles(
        pilot_expected=profile_pilot_from_metrics(metrics, guards)
    )
    quality = validate_quality()
    path = decision_path()
    value = read_json(path, rel(path))
    require(isinstance(value, dict), f"{rel(path)} is not an object")
    disposition = value.get("disposition")
    require(disposition in ("accepted", "rejected"),
            f"{rel(path)} disposition differs")
    observed = value.get("observed_utc", value.get("completed_utc",
                                                   value.get("utc")))
    parse_time(observed, f"{rel(path)} timestamp")
    scope = value.get("scope")
    require(isinstance(scope, str) and scope.strip(),
            f"{rel(path)} scope is missing")
    main_gate = bool(metrics["main_gates"]["all_frozen_main_gates_pass"])
    guard_gate = bool(guards["guard_admission_passed"])
    cap_gate = bool(guards["cap_admission_passed"])
    profile_gate = bool(profiles["gate_passed"])
    quality_gate = quality["status"] == "pass"
    adoption = main_gate and guard_gate and cap_gate and profile_gate and quality_gate
    require(isinstance(value.get("adoption_allowed"), bool)
            and value["adoption_allowed"] is adoption,
            f"{rel(path)} adoption_allowed differs from independent gates")
    require(disposition == ("accepted" if adoption else "rejected"),
            f"{rel(path)} disposition differs from independent gates")
    decision_gate(value, ("native_primary_gate", "main_gate"), main_gate, "main gate")
    decision_gate(value, ("guard_gate",), guard_gate, "guard gate")
    decision_gate(value, ("cap_gate",), cap_gate, "cap gate")
    if profiles["required"]:
        decision_gate(value, ("profile_gate",), profile_gate, "profile gate")
    elif "profile_gate" in value:
        require(isinstance(value["profile_gate"], bool),
                f"{rel(path)} skipped profile gate is malformed")
    decision_gate(value, ("quality_gate",), quality_gate, "quality gate")
    for name, expected_hash in (
        ("metrics_analysis_sha256", metrics["report"]["sha256"]),
        ("main_analysis_sha256", metrics["report"]["sha256"]),
        ("guard_analysis_sha256", guards["report"]["sha256"]),
        ("quality_sha256", quality["canonical_sha256"]),
        ("quality_summary_sha256", quality["canonical_sha256"]),
    ):
        if name in value:
            require(value[name] == expected_hash,
                    f"{rel(path)} {name} differs")
    if "profile_analysis_sha256" in value and profiles.get("decision_sha256"):
        require(value["profile_analysis_sha256"] == profiles["decision_sha256"],
                f"{rel(path)} profile analysis hash differs")
    final_stage = validate_final_source(
        disposition, json.loads((CANDIDATE / "source-manifest.json").read_text()),
        sha(CANDIDATE / "source-manifest.json"),
    )
    final_manifest_sha = final_stage["manifest_sha256"]
    if "final_source" in value:
        require(value["final_source"] == "final",
                f"{rel(path)} final_source must select final stage")
    if "final_source_manifest_sha256" in value:
        require(value["final_source_manifest_sha256"] == final_manifest_sha,
                f"{rel(path)} final source manifest hash differs")
    require(quality["attempts"], "quality attempts are missing")
    canonical = read_json(HERE / "quality.json", "quality.json")
    require(canonical["source_stage"] == "final"
            and canonical["source_manifest_sha256"] == final_manifest_sha,
            "canonical quality is not bound to final source")
    return {
        "status": "pass", "decision": rel(path), "decision_sha256": sha(path),
        "disposition": disposition, "adoption_allowed": adoption,
        "gates": {
            "main": main_gate, "guard": guard_gate, "cap": cap_gate,
            "profile": profile_gate, "quality": quality_gate,
        }, "final_source": final_stage,
        "metrics": metrics["report"], "guards": guards["report"],
        "quality": {"sha256": quality["canonical_sha256"],
                    "attempt": quality["selected_attempt"]},
    }


def receipt_interval(path: Path) -> tuple[dt.datetime, dt.datetime, str]:
    value = read_json(path, rel(path))
    start, end = validate_receipt_common(value, rel(path))
    return start, end, rel(path)


def repeat_receipt_intervals(stage: str, prefix: str, repeat: int
                             ) -> list[tuple[dt.datetime, dt.datetime, str]]:
    paths = sorted((HERE / stage).glob(prefix + "*.receipt.json"))
    result: list[tuple[dt.datetime, dt.datetime, str]] = []
    for path in paths:
        if f"-r{repeat}-" in path.name:
            result.append(receipt_interval(path))
    require(result, f"{stage}/{prefix} repeat {repeat} receipts are missing")
    return result


def check_abba_order(prefix: str) -> dict[str, Any]:
    groups = [
        ("baseline-r1", repeat_receipt_intervals("baseline", prefix, 1)),
        ("candidate-r1", repeat_receipt_intervals("candidate", prefix, 1)),
        ("candidate-r2", repeat_receipt_intervals("candidate", prefix, 2)),
        ("baseline-r2", repeat_receipt_intervals("baseline", prefix, 2)),
    ]
    bounds: list[dict[str, Any]] = []
    for name, rows in groups:
        first = min(row[0] for row in rows)
        last = max(row[1] for row in rows)
        bounds.append({"group": name, "first_start_utc": first.isoformat(),
                       "last_end_utc": last.isoformat(), "count": len(rows)})
    require(all(
        dt.datetime.fromisoformat(left["last_end_utc"])
        <= dt.datetime.fromisoformat(right["first_start_utc"])
        for left, right in zip(bounds, bounds[1:])
    ), f"{prefix} ABBA groups overlap")
    return {"prefix": prefix, "order": [item[0] for item in groups],
            "groups": bounds, "passed": True}


def validate_receipt_inventory() -> dict[str, Any]:
    """Check exact stage receipt names and serial ABBA timing."""
    validate_main_bundle()
    for stage in STAGES:
        include_repeat_two = True
        validate_guard_cap_stage(
            stage, sha(BASELINE / "source-manifest.json"),
            sha(CANDIDATE / "source-manifest.json"), include_repeat_two,
        )
    profile = validate_profiles()
    expected_by_stage: dict[str, set[str]] = {}
    for stage in STAGES:
        names = {
            path.name for path in (HERE / stage).glob("*.receipt.json")
        }
        expected = {
            f"build-{kind}.receipt.json"
            for kind in ("normal", "alloc", "guard-normal", "guard-alloc", "cap")
        }
        expected |= {
            f"{job['name']}.receipt.json"
            for lane in ("preflight", "native", "alloc")
            for job in expected_main_jobs(lane, stage)
        }
        expected |= {
            f"{job['name']}.receipt.json"
            for lane in ("normal", "alloc")
            for job in expected_guard_jobs(lane)
        }
        expected |= {
            f"{job['name']}.receipt.json" for job in expected_cap_jobs()
        }
        if profile["required"]:
            expected |= {f"{job['name']}.receipt.json"
                         for job in expected_profile_jobs()}
        require(names == expected,
                f"{stage} complete receipt inventory differs")
        validate_stage_raw_inventory(stage, True, profile["required"])
        expected_by_stage[stage] = names
    attempts = validate_all_check_attempts()

    def group(stage: str, names: list[str], label: str) -> tuple[dt.datetime, dt.datetime, str]:
        rows = [receipt_interval(HERE / stage / f"{name}.receipt.json")
                for name in names]
        ordered_rows = sorted(rows, key=lambda item: (item[0], item[1], item[2]))
        require([item[2].rsplit("/", 1)[-1][:-len(".receipt.json")]
                 for item in ordered_rows] == names,
                f"{label} serial order differs")
        return min(item[0] for item in rows), max(item[1] for item in rows), label

    def stage_order(stage: str, include_repeat_two: bool) -> list[tuple[dt.datetime, dt.datetime, str]]:
        groups: list[tuple[dt.datetime, dt.datetime, str]] = []

        def add(label: str, names: list[str]) -> None:
            groups.append(group(stage, names, label))

        def jobs(lane: str, repeat: int) -> list[str]:
            return [job["name"] for job in expected_main_jobs(lane, stage)
                    if job["repeat"] == repeat]

        def guard(lane: str, repeat: int) -> list[str]:
            return [job["name"] for job in expected_guard_jobs(lane)
                    if job["repeat"] == repeat]

        def caps(repeat: int) -> list[str]:
            return [job["name"] for job in expected_cap_jobs()
                    if job["repeat"] == repeat]

        add(f"{stage}/build-normal", ["build-normal"])
        add(f"{stage}/preflight-r1", jobs("preflight", 1))
        add(f"{stage}/native-r1", jobs("native", 1))
        if include_repeat_two:
            add(f"{stage}/native-r2", jobs("native", 2))
        add(f"{stage}/build-guard-normal", ["build-guard-normal"])
        add(f"{stage}/guard-native-r1", guard("normal", 1))
        if include_repeat_two:
            add(f"{stage}/guard-native-r2", guard("normal", 2))
        add(f"{stage}/build-cap", ["build-cap"])
        add(f"{stage}/cap-r1", caps(1))
        if include_repeat_two:
            add(f"{stage}/cap-r2", caps(2))
        add(f"{stage}/build-alloc", ["build-alloc"])
        add(f"{stage}/alloc-r1", jobs("alloc", 1))
        if include_repeat_two:
            add(f"{stage}/alloc-r2", jobs("alloc", 2))
        add(f"{stage}/build-guard-alloc", ["build-guard-alloc"])
        add(f"{stage}/guard-alloc-r1", guard("alloc", 1))
        if include_repeat_two:
            add(f"{stage}/guard-alloc-r2", guard("alloc", 2))
        return groups

    baseline_order = stage_order("baseline", False)
    candidate_order = stage_order("candidate", True)
    baseline_r2 = [
        group("baseline", names, label)
        for names, label in (
            ([job["name"] for job in expected_main_jobs("native", "baseline")
              if job["repeat"] == 2], "baseline/native-r2"),
            ([job["name"] for job in expected_guard_jobs("normal")
              if job["repeat"] == 2], "baseline/guard-native-r2"),
            ([job["name"] for job in expected_cap_jobs()
              if job["repeat"] == 2], "baseline/cap-r2"),
            ([job["name"] for job in expected_main_jobs("alloc", "baseline")
              if job["repeat"] == 2], "baseline/alloc-r2"),
            ([job["name"] for job in expected_guard_jobs("alloc")
              if job["repeat"] == 2], "baseline/guard-alloc-r2"),
        )
    ]
    # The baseline stage's first repeat is captured before the candidate;
    # candidate owns both of its repeats; baseline repeat two is replayed last
    # under the candidate execution manifest.  The per-lane ABBA checks below
    # remain the authoritative cross-stage check.
    serial_groups = baseline_order + candidate_order + baseline_r2
    require(all(left[1] <= right[0]
                for left, right in zip(serial_groups, serial_groups[1:])),
            "stage capture serial order differs")
    if profile["required"]:
        profile_names = [job["name"] for job in expected_profile_jobs()]
        profile_groups = [
            group(stage, profile_names, f"{stage}/profiles")
            for stage in STAGES
        ]
        require(serial_groups[-1][1] <= profile_groups[0][0]
                and profile_groups[0][1] <= profile_groups[1][0],
                "profile capture serial order differs")
    prefixes = ("native-", "alloc-", "guard-native-", "guard-alloc-", "cap-")
    abba = [check_abba_order(prefix) for prefix in prefixes]
    preflight_base = repeat_receipt_intervals("baseline", "preflight-", 1)
    preflight_cand = repeat_receipt_intervals("candidate", "preflight-", 1)
    require(max(row[1] for row in preflight_base)
            <= min(row[0] for row in preflight_cand),
            "preflight baseline/candidate order differs")
    return {
        "status": "pass", "stages": {
            stage: {"receipt_count": len(names),
                    "receipts": sorted(names)}
            for stage, names in expected_by_stage.items()
        }, "check_attempts": attempts, "abba": abba,
        "profile": profile,
    }


def validate_receipt_timeline() -> dict[str, Any]:
    """Require every retained child interval to be serial across the bundle."""
    capture_paths = [
        path for stage in STAGES
        for path in (HERE / stage).glob("*.receipt.json")
    ]
    attempt_paths = [
        path / "receipt.json"
        for path in (HERE / "check-attempts").iterdir()
        if path.is_dir() and not path.is_symlink()
    ] if (HERE / "check-attempts").is_dir() else []
    intervals = [receipt_interval(path) for path in capture_paths]
    for path in attempt_paths:
        value = read_json(path, rel(path))
        start, end = interval(value, rel(path))
        intervals.append((start, end, rel(path)))
    require(intervals, "receipt timeline is empty")
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "retained receipt intervals overlap")
    return {
        "status": "pass", "count": len(ordered),
        "first_start_utc": ordered[0][0].isoformat(),
        "last_end_utc": ordered[-1][1].isoformat(),
        "receipts": [item[2] for item in ordered],
    }


def cleanup_custody(value: dict[str, Any], stage: str, kind: str,
                    expected: str, path: Path) -> None:
    hashes = value.get("binary_sha256_by_kind")
    require(isinstance(hashes, dict),
            "cleanup binary_sha256_by_kind map is missing")
    aliases = (f"{stage}/{kind}", f"{stage}:{kind}", kind,
               path.name, str(path), path.as_posix())
    observed = next((hashes.get(alias) for alias in aliases
                     if hashes.get(alias) is not None), None)
    require(observed == expected and is_hash(observed),
            f"cleanup binary custody differs: {stage}/{kind}")


def validate_cleanup() -> dict[str, Any]:
    path = need(HERE / "cleanup.json", "cleanup.json")
    value = read_json(path, "cleanup.json")
    require(isinstance(value, dict), "cleanup.json is not an object")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("removed") == [str(TARGET)]
            and value.get("owned_paths_absent") is True
            and value.get("accessible_process_references") == [],
            "cleanup target custody differs")
    stamp = value.get("observed_utc", value.get("completed_utc",
                                                 value.get("utc")))
    parse_time(stamp, "cleanup timestamp")
    require(not os.path.lexists(TARGET),
            "cleanup claims target absent but owned path remains")
    if "target" in value:
        require(value["target"] == str(TARGET), "cleanup target differs")
    if "scope" in value:
        require(isinstance(value["scope"], str) and value["scope"].strip(),
                "cleanup scope is empty")
    if "python_cache_absent" in value:
        require(value["python_cache_absent"] is True
                and not list(HERE.rglob("__pycache__")),
                "Python cache remains after cleanup")
    input_hashes = value.get("input_sha256")
    if input_hashes is not None:
        require(isinstance(input_hashes, dict),
                "cleanup input hash map is malformed")
        for name, digest in input_hashes.items():
            safe_relative(name, "cleanup input path")
            check_hash(digest, f"cleanup input {name}")
            require(sha(HERE / name) == digest,
                    f"cleanup input hash differs: {name}")
    checked = []
    for stage in STAGES:
        manifest, manifest_sha = stage_manifest(stage)
        for kind in ("normal", "alloc", "guard-normal", "guard-alloc", "cap"):
            descriptor = read_json(HERE / stage / descriptor_name(kind),
                                   f"{stage}/{descriptor_name(kind)}")
            expected = descriptor.get("sha256") if isinstance(descriptor, dict) else None
            binary_path = Path(descriptor.get("path", "")) if isinstance(descriptor, dict) else Path("")
            require(is_hash(expected)
                    and binary_path == SCRATCH_ROOT / stage / retained_binary_name(kind)
                    and descriptor.get("source_manifest_sha256") == manifest_sha,
                    f"{stage}/{kind} descriptor differs during cleanup")
            cleanup_custody(value, stage, kind, expected, binary_path)
            checked.append(f"{stage}/{kind}")
    return {"status": "pass", "sha256": sha(path),
            "owned_paths_absent": True, "binary_custody": checked}


def validate_seal() -> dict[str, Any]:
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2, "SHA256SUMS line differs")
        check_hash(fields[0], "SHA256SUMS digest")
        safe_relative(fields[1], "SHA256SUMS path")
        require(fields[1] != "SHA256SUMS" and fields[1] not in expected,
                "SHA256SUMS inventory is unsafe or duplicated")
        expected[fields[1]] = fields[0]
    actual = {
        rel(item): sha(item) for item in HERE.rglob("*")
        if item.is_file() and not item.is_symlink() and item != SEAL
    }
    require(expected == actual and not any(item.is_symlink() for item in HERE.rglob("*")),
            "SHA256SUMS inventory differs")
    require(not list(HERE.rglob("__pycache__")),
            "Python bytecode cache is retained")
    return {"status": "pass", "sha256": sha(SEAL), "entries": len(expected)}


def validate_candidate_preliminary() -> dict[str, Any]:
    """Validate a complete candidate stage before retained baseline r2."""
    plan = validate_plan()
    validate_supplemental_inputs()
    validate_analysis_inputs()
    validate_workspace_lock()
    validate_public_tests()
    candidate = validate_source("candidate")
    candidate_correctness = validate_candidate_correctness()
    manifest_sha = candidate["manifest_sha256"]
    builds, build_intervals = validate_stage_builds("candidate", manifest_sha)
    main_rows, main_intervals = validate_main_stage(
        "candidate", plan, manifest_sha, manifest_sha, True
    )
    guard_rows, guard_intervals = validate_guard_cap_stage(
        "candidate", manifest_sha, manifest_sha, True
    )
    intervals = build_intervals + main_intervals + guard_intervals
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "candidate build/capture receipts overlap")
    validate_stage_raw_inventory("candidate", True, False)
    return {
        "status": "preliminary", "complete": False,
        "candidate_complete": True, "scope": "candidate-only preliminary",
        "source": candidate, "candidate_correctness": candidate_correctness,
        "builds": builds,
        "main_rows": compact_main_rows(main_rows), "guard_cap": guard_rows,
        "intervals": len(intervals),
        "baseline": {"status": "pending retained baseline r2"},
    }


def validate_inputs_component() -> dict[str, Any]:
    return {
        "status": "pass", "frozen": validate_frozen_inputs(),
        "supplemental": validate_supplemental_inputs(),
        "analysis": validate_analysis_inputs(),
        "workspace_lock": validate_workspace_lock(),
        "quality_plan": validate_quality_plan(),
    }


def validate_source_component() -> dict[str, Any]:
    baseline = validate_source("baseline")
    candidate = validate_source("candidate")
    require(baseline["manifest_sha256"] != candidate["manifest_sha256"],
            "candidate source manifest is unchanged from baseline")
    return {"status": "pass", "baseline": baseline, "candidate": candidate,
            "live_manifest_entries": len(current_source_manifest())}


def validate_all() -> dict[str, Any]:
    """Run the terminal pre-cleanup and post-cleanup custody checks."""
    disposition = validate_disposition()
    inventory = validate_receipt_inventory()
    timeline = validate_receipt_timeline()
    cleanup = validate_cleanup()
    seal = validate_seal()
    return {
        "status": "pass", "disposition": disposition,
        "receipt_inventory": inventory, "receipt_timeline": timeline,
        "cleanup": cleanup, "seal": seal,
    }


def component(name: str) -> dict[str, Any]:
    """Dispatch a read-only preliminary or terminal verifier component."""
    if name == "plan":
        plan = validate_plan()
        return {"status": "pass", "plan": plan,
                "supplemental": validate_supplemental_inputs(),
                "analysis": validate_analysis_inputs()}
    if name in {"inputs", "input"}:
        return validate_inputs_component()
    if name == "host":
        return validate_host()
    if name == "adr":
        return validate_adr()
    if name in {"workspace-lock", "lock"}:
        return validate_workspace_lock()
    if name in {"public-tests", "public"}:
        return validate_public_tests()
    if name in {"public-exact", "exact-public"}:
        return validate_public_exact_tests()
    if name in {"candidate-attempts", "attempts"}:
        return validate_candidate_attempts()
    if name in {"candidate-binding", "source-binding"}:
        return validate_candidate_source_binding()
    if name in {"candidate-correctness", "correctness"}:
        return validate_candidate_correctness()
    if name == "source":
        return validate_source_component()
    if name in {"baseline", "preliminary"}:
        return validate_main_preliminary()
    if name == "candidate":
        return validate_candidate_preliminary()
    if name in {"build", "builds"}:
        plan = validate_plan()
        validate_supplemental_inputs()
        validate_analysis_inputs()
        baseline = validate_source("baseline")
        candidate = validate_source("candidate")
        builds: dict[str, Any] = {}
        intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
        for stage, source in (("baseline", baseline), ("candidate", candidate)):
            builds[stage], times = validate_stage_builds(
                stage, source["manifest_sha256"]
            )
            intervals.extend(times)
        ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
        require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
                "build receipts overlap")
        return {"status": "pass", "plan_sha256": plan["sha256"],
                "builds": builds, "intervals": len(intervals)}
    if name in {"captures", "main"}:
        result = validate_main_bundle()
        if name == "main":
            result["metrics"] = validate_metrics_analysis()
        return result
    if name in {"metrics", "analysis"}:
        return validate_metrics_analysis()
    if name in {"guards", "guard", "cap", "guard-cap", "guard-analysis"}:
        return validate_guard_cap_analysis()
    if name in {"profiles", "profile"}:
        return validate_profiles()
    if name == "quality":
        return validate_quality()
    if name in {"disposition", "decision"}:
        return validate_disposition()
    if name in {"inventory", "receipts"}:
        return validate_receipt_inventory()
    if name in {"serial", "timeline"}:
        return validate_receipt_timeline()
    if name == "precleanup":
        require(os.path.lexists(TARGET),
                "precleanup requires the owned target to remain present")
        disposition = validate_disposition()
        inventory = validate_receipt_inventory()
        timeline = validate_receipt_timeline()
        return {"status": "pass", "disposition": disposition,
                "receipt_inventory": inventory, "receipt_timeline": timeline}
    if name == "cleanup":
        return validate_cleanup()
    if name == "seal":
        return validate_seal()
    if name == "all":
        return validate_all()
    raise VerificationError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    try:
        result = component(selected)
    except IncompleteError as error:
        return {
            "schema": "litchi.xlsx.verification.0552.v1",
            "status": "incomplete", "scope": selected, "error": str(error),
        }
    except (VerificationError, OSError, subprocess.CalledProcessError,
            KeyError, TypeError, AttributeError, IndexError, ValueError) as error:
        return {
            "schema": "litchi.xlsx.verification.0552.v1",
            "status": "fail", "scope": selected, "error": str(error),
        }
    status = result.get("status")
    # A baseline/candidate-only component is useful during assembly, but it
    # does not certify the complete campaign.  Keep that distinction in the
    # machine-readable envelope so missing later evidence cannot look like a
    # terminal pass.
    envelope_status = "incomplete" if status in ("preliminary", "pending") else "pass"
    return {
        "schema": "litchi.xlsx.verification.0552.v1",
        "status": envelope_status, "scope": selected, "result": result,
    }


def verify(sealed: bool = False) -> dict[str, Any]:
    """Programmatic entry point; unsealed verification stops before cleanup."""
    return component("all" if sealed else "precleanup")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=(
        "all", "precleanup", "preliminary", "plan", "inputs", "host", "adr",
        "workspace-lock", "public-tests", "public-exact", "source", "baseline", "candidate",
        "candidate-attempts", "candidate-binding", "candidate-correctness",
        "build", "builds", "captures", "main", "metrics", "analysis",
        "guards", "guard", "cap", "guard-cap", "guard-analysis",
        "profiles", "profile", "quality", "disposition", "decision",
        "inventory", "receipts", "serial", "timeline", "cleanup", "seal",
    ), default="all")
    parser.add_argument("--strict", action="store_true",
                        help="return failure for incomplete evidence")
    parser.add_argument("--output", type=Path,
                        help="write the verification result outside this bundle")
    args = parser.parse_args(argv)
    if args.output is not None:
        try:
            if args.output.resolve().is_relative_to(HERE.resolve()):
                raise SystemExit("verification output must be outside the evidence bundle")
        except AttributeError:
            if str(args.output.resolve()).startswith(str(HERE.resolve()) + os.sep):
                raise SystemExit("verification output must be outside the evidence bundle")
    result = run_bundle(args.component)
    encoded = json.dumps(result, indent=2, sort_keys=True, allow_nan=False) + "\n"
    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded, encoding="utf-8")
    else:
        print(encoded, end="")
    if result["status"] == "fail":
        return 1
    if result["status"] == "incomplete" and args.strict:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
