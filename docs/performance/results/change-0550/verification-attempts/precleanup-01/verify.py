"""Fail-closed, read-only custody verifier for the 0550 XLSX attribution run.

0550 is a source-preserving diagnostic batch.  It has one measured baseline
stage and does not admit a candidate or make a speedup claim.  This verifier
reconstructs the source tree in a private Git index, checks the exact frozen
driver matrix and every serial receipt, replays the retained profile and
metrics analyzers in temporary paths, and verifies cleanup and the recursive
seal.  It never builds, captures, edits the checkout, or writes inside this
evidence directory.

Use ``python3 -B verify.py --precleanup --strict`` before removing the owned
target.  Use ``python3 -B verify.py --strict`` after the target and all
temporary binaries have been removed and ``SHA256SUMS`` has been written.
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
import sys
import tempfile
from typing import Any

sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
CAPTURE = HERE / "capture.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
HOST = HERE / "host.json"
BASE = HERE / "baseline"
TARGET = Path("/home/zhuhe/litchi-goal-0550-target")
SCRATCH_ROOT = TARGET / "retained"
PRIOR = HERE.parent / "change-0549"
PRIOR_FINAL_MANIFEST = PRIOR / "final" / "source-manifest.json"
PRIOR_SEAL = PRIOR / "SHA256SUMS"
QUALITY_CONTINUITY = HERE / "quality-continuity.json"
PRIOR_QUALITY = PRIOR / "quality-summary.json"
TOOL_CONTINUITY = HERE / "tool-source-continuity.json"
PROFILE_ANALYZER = HERE / "analyze_profiles.py"
METRICS_ANALYZER = HERE / "analyze_metrics.py"
ANALYSIS_PHASE = HERE / "analysis_phase.py"
ANALYSIS_INPUTS = HERE / "analysis-inputs.json"
ANALYSIS_RUNS = HERE / "analysis-runs"
PROFILE_REPORT = HERE / "profile-analysis.json"
METRICS_REPORT = HERE / "metrics-analysis.json"
DOCUMENTATION_MANIFEST = HERE / "documentation-manifest.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"

REVISION = "090b15b64ae52da2f8bf765cbb745ef76122792e"
SCOPE = "Current source-backed XLSX MultiSourceEdit commit attribution; no optimization or speedup claim"
PRIORITY = "OLE2/OOXML first; ODF deferred until that optimization goal completes; iWork excluded"
CASES = (
    "xlsx_source_backed_cell_values_one_edit_save",
    "xlsx_source_backed_cell_values_one_percent_edit_save",
)
SHAPES = ("medium", "dense-sparse", "noncompact", "vendor-extension")
PROFILE_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
PROFILE_OWNER = "litchi_xlsx::cell_values::source::MultiSourceEdit::commit"
PROFILE_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
PROFILE_LIFECYCLE = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
EXPECTED_ENVIRONMENT = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
RECEIPT_KEYS = {
    "command", "start_utc", "end_utc", "seconds", "exit_code",
    "execution_stage", "execution_manifest_sha256", "binary_sha256",
    "source_manifest_sha256", "script_sha256", "plan_sha256", "environment",
    "artifacts",
}
QUALITY_NAMES = ("fmt", "harness-fmt", "boundaries", "claims")
DOCUMENTATION_NAMES = (
    "docs/performance/BASELINE.md",
    "docs/performance/HOTSPOTS.md",
    "docs/performance/REPORT.md",
    "docs/performance/ADR_COMPLIANCE.md",
    "docs/performance/GOAL_AUDIT.md",
    "docs/performance/changes/0550-xlsx-source-commit-attribution.md",
    "docs/performance/results/change-0550/README.md",
    "docs/performance/results/change-0550/protocol.md",
    "docs/performance/results/change-0550/adr-compliance.md",
    "docs/performance/results/change-0550/source-review.md",
    "docs/performance/results/change-0550/scope-review.md",
    "docs/performance/results/change-0550/profile-review.md",
    "docs/performance/results/change-0550/results-review.md",
)
ANALYSIS_SPECS = (
    ("profiles", PROFILE_ANALYZER, PROFILE_REPORT),
    ("metrics", METRICS_ANALYZER, METRICS_REPORT),
)
ANALYSIS_RECEIPT_KEYS = {
    "command", "start_utc", "end_utc", "exit_code", "script_sha256",
    "plan_sha256", "output", "output_sha256", "artifacts", "inputs_sha256",
}
ANALYSIS_REQUIRED_INPUTS = frozenset({
    "docs/performance/results/change-0550/analyze_metrics.py",
    "docs/performance/results/change-0550/analyze_profiles.py",
    "docs/performance/results/change-0550/analysis_phase.py",
    "docs/performance/results/change-0550/run.py",
    "docs/performance/results/change-0550/capture.py",
    "docs/performance/results/change-0550/plan.json",
    "docs/performance/results/change-0550/baseline/source-manifest.json",
    "docs/performance/results/change-0550/capture-amendment-inputs.json",
    "tools/validate_perf_corpus_binding.py",
    "docs/performance/results/change-0546/integration/analyze.py",
    "docs/performance/results/change-0521/analyze.py",
    "docs/performance/results/change-0521/analyze_profiles.py",
    "docs/performance/results/change-0519/analyze_profiles.py",
    "docs/performance/results/change-0519/compare_profile_lanes.py",
})
ALLOWED_PERFORMANCE_CLAIMS = {
    None,
    "none",
    "diagnostic-only",
    "none: descriptive baseline attribution only; no admission gate or speedup claim",
}
FAILED_MANIFEST = HERE / "failed-attempts" / "manifest.json"
CAPTURE_AMENDMENT = HERE / "capture-amendment-inputs.json"
CAPTURE_AMENDED = HERE / "capture_amended.py"
RESUME = HERE / "resume.py"
DEBUGGER_CLEANUP = HERE / "failed-attempts" / "debugger-temp-cleanup.json"
FAILED_PROFILE = "profile-r1-medium-c0"
TIME_FORMAT = None  # Kept explicit: native capture uses ``time -v`` on stderr.


class VerificationError(ValueError):
    """Malformed, contradictory, or out-of-scope retained evidence."""


class IncompleteError(VerificationError):
    """Evidence required by a selected component has not arrived yet."""


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


def valid_digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and "." not in path.parts and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise VerificationError(f"{label} timestamp is malformed") from error
    require(result.tzinfo is not None, f"{label} has no timezone")
    return result


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and seconds >= 0 and end > start,
            f"{label} interval is invalid")
    wall = (end - start).total_seconds()
    tolerance = max(0.25, wall * 0.02 + 0.05)
    require(abs(float(seconds) - wall) <= tolerance,
            f"{label} seconds does not match its UTC interval")
    return start, end


def git(args: list[str], *, env: dict[str, str] | None = None,
        input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, env=env, input=input_data,
                                       stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise VerificationError(f"Git command failed ({' '.join(args)}): "
                                f"{detail.decode(errors='replace')[-2000:]}") from error


def source_name(name: str) -> bool:
    return name in SOURCE_EXACT or any(name.startswith(prefix) for prefix in SOURCE_PREFIXES)


def source_manifest(path: Path, label: str | None = None) -> dict[str, str]:
    label = label or rel(path)
    value = read_json(path, label)
    require(isinstance(value, dict) and value, f"{label} is an empty source manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label} source path")
        require(source_name(name) and valid_digest(digest),
                f"{label} source entry is invalid: {name}")
        require(name not in result, f"{label} repeats {name}")
        result[name] = digest
    return result


def git_tree_manifest(revision: str) -> dict[str, str]:
    raw = git(["git", "ls-tree", "-r", "-z", revision, "--", "crates",
               "tools/perf-baseline", "Cargo.toml", "Cargo.lock", ".cargo",
               "rust-toolchain.toml"])
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
        require(len(fields) == 3 and fields[1] == b"blob" and source_name(name),
                f"Git source tree entry is out of scope: {name}")
        entries.append((name, fields[2].decode()))
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
    return result


def current_source_manifest() -> dict[str, str]:
    tracked = git(["git", "ls-files", "-z", "crates", "tools/perf-baseline",
                   "Cargo.toml", "Cargo.lock", ".cargo", "rust-toolchain.toml"])
    untracked = git(["git", "ls-files", "--others", "--exclude-standard", "-z", "--",
                     "crates", "tools/perf-baseline"])
    names = {item.decode("utf-8") for item in tracked.split(b"\0") if item}
    names.update(item.decode("utf-8") for item in untracked.split(b"\0")
                 if item and item.endswith(b".rs"))
    names = {name for name in names if source_name(name)}
    result: dict[str, str] = {}
    for name in sorted(names):
        path = REPO / name
        require(path.is_file() and not path.is_symlink(),
                f"current source path is missing: {name}")
        result[name] = sha(path)
    return result


def parse_index(index: Path) -> dict[str, str]:
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    raw = git(["git", "ls-files", "-s", "-z"], env=env)
    result: dict[str, str] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, encoded = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3 and fields[0] in (b"100644", b"100755"),
                "private source index entry is malformed")
        name = encoded.decode("utf-8")
        require(name not in result, f"private source index repeats {name}")
        result[name] = fields[1].decode()
    return result


def index_hashes(index: Path, oids: set[str]) -> dict[str, str]:
    if not oids:
        return {}
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    response = git(["git", "cat-file", "--batch"], env=env,
                   input_data=("\n".join(sorted(oids)) + "\n").encode())
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


def validate_frozen_inputs() -> dict[str, Any]:
    value = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(value, dict)
            and set(value) == {"run.py", "capture.py", "plan.json", "adr-manifest.json"},
            "frozen-inputs.json inventory differs")
    for name, digest in value.items():
        require(valid_digest(digest), f"frozen input digest is malformed: {name}")
        path = HERE / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"frozen input hash differs: {name}")
    return {"status": "pass", "sha256": sha(FROZEN),
            "files": {name: value[name] for name in sorted(value)}}


def validate_plan() -> dict[str, Any]:
    frozen = validate_frozen_inputs()
    value = read_json(PLAN, "plan.json")
    expected_keys = {
        "revision", "created_utc", "previous_turn", "priority", "scope",
        "hypothesis", "cpu", "owned_paths", "cases", "shapes", "native",
        "alloc", "profile", "drift_review_percent", "order", "limitations",
    }
    require(isinstance(value, dict) and set(value) == expected_keys,
            "0550 plan envelope differs")
    parse_time(value.get("created_utc"), "plan.created_utc")
    require(value.get("revision") == REVISION and re.fullmatch(r"[0-9a-f]{40}", REVISION),
            "plan revision differs")
    try:
        subprocess.run(["git", "cat-file", "-e", REVISION + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError("plan revision is not a Git commit") from error
    require(value.get("previous_turn") ==
            "progress: committed 0549 rejection with strict post-commit replay and cleanup",
            "plan previous-turn continuity differs")
    require(value.get("priority") == PRIORITY and value.get("scope") == SCOPE,
            "plan priority or scope differs")
    require(value.get("hypothesis") == (
        "Fresh attribution determines whether full rewritten validation and reduced readback "
        "are substantial repeated work relative to rewrite and semantic merge. Inclusive shares "
        "are not removable fractions."
    ), "plan hypothesis differs")
    require(value.get("cpu") == 2 and value.get("owned_paths") == [str(TARGET)]
            and value.get("cases") == list(CASES)
            and value.get("shapes") == list(SHAPES), "plan workload envelope differs")
    for section_name, expected in (
        ("native", {"repeats": 2, "warmup": 3, "samples": 30}),
        ("alloc", {"repeats": 2, "warmup": 3, "samples": 30}),
    ):
        section = value.get(section_name)
        require(section == expected, f"plan {section_name} matrix differs")
    require(value.get("profile") == {
        "repeats": 2, "warmup": 0, "samples": 1, "case": PROFILE_CASE,
        "owner": PROFILE_OWNER, "parent": PROFILE_PARENT, "lifecycle": PROFILE_LIFECYCLE,
    }, "plan profile owner or matrix differs")
    require(value.get("drift_review_percent") == 5
            and value.get("order") == [
                "build normal", "preflight", "native r1", "profiles r1",
                "profiles r2", "native r2", "build allocator", "alloc r1",
                "alloc r2", "metadata checks",
            ], "plan order or review threshold differs")
    expected_limitations = [
        "30 native samples descriptive only; no registered latency claim",
        "Native commit phase includes edit staging; exact owner profile excludes it",
        "Profiles simulated Ir only; instrumented elapsed excluded",
        "No new counters for copied/reduced bytes or fallback route; source/profile inference must be labeled",
        "Single SourceEdit owner and managed execution not measured in this batch",
        "No new hardware, physical-cold, provider, concurrency, native Office or fuzz claim",
    ]
    require(value.get("limitations") == expected_limitations,
            "plan limitations differ")
    return {"status": "pass", "sha256": sha(PLAN), "frozen": frozen}


def validate_adr() -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    previous = read_json(PRIOR / "adr-manifest.json", "change-0549/adr-manifest.json")
    require(isinstance(value, dict) and isinstance(previous, dict)
            and set(value) == {"checked_utc", "files", "status"}
            and value.get("files") == previous.get("files")
            and value.get("status") == "all previously read ADRs unchanged",
            "ADR manifest continuity differs")
    parse_time(value.get("checked_utc"), "adr-manifest.checked_utc")
    files = value["files"]
    require(isinstance(files, dict) and files, "ADR file map is empty")
    for name, digest in files.items():
        safe_relative(name, "ADR path")
        path = REPO / name
        require(name.startswith("docs/adr/") and valid_digest(digest)
                and path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"ADR hash differs: {name}")
    return {"status": "pass", "entries": len(files), "sha256": sha(ADR)}


def validate_quality_continuity() -> dict[str, Any]:
    """Bind the unchanged-source claim to the accepted 0549 quality record."""
    value = read_json(QUALITY_CONTINUITY, "quality-continuity.json")
    expected_scope = (
        "Exact unchanged source inventory. Prior 15checks/4382executions are OLE "
        "owner tests and workspace/harness checks, not a fresh XLSX suite. This "
        "diagnostic runs fresh build/case oracles and four metadata checks; no new "
        "broad test/fuzz/native-Office claim."
    )
    require(isinstance(value, dict)
            and set(value) == {
                "source_manifest_sha256", "prior_manifest", "prior_manifest_sha256",
                "prior_quality", "prior_quality_sha256", "prior_checks",
                "prior_executed_tests", "scope",
            }, "quality continuity envelope differs")
    require(value.get("source_manifest_sha256") == sha(BASE / "source-manifest.json")
            and value.get("prior_manifest") ==
            "docs/performance/results/change-0549/final/source-manifest.json"
            and value.get("prior_manifest_sha256") == sha(PRIOR_FINAL_MANIFEST)
            and value.get("prior_quality") ==
            "docs/performance/results/change-0549/quality-summary.json"
            and value.get("prior_quality_sha256") == sha(PRIOR_QUALITY)
            and value.get("prior_checks") == 15
            and value.get("prior_executed_tests") == 4382
            and value.get("scope") == expected_scope,
            "quality continuity binding differs")
    return {"status": "pass", "sha256": sha(QUALITY_CONTINUITY),
            "prior_manifest_sha256": value["prior_manifest_sha256"],
            "prior_quality_sha256": value["prior_quality_sha256"]}


def validate_tool_source_continuity() -> dict[str, Any]:
    """Verify helper/checker bytes against the frozen Git revision."""
    value = read_json(TOOL_CONTINUITY, "tool-source-continuity.json")
    expected_files = {
        "tools/check_crate_boundaries.py",
        "tools/check_perf_claims.py",
        "tools/validate_perf_corpus_binding.py",
    }
    require(isinstance(value, dict)
            and set(value) == {"revision", "files", "scope"}
            and value.get("revision") == REVISION
            and value.get("scope") ==
            "Current checker/helper bytes equal frozen baseline Git revision; observed after terminal metadata checks.",
            "tool source continuity envelope differs")
    files = value.get("files")
    require(isinstance(files, dict) and set(files) == expected_files,
            "tool source continuity inventory differs")
    for name in sorted(expected_files):
        safe_relative(name, "tool source path")
        path = REPO / name
        require(path.is_file() and not path.is_symlink() and valid_digest(files[name])
                and sha(path) == files[name],
                f"current tool source hash differs: {name}")
        frozen = git(["git", "show", f"{REVISION}:{name}"])
        require(hashlib.sha256(frozen).hexdigest() == files[name],
                f"frozen tool source hash differs: {name}")
    return {"status": "pass", "sha256": sha(TOOL_CONTINUITY),
            "entries": len(files), "revision": REVISION}


def validate_host() -> dict[str, Any]:
    value = read_json(HOST, "host.json")
    commands = [
        ["rustc", "-vV"], ["cargo", "-V"], ["valgrind", "--version"],
        ["perf", "--version"], ["uname", "-a"], ["lscpu"],
        ["rustup", "toolchain", "list"], ["cargo", "fuzz", "--version"],
    ]
    require(isinstance(value, dict)
            and set(value) == {"observations", "affinity", "environment"}
            and isinstance(value.get("observations"), list)
            and len(value["observations"]) == len(commands),
            "host observation envelope differs")
    for row, command in zip(value["observations"], commands):
        require(isinstance(row, dict)
                and set(row) == {"command", "exit_code", "stdout", "stderr"}
                and row.get("command") == command
                and isinstance(row.get("exit_code"), int)
                and isinstance(row.get("stdout"), str)
                and isinstance(row.get("stderr"), str),
                "host observation row differs")
    require(value["observations"][-1]["exit_code"] == 101,
            "cargo-fuzz availability observation differs")
    require(all(row["exit_code"] == 0 for row in value["observations"][:-1]),
            "host tool observation failed unexpectedly")
    affinity = value.get("affinity")
    require(isinstance(affinity, list) and affinity == list(range(32)),
            "host affinity observation differs")
    environment = value.get("environment")
    require(isinstance(environment, dict)
            and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            "host instrumentation environment differs")
    return {"status": "pass", "sha256": sha(HOST), "affinity": affinity}


def validate_source(plan: dict[str, Any]) -> dict[str, Any]:
    manifest_path = need(BASE / "source-manifest.json", "baseline/source-manifest.json")
    patch_path = need(BASE / "source.patch", "baseline/source.patch")
    manifest = source_manifest(manifest_path, "baseline/source-manifest.json")
    require(patch_path.read_bytes() == b"", "baseline source.patch is not empty")
    tree = git_tree_manifest(REVISION)
    require(manifest == tree, "baseline source manifest differs from frozen Git revision")
    current = current_source_manifest()
    require(current == manifest, "current source manifest differs from baseline")
    prior = source_manifest(PRIOR_FINAL_MANIFEST, "change-0549/final/source-manifest.json")
    require(manifest == prior,
            "0550 baseline source does not preserve the 0549 final source")
    prior_seal = read_text(PRIOR_SEAL, "change-0549/SHA256SUMS")
    binding = next((line.split("  ", 1)[0] for line in prior_seal.splitlines()
                    if line.endswith("  final/source-manifest.json")), None)
    require(binding == sha(PRIOR_FINAL_MANIFEST),
            "0549 seal does not bind its final source manifest")
    return {
        "status": "pass", "manifest_sha256": sha(manifest_path),
        "manifest_entries": len(manifest), "current_matches": True,
        "prior_0549_final_matches": True,
    }


def validate_private_source_replay(plan: dict[str, Any]) -> dict[str, Any]:
    """Apply the retained empty patch through a private index as a custody check."""
    patch = need(BASE / "source.patch", "baseline/source.patch")
    with tempfile.TemporaryDirectory(prefix=".litchi-0550-source-", dir="/home/zhuhe") as folder:
        index = Path(folder) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", REVISION], env=env)
        if patch.stat().st_size:
            git(["git", "apply", "--cached", "--binary", str(patch)], env=env)
        indexed_all = parse_index(index)
        indexed = {name: oid for name, oid in indexed_all.items() if source_name(name)}
        changed = set(git(["git", "diff", "--cached", "--name-only", REVISION],
                           env=env).decode().splitlines())
        require(not changed and set(indexed) == set(source_manifest(BASE / "source-manifest.json")),
                "private source replay differs from baseline")
        hashes = index_hashes(index, set(indexed.values()))
        expected = source_manifest(BASE / "source-manifest.json")
        for name, oid in indexed.items():
            require(hashes.get(oid) == expected.get(name),
                    f"private source replay blob differs: {name}")
    return {"status": "pass", "changed_files": [], "index_replayed": True}


def original_profile_command(job: dict[str, Any], kind: str = "normal") -> list[str]:
    """The command retained for the failed, pre-amendment profile attempt."""
    return [
        "taskset", "-c", "2", "/usr/bin/time", "-v", "valgrind",
        "--tool=callgrind", "--collect-atstart=no",
        "--toggle-collect=" + PROFILE_OWNER,
        "--zero-before=" + PROFILE_OWNER,
        "--dump-after=" + PROFILE_OWNER,
        "--callgrind-out-file=" + str(BASE / f"{job['name']}.callgrind"),
        str(SCRATCH_ROOT / "baseline" / kind), "--case", job["case"],
        "--xlsx-cell-crud-shape", job["shape"], "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]), "--json", str(BASE / f"{job['name']}.json"),
        "--corpus-manifest", str(BASE / f"{job['name']}.catalog.json"),
    ]


def validate_failed_attempt() -> dict[str, Any]:
    """Retain and independently bind the first failed profile invocation.

    The first frozen driver attempted to start Valgrind's default gdbserver
    and failed before the profiled child ran.  The amended driver is admitted
    only because this exact failure, its raw files, and the one-option source
    amendment are retained.  A verifier must never turn that attempt into a
    successful measurement or silently discard it.
    """
    manifest_path = need(FAILED_MANIFEST, "failed-attempts/manifest.json")
    folder = need(HERE / "failed-attempts", "failed-attempts", directory=True)
    attempt_folder = need(folder / FAILED_PROFILE, f"failed-attempts/{FAILED_PROFILE}", directory=True)
    manifest = read_json(manifest_path, rel(manifest_path))
    expected_names = {
        f"baseline/{FAILED_PROFILE}.callgrind",
        f"baseline/{FAILED_PROFILE}.host.json",
        f"baseline/{FAILED_PROFILE}.receipt.json",
        f"baseline/{FAILED_PROFILE}.stderr",
        f"baseline/{FAILED_PROFILE}.stdout",
    }
    require(isinstance(manifest, dict) and set(manifest) == expected_names,
            "failed-attempt manifest inventory differs")
    require({path.name for path in folder.iterdir()} == {
        "manifest.json", "debugger-temp-cleanup.json", FAILED_PROFILE,
    }, "failed-attempts inventory contains an unretained file")
    expected_retained = {
        name: f"failed-attempts/{FAILED_PROFILE}/{Path(name).name}"
        for name in expected_names
    }
    for source_name_value, row in manifest.items():
        safe_relative(row.get("retained") if isinstance(row, dict) else None,
                      "failed-attempt retained path")
        require(isinstance(row, dict)
                and set(row) == {"retained", "sha256"}
                and row.get("retained") == expected_retained[source_name_value]
                and valid_digest(row.get("sha256")),
                f"failed-attempt manifest row differs: {source_name_value}")
        retained = HERE / row["retained"]
        require(retained.is_file() and not retained.is_symlink()
                and sha(retained) == row["sha256"],
                f"failed-attempt custody differs: {source_name_value}")
    require({path.name for path in attempt_folder.iterdir()} == {
        f"{FAILED_PROFILE}.callgrind", f"{FAILED_PROFILE}.host.json",
        f"{FAILED_PROFILE}.receipt.json", f"{FAILED_PROFILE}.stderr",
        f"{FAILED_PROFILE}.stdout",
    }, "failed profile raw inventory differs")

    receipt_path = attempt_folder / f"{FAILED_PROFILE}.receipt.json"
    receipt = read_json(receipt_path, rel(receipt_path))
    job = {"name": FAILED_PROFILE, "case": PROFILE_CASE, "shape": "medium",
           "warmup": 0, "samples": 1}
    require(isinstance(receipt, dict) and set(receipt) == RECEIPT_KEYS
            and receipt.get("command") == original_profile_command(job)
            and receipt.get("exit_code") == 1,
            "failed profile receipt does not preserve the original failed command")
    source_sha = sha(BASE / "source-manifest.json")
    require(receipt.get("execution_stage") == "baseline"
            and receipt.get("execution_manifest_sha256") == source_sha
            and receipt.get("source_manifest_sha256") == source_sha
            and receipt.get("script_sha256") == sha(RUN)
            and receipt.get("plan_sha256") == sha(PLAN)
            and receipt.get("binary_sha256") ==
            read_json(BASE / "binary-normal.json", "baseline/binary-normal.json").get("sha256"),
            "failed profile source/driver binding differs")
    environment = receipt.get("environment")
    require(isinstance(environment, dict) and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            "failed profile instrumentation environment differs")
    interval(receipt, rel(receipt_path))
    artifacts = receipt.get("artifacts")
    expected_artifacts_set = {
        f"{FAILED_PROFILE}.callgrind", f"{FAILED_PROFILE}.host.json",
        f"{FAILED_PROFILE}.stderr", f"{FAILED_PROFILE}.stdout",
    }
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts_set,
            "failed profile receipt artifacts differ")
    for name, digest in artifacts.items():
        path = attempt_folder / name
        require(valid_digest(digest) and sha(path) == digest,
                f"failed profile artifact hash differs: {name}")
        if name.endswith(".host.json"):
            validate_host_sidecar(path, f"failed-attempts/{FAILED_PROFILE}/{name}")
    require((attempt_folder / f"{FAILED_PROFILE}.callgrind").stat().st_size == 0
            and (attempt_folder / f"{FAILED_PROFILE}.stdout").stat().st_size == 0,
            "failed profile retained measurement artifacts are not empty")
    failed_stderr = read_text(attempt_folder / f"{FAILED_PROFILE}.stderr",
                              f"failed-attempts/{FAILED_PROFILE}/{FAILED_PROFILE}.stderr")
    require("error writing" in failed_stderr and "shared mem" in failed_stderr
            and "Command exited with non-zero status 1" in failed_stderr,
            "failed profile stderr does not preserve the pre-measurement failure")

    amendment = read_json(CAPTURE_AMENDMENT, "capture-amendment-inputs.json")
    require(isinstance(amendment, dict)
            and set(amendment) == {"utc", "reason", "inputs", "failed_artifacts"},
            "capture amendment envelope differs")
    amendment_time = parse_time(amendment.get("utc"), "capture-amendment-inputs.utc")
    failed_start = parse_time(receipt.get("start_utc"), "failed profile start")
    failed_end = parse_time(receipt.get("end_utc"), "failed profile end")
    require(failed_end < amendment_time and failed_start < failed_end,
            "capture amendment does not follow the failed attempt")

    # The failed invocation belongs exactly in the frozen serial gap between
    # native repeat one and the amended profile repeat one.  It is retained as
    # a failed attempt and must not be allowed to overlap or replace a
    # successful receipt.
    native_last = read_json(BASE / "native-r1-vendor-extension-c1.receipt.json",
                            "baseline/native-r1-vendor-extension-c1.receipt.json")
    profile_first = read_json(BASE / "profile-r1-medium-c0.receipt.json",
                              "baseline/profile-r1-medium-c0.receipt.json")
    native_end = parse_time(native_last.get("end_utc"),
                            "baseline/native-r1-vendor-extension-c1 end")
    profile_start = parse_time(profile_first.get("start_utc"),
                               "baseline/profile-r1-medium-c0 start")
    require(native_end <= failed_start < failed_end <= profile_start,
            "failed profile attempt is outside the serial capture gap")
    reason = amendment.get("reason")
    require(isinstance(reason, str) and all(fragment in reason for fragment in (
        "gdbserver", "owned target", "No production", "native timing"
    )) and ("vgdb" in reason or "--vgdb=no" in reason),
            "capture amendment reason is incomplete")
    inputs = amendment.get("inputs")
    expected_inputs = {
        "plan.json": PLAN,
        "run.py": RUN,
        "capture.py": CAPTURE,
        "capture_amended.py": CAPTURE_AMENDED,
        "resume.py": RESUME,
        "failed-attempts/manifest.json": FAILED_MANIFEST,
    }
    require(isinstance(inputs, dict) and set(inputs) == set(expected_inputs),
            "capture amendment input inventory differs")
    for name, path in expected_inputs.items():
        require(inputs.get(name) == sha(path), f"capture amendment input hash differs: {name}")
    failed_artifacts = amendment.get("failed_artifacts")
    require(failed_artifacts == manifest,
            "capture amendment failed-artifact map differs")

    amended_text = read_text(CAPTURE_AMENDED, "capture_amended.py")
    original_text = read_text(CAPTURE, "capture.py")
    old = "command += ['valgrind', '--tool=callgrind', '--collect-atstart=no',"
    new = "command += ['valgrind', '--vgdb=no', '--vgdb-prefix=' + str(R.TARGET / 'tmp/vgdb'), '--tool=callgrind', '--collect-atstart=no',"
    require(original_text.count(old) == 1 and amended_text.count(new) == 1
            and amended_text == original_text.replace(old, new),
            "capture amendment changes more than the approved Valgrind options")
    resume_text = read_text(RESUME, "resume.py")
    require("from capture_amended import capture" in resume_text,
            "resume.py does not use the approved amended capture driver")
    try:
        resume_tree = ast.parse(resume_text, filename=str(RESUME))
    except SyntaxError as error:
        raise VerificationError("resume.py is not valid Python") from error
    resume_calls: list[tuple[str, int]] = []
    for node in ast.walk(resume_tree):
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Name):
            continue
        if node.func.id != "capture" or len(node.args) != 2:
            continue
        if isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str) \
                and isinstance(node.args[1], ast.Constant) and isinstance(node.args[1].value, int):
            resume_calls.append((node.args[0].value, node.args[1].value))
    require(resume_calls == [
        ("profile", 1), ("profile", 2), ("native", 2), ("alloc", 1), ("alloc", 2)
    ], "resume.py does not retain the exact serial continuation")

    debugger = read_json(DEBUGGER_CLEANUP, "failed-attempts/debugger-temp-cleanup.json")
    require(isinstance(debugger, dict)
            and set(debugger) == {"failed_pid", "removed", "scope"}
            and isinstance(debugger.get("failed_pid"), int)
            and debugger["failed_pid"] > 0
            and debugger.get("scope") == "Only exact failed-child debugger paths; failed capture is terminal.",
            "debugger temporary cleanup record differs")
    removed = debugger.get("removed")
    require(isinstance(removed, list) and removed, "debugger temporary cleanup is empty")
    for row in removed:
        require(isinstance(row, dict) and set(row) == {"path", "uid", "mode", "bytes"}
                and isinstance(row.get("path"), str)
                and row["path"].startswith("/tmp/vgdb-pipe-shared-mem-vgdb-")
                and isinstance(row.get("uid"), int) and isinstance(row.get("mode"), int)
                and isinstance(row.get("bytes"), int) and row["bytes"] == 0,
                "debugger temporary cleanup row differs")
    return {"status": "pass", "failed_profile": FAILED_PROFILE,
            "receipt_sha256": sha(receipt_path), "amendment_sha256": sha(CAPTURE_AMENDMENT),
            "manifest_sha256": sha(FAILED_MANIFEST)}


def validate_host_sidecar(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict)
            and set(value) == {"observed_utc", "compiler_processes", "scope"}
            and value.get("scope") == "Accessible compiler processes; no host quiescence guarantee"
            and isinstance(value.get("compiler_processes"), list),
            f"{label} host sidecar differs")
    parse_time(value.get("observed_utc"), f"{label}.observed_utc")
    for process in value["compiler_processes"]:
        require(isinstance(process, dict) and isinstance(process.get("pid"), int)
                and process["pid"] > 0 and process.get("comm") in {"cargo", "rustc"}
                and isinstance(process.get("cwd"), str) and process["cwd"],
                f"{label} compiler process row differs")


def expected_artifacts(name: str, lane: str | None = None) -> set[str]:
    if name.startswith(("build-", "check-")) or name == "symbols":
        return {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"}
    require(lane in {"preflight", "native", "alloc", "profile"},
            f"cannot infer lane for {name}")
    result = {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr",
              f"{name}.json", f"{name}.catalog.json"}
    if lane == "profile":
        result.add(f"{name}.callgrind")
        numbered = sorted(
            int(path.name.rsplit(".", 1)[1])
            for path in BASE.glob(f"{name}.callgrind.*")
            if path.is_file() and not path.is_symlink()
            and path.name.rsplit(".", 1)[1].isdigit()
        )
        require(numbered == list(range(1, len(numbered) + 1)) and numbered,
                f"{name} numbered Callgrind inventory differs")
        result.update(f"{name}.callgrind.{number}" for number in numbered)
    return result


def validate_artifacts(folder: Path, value: dict[str, Any], expected: set[str], label: str,
                       *, allowed_extra: set[str] | None = None) -> None:
    allowed_extra = set() if allowed_extra is None else set(allowed_extra)
    require(all(isinstance(name, str) and Path(name).name == name
                for name in allowed_extra),
            f"{label} allowed extra artifact path is unsafe")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} artifact inventory differs")
    stem = label.rsplit("/", 1)[-1].removesuffix(".receipt.json")
    actual = {path.name for path in folder.iterdir()
              if path.name.startswith(stem + ".") and path.name != stem + ".receipt.json"}
    require(actual - allowed_extra == expected, f"{label} raw artifact files differ")
    for filename in allowed_extra & actual:
        path = folder / filename
        require(path.is_file() and not path.is_symlink(),
                f"{label} extra artifact is not a regular file: {filename}")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename
                and valid_digest(digest), f"{label} artifact entry is malformed")
        path = folder / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label} artifact custody differs: {filename}")
        if filename.endswith(".host.json"):
            validate_host_sidecar(path, f"{label}/{filename}")


def validate_receipt(path: Path, name: str, command: list[str], binary: str | None,
                     expected: set[str], *, allow_failure: bool = False) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS
            and value.get("command") == command,
            f"{rel(path)} receipt command or schema differs")
    require(isinstance(value.get("exit_code"), int)
            and (allow_failure or value["exit_code"] == 0),
            f"{rel(path)} did not exit successfully")
    source_sha = sha(BASE / "source-manifest.json")
    require(value.get("execution_stage") == "baseline"
            and value.get("execution_manifest_sha256") == source_sha
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("script_sha256") == sha(RUN)
            and value.get("source_manifest_sha256") == source_sha
            and value.get("binary_sha256") == binary,
            f"{rel(path)} source/driver/binary binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            f"{rel(path)} instrumentation environment differs")
    times = interval(value, rel(path))
    extras = ({f"{name}.inclusive.txt", f"{name}.self.txt"}
              if name.startswith("profile-") else set())
    validate_artifacts(path.parent, value, expected, rel(path), allowed_extra=extras)
    return value, times


def build_command(kind: str) -> list[str]:
    executable = "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else "")
    command = ["env", "TMPDIR=" + str(TARGET / "tmp"), "CARGO_BUILD_JOBS=2",
               "CARGO_INCREMENTAL=0", "cargo", "build", "--release", "--locked",
               "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", executable,
               "--target-dir", str(TARGET)]
    if kind == "alloc":
        command += ["--features", "allocator-metrics"]
    return command


def binary_identity(kind: str) -> dict[str, Any]:
    path = need(BASE / f"binary-{kind}.json", f"baseline/binary-{kind}.json")
    value = read_json(path, rel(path))
    expected_path = SCRATCH_ROOT / "baseline" / kind
    require(isinstance(value, dict)
            and set(value) == {"path", "sha256", "bytes", "build_receipt_sha256",
                               "source_manifest_sha256"}
            and value.get("path") == str(expected_path)
            and valid_digest(value.get("sha256"))
            and isinstance(value.get("bytes"), int) and value["bytes"] > 0
            and value.get("source_manifest_sha256") == sha(BASE / "source-manifest.json")
            and value.get("build_receipt_sha256") == sha(BASE / f"build-{kind}.receipt.json"),
            f"{rel(path)} identity differs")
    binary = Path(value["path"])
    if binary.exists():
        require(binary.is_file() and not binary.is_symlink()
                and TARGET.is_dir() and not TARGET.is_symlink()
                and SCRATCH_ROOT.is_dir() and not SCRATCH_ROOT.is_symlink()
                and binary.resolve() == expected_path
                and sha(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                f"{rel(path)} binary custody differs")
    else:
        cleanup = read_json(CLEANUP, "cleanup.json")
        hashes = cleanup.get("binary_sha256_by_kind") if isinstance(cleanup, dict) else None
        require(isinstance(hashes, dict) and hashes.get(kind) == value["sha256"],
                f"cleanup does not retain {kind} binary hash")
    return {"kind": kind, "path": str(binary), "sha256": value["sha256"],
            "bytes": value["bytes"], "identity_sha256": sha(path)}


def validate_builds() -> tuple[dict[str, dict[str, Any]], list[tuple[dt.datetime, dt.datetime]]]:
    identities: dict[str, dict[str, Any]] = {}
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for kind in ("normal", "alloc"):
        path = need(BASE / f"build-{kind}.receipt.json", f"baseline/build-{kind} receipt")
        value, times = validate_receipt(path, f"build-{kind}", build_command(kind), None,
                                         expected_artifacts(f"build-{kind}"))
        identities[kind] = binary_identity(kind)
        require(value.get("binary_sha256") is None, f"{rel(path)} carries a build binary hash")
        intervals.append(times)
    return identities, intervals


def capture_jobs(lane: str) -> list[dict[str, Any]]:
    require(lane in {"preflight", "native", "alloc", "profile"},
            f"unknown capture lane: {lane}")
    jobs: list[dict[str, Any]] = []
    repeats = 1 if lane == "preflight" else 2
    for repeat in range(1, repeats + 1):
        shapes = list(SHAPES) if repeat == 1 else list(reversed(SHAPES))
        cases = [PROFILE_CASE] if lane == "profile" else list(CASES)
        for shape in shapes:
            for index, case in enumerate(cases):
                jobs.append({
                    "name": f"{lane}-r{repeat}-{shape}-c{index}",
                    "lane": lane, "repeat": repeat, "shape": shape, "case": case,
                    "warmup": 0 if lane in {"preflight", "profile"} else 3,
                    "samples": 1 if lane == "preflight" else (
                        1 if lane == "profile" else 30),
                })
    return jobs


def capture_command(job: dict[str, Any], kind: str) -> list[str]:
    lane = job["lane"]
    name = job["name"]
    command = ["taskset", "-c", "2", "/usr/bin/time", "-v"]
    if lane == "profile":
        command += ["valgrind", "--vgdb=no",
                    "--vgdb-prefix=" + str(TARGET / "tmp/vgdb"),
                    "--tool=callgrind", "--collect-atstart=no",
                    "--toggle-collect=" + PROFILE_OWNER,
                    "--zero-before=" + PROFILE_OWNER,
                    "--dump-after=" + PROFILE_OWNER,
                    "--callgrind-out-file=" + str(BASE / f"{name}.callgrind")]
    command += [str(SCRATCH_ROOT / "baseline" / kind), "--case", job["case"],
                "--xlsx-cell-crud-shape", job["shape"], "--warmup", str(job["warmup"]),
                "--samples", str(job["samples"]), "--json", str(BASE / f"{name}.json"),
                "--corpus-manifest", str(BASE / f"{name}.catalog.json")]
    return command


def validate_corpus_report(path: Path, job: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    value = read_json(path, rel(path))
    label = rel(path)
    require(isinstance(value, dict) and value.get("schema_version") == 1,
            f"{label} report schema differs")
    tool = value.get("tool")
    expected_binary_name = "litchi-perf-baseline-alloc" if binary["kind"] == "alloc" else "litchi-perf-baseline"
    require(isinstance(tool, dict) and tool.get("name") == "litchi-perf-baseline"
            and tool.get("binary") == expected_binary_name
            and tool.get("profile") == "release"
            and isinstance(tool.get("instrumentation"), str),
            f"{label} tool identity differs")
    identity = value.get("binary_identity")
    require(isinstance(identity, dict)
            and identity.get("path") == binary["path"]
            and identity.get("binary_sha256") == binary["sha256"]
            and identity.get("profile") == "release",
            f"{label} binary identity differs")
    environment = value.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == REVISION
            and environment.get("cpu_affinity") == "2",
            f"{label} benchmark environment differs")
    configuration = value.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("cases") == [job["case"]]
            and configuration.get("xlsx_cell_crud_shapes") == [job["shape"]]
            and configuration.get("samples_per_case") == job["samples"]
            and configuration.get("warmup_iterations_per_case") == job["warmup"],
            f"{label} benchmark configuration differs")
    results = value.get("results")
    require(isinstance(results, list) and len(results) == 1
            and isinstance(results[0], dict)
            and results[0].get("case") == job["case"],
            f"{label} result matrix differs")
    corpus = results[0].get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == job["shape"],
            f"{label} corpus shape differs")
    elapsed = results[0].get("elapsed_ns")
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns"
            and isinstance(elapsed.get("samples"), list)
            and len(elapsed["samples"]) == job["samples"]
            and all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                    for item in elapsed["samples"]),
            f"{label} elapsed samples differ")
    output_digest = results[0].get("output_sha256")
    require(valid_digest(output_digest), f"{label} output digest is malformed")
    return value


def validate_captures(identities: dict[str, dict[str, Any]]) -> tuple[list[tuple[dt.datetime, dt.datetime]], dict[str, Any]]:
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    reports: dict[str, Any] = {}
    for lane in ("preflight", "native", "profile", "alloc"):
        kind = "alloc" if lane == "alloc" else "normal"
        for job in capture_jobs(lane):
            path = need(BASE / f"{job['name']}.receipt.json",
                        f"baseline/{job['name']} receipt")
            value, times = validate_receipt(
                path, job["name"], capture_command(job, kind), identities[kind]["sha256"],
                expected_artifacts(job["name"], lane))
            report = validate_corpus_report(BASE / f"{job['name']}.json", job, identities[kind])
            intervals.append(times)
            reports[job["name"]] = {
                "receipt_sha256": sha(path), "report_sha256": sha(BASE / f"{job['name']}.json"),
                "lane": lane, "repeat": job["repeat"], "shape": job["shape"],
                "case": job["case"], "samples": len(report["results"][0]["elapsed_ns"]["samples"]),
            }
    return intervals, reports


def metadata_commands() -> list[tuple[str, list[str]]]:
    return [
        ("fmt", ["cargo", "fmt", "--all", "--check"]),
        ("harness-fmt", ["cargo", "fmt", "--manifest-path",
                         "tools/perf-baseline/Cargo.toml", "--all", "--check"]),
        ("boundaries", ["python3", "-B", "tools/check_crate_boundaries.py"]),
        ("claims", ["python3", "-B", "tools/check_perf_claims.py", "--registry",
                     "docs/performance/claim-registry-v1.json", "--repo-root", ".",
                     "--evidence-root", ".", "--mode", "strict"]),
    ]


def validate_metadata() -> tuple[list[tuple[dt.datetime, dt.datetime]], dict[str, Any]]:
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    rows: dict[str, Any] = {}
    for name, command in metadata_commands():
        receipt_path = need(BASE / f"check-{name}.receipt.json",
                            f"baseline/check-{name} receipt")
        value, times = validate_receipt(
            receipt_path, f"check-{name}", command, None,
            expected_artifacts(f"check-{name}"))
        intervals.append(times)
        rows[name] = {"receipt_sha256": sha(receipt_path), "exit_code": value["exit_code"]}
    return intervals, rows


def validate_symbols(identity: dict[str, Any]) -> dict[str, Any]:
    path = need(BASE / "symbols.receipt.json", "baseline/symbols receipt")
    value, times = validate_receipt(
        path, "symbols", ["nm", "-C", str(SCRATCH_ROOT / "baseline" / "normal")], identity["sha256"],
        expected_artifacts("symbols"))
    require(value.get("binary_sha256") == identity["sha256"],
            "symbols binary binding differs")
    observation = read_json(HERE / "symbol-observation.json", "symbol-observation.json")
    lines = read_text(BASE / "symbols.stdout", "baseline/symbols.stdout").splitlines()
    require(isinstance(observation, dict)
            and set(observation) == {"owner", "rows", "plan_sha256", "symbols_sha256"}
            and observation.get("owner") == PROFILE_OWNER
            and observation.get("plan_sha256") == sha(PLAN)
            and observation.get("symbols_sha256") == sha(BASE / "symbols.stdout")
            and isinstance(observation.get("rows"), list)
            and observation["rows"]
            and all(isinstance(item, str) and item in lines for item in observation["rows"])
            and all(item.endswith(" " + PROFILE_OWNER) for item in observation["rows"]),
            "symbol observation binding differs")
    return {"status": "pass", "owner": PROFILE_OWNER,
            "receipt_sha256": sha(path), "stdout_sha256": sha(BASE / "symbols.stdout"),
            "interval": [times[0].isoformat(), times[1].isoformat()]}


def bundle_snapshot() -> dict[str, tuple[Any, ...]]:
    """Record every bundle entry so analyzer replay cannot mutate evidence."""
    result: dict[str, tuple[Any, ...]] = {}
    for path in HERE.rglob("*"):
        relative = rel(path)
        if path.is_symlink():
            try:
                result[relative] = ("symlink", os.readlink(path))
            except OSError as error:
                raise VerificationError(
                    f"cannot inspect evidence symlink during replay: {relative}") from error
            continue
        try:
            stat = path.stat()
        except OSError as error:
            raise VerificationError(
                f"cannot inspect evidence entry during replay: {relative}") from error
        if path.is_dir():
            result[relative] = ("directory", stat.st_mtime_ns, stat.st_ino)
        elif path.is_file():
            result[relative] = ("file", stat.st_size, stat.st_mtime_ns, stat.st_ino)
        else:
            raise VerificationError(f"unsupported evidence entry during replay: {relative}")
    return result


def run_analyzer(spec: tuple[str, Path, Path], receipt: dict[str, Any]) -> dict[str, Any]:
    name, script, retained = spec
    need(script, rel(script))
    need(retained, rel(retained))
    command = receipt.get("command")
    require(isinstance(command, list) and command and
            isinstance(command[0], str) and Path(command[0]).name.startswith("python"),
            f"analysis-runs/{name} interpreter differs")
    with tempfile.TemporaryDirectory(prefix=f".litchi-0550-{name}-", dir="/home/zhuhe") as folder:
        output = Path(folder) / retained.name
        replay = list(command)
        script_tokens = {str(script), rel(script), script.name,
                         script.relative_to(REPO).as_posix()}
        script_replaced = False
        output_tokens = {str(retained), rel(retained), retained.name,
                         retained.relative_to(REPO).as_posix()}
        for index, token in enumerate(replay):
            if token in script_tokens:
                replay[index] = str(script)
                script_replaced = True
            if token in output_tokens:
                replay[index] = str(output)
        require(script_replaced,
                f"analysis-runs/{name} command has no canonical analyzer path")
        require(any(token == str(output) for token in replay),
                f"analysis-runs/{name} command has no retained output path")
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
        before = bundle_snapshot()
        try:
            result = subprocess.run(replay, cwd=REPO, env=environment,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    text=True, check=False)
        except OSError as error:
            raise VerificationError(f"{rel(script)} replay could not start: {error}") from error
        require(result.returncode == 0,
                f"{rel(script)} replay failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == retained.read_bytes(),
                f"{rel(script)} replay differs from {rel(retained)}")
        after = bundle_snapshot()
        require(before == after, f"{rel(script)} replay modified the evidence bundle")
    value = read_json(retained, rel(retained))
    require(isinstance(value, dict) and value.get("status") == "pass"
            and value.get("plan_sha256") == sha(PLAN),
            f"{rel(retained)} analyzer envelope differs")
    if "scope" in value:
        require(value["scope"] == SCOPE, f"{rel(retained)} scope differs")
    if "performance_claim" in value:
        require(value["performance_claim"] in ALLOWED_PERFORMANCE_CLAIMS,
                f"{rel(retained)} makes an unsupported performance claim")
    return value


def resolve_input_path(name: str) -> Path:
    safe_relative(name, "analysis input path")
    repository_path = REPO / name
    bundle_path = HERE / name
    if repository_path.is_file() and not repository_path.is_symlink():
        return repository_path
    if bundle_path.is_file() and not bundle_path.is_symlink():
        return bundle_path
    raise IncompleteError(f"analysis input is not retained: {name}")


def validate_analysis_inputs() -> dict[str, Any]:
    value = read_json(ANALYSIS_INPUTS, "analysis-inputs.json")
    require(isinstance(value, dict) and value, "analysis-inputs.json is empty")
    for name, digest in value.items():
        require(isinstance(name, str) and valid_digest(digest),
                "analysis input map entry is malformed")
        safe_relative(name, "analysis input path")
        require("analysis-runs" not in Path(name).parts
                and "__pycache__" not in Path(name).parts,
                f"analysis input path is generated evidence: {name}")
        path = resolve_input_path(name)
        require(sha(path) == digest, f"analysis input hash differs: {name}")
    require(ANALYSIS_REQUIRED_INPUTS <= set(value),
            "analysis input map omits a driver, frozen input, or imported helper")
    return {"status": "pass", "sha256": sha(ANALYSIS_INPUTS), "entries": len(value)}


def validate_analysis_runs(selected: set[str] | None = None) -> dict[str, Any]:
    inputs = validate_analysis_inputs()
    need(ANALYSIS_RUNS, "analysis-runs", directory=True)
    expected_names = {name for name, _, _ in ANALYSIS_SPECS}
    actual_names = {path.name for path in ANALYSIS_RUNS.iterdir()}
    selected = expected_names if selected is None else set(selected)
    require(selected <= expected_names, "analysis analyzer selection differs")
    if selected == expected_names:
        require(actual_names == expected_names, "analysis-runs inventory differs")
    else:
        require(selected <= actual_names,
                "selected analysis run is missing")
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    rows: dict[str, Any] = {}
    for name, script, retained in ANALYSIS_SPECS:
        if name not in selected:
            continue
        folder = need(ANALYSIS_RUNS / name, f"analysis-runs/{name}", directory=True)
        expected_files = {"stdout", "stderr", "receipt.json"}
        require({path.name for path in folder.iterdir()} == expected_files,
                f"analysis-runs/{name} artifact inventory differs")
        receipt_path = need(folder / "receipt.json", f"analysis-runs/{name}/receipt.json")
        receipt = read_json(receipt_path, rel(receipt_path))
        require(isinstance(receipt, dict) and set(receipt) == ANALYSIS_RECEIPT_KEYS,
                f"analysis-runs/{name} receipt schema differs")
        command = receipt.get("command")
        require(isinstance(command, list) and len(command) == 5
                and Path(str(command[0])).name.startswith("python")
                and command[1] == "-B"
                and command[2] in {str(script), rel(script), script.name}
                and command[3] == "--output"
                and command[4] in {str(retained), rel(retained), retained.name,
                                   retained.relative_to(REPO).as_posix()},
                f"analysis-runs/{name} command differs")
        require(receipt.get("script_sha256") == sha(script)
                and receipt.get("plan_sha256") == sha(PLAN)
                and receipt.get("inputs_sha256") == inputs["sha256"]
                and receipt.get("output") == rel(retained)
                and valid_digest(receipt.get("output_sha256"))
                and receipt.get("output_sha256") == sha(retained)
                and receipt.get("exit_code") == 0,
                f"analysis-runs/{name} custody binding differs")
        artifacts = receipt.get("artifacts")
        require(isinstance(artifacts, dict) and set(artifacts) == {"stdout", "stderr"}
                and artifacts.get("stdout") == sha(folder / "stdout")
                and artifacts.get("stderr") == sha(folder / "stderr"),
                f"analysis-runs/{name} artifact binding differs")
        if "seconds" in receipt:
            times = interval(receipt, rel(receipt_path))
        else:
            start = parse_time(receipt.get("start_utc"), rel(receipt_path))
            end = parse_time(receipt.get("end_utc"), rel(receipt_path))
            require(end > start, f"analysis-runs/{name} interval is inverted")
            times = (start, end)
        intervals.append(times)
        report = run_analyzer((name, script, retained), receipt)
        rows[name] = {"receipt_sha256": sha(receipt_path), "report_sha256": sha(retained),
                      "status": report.get("status")}
    ordered = sorted(intervals, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "analysis run intervals overlap")
    return {"status": "pass", "inputs": inputs, "runs": rows,
            "intervals": len(intervals)}


def validate_review_and_documentation() -> dict[str, Any]:
    value = read_json(DOCUMENTATION_MANIFEST, "documentation-manifest.json")
    require(isinstance(value, dict) and set(value) == set(DOCUMENTATION_NAMES),
            "documentation manifest inventory differs")
    for name in DOCUMENTATION_NAMES:
        safe_relative(name, "documentation path")
        path = REPO / name
        require(path.is_file() and not path.is_symlink() and valid_digest(value[name])
                and sha(path) == value[name], f"documentation binding differs: {name}")
        require(path.stat().st_size > 0, f"documentation file is empty: {name}")
    for name in ("source-review.md", "scope-review.md", "profile-review.md", "results-review.md"):
        text = read_text(HERE / name, name)
        require(text.strip(), f"{name} is empty")
    return {"status": "pass", "entries": len(value),
            "sha256": sha(DOCUMENTATION_MANIFEST)}


def validate_receipt_inventory() -> dict[str, Any]:
    expected: set[str] = {"build-normal", "build-alloc", "symbols"}
    expected.update(job["name"] for lane in ("preflight", "native", "profile", "alloc")
                    for job in capture_jobs(lane))
    expected.update(f"check-{name}" for name in QUALITY_NAMES)
    actual = {path.name.removesuffix(".receipt.json") for path in BASE.glob("*.receipt.json")}
    require(actual == expected,
            f"baseline receipt inventory differs (expected {len(expected)}, found {len(actual)})")
    expected_files = {"source-manifest.json", "source.patch",
                      "binary-normal.json", "binary-alloc.json"}
    for name in sorted(expected):
        if name.startswith("preflight-"):
            lane = "preflight"
        elif name.startswith("native-"):
            lane = "native"
        elif name.startswith("profile-"):
            lane = "profile"
        elif name.startswith("alloc-"):
            lane = "alloc"
        else:
            lane = None
        expected_files.add(f"{name}.receipt.json")
        expected_files.update(expected_artifacts(name, lane))
        if lane == "profile" and PROFILE_REPORT.exists():
            expected_files.update({f"{name}.inclusive.txt", f"{name}.self.txt"})
    actual_files = {path.name for path in BASE.iterdir()
                    if path.is_file() and not path.is_symlink()}
    require(actual_files == expected_files,
            "baseline raw artifact inventory contains an unexpected file")
    return {"status": "pass", "receipts": len(actual),
            "files": len(actual_files),
            "names_sha256": hashlib.sha256("\n".join(sorted(actual)).encode()).hexdigest()}


def validate_serial_timeline() -> dict[str, Any]:
    expected: list[str] = ["build-normal", "symbols"]
    expected += [job["name"] for job in capture_jobs("preflight")]
    expected += [job["name"] for job in capture_jobs("native") if job["repeat"] == 1]
    expected += [job["name"] for job in capture_jobs("profile") if job["repeat"] == 1]
    expected += [job["name"] for job in capture_jobs("profile") if job["repeat"] == 2]
    expected += [job["name"] for job in capture_jobs("native") if job["repeat"] == 2]
    expected += ["build-alloc"]
    expected += [job["name"] for job in capture_jobs("alloc") if job["repeat"] == 1]
    expected += [job["name"] for job in capture_jobs("alloc") if job["repeat"] == 2]
    expected += [f"check-{name}" for name in QUALITY_NAMES]
    rows: list[tuple[str, dt.datetime, dt.datetime]] = []
    for name in expected:
        value = read_json(BASE / f"{name}.receipt.json", rel(BASE / f"{name}.receipt.json"))
        start, end = interval(value, f"baseline/{name}.receipt.json")
        rows.append((name, start, end))
    require(all(rows[index][2] <= rows[index + 1][1]
               for index in range(len(rows) - 1)),
            "baseline receipt intervals overlap or order differs")
    return {"status": "pass", "receipts": len(rows), "first": rows[0][0],
            "last": rows[-1][0], "order": [row[0] for row in rows]}


def process_references(target: Path) -> list[dict[str, Any]]:
    needle = str(target)
    found: list[dict[str, Any]] = []
    proc = Path("/proc")
    if not proc.is_dir():
        return found
    for item in proc.iterdir():
        if not item.name.isdigit():
            continue
        pid = item.name
        fields: dict[str, str] = {}
        try:
            fields["cmdline"] = (item / "cmdline").read_bytes().replace(b"\0", b" ").decode(errors="replace")
        except OSError:
            pass
        try:
            fields["cwd"] = os.readlink(item / "cwd")
        except OSError:
            pass
        try:
            fields["exe"] = os.readlink(item / "exe")
        except OSError:
            pass
        if any(needle in value for value in fields.values()):
            found.append({"pid": int(pid), "fields": fields})
            continue
        try:
            for fd in (item / "fd").iterdir():
                try:
                    if needle in os.readlink(fd):
                        found.append({"pid": int(pid), "fd": fd.name,
                                      "target": os.readlink(fd)})
                        break
                except OSError:
                    continue
        except OSError:
            continue
    return found


def validate_cleanup(*, required: bool) -> dict[str, Any]:
    if not CLEANUP.exists():
        if required:
            raise IncompleteError("cleanup.json is missing")
        return {"status": "not-present"}
    value = read_json(CLEANUP, "cleanup.json")
    expected_keys = {
        "accessible_process_references", "binary_sha256_by_kind", "input_sha256",
        "observed_utc", "owned_paths_absent", "plan_sha256", "python_cache_absent",
        "removed", "scope", "target",
    }
    require(isinstance(value, dict) and set(value) == expected_keys,
            "cleanup envelope differs")
    parse_time(value.get("observed_utc"), "cleanup.observed_utc")
    require(value.get("target") == str(TARGET)
            and value.get("removed") == [str(TARGET)]
            and value.get("owned_paths_absent") is True
            and value.get("accessible_process_references") == []
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("python_cache_absent") is True
            and not list(HERE.rglob("__pycache__")), "cleanup envelope differs")
    require(all(not os.path.lexists(path) for path in [str(TARGET)]),
            "owned target remains after claimed cleanup")
    require(process_references(TARGET) == [], "owned target has a live process reference")
    hashes = value.get("binary_sha256_by_kind")
    require(isinstance(hashes, dict) and set(hashes) == {"normal", "alloc"}
            and all(valid_digest(item) for item in hashes.values()),
            "cleanup binary hash map differs")
    for kind in ("normal", "alloc"):
        descriptor = read_json(BASE / f"binary-{kind}.json",
                               f"baseline/binary-{kind}.json")
        require(isinstance(descriptor, dict)
                and hashes[kind] == descriptor.get("sha256"),
                f"cleanup binary hash differs: {kind}")
    inputs = value.get("input_sha256")
    require(isinstance(inputs, dict)
            and set(inputs) == {"plan.json", "run.py", "capture.py", "frozen-inputs.json"},
            "cleanup input hash map differs")
    for name, digest in inputs.items():
        require(valid_digest(digest) and sha(HERE / name) == digest,
                f"cleanup input hash differs: {name}")
    require(isinstance(value.get("scope"), str) and value["scope"].strip(),
            "cleanup scope is missing")
    return {"status": "pass", "sha256": sha(CLEANUP),
            "owned_paths_absent": True, "binary_kinds": sorted(hashes)}


def validate_seal() -> dict[str, Any]:
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), "SHA256SUMS line is malformed")
        safe_relative(fields[1], "SHA256SUMS path")
        require(fields[1] != "SHA256SUMS" and fields[1] not in expected,
                "SHA256SUMS inventory is unsafe or duplicated")
        expected[fields[1]] = fields[0]
    actual = {rel(item): sha(item) for item in HERE.rglob("*")
              if item.is_file() and not item.is_symlink() and item != path}
    require(expected == actual and not any(item.is_symlink() for item in HERE.rglob("*")),
            "SHA256SUMS inventory differs")
    return {"status": "pass", "sha256": sha(path), "entries": len(expected)}


def all_main_evidence(*, require_cleanup: bool) -> dict[str, Any]:
    plan_result = validate_plan()
    host_result = validate_host()
    adr_result = validate_adr()
    source_result = validate_source(plan_result)
    quality_result = validate_quality_continuity()
    tool_result = validate_tool_source_continuity()
    replay_result = validate_private_source_replay(plan_result)
    failed_attempt = validate_failed_attempt()
    identities, build_intervals = validate_builds()
    symbol_result = validate_symbols(identities["normal"])
    capture_intervals, capture_result = validate_captures(identities)
    metadata_intervals, metadata_result = validate_metadata()
    validate_receipt_inventory()
    timeline = validate_serial_timeline()
    all_intervals = build_intervals + capture_intervals + metadata_intervals
    require(all(left[1] <= right[0] for left, right in zip(
        sorted(all_intervals, key=lambda item: item[0]),
        sorted(all_intervals, key=lambda item: item[0])[1:])),
        "all retained main receipts overlap")
    analysis = validate_analysis_runs()
    documentation = validate_review_and_documentation()
    cleanup = validate_cleanup(required=require_cleanup)
    seal = validate_seal() if require_cleanup else {"status": "not-required"}
    return {
        "status": "pass", "plan": plan_result, "host": host_result,
        "adr": adr_result,
        "source": source_result, "private_source_replay": replay_result,
        "quality_continuity": quality_result,
        "tool_source_continuity": tool_result,
        "failed_attempt": failed_attempt,
        "builds": identities, "symbols": symbol_result,
        "captures": {"jobs": len(capture_result), "reports": capture_result},
        "metadata": metadata_result, "timeline": timeline,
        "analysis": analysis, "documentation": documentation,
        "cleanup": cleanup, "seal": seal,
    }


def component(name: str) -> dict[str, Any]:
    if name == "plan":
        return validate_plan()
    plan_result = validate_plan()
    if name == "host":
        return validate_host()
    if name == "adr":
        return validate_adr()
    if name == "source":
        return {"source": validate_source(plan_result),
                "private_source_replay": validate_private_source_replay(plan_result),
                "failed_attempt": validate_failed_attempt(),
                "quality_continuity": validate_quality_continuity(),
                "tool_source_continuity": validate_tool_source_continuity()}
    if name == "build":
        identities, intervals = validate_builds()
        require(intervals[0][1] <= intervals[1][0], "build receipts overlap")
        return {"builds": identities, "intervals": len(intervals)}
    identities, build_intervals = validate_builds()
    if name == "symbols":
        return validate_symbols(identities["normal"])
    if name in {"captures", "native", "allocation", "profile"}:
        capture_intervals, captures = validate_captures(identities)
        require(all(left[1] <= right[0] for left, right in zip(
            sorted(build_intervals + capture_intervals, key=lambda item: item[0]),
            sorted(build_intervals + capture_intervals, key=lambda item: item[0])[1:])),
            "capture receipts overlap")
        return {"jobs": len(captures), "reports": captures}
    if name == "metadata":
        return {"metadata": validate_metadata()[1]}
    if name in {"analysis", "profiles", "metrics"}:
        result = validate_analysis_runs(
            None if name == "analysis" else {name})
        return result
    if name in {"review", "documentation"}:
        return validate_review_and_documentation()
    if name == "cleanup":
        return validate_cleanup(required=True)
    if name == "seal":
        return validate_seal()
    if name == "precleanup":
        return all_main_evidence(require_cleanup=False)
    if name == "all":
        return all_main_evidence(require_cleanup=True)
    raise VerificationError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    try:
        result = component(selected)
        return {"schema": "litchi-0550-xlsx-attribution-verification-v1",
                "status": "pass", "scope": selected,
                "performance_claim": "none", "result": result}
    except IncompleteError as error:
        return {"schema": "litchi-0550-xlsx-attribution-verification-v1",
                "status": "incomplete", "scope": selected,
                "performance_claim": "none", "error": str(error)}
    except (VerificationError, OSError, subprocess.CalledProcessError,
            KeyError, TypeError, AttributeError, IndexError, ValueError) as error:
        return {"schema": "litchi-0550-xlsx-attribution-verification-v1",
                "status": "fail", "scope": selected,
                "performance_claim": "none", "error": str(error)}


def verify(sealed: bool = False) -> dict[str, Any]:
    """Run the complete pre-cleanup or sealed verification for API callers."""
    return component("all" if sealed else "precleanup")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true",
                        help="run the complete pre-cleanup component")
    parser.add_argument("--component", choices=(
        "all", "precleanup", "plan", "host", "adr", "source", "build", "symbols",
        "captures", "native", "allocation", "profile", "metadata", "analysis",
        "profiles", "metrics", "review", "documentation", "cleanup", "seal",
    ), default=None)
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    selected = "precleanup" if args.precleanup else (args.component or "all")
    if args.precleanup and args.component is not None:
        parser.error("--precleanup and --component cannot be combined")
    if args.output and args.output.resolve().is_relative_to(HERE.resolve()):
        parser.error("verification output must be outside the evidence bundle")
    report = run_bundle(selected)
    text = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
    print(text, end="")
    return 1 if report["status"] == "fail" or (
        args.strict and report["status"] == "incomplete") else 0


if __name__ == "__main__":
    raise SystemExit(main())
