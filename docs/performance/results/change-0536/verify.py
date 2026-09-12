"""Fail-closed custody verifier for the matched 0536 CFB experiment.

This module only reads retained evidence.  It replays source patches through a
private Git index, validates every stage receipt and raw artifact, replays the
retained diagnostic analyzers in temporary directories outside the bundle, and
checks the selected quality stage, cleanup, and recursive seal.  It never
builds, captures, or writes a report in the evidence bundle.
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
import shutil
import subprocess
import sys
import tempfile
from typing import Any

sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
CHECKS = HERE / "checks.py"
SUMMARIZE = HERE / "summarize_quality.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
BASELINE = HERE / "baseline"
CANDIDATE = HERE / "candidate"
FINAL = HERE / "final"
CANDIDATE_INPUTS = HERE / "candidate-inputs.json"
ANALYZER = HERE / "analyze.py"
PROFILE_ANALYZER = HERE / "analyze_profiles.py"
HARDWARE_ANALYZER = HERE / "analyze_hardware.py"
INSTRUCTION_ANALYZER = HERE / "instruction_analysis.py"
DECIDER = HERE / "decide.py"
QUALITY_REUSE = HERE / "baseline-quality-reuse.json"
QUALITY_SUMMARY = HERE / "quality-summary.json"
ADVERSE_REVIEW = HERE / "adverse-review.json"
VARIATION_REVIEW = HERE / "variation-review.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
SCRATCH_ROOT = Path("/tmp/litchi-goal-0536")
TARGET = Path("/home/zhuhe/litchi-goal-0536-target")
CANONICAL_BINARY_ROOT = TARGET / "retained-binaries"

PRIOR_0534 = HERE.parent / "change-0534"
PRIOR_0535 = HERE.parent / "change-0535"
STAGES = ("baseline", "candidate", "final")
MEASURED_STAGES = ("baseline", "candidate")
SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
XLS_CASES = (
    "xls_semantic_open", "xls_eager_open_list_worksheets",
    "xls_eager_open_one_cell", "xls_source_backed_open",
    "xls_source_backed_open_list_worksheets", "xls_source_backed_open_one_cell",
    "xls_owned_source_open", "xls_owned_source_open_list_worksheets",
    "xls_owned_source_open_one_cell",
)
CFB_CASE = "cfb_open"
CFB_SHAPES = ("tiny", "many-small", "few-large")
XLS_OWNER_CASE = "xls_owned_source_open_one_cell"
QUALITY_NAMES = (
    "fmt", "harness-fmt", "cfb-tests", "cfb-no-default-tests", "xls-tests",
    "doc-tests", "ppt-tests", "workspace-check", "ole-clippy", "harness-clippy",
    "ole-rustdoc", "harness-rustdoc", "boundaries", "claims",
)
ASSEMBLY_OWNERS = ("SectorChainScratch", "CheckedBitSet", "chain_error")
ASSEMBLY_REQUIRED = ("collect_exact", "insert")
EXPECTED_ENVIRONMENT = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
RECEIPT_KEYS = {
    "command", "start_utc", "end_utc", "seconds", "exit_code", "execution_stage",
    "execution_manifest_sha256", "plan_sha256", "script_sha256",
    "source_manifest_sha256", "binary_sha256", "environment", "artifacts",
}
TIME_FORMAT = '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}'


class VerificationError(ValueError):
    """Malformed, contradictory, or out-of-scope retained evidence."""


class IncompleteError(VerificationError):
    """Evidence needed by a selected component has not arrived yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def rel(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError:
        return str(path)


def need(path: Path, label: str | None = None, *, directory: bool = False) -> Path:
    label = label or rel(path)
    if directory:
        if not path.is_dir() or path.is_symlink():
            raise IncompleteError(f"{label} is missing or not a regular directory")
    elif not path.is_file() or path.is_symlink():
        raise IncompleteError(f"{label} is missing or not a regular file")
    return path


def read_json(path: Path, label: str | None = None) -> Any:
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"cannot read {label or rel(path)}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    need(path, label)
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise VerificationError(f"cannot read {label or rel(path)}: {error}") from error


def sha(path: Path) -> str:
    if not path.is_file() or path.is_symlink():
        raise VerificationError(f"cannot hash non-regular file: {rel(path)}")
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
    require(path.as_posix() == value and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise VerificationError(f"{label} timestamp is malformed") from error
    require(parsed.tzinfo is not None, f"{label} timestamp has no timezone")
    return parsed


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


def source_name(name: str) -> bool:
    return name in SOURCE_EXACT or any(name.startswith(prefix) for prefix in SOURCE_PREFIXES)


def git(args: list[str], *, env: dict[str, str] | None = None,
        input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, env=env, input=input_data,
                                       stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise VerificationError(f"Git command failed ({' '.join(args)}): "
                                f"{detail.decode(errors='replace')[-2000:]}") from error


def source_manifest(path: Path, label: str | None = None) -> dict[str, str]:
    value = read_json(path, label or rel(path))
    require(isinstance(value, dict) and value, f"{label or rel(path)} source manifest is empty")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label or rel(path)} source path")
        require(source_name(name) and valid_digest(digest),
                f"{label or rel(path)} source entry is invalid: {name}")
        require(name not in result, f"{label or rel(path)} repeats {name}")
        result[name] = digest
    return result


def tree_manifest(revision: str) -> dict[str, str]:
    raw = git(["git", "ls-tree", "-r", "-z", revision])
    objects: dict[str, str] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, name_bytes = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "frozen Git source tree entry is malformed")
        name = name_bytes.decode()
        if source_name(name):
            objects[name] = fields[2].decode()
    require(objects, "frozen Git revision has no source objects")
    oids = sorted(set(objects.values()))
    raw_batch = git(["git", "cat-file", "--batch"],
                    input_data=("\n".join(oids) + "\n").encode())
    hashes: dict[str, str] = {}
    position = 0
    while position < len(raw_batch):
        end = raw_batch.find(b"\n", position)
        require(end >= 0, "Git batch response is malformed")
        fields = raw_batch[position:end].split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "Git source object is not a blob")
        oid = fields[0].decode()
        length = int(fields[2])
        position = end + 1
        data = raw_batch[position:position + length]
        require(len(data) == length, "Git source object is truncated")
        hashes[oid] = hashlib.sha256(data).hexdigest()
        position += length
        require(raw_batch[position:position + 1] == b"\n",
                "Git batch separator is missing")
        position += 1
    return {name: hashes[oid] for name, oid in objects.items()}


def parse_index(index: Path) -> dict[str, str]:
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    raw = git(["git", "ls-files", "-s", "-z"], env=env)
    result: dict[str, str] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, name_bytes = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3 and fields[0] in (b"100644", b"100755"),
                "private source replay entry is malformed")
        name = name_bytes.decode()
        require(name not in result, f"private source replay repeats {name}")
        result[name] = fields[1].decode()
    return result


def index_hashes(index: Path, oids: set[str]) -> dict[str, str]:
    if not oids:
        return {}
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    raw = git(["git", "cat-file", "--batch"], env=env,
              input_data=("\n".join(sorted(oids)) + "\n").encode())
    result: dict[str, str] = {}
    position = 0
    while position < len(raw):
        end = raw.find(b"\n", position)
        require(end >= 0, "private Git batch response is malformed")
        oid, kind, size = raw[position:end].split()
        require(kind == b"blob", "private source object is not a blob")
        position = end + 1
        length = int(size)
        data = raw[position:position + length]
        require(len(data) == length, "private Git batch response is truncated")
        result[oid.decode()] = hashlib.sha256(data).hexdigest()
        position += length
        require(raw[position:position + 1] == b"\n", "private Git separator is missing")
        position += 1
    return result


def sidecar_sources(folder: Path) -> dict[str, str]:
    path = folder / "new-files.json"
    if not path.exists():
        return {}
    value = read_json(path, rel(path))
    rows = value.get("files") if isinstance(value, dict) else None
    require(isinstance(rows, dict), f"{rel(path)}.files is not an object")
    result: dict[str, str] = {}
    for name, item in rows.items():
        safe_relative(name, f"{rel(path)} source path")
        require(source_name(name), f"{rel(path)} source path is out of scope: {name}")
        if isinstance(item, str):
            expected, artifact = item, name
        else:
            require(isinstance(item, dict), f"{rel(path)} entry is malformed: {name}")
            expected, artifact = item.get("sha256"), item.get("artifact", name)
        safe_relative(artifact, f"{rel(path)} artifact")
        require(valid_digest(expected), f"{rel(path)} digest is malformed: {name}")
        artifact_path = folder / artifact
        require(artifact_path.is_file() and not artifact_path.is_symlink()
                and sha(artifact_path) == expected,
                f"{rel(path)} source artifact custody differs: {artifact}")
        result[name] = expected
    return result


def replay_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    require(stage in STAGES, f"invalid source stage: {stage}")
    folder = HERE / stage
    need(folder, f"{stage} evidence directory", directory=True)
    manifest_path = need(folder / "source-manifest.json", f"{stage} source manifest")
    patch_path = need(folder / "source.patch", f"{stage} source patch")
    expected = source_manifest(manifest_path)
    with tempfile.TemporaryDirectory(prefix=".litchi-0536-source-", dir="/home/zhuhe") as directory:
        index = Path(directory) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", plan["revision"]], env=env)
        if patch_path.stat().st_size:
            git(["git", "apply", "--cached", "--binary", str(patch_path)], env=env)
        indexed = parse_index(index)
        changed = set(git(["git", "diff", "--cached", "--name-only", plan["revision"]],
                          env=env).decode().splitlines())
        require(all(source_name(name) for name in changed),
                f"{stage} source patch changes an out-of-scope file")
        scoped = {name: oid for name, oid in indexed.items() if source_name(name)}
        hashes = index_hashes(index, set(scoped.values()))
    extras = sidecar_sources(folder)
    require(set(expected) == set(scoped) | set(extras),
            f"{stage} source manifest/index inventory differs")
    for name, expected_digest in expected.items():
        if name in scoped:
            require(hashes.get(scoped[name]) == expected_digest,
                    f"{stage} replay blob differs: {name}")
        else:
            require(extras.get(name) == expected_digest,
                    f"{stage} sidecar source differs: {name}")
    if stage == "baseline":
        require(not changed and not extras, "baseline source patch is not empty")
        require(expected == tree_manifest(plan["revision"]),
                "baseline source manifest differs from the frozen revision")
        prior = source_manifest(PRIOR_0535 / "baseline/source-manifest.json",
                                "change-0535 baseline source manifest")
        require(expected == prior, "0536 baseline is not the exact 0535 source")
    elif stage == "candidate":
        require(changed == {"crates/litchi-cfb/src/file.rs"},
                "candidate source patch does not contain only the CFB production file")
    else:
        require(changed <= {"crates/litchi-cfb/src/file.rs"},
                "final source patch changes an unexpected file")
    return {"stage": stage, "manifest_sha256": sha(manifest_path),
            "manifest_entries": len(expected), "patch_sha256": sha(patch_path),
            "changed_files": sorted(changed), "new_files": sorted(extras)}


def replay_patch_manifest(patch_path: Path, plan: dict[str, Any],
                          expected: dict[str, str], label: str) -> None:
    """Apply a retained top-level patch in a private index and bind its tree."""
    need(patch_path, label)
    with tempfile.TemporaryDirectory(prefix=".litchi-0536-patch-", dir="/home/zhuhe") as directory:
        index = Path(directory) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        git(["git", "read-tree", plan["revision"]], env=env)
        if patch_path.stat().st_size:
            git(["git", "apply", "--cached", "--binary", str(patch_path)], env=env)
        changed = set(git(["git", "diff", "--cached", "--name-only", plan["revision"]],
                          env=env).decode().splitlines())
        require(all(source_name(name) for name in changed),
                f"{label} changes an out-of-scope file")
        indexed = parse_index(index)
        scoped = {name: oid for name, oid in indexed.items() if source_name(name)}
        hashes = index_hashes(index, set(scoped.values()))
    actual = {name: hashes[oid] for name, oid in scoped.items()}
    require(actual == expected,
            f"{label} does not reproduce the retained source manifest")


def current_source_manifest() -> dict[str, str]:
    tracked = git(["git", "ls-files", "-z", "crates", "tools/perf-baseline",
                   "Cargo.toml", "Cargo.lock", ".cargo", "rust-toolchain.toml"])
    names = {item.decode() for item in tracked.split(b"\0") if item}
    untracked = git(["git", "ls-files", "--others", "--exclude-standard", "-z", "--",
                     "crates", "tools/perf-baseline"])
    names.update(item.decode() for item in untracked.split(b"\0")
                 if item and item.endswith(b".rs"))
    names = {name for name in names if source_name(name)}
    return {name: sha(REPO / name) for name in sorted(names)
            if (REPO / name).is_file() and not (REPO / name).is_symlink()}


def validate_adr() -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    prior = read_json(PRIOR_0534 / "adr-manifest.json", "change-0534/adr-manifest.json")
    files = value.get("files") if isinstance(value, dict) else None
    require(isinstance(files, dict) and files and value == prior,
            "ADR manifest differs from the accepted prior manifest")
    for name, digest in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith("docs/adr/") and valid_digest(digest)
                and (REPO / name).is_file() and sha(REPO / name) == digest,
                f"ADR hash differs: {name}")
    require(valid_digest(value.get("prior_audit_source_binding_sha256")),
            "ADR prior audit binding is malformed")
    return {"status": "pass", "entries": len(files), "sha256": sha(ADR)}


def check_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict) and set(frozen) == {"created_utc", "files"},
            "frozen input envelope differs")
    parse_time(frozen.get("created_utc"), "frozen-inputs.json.created_utc")
    files = frozen.get("files")
    expected_files = {
        "plan.json", "run.py", "capture_baseline.py", "capture_candidate.py",
        "inspect_assembly.py", "adr-manifest.json",
    }
    require(isinstance(files, dict) and set(files) == expected_files,
            "frozen input inventory differs")
    for name, expected in files.items():
        require(isinstance(name, str) and Path(name).name == name and valid_digest(expected),
                f"frozen input entry is malformed: {name!r}")
        path = HERE / name
        require(path.is_file() and not path.is_symlink() and sha(path) == expected,
                f"frozen input digest differs: {name}")

    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict)
            and plan.get("status") == "frozen-before-build-capture-and-candidate"
            and plan.get("scope") == "Matched CFB collector cold-error-layout experiment"
            and plan.get("priority") ==
            "OLE2/OOXML active; ODF deferred until that goal completes; iWork excluded"
            and plan.get("hypothesis") == (
                "Move rare formatted collector errors into private cold helpers while "
                "preserving all successful checks and exact error/allocation/reset ordering"
            )
            and plan.get("cpu") == 2
            and plan.get("owned_paths") == [str(SCRATCH_ROOT), str(TARGET)],
            "plan envelope differs")
    revision = plan.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError("plan revision is not a Git commit") from error
    groups = plan.get("groups")
    require(isinstance(groups, dict)
            and groups.get("xls", {}).get("cases") == list(XLS_CASES)
            and groups.get("cfb", {}).get("cases") == [CFB_CASE]
            and groups.get("cfb", {}).get("shapes") == list(CFB_SHAPES)
            and groups.get("cfb", {}).get("payload") == "incompressible",
            "plan source groups differ")
    native = plan.get("native")
    require(isinstance(native, dict) and native.get("repeats") == 2
            and native.get("warmup") == 20 and native.get("samples") == 1000
            and native.get("order") == [
                "baseline r1 xls/cfb", "candidate r1 xls/cfb",
                "candidate r2 cfb/xls", "baseline r2 cfb/xls",
            ], "native plan differs")
    allocation = plan.get("allocation")
    require(isinstance(allocation, dict) and allocation.get("repeats") == 2
            and allocation.get("warmup") == 3 and allocation.get("samples") == 30
            and allocation.get("scope") == (
                "Existing constructor/operation clock with canonical operation-global "
                "System allocator region; excludes construction of input fixtures, "
                "correctness oracles, report construction and object drop"
            ),
            "allocation plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict) and profile.get("repeats") == 2
            and profile.get("warmup") == 0 and profile.get("samples") == 5
            and profile.get("jobs") == [
                "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large",
            ]
            and profile.get("xls_owner") ==
            "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
            and profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open"
            and profile.get("scope") == (
                "zero-before/dump-after/toggle exact constructor; retain setup CFB open "
                "dump and use only positive incoming benchmark-runner dumps for operation "
                "attribution"
            )
            and profile.get("instruction_flags") == [
                "--dump-instr=yes", "--dump-line=no", "--compress-pos=no",
                "--collect-jumps=yes",
            ], "profile plan differs")
    hardware = plan.get("hardware")
    require(isinstance(hardware, dict) and hardware.get("repeats") == 2
            and hardware.get("case") == XLS_OWNER_CASE
            and hardware.get("warmup") == 0 and hardware.get("samples") == 1000
            and hardware.get("events") ==
            "{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations",
            "hardware plan differs")
    require(hardware.get("scope") == (
        "Whole-child diagnostic including setup/copies/queries/oracles/drop/report; "
        "no operation-local hardware claim; unavailable or multiplexed groups cannot "
        "support claims"
    ), "hardware scope differs")
    assembly = plan.get("assembly")
    require(isinstance(assembly, dict)
            and assembly.get("owners") == list(ASSEMBLY_OWNERS)
            and assembly.get("required") == list(ASSEMBLY_REQUIRED),
            "assembly plan differs")
    review = plan.get("review")
    require(isinstance(review, dict) and review.get("same_build_adverse_percent") == 5
            and review.get("matched_adverse_percent") == 5
            and review.get("primary_cases") == [
                "xls_source_backed_open", "xls_source_backed_open_one_cell",
                "xls_owned_source_open", "xls_owned_source_open_one_cell",
            ]
            and review.get("policy") == (
                "Adopt only if all four primary workflow p50 values improve by at least "
                "3% in both paired repeats; XLS-owned constructor Ir and collector "
                "exclusive Ir in XLS-owned and CFB-few-large decrease in both repeats; "
                "allocation calls/bytes/incremental peak do not grow materially; and "
                "correctness/quality pass. Review every >5% matched latency/RSS regression "
                "and every absolute >5% same-build variation. No stable-tail or broad "
                "provider/scaling claim. Revert production if admission fails, retain "
                "independent contract tests after restored quality. The exact final "
                "runtime-plus-test source and binary must be measured; any changed rebuild "
                "requires fresh full ABBA."
            ),
            "review plan differs")
    reuse = plan.get("baseline_quality_reuse")
    require(isinstance(reuse, dict)
            and reuse.get("source_bundle") == "change-0535"
            and reuse.get("quality_bundle") == "change-0534"
            and reuse.get("quality_stage") == "final"
            and reuse.get("checks") == 14 and reuse.get("executed_tests") == 4382
            and "Exact source manifest equality" in str(reuse.get("condition", "")),
            "baseline quality reuse plan differs")
    require(assembly.get("scope") ==
            "Same measured binary; all matching symbols captured; instruction mapping, "
            "not static latency", "assembly scope differs")
    require(plan.get("temporary_storage") == (
        "Use owned target/tmp from first child; prior0532 confirmed /tmp EDQUOT; "
        "vgdb=no from first profile"
    ), "temporary storage policy differs")
    return plan


def validate_candidate_inputs() -> dict[str, Any]:
    """Bind the frozen production patch and quality-driver inputs.

    ``candidate-inputs.json`` is created after the campaign plan was frozen,
    so it is checked here as a separate custody record.  The two hashes are
    deliberately checked against the files actually used by the candidate
    stage; this prevents a later replacement of either the production patch
    or the quality command inventory from being silently accepted.
    """
    value = read_json(CANDIDATE_INPUTS, "candidate-inputs.json")
    require(isinstance(value, dict)
            and set(value) == {
                "created_utc", "files", "reviewed_preformat_patch_sha256",
                "formatting",
            }, "candidate-inputs envelope differs")
    parse_time(value.get("created_utc"), "candidate-inputs.json.created_utc")
    files = value.get("files")
    require(isinstance(files, dict) and set(files) == {"candidate.patch", "checks.py"},
            "candidate-inputs file inventory differs")
    expected = {
        "candidate.patch": HERE / "candidate.patch",
        "checks.py": CHECKS,
    }
    for name, path in expected.items():
        require(files.get(name) == sha(path),
                f"candidate-inputs digest differs: {name}")
    require(valid_digest(value.get("reviewed_preformat_patch_sha256"))
            and value.get("formatting") ==
            "root rustfmt edition2024 applied beforecandidatefreeze; finaldiffreviewrequired",
            "candidate-inputs review binding differs")
    return {"status": "pass", "sha256": sha(CANDIDATE_INPUTS),
            "files": {name: files[name] for name in sorted(files)}}


def validate_source(plan: dict[str, Any]) -> dict[str, Any]:
    candidate_inputs = validate_candidate_inputs()
    baseline = replay_stage("baseline", plan)
    candidate = replay_stage("candidate", plan)
    require(baseline["manifest_sha256"] != candidate["manifest_sha256"],
            "candidate source manifest is unchanged")
    candidate_manifest = source_manifest(CANDIDATE / "source-manifest.json")
    replay_patch_manifest(HERE / "candidate.patch", plan, candidate_manifest,
                          "candidate.patch")
    require((CANDIDATE / "source.patch").read_bytes() ==
            (HERE / "candidate.patch").read_bytes(),
            "candidate source.patch differs from candidate.patch")
    final = replay_stage("final", plan) if FINAL.exists() else None
    manifests = {
        "baseline": source_manifest(BASELINE / "source-manifest.json"),
        "candidate": source_manifest(CANDIDATE / "source-manifest.json"),
    }
    if final is not None:
        manifests["final"] = source_manifest(FINAL / "source-manifest.json")
    decision_path = HERE / "decision.json"
    selected = None
    if decision_path.is_file():
        decision = read_json(decision_path, "decision.json")
        selected = decision.get("final_source")
        if selected is None:
            selected = "candidate" if decision.get("disposition") == "accepted" else "final"
        require(selected in manifests, "decision selects an unretained source stage")
        if decision.get("disposition") == "rejected":
            require(selected == "final" and final is not None
                    and manifests["final"] == manifests["baseline"]
                    and final.get("changed_files") == []
                    and (FINAL / "source.patch").stat().st_size == 0,
                    "rejected decision does not restore the baseline source")
        require(current_source_manifest() == manifests[selected],
                "current checkout does not match the selected source manifest")
    else:
        require(current_source_manifest() in manifests.values(),
                "current checkout does not match a retained source stage")
    return {"status": "pass", "candidate_inputs": candidate_inputs,
            "baseline": baseline, "candidate": candidate,
            "final": final, "selected": selected, "adr": validate_adr(),
            "manifests": {name: sha(HERE / name / "source-manifest.json")
                          for name in manifests}}


def validate_recursive_seal(root: Path, seal: Path, label: str) -> dict[str, Any]:
    need(seal, label)
    expected: dict[str, str] = {}
    for line in read_text(seal, label).splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]),
                f"{label} line differs")
        safe_relative(fields[1], f"{label} path")
        require(fields[1] != seal.name and fields[1] not in expected,
                f"{label} inventory is unsafe or duplicated")
        expected[fields[1]] = fields[0]
    actual = {path.relative_to(root).as_posix(): sha(path)
              for path in root.rglob("*")
              if path.is_file() and not path.is_symlink() and path != seal}
    require(expected == actual and not any(path.is_symlink() for path in root.rglob("*")),
            f"{label} recursive inventory differs")
    return {"status": "pass", "sha256": sha(seal), "entries": len(expected)}


def validate_quality_reuse(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(QUALITY_REUSE, "baseline-quality-reuse.json")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0535-diagnostic-verification-v1"
            and value.get("status") == "pass"
            and value.get("scope") == "all"
            and value.get("performance_claim") == "none; baseline diagnostics only"
            and value.get("prior_verifier_sha256") == sha(PRIOR_0535 / "verify.py")
            and isinstance(value.get("result"), dict)
            and value["result"].get("status") == "pass"
            and value.get("reused_by") == "0536 baseline only",
            "baseline quality reuse envelope differs")
    result = value["result"]
    source = result.get("source")
    require(isinstance(source, dict) and source.get("status") == "pass"
            and source.get("manifest_sha256") ==
            sha(PRIOR_0535 / "baseline/source-manifest.json"),
            "baseline quality reuse source binding differs")
    prior = result.get("prior_quality")
    require(isinstance(prior, dict) and prior.get("status") == "pass"
            and prior.get("checks") == 14 and prior.get("executed_tests") == 4382,
            "baseline quality reuse gate count differs")
    require(prior.get("summary_sha256") == sha(PRIOR_0534 / "quality-summary.json")
            and prior.get("verification_sha256") == sha(PRIOR_0534 / "verification.json")
            and prior.get("seal", {}).get("sha256") == sha(PRIOR_0534 / "SHA256SUMS"),
            "baseline quality reuse prior-final hashes differ")
    for key in ("cleanup", "seal", "numeric", "instruction", "inventory"):
        require(isinstance(result.get(key), dict)
                and result[key].get("status") == "pass",
                f"baseline quality reuse lacks passing {key} custody")
    retained_seal = validate_recursive_seal(
        PRIOR_0535, PRIOR_0535 / "SHA256SUMS", "change-0535/SHA256SUMS"
    )
    require(result["seal"].get("sha256") == retained_seal["sha256"],
            "baseline quality reuse 0535 seal differs")
    return {"status": "pass", "sha256": sha(QUALITY_REUSE),
            "prior_verifier_sha256": value["prior_verifier_sha256"],
            "source_manifest_sha256": source["manifest_sha256"],
            "checks": prior["checks"], "executed_tests": prior["executed_tests"],
            "seal": retained_seal}


def validate_host(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict)
            and set(value) == {"observed_utc", "compiler_processes", "scope"}
            and value.get("scope") ==
            "Accessible compiler processes; no host quiescence guarantee"
            and isinstance(value.get("compiler_processes"), list),
            f"{label} host observation differs")
    parse_time(value.get("observed_utc"), f"{label}.observed_utc")
    for process in value["compiler_processes"]:
        require(isinstance(process, dict) and isinstance(process.get("pid"), int)
                and process["pid"] > 0 and process.get("comm") in {"cargo", "rustc"}
                and isinstance(process.get("cwd"), str) and process["cwd"],
                f"{label} compiler process row differs")


def expected_artifacts(folder: Path, name: str, lane: str | None = None) -> set[str]:
    if name.startswith(("build-", "check-", "assembly-")) or name == "symbols":
        return {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"}
    if lane is None:
        if name.startswith("native-"):
            lane = "native"
        elif name.startswith("alloc-"):
            lane = "alloc"
        elif name.startswith("profile-"):
            lane = "profile"
        elif name.startswith("hardware-"):
            lane = "hardware"
        else:
            raise VerificationError(f"cannot infer artifact lane for {name}")
    result = {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr",
              f"{name}.json", f"{name}.catalog.json"}
    if lane == "native":
        result.add(f"{name}.rss.json")
    elif lane == "alloc":
        pass
    elif lane == "hardware":
        result.add(f"{name}.csv")
    elif lane == "profile":
        result.add(f"{name}.callgrind")
        suffixes = sorted(
            int(path.name.rsplit(".", 1)[1])
            for path in folder.glob(f"{name}.callgrind.*")
            if path.is_file() and path.name.rsplit(".", 1)[1].isdigit()
        )
        expected_count = 5 if name.endswith("xls-owned") else 6
        require(suffixes == list(range(1, expected_count + 1)),
                f"{rel(folder)}/{name} numbered Callgrind inventory differs")
        result.update(f"{name}.callgrind.{number}" for number in suffixes)
    else:
        raise VerificationError(f"unknown artifact lane: {lane}")
    return result


def is_derived_annotation(name: str) -> bool:
    """Recognize only the profile analyzer's deterministic sidecar outputs."""
    return re.fullmatch(
        r"profile-[A-Za-z0-9_-]+\.part-\d+\.(?:inclusive|self)\.txt", name
    ) is not None


def validate_artifacts(folder: Path, value: dict[str, Any], expected: set[str], label: str) -> None:
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} raw artifact inventory differs")
    stems = {filename.split(".", 1)[0] for filename in expected}
    require(len(stems) == 1 and next(iter(stems)),
            f"{label} artifact names are malformed")
    stem = next(iter(stems))
    try:
        actual = {path.name for path in folder.iterdir()
                  if path.name.startswith(stem + ".")
                  and path.name != stem + ".receipt.json"
                  and not is_derived_annotation(path.name)}
    except OSError as error:
        raise VerificationError(f"{label} artifact directory cannot be read") from error
    require(actual == expected,
            f"{label} raw artifact files differ from the receipt inventory")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename
                and valid_digest(digest), f"{label} artifact entry differs")
        require(not is_derived_annotation(filename),
                f"{label} receipt includes a derived annotation")
        path = folder / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label} raw artifact custody differs: {filename}")
        if filename.endswith(".host.json"):
            validate_host(path, f"{label}/{filename}")


def validate_receipt(path: Path, name: str, stage: str, plan: dict[str, Any],
                     command: list[str], expected_binary: str | None,
                     expected_artifacts_set: set[str], *, allow_failure: bool = False,
                     expected_execution_stage: str | None = None) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS
            and value.get("command") == command,
            f"{rel(path)} command or receipt schema differs")
    exit_code = value.get("exit_code")
    require(isinstance(exit_code, int), f"{rel(path)} exit code is invalid")
    require(allow_failure or exit_code == 0, f"{rel(path)} did not exit successfully")
    source_sha = sha(HERE / stage / "source-manifest.json")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("script_sha256") == sha(RUN)
            and value.get("source_manifest_sha256") == source_sha,
            f"{rel(path)} plan/script/source binding differs")
    execution = value.get("execution_stage")
    require(execution in STAGES, f"{rel(path)} execution stage is invalid")
    expected_execution_stage = expected_execution_stage or stage
    require(execution == expected_execution_stage
            and value.get("execution_manifest_sha256") ==
            sha(HERE / execution / "source-manifest.json"),
            f"{rel(path)} execution binding differs")
    require(value.get("binary_sha256") == expected_binary,
            f"{rel(path)} binary binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            f"{rel(path)} instrumentation environment differs")
    times = interval(value, rel(path))
    validate_artifacts(path.parent, value, expected_artifacts_set, rel(path))
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


def binary_identity(stage: str, kind: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    path = need(folder / f"binary-{kind}.json", f"{stage} {kind} binary identity")
    value = read_json(path, rel(path))
    expected_path = SCRATCH_ROOT / stage / kind
    require(isinstance(value, dict)
            and set(value) == {
                "path", "sha256", "bytes", "build_receipt_sha256",
                "source_manifest_sha256",
            }
            and value.get("path") == str(expected_path)
            and valid_digest(value.get("sha256"))
            and isinstance(value.get("bytes"), int) and value["bytes"] > 0
            and value.get("source_manifest_sha256") == sha(folder / "source-manifest.json")
            and value.get("build_receipt_sha256") == sha(folder / f"build-{kind}.receipt.json"),
            f"{rel(path)} identity differs")
    binary = Path(value["path"])
    if binary.exists():
        require(binary.is_file() and not binary.is_symlink()
                and SCRATCH_ROOT.is_symlink()
                and SCRATCH_ROOT.resolve() == CANONICAL_BINARY_ROOT
                and CANONICAL_BINARY_ROOT.is_dir()
                and not CANONICAL_BINARY_ROOT.is_symlink()
                and binary.resolve() == CANONICAL_BINARY_ROOT / stage / kind
                and sha(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                f"{rel(path)} binary alias or custody differs")
    else:
        cleanup = read_json(CLEANUP, "cleanup.json")
        require(cleanup.get("plan_sha256") == sha(PLAN)
                and cleanup.get("removed") == plan["owned_paths"]
                and cleanup.get("owned_paths_absent") is True
                and all(not os.path.lexists(item) for item in plan["owned_paths"]),
                f"{rel(path)} missing binary lacks cleanup custody")
    return {"kind": kind, "path": str(binary), "sha256": value["sha256"],
            "bytes": value["bytes"], "identity_sha256": sha(path)}


def validate_builds(stage: str, plan: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], list[tuple[dt.datetime, dt.datetime]]]:
    identities: dict[str, dict[str, Any]] = {}
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for kind in ("normal", "alloc"):
        name = f"build-{kind}"
        path = need(HERE / stage / f"{name}.receipt.json", f"{stage} {name} receipt")
        value, times = validate_receipt(
            path, name, stage, plan, build_command(kind), None,
            expected_artifacts(HERE / stage, name),
        )
        require(value.get("binary_sha256") is None,
                f"{rel(path)} unexpectedly carries a binary hash")
        intervals.append(times)
        identities[kind] = binary_identity(stage, kind, plan)
    return identities, intervals


def capture_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    if lane == "native":
        jobs = []
        for repeat in range(1, 3):
            groups = ("xls", "cfb") if repeat == 1 else ("cfb", "xls")
            for group in groups:
                jobs.append({"name": f"native-r{repeat}-{group}", "lane": lane,
                             "repeat": repeat, "group": group,
                             "selection": plan["groups"][group], "warmup": 20,
                             "samples": 1000})
        return jobs
    if lane == "alloc":
        return [{"name": f"alloc-r{repeat}-{group}", "lane": lane,
                 "repeat": repeat, "group": group,
                 "selection": plan["groups"][group], "warmup": 3, "samples": 30}
                for repeat in range(1, 3) for group in ("xls", "cfb")]
    if lane == "profile":
        profile = plan["profile"]
        jobs = []
        for repeat in range(1, 3):
            jobs.append({"name": f"profile-r{repeat}-xls-owned", "lane": lane,
                         "repeat": repeat, "group": "xls-owned", "shape": None,
                         "owner": profile["xls_owner"], "warmup": 0, "samples": 5,
                         "selection": {"cases": [XLS_OWNER_CASE]}})
            for shape in CFB_SHAPES:
                jobs.append({"name": f"profile-r{repeat}-cfb-{shape}", "lane": lane,
                             "repeat": repeat, "group": f"cfb-{shape}", "shape": shape,
                             "owner": profile["cfb_owner"], "warmup": 0, "samples": 5,
                             "selection": {"cases": [CFB_CASE], "shapes": [shape],
                                            "payload": "incompressible"}})
        return jobs
    if lane == "hardware":
        return [{"name": f"hardware-r{repeat}-xls-owned", "lane": lane,
                 "repeat": repeat, "group": "xls-owned", "warmup": 0, "samples": 1000,
                 "selection": {"cases": [plan["hardware"]["case"]]}}
                for repeat in range(1, 3)]
    raise VerificationError(f"unknown capture lane: {lane}")


def capture_command(stage: str, job: dict[str, Any], plan: dict[str, Any], kind: str) -> list[str]:
    folder = HERE / stage
    binary = SCRATCH_ROOT / stage / kind
    command = ["taskset", "-c", str(plan["cpu"])]
    lane = job["lane"]
    if lane == "native":
        command += ["/usr/bin/time", "-f", TIME_FORMAT, "-o",
                    str(folder / f"{job['name']}.rss.json")]
    elif lane == "profile":
        command += ["valgrind", "--vgdb=no", "--tool=callgrind", "--collect-atstart=no",
                    "--toggle-collect=" + job["owner"],
                    "--zero-before=" + job["owner"], "--dump-after=" + job["owner"],
                    "--callgrind-out-file=" + str(folder / f"{job['name']}.callgrind"),
                    *plan["profile"]["instruction_flags"]]
    elif lane == "hardware":
        command += ["perf", "stat", "-x", ",", "-o",
                    str(folder / f"{job['name']}.csv"), "-e", plan["hardware"]["events"], "--"]
    selection = job["selection"]
    command += [str(binary), "--case", ",".join(selection["cases"]), "--warmup",
                str(job["warmup"]), "--samples", str(job["samples"]), "--json",
                str(folder / f"{job['name']}.json"), "--corpus-manifest",
                str(folder / f"{job['name']}.catalog.json")]
    if "shapes" in selection:
        command += ["--shape", ",".join(selection["shapes"]), "--payload", selection["payload"]]
    return command


def validate_captures(stage: str, plan: dict[str, Any],
                      identities: dict[str, dict[str, Any]]) -> list[tuple[dt.datetime, dt.datetime]]:
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    folder = HERE / stage
    for lane in ("native", "alloc", "profile", "hardware"):
        for job in capture_jobs(plan, lane):
            kind = "alloc" if lane == "alloc" else "normal"
            expected_execution = (
                "candidate" if stage == "baseline" and job["name"] in {
                    "native-r2-cfb", "native-r2-xls",
                } else stage
            )
            path = need(folder / f"{job['name']}.receipt.json",
                        f"{stage} {job['name']} receipt")
            value, times = validate_receipt(
                path, job["name"], stage, plan,
                capture_command(stage, job, plan, kind), identities[kind]["sha256"],
                expected_artifacts(folder, job["name"], lane),
                allow_failure=lane == "hardware",
                expected_execution_stage=expected_execution,
            )
            intervals.append(times)
            require(value["command"] == capture_command(stage, job, plan, kind),
                    f"{stage} {job['name']} command is not frozen")
    return intervals


def validate_assembly(stage: str, plan: dict[str, Any],
                      binary: dict[str, Any]) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    folder = HERE / stage
    index_path = need(folder / "assembly-index.json", f"{stage} assembly index")
    index = read_json(index_path, rel(index_path))
    require(isinstance(index, dict)
            and set(index) == {"plan_sha256", "binary_sha256", "source_manifest_sha256",
                               "script_sha256", "rows"}
            and index.get("plan_sha256") == sha(PLAN)
            and index.get("binary_sha256") == binary["sha256"]
            and index.get("source_manifest_sha256") == sha(folder / "source-manifest.json")
            and index.get("script_sha256") == sha(HERE / "inspect_assembly.py")
            and isinstance(index.get("rows"), list) and index["rows"],
            f"{rel(index_path)} envelope differs")
    symbols_path = need(folder / "symbols.receipt.json", f"{stage} symbols receipt")
    _, symbols_interval = validate_receipt(
        symbols_path, "symbols", stage, plan,
        ["nm", "-S", "--defined-only", str(SCRATCH_ROOT / stage / "normal")],
        binary["sha256"], expected_artifacts(folder, "symbols"),
    )
    symbols: dict[str, tuple[int, int]] = {}
    for line in read_text(folder / "symbols.stdout", rel(folder / "symbols.stdout")).splitlines():
        fields = line.split()
        if len(fields) == 4 and fields[2] in ("t", "T"):
            try:
                symbols[fields[3]] = (int(fields[0], 16), int(fields[1], 16))
            except ValueError:
                continue
    matching = {symbol: value for symbol, value in symbols.items()
                if any(owner in symbol for owner in ASSEMBLY_OWNERS)}
    rows = index["rows"]
    require({row.get("symbol") for row in rows if isinstance(row, dict)} == set(matching),
            f"{rel(index_path)} does not retain every matching owner symbol")
    seen_names: set[str] = set()
    seen_symbols: set[str] = set()
    intervals = [symbols_interval]
    for number, row in enumerate(rows):
        require(isinstance(row, dict)
                and set(row) == {"name", "symbol", "address_hex", "size_bytes", "receipt_sha256"}
                and row.get("name") == f"assembly-{number}"
                and isinstance(row.get("symbol"), str) and row["symbol"] in matching
                and row["name"] not in seen_names and row["symbol"] not in seen_symbols
                and valid_digest(row.get("receipt_sha256"))
                and re.fullmatch(r"[0-9a-fA-F]+", row.get("address_hex", "")) is not None
                and isinstance(row.get("size_bytes"), int) and row["size_bytes"] > 0,
                f"{rel(index_path)} assembly row {number} differs")
        symbol = row["symbol"]
        require(matching[symbol] == (int(row["address_hex"], 16), row["size_bytes"]),
                f"{rel(index_path)} symbol is not bound to nm: {symbol}")
        receipt_path = need(folder / f"{row['name']}.receipt.json",
                            f"{stage} {row['name']} receipt")
        _, times = validate_receipt(
            receipt_path, row["name"], stage, plan,
            ["objdump", "-d", "--disassemble=" + symbol,
             str(SCRATCH_ROOT / stage / "normal")], binary["sha256"],
            expected_artifacts(folder, row["name"], "assembly"),
        )
        require(row["receipt_sha256"] == sha(receipt_path),
                f"{rel(receipt_path)} hash differs")
        seen_names.add(row["name"])
        seen_symbols.add(symbol)
        intervals.append(times)
    require(all(any(required in symbol for symbol in matching)
                for required in ASSEMBLY_REQUIRED),
            f"{rel(index_path)} omits a required function")
    return {"status": "pass", "rows": len(rows), "symbols": len(symbols),
            "index_sha256": sha(index_path)}, intervals


def replay_instruction_report(args: list[str], retained: Path) -> Any:
    """Replay instruction analysis from an unsealed temporary evidence view.

    The retained instruction tool refuses to create any report once a
    ``SHA256SUMS`` file exists.  Symlinking the immutable evidence into a
    temporary directory keeps that guard meaningful while allowing the
    verifier to replay the analyzer after the real bundle has been sealed.
    """
    need(retained, rel(retained))
    before = {path: (path.stat().st_size, path.stat().st_mtime_ns)
              for path in HERE.rglob("*")
              if path.is_file() and not path.is_symlink()}
    with tempfile.TemporaryDirectory(prefix=".litchi-0536-instruction-", dir="/home/zhuhe") as directory:
        view = Path(directory) / "bundle"
        view.mkdir()
        for item in HERE.iterdir():
            if item.name in {"SHA256SUMS", "instruction_analysis.py"}:
                continue
            os.symlink(item, view / item.name, target_is_directory=item.is_dir())
        script = view / "instruction_analysis.py"
        shutil.copy2(INSTRUCTION_ANALYZER, script)
        output = Path(directory) / retained.name
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
        command = [sys.executable, "-B", str(script), *args,
                   "--output", str(output)]
        try:
            result = subprocess.run(command, cwd=REPO, env=environment,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    text=True, check=False)
        except OSError as error:
            raise VerificationError(
                f"{rel(INSTRUCTION_ANALYZER)} replay could not start: {error}"
            ) from error
        require(result.returncode == 0,
                f"{rel(INSTRUCTION_ANALYZER)} replay failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == retained.read_bytes(),
                f"{rel(INSTRUCTION_ANALYZER)} replay differs from {rel(retained)}")
    after = {path: (path.stat().st_size, path.stat().st_mtime_ns)
             for path in HERE.rglob("*")
             if path.is_file() and not path.is_symlink()}
    require(before == after,
            f"{rel(INSTRUCTION_ANALYZER)} replay modified the evidence bundle")
    return read_json(retained, rel(retained))


def replay_report(script: Path, args: list[str], retained: Path) -> Any:
    need(script, rel(script))
    need(retained, rel(retained))
    if script == INSTRUCTION_ANALYZER:
        return replay_instruction_report(args, retained)
    if script == PROFILE_ANALYZER:
        for stage in MEASURED_STAGES:
            for dump in sorted((HERE / stage).glob("profile-*.callgrind.*")):
                suffix = dump.name.rsplit(".", 1)[-1]
                if not suffix.isdigit() or ("-cfb-" in dump.name and suffix == "1"):
                    continue
                stem = dump.name.rsplit(".callgrind.", 1)[0]
                for kind in ("inclusive", "self"):
                    annotation = HERE / stage / f"{stem}.part-{suffix}.{kind}.txt"
                    need(annotation, rel(annotation))
    before = {path: (path.stat().st_size, path.stat().st_mtime_ns)
              for path in HERE.rglob("*")
              if path.is_file() and not path.is_symlink()}
    with tempfile.TemporaryDirectory(prefix=".litchi-0536-report-", dir="/home/zhuhe") as directory:
        output = Path(directory) / retained.name
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
        instruction_args = script == INSTRUCTION_ANALYZER
        command = [sys.executable, "-B", str(script), *args]
        if instruction_args:
            command += ["--output", str(output)]
        else:
            command += [str(output)]
        try:
            result = subprocess.run(
                command,
                cwd=REPO, env=environment, stdout=subprocess.PIPE,
                stderr=subprocess.PIPE, text=True, check=False,
            )
        except OSError as error:
            raise VerificationError(f"{rel(script)} replay could not start: {error}") from error
        require(result.returncode == 0,
                f"{rel(script)} replay failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == retained.read_bytes(),
                f"{rel(script)} replay differs from {rel(retained)}")
    after = {path: (path.stat().st_size, path.stat().st_mtime_ns)
             for path in HERE.rglob("*")
             if path.is_file() and not path.is_symlink()}
    require(before == after, f"{rel(script)} replay modified the evidence bundle")
    return read_json(retained, rel(retained))


def validate_numeric(plan: dict[str, Any]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    for stage in MEASURED_STAGES:
        retained = HERE / stage / "analysis.json"
        value = replay_report(ANALYZER, ["--stage", stage], retained)
        require(isinstance(value, dict)
                and value.get("schema") == "cfb_ole2_matched_numeric_analysis_v1"
                and value.get("status") == "pass"
                and value.get("stage") == stage
                and value.get("scope") == plan["scope"]
                and value.get("priority") == plan["priority"]
                and value.get("plan_sha256") == sha(PLAN)
                and value.get("native_samples") == 24000
                and value.get("allocation_samples") == 720
                and isinstance(value.get("rows"), list)
                and len(value["rows"]) == 48,
                f"{stage} numeric analysis envelope differs")
        values[stage] = value
    comparison_path = HERE / "comparison.json"
    comparison = replay_report(ANALYZER, ["--compare"], comparison_path)
    require(isinstance(comparison, dict)
            and comparison.get("schema") == "cfb_ole2_matched_comparison_v1"
            and comparison.get("status") == "pass"
            and comparison.get("stage") == "compare"
            and comparison.get("plan_sha256") == sha(PLAN)
            and comparison.get("baseline", {}).get("stage") == "baseline"
            and comparison.get("candidate", {}).get("stage") == "candidate"
            and isinstance(comparison.get("comparison"), dict),
            "numeric comparison envelope differs")
    return {"status": "pass", "stages": values,
            "comparison": comparison,
            "analysis_sha256": {stage: sha(HERE / stage / "analysis.json")
                                 for stage in MEASURED_STAGES},
            "comparison_sha256": sha(comparison_path)}


def validate_profiles(plan: dict[str, Any]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    for stage in MEASURED_STAGES:
        retained = HERE / stage / "profile-analysis.json"
        value = replay_report(PROFILE_ANALYZER, ["--stage", stage], retained)
        require(isinstance(value, dict)
                and value.get("schema") ==
                "cfb_ole2_constructor_callgrind_profile_analysis_v2"
                and value.get("status") == "pass"
                and value.get("stage_selection") == [stage]
                and value.get("stage", {}).get("stage") == stage
                and value.get("performance_claim") == "diagnostic-only"
                and value.get("plan_sha256") == sha(PLAN)
                and value.get("profile_count") == 8
                and value.get("timed_constructor_dump_count") == 40
                and value.get("setup_dump_count") == 6
                and value.get("comparison") is None,
                f"{stage} profile analysis envelope differs")
        values[stage] = value
    comparison_path = HERE / "profile-comparison.json"
    comparison = replay_report(PROFILE_ANALYZER, ["--compare"], comparison_path)
    require(isinstance(comparison, dict)
            and comparison.get("schema") ==
            "cfb_ole2_constructor_callgrind_profile_comparison_v1"
            and comparison.get("status") == "pass"
            and comparison.get("stage_selection") == list(MEASURED_STAGES)
            and comparison.get("comparison", {}).get("available") is True,
            "profile comparison envelope differs")
    return {"status": "pass", "profiles": values, "comparison": comparison,
            "profile_sha256": {stage: sha(HERE / stage / "profile-analysis.json")
                                for stage in MEASURED_STAGES},
            "comparison_sha256": sha(comparison_path)}


def validate_instruction(plan: dict[str, Any]) -> dict[str, Any]:
    """Replay the instruction-position diagnostic and its stage comparison."""
    scope = (
        "Stage-local Callgrind positions:instr attribution for exclusive "
        "SectorChainScratch::collect_exact; call/jump metadata is separate "
        "and is not an operation-local count or latency claim."
    )
    values: dict[str, Any] = {}
    expected_validation = {
        "raw_dump_count", "timed_dump_count", "setup_dump_count",
        "positions_instr_only", "events_ir_only", "single_bias_per_job",
        "single_bias_across_profiles", "collector_instruction_ir_equals_function_self",
        "collector_parent_attributed", "setup_and_timed_separated",
        "direct_callees_separate_from_instruction_ir",
        "no_operation_local_timing_claim", "chain_error_fragment_is_optional_and_explicit",
    }
    for stage in MEASURED_STAGES:
        retained = HERE / stage / "instruction-analysis.json"
        value = replay_report(INSTRUCTION_ANALYZER, ["--stage", stage], retained)
        binary = read_json(HERE / stage / "binary-normal.json",
                           f"{stage} binary-normal.json")
        require(isinstance(value, dict)
                and set(value) == {
                    "schema", "stage", "scope", "performance_claim", "plan_sha256",
                    "source_manifest_sha256", "build_receipt", "build_receipt_sha256",
                    "binary_sha256", "binary_identity", "binary_identity_sha256",
                    "assembly_index_sha256", "assembly", "jobs", "aggregates",
                    "validation",
                }
                and value.get("schema") == "litchi-ole2-change-0536-instruction-analysis-v1"
                and value.get("stage") == stage
                and value.get("scope") == scope
                and value.get("performance_claim") is None
                and value.get("plan_sha256") == sha(PLAN)
                and value.get("source_manifest_sha256") ==
                sha(HERE / stage / "source-manifest.json")
                and value.get("build_receipt") == f"{stage}/build-normal.receipt.json"
                and value.get("build_receipt_sha256") ==
                sha(HERE / stage / "build-normal.receipt.json")
                and value.get("binary_sha256") == binary.get("sha256")
                and value.get("binary_identity") == f"{stage}/binary-normal.json"
                and value.get("binary_identity_sha256") ==
                sha(HERE / stage / "binary-normal.json")
                and value.get("assembly_index_sha256") ==
                sha(HERE / stage / "assembly-index.json")
                and isinstance(value.get("jobs"), list)
                and len(value["jobs"]) == 8
                and isinstance(value.get("aggregates"), dict)
                and set(value["aggregates"]) == {"setup", "timed"}
                and isinstance(value.get("validation"), dict)
                and set(value["validation"]) == expected_validation,
                f"{stage} instruction analysis envelope differs")
        names = {job["name"] for job in capture_jobs(plan, "profile")}
        require({job.get("name") for job in value["jobs"]} == names,
                f"{stage} instruction profile job matrix differs")
        for job in value["jobs"]:
            require(isinstance(job, dict)
                    and job.get("timed_dump_count") == 5
                    and job.get("setup_dump_count") in {0, 1}
                    and job.get("dump_count") ==
                    job["timed_dump_count"] + job["setup_dump_count"],
                    f"{stage} instruction job dump counts differ")
        validation = value["validation"]
        boolean_validation = {
            key: item for key, item in validation.items()
            if key not in {"raw_dump_count", "timed_dump_count", "setup_dump_count"}
        }
        require(all(item is True for item in boolean_validation.values())
                and validation["raw_dump_count"] == 46
                and validation["timed_dump_count"] == 40
                and validation["setup_dump_count"] == 6,
                f"{stage} instruction validation flags differ")
        assembly = value["assembly"]
        index = read_json(HERE / stage / "assembly-index.json",
                          f"{stage} assembly index")
        require(isinstance(assembly, dict)
                and set(assembly) == {
                    "rows", "collector_rows", "insert_rows", "chain_error_rows",
                    "chain_error_symbols", "collector_symbols", "insert_symbols",
                    "collector_instruction_count", "required_fragments_present",
                    "optional_fragments_observed",
                }
                and assembly.get("rows") == len(index.get("rows", []))
                and assembly.get("rows") > 0
                and assembly.get("collector_rows") > 0
                and assembly.get("insert_rows") > 0
                and isinstance(assembly.get("chain_error_rows"), int)
                and assembly.get("chain_error_rows") >= 0
                and isinstance(assembly.get("collector_instruction_count"), int)
                and assembly.get("collector_instruction_count") > 0
                and assembly.get("required_fragments_present") == ["collect_exact", "insert"]
                and assembly.get("optional_fragments_observed") in ([], ["chain_error"]),
                f"{stage} instruction assembly summary differs")
        values[stage] = value
    retained = HERE / "instruction-analysis-comparison.json"
    comparison = replay_report(INSTRUCTION_ANALYZER, ["--compare"], retained)
    require(isinstance(comparison, dict)
            and set(comparison) == {
                "schema", "scope", "performance_claim", "plan_sha256", "stages",
                "job_matrix", "validation",
            }
            and comparison.get("schema") ==
            "litchi-ole2-change-0536-instruction-comparison-v1"
            and comparison.get("scope") == (
                "Matched stage instruction mapping for collect_exact and optional "
                "chain_error helper code shape; Ir and call metadata remain diagnostic"
            )
            and comparison.get("performance_claim") is None
            and comparison.get("plan_sha256") == sha(PLAN)
            and isinstance(comparison.get("stages"), dict)
            and set(comparison["stages"]) == set(MEASURED_STAGES)
            and comparison.get("job_matrix") == sorted(
                job["name"] for job in capture_jobs(plan, "profile")
            )
            and isinstance(comparison.get("validation"), dict)
            and set(comparison["validation"]) == {
                "stage_reports_valid", "matched_job_matrix",
                "dynamic_chain_error_symbols_retained", "no_static_instruction_latency_claim",
            }
            and all(item is True for item in comparison["validation"].values()),
            "instruction comparison envelope differs")
    for stage in MEASURED_STAGES:
        row = comparison["stages"][stage]
        require(isinstance(row, dict)
                and row.get("report_stage") == stage
                and row.get("source_manifest_sha256") ==
                sha(HERE / stage / "source-manifest.json")
                and row.get("binary_sha256") == values[stage]["binary_sha256"],
                f"{stage} instruction comparison binding differs")
    return {"status": "pass", "instruction_sha256": {
                stage: sha(HERE / stage / "instruction-analysis.json")
                for stage in MEASURED_STAGES
            }, "comparison_sha256": sha(retained)}


def validate_hardware(plan: dict[str, Any]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    for stage in MEASURED_STAGES:
        retained = HERE / stage / "hardware-analysis.json"
        value = replay_report(HARDWARE_ANALYZER, ["--stage", stage], retained)
        require(isinstance(value, dict)
                and value.get("status") == "pass"
                and value.get("scope") == plan["hardware"]["scope"]
                and isinstance(value.get("captures"), list)
                and len(value["captures"]) == 2
                and value.get("latency_samples_excluded_from_native") == 2000
                and value.get("no_operation_local_hardware_or_speedup_claim") is True,
                f"{stage} hardware analysis envelope differs")
        for row in value["captures"]:
            require(isinstance(row, dict)
                    and row.get("name") in {job["name"] for job in capture_jobs(plan, "hardware")}
                    and row.get("status") in {
                        "measured", "unavailable", "unavailable_for_group_claim",
                    }, f"{stage} hardware analysis row differs")
        values[stage] = value
    return {"status": "pass", "hardware": values,
            "hardware_sha256": {stage: sha(HERE / stage / "hardware-analysis.json")
                                for stage in MEASURED_STAGES}}


def _profile_row(profiles: dict[str, Any], stage: str, group: str,
                 repeat: int, shape: str | None = None) -> dict[str, Any]:
    require(stage in MEASURED_STAGES and isinstance(profiles.get(stage), dict),
            f"{stage} profile analysis is missing")
    rows = profiles[stage].get("profiles")
    require(isinstance(rows, list), f"{stage} profile matrix is missing")
    matches = [item for item in rows
               if isinstance(item, dict) and item.get("group") == group
               and item.get("repeat") == repeat
               and (shape is None or item.get("shape") == shape)]
    require(len(matches) == 1,
            f"{stage} profile {group}/{shape or 'none'} repeat {repeat} is not unique")
    return matches[0]


def _constructor_ir(profile: dict[str, Any], label: str) -> int:
    attribution = profile.get("constructor_attribution")
    require(isinstance(attribution, list) and attribution,
            f"{label} constructor attribution is missing")
    total = 0
    for item in attribution:
        require(isinstance(item, dict) and isinstance(item.get("constructor"), dict),
                f"{label} constructor attribution row is malformed")
        value = item["constructor"].get("inclusive_ir")
        require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
                f"{label} constructor Ir is not positive")
        total += value
    require(total > 0, f"{label} constructor Ir total is not positive")
    return total


def profile_ir_gate(profiles: dict[str, Any]) -> tuple[list[dict[str, Any]], bool]:
    """Replay the independent XLS-owned constructor-inclusive-Ir gate."""
    rows: list[dict[str, Any]] = []
    for repeat in (1, 2):
        values = {
            stage: _constructor_ir(
                _profile_row(profiles, stage, "xls-owned", repeat),
                f"{stage} XLS-owned repeat {repeat}",
            )
            for stage in MEASURED_STAGES
        }
        rows.append({"repeat": repeat, "baseline": values["baseline"],
                     "candidate": values["candidate"],
                     "change_percent": 100 * (values["candidate"] /
                                               values["baseline"] - 1)})
    return rows, all(row["candidate"] < row["baseline"] for row in rows)


COLLECTOR_TARGET = "litchi_cfb::file::SectorChainScratch::collect_exact"


def _collector_target(item: dict[str, Any], label: str) -> dict[str, Any]:
    functions = item.get("functions")
    require(isinstance(functions, dict), f"{label} function attribution is missing")
    matches = [value for value in functions.values()
               if isinstance(value, dict) and value.get("target") == COLLECTOR_TARGET]
    if not matches:
        # The analyzer may expose a stable alias while retaining the exact
        # target identity in the row.  The alias is accepted only after that
        # identity check below.
        for key in ("collector", "collect_exact"):
            value = functions.get(key)
            if isinstance(value, dict):
                matches.append(value)
    require(len(matches) == 1,
            f"{label} collect_exact attribution is not unique")
    target = matches[0]
    require(target.get("target") == COLLECTOR_TARGET,
            f"{label} collect_exact target identity differs")
    return target


def _collector_ir(profile: dict[str, Any], label: str) -> int:
    attribution = profile.get("constructor_attribution")
    require(isinstance(attribution, list) and attribution,
            f"{label} collector attribution is missing")
    total = 0
    for item in attribution:
        require(isinstance(item, dict), f"{label} collector attribution row is malformed")
        target = _collector_target(item, label)
        require(target.get("out_of_line") is True
                and target.get("inlined_or_absent") is False
                and isinstance(target.get("incoming_edge_count"), int)
                and target["incoming_edge_count"] > 0,
                f"{label} collect_exact is not a positive out-of-line target")
        value = target.get("self_ir")
        require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
                f"{label} collect_exact exclusive Ir is malformed")
        total += value
    return total


def collector_profile_gate(profiles: dict[str, Any]) -> tuple[list[dict[str, Any]], bool]:
    """Replay collect_exact exclusive-Ir for both required workload shapes."""
    rows: list[dict[str, Any]] = []
    selected = (("xls-owned", None), ("cfb-few-large", "few-large"))
    for repeat in (1, 2):
        for group, shape in selected:
            values = {
                stage: _collector_ir(
                    _profile_row(profiles, stage, group, repeat, shape),
                    f"{stage} {group} repeat {repeat}",
                )
                for stage in MEASURED_STAGES
            }
            require(values["baseline"] > 0,
                    f"baseline {group} repeat {repeat} collect_exact Ir is not positive")
            rows.append({
                "repeat": repeat,
                "group": group,
                "shape": shape,
                "baseline": values["baseline"],
                "candidate": values["candidate"],
                "change_percent": 100 * (values["candidate"] /
                                          values["baseline"] - 1),
            })
    return rows, all(row["candidate"] < row["baseline"] for row in rows)


def review_rows(raw: list[Any], checked: Any, label: str) -> None:
    require(isinstance(checked, list) and len(checked) == len(raw),
            f"{label} review length differs")
    remaining = list(checked)
    for item in raw:
        require(isinstance(item, dict), f"{label} raw row is malformed")
        index = next((i for i, candidate in enumerate(remaining)
                      if isinstance(candidate, dict)
                      and all(candidate.get(key) == value
                              for key, value in item.items())), None)
        require(index is not None, f"{label} row is not individually reviewed")
        candidate = remaining.pop(index)
        explanation = candidate.get("review", candidate.get("reason"))
        require(isinstance(explanation, str) and explanation.strip(),
                f"{label} row lacks an explanation")
    require(not remaining, f"{label} review has extra rows")


def validate_variation(analysis: dict[str, Any]) -> dict[str, Any]:
    path = need(VARIATION_REVIEW, "variation-review.json")
    value = read_json(path, rel(path))
    raw: list[Any] = []
    stages = analysis.get("stages")
    require(isinstance(stages, dict), "numeric stage reports are missing")
    for stage in MEASURED_STAGES:
        report = stages.get(stage)
        require(isinstance(report, dict), f"{stage} numeric report is missing")
        rows = report.get("same_build_variations_over_five_percent")
        require(isinstance(rows, list), f"{stage} same-build variation rows are missing")
        raw.extend(rows)
    reviewed = value.get("variations") if isinstance(value, dict) else None
    allowed_hashes = set(analysis.get("analysis_sha256", {}).values())
    comparison_hash = analysis.get("comparison_sha256")
    if valid_digest(comparison_hash):
        allowed_hashes.add(comparison_hash)
    require(isinstance(reviewed, list)
            and valid_digest(value.get("analysis_sha256"))
            and value.get("analysis_sha256") in allowed_hashes
            and value.get("threshold_percent") == 5
            and isinstance(value.get("review"), str) and value["review"].strip()
            and value.get("complete", True) is True,
            "variation review envelope differs")
    review_rows(raw, reviewed, "same-build variation")
    return {"status": "pass", "sha256": sha(path), "variations": len(raw)}


def quality_commands() -> list[tuple[str, list[str]]]:
    tree = ast.parse(read_text(CHECKS, "checks.py"), filename=str(CHECKS))
    value = None
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(target, ast.Name) and target.id == "COMMANDS"
                for target in node.targets):
            value = ast.literal_eval(node.value)
            break
    require(isinstance(value, list), "checks.py COMMANDS is missing")
    result: list[tuple[str, list[str]]] = []
    for row in value:
        require(isinstance(row, (list, tuple)) and len(row) == 2
                and isinstance(row[0], str) and isinstance(row[1], list),
                "checks.py command row is malformed")
        require(row[0] not in {name for name, _ in result},
                "checks.py repeats a quality command")
        result.append((row[0], row[1]))
    require([name for name, _ in result] == list(QUALITY_NAMES),
            "checks.py quality command inventory differs")
    return result


def quality_prefix() -> list[str]:
    return ["env", "TMPDIR=" + str(TARGET / "test-tmp"),
            "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
            "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings"]


def test_count(stdout: str) -> int:
    return sum(int(item) for item in
               re.findall(r"test result: ok\. (\d+) passed;", stdout))


def validate_quality(plan: dict[str, Any]) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    summary = read_json(QUALITY_SUMMARY, "quality-summary.json")
    require(isinstance(summary, dict)
            and set(summary) == {"status", "stage", "checks", "executed_tests"}
            and summary.get("status") == "pass"
            and summary.get("stage") in {"candidate", "final"},
            "quality summary envelope differs")
    stage = summary["stage"]
    commands = dict(quality_commands())
    rows = summary.get("checks")
    require(isinstance(rows, list) and len(rows) == len(commands),
            "quality summary check count differs")
    expected_names = {f"check-{name}.receipt.json" for name in commands}
    seen: set[str] = set()
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    total = 0
    for row in rows:
        require(isinstance(row, dict)
                and set(row) == {"name", "receipt_sha256", "executed_tests"},
                "quality summary row schema differs")
        name = row["name"]
        require(name in expected_names and name not in seen and
                valid_digest(row["receipt_sha256"]),
                "quality summary receipt name differs")
        seen.add(name)
        command_name = name.removeprefix("check-").removesuffix(".receipt.json")
        path = need(HERE / stage / name, f"{stage} {name}")
        value, times = validate_receipt(
            path, name.removesuffix(".receipt.json"), stage, plan,
            quality_prefix() + commands[command_name], None,
            expected_artifacts(HERE / stage, name.removesuffix(".receipt.json")),
        )
        require(value.get("exit_code") == 0 and row["receipt_sha256"] == sha(path),
                f"{rel(path)} quality binding differs")
        stdout = path.with_name(path.name.replace(".receipt.json", ".stdout"))
        count = test_count(read_text(stdout, rel(stdout)))
        require(isinstance(row["executed_tests"], int)
                and not isinstance(row["executed_tests"], bool)
                and row["executed_tests"] == count,
                f"{rel(path)} test count differs")
        intervals.append(times)
        total += count
    require(seen == expected_names and isinstance(summary.get("executed_tests"), int)
            and summary["executed_tests"] == total,
            "quality summary aggregate differs")
    return ({"status": "pass", "stage": stage, "checks": len(rows),
             "executed_tests": total, "sha256": sha(QUALITY_SUMMARY)}, intervals)


def validate_optional_quality_receipts(
        plan: dict[str, Any], selected_stage: str
) -> list[tuple[dt.datetime, dt.datetime]]:
    """Validate complete retained quality runs outside the selected summary."""
    commands = dict(quality_commands())
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    baseline_found = sorted(BASELINE.glob("check-*.receipt.json"))
    require(not baseline_found,
            "baseline contains new quality receipts despite sealed reuse")
    for stage in ("candidate", "final"):
        if stage == selected_stage:
            continue
        folder = HERE / stage
        found = sorted(folder.glob("check-*.receipt.json")) if folder.is_dir() else []
        if not found:
            continue
        expected = {f"check-{name}.receipt.json" for name in commands}
        require({path.name for path in found} == expected,
                f"{stage} quality receipt inventory differs")
        for path in found:
            name = path.name.removesuffix(".receipt.json")
            command_name = name.removeprefix("check-")
            value, times = validate_receipt(
                path, name, stage, plan, quality_prefix() + commands[command_name], None,
                expected_artifacts(folder, name),
            )
            require(value.get("exit_code") == 0, f"{rel(path)} quality check failed")
            intervals.append(times)
    return intervals


def expected_core_receipts(stage: str, plan: dict[str, Any]) -> set[str]:
    names = {"build-normal", "build-alloc", "symbols"}
    for lane in ("native", "alloc", "profile", "hardware"):
        names.update(job["name"] for job in capture_jobs(plan, lane))
    index = read_json(HERE / stage / "assembly-index.json",
                      rel(HERE / stage / "assembly-index.json"))
    rows = index.get("rows") if isinstance(index, dict) else None
    require(isinstance(rows, list) and rows,
            f"{stage} assembly rows are missing")
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("name"), str),
                f"{stage} assembly receipt name is malformed")
        names.add(row["name"])
    return names


def validate_receipt_inventory(plan: dict[str, Any]) -> dict[str, Any]:
    expected: dict[str, set[str]] = {}
    check_names = {f"check-{name}" for name in QUALITY_NAMES}
    for stage in MEASURED_STAGES:
        core = expected_core_receipts(stage, plan)
        folder = HERE / stage
        found = {path.name.removesuffix(".receipt.json")
                 for path in folder.glob("*.receipt.json")}
        found_checks = found & check_names
        require(found_checks in (set(), check_names),
                f"{stage} quality receipt inventory is partial")
        require(found == core | found_checks,
                f"{stage} receipt inventory differs: expected {len(core | found_checks)}, "
                f"found {len(found)}")
        # Baseline quality is intentionally represented by the sealed reuse
        # record; no fresh baseline test receipt may appear in this bundle.
        require(stage != "baseline" or not found_checks,
                "baseline quality receipts are not allowed with sealed reuse")
        expected[stage] = found
    if FINAL.is_dir():
        folder = FINAL
        found = {path.name.removesuffix(".receipt.json")
                 for path in folder.glob("*.receipt.json")}
        if found:
            require(found == check_names or
                    found == expected_core_receipts("final", plan) | check_names,
                    "final receipt inventory differs")
            expected["final"] = found
    count = sum(len(names) for names in expected.values())
    inventory_text = "\n".join(
        f"{stage}/{name}" for stage in sorted(expected)
        for name in sorted(expected[stage])
    ) + "\n"
    return {
        "status": "pass",
        "stages": {stage: len(names) for stage, names in expected.items()},
        "successful_receipts": count,
        "receipt_names_sha256": hashlib.sha256(inventory_text.encode()).hexdigest(),
    }


def validate_receipt_timeline() -> dict[str, Any]:
    rows: list[tuple[dt.datetime, dt.datetime, str]] = []
    for stage in STAGES:
        folder = HERE / stage
        if not folder.is_dir():
            continue
        for path in folder.glob("*.receipt.json"):
            value = read_json(path, rel(path))
            start, end = interval(value, rel(path))
            rows.append((start, end,
                         f"{stage}/{path.name.removesuffix('.receipt.json')}"))
    require(rows, "no serial receipts retained")
    ordered = sorted(rows, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "retained receipt intervals overlap")
    native = [row[2] for row in ordered if "/native-" in row[2]]
    require(native == [
        "baseline/native-r1-xls", "baseline/native-r1-cfb",
        "candidate/native-r1-xls", "candidate/native-r1-cfb",
        "candidate/native-r2-cfb", "candidate/native-r2-xls",
        "baseline/native-r2-cfb", "baseline/native-r2-xls",
    ], f"native ABBA order differs: {native}")
    return {"status": "pass", "serial_intervals": len(ordered),
            "first": ordered[0][2], "last": ordered[-1][2],
            "native_abba": native}


def validate_decision(plan: dict[str, Any], analysis: dict[str, Any],
                      profiles: dict[str, Any], quality: dict[str, Any]) -> dict[str, Any]:
    path = need(HERE / "decision.json", "decision.json")
    value = read_json(path, rel(path))
    require(isinstance(value, dict)
            and value.get("disposition") in {"accepted", "rejected"}
            and isinstance(value.get("native_primary_gate"), bool)
            and isinstance(value.get("memory_gate"), bool)
            and isinstance(value.get("profile_gate"), bool)
            and isinstance(value.get("collector_profile_gate"), bool)
            and isinstance(value.get("adoption_allowed"), bool)
            and isinstance(value.get("constructor_ir"), list)
            and isinstance(value.get("collector_ir"), list)
            and value.get("adverse_review_complete") is True
            and value.get("quality_gates") == 14
            and value.get("comparison_sha256") == analysis["comparison_sha256"]
            and value.get("profiles_sha256") == {
                stage: sha(HERE / stage / "profile-analysis.json")
                for stage in MEASURED_STAGES
            }
            and value.get("quality_summary_sha256") == sha(QUALITY_SUMMARY)
            and value.get("adverse_review_sha256") == sha(ADVERSE_REVIEW)
            and value.get("scope") == (
                "Matched CFB collector cold-error-layout experiment; OLE2/OOXML active, "
                "ODF deferred, iWork excluded"
            ), "decision envelope differs")

    comparison = analysis["comparison"].get("comparison")
    require(isinstance(comparison, dict), "decision comparison payload is missing")
    admission = comparison.get("admission")
    require(isinstance(admission, dict), "decision admission payload is missing")
    primary = admission.get("primary_workflow_p50", {}).get(
        "all_four_cases_both_repeats_pass")
    memory = admission.get("allocation_guard", {}).get("passes")
    require(isinstance(primary, bool) and isinstance(memory, bool),
            "decision numeric gate values are malformed")
    constructor_ir, profile_gate = profile_ir_gate(profiles["profiles"])
    collector_ir, collector_gate = collector_profile_gate(profiles["profiles"])
    require(value["native_primary_gate"] is primary
            and value["memory_gate"] is memory
            and value["profile_gate"] is profile_gate
            and value["collector_profile_gate"] is collector_gate
            and value["constructor_ir"] == constructor_ir
            and value["collector_ir"] == collector_ir,
            "decision gate values differ from independent replay")

    review = read_json(ADVERSE_REVIEW, "adverse-review.json")
    require(isinstance(review, dict)
            and review.get("comparison_sha256") == analysis["comparison_sha256"]
            and review.get("complete") is True
            and isinstance(review.get("adoption_allowed"), bool),
            "adverse review envelope differs")
    review_rows(comparison.get("matched_adverse_flags_over_five_percent"),
                review.get("matched"), "matched")
    review_rows(comparison.get("same_build_variations_over_five_percent"),
                review.get("same_build"), "same-build")
    adoption = bool(primary and memory and profile_gate and collector_gate
                   and review["adoption_allowed"])
    expected_disposition = "accepted" if adoption else "rejected"
    require(value["adoption_allowed"] is adoption
            and value["disposition"] == expected_disposition,
            "decision disposition differs from independent gates")
    require(value.get("other_rejection_reason") == review.get("rejection_reason"),
            "decision rejection reason differs from adverse review")

    # Re-run only the pure decision script in a temporary directory.  The
    # retained decision must be byte-identical and no analyzer may write the
    # evidence bundle as a side effect.
    replayed = replay_report(DECIDER, [], path)
    require(replayed == value, "decision does not replay from decide.py")
    selected = value.get("final_source")
    if selected is None:
        selected = "candidate" if value["disposition"] == "accepted" else "final"
    require(selected in {"candidate", "final"},
            "decision final source is invalid")
    manifest_path = HERE / selected / "source-manifest.json"
    require(manifest_path.is_file(), "decision final source manifest is missing")
    require(current_source_manifest() == source_manifest(manifest_path),
            "current checkout does not match decision final source")
    if value["disposition"] == "rejected":
        require(selected == "final" and quality["stage"] == "final",
                "rejected decision does not retain final restored quality")
    else:
        require(selected in {"candidate", "final"} and quality["stage"] == selected,
                "accepted decision does not retain candidate-equivalent quality")
    return {"status": "pass", "disposition": value["disposition"],
            "final_source": selected, "sha256": sha(path),
            "constructor_ir": constructor_ir, "collector_ir": collector_ir,
            "profile_gate": profile_gate, "collector_profile_gate": collector_gate}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict)
            and set(value) == {
                "observed_utc", "plan_sha256", "removed",
                "accessible_process_references", "owned_paths_absent",
                "python_cache_absent", "scope",
            }
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("removed") == plan["owned_paths"]
            and value.get("accessible_process_references") == []
            and value.get("owned_paths_absent") is True
            and value.get("python_cache_absent") is True,
            "cleanup receipt differs")
    parse_time(value.get("observed_utc"), "cleanup.json.observed_utc")
    require(value.get("scope") == (
        "Accessible process cwd/exe, arguments, mappings, and file descriptors "
        "checked before removal."
    ), "cleanup scope differs")
    require(all(not os.path.lexists(path) for path in plan["owned_paths"])
            and not list(HERE.rglob("__pycache__")),
            "owned path or Python cache remains")
    return {"status": "pass", "sha256": sha(CLEANUP),
            "owned_paths_absent": True, "python_cache_absent": True}


def validate_seal() -> dict[str, Any]:
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]),
                "SHA256SUMS line differs")
        safe_relative(fields[1], "SHA256SUMS path")
        require(fields[1] != "SHA256SUMS" and fields[1] not in expected,
                "SHA256SUMS inventory is unsafe or duplicated")
        expected[fields[1]] = fields[0]
    actual = {rel(item): sha(item) for item in HERE.rglob("*")
              if item.is_file() and not item.is_symlink() and item != SEAL}
    require(expected == actual
            and not any(item.is_symlink() for item in HERE.rglob("*")),
            "SHA256SUMS inventory differs")
    return {"status": "pass", "sha256": sha(SEAL), "entries": len(expected)}


def check_serial(intervals: list[tuple[dt.datetime, dt.datetime]], label: str) -> None:
    ordered = sorted(intervals, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{label} receipt intervals overlap")


def component(name: str) -> dict[str, Any]:
    plan = check_plan()
    if name == "plan":
        return {"status": "pass", "plan_sha256": sha(PLAN),
                "frozen_inputs_sha256": sha(FROZEN)}
    reuse = validate_quality_reuse(plan)
    if name in {"quality-reuse", "reuse"}:
        return reuse
    source = validate_source(plan)
    if name == "source":
        return source
    if name == "cleanup":
        return validate_cleanup(plan)
    if name == "seal":
        return validate_seal()

    identities: dict[str, dict[str, dict[str, Any]]] = {}
    build_intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in MEASURED_STAGES:
        identities[stage], times = validate_builds(stage, plan)
        build_intervals.extend(times)
    if name in {"build", "builds"}:
        check_serial(build_intervals, "build")
        return {"status": "pass", "builds": identities,
                "intervals": len(build_intervals)}

    assemblies: dict[str, Any] = {}
    assembly_intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in MEASURED_STAGES:
        assemblies[stage], times = validate_assembly(
            stage, plan, identities[stage]["normal"]
        )
        assembly_intervals.extend(times)
    if name == "assembly":
        check_serial(build_intervals + assembly_intervals, "build/assembly")
        return {"status": "pass", "builds": identities, "assembly": assemblies,
                "intervals": len(build_intervals) + len(assembly_intervals)}

    capture_intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in MEASURED_STAGES:
        capture_intervals.extend(validate_captures(stage, plan, identities[stage]))
    if name in {"captures", "native", "allocation"}:
        check_serial(build_intervals + assembly_intervals + capture_intervals,
                     "build/assembly/capture")
        return {"status": "pass", "builds": identities, "assembly": assemblies,
                "captures": len(capture_intervals),
                "intervals": len(build_intervals) + len(assembly_intervals) +
                len(capture_intervals)}

    # Receipt inventory is intentionally available before cleanup.  Keep the
    # serial check in this branch so callers cannot obtain a passing inventory
    # result while omitting the ABBA/timestamp custody check.
    if name == "inventory":
        quality, quality_intervals = validate_quality(plan)
        optional = validate_optional_quality_receipts(plan, quality["stage"])
        inventory = validate_receipt_inventory(plan)
        timeline = validate_receipt_timeline()
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals + optional, "all retained receipts")
        return {"status": "pass", "inventory": inventory, "timeline": timeline,
                "quality": quality, "intervals": len(build_intervals) +
                len(assembly_intervals) + len(capture_intervals) +
                len(quality_intervals) + len(optional)}

    analysis = validate_numeric(plan)
    if name == "analysis":
        return analysis
    profiles = validate_profiles(plan)
    if name == "profile":
        return profiles
    instruction = None
    if name in {"instruction", "precleanup", "all"}:
        instruction = validate_instruction(plan)
        if name == "instruction":
            return instruction
    hardware = validate_hardware(plan)
    if name == "hardware":
        return hardware
    variation = validate_variation(analysis)
    if name == "variation":
        return variation
    quality, quality_intervals = validate_quality(plan)
    optional_quality_intervals = validate_optional_quality_receipts(plan, quality["stage"])
    if name == "quality":
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals + optional_quality_intervals,
                     "all quality intervals")
        return quality
    decision = validate_decision(plan, analysis, profiles, quality)
    if name == "decision":
        return {"status": "pass", "analysis": analysis["comparison_sha256"],
                "profiles": profiles["comparison_sha256"], "quality": quality,
                "decision": decision}
    if name == "precleanup":
        inventory = validate_receipt_inventory(plan)
        timeline = validate_receipt_timeline()
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals + optional_quality_intervals,
                     "precleanup retained receipts")
        return {"status": "pass", "source": source, "reuse": reuse,
                "builds": identities, "assembly": assemblies,
                "captures": len(capture_intervals), "analysis": analysis,
                "profiles": profiles, "instruction": instruction,
                "hardware": hardware, "variation": variation,
                "quality": quality, "decision": decision,
                "receipt_inventory": inventory, "timeline": timeline}
    if name == "all":
        inventory = validate_receipt_inventory(plan)
        timeline = validate_receipt_timeline()
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals + optional_quality_intervals,
                     "build/assembly/capture/quality")
        cleanup = validate_cleanup(plan)
        seal = validate_seal()
        return {"status": "pass", "source": source, "reuse": reuse,
                "builds": identities, "assembly": assemblies,
                "captures": len(capture_intervals), "analysis": analysis,
                "profiles": profiles, "instruction": instruction,
                "hardware": hardware,
                "variation": variation, "quality": quality,
                "decision": decision, "receipt_inventory": inventory,
                "timeline": timeline, "cleanup": cleanup, "seal": seal}
    raise VerificationError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    if selected == "precleanup":
        names = ["precleanup"]
    elif selected == "all":
        names = ["all"]
    else:
        names = [selected]
    results: dict[str, Any] = {}
    for name in names:
        try:
            results[name] = {"status": "pass", "result": component(name)}
        except IncompleteError as error:
            results[name] = {"status": "incomplete", "error": str(error)}
        except (VerificationError, OSError, subprocess.CalledProcessError,
                KeyError, TypeError, AttributeError, IndexError, ValueError) as error:
            results[name] = {"status": "fail", "error": str(error)}
    statuses = [item["status"] for item in results.values()]
    status = ("fail" if "fail" in statuses else
              "incomplete" if "incomplete" in statuses else "pass")
    return {"schema": "litchi-0536-matched-verification-v1", "status": status,
            "scope": selected,
            "performance_claim": "candidate admission only; no broad format claim",
            "components": results}


def verify(sealed: bool = False) -> dict[str, Any]:
    result = component("all")
    if not sealed:
        result.pop("seal", None)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=(
        "all", "precleanup", "plan", "source", "quality-reuse", "reuse",
        "build", "builds", "assembly", "captures", "native", "allocation",
        "analysis", "profile", "instruction", "hardware", "variation", "quality", "decision",
        "inventory", "cleanup", "seal",
    ), default="all")
    parser.add_argument("--sealed", action="store_true")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.output and args.output.resolve().is_relative_to(HERE.resolve()):
        raise SystemExit("verification output must be outside the evidence bundle")
    report = run_bundle(args.component)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n",
                               encoding="utf-8")
    print(json.dumps(report, indent=2, sort_keys=True))
    return 2 if args.strict and report["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
