"""Independent custody and replay verifier for the matched 0534 campaign.

The verifier is deliberately read-only with respect to the evidence bundle.
It replays source patches through a private Git index, checks the serial
receipt matrix and ABBA order, delegates raw report semantics to the retained
analyzers in a temporary directory outside this bundle, and validates the
review decision, cleanup record, and final recursive seal.  It never builds,
captures, or runs Rust itself.
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
from typing import Any

# Verification must not leave interpreter caches inside the sealed bundle.
sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
CHECKS = HERE / "checks.py"
FROZEN = HERE / "frozen-inputs.json"
CAPTURE_INPUTS = HERE / "capture-inputs.json"
ADR = HERE / "adr-manifest.json"
BASELINE = HERE / "baseline"
CANDIDATE = HERE / "candidate"
FINAL = HERE / "final"
ANALYZER = HERE / "analyze.py"
ASSEMBLY_ANALYZER = HERE / "analyze_assembly.py"
ASSEMBLY_ANALYSIS = HERE / "assembly-analysis.json"
PROFILE_ANALYZER = HERE / "analyze_profiles.py"
HARDWARE_ANALYZER = HERE / "analyze_hardware.py"
DECIDER = HERE / "decide.py"
COMPARISON = HERE / "comparison.json"
PROFILE_ANALYSIS = "profile-analysis.json"
HARDWARE_ANALYSIS = "hardware-analysis.json"
QUALITY_SUMMARY = HERE / "quality-summary.json"
ADVERSE_REVIEW = HERE / "adverse-review.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
SCRATCH_ROOT = Path("/tmp/litchi-goal-0534")
TARGET = Path("/home/zhuhe/litchi-goal-0534-target")
CANONICAL_BINARY_ROOT = TARGET / "retained-binaries"

SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
STAGES = ("baseline", "candidate", "final")
MEASURED_STAGES = ("baseline", "candidate")
EXPECTED_ENVIRONMENT = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
TIME_FORMAT = '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}'
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
ASSEMBLY_OWNERS = {
    "validate_stream_allocations",
    "validate_physical_sector_layout",
}


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
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        raise VerificationError(f"cannot hash {rel(path)}: {error}") from error
    return value.hexdigest()


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
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise VerificationError(f"{label} timestamp is malformed") from error
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
    require(isinstance(value, dict) and value,
            f"{label or rel(path)} is not a nonempty source manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label or rel(path)} source path")
        require(source_name(name) and valid_digest(digest),
                f"{label or rel(path)} has an invalid source entry: {name}")
        require(name not in result, f"{label or rel(path)} repeats {name}")
        result[name] = digest
    return result


def check_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict) and set(frozen) == {
        "created_utc", "files",
    }, "frozen input envelope differs")
    parse_time(frozen.get("created_utc"), "frozen-inputs.json.created_utc")
    files = frozen.get("files")
    require(isinstance(files, dict) and set(files) == {
        "plan.json", "run.py", "checks.py", "adr-manifest.json",
    }, "frozen input inventory differs")
    for name, path in (("plan.json", PLAN), ("run.py", RUN),
                       ("checks.py", CHECKS), ("adr-manifest.json", ADR)):
        require(files.get(name) == sha(path), f"frozen {name} digest differs")
    for name, expected in files.items():
        require(isinstance(name, str) and Path(name).name == name
                and name != "SHA256SUMS" and valid_digest(expected),
                f"frozen input entry is malformed: {name!r}")
        path = HERE / name
        require(path.is_file() and not path.is_symlink() and sha(path) == expected,
                f"frozen input digest differs: {name}")
    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict)
            and plan.get("status") == "frozen-before-build-capture-and-candidate",
            "plan is not the frozen 0534 plan")
    revision = plan.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError("plan revision is not a Git commit") from error
    require(plan.get("cpu") == 2
            and plan.get("priority") ==
            "OLE2/OOXML active; ODF deferred until that goal completes; iWork excluded"
            and plan.get("owned_paths") == [str(SCRATCH_ROOT), str(TARGET)],
            "plan priority, CPU, or owned paths differ")
    groups = plan.get("groups")
    require(isinstance(groups, dict)
            and groups.get("xls", {}).get("cases") == list(XLS_CASES)
            and groups.get("cfb", {}).get("cases") == [CFB_CASE]
            and groups.get("cfb", {}).get("shapes") == list(CFB_SHAPES)
            and groups.get("cfb", {}).get("payload") == "incompressible",
            "plan source groups differ")
    native = plan.get("native")
    require(isinstance(native, dict)
            and native.get("repeats") == 2 and native.get("warmup") == 20
            and native.get("samples") == 1000
            and native.get("order") == [
                "baseline r1 xls/cfb", "candidate r1 xls/cfb",
                "candidate r2 cfb/xls", "baseline r2 cfb/xls",
            ], "native plan differs")
    allocation = plan.get("allocation")
    require(isinstance(allocation, dict)
            and allocation.get("repeats") == 2 and allocation.get("warmup") == 3
            and allocation.get("samples") == 30
            and "operation-global" in str(allocation.get("scope", "")).lower(),
            "allocation plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict)
            and profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 5
            and profile.get("jobs") == [
                "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large",
            ]
            and profile.get("xls_owner") ==
            "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
            and profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open",
            "profile plan differs")
    hardware = plan.get("hardware")
    require(isinstance(hardware, dict)
            and hardware.get("repeats") == 2
            and hardware.get("case") == XLS_OWNER_CASE
            and hardware.get("warmup") == 0 and hardware.get("samples") == 1000
            and hardware.get("events") ==
            "{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations",
            "hardware plan differs")
    review = plan.get("review")
    require(isinstance(review, dict)
            and review.get("same_build_adverse_percent") == 5
            and review.get("matched_adverse_percent") == 5
            and review.get("primary_cases") == [
                "xls_source_backed_open", "xls_source_backed_open_one_cell",
                "xls_owned_source_open", "xls_owned_source_open_one_cell",
            ]
            and isinstance(review.get("policy"), str)
            and "physical-reconciliation" in review["policy"]
            and "both repeats" in review["policy"], "review plan differs")
    return plan


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
    unique = sorted(set(objects.values()))
    raw_batch = git(["git", "cat-file", "--batch"],
                    input_data=("\n".join(unique) + "\n").encode())
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
    with tempfile.TemporaryDirectory(prefix=".litchi-0534-source-", dir="/home/zhuhe") as directory:
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
    elif stage == "candidate":
        require(changed == {"crates/litchi-cfb/src/file.rs"},
                "candidate source patch does not contain the CFB candidate")
    return {
        "stage": stage, "manifest_sha256": sha(manifest_path),
        "manifest_entries": len(expected), "patch_sha256": sha(patch_path),
        "changed_files": sorted(changed), "new_files": sorted(extras),
    }


def validate_intended_patch_composition(plan: dict[str, Any]) -> dict[str, Any]:
    source = "crates/litchi-cfb/src/file.rs"
    patches = [need(HERE / "candidate.patch"), need(HERE / "tests.patch")]
    results = {}
    stages = ["candidate"] + (["final"] if FINAL.exists() else [])
    for stage in stages:
        selected = patches if stage == "candidate" else patches[1:]
        with tempfile.TemporaryDirectory(prefix=".litchi-0534-compose-", dir="/home/zhuhe") as directory:
            index = Path(directory) / "index"
            env = dict(os.environ, GIT_INDEX_FILE=str(index))
            git(["git", "read-tree", plan["revision"]], env=env)
            git(["git", "apply", "--cached", *[str(path) for path in selected]], env=env)
            changed = set(git(["git", "diff", "--cached", "--name-only", plan["revision"]],
                              env=env).decode().splitlines())
            require(changed == {source}, "intended patches change another source file")
            digest = hashlib.sha256(git(["git", "show", ":" + source], env=env)).hexdigest()
        require(source_manifest(HERE / stage / "source-manifest.json")[source] == digest,
                f"{stage} source is not the intended production/test patch composition")
        results[stage] = digest
    return {"status": "pass", "source_hashes": results,
            "patch_hashes": {path.name: sha(path) for path in patches}}


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
    files = value.get("files") if isinstance(value, dict) else None
    require(isinstance(files, dict) and files, "ADR manifest is malformed")
    for name, expected in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith("docs/adr/") and valid_digest(expected)
                and (REPO / name).is_file() and sha(REPO / name) == expected,
                f"ADR hash differs: {name}")
    require(valid_digest(value.get("prior_audit_source_binding_sha256")),
            "ADR prior audit binding is malformed")
    return {"status": "pass", "entries": len(files), "sha256": sha(ADR)}


def validate_source(plan: dict[str, Any]) -> dict[str, Any]:
    baseline = replay_stage("baseline", plan)
    candidate = replay_stage("candidate", plan)
    require(baseline["manifest_sha256"] != candidate["manifest_sha256"],
            "candidate source manifest is unchanged")
    final = None
    if FINAL.exists():
        final = replay_stage("final", plan)
    manifests = {
        "baseline": source_manifest(BASELINE / "source-manifest.json"),
        "candidate": source_manifest(CANDIDATE / "source-manifest.json"),
    }
    if final is not None:
        manifests["final"] = source_manifest(FINAL / "source-manifest.json")
    selected = None
    decision_path = HERE / "decision.json"
    if decision_path.is_file():
        decision = read_json(decision_path, "decision.json")
        selected = decision.get("final_source")
        if selected is None:
            selected = "candidate" if decision.get("disposition") == "accepted" else "final"
        require(selected in manifests, "decision selects an unretained source stage")
        require(current_source_manifest() == manifests[selected],
                "current checkout does not match the selected source manifest")
    else:
        current = current_source_manifest()
        require(current in manifests.values(),
                "current checkout does not match a retained baseline or candidate source")
    capture_inputs = validate_capture_inputs()
    composition = validate_intended_patch_composition(plan)
    return {"status": "pass", "baseline": baseline, "candidate": candidate,
            "final": final, "adr": validate_adr(), "selected": selected,
            "capture_inputs": capture_inputs, "patch_composition": composition,
            "manifests": {name: sha(HERE / name / "source-manifest.json")
                          for name in manifests}}


def validate_capture_inputs() -> dict[str, Any]:
    """Bind the frozen wrappers used to finish each stage's capture matrix."""
    value = read_json(CAPTURE_INPUTS, "capture-inputs.json")
    require(isinstance(value, dict)
            and value.get("status") == "frozen-before-captures"
            and isinstance(value.get("files"), dict)
            and set(value["files"]) == {
                "capture_baseline.py", "capture_candidate.py",
                "inspect_assembly.py", "candidate.patch",
            }, "capture-inputs envelope differs")
    parse_time(value.get("created_utc"), "capture-inputs.json.created_utc")
    for name, expected in value["files"].items():
        require(valid_digest(expected), f"capture input hash is malformed: {name}")
        path = HERE / name
        require(path.is_file() and not path.is_symlink() and sha(path) == expected,
                f"capture input custody differs: {name}")
    return {"status": "pass", "sha256": sha(CAPTURE_INPUTS),
            "files": dict(value["files"])}


def validate_host(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict)
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
            lane = "assembly"
    result = {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr",
              f"{name}.json", f"{name}.catalog.json"}
    if lane == "native":
        result.add(f"{name}.rss.json")
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
    return result


def validate_artifacts(folder: Path, value: dict[str, Any], expected: set[str],
                       label: str) -> None:
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} artifact inventory differs")
    for filename, expected_digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename
                and valid_digest(expected_digest), f"{label} artifact entry differs")
        # callgrind_annotate output is derived review material.  It is hashed
        # by the recursive bundle seal and referenced by profile analysis, but
        # it is intentionally outside the child receipt's raw-artifact set.
        require(not filename.endswith((".inclusive.txt", ".self.txt")),
                f"{label} receipt includes a derived annotation")
        path = folder / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == expected_digest,
                f"{label} artifact custody differs: {filename}")
        if filename.endswith(".host.json"):
            validate_host(path, f"{label}/{filename}")


def validate_receipt(path: Path, name: str, stage: str, plan: dict[str, Any],
                     command: list[str], expected_binary: str | None,
                     expected_artifact_names: set[str], *, allow_failure: bool = False,
                     expected_execution_stage: str | None = None) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    value = read_json(path, rel(path))
    require(value.get("command") == command, f"{rel(path)} command differs")
    require(isinstance(value.get("exit_code"), int), f"{rel(path)} exit code is invalid")
    require(allow_failure or value["exit_code"] == 0, f"{rel(path)} did not exit successfully")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("script_sha256") == sha(RUN)
            and value.get("source_manifest_sha256") == sha(HERE / stage / "source-manifest.json"),
            f"{rel(path)} plan/script/source binding differs")
    execution = value.get("execution_stage")
    require(execution in STAGES, f"{rel(path)} execution stage is invalid")
    expected_execution_stage = expected_execution_stage or stage
    require(execution == expected_execution_stage,
            f"{rel(path)} execution stage differs from {expected_execution_stage}")
    require(value.get("execution_manifest_sha256") ==
            sha(HERE / execution / "source-manifest.json"),
            f"{rel(path)} execution manifest binding differs")
    require(value.get("binary_sha256") == expected_binary,
            f"{rel(path)} binary binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            f"{rel(path)} instrumentation environment differs")
    start_end = interval(value, rel(path))
    validate_artifacts(path.parent, value, expected_artifact_names, rel(path))
    return value, start_end


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
    require(isinstance(value, dict) and value.get("path") == str(expected_path)
            and valid_digest(value.get("sha256"))
            and isinstance(value.get("bytes"), int) and value["bytes"] > 0
            and value.get("source_manifest_sha256") == sha(folder / "source-manifest.json"),
            f"{rel(path)} identity differs")
    build_path = folder / f"build-{kind}.receipt.json"
    require(value.get("build_receipt_sha256") == sha(build_path),
            f"{rel(path)} build receipt binding differs")
    binary = Path(value["path"])
    if binary.exists():
        require(SCRATCH_ROOT.is_symlink()
                and SCRATCH_ROOT.resolve() == CANONICAL_BINARY_ROOT
                and CANONICAL_BINARY_ROOT.is_dir()
                and not CANONICAL_BINARY_ROOT.is_symlink(),
                "scratch binary alias is not bound to target/retained-binaries")
        require(binary.is_file() and not binary.is_symlink() and sha(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                f"{rel(path)} binary custody differs")
    else:
        cleanup = read_json(CLEANUP, "cleanup.json")
        require(cleanup.get("plan_sha256") == sha(PLAN)
                and cleanup.get("owned_paths_absent") is True
                and cleanup.get("removed") == plan["owned_paths"]
                and all(not os.path.lexists(path) for path in plan["owned_paths"]),
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
        require(value.get("binary_sha256") is None, f"{rel(path)} unexpectedly has a binary hash")
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
        jobs = []
        profile = plan["profile"]
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


def option(command: list[str], name: str) -> str:
    values = [item.split("=", 1)[1] for item in command if item.startswith(name + "=")]
    values += [command[index + 1] for index, item in enumerate(command[:-1]) if item == name]
    require(len(values) == 1, f"capture command option {name} is missing or repeated")
    return values[0]


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
                    "--callgrind-out-file=" + str(folder / f"{job['name']}.callgrind")]
    elif lane == "hardware":
        command += ["perf", "stat", "-x", ",", "-o", str(folder / f"{job['name']}.csv"),
                    "-e", plan["hardware"]["events"], "--"]
    selection = job["selection"]
    command += [str(binary), "--case", ",".join(selection["cases"]), "--warmup",
                str(job["warmup"]), "--samples", str(job["samples"]), "--json",
                str(folder / f"{job['name']}.json"), "--corpus-manifest",
                str(folder / f"{job['name']}.catalog.json")]
    if "shapes" in selection:
        command += ["--shape", ",".join(selection["shapes"]), "--payload", selection["payload"]]
    return command


def validate_capture_command(command: list[str], stage: str, job: dict[str, Any],
                             plan: dict[str, Any], kind: str) -> None:
    require(command == capture_command(stage, job, plan, kind),
            f"{job['name']} capture command differs")
    require(option(command, "--case") == ",".join(job["selection"]["cases"])
            and option(command, "--warmup") == str(job["warmup"])
            and option(command, "--samples") == str(job["samples"]),
            f"{job['name']} capture selectors differ")


def validate_captures(stage: str, plan: dict[str, Any],
                      identities: dict[str, dict[str, Any]]) -> list[tuple[dt.datetime, dt.datetime]]:
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    folder = HERE / stage
    for lane in ("native", "alloc", "profile", "hardware"):
        for job in capture_jobs(plan, lane):
            kind = "alloc" if lane == "alloc" else "normal"
            path = need(folder / f"{job['name']}.receipt.json", f"{stage} {job['name']} receipt")
            expected_execution = "candidate" if stage == "baseline" and job["name"] in {
                "native-r2-xls", "native-r2-cfb",
            } else stage
            value, times = validate_receipt(
                path, job["name"], stage, plan,
                capture_command(stage, job, plan, kind), identities[kind]["sha256"],
                expected_artifacts(folder, job["name"], lane),
                allow_failure=lane == "hardware",
                expected_execution_stage=expected_execution,
            )
            validate_capture_command(value["command"], stage, job, plan, kind)
            intervals.append(times)
    return intervals


def assembly_rows(stage: str, plan: dict[str, Any],
                  identity: dict[str, Any]) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    folder = HERE / stage
    plan_path = need(folder / "assembly-plan.json", f"{stage} assembly plan")
    assembly_plan = read_json(plan_path, rel(plan_path))
    require(isinstance(assembly_plan, dict)
            and assembly_plan.get("schema") == "litchi-0534-assembly-plan-v1"
            and assembly_plan.get("status") == "frozen-before-inspection"
            and assembly_plan.get("owners") == [
                "validate_physical_sector_layout", "validate_stream_allocations",
            ]
            and assembly_plan.get("binary_sha256") == identity["sha256"]
            and assembly_plan.get("source_manifest_sha256") == sha(folder / "source-manifest.json")
            and assembly_plan.get("script_sha256") == sha(HERE / "inspect_assembly.py"),
            f"{rel(plan_path)} envelope differs")
    index_path = need(folder / "assembly-index.json", f"{stage} assembly index")
    index = read_json(index_path, rel(index_path))
    require(isinstance(index, dict)
            and index.get("schema") == "litchi-0534-assembly-index-v1"
            and index.get("plan_sha256") == sha(plan_path),
            f"{rel(index_path)} envelope differs")
    symbols_path = need(folder / "symbols.receipt.json", f"{stage} symbols receipt")
    symbols_value, symbols_interval = validate_receipt(
        symbols_path, "symbols", stage, plan,
        ["nm", "-S", "--defined-only", str(SCRATCH_ROOT / stage / "normal")],
        identity["sha256"], expected_artifacts(folder, "symbols"),
    )
    require(index.get("symbols_receipt_sha256") == sha(symbols_path),
            f"{rel(index_path)} symbols binding differs")
    symbols: dict[str, tuple[int, int]] = {}
    for line in read_text(folder / "symbols.stdout", rel(folder / "symbols.stdout")).splitlines():
        fields = line.split()
        if len(fields) == 4 and fields[2] in ("t", "T"):
            try:
                symbols[fields[3]] = (int(fields[0], 16), int(fields[1], 16))
            except ValueError:
                continue
    rows = index.get("rows")
    require(isinstance(rows, list) and rows, f"{rel(index_path)} rows are missing")
    seen_names: set[str] = set()
    seen_symbols: set[str] = set()
    intervals = [symbols_interval]
    for row in rows:
        require(isinstance(row, dict)
                and set(row) == {"name", "owner", "symbol", "address_hex", "size_bytes", "receipt_sha256"},
                f"{rel(index_path)} row schema differs")
        name, owner, symbol = row["name"], row["owner"], row["symbol"]
        require(isinstance(name, str) and name.startswith("assembly-") and name not in seen_names
                and owner in ASSEMBLY_OWNERS and isinstance(symbol, str) and symbol not in seen_symbols
                and valid_digest(row["receipt_sha256"])
                and re.fullmatch(r"[0-9a-fA-F]+", row["address_hex"]) is not None
                and isinstance(row["size_bytes"], int) and row["size_bytes"] > 0,
                f"{rel(index_path)} row identity differs")
        require(symbol in symbols and symbols[symbol] ==
                (int(row["address_hex"], 16), row["size_bytes"]),
                f"{rel(index_path)} symbol is not bound by nm output: {symbol}")
        receipt_path = need(folder / f"{name}.receipt.json", f"{stage} {name} receipt")
        value, times = validate_receipt(
            receipt_path, name, stage, plan,
            ["objdump", "-d", "--disassemble=" + symbol, str(SCRATCH_ROOT / stage / "normal")],
            identity["sha256"], expected_artifacts(folder, name, "assembly"),
        )
        require(row["receipt_sha256"] == sha(receipt_path), f"{rel(receipt_path)} hash differs")
        seen_names.add(name)
        seen_symbols.add(symbol)
        intervals.append(times)
    require({row["owner"] for row in rows} >= {
        "validate_stream_allocations", "validate_physical_sector_layout",
    }, f"{rel(index_path)} omits required CFB attribution owners")
    return ({"status": "pass", "plan_sha256": sha(plan_path),
             "index_sha256": sha(index_path), "symbols": len(symbols),
             "disassemblies": len(rows)}, intervals)


def validate_assembly_analysis() -> dict[str, Any]:
    """Replay the stage-paired static caller/symbol analysis exactly."""
    value = replay_report(ASSEMBLY_ANALYZER, [], ASSEMBLY_ANALYSIS)
    stages = value.get("stages") if isinstance(value, dict) else None
    require(isinstance(value, dict) and value.get("status") == "pass"
            and isinstance(stages, dict)
            and set(stages) == set(MEASURED_STAGES)
            and "static code" in str(value.get("scope", "")),
            "assembly analysis envelope differs")
    for stage in MEASURED_STAGES:
        row = stages[stage]
        require(isinstance(row, dict) and valid_digest(row.get("binary_sha256"))
                and isinstance(row.get("binary_bytes"), int) and row["binary_bytes"] > 0
                and row.get("source_manifest_sha256") ==
                sha(HERE / stage / "source-manifest.json")
                and row.get("assembly_index_sha256") ==
                sha(HERE / stage / "assembly-index.json")
                and isinstance(row.get("rows"), list)
                and isinstance(row.get("owners"), dict),
                f"{stage} assembly analysis row differs")
    return {"status": "pass", "sha256": sha(ASSEMBLY_ANALYSIS),
            "stages": list(MEASURED_STAGES)}


def check_serial(intervals: list[tuple[dt.datetime, dt.datetime]], label: str) -> None:
    ordered = sorted(intervals, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{label} receipt intervals overlap")


def replay_report(script: Path, args: list[str], retained: Path) -> Any:
    need(script, rel(script))
    need(retained, rel(retained))
    if script == PROFILE_ANALYZER:
        # The retained profile analyzer can materialize callgrind annotations
        # when they are absent.  Refuse that input before launching it so this
        # verifier remains read-only even before the recursive seal exists.
        for stage in MEASURED_STAGES:
            for dump in sorted((HERE / stage).glob("profile-*.callgrind.*")):
                suffix = dump.name.rsplit(".", 1)[-1]
                if not suffix.isdigit() or ("-cfb-" in dump.name and suffix == "1"):
                    continue
                stem = dump.name.rsplit(".callgrind.", 1)[0]
                for kind in ("inclusive", "self"):
                    annotation = HERE / stage / f"{stem}.part-{suffix}.{kind}.txt"
                    need(annotation, rel(annotation))
    # Report replay must be observational.  In particular, the profile
    # analyzer can generate a missing callgrind annotation as a convenience;
    # a complete sealed bundle must already contain every derived annotation.
    # Capture file identity metadata and reject any analyzer-side bundle write.
    before = {
        path: (path.stat().st_size, path.stat().st_mtime_ns)
        for path in HERE.rglob("*") if path.is_file() and not path.is_symlink()
    }
    with tempfile.TemporaryDirectory(prefix=".litchi-0534-report-", dir="/home/zhuhe") as directory:
        output = Path(directory) / retained.name
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
        try:
            result = subprocess.run([sys.executable, "-B", str(script), *args, str(output)],
                                    cwd=REPO, env=environment, stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, text=True, check=False)
        except OSError as error:
            raise VerificationError(f"{rel(script)} replay could not start: {error}") from error
        require(result.returncode == 0,
                f"{rel(script)} replay failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == retained.read_bytes(),
                f"{rel(script)} replay differs from {rel(retained)}")
    after = {
        path: (path.stat().st_size, path.stat().st_mtime_ns)
        for path in HERE.rglob("*") if path.is_file() and not path.is_symlink()
    }
    require(before == after, f"{rel(script)} replay modified the evidence bundle")
    return read_json(retained, rel(retained))


def validate_analysis(plan: dict[str, Any]) -> dict[str, Any]:
    baseline = replay_report(ANALYZER, ["--stage", "baseline"], BASELINE / "analysis.json")
    candidate = replay_report(ANALYZER, ["--stage", "candidate"], CANDIDATE / "analysis.json")
    for stage, value in (("baseline", baseline), ("candidate", candidate)):
        require(isinstance(value, dict) and value.get("status") == "pass"
                and value.get("stage") == stage and value.get("scope") == plan["scope"]
                and value.get("priority") == plan["priority"]
                and value.get("plan_sha256") == sha(PLAN)
                and value.get("native_samples") == 24000
                and value.get("allocation_samples") == 720
                and len(value.get("rows", [])) == 48,
                f"{stage} numeric analysis envelope differs")
    comparison = replay_report(ANALYZER, ["--compare"], COMPARISON)
    require(isinstance(comparison, dict) and comparison.get("status") == "pass"
            and comparison.get("schema") == "cfb_ole2_matched_comparison_v1"
            and comparison.get("plan_sha256") == sha(PLAN)
            and comparison.get("baseline", {}).get("stage") == "baseline"
            and comparison.get("candidate", {}).get("stage") == "candidate",
            "comparison envelope differs")
    return {"status": "pass", "baseline": baseline, "candidate": candidate,
            "comparison": comparison, "analysis_sha256": {
                "baseline": sha(BASELINE / "analysis.json"),
                "candidate": sha(CANDIDATE / "analysis.json"),
            }, "comparison_sha256": sha(COMPARISON)}


def validate_profiles(plan: dict[str, Any]) -> dict[str, Any]:
    values = {}
    for stage in MEASURED_STAGES:
        retained = HERE / stage / PROFILE_ANALYSIS
        value = replay_report(PROFILE_ANALYZER, ["--stage", stage], retained)
        require(isinstance(value, dict)
                and value.get("schema") == "cfb_ole2_constructor_callgrind_profile_analysis_v2"
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
    require(comparison.get("status") == "pass"
            and comparison.get("stage_selection") == list(MEASURED_STAGES)
            and comparison.get("comparison", {}).get("available") is True,
            "matched profile comparison envelope differs")
    constructor_ir, constructor_gate = profile_ir_gate(values)
    physical_ir, physical_gate = physical_profile_gate(values)
    return {"status": "pass", "profiles": values,
            "comparison_sha256": sha(comparison_path),
            "sha256": {stage: sha(HERE / stage / PROFILE_ANALYSIS)
                       for stage in MEASURED_STAGES},
            "constructor_ir": constructor_ir,
            "profile_gate": constructor_gate,
            "physical_ir": physical_ir,
            "physical_profile_gate": physical_gate}


def validate_hardware(plan: dict[str, Any]) -> dict[str, Any]:
    values = {}
    for stage in MEASURED_STAGES:
        retained = HERE / stage / HARDWARE_ANALYSIS
        value = replay_report(HARDWARE_ANALYZER, ["--stage", stage], retained)
        require(isinstance(value, dict) and value.get("status") == "pass"
                and value.get("scope") == plan["hardware"]["scope"]
                and len(value.get("captures", [])) == 2
                and value.get("latency_samples_excluded_from_native") == 2000
                and value.get("no_operation_local_hardware_or_speedup_claim") is True,
                f"{stage} hardware analysis envelope differs")
        for row in value["captures"]:
            require(isinstance(row, dict) and row.get("name", "").startswith("hardware-r")
                    and row.get("status") in {
                        "measured", "unavailable", "unavailable_for_group_claim",
                    }, f"{stage} hardware analysis row differs")
        values[stage] = value
    return {"status": "pass", "hardware": values,
            "sha256": {stage: sha(HERE / stage / HARDWARE_ANALYSIS)
                       for stage in MEASURED_STAGES}}


def validate_variation(analysis: dict[str, Any]) -> dict[str, Any]:
    path = need(HERE / "variation-review.json", "variation-review.json")
    value = read_json(path, rel(path))
    raw = []
    for stage in MEASURED_STAGES:
        report = analysis[stage]
        raw.extend(report.get("same_build_variations_over_five_percent", []))
    reviewed = value.get("variations") if isinstance(value, dict) else None
    require(isinstance(reviewed, list)
            and valid_digest(value.get("analysis_sha256"))
            and value.get("analysis_sha256") in {
                analysis["analysis_sha256"]["baseline"],
                analysis["analysis_sha256"]["candidate"],
                analysis.get("comparison_sha256"),
            }
            and value.get("threshold_percent") == 5
            and isinstance(value.get("review"), str) and value["review"].strip(),
            "variation review envelope differs")
    # A review may intentionally use the comparison hash or a combined hash;
    # require every raw row to be present and individually explained.
    remaining = list(reviewed)
    for item in raw:
        require(isinstance(item, dict), "raw variation row is malformed")
        found = next((index for index, candidate in enumerate(remaining)
                      if isinstance(candidate, dict)
                      and all(candidate.get(key) == value for key, value in item.items())), None)
        require(found is not None, "same-build variation is not individually reviewed")
        candidate = remaining.pop(found)
        explanation = candidate.get("review", candidate.get("reason"))
        require(isinstance(explanation, str) and explanation.strip(),
                "same-build variation review lacks an explanation")
    require(not remaining and value.get("complete", True) is True,
            "same-build variation review is incomplete")
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
        result.append((row[0], row[1]))
    require([name for name, _ in result] == list(QUALITY_NAMES),
            "checks.py quality command inventory differs")
    return result


def quality_prefix() -> list[str]:
    return ["env", "TMPDIR=" + str(TARGET / "test-tmp"),
            "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
            "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings"]


def test_count(stdout: str) -> int:
    return sum(int(item) for item in re.findall(r"test result: ok\. (\d+) passed;", stdout))


def validate_quality(plan: dict[str, Any]) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    summary = read_json(QUALITY_SUMMARY, "quality-summary.json")
    require(isinstance(summary, dict) and summary.get("status") == "pass"
            and summary.get("stage") in ("candidate", "final"),
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
        require(name in expected_names and name not in seen and valid_digest(row["receipt_sha256"]),
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
        count = test_count(read_text(path.with_name(path.name.replace(".receipt.json", ".stdout"))))
        require(isinstance(row["executed_tests"], int) and row["executed_tests"] == count,
                f"{rel(path)} test count differs")
        intervals.append(times)
        total += count
    require(seen == expected_names and summary.get("executed_tests") == total,
            "quality summary aggregate differs")
    return ({"status": "pass", "stage": stage, "checks": len(rows),
             "executed_tests": total, "sha256": sha(QUALITY_SUMMARY)}, intervals)


def validate_optional_quality_receipts(plan: dict[str, Any]) -> list[tuple[dt.datetime, dt.datetime]]:
    """Validate a complete retained quality run in a stage not selected by the summary."""
    commands = dict(quality_commands())
    selected_stage = read_json(QUALITY_SUMMARY)["stage"]
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in ("baseline", "candidate", "final"):
        if stage == selected_stage:
            continue
        folder = HERE / stage
        found = sorted(folder.glob("check-*.receipt.json")) if folder.is_dir() else []
        if not found:
            continue
        require({path.name for path in found} == {
            f"check-{name}.receipt.json" for name in commands
        }, f"{stage} quality receipt inventory differs")
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
    index = read_json(HERE / stage / "assembly-index.json", rel(HERE / stage / "assembly-index.json"))
    rows = index.get("rows") if isinstance(index, dict) else None
    require(isinstance(rows, list), f"{stage} assembly rows are missing")
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("name"), str),
                f"{stage} assembly receipt name is malformed")
        names.add(row["name"])
    return names


def validate_receipt_inventory(plan: dict[str, Any]) -> dict[str, Any]:
    expected: dict[str, set[str]] = {}
    for stage in ("baseline", "candidate"):
        core = expected_core_receipts(stage, plan)
        check_names = {f"check-{name}" for name in QUALITY_NAMES}
        found_checks = {
            path.name.removesuffix(".receipt.json")
            for path in (HERE / stage).glob("check-*.receipt.json")
        }
        require(not found_checks or found_checks == check_names,
                f"{stage} quality receipt inventory is partial")
        expected[stage] = core | found_checks
    if FINAL.is_dir():
        actual_final = {path.name.removesuffix(".receipt.json")
                        for path in FINAL.glob("*.receipt.json")}
        if actual_final:
            core = expected_core_receipts("final", plan) if (FINAL / "assembly-index.json").is_file() else set()
            check_names = {f"check-{name}" for name in QUALITY_NAMES}
            require(actual_final in (check_names, core | check_names),
                    "final receipt inventory differs")
            expected["final"] = actual_final
    for stage, names in expected.items():
        actual = {path.name.removesuffix(".receipt.json")
                  for path in (HERE / stage).glob("*.receipt.json")}
        require(actual == names,
                f"{stage} receipt inventory differs: expected {len(names)}, found {len(actual)}")
    count = sum(len(names) for names in expected.values())
    return {"status": "pass", "stages": {stage: len(names) for stage, names in expected.items()},
            "successful_receipts": count,
            "receipt_names_sha256": hashlib.sha256(
                ("\n".join(f"{stage}/{name}" for stage in sorted(expected)
                           for name in sorted(expected[stage])) + "\n").encode()).hexdigest()}


def validate_receipt_timeline() -> dict[str, Any]:
    rows: list[tuple[dt.datetime, dt.datetime, str]] = []
    for stage in STAGES:
        folder = HERE / stage
        if not folder.is_dir():
            continue
        for path in folder.glob("*.receipt.json"):
            value = read_json(path, rel(path))
            start, end = interval(value, rel(path))
            rows.append((start, end, f"{stage}/{path.name.removesuffix('.receipt.json')}"))
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
    """Replay the independent XLS-owned constructor-Ir admission gate."""
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


def _physical_ir(profile: dict[str, Any], label: str) -> int:
    """Sum physical-reconciliation self Ir across the five timed dumps."""
    attribution = profile.get("constructor_attribution")
    require(isinstance(attribution, list) and attribution,
            f"{label} physical attribution is missing")
    total = 0
    for item in attribution:
        require(isinstance(item, dict) and isinstance(item.get("functions"), dict),
                f"{label} physical attribution row is malformed")
        target = item["functions"].get("physical_reconciliation")
        require(isinstance(target, dict)
                and target.get("target") ==
                "litchi_cfb::file::OleFile<R>::validate_physical_sector_layout"
                and target.get("out_of_line") is True
                and target.get("inlined_or_absent") is False
                and isinstance(target.get("incoming_edge_count"), int)
                and target["incoming_edge_count"] > 0,
                f"{label} physical reconciliation is not a positive out-of-line target")
        value = target.get("self_ir")
        require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
                f"{label} physical reconciliation self Ir is malformed")
        total += value
    return total


def physical_profile_gate(profiles: dict[str, Any]) -> tuple[list[dict[str, Any]], bool]:
    """Replay the paired physical-reconciliation self-Ir gate.

    The gate deliberately covers the XLS-owned constructor and the CFB
    few-large shape for both repeats.  Other profile shapes remain diagnostic
    evidence and cannot satisfy this admission requirement by substitution.
    """
    rows: list[dict[str, Any]] = []
    selected = (("xls-owned", None), ("cfb-few-large", "few-large"))
    for repeat in (1, 2):
        for group, shape in selected:
            values = {
                stage: _physical_ir(
                    _profile_row(profiles, stage, group, repeat, shape),
                    f"{stage} {group} repeat {repeat}",
                )
                for stage in MEASURED_STAGES
            }
            require(values["baseline"] > 0,
                    f"baseline {group} repeat {repeat} physical Ir is not positive")
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
                      and all(candidate.get(key) == value for key, value in item.items())), None)
        require(index is not None, f"{label} row is not individually reviewed")
        candidate = remaining.pop(index)
        explanation = candidate.get("review", candidate.get("reason"))
        require(isinstance(explanation, str) and explanation.strip(),
                f"{label} row lacks an explanation")
    require(not remaining, f"{label} review has extra rows")


def load_decider() -> Any:
    need(DECIDER, "decide.py")
    old_path = list(sys.path)
    old_run = sys.modules.get("run")
    try:
        sys.path.insert(0, str(HERE))
        run_spec = importlib.util.spec_from_file_location("litchi_0534_decide_run", RUN)
        require(run_spec is not None and run_spec.loader is not None, "cannot load run.py")
        run_module = importlib.util.module_from_spec(run_spec)
        sys.modules["run"] = run_module
        run_spec.loader.exec_module(run_module)
        spec = importlib.util.spec_from_file_location("litchi_0534_decider", DECIDER)
        require(spec is not None and spec.loader is not None, "cannot load decide.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    except (OSError, ValueError, KeyError, TypeError, AttributeError, AssertionError) as error:
        raise VerificationError(f"decision replay setup failed: {error}") from error
    finally:
        sys.path[:] = old_path
        if old_run is None:
            sys.modules.pop("run", None)
        else:
            sys.modules["run"] = old_run


def validate_decision(plan: dict[str, Any], analysis: dict[str, Any],
                      profiles: dict[str, Any], quality: dict[str, Any]) -> dict[str, Any]:
    path = need(HERE / "decision.json", "decision.json")
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and value.get("disposition") in {"accepted", "rejected"}
            and isinstance(value.get("native_primary_gate"), bool)
            and isinstance(value.get("memory_gate"), bool)
            and isinstance(value.get("profile_gate"), bool)
            and isinstance(value.get("physical_profile_gate"), bool)
            and isinstance(value.get("physical_ir"), list)
            and value.get("adverse_review_complete") is True
            and value.get("quality_gates") == 14
            and value.get("comparison_sha256") == sha(COMPARISON)
            and value.get("profiles_sha256") == {
                stage: sha(HERE / stage / PROFILE_ANALYSIS) for stage in MEASURED_STAGES
            }
            and value.get("quality_summary_sha256") == sha(QUALITY_SUMMARY)
            and value.get("adverse_review_sha256") == sha(ADVERSE_REVIEW),
            "decision envelope differs")
    comparison = analysis["comparison"]["comparison"]
    primary = comparison["admission"]["primary_workflow_p50"]["all_four_cases_both_repeats_pass"]
    memory = comparison["admission"]["allocation_guard"]["passes"]
    constructor_ir, profile_gate = profile_ir_gate(profiles["profiles"])
    physical_ir, physical_gate = physical_profile_gate(profiles["profiles"])
    require(value["native_primary_gate"] is primary and value["memory_gate"] is memory
            and value["profile_gate"] is profile_gate
            and value["constructor_ir"] == constructor_ir
            and value["physical_profile_gate"] is physical_gate
            and value["physical_ir"] == physical_ir,
            "decision gate values differ from independent replay")
    review = read_json(ADVERSE_REVIEW, "adverse-review.json")
    require(isinstance(review, dict) and review.get("comparison_sha256") == sha(COMPARISON)
            and review.get("complete") is True and isinstance(review.get("adoption_allowed"), bool),
            "adverse review envelope differs")
    review_rows(comparison["matched_adverse_flags_over_five_percent"], review.get("matched"), "matched")
    review_rows(comparison["same_build_variations_over_five_percent"], review.get("same_build"), "same-build")
    expected_disposition = (
        "accepted" if primary and memory and profile_gate and physical_gate
        and review["adoption_allowed"] else "rejected"
    )
    require(value["disposition"] == expected_disposition,
            "decision disposition differs from independent gates")
    expected = load_decider().evaluate()
    require(value == expected, "decision does not replay from decide.py")
    selected = value.get("final_source")
    if selected is None:
        selected = "candidate" if value["disposition"] == "accepted" else "final"
    require(selected in MEASURED_STAGES or selected == "final",
            "decision final source is invalid")
    manifest_path = HERE / selected / "source-manifest.json"
    require(manifest_path.is_file(), "decision final source manifest is missing")
    require(current_source_manifest() == source_manifest(manifest_path),
            "decision current source differs from final source")
    if value["disposition"] == "rejected":
        require(selected == "final" and quality["stage"] == "final",
                "rejected decision does not retain final restored quality")
    else:
        require(selected in ("candidate", "final") and quality["stage"] == selected,
                "accepted decision does not retain candidate-equivalent quality")
    return {"status": "pass", "disposition": value["disposition"],
            "final_source": selected, "sha256": sha(path),
            "constructor_ir": constructor_ir, "physical_ir": physical_ir,
            "physical_profile_gate": physical_gate}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict)
            and valid_digest(value.get("plan_sha256"))
            and value.get("plan_sha256") == sha(PLAN)
            and isinstance(value.get("removed"), list)
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
    return {"status": "pass", "sha256": sha(CLEANUP), "owned_paths_absent": True,
            "python_cache_absent": True}


def validate_seal() -> dict[str, Any]:
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), "SHA256SUMS line differs")
        safe_relative(fields[1], "SHA256SUMS path")
        require(fields[1] != "SHA256SUMS" and fields[1] not in expected,
                "SHA256SUMS inventory is unsafe or duplicated")
        expected[fields[1]] = fields[0]
    actual = {rel(path): sha(path) for path in HERE.rglob("*")
              if path.is_file() and not path.is_symlink() and path != SEAL}
    require(expected == actual and not any(path.is_symlink() for path in HERE.rglob("*")),
            "SHA256SUMS inventory differs")
    return {"status": "pass", "sha256": sha(SEAL), "entries": len(expected)}


def component(name: str) -> dict[str, Any]:
    plan = check_plan()
    source = validate_source(plan)
    if name == "source":
        return source
    identities: dict[str, dict[str, dict[str, Any]]] = {}
    build_intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in MEASURED_STAGES:
        identities[stage], times = validate_builds(stage, plan)
        build_intervals.extend(times)
    if name in {"build", "builds"}:
        check_serial(build_intervals, "build")
        return {"status": "pass", "builds": identities, "intervals": len(build_intervals)}
    assemblies: dict[str, Any] = {}
    assembly_intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in MEASURED_STAGES:
        assemblies[stage], times = assembly_rows(stage, plan, identities[stage]["normal"])
        assembly_intervals.extend(times)
    assembly_diagnostic = validate_assembly_analysis()
    if name == "assembly":
        check_serial(build_intervals + assembly_intervals, "build/assembly")
        return {"status": "pass", "builds": identities, "assembly": assemblies,
                "assembly_analysis": assembly_diagnostic,
                "intervals": len(build_intervals) + len(assembly_intervals)}
    capture_intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in MEASURED_STAGES:
        capture_intervals.extend(validate_captures(stage, plan, identities[stage]))
    if name in {"captures", "native", "allocation"}:
        check_serial(build_intervals + assembly_intervals + capture_intervals,
                     "build/assembly/capture")
        return {"status": "pass", "builds": identities, "assembly": assemblies,
                "assembly_analysis": assembly_diagnostic,
                "captures": len(capture_intervals), "intervals": len(build_intervals) +
                len(assembly_intervals) + len(capture_intervals)}
    analysis = validate_analysis(plan)
    if name == "analysis":
        return analysis
    profiles = validate_profiles(plan)
    if name == "profile":
        return profiles
    hardware = validate_hardware(plan)
    if name == "hardware":
        return hardware
    variation = validate_variation(analysis)
    if name == "variation":
        return variation
    quality, quality_intervals = validate_quality(plan)
    optional_quality_intervals = validate_optional_quality_receipts(plan)
    if name == "quality":
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals + optional_quality_intervals, "all quality intervals")
        return quality
    decision = validate_decision(plan, analysis, profiles, quality)
    if name == "decision":
        return {"status": "pass", "analysis": analysis["comparison_sha256"],
                "profiles": profiles["sha256"], "quality": quality,
                "decision": decision}
    if name == "cleanup":
        return validate_cleanup(plan)
    if name == "seal":
        return validate_seal()
    if name == "all":
        inventory = validate_receipt_inventory(plan)
        timeline = validate_receipt_timeline()
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals + optional_quality_intervals,
                     "build/assembly/capture/quality")
        cleanup = validate_cleanup(plan)
        seal = validate_seal()
        return {"status": "pass", "source": source, "builds": identities,
                "assembly": assemblies, "assembly_analysis": assembly_diagnostic,
                "captures": len(capture_intervals),
                "analysis": analysis, "profiles": profiles, "hardware": hardware,
                "variation": variation, "quality": quality, "decision": decision,
                "receipt_inventory": inventory, "timeline": timeline,
                "cleanup": cleanup, "seal": seal}
    raise VerificationError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    if selected == "precleanup":
        names = ["source", "builds", "assembly", "captures", "analysis", "profile",
                 "hardware", "variation", "quality", "decision"]
    elif selected == "all":
        names = ["all"]
    else:
        names = [selected]
    results: dict[str, Any] = {}
    for name in names:
        try:
            value = component(name)
            results[name] = {"status": "pass", "result": value}
        except IncompleteError as error:
            results[name] = {"status": "incomplete", "error": str(error)}
        except (VerificationError, OSError, subprocess.CalledProcessError,
                KeyError, TypeError, AttributeError, IndexError) as error:
            results[name] = {"status": "fail", "error": str(error)}
    statuses = [item["status"] for item in results.values()]
    status = ("fail" if "fail" in statuses else
              "incomplete" if "incomplete" in statuses else "pass")
    return {"schema": "litchi-0534-matched-verification-v1", "status": status,
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
        "all", "precleanup", "source", "build", "builds", "assembly", "captures",
        "native", "allocation", "analysis", "profile", "hardware", "variation",
        "quality", "decision", "cleanup", "seal",
    ), default="all")
    parser.add_argument("--sealed", action="store_true")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.output and args.output.resolve().is_relative_to(HERE.resolve()):
        raise SystemExit("verification output must be outside the evidence bundle")
    report = run_bundle(args.component)
    if args.sealed and report["status"] == "pass" and args.component == "all":
        # component('all') already validates the seal; this flag documents the
        # caller's intent and keeps the CLI compatible with earlier bundles.
        pass
    text = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
    print(text, end="")
    return 2 if args.strict and report["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
