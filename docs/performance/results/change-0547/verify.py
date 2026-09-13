#!/usr/bin/env python3
"""Fail-closed custody verifier for the baseline-only 0547 OLE2 campaign.

The campaign deliberately has one measured source stage.  This verifier
checks that the stage is the frozen Git tree, that the release binary and all
Callgrind/assembly receipts refer to that tree, and that the retained
diagnostic analyzers reproduce their reports from immutable raw evidence.  It
does not build, benchmark, profile, or clean anything.  ``--precleanup``
checks the complete measured bundle before the owned target is removed;
``--strict`` additionally requires cleanup and the recursive seal.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
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
CAPTURE = HERE / "capture.py"
INSPECT = HERE / "inspect_assembly.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
BASELINE = HERE / "baseline"
INSTRUCTION_ANALYZER = HERE / "instruction_analysis.py"
NUMERIC_ANALYZER = HERE / "analyze.py"
PROFILE_ANALYZER = HERE / "analyze_profiles.py"
PROOF_SCRIPT = HERE / "proof_check.py"
SUBOPERATION_SCRIPT = HERE / "suboperations.py"
QUALITY_REUSE = HERE / "quality-reuse.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
PRIOR = HERE.parent / "change-0546" / "integration"
TARGET = Path("/home/zhuhe/litchi-goal-0547-target")

SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
RECEIPT_KEYS = {
    "command", "start_utc", "end_utc", "seconds", "exit_code",
    "execution_stage", "execution_manifest_sha256", "binary_sha256",
    "source_manifest_sha256", "script_sha256", "plan_sha256", "environment",
    "artifacts",
}
ENVIRONMENT_KEYS = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
PROFILE_OWNER_XLS = (
    "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
)
PROFILE_OWNER_CFB = "litchi_cfb::file::OleFile<R>::open"
ASSEMBLY_OWNERS = ("SectorChainScratch", "CheckedBitSet", "chain_error")
ASSEMBLY_REQUIRED = ("collect_exact", "insert")


class VerificationError(ValueError):
    """Malformed, contradictory, or out-of-scope retained evidence."""


class IncompleteError(VerificationError):
    """Evidence needed by the selected verification mode is not present."""


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
    if directory:
        if not path.is_dir() or path.is_symlink():
            raise IncompleteError(f"{label} is missing or not a regular directory")
    elif not path.is_file() or path.is_symlink():
        raise IncompleteError(f"{label} is missing or not a regular file")
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
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdef" for character in value
    )


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
    require(result.tzinfo is not None, f"{label} timestamp has no timezone")
    return result


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and float(seconds) > 0.0,
            f"{label}.seconds is not positive")
    require(end > start, f"{label} interval is inverted")
    return start, end


def git(args: list[str], *, env: dict[str, str] | None = None,
        input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, env=env, input=input_data,
                                       stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise VerificationError(
            f"Git command failed ({' '.join(args)}): {detail.decode(errors='replace')[-2000:]}"
        ) from error


def source_name(name: str) -> bool:
    return name in SOURCE_EXACT or any(name.startswith(prefix) for prefix in SOURCE_PREFIXES)


def source_manifest(path: Path, label: str | None = None) -> dict[str, str]:
    label = label or rel(path)
    value = read_json(path, label)
    require(isinstance(value, dict) and value, f"{label} is not a nonempty source manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{label} source path")
        require(source_name(name) and valid_digest(digest),
                f"{label} source entry is invalid: {name}")
        require(name not in result, f"{label} repeats {name}")
        result[name] = digest
    return result


def tree_manifest(revision: str) -> dict[str, str]:
    """Return SHA-256 content hashes for the source tree at a Git revision."""
    raw = git(["git", "ls-tree", "-r", "-z", revision])
    objects: dict[str, str] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, name_bytes = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "Git source tree entry is malformed")
        name = name_bytes.decode()
        if source_name(name):
            objects[name] = fields[2].decode()
    require(objects, "Git source tree has no source objects")
    raw_batch = git(["git", "cat-file", "--batch"],
                    input_data=("\n".join(sorted(set(objects.values()))) + "\n").encode())
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


def current_source_manifest() -> dict[str, str]:
    tracked = git(["git", "ls-files", "-z", "crates", "tools/perf-baseline",
                   "Cargo.toml", "Cargo.lock", ".cargo", "rust-toolchain.toml"])
    names = {item.decode() for item in tracked.split(b"\0") if item}
    untracked = git(["git", "ls-files", "--others", "--exclude-standard", "-z", "--",
                     "crates", "tools/perf-baseline"])
    names.update(item.decode() for item in untracked.split(b"\0")
                 if item and item.endswith(b".rs"))
    return {name: sha(REPO / name) for name in sorted(names)
            if source_name(name) and (REPO / name).is_file()
            and not (REPO / name).is_symlink()}


def check_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict)
            and set(frozen) == {"utc", "files", "scope"},
            "frozen-inputs envelope differs")
    parse_time(frozen.get("utc"), "frozen-inputs.utc")
    require(frozen.get("scope") ==
            "Capture inputs frozen; analysis consumes immutable raw output and cannot alter measured sources or captures.",
            "frozen-inputs scope differs")
    frozen_files = frozen.get("files")
    expected_frozen = {"run.py", "inspect_assembly.py", "capture.py", "plan.json", "adr-manifest.json"}
    require(isinstance(frozen_files, dict) and set(frozen_files) == expected_frozen,
            "frozen-inputs file inventory differs")
    for name, digest in frozen_files.items():
        require(valid_digest(digest), f"frozen input digest is invalid: {name}")
        require(sha(HERE / name) == digest, f"frozen input digest differs: {name}")

    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("status") == "frozen-before-build-and-capture",
            "plan is not the frozen 0547 plan")
    require(plan.get("scope") == "Fresh baseline-only OLE2 collector sub-operation attribution",
            "plan scope differs")
    require(plan.get("priority") ==
            "OLE2/OOXML active; ODF deferred until that goal completes; iWork excluded",
            "plan priority differs")
    require(plan.get("cpu") == 2, "plan CPU differs")
    revision = plan.get("revision")
    require(isinstance(revision, str) and len(revision) == 40
            and all(c in "0123456789abcdef" for c in revision),
            "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError("plan revision is not a Git commit") from error
    require(plan.get("owned_paths") == [str(TARGET)], "plan owned paths differ")
    groups = plan.get("groups")
    expected_xls = [
        "xls_semantic_open", "xls_eager_open_list_worksheets", "xls_eager_open_one_cell",
        "xls_source_backed_open", "xls_source_backed_open_list_worksheets",
        "xls_source_backed_open_one_cell", "xls_owned_source_open",
        "xls_owned_source_open_list_worksheets", "xls_owned_source_open_one_cell",
    ]
    require(isinstance(groups, dict)
            and groups.get("xls", {}).get("cases") == expected_xls
            and groups.get("cfb", {}).get("cases") == ["cfb_open"]
            and groups.get("cfb", {}).get("shapes") == ["tiny", "many-small", "few-large"]
            and groups.get("cfb", {}).get("payload") == "incompressible",
            "plan case matrix differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict)
            and profile.get("repeats") == 1 and profile.get("warmup") == 0
            and profile.get("samples") == 5
            and profile.get("jobs") == ["xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large"]
            and profile.get("xls_owner") == PROFILE_OWNER_XLS
            and profile.get("cfb_owner") == PROFILE_OWNER_CFB
            and profile.get("scope") == (
                "zero-before/dump-after/toggle exact constructor; retain setup CFB open "
                "dump and use only positive incoming benchmark-runner dumps for operation attribution"
            )
            and profile.get("instruction_flags") == [
                "--dump-instr=yes", "--dump-line=no", "--compress-pos=no", "--collect-jumps=yes",
            ], "plan profile lane differs")
    assembly = plan.get("assembly")
    require(isinstance(assembly, dict)
            and assembly.get("owners") == list(ASSEMBLY_OWNERS)
            and assembly.get("required") == list(ASSEMBLY_REQUIRED)
            and assembly.get("scope") == "Same measured binary; all matching symbols captured; instruction mapping, not static latency",
            "plan assembly lane differs")
    require(plan.get("performance_claim") ==
            "none: single-repeat baseline instruction attribution, not candidate comparison or native latency",
            "plan performance claim differs")
    reuse = plan.get("quality_reuse")
    require(isinstance(reuse, dict)
            and reuse.get("bundle") == "change-0546/integration"
            and reuse.get("stage") == "candidate"
            and isinstance(reuse.get("condition"), str)
            and "Exact complete source-manifest equality" in reuse["condition"],
            "plan quality reuse differs")
    require(plan.get("temporary_storage") == (
        "Use owned target/tmp from first child; prior0532 confirmed /tmp EDQUOT; vgdb=no from first profile"
    ), "plan temporary storage differs")
    return plan


def validate_adr() -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    require(isinstance(value, dict) and set(value) == {"files", "checked_utc", "status"},
            "ADR manifest envelope differs")
    parse_time(value.get("checked_utc"), "adr-manifest.checked_utc")
    require(value.get("status") == "all previously read ADRs unchanged",
            "ADR manifest status differs")
    files = value.get("files")
    require(isinstance(files, dict) and len(files) == 30, "ADR file inventory differs")
    for name, digest in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith("docs/adr/") and valid_digest(digest)
                and (REPO / name).is_file() and sha(REPO / name) == digest,
                f"ADR hash differs: {name}")
    prior = PRIOR / "adr-manifest.json"
    if prior.is_file():
        require(read_bytes(ADR) == read_bytes(prior), "ADR manifest differs from 0546")
    return {"status": "pass", "entries": len(files), "sha256": sha(ADR)}


def validate_source(plan: dict[str, Any]) -> dict[str, Any]:
    stage_manifest_path = need(BASELINE / "source-manifest.json", "baseline source manifest")
    manifest = source_manifest(stage_manifest_path, "baseline/source-manifest.json")
    expected = tree_manifest(plan["revision"])
    require(manifest == expected, "baseline source manifest differs from frozen Git revision")
    require(read_bytes(BASELINE / "source.patch") == b"",
            "baseline source patch is not empty")
    current = current_source_manifest()
    require(current == manifest, "current source tree differs from the frozen baseline")
    prior_manifest_path = PRIOR / "candidate" / "source-manifest.json"
    require(prior_manifest_path.is_file(), "0546 candidate source manifest is missing")
    require(read_bytes(stage_manifest_path) == read_bytes(prior_manifest_path),
            "0547 baseline is not the sealed 0546 candidate source")
    return {
        "status": "pass", "manifest_entries": len(manifest),
        "manifest_sha256": sha(stage_manifest_path),
        "source_patch_sha256": sha(BASELINE / "source.patch"),
        "prior_manifest_sha256": sha(prior_manifest_path),
        "adr": validate_adr(),
    }


def validate_campaign_host() -> dict[str, Any]:
    path = need(HERE / "host.json", "host.json")
    value = read_json(path, "host.json")
    require(isinstance(value, dict)
            and set(value) == {"utc", "affinity", "commands"},
            "host.json envelope differs")
    parse_time(value.get("utc"), "host.json.utc")
    affinity = value.get("affinity")
    require(isinstance(affinity, list) and affinity
            and all(isinstance(cpu, int) and not isinstance(cpu, bool) and cpu >= 0
                    for cpu in affinity),
            "host.json affinity differs")
    commands = value.get("commands")
    require(isinstance(commands, dict) and commands
            and all(isinstance(name, str) and name and isinstance(output, str)
                    and output for name, output in commands.items()),
            "host.json command inventory differs")
    return {"status": "pass", "sha256": sha(path), "affinity": affinity}


def validate_host(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict)
            and set(value) == {"observed_utc", "compiler_processes", "scope"},
            f"{label} host envelope differs")
    parse_time(value.get("observed_utc"), f"{label}.observed_utc")
    require(value.get("scope") == HOST_SCOPE and isinstance(value.get("compiler_processes"), list),
            f"{label} host scope differs")
    for process in value["compiler_processes"]:
        require(isinstance(process, dict) and isinstance(process.get("pid"), int)
                and process["pid"] > 0 and process.get("comm") in {"cargo", "rustc"}
                and isinstance(process.get("cwd"), str) and process["cwd"],
                f"{label} compiler process row differs")


def validate_artifacts(folder: Path, receipt: dict[str, Any], expected: set[str], label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} artifact inventory differs")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename
                and valid_digest(digest), f"{label} artifact entry differs")
        path = folder / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label} artifact custody differs: {filename}")
        if filename.endswith(".host.json"):
            validate_host(path, f"{label}/{filename}")


def validate_receipt(path: Path, expected_command: list[str], stage: str,
                     expected_binary: str | None, expected_artifacts: set[str]) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS,
            f"{rel(path)} receipt schema differs")
    require(value.get("command") == expected_command, f"{rel(path)} command differs")
    require(value.get("exit_code") == 0, f"{rel(path)} did not exit successfully")
    stage_manifest_path = HERE / stage / "source-manifest.json"
    require(value.get("source_manifest_sha256") == sha(stage_manifest_path)
            and value.get("execution_stage") == stage
            and value.get("execution_manifest_sha256") == sha(stage_manifest_path)
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("script_sha256") == sha(RUN)
            and value.get("binary_sha256") == expected_binary,
            f"{rel(path)} source/binary/script binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and set(environment) == ENVIRONMENT_KEYS
            and all(item is None for item in environment.values()),
            f"{rel(path)} instrumentation environment differs")
    times = interval(value, rel(path))
    validate_artifacts(path.parent, value, expected_artifacts, rel(path))
    return value, times


def binary_identity() -> dict[str, Any]:
    path = need(BASELINE / "binary-normal.json", "baseline binary identity")
    value = read_json(path, rel(path))
    require(isinstance(value, dict)
            and set(value) == {"path", "sha256", "bytes", "build_receipt_sha256", "source_manifest_sha256"},
            "binary identity envelope differs")
    expected_path = TARGET / "retained" / "baseline" / "normal"
    require(value.get("path") == str(expected_path) and valid_digest(value.get("sha256"))
            and isinstance(value.get("bytes"), int) and value["bytes"] > 0
            and value.get("build_receipt_sha256") == sha(BASELINE / "build-normal.receipt.json")
            and value.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json"),
            "binary identity binding differs")
    binary = Path(value["path"])
    if binary.exists():
        require(binary.is_file() and not binary.is_symlink() and sha(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                "retained binary contents differ")
    else:
        cleanup = read_json(CLEANUP, "cleanup.json")
        require(cleanup.get("owned_paths_absent") is True
                and str(TARGET) in cleanup.get("removed", []),
                "missing retained binary has no cleanup custody")
    return {"path": str(binary), "sha256": value["sha256"], "bytes": value["bytes"],
            "identity_sha256": sha(path)}


def build_command(plan: dict[str, Any]) -> list[str]:
    return ["env", f"TMPDIR={TARGET}/tmp", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0",
            "cargo", "build", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
            "--target-dir", str(TARGET)]


def validate_build(plan: dict[str, Any]) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    identity = binary_identity()
    receipt, times = validate_receipt(
        BASELINE / "build-normal.receipt.json", build_command(plan), "baseline", None,
        {"build-normal.host.json", "build-normal.stdout", "build-normal.stderr"},
    )
    require(receipt.get("binary_sha256") is None, "build receipt carries a runtime binary hash")
    return identity, times


def option(command: list[str], name: str) -> str:
    found = []
    for index, item in enumerate(command):
        if item == name:
            require(index + 1 < len(command), f"command omits value for {name}")
            found.append(command[index + 1])
        elif item.startswith(name + "="):
            found.append(item.split("=", 1)[1])
    require(len(found) == 1, f"command omits or repeats {name}")
    return found[0]


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    jobs = [{
        "name": "profile-r1-xls-owned", "owner": profile["xls_owner"], "shape": None,
        "cases": ["xls_owned_source_open_one_cell"], "runner": "xls",
        "warmup": profile["warmup"], "samples": profile["samples"],
    }]
    for shape in plan["groups"]["cfb"]["shapes"]:
        jobs.append({
            "name": f"profile-r1-cfb-{shape}", "owner": profile["cfb_owner"],
            "shape": shape, "cases": ["cfb_open"], "runner": "cfb",
            "warmup": profile["warmup"], "samples": profile["samples"],
        })
    return jobs


def profile_command(job: dict[str, Any], identity: dict[str, Any], plan: dict[str, Any]) -> list[str]:
    name = job["name"]
    command = ["taskset", "-c", str(plan["cpu"]), "valgrind", "--vgdb=no",
               "--tool=callgrind", "--collect-atstart=no",
               f"--toggle-collect={job['owner']}", f"--zero-before={job['owner']}",
               f"--dump-after={job['owner']}",
               f"--callgrind-out-file={BASELINE / (name + '.callgrind')}",
               *plan["profile"]["instruction_flags"], str(identity["path"]),
               "--case", ",".join(job["cases"]), "--warmup", str(job["warmup"]),
               "--samples", str(job["samples"]), "--json", str(BASELINE / (name + ".json")),
               "--corpus-manifest", str(BASELINE / (name + ".catalog.json"))]
    if job["shape"] is not None:
        command.extend(["--shape", job["shape"], "--payload", "incompressible"])
    return command


def numbered_dumps(name: str) -> list[tuple[int, Path]]:
    found = []
    prefix = f"{name}.callgrind."
    for path in BASELINE.glob(prefix + "*"):
        suffix = path.name[len(prefix):]
        if suffix.isdigit() and path.is_file() and not path.is_symlink():
            found.append((int(suffix), path))
    found.sort()
    require(found and [number for number, _ in found] == list(range(1, len(found) + 1)),
            f"{name} numbered Callgrind inventory differs")
    return found


def validate_profile(job: dict[str, Any], identity: dict[str, Any], plan: dict[str, Any]) -> tuple[tuple[dt.datetime, dt.datetime], dict[str, Any]]:
    name = job["name"]
    numbers = numbered_dumps(name)
    expected_count = 5 if job["shape"] is None else 6
    require(len(numbers) == expected_count, f"{name} dump count differs")
    expected = {f"{name}.host.json", f"{name}.json", f"{name}.catalog.json",
                f"{name}.stdout", f"{name}.stderr", f"{name}.callgrind"}
    expected.update(f"{name}.callgrind.{number}" for number, _ in numbers)
    receipt, times = validate_receipt(BASELINE / f"{name}.receipt.json",
                                      profile_command(job, identity, plan), "baseline",
                                      identity["sha256"], expected)
    command = receipt["command"]
    require(command[:4] == ["taskset", "-c", str(plan["cpu"]), "valgrind"]
            and command.count("--vgdb=no") == 1
            and option(command, "--callgrind-out-file") == str(BASELINE / (name + ".callgrind")),
            f"{name} Callgrind command controls differ")
    for _, dump in numbers:
        text = read_text(dump, rel(dump))
        require("events: Ir" in text and "Trigger: --dump-after=" in text,
                f"{rel(dump)} is not a selected-owner Ir dump")
    final_text = read_text(BASELINE / f"{name}.callgrind", f"{name}.callgrind")
    require("Program termination" in final_text and "events: Ir" in final_text,
            f"{name} final Callgrind dump is malformed")
    return times, {"name": name, "receipt_sha256": sha(BASELINE / f"{name}.receipt.json"),
                   "dumps": len(numbers), "setup_dumps": 0 if job["shape"] is None else 1,
                   "timed_dumps": 5}


def parse_symbols(path: Path) -> dict[str, tuple[int, int]]:
    symbols: dict[str, tuple[int, int]] = {}
    for line in read_text(path, rel(path)).splitlines():
        fields = line.split()
        if len(fields) == 4 and fields[2] in {"t", "T"}:
            try:
                symbols[fields[3]] = (int(fields[0], 16), int(fields[1], 16))
            except ValueError:
                pass
    return symbols


def validate_assembly(identity: dict[str, Any], plan: dict[str, Any]) -> tuple[list[tuple[dt.datetime, dt.datetime]], dict[str, Any]]:
    index_path = need(BASELINE / "assembly-index.json", "assembly index")
    index = read_json(index_path, rel(index_path))
    require(isinstance(index, dict)
            and set(index) == {"plan_sha256", "binary_sha256", "source_manifest_sha256", "script_sha256", "rows"}
            and index.get("plan_sha256") == sha(PLAN)
            and index.get("binary_sha256") == identity["sha256"]
            and index.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json")
            and index.get("script_sha256") == sha(INSPECT)
            and isinstance(index.get("rows"), list) and index["rows"],
            "assembly index envelope differs")
    symbols_receipt = BASELINE / "symbols.receipt.json"
    symbol_command = ["nm", "-S", "--defined-only", str(identity["path"])]
    _, symbols_times = validate_receipt(symbols_receipt, symbol_command, "baseline",
                                         identity["sha256"],
                                         {"symbols.host.json", "symbols.stdout", "symbols.stderr"})
    symbols = parse_symbols(BASELINE / "symbols.stdout")
    matching = {name: value for name, value in symbols.items()
                if any(owner in name for owner in plan["assembly"]["owners"])}
    rows = index["rows"]
    seen_symbols: set[str] = set()
    intervals = [symbols_times]
    for number, row in enumerate(rows):
        require(isinstance(row, dict)
                and set(row) == {"name", "symbol", "address_hex", "size_bytes", "receipt_sha256"}
                and row.get("name") == f"assembly-{number}"
                and isinstance(row.get("symbol"), str) and row["symbol"] in matching
                and row["symbol"] not in seen_symbols and valid_digest(row.get("receipt_sha256"))
                and isinstance(row.get("size_bytes"), int) and row["size_bytes"] > 0,
                f"assembly row {number} differs")
        try:
            address = int(row["address_hex"], 16)
        except (TypeError, ValueError) as error:
            raise VerificationError(f"assembly row {number} address is malformed") from error
        require(matching[row["symbol"]] == (address, row["size_bytes"]),
                f"assembly row {number} is not bound to nm")
        receipt_path = BASELINE / f"assembly-{number}.receipt.json"
        _, times = validate_receipt(receipt_path,
                                     ["objdump", "-d", "--disassemble=" + row["symbol"], str(identity["path"])],
                                     "baseline", identity["sha256"],
                                     {f"assembly-{number}.host.json", f"assembly-{number}.stdout",
                                      f"assembly-{number}.stderr"})
        require(row["receipt_sha256"] == sha(receipt_path),
                f"assembly row {number} receipt hash differs")
        require(read_bytes(BASELINE / f"assembly-{number}.stdout"),
                f"assembly row {number} has empty disassembly")
        seen_symbols.add(row["symbol"])
        intervals.append(times)
    require(set(matching) == seen_symbols, "assembly index does not retain every owner symbol")
    require(all(any(fragment in symbol for symbol in seen_symbols) for fragment in plan["assembly"]["required"]),
            "assembly index omits a required function")
    return intervals, {"rows": len(rows), "matching_symbols": len(matching),
                       "index_sha256": sha(index_path)}


def snapshot_bundle() -> dict[Path, str]:
    return {path: sha(path) for path in HERE.rglob("*") if path.is_file() and not path.is_symlink()}


def mirror_bundle(root: Path) -> Path:
    """Make a hard-linked read view so analyzer annotations cannot touch HERE."""
    view = root / "change-0547"
    shutil.copytree(HERE, view, copy_function=os.link,
                    ignore=shutil.ignore_patterns("SHA256SUMS", "__pycache__"))
    for sibling in HERE.parent.glob("change-*"):
        if sibling.name == "change-0547":
            continue
        link = root / sibling.name
        try:
            link.symlink_to(sibling, target_is_directory=True)
        except FileExistsError:
            pass
    return view


def replay_analyzer(script: Path, args: list[str], retained: Path) -> dict[str, Any]:
    need(script, rel(script))
    need(retained, rel(retained))
    before = snapshot_bundle()
    with tempfile.TemporaryDirectory(prefix=".litchi-0547-replay-", dir="/home/zhuhe") as temp:
        output = Path(temp) / (retained.stem + ".replayed.json")
        # Keep the analyzer's real ``HERE``.  Its retained receipt checks
        # deliberately bind absolute raw-artifact paths to this bundle; a
        # hard-linked mirror would change those paths and make a valid report
        # fail replay.  The explicit output is outside HERE, and the before /
        # after snapshot below rejects any evidence mutation.
        command = [sys.executable, "-B", str(script), *args,
                   "--output", str(output)]
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
        try:
            result = subprocess.run(command, cwd=REPO, env=environment,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    text=True, check=False)
        except OSError as error:
            raise VerificationError(f"{rel(script)} replay could not start: {error}") from error
        require(result.returncode == 0,
                f"{rel(script)} replay failed: {result.stderr[-3000:]}")
        require(output.is_file() and output.read_bytes() == read_bytes(retained),
                f"{rel(script)} replay differs from {rel(retained)}")
    after = snapshot_bundle()
    require(before == after, f"{rel(script)} replay modified the evidence bundle")
    return read_json(retained, rel(retained))


def replay_stdout_json(script: Path, retained: Path) -> dict[str, Any]:
    """Replay a deterministic stdout-only model without touching the bundle."""
    need(script, rel(script))
    need(retained, rel(retained))
    before = snapshot_bundle()
    with tempfile.TemporaryDirectory(prefix=".litchi-0547-model-", dir="/home/zhuhe") as temp:
        view = mirror_bundle(Path(temp))
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
        try:
            result = subprocess.run(
                [sys.executable, "-B", str(view / script.name)], cwd=REPO,
                env=environment, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                check=False,
            )
        except OSError as error:
            raise VerificationError(f"{rel(script)} replay could not start: {error}") from error
        require(result.returncode == 0,
                f"{rel(script)} replay failed: {result.stderr.decode(errors='replace')[-3000:]}")
        require(result.stderr == b"", f"{rel(script)} replay wrote stderr")
        require(result.stdout == read_bytes(retained),
                f"{rel(script)} replay differs from {rel(retained)}")
        try:
            value = json.loads(result.stdout)
        except json.JSONDecodeError as error:
            raise VerificationError(f"{rel(script)} retained stdout is not JSON") from error
    after = snapshot_bundle()
    require(before == after, f"{rel(script)} replay modified the evidence bundle")
    require(isinstance(value, dict), f"{rel(script)} JSON result is not an object")
    return value


def validate_model_receipt(name: str, script: Path, output: Path,
                           receipt_path: Path, _script: Path) -> dict[str, Any]:
    """Bind a post-profile model's retained stdout to its execution receipt."""
    receipt = read_json(receipt_path, rel(receipt_path))
    require(isinstance(receipt, dict), f"{rel(receipt_path)} is not an object")
    expected_command = ["python3", "-B", f"docs/performance/results/change-0547/{script.name}"]
    require(receipt.get("command") == expected_command and receipt.get("exit_code") == 0
            and receipt.get("script_sha256") == sha(script)
            and receipt.get("stdout_sha256") == sha(output),
            f"{rel(receipt_path)} command/hash binding differs")
    if name == "proof-check":
        require(set(receipt) == {"command", "generated_utc", "exit_code", "script_sha256",
                                 "stdout_sha256", "status", "scope", "result"}
                and receipt.get("status") == "pass"
                and receipt.get("scope") == (
                    "semantic ordered-error model; no Rust allocation, native latency, or aliasing equivalence claim"
                ), f"{rel(receipt_path)} proof envelope differs")
        parse_time(receipt.get("generated_utc"), f"{rel(receipt_path)}.generated_utc")
        result = receipt.get("result")
        require(isinstance(result, dict) and result.get("cases") == 376264
                and result.get("cycle_cases") == 13778
                and result.get("short_cycle_table_len") == 32,
                f"{rel(receipt_path)} proof result differs")
    else:
        require(set(receipt) == {"command", "start_utc", "end_utc", "seconds", "exit_code",
                                 "script_sha256", "stdout_sha256", "stderr_sha256"},
                f"{rel(receipt_path)} model envelope differs")
        start, end = interval(receipt, rel(receipt_path))
        stderr = HERE / "suboperations.stderr"
        require(receipt.get("stderr_sha256") == sha(stderr)
                and read_bytes(stderr) == b"", f"{rel(receipt_path)} stderr binding differs")
        return {"status": "pass", "start": start, "end": end}
    return {"status": "pass", "generated": parse_time(
        receipt["generated_utc"], f"{rel(receipt_path)}.generated_utc"
    )}


def validate_reports(plan: dict[str, Any], identity: dict[str, Any]) -> dict[str, Any]:
    reports: dict[str, Any] = {}
    report_specs = []
    require(NUMERIC_ANALYZER.is_file() and (HERE / "analysis.json").is_file(),
            "normal-report analyzer/report is missing")
    report_specs.append((NUMERIC_ANALYZER, [], HERE / "analysis.json", "normal"))
    require(PROFILE_ANALYZER.is_file() and (HERE / "profile-analysis.json").is_file(),
            "profile analyzer/report is missing")
    report_specs.append((PROFILE_ANALYZER, [], HERE / "profile-analysis.json", "profile"))
    require(INSTRUCTION_ANALYZER.is_file() and (HERE / "instruction-analysis.json").is_file(),
            "instruction analyzer/report is missing")
    report_specs.append((INSTRUCTION_ANALYZER, ["--stage", "baseline"], HERE / "instruction-analysis.json", "instruction"))
    for script, args, retained, key in report_specs:
        value = replay_analyzer(script, args, retained)
        require(isinstance(value, dict) and value.get("status") == "pass",
                f"{rel(retained)} does not report pass")
        require(value.get("plan_sha256") == sha(PLAN),
                f"{rel(retained)} plan binding differs")
        if value.get("source_manifest_sha256") is not None:
            require(value.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json"),
                    f"{rel(retained)} source binding differs")
        reports[key] = {"schema": value.get("schema"), "sha256": sha(retained),
                        "status": value.get("status")}
    instruction = read_json(HERE / "instruction-analysis.json")
    require(instruction.get("stage") == "baseline"
            and isinstance(instruction.get("jobs"), list) and len(instruction["jobs"]) == 4
            and instruction.get("validation", {}).get("timed_dump_count") == 20
            and instruction.get("validation", {}).get("setup_dump_count") == 3,
            "instruction analysis matrix differs from one-repeat plan")
    normal = read_json(HERE / "analysis.json")
    require(normal.get("stage") == "baseline" and normal.get("profile_count") == 4
            and normal.get("binary_sha256") == identity["sha256"],
            "normal-report analysis matrix or binary binding differs")
    profile = read_json(HERE / "profile-analysis.json")
    metadata = profile.get("metadata")
    require(profile.get("profile_count") == 4
            and profile.get("timed_constructor_dump_count") == 20
            and profile.get("setup_dump_count") == 3
            and isinstance(metadata, dict)
            and metadata.get("sha256") == identity["sha256"]
            and metadata.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json"),
            "profile analysis matrix or binary binding differs")
    require(instruction.get("binary_sha256") == identity["sha256"],
            "instruction analysis binary binding differs")

    # The two small scripts below are intentionally stdout-only model checks.
    # Their retained JSON is captured after profiling and is replayed in a
    # temporary working directory so this verifier never creates or replaces
    # an evidence report.
    model_specs = [
        (PROOF_SCRIPT, HERE / "proof-check.stdout", "proof"),
        (SUBOPERATION_SCRIPT, HERE / "suboperations.json", "suboperations"),
    ]
    for script, retained, key in model_specs:
        require(script.is_file() and retained.is_file(),
                f"{key} model script/report is missing")
        value = replay_stdout_json(script, retained)
        require(isinstance(value, dict) and value.get("status") == "pass",
                f"{rel(retained)} does not report pass")
        reports[key] = {"schema": value.get("schema"), "sha256": sha(retained),
                        "status": value.get("status")}
    validate_model_receipt("proof-check", PROOF_SCRIPT, HERE / "proof-check.stdout",
                           HERE / "proof-check.receipt.json", model_specs[0][0])
    validate_model_receipt("suboperations", SUBOPERATION_SCRIPT, HERE / "suboperations.json",
                           HERE / "suboperations.receipt.json", model_specs[1][0])
    return {"status": "pass", "reports": reports}


def validate_quality_reuse() -> dict[str, Any]:
    value = read_json(QUALITY_REUSE, "quality-reuse.json")
    require(isinstance(value, dict)
            and set(value) == {"scope", "prior_manifest_sha256", "current_manifest_sha256",
                               "quality_summary_sha256", "prior_seal_sha256", "receipts",
                               "tests_previously_passed"}
            and value.get("scope") == (
                "Exact full source manifest equals sealed0546 final candidate; no production/harness edits. Prior nine quality commands and1306 tests reused, not rerun or counted as0547 executions."
            )
            and value.get("prior_manifest_sha256") == value.get("current_manifest_sha256")
            and value.get("prior_manifest_sha256") == sha(PRIOR / "candidate" / "source-manifest.json")
            and value.get("quality_summary_sha256") == sha(PRIOR / "quality-summary.json")
            and value.get("prior_seal_sha256") == sha(PRIOR / "SHA256SUMS")
            and value.get("tests_previously_passed") == 1306,
            "quality reuse envelope differs")
    receipts = value.get("receipts")
    expected = {
        "final-quality-boundaries.receipt.json", "final-quality-cap-clippy.receipt.json",
        "final-quality-check.receipt.json", "final-quality-clippy.receipt.json",
        "final-quality-fmt.receipt.json", "final-quality-guard-clippy.receipt.json",
        "final-quality-guard-fmt.receipt.json", "final-quality-rustdoc.receipt.json",
        "final-quality-tests.receipt.json",
    }
    require(isinstance(receipts, dict) and set(receipts) == expected,
            "quality reuse receipt inventory differs")
    for name, digest in receipts.items():
        path = PRIOR / "candidate" / name
        require(valid_digest(digest) and path.is_file() and sha(path) == digest,
                f"reused quality receipt custody differs: {name}")
        receipt = read_json(path, rel(path))
        require(receipt.get("exit_code") == 0, f"reused quality receipt failed: {name}")
    summary = read_json(PRIOR / "quality-summary.json")
    require(summary.get("status") == "pass" and summary.get("full_test_count") == 1306
            and set(summary.get("receipts", {})) == expected,
            "prior quality summary differs")
    validate_recursive_seal(PRIOR, PRIOR / "SHA256SUMS", "0546 integration seal")
    return {"status": "pass", "tests": 1306, "checks": 9,
            "quality_summary_sha256": sha(PRIOR / "quality-summary.json"),
            "seal_sha256": sha(PRIOR / "SHA256SUMS")}


def validate_recursive_seal(root: Path, seal: Path, label: str) -> None:
    """Check a retained checksum inventory before trusting reused evidence."""
    need(seal, label)
    expected: dict[str, str] = {}
    for line in read_text(seal, label).splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), f"{label} line differs")
        name = safe_relative(fields[1], f"{label} path")
        require(name != seal.name and name not in expected,
                f"{label} contains a duplicate or self entry")
        expected[name] = fields[0]
    actual = {
        path.relative_to(root).as_posix(): sha(path)
        for path in root.rglob("*")
        if path.is_file() and not path.is_symlink() and path != seal
    }
    require(expected == actual and not any(path.is_symlink() for path in root.rglob("*")),
            f"{label} recursive inventory differs")


def check_serial(intervals: list[tuple[dt.datetime, dt.datetime]], label: str,
                 lower_bound: dt.datetime | None = None) -> None:
    require(intervals, f"{label} interval list is empty")
    if lower_bound is not None:
        require(intervals[0][0] >= lower_bound,
                f"{label} starts before the frozen input timestamp")
    for previous, current in zip(intervals, intervals[1:]):
        require(previous[1] <= current[0], f"{label} receipts overlap")


def receipt_inventory(plan: dict[str, Any]) -> dict[str, Any]:
    expected = {"build-normal", "symbols"}
    expected.update(job["name"] for job in profile_jobs(plan))
    expected.update(f"assembly-{number}"
                   for number in range(len(read_json(BASELINE / "assembly-index.json")["rows"])))
    actual = {path.name.removesuffix(".receipt.json")
              for path in BASELINE.glob("*.receipt.json")}
    require(actual == expected,
            f"baseline receipt inventory differs: expected={sorted(expected)} actual={sorted(actual)}")
    return {"status": "pass", "receipts": len(actual),
            "names_sha256": hashlib.sha256("\n".join(sorted(actual)).encode()).hexdigest()}


def validate_cleanup(plan: dict[str, Any], binary: dict[str, Any], *, required: bool) -> dict[str, Any] | None:
    if not CLEANUP.is_file():
        if required:
            raise IncompleteError("cleanup.json is missing")
        return None
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict), "cleanup record is not an object")
    removed = value.get("removed")
    require(value.get("target") == str(TARGET)
            and value.get("binary_sha256") == binary["sha256"]
            and isinstance(value.get("input_sha256"), dict)
            and removed == [str(TARGET)]
            and value.get("owned_paths_absent") is True,
            "cleanup target/binary binding differs")
    require(value.get("plan_sha256") in (None, sha(PLAN)),
            "cleanup plan binding differs")
    require(not value.get("accessible_process_references"),
            "cleanup retains accessible process references")
    require(not os.path.lexists(TARGET), "owned target still exists after cleanup")
    inputs = value["input_sha256"]
    require(inputs, "cleanup input bindings are empty")
    for name, digest in inputs.items():
        safe_relative(name, "cleanup input path")
        path = HERE / name
        require(valid_digest(digest) and path.is_file() and not path.is_symlink()
                and digest == sha(path), f"cleanup input binding differs: {name}")
    for name in ("plan.json", "run.py", "capture.py", "inspect_assembly.py", "adr-manifest.json"):
        require(inputs.get(name) == sha(HERE / name),
                f"cleanup omits frozen input binding: {name}")
    parse_time(value.get("utc", value.get("completed_utc")), "cleanup timestamp")
    return {"status": "pass", "sha256": sha(CLEANUP), "removed": removed}


def validate_seal() -> dict[str, Any]:
    need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(SEAL, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), "SHA256SUMS line differs")
        name = safe_relative(fields[1], "SHA256SUMS path")
        require(name != "SHA256SUMS" and name not in expected, "SHA256SUMS repeats or includes itself")
        expected[name] = fields[0]
    actual = {path.relative_to(HERE).as_posix(): sha(path)
              for path in HERE.rglob("*") if path.is_file() and not path.is_symlink() and path != SEAL}
    require(expected == actual, "SHA256SUMS recursive inventory differs")
    require(not any(path.is_symlink() for path in HERE.rglob("*")),
            "evidence bundle contains a symlink")
    return {"status": "pass", "entries": len(expected), "sha256": sha(SEAL)}


def verify(*, precleanup: bool = False) -> dict[str, Any]:
    documents = read_json(HERE / "documentation-manifest.json")
    expected_documents = {"docs/performance/" + name for name in
                          ("BASELINE.md", "HOTSPOTS.md", "REPORT.md", "ADR_COMPLIANCE.md", "GOAL_AUDIT.md")}
    expected_documents.add("docs/performance/changes/0547-ole2-collector-attribution.md")
    require(set(documents) == expected_documents, "documentation manifest inventory differs")
    for name, digest in documents.items():
        require(sha(REPO / name) == digest, f"documentation binding differs: {name}")
    adaptation = read_json(HERE / "analysis-adaptation.json")
    require(adaptation["canonical_instruction_analysis_sha256"] == sha(HERE / "instruction-analysis.json"),
            "canonical analysis adaptation binding differs")
    for name, digest in adaptation["superseded"].items():
        safe_relative(name, "superseded attempt path")
        require(sha(HERE / name) == digest, f"superseded attempt binding differs: {name}")
    require(read_json(HERE / "superseded-suboperations/suboperations.json")["rows"] ==
            read_json(HERE / "suboperations.json")["rows"], "final attribution rows changed")
    plan = check_plan()
    host = validate_campaign_host()
    source = validate_source(plan)
    build, build_interval = validate_build(plan)
    intervals = [build_interval]
    profiles = []
    for job in profile_jobs(plan):
        times, row = validate_profile(job, build, plan)
        intervals.append(times)
        profiles.append(row)
    assembly_intervals, assembly = validate_assembly(build, plan)
    intervals.extend(assembly_intervals)
    reports = validate_reports(plan, build)
    quality = validate_quality_reuse()
    check_serial(intervals, "build/profile/assembly", parse_time(
        read_json(FROZEN, "frozen-inputs.json")["utc"], "frozen-inputs.utc"
    ))
    inventory = receipt_inventory(plan)
    cleanup = validate_cleanup(plan, build, required=not precleanup)
    seal = None if precleanup else validate_seal()
    return {"status": "pass", "mode": "precleanup" if precleanup else "strict",
            "host": host, "source": source, "build": build, "profiles": profiles,
            "assembly": assembly, "reports": reports, "quality_reuse": quality,
            "inventory": inventory, "cleanup": cleanup, "seal": seal}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--precleanup", action="store_true",
                       help="check complete captured evidence before cleanup/sealing")
    group.add_argument("--strict", action="store_true",
                       help="require cleanup and recursive seal (default)")
    args = parser.parse_args(argv)
    try:
        result = verify(precleanup=args.precleanup)
    except (VerificationError, OSError, subprocess.SubprocessError) as error:
        print(f"0547 verify: {error}", file=sys.stderr)
        return 2
    print("PASS: frozen 0547 baseline source, serial build/profiles/assembly, analyzer replay, quality reuse"
          + (", cleanup, and recursive seal" if not args.precleanup else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
