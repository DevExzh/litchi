"""Fail-closed custody verifier for the 0535 diagnostic baseline.

This campaign has no candidate and makes no speedup claim.  The verifier
binds the baseline source to the sealed 0534 final source, replays the prior
14 quality gates, checks every new receipt and raw artifact, and replays the
numeric and instruction reports in temporary directories outside this bundle.
It never builds, captures, or writes an evidence report in the bundle.
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
INSPECT = HERE / "inspect_assembly.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
BASELINE = HERE / "baseline"
ANALYZER = HERE / "analyze.py"
INSTRUCTION_ANALYZER = HERE / "instruction_analysis.py"
INSTRUCTION_ANALYSIS = HERE / "instruction-analysis.json"
ASSEMBLY_INDEX = BASELINE / "assembly-index.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"

PRIOR = HERE.parent / "change-0534"
PRIOR_PLAN = PRIOR / "plan.json"
PRIOR_RUN = PRIOR / "run.py"
PRIOR_CHECKS = PRIOR / "checks.py"
PRIOR_ADR = PRIOR / "adr-manifest.json"
PRIOR_FINAL = PRIOR / "final"
PRIOR_QUALITY = PRIOR / "quality-summary.json"
PRIOR_VERIFICATION = PRIOR / "verification.json"
PRIOR_SEAL = PRIOR / "SHA256SUMS"

SCRATCH_ROOT = Path("/tmp/litchi-goal-0535")
TARGET = Path("/home/zhuhe/litchi-goal-0535-target")
CANONICAL_BINARY_ROOT = TARGET / "retained-binaries"
PRIOR_TARGET = Path("/home/zhuhe/litchi-goal-0534-target")

SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
EXPECTED_ENVIRONMENT = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
RECEIPT_KEYS = {
    "command", "start_utc", "end_utc", "seconds", "exit_code", "execution_stage",
    "execution_manifest_sha256", "plan_sha256", "script_sha256",
    "source_manifest_sha256", "binary_sha256", "environment", "artifacts",
}
QUALITY_NAMES = (
    "fmt", "harness-fmt", "cfb-tests", "cfb-no-default-tests", "xls-tests",
    "doc-tests", "ppt-tests", "workspace-check", "ole-clippy", "harness-clippy",
    "ole-rustdoc", "harness-rustdoc", "boundaries", "claims",
)
ASSEMBLY_OWNERS = ("SectorChainScratch", "CheckedBitSet")
ASSEMBLY_REQUIRED = ("collect_exact", "insert")
TIME_FORMAT = '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}'
XLS_CASE = "xls_owned_source_open_one_cell"
CFB_CASE = "cfb_open"
PROFILE_OPTIONS = [
    "--dump-instr=yes", "--dump-line=no", "--compress-pos=no", "--collect-jumps=yes",
]


class VerificationError(ValueError):
    """Malformed, contradictory, or out-of-scope evidence."""


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


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and ".." not in path.parts,
            f"{label} escapes its root")
    return value


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
    result: dict[str, str] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, name_bytes = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "frozen Git source entry is malformed")
        name = name_bytes.decode()
        if source_name(name):
            result[name] = hashlib.sha256(git(["git", "show", f"{revision}:{name}"])).hexdigest()
    require(result, "frozen Git revision has no source objects")
    return result


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


def check_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict) and set(frozen) == {"created_utc", "files"},
            "frozen input envelope differs")
    parse_time(frozen.get("created_utc"), "frozen-inputs.json.created_utc")
    files = frozen.get("files")
    expected_files = {"plan.json", "run.py", "capture.py", "inspect_assembly.py", "adr-manifest.json"}
    require(isinstance(files, dict) and set(files) == expected_files,
            "frozen input inventory differs")
    for name, expected in files.items():
        require(Path(name).name == name and valid_digest(expected),
                f"frozen input entry is malformed: {name!r}")
        path = HERE / name
        require(path.is_file() and not path.is_symlink() and sha(path) == expected,
                f"frozen input digest differs: {name}")

    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict)
            and plan.get("status") == "frozen-before-build-and-captures"
            and plan.get("priority") ==
            "OLE2/OOXML active; ODF deferred until that goal completes; iWork excluded"
            and plan.get("performance_claim") ==
            "none; diagnostic prerequisite to a separate measured candidate"
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
            and groups.get("xls", {}).get("cases") == [XLS_CASE]
            and groups.get("cfb", {}).get("cases") == [CFB_CASE]
            and groups.get("cfb", {}).get("shapes") == ["few-large"]
            and groups.get("cfb", {}).get("payload") == "incompressible",
            "plan case matrix differs")
    native = plan.get("native")
    require(isinstance(native, dict) and native.get("repeats") == 2
            and native.get("warmup") == 20 and native.get("samples") == 1000
            and "no candidate" in str(native.get("scope", "")).lower(),
            "native plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict) and profile.get("repeats") == 2
            and profile.get("warmup") == 0 and profile.get("samples") == 5
            and profile.get("jobs") == ["xls-owned", "cfb-few-large"]
            and profile.get("xls_owner") ==
            "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
            and profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open"
            and profile.get("options") == PROFILE_OPTIONS,
            "profile plan differs")
    assembly = plan.get("assembly")
    require(isinstance(assembly, dict)
            and assembly.get("owners") == list(ASSEMBLY_OWNERS)
            and assembly.get("required") == list(ASSEMBLY_REQUIRED),
            "assembly plan differs")
    reuse = plan.get("quality_reuse")
    require(isinstance(reuse, dict)
            and reuse.get("bundle") == "change-0534"
            and reuse.get("stage") == "final"
            and reuse.get("scope", "").find("14") >= 0
            and reuse.get("scope", "").find("4382") >= 0
            and all(valid_digest(reuse.get(key)) for key in (
                "manifest_sha256", "summary_sha256", "verification_sha256", "seal_sha256")),
            "quality reuse binding differs")
    return plan


def validate_adr() -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    prior = read_json(PRIOR_ADR, "change-0534/adr-manifest.json")
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


def validate_source(plan: dict[str, Any]) -> dict[str, Any]:
    need(BASELINE, "baseline evidence directory", directory=True)
    manifest = source_manifest(BASELINE / "source-manifest.json", "baseline source manifest")
    patch = need(BASELINE / "source.patch", "baseline source.patch")
    require(patch.stat().st_size == 0, "baseline source.patch is not empty")
    prior_manifest = source_manifest(PRIOR_FINAL / "source-manifest.json",
                                     "change-0534 final source manifest")
    require(manifest == prior_manifest
            and sha(BASELINE / "source-manifest.json") ==
            plan["quality_reuse"]["manifest_sha256"],
            "0535 baseline is not the exact 0534 final source")
    require(manifest == tree_manifest(plan["revision"]),
            "baseline source differs from the frozen Git revision")
    require(manifest == current_source_manifest(),
            "current checkout differs from the retained baseline source")
    require(not (HERE / "candidate").exists() and not (HERE / "final").exists(),
            "diagnostic baseline unexpectedly contains a candidate or final stage")
    return {"status": "pass", "manifest_sha256": sha(BASELINE / "source-manifest.json"),
            "manifest_entries": len(manifest), "source_patch_sha256": sha(patch),
            "prior_manifest_sha256": sha(PRIOR_FINAL / "source-manifest.json"),
            "adr": validate_adr()}


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


def validate_artifacts(folder: Path, value: dict[str, Any], expected: set[str], label: str) -> None:
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} raw artifact inventory differs")
    stems = {filename.split(".", 1)[0] for filename in expected}
    require(len(stems) == 1 and next(iter(stems)),
            f"{label} receipt artifact names are malformed")
    stem = next(iter(stems))
    try:
        actual = {path.name for path in folder.iterdir()
                  if path.name.startswith(stem + ".")
                  and path.name != stem + ".receipt.json"}
    except OSError as error:
        raise VerificationError(f"{label} raw artifact directory cannot be read") from error
    require(actual == expected,
            f"{label} raw artifact files differ from the receipt inventory")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename
                and valid_digest(digest), f"{label} raw artifact entry differs")
        require(not filename.endswith((".inclusive.txt", ".self.txt")),
                f"{label} receipt includes a derived annotation")
        path = folder / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label} raw artifact custody differs: {filename}")
        if filename.endswith(".host.json"):
            validate_host(path, f"{label}/{filename}")


def expected_artifacts(folder: Path, name: str, lane: str) -> set[str]:
    result = {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"}
    if lane == "build" or lane == "assembly":
        return result
    result.update({f"{name}.json", f"{name}.catalog.json"})
    if lane == "native":
        result.add(f"{name}.rss.json")
    elif lane == "profile":
        result.add(f"{name}.callgrind")
        suffixes = sorted(
            int(path.name.rsplit(".", 1)[1])
            for path in folder.glob(f"{name}.callgrind.*")
            if path.is_file() and path.name.rsplit(".", 1)[1].isdigit()
        )
        expected_count = 5 if name.endswith("xls-owned") else 6
        require(suffixes == list(range(1, expected_count + 1)),
                f"{rel(folder)}/{name} Callgrind inventory differs")
        result.update(f"{name}.callgrind.{number}" for number in suffixes)
    else:
        raise VerificationError(f"unknown capture lane: {lane}")
    return result


def validate_receipt(path: Path, name: str, command: list[str], source_sha: str,
                     binary_sha: str | None, expected: set[str], *,
                     execution_stage: str = "baseline") -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS
            and value.get("command") == command,
            f"{rel(path)} command differs")
    require(value.get("exit_code") == 0, f"{rel(path)} did not exit successfully")
    require(value.get("script_sha256") == sha(RUN)
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("source_manifest_sha256") == source_sha,
            f"{rel(path)} source/plan/script binding differs")
    require(value.get("execution_stage") == execution_stage
            and value.get("execution_manifest_sha256") == source_sha,
            f"{rel(path)} execution binding differs")
    require(value.get("binary_sha256") == binary_sha,
            f"{rel(path)} binary binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            f"{rel(path)} instrumentation environment differs")
    times = interval(value, rel(path))
    validate_artifacts(path.parent, value, expected, rel(path))
    return value, times


def build_command() -> list[str]:
    return ["env", "TMPDIR=" + str(TARGET / "tmp"), "CARGO_BUILD_JOBS=2",
            "CARGO_INCREMENTAL=0", "cargo", "build", "--release", "--locked",
            "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
            "litchi-perf-baseline", "--target-dir", str(TARGET)]


def validate_binary(plan: dict[str, Any]) -> dict[str, Any]:
    path = need(BASELINE / "binary-normal.json", "baseline binary identity")
    value = read_json(path, rel(path))
    expected_path = SCRATCH_ROOT / "baseline" / "normal"
    require(isinstance(value, dict) and value.get("path") == str(expected_path)
            and valid_digest(value.get("sha256"))
            and isinstance(value.get("bytes"), int) and value["bytes"] > 0
            and value.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json")
            and value.get("build_receipt_sha256") == sha(BASELINE / "build-normal.receipt.json"),
            "baseline binary descriptor differs")
    binary = Path(value["path"])
    if binary.exists():
        require(binary.is_file() and not binary.is_symlink()
                and SCRATCH_ROOT.is_symlink()
                and SCRATCH_ROOT.resolve() == CANONICAL_BINARY_ROOT
                and CANONICAL_BINARY_ROOT.is_dir()
                and not CANONICAL_BINARY_ROOT.is_symlink()
                and binary.resolve() == CANONICAL_BINARY_ROOT / "baseline" / "normal"
                and sha(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                "baseline binary alias or custody differs")
    else:
        cleanup = read_json(CLEANUP, "cleanup.json")
        require(cleanup.get("plan_sha256") == sha(PLAN)
                and cleanup.get("removed") == plan["owned_paths"]
                and cleanup.get("owned_paths_absent") is True
                and all(not os.path.lexists(item) for item in plan["owned_paths"]),
                "missing baseline binary lacks cleanup custody")
    return {"path": str(binary), "sha256": value["sha256"],
            "bytes": value["bytes"], "identity_sha256": sha(path)}


def validate_build(plan: dict[str, Any]) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    source_sha = sha(BASELINE / "source-manifest.json")
    path = need(BASELINE / "build-normal.receipt.json", "baseline build receipt")
    value, times = validate_receipt(path, "build-normal", build_command(), source_sha, None,
                                    expected_artifacts(BASELINE, "build-normal", "build"))
    binary = validate_binary(plan)
    return {"receipt_sha256": sha(path), "binary": binary,
            "source_manifest_sha256": source_sha}, [times]


def native_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    config = plan["native"]
    return [
        {"name": "native-r1-xls", "repeat": 1, "group": "xls",
         "selection": plan["groups"]["xls"], "warmup": config["warmup"],
         "samples": config["samples"]},
        {"name": "native-r1-cfb", "repeat": 1, "group": "cfb",
         "selection": plan["groups"]["cfb"], "warmup": config["warmup"],
         "samples": config["samples"]},
        {"name": "native-r2-cfb", "repeat": 2, "group": "cfb",
         "selection": plan["groups"]["cfb"], "warmup": config["warmup"],
         "samples": config["samples"]},
        {"name": "native-r2-xls", "repeat": 2, "group": "xls",
         "selection": plan["groups"]["xls"], "warmup": config["warmup"],
         "samples": config["samples"]},
    ]


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    config = plan["profile"]
    jobs = []
    for repeat in (1, 2):
        jobs.append({"name": f"profile-r{repeat}-xls-owned", "repeat": repeat,
                     "group": "xls-owned", "shape": None,
                     "owner": config["xls_owner"], "selection": {"cases": [XLS_CASE]},
                     "warmup": config["warmup"], "samples": config["samples"]})
        jobs.append({"name": f"profile-r{repeat}-cfb-few-large", "repeat": repeat,
                     "group": "cfb-few-large", "shape": "few-large",
                     "owner": config["cfb_owner"],
                     "selection": {"cases": [CFB_CASE], "shapes": ["few-large"],
                                    "payload": "incompressible"},
                     "warmup": config["warmup"], "samples": config["samples"]})
    return jobs


def capture_command(job: dict[str, Any], lane: str) -> list[str]:
    folder = BASELINE
    binary = SCRATCH_ROOT / "baseline" / "normal"
    command = ["taskset", "-c", "2"]
    if lane == "native":
        command += ["/usr/bin/time", "-f", TIME_FORMAT, "-o",
                    str(folder / f"{job['name']}.rss.json")]
    else:
        command += ["valgrind", "--vgdb=no", "--tool=callgrind", "--collect-atstart=no",
                    *PROFILE_OPTIONS, "--toggle-collect=" + job["owner"],
                    "--zero-before=" + job["owner"], "--dump-after=" + job["owner"],
                    "--callgrind-out-file=" + str(folder / f"{job['name']}.callgrind")]
    selection = job["selection"]
    command += [str(binary), "--case", ",".join(selection["cases"]), "--warmup",
                str(job["warmup"]), "--samples", str(job["samples"]), "--json",
                str(folder / f"{job['name']}.json"), "--corpus-manifest",
                str(folder / f"{job['name']}.catalog.json")]
    if "shapes" in selection:
        command += ["--shape", ",".join(selection["shapes"]), "--payload", selection["payload"]]
    return command


def validate_captures(plan: dict[str, Any], binary_sha: str) -> list[tuple[dt.datetime, dt.datetime]]:
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    source_sha = sha(BASELINE / "source-manifest.json")
    jobs = [(job, "native") for job in native_jobs(plan)]
    jobs += [(job, "profile") for job in profile_jobs(plan)]
    for job, lane in jobs:
        path = need(BASELINE / f"{job['name']}.receipt.json",
                    f"baseline {job['name']} receipt")
        value, times = validate_receipt(
            path, job["name"], capture_command(job, lane), source_sha, binary_sha,
            expected_artifacts(BASELINE, job["name"], lane),
        )
        require(value["command"] == capture_command(job, lane),
                f"{job['name']} command is not frozen")
        intervals.append(times)
    return intervals


def validate_assembly(plan: dict[str, Any], binary_sha: str) -> list[tuple[dt.datetime, dt.datetime]]:
    value = read_json(ASSEMBLY_INDEX, rel(ASSEMBLY_INDEX))
    require(isinstance(value, dict)
            and set(value) == {"plan_sha256", "binary_sha256", "source_manifest_sha256",
                               "script_sha256", "rows"}
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("binary_sha256") == binary_sha
            and value.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json")
            and value.get("script_sha256") == sha(INSPECT)
            and isinstance(value.get("rows"), list) and value["rows"],
            "assembly index envelope differs")
    symbols_path = need(BASELINE / "symbols.receipt.json", "symbols receipt")
    symbols_command = ["nm", "-S", "--defined-only", str(SCRATCH_ROOT / "baseline" / "normal")]
    _, symbols_interval = validate_receipt(
        symbols_path, "symbols", symbols_command, sha(BASELINE / "source-manifest.json"),
        binary_sha, expected_artifacts(BASELINE, "symbols", "assembly"),
    )
    symbols: dict[str, tuple[int, int]] = {}
    for line in read_text(BASELINE / "symbols.stdout", "symbols.stdout").splitlines():
        fields = line.split()
        if len(fields) == 4 and fields[2] in ("t", "T"):
            try:
                symbols[fields[3]] = (int(fields[0], 16), int(fields[1], 16))
            except ValueError:
                continue
    matching = {symbol: identity for symbol, identity in symbols.items()
                if any(owner in symbol for owner in ASSEMBLY_OWNERS)}
    rows = value["rows"]
    expected_names = [f"assembly-{index}" for index in range(len(rows))]
    seen_symbols: set[str] = set()
    intervals = [symbols_interval]
    require(len(rows) > 0 and {row.get("symbol") for row in rows if isinstance(row, dict)} == set(matching),
            "assembly index does not retain every matching symbol")
    for index, row in enumerate(rows):
        require(isinstance(row, dict)
                and set(row) == {"name", "symbol", "address_hex", "size_bytes", "receipt_sha256"}
                and row.get("name") == expected_names[index]
                and isinstance(row.get("symbol"), str) and row["symbol"] in matching
                and row["symbol"] not in seen_symbols
                and valid_digest(row.get("receipt_sha256"))
                and re.fullmatch(r"[0-9a-fA-F]+", row.get("address_hex", "")) is not None
                and isinstance(row.get("size_bytes"), int) and row["size_bytes"] > 0,
                f"assembly row {index} differs")
        symbol = row["symbol"]
        require(matching[symbol] == (int(row["address_hex"], 16), row["size_bytes"]),
                f"assembly row {index} is not bound to nm")
        receipt_path = need(BASELINE / f"{row['name']}.receipt.json",
                            f"{row['name']} receipt")
        _, times = validate_receipt(
            receipt_path, row["name"],
            ["objdump", "-d", "--disassemble=" + symbol,
             str(SCRATCH_ROOT / "baseline" / "normal")],
            sha(BASELINE / "source-manifest.json"), binary_sha,
            expected_artifacts(BASELINE, row["name"], "assembly"),
        )
        require(row["receipt_sha256"] == sha(receipt_path),
                f"{row['name']} receipt hash differs")
        seen_symbols.add(symbol)
        intervals.append(times)
    require(all(any(required in symbol for symbol in matching)
                for required in ASSEMBLY_REQUIRED),
            "assembly index omits a required function")
    return intervals


def test_count(stdout: str) -> int:
    return sum(int(item) for item in re.findall(r"test result: ok\. (\d+) passed;", stdout))


def quality_commands(checks_path: Path, target: Path) -> list[tuple[str, list[str]]]:
    tree = ast.parse(read_text(checks_path, rel(checks_path)), filename=str(checks_path))
    value = None
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(target_node, ast.Name) and target_node.id == "COMMANDS"
                for target_node in node.targets):
            value = ast.literal_eval(node.value)
            break
    require(isinstance(value, list), f"{rel(checks_path)} COMMANDS is missing")
    result = []
    for row in value:
        require(isinstance(row, (list, tuple)) and len(row) == 2
                and isinstance(row[0], str) and isinstance(row[1], list),
                f"{rel(checks_path)} command row is malformed")
        result.append((row[0], ["env", "TMPDIR=" + str(target / "test-tmp"),
                                 "CARGO_TARGET_DIR=" + str(target), "CARGO_BUILD_JOBS=2",
                                 "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings", *row[1]]))
    require([name for name, _ in result] == list(QUALITY_NAMES),
            f"{rel(checks_path)} quality command inventory differs")
    return result


def validate_prior_seal() -> dict[str, Any]:
    need(PRIOR_SEAL, "change-0534/SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(PRIOR_SEAL, "change-0534/SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]),
                "prior SHA256SUMS line differs")
        safe_relative(fields[1], "prior SHA256SUMS path")
        require(fields[1] != "SHA256SUMS" and fields[1] not in expected,
                "prior SHA256SUMS inventory is unsafe")
        expected[fields[1]] = fields[0]
    actual = {path.relative_to(PRIOR).as_posix(): sha(path)
              for path in PRIOR.rglob("*")
              if path.is_file() and not path.is_symlink() and path != PRIOR_SEAL}
    require(expected == actual and not any(path.is_symlink() for path in PRIOR.rglob("*")),
            "prior 0534 recursive seal differs")
    return {"status": "pass", "sha256": sha(PRIOR_SEAL), "entries": len(expected)}


def validate_prior_quality(plan: dict[str, Any]) -> dict[str, Any]:
    reuse = plan["quality_reuse"]
    require(sha(PRIOR_FINAL / "source-manifest.json") == reuse["manifest_sha256"],
            "prior final source manifest binding differs")
    require(sha(PRIOR_QUALITY) == reuse["summary_sha256"]
            and sha(PRIOR_VERIFICATION) == reuse["verification_sha256"]
            and sha(PRIOR_SEAL) == reuse["seal_sha256"],
            "prior quality reuse digest differs")
    verification = read_json(PRIOR_VERIFICATION, "change-0534/verification.json")
    require(isinstance(verification, dict) and verification.get("status") == "pass"
            and verification.get("source", {}).get("selected") == "final"
            and verification.get("quality", {}).get("stage") == "final",
            "prior verification did not accept the final source")

    summary = read_json(PRIOR_QUALITY, "change-0534/quality-summary.json")
    require(isinstance(summary, dict) and summary.get("status") == "pass"
            and summary.get("stage") == "final"
            and isinstance(summary.get("checks"), list)
            and len(summary["checks"]) == 14
            and summary.get("executed_tests") == 4382,
            "prior quality summary differs")
    commands = dict(quality_commands(PRIOR_CHECKS, PRIOR_TARGET))
    expected_names = {f"check-{name}.receipt.json" for name in QUALITY_NAMES}
    seen: set[str] = set()
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    total = 0
    source_sha = sha(PRIOR_FINAL / "source-manifest.json")
    for row in summary["checks"]:
        require(isinstance(row, dict)
                and set(row) == {"name", "receipt_sha256", "executed_tests"}
                and row["name"] in expected_names and row["name"] not in seen
                and valid_digest(row["receipt_sha256"]),
                "prior quality summary row differs")
        seen.add(row["name"])
        name = row["name"].removesuffix(".receipt.json")
        path = need(PRIOR_FINAL / row["name"], f"change-0534/final/{row['name']}")
        value, times = validate_prior_receipt(
            path, name, commands[name.removeprefix("check-")], source_sha
        )
        require(row["receipt_sha256"] == sha(path),
                f"prior quality receipt digest differs: {name}")
        count = test_count(read_text(path.with_name(path.name.replace(".receipt.json", ".stdout"))))
        require(isinstance(row["executed_tests"], int) and row["executed_tests"] == count,
                f"prior quality test count differs: {name}")
        intervals.append(times)
        total += count
    require(seen == expected_names and total == 4382, "prior quality check inventory differs")
    check_serial(intervals, "prior quality")
    seal = validate_prior_seal()
    return {"status": "pass", "checks": len(seen), "executed_tests": total,
            "summary_sha256": sha(PRIOR_QUALITY), "verification_sha256": sha(PRIOR_VERIFICATION),
            "seal": seal}


def validate_prior_receipt(path: Path, name: str, command: list[str],
                           source_sha: str) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    value = read_json(path, rel(path))
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS
            and value.get("command") == command
            and value.get("exit_code") == 0
            and value.get("script_sha256") == sha(PRIOR_RUN)
            and value.get("plan_sha256") == sha(PRIOR_PLAN)
            and value.get("source_manifest_sha256") == source_sha
            and value.get("execution_stage") == "final"
            and value.get("execution_manifest_sha256") == source_sha
            and value.get("binary_sha256") is None,
            f"{rel(path)} prior quality binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            f"{rel(path)} prior quality environment differs")
    times = interval(value, rel(path))
    validate_artifacts(path.parent, value,
                       {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"},
                       rel(path))
    return value, times


def check_serial(intervals: list[tuple[dt.datetime, dt.datetime]], label: str) -> None:
    ordered = sorted(intervals, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{label} receipt intervals overlap")


def replay_report(script: Path, args: list[str], retained: Path) -> Any:
    need(script, rel(script))
    need(retained, rel(retained))
    before = {path: (path.stat().st_size, path.stat().st_mtime_ns)
              for path in HERE.rglob("*") if path.is_file() and not path.is_symlink()}
    with tempfile.TemporaryDirectory(prefix=".litchi-0535-report-", dir="/home/zhuhe") as directory:
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
    after = {path: (path.stat().st_size, path.stat().st_mtime_ns)
             for path in HERE.rglob("*") if path.is_file() and not path.is_symlink()}
    require(before == after, f"{rel(script)} replay modified the evidence bundle")
    return read_json(retained, rel(retained))


def validate_numeric(plan: dict[str, Any], binary_sha: str) -> dict[str, Any]:
    value = replay_report(ANALYZER, [], HERE / "analysis.json")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0535-diagnostic-native-v1"
            and value.get("status") == "pass"
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json")
            and value.get("binary_sha256") == binary_sha
            and value.get("native_samples") == 4000
            and value.get("profile_timing_samples_excluded") == 20
            and isinstance(value.get("native_rows"), list)
            and len(value["native_rows"]) == 4
            and isinstance(value.get("profile_rows"), list)
            and len(value["profile_rows"]) == 4
            and isinstance(value.get("same_build_variations_over_five_percent"), list)
            and "no candidate" in str(value.get("scope", "")).lower()
            and "speedup" in str(value.get("scope", "")).lower(),
            "numeric diagnostic report envelope differs")
    native_names = {job["name"] for job in native_jobs(plan)}
    profile_names = {job["name"] for job in profile_jobs(plan)}
    require({row.get("name") for row in value["native_rows"]} == native_names
            and {row.get("name") for row in value["profile_rows"]} == profile_names,
            "numeric report job inventory differs")
    for row in value["same_build_variations_over_five_percent"]:
        require(isinstance(row, dict) and isinstance(row.get("review"), str)
                and row["review"].strip(), "numeric variation row lacks review")
    return {"status": "pass", "sha256": sha(HERE / "analysis.json"),
            "native_rows": len(value["native_rows"]),
            "profile_rows": len(value["profile_rows"]),
            "same_build_variations": len(value["same_build_variations_over_five_percent"])}


def validate_instruction_analysis(plan: dict[str, Any], binary_sha: str) -> dict[str, Any]:
    value = replay_report(INSTRUCTION_ANALYZER, ["--output"], INSTRUCTION_ANALYSIS)
    require(isinstance(value, dict)
            and value.get("schema") ==
            "litchi-ole2-change-0535-instruction-analysis-v1"
            and value.get("performance_claim") is None
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("source_manifest_sha256") == sha(BASELINE / "source-manifest.json")
            and value.get("binary_sha256") == binary_sha
            and value.get("assembly_index_sha256") == sha(ASSEMBLY_INDEX)
            and isinstance(value.get("jobs"), list)
            and [job.get("name") for job in value["jobs"]] ==
            [job["name"] for job in profile_jobs(plan)],
            "instruction analysis envelope differs")
    assembly = value.get("assembly")
    require(isinstance(assembly, dict)
            and assembly.get("required_fragments_present") == ["collect_exact", "insert"]
            and assembly.get("rows") == len(json.loads(ASSEMBLY_INDEX.read_text())["rows"])
            and assembly.get("collector_rows") == 1
            and assembly.get("insert_rows") == 1
            and isinstance(assembly.get("collector_instruction_count"), int)
            and assembly["collector_instruction_count"] > 0,
            "instruction analysis assembly summary differs")
    timed = setup = 0
    for job, expected in zip(value["jobs"], profile_jobs(plan)):
        expected_dumps = 5 if expected["group"] == "xls-owned" else 6
        expected_kind = "xls" if expected["group"] == "xls-owned" else "cfb"
        expected_runner = (
            "litchi_perf_baseline::run_xls_owned_source_case"
            if expected_kind == "xls" else "litchi_perf_baseline::run_cfb_open"
        )
        require(isinstance(job, dict)
                and job.get("name") == expected["name"]
                and job.get("repeat") == expected["repeat"]
                and job.get("kind") == expected_kind
                and job.get("runner") == expected_runner
                and job.get("owner") == expected["owner"]
                and job.get("dump_count") == expected_dumps
                and job.get("timed_dump_count") == 5
                and job.get("setup_dump_count") == (1 if expected["group"] == "cfb-few-large" else 0),
                f"instruction analysis job differs: {expected['name']}")
        dumps = job.get("dumps")
        require(isinstance(dumps, list) and len(dumps) == expected_dumps,
                f"instruction analysis dump inventory differs: {expected['name']}")
        bias = job.get("relocation_bias_hex")
        require(isinstance(bias, str)
                and re.fullmatch(r"-?0x[0-9a-f]+", bias) is not None,
                f"instruction analysis relocation bias differs: {expected['name']}")
        for number, dump in enumerate(dumps, 1):
            role = "setup" if expected_kind == "cfb" and number == 1 else "timed"
            raw = BASELINE / f"{expected['name']}.callgrind.{number}"
            require(isinstance(dump, dict)
                    and dump.get("number") == number
                    and dump.get("part") == number
                    and dump.get("path") == rel(raw)
                    and valid_digest(dump.get("sha256"))
                    and dump["sha256"] == sha(raw)
                    and dump.get("role") == role
                    and dump.get("relocation_bias_hex") == bias
                    and dump.get("positions") == ["instr"]
                    and dump.get("events") == ["Ir"],
                    f"instruction analysis dump differs: {expected['name']} part {number}")
        timed += job["timed_dump_count"]
        setup += job["setup_dump_count"]
    validation = value.get("validation")
    require(isinstance(validation, dict)
            and validation.get("raw_dump_count") == 22
            and validation.get("timed_dump_count") == 20
            and validation.get("setup_dump_count") == 2
            and validation.get("positions_instr_only") is True
            and validation.get("events_ir_only") is True
            and validation.get("single_bias_per_job") is True
            and validation.get("single_bias_across_profiles") is True
            and validation.get("collector_instruction_ir_equals_function_self") is True
            and validation.get("setup_and_timed_separated") is True
            and validation.get("direct_callees_separate_from_instruction_ir") is True
            and validation.get("no_operation_local_timing_claim") is True
            and timed == 20 and setup == 2,
            "instruction analysis validation summary differs")
    return {"status": "pass", "sha256": sha(INSTRUCTION_ANALYSIS),
            "profile_rows": len(value["jobs"]), "timed_dumps": timed,
            "setup_dumps": setup}


def validate_inventory(plan: dict[str, Any]) -> dict[str, Any]:
    index = read_json(ASSEMBLY_INDEX, rel(ASSEMBLY_INDEX))
    rows = index.get("rows") if isinstance(index, dict) else None
    require(isinstance(rows, list), "assembly rows are missing")
    expected = {"build-normal", "symbols"}
    expected.update(row["name"] for row in rows if isinstance(row, dict) and isinstance(row.get("name"), str))
    expected.update(job["name"] for job in native_jobs(plan))
    expected.update(job["name"] for job in profile_jobs(plan))
    actual = {path.name.removesuffix(".receipt.json")
              for path in BASELINE.glob("*.receipt.json")}
    require(actual == expected,
            f"baseline receipt inventory differs: expected {len(expected)}, found {len(actual)}")
    return {"status": "pass", "receipts": len(actual),
            "receipt_names_sha256": hashlib.sha256(
                ("\n".join(sorted(actual)) + "\n").encode()).hexdigest()}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict)
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
    return {"status": "pass", "sha256": sha(CLEANUP)}


def validate_seal() -> dict[str, Any]:
    need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(SEAL, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), "SHA256SUMS line differs")
        safe_relative(fields[1], "SHA256SUMS path")
        require(fields[1] != "SHA256SUMS" and fields[1] not in expected,
                "SHA256SUMS inventory is unsafe")
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
    prior = validate_prior_quality(plan)
    if name in {"prior-quality", "quality-reuse"}:
        return prior
    builds, build_intervals = validate_build(plan)
    if name in {"build", "builds"}:
        check_serial(build_intervals, "build")
        return {"status": "pass", "build": builds}
    binary_sha = builds["binary"]["sha256"]
    assembly_intervals = validate_assembly(plan, binary_sha)
    if name == "assembly":
        check_serial(build_intervals + assembly_intervals, "build/assembly")
        return {"status": "pass", "build": builds, "assembly_receipts": len(assembly_intervals)}
    capture_intervals = validate_captures(plan, binary_sha)
    if name in {"captures", "native", "profile"}:
        check_serial(build_intervals + assembly_intervals + capture_intervals,
                     "build/assembly/capture")
        return {"status": "pass", "build": builds, "capture_receipts": len(capture_intervals)}
    intervals = build_intervals + assembly_intervals + capture_intervals
    numeric = validate_numeric(plan, binary_sha)
    if name == "analysis":
        return numeric
    instruction = validate_instruction_analysis(plan, binary_sha)
    if name in {"instruction", "instructions"}:
        return instruction
    check_serial(intervals, "all measurement")
    inventory = validate_inventory(plan)
    if name == "inventory":
        return inventory
    if name == "cleanup":
        return validate_cleanup(plan)
    if name == "seal":
        return validate_seal()
    if name == "all":
        cleanup = validate_cleanup(plan)
        seal = validate_seal()
        return {"status": "pass", "source": source, "prior_quality": prior,
                "build": builds, "assembly_receipts": len(assembly_intervals),
                "capture_receipts": len(capture_intervals), "numeric": numeric,
                "instruction": instruction, "inventory": inventory,
                "cleanup": cleanup, "seal": seal}
    raise VerificationError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    try:
        result = component(selected)
        return {"schema": "litchi-0535-diagnostic-verification-v1", "status": "pass",
                "scope": selected, "performance_claim": "none; baseline diagnostics only",
                "result": result}
    except IncompleteError as error:
        return {"schema": "litchi-0535-diagnostic-verification-v1", "status": "incomplete",
                "scope": selected, "performance_claim": "none; baseline diagnostics only",
                "error": str(error)}
    except (VerificationError, OSError, subprocess.CalledProcessError,
            KeyError, TypeError, AttributeError, IndexError) as error:
        return {"schema": "litchi-0535-diagnostic-verification-v1", "status": "fail",
                "scope": selected, "performance_claim": "none; baseline diagnostics only",
                "error": str(error)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=(
        "all", "source", "prior-quality", "quality-reuse", "build", "builds",
        "assembly", "captures", "native", "profile", "analysis", "instruction",
        "instructions", "inventory", "cleanup", "seal",
    ), default="all")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.output and args.output.resolve().is_relative_to(HERE.resolve()):
        raise SystemExit("verification output must be outside the evidence bundle")
    report = run_bundle(args.component)
    text = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
    print(text, end="")
    return 2 if args.strict and report["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
