"""Independent verifier for the 0532 measurement-only CFB/OLE2 bundle.

The bundle contains a current-head baseline, with no candidate or speedup
claim.  This verifier checks the frozen source and driver, replays Git blob
custody, validates every serial build/capture/quality receipt, delegates raw
native/profile/hardware semantics to the retained analyzers in temporary
output files, binds the individual same-build variation review, and checks
cleanup and the final seal.  It never builds, captures, or runs Rust.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
BASELINE = HERE / "baseline"
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
ANALYSIS = HERE / "analysis.json"
ANALYZER = HERE / "analyze.py"
PROFILE_ANALYSIS = HERE / "profile-analysis.json"
PROFILE_ANALYZER = HERE / "analyze_profiles.py"
PROFILE_RETRY_PLAN = HERE / "profile-retry-plan.json"
PROFILE_RETRY_FROZEN = HERE / "profile-retry-frozen.json"
PROFILE_RETRY_SCRIPT = HERE / "capture_profiles.py"
FAILED_PROFILE_DIR = HERE / "failed-profile-initial"
HARDWARE_ANALYSIS = HERE / "hardware-analysis.json"
HARDWARE_ANALYZER = HERE / "analyze_hardware.py"
MECHANISM_ANALYSIS = HERE / "mechanism-analysis.json"
MECHANISM_ANALYZER = HERE / "analyze_mechanism.py"
VARIATION_REVIEW = HERE / "variation-review.json"
ASSEMBLY_PLAN = HERE / "assembly-plan.json"
ASSEMBLY_INDEX = HERE / "assembly-index.json"
ASSEMBLY_SCRIPT = HERE / "inspect_assembly.py"
QUALITY_PLAN = HERE / "quality-plan.json"
QUALITY_CHECKS = HERE / "checks.py"
QUALITY_REUSE = HERE / "quality-reuse.json"
QUALITY_SUMMARY = HERE / "quality-summary.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
PRIOR = HERE.parent / "change-0531"
PRIOR_RESTORED_MANIFEST = PRIOR / "restored" / "source-manifest.json"
PRIOR_RESTORED_QUALITY = PRIOR / "restored-quality-summary.json"
PRIOR_VERIFIER = PRIOR / "verify.py"
PRIOR_SEAL = PRIOR / "SHA256SUMS"
SCRATCH = Path("/tmp/litchi-goal-0532")
TARGET = Path("/home/zhuhe/litchi-goal-0532-target")
SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
EXPECTED_ENVIRONMENT = {
    "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
TIME_FORMAT = '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}'
CFB_SHAPES = ("tiny", "many-small", "few-large")
CFB_CASE = "cfb_open"
XLS_OWNER_CASE = "xls_owned_source_open_one_cell"
XLS_CASES = (
    "xls_semantic_open", "xls_eager_open_list_worksheets",
    "xls_eager_open_one_cell", "xls_source_backed_open",
    "xls_source_backed_open_list_worksheets", "xls_source_backed_open_one_cell",
    "xls_owned_source_open", "xls_owned_source_open_list_worksheets",
    "xls_owned_source_open_one_cell",
)
QUALITY_NAMES = (
    "cfb-tests", "cfb-no-default-tests", "xls-tests", "owner-clippy",
    "owner-rustdoc",
)


class VerificationError(ValueError):
    """Malformed, missing, or contradictory retained evidence."""


class IncompleteError(VerificationError):
    """Evidence has not arrived yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def need(path: Path, label: str | None = None) -> Path:
    label = label or path.as_posix()
    if not path.is_file() or path.is_symlink():
        raise IncompleteError(f"{label} is missing or not a regular file")
    return path


def read_json(path: Path, label: str | None = None) -> Any:
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"cannot read {label or path.name}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    need(path, label)
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise VerificationError(f"cannot read {label or path.name}: {error}") from error


def digest(path: Path) -> str:
    if not path.is_file() or path.is_symlink():
        raise VerificationError(f"cannot hash non-regular file: {path}")
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        raise VerificationError(f"cannot hash {path}: {error}") from error
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


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError:
        return str(path)


def git(args: list[str], *, input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, input=input_data,
                                       stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        stderr = getattr(error, "stderr", b"")
        detail = stderr.decode(errors="replace")[-2000:]
        raise VerificationError(f"Git command failed ({' '.join(args)}): {detail}") from error


def source_name(name: str) -> bool:
    return name in SOURCE_EXACT or any(name.startswith(prefix) for prefix in SOURCE_PREFIXES)


def source_manifest(path: Path, label: str | None = None) -> dict[str, str]:
    value = read_json(path, label or relative(path))
    require(isinstance(value, dict) and value,
            f"{label or relative(path)} is not a non-empty source manifest")
    result: dict[str, str] = {}
    for name, value in value.items():
        safe_relative(name, f"{label or relative(path)} source path")
        require(source_name(name) and valid_digest(value),
                f"{label or relative(path)} has an invalid source entry: {name}")
        require(name not in result, f"{label or relative(path)} repeats {name}")
        result[name] = value
    return result


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise VerificationError(f"{label} timestamp is malformed") from error
    require(result.tzinfo is not None, f"{label} has no timezone")
    return result


def receipt_interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    require(end > start, f"{label} interval is not positive")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and float(seconds) > 0, f"{label} duration is invalid")
    return start, end


def check_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict) and set(frozen) == {"plan.json", "run.py"},
            "frozen input inventory differs")
    require(frozen.get("plan.json") == digest(PLAN)
            and frozen.get("run.py") == digest(RUN),
            "frozen plan/run digest differs")
    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict)
            and plan.get("status") == "frozen-before-build-and-capture",
            "plan is not frozen before build/capture")
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
            and plan.get("owned_paths") == [str(SCRATCH), str(TARGET)],
            "plan priority, CPU, or owned paths differ")
    groups = plan.get("groups")
    require(isinstance(groups, dict)
            and groups.get("xls", {}).get("cases") == list(XLS_CASES)
            and groups.get("cfb", {}).get("cases") == [CFB_CASE]
            and groups.get("cfb", {}).get("shapes") == list(CFB_SHAPES)
            and groups["cfb"].get("payload") == "incompressible",
            "plan source groups differ")
    native = plan.get("native")
    require(isinstance(native, dict)
            and native.get("repeats") == 2 and native.get("warmup") == 20
            and native.get("samples") == 1000
            and native.get("order") == ["r1 xls", "r1 cfb", "r2 cfb", "r2 xls"],
            "native plan differs")
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
            and profile.get("jobs") == ["xls-owned", "cfb-tiny", "cfb-many-small",
                                         "cfb-few-large"]
            and profile.get("xls_owner") ==
            "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
            and profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open",
            "profile plan differs")
    hardware = plan.get("hardware")
    require(isinstance(hardware, dict)
            and hardware.get("repeats") == 2
            and hardware.get("case") == "xls_owned_source_open_one_cell"
            and hardware.get("warmup") == 0 and hardware.get("samples") == 1000
            and hardware.get("events") ==
            "{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations",
            "hardware plan differs")
    review = plan.get("review")
    require(isinstance(review, dict)
            and review.get("same_build_adverse_percent") == 5,
            "same-build review threshold differs")
    return plan


def git_tree_source_manifest(revision: str) -> dict[str, str]:
    raw = git(["git", "ls-tree", "-r", "-z", revision])
    objects: dict[str, str] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, name_bytes = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "Git source tree entry is not a blob")
        name = name_bytes.decode()
        if source_name(name):
            objects[name] = fields[2].decode()
    require(objects, "Git revision has no source objects")
    payload = b"\n".join(value.encode() for value in sorted(set(objects.values()))) + b"\n"
    batch = git(["git", "cat-file", "--batch"], input_data=payload)
    hashes: dict[str, str] = {}
    position = 0
    while position < len(batch):
        line_end = batch.find(b"\n", position)
        require(line_end >= 0, "Git batch response is malformed")
        fields = batch[position:line_end].split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "Git source object is not a blob")
        oid = fields[0].decode()
        length = int(fields[2])
        position = line_end + 1
        data = batch[position:position + length]
        require(len(data) == length, "Git source object is truncated")
        hashes[oid] = hashlib.sha256(data).hexdigest()
        position += length
        require(batch[position:position + 1] == b"\n", "Git batch separator is missing")
        position += 1
    return {name: hashes[oid] for name, oid in objects.items()}


def validate_source(plan: dict[str, Any]) -> dict[str, Any]:
    manifest_path = need(BASELINE / "source-manifest.json", "baseline source manifest")
    patch_path = need(BASELINE / "source.patch", "baseline source patch")
    manifest = source_manifest(manifest_path)
    require(patch_path.read_bytes() == b"", "baseline source patch is not empty")
    tree = git_tree_source_manifest(plan["revision"])
    require(manifest == tree, "baseline source manifest differs from revision blobs")
    current_revision = git(["git", "rev-parse", "HEAD"]).decode().strip()
    require(current_revision == plan["revision"],
            "current checkout revision differs from frozen baseline revision")
    for name, expected in manifest.items():
        path = REPO / name
        require(path.is_file() and not path.is_symlink() and digest(path) == expected,
                f"current source differs from frozen baseline: {name}")
    prior = source_manifest(PRIOR_RESTORED_MANIFEST, "0531 restored source manifest")
    require(prior == manifest, "0532 source map differs from 0531 restored source")
    adr = read_json(ADR, "adr-manifest.json")
    files = adr.get("files") if isinstance(adr, dict) else None
    require(isinstance(files, dict) and files, "ADR manifest is malformed")
    for name, expected in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith("docs/adr/") and valid_digest(expected)
                and (REPO / name).is_file() and digest(REPO / name) == expected,
                f"ADR hash differs: {name}")
    require(valid_digest(adr.get("prior_audit_source_binding_sha256")),
            "ADR prior audit binding is malformed")
    return {"status": "pass", "manifest_sha256": digest(manifest_path),
            "manifest_entries": len(manifest), "patch_sha256": digest(patch_path),
            "revision": plan["revision"], "adr_entries": len(files)}


def validate_host(path: Path, label: str) -> None:
    value = read_json(path, label)
    require(isinstance(value, dict)
            and value.get("scope") ==
            "Accessible cargo/rustc process observation; no host quiescence guarantee"
            and isinstance(value.get("compiler_processes"), list),
            f"{label} host observation differs")
    for process in value["compiler_processes"]:
        require(isinstance(process, dict)
                and isinstance(process.get("pid"), int) and process["pid"] > 0
                and process.get("comm") in ("cargo", "rustc")
                and isinstance(process.get("cwd"), str) and process["cwd"],
                f"{label} host process row differs")


def validate_artifacts(folder: Path, receipt: dict[str, Any], expected: set[str],
                       label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{label} artifact inventory differs")
    for filename, expected_digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename
                and valid_digest(expected_digest), f"{label} artifact entry differs")
        path = folder / filename
        require(path.is_file() and not path.is_symlink() and digest(path) == expected_digest,
                f"{label} artifact custody differs: {filename}")
    host = folder / f"{label.split('/')[-1]}.host.json"
    validate_host(host, f"{label} host")


def validate_receipt(path: Path, name: str, manifest_sha: str,
                     command: list[str], binary_sha: str | None,
                     expected_artifacts: set[str], *, allow_failure: bool = False
                     ) -> tuple[dict[str, Any], tuple[dt.datetime, dt.datetime]]:
    value = read_json(path, relative(path))
    require(value.get("command") == command, f"{relative(path)} command differs")
    require(allow_failure or value.get("exit_code") == 0,
            f"{relative(path)} did not exit successfully")
    require(value.get("plan_sha256") == digest(PLAN)
            and value.get("script_sha256") == digest(RUN)
            and value.get("source_manifest_sha256") == manifest_sha,
            f"{relative(path)} source/driver binding differs")
    require(value.get("binary_sha256") == binary_sha,
            f"{relative(path)} binary binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and set(environment) == EXPECTED_ENVIRONMENT
            and all(item is None for item in environment.values()),
            f"{relative(path)} instrumentation environment differs")
    start_end = receipt_interval(value, relative(path))
    validate_artifacts(path.parent, value, expected_artifacts, f"{relative(path.parent)}/{name}")
    return value, start_end


def build_command(kind: str) -> list[str]:
    executable = "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else "")
    result = ["env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
              "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
              "--bin", executable, "--target-dir", str(TARGET)]
    if kind == "alloc":
        result += ["--features", "allocator-metrics"]
    return result


def load_binary_identity(kind: str, manifest_sha: str,
                         build_receipt_sha: str) -> dict[str, Any]:
    path = need(BASELINE / f"binary-{kind}.json", f"baseline {kind} binary identity")
    value = read_json(path, relative(path))
    expected_path = SCRATCH / kind
    require(isinstance(value, dict) and value.get("path") == str(expected_path)
            and valid_digest(value.get("sha256"))
            and isinstance(value.get("bytes"), int) and value["bytes"] > 0
            and value.get("source_manifest_sha256") == manifest_sha
            and value.get("build_receipt_sha256") == build_receipt_sha,
            f"{relative(path)} identity differs")
    binary = Path(value["path"])
    if binary.is_file() and not binary.is_symlink():
        require(digest(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                f"{relative(path)} binary custody differs")
    else:
        cleanup = read_json(CLEANUP, "cleanup receipt")
        require(cleanup.get("owned_paths_absent") is True
                and set(cleanup.get("removed", [])) == {str(SCRATCH), str(TARGET)},
                f"{relative(path)} missing binary lacks cleanup custody")
    return {"kind": kind, "path": str(binary), "sha256": value["sha256"],
            "bytes": value["bytes"], "identity_sha256": digest(path)}


def validate_builds(plan: dict[str, Any], manifest_sha: str) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    identities: dict[str, dict[str, Any]] = {}
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    for kind in ("normal", "alloc"):
        name = f"build-{kind}"
        path = need(BASELINE / f"{name}.receipt.json", f"baseline {name} receipt")
        expected = {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"}
        _value, interval = validate_receipt(path, name, manifest_sha,
                                            build_command(kind), None, expected)
        intervals.append(interval)
        identities[kind] = load_binary_identity(kind, manifest_sha, digest(path))
    return identities, intervals


def capture_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    if lane == "native":
        config = plan["native"]
        groups = ["xls", "cfb"]
        jobs = []
        for repeat in range(1, config["repeats"] + 1):
            selected = groups if repeat == 1 else list(reversed(groups))
            for group in selected:
                jobs.append({"name": f"native-r{repeat}-{group}", "lane": lane,
                             "repeat": repeat, "group": group,
                             "selection": plan["groups"][group],
                             "warmup": config["warmup"], "samples": config["samples"]})
        return jobs
    if lane == "alloc":
        config = plan["allocation"]
        return [{"name": f"alloc-r{repeat}-{group}", "lane": lane,
                 "repeat": repeat, "group": group,
                 "selection": plan["groups"][group],
                 "warmup": config["warmup"], "samples": config["samples"]}
                for repeat in range(1, config["repeats"] + 1)
                for group in ("xls", "cfb")]
    if lane == "profile":
        config = plan["profile"]
        jobs = []
        for repeat in range(1, config["repeats"] + 1):
            jobs.append({"name": f"profile-r{repeat}-xls-owned", "lane": lane,
                         "repeat": repeat, "group": "xls-owned", "shape": None,
                         "owner": config["xls_owner"], "warmup": config["warmup"],
                         "samples": config["samples"],
                         "selection": {"cases": [XLS_OWNER_CASE]}})
            for shape in CFB_SHAPES:
                jobs.append({"name": f"profile-r{repeat}-cfb-{shape}", "lane": lane,
                             "repeat": repeat, "group": f"cfb-{shape}", "shape": shape,
                             "owner": config["cfb_owner"], "warmup": config["warmup"],
                             "samples": config["samples"],
                             "selection": {"cases": [CFB_CASE], "shapes": [shape],
                                            "payload": "incompressible"}})
        return jobs
    if lane == "hardware":
        config = plan["hardware"]
        return [{"name": f"hardware-r{repeat}-xls-owned", "lane": lane,
                 "repeat": repeat, "group": "xls-owned", "warmup": config["warmup"],
                 "samples": config["samples"],
                 "selection": {"cases": [config["case"]]}}
                for repeat in range(1, config["repeats"] + 1)]
    raise VerificationError(f"unknown capture lane: {lane}")


def capture_command(job: dict[str, Any], plan: dict[str, Any]) -> list[str]:
    lane = job["lane"]
    kind = "alloc" if lane == "alloc" else "normal"
    folder = BASELINE
    name = job["name"]
    command = ["taskset", "-c", str(plan["cpu"])]
    if lane == "native":
        command += ["/usr/bin/time", "-f", TIME_FORMAT, "-o",
                    str(folder / f"{name}.rss.json")]
    elif lane == "profile":
        command += ["valgrind", "--vgdb=no", "--tool=callgrind", "--collect-atstart=no",
                    "--toggle-collect=" + job["owner"],
                    "--zero-before=" + job["owner"],
                    "--dump-after=" + job["owner"],
                    "--callgrind-out-file=" + str(folder / f"{name}.callgrind")]
    elif lane == "hardware":
        command += ["perf", "stat", "-x", ",", "-o", str(folder / f"{name}.csv"),
                    "-e", plan["hardware"]["events"], "--"]
    command += [str(SCRATCH / kind), "--case", ",".join(job["selection"]["cases"]),
                "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
                "--json", str(folder / f"{name}.json"),
                "--corpus-manifest", str(folder / f"{name}.catalog.json")]
    if "shapes" in job["selection"]:
        command += ["--shape", ",".join(job["selection"]["shapes"]),
                    "--payload", job["selection"]["payload"]]
    return command


def capture_artifacts(job: dict[str, Any]) -> set[str]:
    name = job["name"]
    result = {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr",
              f"{name}.json", f"{name}.catalog.json"}
    if job["lane"] == "native":
        result.add(f"{name}.rss.json")
    elif job["lane"] == "hardware":
        result.add(f"{name}.csv")
    elif job["lane"] == "profile":
        result.add(f"{name}.callgrind")
        expected_numbered = 5 if job["shape"] is None else 6
        numbers = sorted(
            int(path.name.rsplit(".", 1)[1])
            for path in BASELINE.glob(f"{name}.callgrind.*")
            if path.name.rsplit(".", 1)[1].isdigit()
        )
        require(numbers == list(range(1, expected_numbered + 1)),
                f"{name} numbered Callgrind dump inventory differs")
        result.update(f"{name}.callgrind.{number}" for number in numbers)
    return result


def validate_profile_retry(plan: dict[str, Any], normal: dict[str, Any],
                           manifest_sha: str) -> dict[str, Any]:
    """Bind the one failed profiler launch and the frozen retry protocol."""
    retry = read_json(PROFILE_RETRY_PLAN, "profile-retry-plan.json")
    frozen = read_json(PROFILE_RETRY_FROZEN, "profile-retry-frozen.json")
    require(isinstance(retry, dict)
            and retry.get("schema") == "litchi-0532-profile-retry-plan-v1"
            and retry.get("status") == "frozen-before-retry"
            and retry.get("primary_plan_sha256") == digest(PLAN)
            and retry.get("frozen_run_sha256") == digest(RUN)
            and retry.get("retry_extra_argument") == "--vgdb=no"
            and "same binary" in str(retry.get("scope", ""))
            and retry.get("failed_receipt_sha256") == digest(
                FAILED_PROFILE_DIR / "profile-r1-xls-owned.receipt.json"),
            "profile retry plan binding differs")
    require(isinstance(frozen, dict)
            and set(frozen) == {"profile-retry-plan.json", "capture_profiles.py"}
            and frozen.get("profile-retry-plan.json") == digest(PROFILE_RETRY_PLAN)
            and frozen.get("capture_profiles.py") == digest(PROFILE_RETRY_SCRIPT),
            "profile retry frozen-input binding differs")
    failed_name = "profile-r1-xls-owned"
    failed_receipt_path = need(FAILED_PROFILE_DIR / f"{failed_name}.receipt.json",
                               "failed profile receipt")
    failed_job = capture_jobs(plan, "profile")[0]
    failed_command = capture_command(failed_job, plan)
    failed_command.pop(failed_command.index("--vgdb=no"))
    failed, _interval = validate_receipt(
        failed_receipt_path, failed_name, manifest_sha, failed_command,
        normal["sha256"], {
            f"{failed_name}.host.json", f"{failed_name}.stdout",
            f"{failed_name}.stderr", f"{failed_name}.callgrind",
        }, allow_failure=True,
    )
    require(failed.get("exit_code") != 0,
            "failed profile launch unexpectedly succeeded")
    failed_artifacts = retry.get("failed_artifacts")
    require(isinstance(failed_artifacts, dict)
            and set(failed_artifacts) == {
                f"{failed_name}.receipt.json", f"{failed_name}.stderr",
                f"{failed_name}.callgrind", f"{failed_name}.stdout",
                f"{failed_name}.host.json"},
            "failed profile artifact inventory differs")
    require(failed.get("artifacts") == {
                filename: expected for filename, expected in failed_artifacts.items()
                if filename != f"{failed_name}.receipt.json"
            },
            "failed profile receipt artifact inventory differs")
    for filename, expected in failed_artifacts.items():
        path = FAILED_PROFILE_DIR / filename
        require(path.is_file() and not path.is_symlink()
                and valid_digest(expected) and digest(path) == expected,
                f"failed profile artifact custody differs: {filename}")
    validate_host(FAILED_PROFILE_DIR / f"{failed_name}.host.json",
                  "failed profile host")
    return {"status": "pass", "retry_plan_sha256": digest(PROFILE_RETRY_PLAN),
            "frozen_sha256": digest(PROFILE_RETRY_FROZEN),
            "failed_receipt_sha256": digest(failed_receipt_path),
            "retry_argument": "--vgdb=no"}


def validate_captures(plan: dict[str, Any], manifest_sha: str,
                      identities: dict[str, dict[str, Any]]) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    counts: dict[str, int] = {}
    retry = validate_profile_retry(plan, identities["normal"], manifest_sha)
    for lane in ("native", "alloc", "profile", "hardware"):
        jobs = capture_jobs(plan, lane)
        kind = "alloc" if lane == "alloc" else "normal"
        lane_intervals: list[tuple[str, tuple[dt.datetime, dt.datetime]]] = []
        for job in jobs:
            path = need(BASELINE / f"{job['name']}.receipt.json",
                        f"baseline {job['name']} receipt")
            _value, interval = validate_receipt(
                path, job["name"], manifest_sha, capture_command(job, plan),
                identities[kind]["sha256"], capture_artifacts(job),
                allow_failure=(lane == "hardware"),
            )
            intervals.append(interval)
            lane_intervals.append((job["name"], interval))
            if lane == "profile":
                binding_path = need(BASELINE / f"{job['name']}.profile-binding.json",
                                    f"{job['name']} retry binding")
                binding = read_json(binding_path, relative(binding_path))
                require(isinstance(binding, dict)
                        and set(binding) == {"retry_plan_sha256", "capture_script_sha256",
                                             "receipt_sha256"}
                        and binding.get("retry_plan_sha256") == digest(PROFILE_RETRY_PLAN)
                        and binding.get("capture_script_sha256") == digest(PROFILE_RETRY_SCRIPT)
                        and binding.get("receipt_sha256") == digest(path),
                        f"{job['name']} retry binding differs")
            counts[lane] = counts.get(lane, 0) + 1
        observed_order = [name for _start, name in
                          sorted((interval[0], name) for name, interval in lane_intervals)]
        require(observed_order == [job["name"] for job in jobs],
                f"{lane} receipt timeline order differs")
    return {"status": "pass", "jobs": counts, "profile_retry": retry,
            "receipts": sum(counts.values())}, intervals


def validate_assembly(plan: dict[str, Any], manifest_sha: str,
                      normal: dict[str, Any]) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    """Validate symbol and disassembly receipts produced before capture."""
    plan_value = read_json(ASSEMBLY_PLAN, "assembly-plan.json")
    require(isinstance(plan_value, dict)
            and plan_value.get("schema") == "litchi-0532-assembly-plan-v1"
            and plan_value.get("status") == "frozen-before-inspection"
            and plan_value.get("owners") == ["claim_sector",
                                               "validate_stream_allocations",
                                               "validate_physical_sector_layout"]
            and plan_value.get("binary_sha256") == normal["sha256"]
            and plan_value.get("source_manifest_sha256") == manifest_sha
            and plan_value.get("script_sha256") == digest(ASSEMBLY_SCRIPT),
            "assembly plan binding differs")
    index = read_json(ASSEMBLY_INDEX, "assembly-index.json")
    require(isinstance(index, dict)
            and index.get("schema") == "litchi-0532-assembly-index-v1"
            and index.get("plan_sha256") == digest(ASSEMBLY_PLAN),
            "assembly index envelope differs")
    symbols_path = need(BASELINE / "symbols.receipt.json", "symbols receipt")
    _symbols_value, symbols_interval = validate_receipt(
        symbols_path, "symbols", manifest_sha,
        ["nm", "-S", "--defined-only", str(SCRATCH / "normal")],
        normal["sha256"], {"symbols.host.json", "symbols.stdout", "symbols.stderr"},
    )
    require(index.get("symbols_receipt_sha256") == digest(symbols_path),
            "assembly index symbols receipt binding differs")
    symbols = read_text(BASELINE / "symbols.stdout", "symbols stdout")
    symbol_rows = {}
    for line in symbols.splitlines():
        fields = line.split()
        if len(fields) == 4 and fields[2] in ("t", "T"):
            try:
                symbol_rows[fields[3]] = (int(fields[0], 16), int(fields[1], 16))
            except ValueError:
                continue
    rows = index.get("rows")
    require(isinstance(rows, list) and rows, "assembly index rows are missing")
    owners = {"claim_sector", "validate_stream_allocations",
              "validate_physical_sector_layout"}
    seen_names: set[str] = set()
    seen_symbols: set[str] = set()
    intervals = [symbols_interval]
    for row in rows:
        require(isinstance(row, dict)
                and set(row) == {"name", "owner", "symbol", "address_hex",
                                  "size_bytes", "receipt_sha256"},
                "assembly index row schema differs")
        name, owner, symbol = row["name"], row["owner"], row["symbol"]
        require(isinstance(name, str) and name.startswith("assembly-")
                and name not in seen_names and owner in owners
                and isinstance(symbol, str) and symbol not in seen_symbols
                and valid_digest(row["receipt_sha256"])
                and re.fullmatch(r"[0-9a-fA-F]+", row["address_hex"]) is not None
                and isinstance(row["size_bytes"], int) and row["size_bytes"] > 0,
                "assembly index row identity differs")
        require(symbol in symbol_rows
                and symbol_rows[symbol] == (int(row["address_hex"], 16),
                                             row["size_bytes"]),
                f"assembly symbol is not bound by nm output: {symbol}")
        receipt_path = need(BASELINE / f"{name}.receipt.json", f"{name} receipt")
        _value, interval = validate_receipt(
            receipt_path, name, manifest_sha,
            ["objdump", "-d", "--disassemble=" + symbol, str(SCRATCH / "normal")],
            normal["sha256"], {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"},
        )
        require(row["receipt_sha256"] == digest(receipt_path),
                f"{name} receipt hash differs")
        seen_names.add(name)
        seen_symbols.add(symbol)
        intervals.append(interval)
    require({row["owner"] for row in rows} == owners,
            "assembly index omits an attribution owner")
    return ({"status": "pass", "plan_sha256": digest(ASSEMBLY_PLAN),
             "index_sha256": digest(ASSEMBLY_INDEX), "symbols": 1,
             "disassemblies": len(rows), "intervals": len(intervals),
             "symbols_receipt": digest(symbols_path)}, intervals)


def check_serial(intervals: list[tuple[dt.datetime, dt.datetime]], label: str) -> None:
    ordered = sorted(intervals, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{label} receipt intervals overlap")


def replay_analyzer(script: Path, retained: Path, *args: str) -> dict[str, Any]:
    need(script, relative(script))
    need(retained, relative(retained))
    # /tmp was independently observed to return EDQUOT during this campaign;
    # use a transient directory on the parent home filesystem so post-cleanup
    # replay remains possible without changing the sealed evidence directory.
    with tempfile.TemporaryDirectory(prefix="litchi-0532-replay-", dir="/home/zhuhe") as folder:
        output = Path(folder) / retained.name
        command = [sys.executable, "-B", str(script), str(output), *args]
        environment = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
        try:
            result = subprocess.run(command, cwd=REPO, env=environment,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    text=True, check=False)
        except OSError as error:
            raise VerificationError(f"{relative(script)} replay could not start: {error}") from error
        require(result.returncode == 0,
                f"{relative(script)} replay failed: {result.stderr[-2000:]}")
        require(output.is_file() and output.read_bytes() == retained.read_bytes(),
                f"{relative(script)} replay differs from {relative(retained)}")
    return read_json(retained, relative(retained))


def replay_pure_function(script: Path, retained: Path, function: str) -> dict[str, Any]:
    """Replay a report producer whose public API returns the document."""
    need(script, relative(script))
    need(retained, relative(retained))
    try:
        spec = importlib.util.spec_from_file_location(
            "litchi_0532_" + script.stem.replace("-", "_"), script)
        require(spec is not None and spec.loader is not None,
                f"cannot load {relative(script)}")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        producer = getattr(module, function)
        value = producer()
    except (OSError, ValueError, KeyError, TypeError, AttributeError,
            AssertionError) as error:
        raise VerificationError(f"{relative(script)} pure replay failed: {error}") from error
    expected = read_json(retained, relative(retained))
    require(value == expected, f"{relative(script)} pure replay differs from {relative(retained)}")
    rendered = (json.dumps(value, indent=2) + "\n").encode()
    require(rendered == retained.read_bytes(),
            f"{relative(script)} pure replay serialization differs from {relative(retained)}")
    return expected


def validate_analysis(plan: dict[str, Any], captures: dict[str, Any]) -> dict[str, Any]:
    value = replay_analyzer(ANALYZER, ANALYSIS)
    require(isinstance(value, dict) and value.get("status") == "pass"
            and value.get("plan_sha256") == digest(PLAN)
            and value.get("performance_claim") is None
            and isinstance(value.get("rows"), list)
            and value.get("native_samples") == 24000
            and value.get("allocation_samples") == 720
            and isinstance(value.get("same_build_variations_over_five_percent"), list),
            "native/allocation analysis envelope differs")
    return {"status": "pass", "path": relative(ANALYSIS), "sha256": digest(ANALYSIS),
            "native_samples": value.get("native_samples"),
            "allocation_samples": value.get("allocation_samples"),
            "rows": len(value["rows"]),
            "same_build_variations": len(value["same_build_variations_over_five_percent"]),
            "captures": captures}


def validate_profile(plan: dict[str, Any]) -> dict[str, Any]:
    value = replay_analyzer(PROFILE_ANALYZER, PROFILE_ANALYSIS)
    require(isinstance(value, dict)
            and value.get("schema") == "cfb_ole2_constructor_callgrind_profile_analysis_v1"
            and value.get("status") == "pass"
            and value.get("scope") == plan["scope"]
            and value.get("plan_sha256") == digest(PLAN)
            and value.get("stage_selection") == ["baseline"]
            and value.get("performance_claim") == "none"
            and value.get("profile_count") == 8
            and value.get("timed_constructor_dump_count") == 40
            and value.get("setup_dump_count") == 6
            and value.get("comparison") is None
            and value.get("validation", {}).get("no_speedup_comparison_performed") is True,
            "profile analysis envelope differs")
    return {"status": "pass", "path": relative(PROFILE_ANALYSIS),
            "sha256": digest(PROFILE_ANALYSIS), "profiles": value["profile_count"],
            "timed_constructor_dumps": value.get("timed_constructor_dump_count"),
            "setup_dumps": value.get("setup_dump_count")}


def validate_hardware(plan: dict[str, Any]) -> dict[str, Any]:
    value = replay_analyzer(HARDWARE_ANALYZER, HARDWARE_ANALYSIS)
    require(isinstance(value, dict) and value.get("status") == "pass"
            and isinstance(value.get("captures"), list)
            and len(value["captures"]) == 2
            and value.get("scope") == plan["hardware"]["scope"]
            and value.get("latency_samples_excluded_from_native") == 2000
            and value.get("no_operation_local_hardware_or_speedup_claim") is True,
            "hardware analysis envelope differs")
    for row in value["captures"]:
        require(isinstance(row, dict) and row.get("name", "").startswith("hardware-r"),
                "hardware analysis row differs")
        require(row.get("status") in ("measured", "unavailable", "unavailable_for_group_claim"),
                "hardware analysis status differs")
    return {"status": "pass", "path": relative(HARDWARE_ANALYSIS),
            "sha256": digest(HARDWARE_ANALYSIS), "captures": len(value["captures"]),
            "statuses": [row["status"] for row in value["captures"]]}


def validate_mechanism(plan: dict[str, Any], source: dict[str, Any],
                       profile: dict[str, Any], analysis: dict[str, Any],
                       assembly: dict[str, Any]) -> dict[str, Any]:
    value = replay_pure_function(MECHANISM_ANALYZER, MECHANISM_ANALYSIS, "analyze")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0532-claim-mechanism-v1"
            and value.get("status") == "pass"
            and value.get("performance_claim") == "none",
            "mechanism analysis envelope differs")
    bindings = value.get("bindings")
    require(bindings == {
        "profile-analysis.json": digest(PROFILE_ANALYSIS),
        "analysis.json": digest(ANALYSIS),
        "assembly-index.json": digest(ASSEMBLY_INDEX),
        "baseline/source-manifest.json": digest(BASELINE / "source-manifest.json"),
    }, "mechanism analysis bindings differ")
    geometry = value.get("source_geometry")
    require(geometry == {
        "generator": "tools/perf-baseline/src/lib.rs",
        "writer_default": "crates/litchi-cfb/src/writer/core.rs",
        "sector_bytes": 512,
    }, "mechanism source geometry differs")
    code = value.get("code")
    require(isinstance(code, list) and len(code) == 6,
            "mechanism assembly attribution differs")
    for row in code:
        require(isinstance(row, dict)
                and row.get("bytes") == 363
                and row.get("stack_reservation_bytes") == 112
                and row.get("success_path_instruction_count") == 15,
                "mechanism code-shape attribution differs")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == profile["profiles"],
            "mechanism profile equation inventory differs")
    return {"status": "pass", "path": relative(MECHANISM_ANALYSIS),
            "sha256": digest(MECHANISM_ANALYSIS), "assembly_variants": len(code),
            "profile_rows": len(rows), "performance_claim": "none"}


def validate_variation_review(analysis: dict[str, Any]) -> dict[str, Any]:
    path = need(VARIATION_REVIEW, "variation-review.json")
    value = read_json(path, "variation-review.json")
    raw = read_json(ANALYSIS, "analysis.json").get("same_build_variations_over_five_percent")
    reviewed = value.get("variations") if isinstance(value, dict) else None
    require(isinstance(raw, list) and isinstance(reviewed, list)
            and value.get("schema") == "litchi-0532-variation-review-v1"
            and value.get("analysis_sha256") == analysis["sha256"]
            and value.get("threshold_percent") == 5
            and isinstance(value.get("review"), str)
            and value["review"].strip()
            and len(reviewed) == len(raw),
            "same-build variation review binding differs")
    remaining = list(reviewed)
    for item in raw:
        require(isinstance(item, dict), "analysis variation row is malformed")
        index = next((i for i, candidate in enumerate(remaining)
                      if isinstance(candidate, dict)
                      and all(candidate.get(key) == val for key, val in item.items())), None)
        require(index is not None, "same-build variation is not individually reviewed")
        candidate = remaining.pop(index)
        explanation = candidate.get("review", candidate.get("reason"))
        require(isinstance(explanation, str) and explanation.strip(),
                "same-build variation review lacks an explanation")
    require(not remaining
            and (value.get("complete") in (None, True)),
            "same-build variation review has extra or incomplete rows")
    return {"status": "pass", "path": relative(path), "sha256": digest(path),
            "variations": len(raw)}


def expected_quality_commands() -> dict[str, list[str]]:
    value = read_json(QUALITY_PLAN, "quality-plan.json")
    require(isinstance(value, dict) and value.get("status") == "frozen-before-quality"
            and value.get("script_sha256") == digest(QUALITY_CHECKS),
            "quality plan envelope differs")
    rows = value.get("commands")
    require(isinstance(rows, list) and [row[0] for row in rows] == list(QUALITY_NAMES),
            "quality command inventory differs")
    result: dict[str, list[str]] = {}
    expected = {
        "cfb-tests": ["cargo", "test", "--locked", "-p", "litchi-cfb",
                       "--all-features", "--", "--test-threads=2"],
        "cfb-no-default-tests": ["cargo", "test", "--locked", "-p", "litchi-cfb",
                                 "--no-default-features", "--", "--test-threads=2"],
        "xls-tests": ["cargo", "test", "--locked", "-p", "litchi-xls",
                       "--all-features", "--", "--test-threads=2"],
        "owner-clippy": ["cargo", "clippy", "--locked", "-p", "litchi-cfb", "-p",
                          "litchi-xls", "--all-features", "--lib", "--", "-D", "warnings"],
        "owner-rustdoc": ["cargo", "doc", "--locked", "-p", "litchi-cfb", "-p",
                           "litchi-xls", "--all-features", "--no-deps"],
    }
    for row in rows:
        require(isinstance(row, list) and len(row) == 2 and row[0] in expected
                and row[1] == expected[row[0]], "quality command row differs")
        result[row[0]] = row[1]
    return result


def replay_prior_quality_and_seal() -> dict[str, Any]:
    """Run the prior bundle's own restored-quality and seal validators."""
    need(PRIOR_VERIFIER, "0531 verifier")
    try:
        spec = importlib.util.spec_from_file_location(
            "litchi_0531_quality_replay", PRIOR_VERIFIER)
        require(spec is not None and spec.loader is not None,
                "cannot load 0531 verifier")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        quality = module.validate_quality("restored")
        seal = module.validate_seal()
    except (OSError, ValueError, KeyError, TypeError, AttributeError) as error:
        raise VerificationError(f"0531 restored quality/seal replay failed: {error}") from error
    require(isinstance(quality, dict) and quality.get("stage") == "restored"
            and quality.get("checks") == 8
            and quality.get("executed_tests") == 4757,
            "0531 restored quality replay envelope differs")
    require(isinstance(seal, dict), "0531 seal replay envelope differs")
    return {"quality": quality, "seal": seal}


def validate_quality_reuse(manifest_sha: str) -> dict[str, Any]:
    value = read_json(QUALITY_REUSE, "quality-reuse.json")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0532-quality-reuse-v1"
            and value.get("source_manifest_sha256") == manifest_sha
            and value.get("prior_bundle") == "change-0531"
            and value.get("prior_source_manifest") == "restored/source-manifest.json"
            and value.get("prior_source_manifest_sha256") == digest(PRIOR_RESTORED_MANIFEST)
            and value.get("prior_quality_summary") == "restored-quality-summary.json"
            and value.get("prior_quality_summary_sha256") == digest(PRIOR_RESTORED_QUALITY)
            and value.get("prior_verifier_sha256") == digest(PRIOR_VERIFIER)
            and value.get("prior_seal_sha256") == digest(PRIOR_SEAL)
            and value.get("source_maps_equal") is True
            and value.get("successful_test_executions") == 4757,
            "quality reuse binding differs")
    prior_quality = read_json(PRIOR_RESTORED_QUALITY, "0531 restored quality summary")
    require(prior_quality.get("status") == "pass"
            and prior_quality.get("stage") == "restored"
            and len(prior_quality.get("checks", [])) == 8
            and prior_quality.get("successful_test_executions") == 4757,
            "prior restored quality summary differs")
    replay = replay_prior_quality_and_seal()
    return {"status": "pass", "path": relative(QUALITY_REUSE),
            "sha256": digest(QUALITY_REUSE), "prior_tests": 4757,
            "prior_replay": replay}


def quality_prefix() -> list[str]:
    return ["env", "TMPDIR=" + str(TARGET / "tmp"),
            "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
            "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings"]


def validate_quality(plan: dict[str, Any], manifest_sha: str) -> tuple[dict[str, Any], list[tuple[dt.datetime, dt.datetime]]]:
    commands = expected_quality_commands()
    reuse = validate_quality_reuse(manifest_sha)
    summary = read_json(QUALITY_SUMMARY, "quality-summary.json")
    require(isinstance(summary, dict) and summary.get("status") == "pass"
            and summary.get("source_manifest_sha256") == manifest_sha
            and summary.get("checks_script_sha256") == digest(QUALITY_CHECKS),
            "quality summary envelope differs")
    rows = summary.get("checks")
    require(isinstance(rows, list) and len(rows) == len(commands),
            "quality summary checks differ")
    require(all(isinstance(item, dict) for item in rows),
            "quality summary rows are malformed")
    require([item.get("name") for item in rows] ==
            ["check-" + name for name in commands],
            "quality summary check order differs")
    intervals: list[tuple[dt.datetime, dt.datetime]] = []
    seen: set[str] = set()
    total = 0
    for item in rows:
        require(isinstance(item, dict)
                and set(item) == {"name", "receipt_sha256", "executed_tests"},
                "quality summary row schema differs")
        name = item["name"]
        require(name in {"check-" + value for value in commands} and name not in seen,
                "quality summary row name differs")
        seen.add(name)
        command_name = name.removeprefix("check-")
        receipt_path = need(BASELINE / f"{name}.receipt.json", f"{name} receipt")
        command = quality_prefix() + commands[command_name]
        value, interval = validate_receipt(
            receipt_path, name, manifest_sha, command, None,
            {f"{name}.host.json", f"{name}.stdout", f"{name}.stderr"},
        )
        intervals.append(interval)
        require(item["receipt_sha256"] == digest(receipt_path)
                and value.get("exit_code") == 0,
                f"{name} quality receipt binding differs")
        stdout = BASELINE / f"{name}.stdout"
        count = sum(int(number) for number in re.findall(
            r"test result: ok\. (\d+) passed;", read_text(stdout, relative(stdout))))
        require(item["executed_tests"] == count, f"{name} test count differs")
        total += count
    require(seen == {"check-" + value for value in commands}
            and summary.get("executed_tests") == total,
            "quality aggregate differs")
    return {"status": "pass", "path": relative(QUALITY_SUMMARY),
            "sha256": digest(QUALITY_SUMMARY), "checks": len(rows),
            "executed_tests": total, "reuse": reuse}, intervals


def expected_receipt_names(plan: dict[str, Any]) -> set[str]:
    names = {"build-normal", "build-alloc", "symbols"}
    for lane in ("native", "alloc", "profile", "hardware"):
        names.update(job["name"] for job in capture_jobs(plan, lane))
    names.update("check-" + name for name in QUALITY_NAMES)
    assembly = read_json(ASSEMBLY_INDEX, "assembly-index.json")
    rows = assembly.get("rows") if isinstance(assembly, dict) else None
    require(isinstance(rows, list), "assembly index rows are missing")
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("name"), str),
                "assembly index receipt name is malformed")
        names.add(row["name"])
    return names


def validate_receipt_inventory(plan: dict[str, Any]) -> dict[str, Any]:
    """Require exactly the successful receipt set expected by the frozen plan."""
    expected = expected_receipt_names(plan)
    actual = {
        path.name.removesuffix(".receipt.json")
        for path in BASELINE.glob("*.receipt.json")
    }
    require(actual == expected,
            f"baseline receipt inventory differs: expected {len(expected)}, found {len(actual)}")
    return {"status": "pass", "successful_receipts": len(actual),
            "receipt_names_sha256": hashlib.sha256(
                ("\n".join(sorted(actual)) + "\n").encode()).hexdigest()}


def validate_receipt_timeline() -> dict[str, Any]:
    """Check non-overlap for all successful receipts and the retained failed attempt."""
    intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for path in BASELINE.glob("*.receipt.json"):
        value = read_json(path, relative(path))
        start, end = receipt_interval(value, relative(path))
        intervals.append((start, end, path.name.removesuffix(".receipt.json")))
    failed_path = FAILED_PROFILE_DIR / "profile-r1-xls-owned.receipt.json"
    failed = read_json(failed_path, "failed profile receipt")
    start, end = receipt_interval(failed, relative(failed_path))
    intervals.append((start, end, "failed-profile-initial/profile-r1-xls-owned"))
    ordered = sorted(intervals, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "retained receipt intervals overlap")
    return {"status": "pass", "successful_receipts": len(intervals) - 1,
            "retained_failed_attempts": 1, "serial_intervals": len(ordered),
            "first": ordered[0][2], "last": ordered[-1][2],
            "order": [item[2] for item in ordered]}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict)
            and value.get("plan_sha256") == digest(PLAN)
            and value.get("removed") == plan["owned_paths"]
            and value.get("accessible_process_references") == []
            and value.get("owned_paths_absent") is True
            and value.get("python_cache_absent") is True,
            "cleanup receipt differs")
    require(all(not os.path.lexists(path) for path in plan["owned_paths"])
            and not list(HERE.rglob("__pycache__")),
            "owned path or Python cache remains")
    return {"status": "pass", "path": relative(CLEANUP), "sha256": digest(CLEANUP),
            "owned_paths_absent": True, "python_cache_absent": True}


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
    actual = {relative(item): digest(item) for item in HERE.rglob("*")
              if item.is_file() and not item.is_symlink() and item != path}
    require(expected == actual and not any(item.is_symlink() for item in HERE.rglob("*")),
            "SHA256SUMS inventory differs")
    return {"status": "pass", "path": relative(path), "sha256": digest(path),
            "entries": len(expected)}


def component(name: str) -> dict[str, Any]:
    plan = check_plan()
    source = validate_source(plan)
    if name == "source":
        return source
    identities, build_intervals = validate_builds(plan, source["manifest_sha256"])
    if name in ("build", "builds"):
        check_serial(build_intervals, "build")
        return {"status": "pass", "builds": identities,
                "intervals": len(build_intervals)}
    assembly, assembly_intervals = validate_assembly(
        plan, source["manifest_sha256"], identities["normal"])
    if name == "assembly":
        check_serial(build_intervals + assembly_intervals, "build/assembly")
        return {"status": "pass", "builds": identities, "assembly": assembly,
                "intervals": len(build_intervals) + len(assembly_intervals)}
    captures, capture_intervals = validate_captures(
        plan, source["manifest_sha256"], identities)
    if name in ("captures", "native", "allocation"):
        check_serial(build_intervals + assembly_intervals + capture_intervals,
                     "build/assembly/capture")
        return {"status": "pass", "builds": identities, "captures": captures,
                "assembly": assembly,
                "intervals": (len(build_intervals) + len(assembly_intervals) +
                              len(capture_intervals))}
    analysis = validate_analysis(plan, captures)
    if name == "analysis":
        return analysis
    profile = validate_profile(plan)
    if name == "profile":
        return profile
    mechanism = validate_mechanism(plan, source, profile, analysis, assembly)
    if name == "mechanism":
        return mechanism
    hardware = validate_hardware(plan)
    if name == "hardware":
        return hardware
    variation = validate_variation_review(analysis)
    if name == "variation":
        return variation
    quality, quality_intervals = validate_quality(plan, source["manifest_sha256"])
    if name == "quality":
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals, "build/assembly/capture/quality")
        return quality
    if name == "cleanup":
        return validate_cleanup(plan)
    if name == "seal":
        return validate_seal()
    if name == "all":
        receipt_inventory = validate_receipt_inventory(plan)
        timeline = validate_receipt_timeline()
        check_serial(build_intervals + assembly_intervals + capture_intervals +
                     quality_intervals, "build/assembly/capture/quality")
        cleanup = validate_cleanup(plan)
        seal = validate_seal()
        return {"status": "pass", "source": source, "builds": identities,
                "assembly": assembly, "captures": captures,
                "analysis": analysis, "profile": profile, "mechanism": mechanism,
                "hardware": hardware, "variation": variation, "quality": quality,
                "receipt_inventory": receipt_inventory, "timeline": timeline,
                "cleanup": cleanup, "seal": seal}
    raise VerificationError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    if selected == "all":
        # ``component('all')`` performs one bounded pass, avoiding repeated
        # expensive analyzer replays while retaining every sub-result.
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
        except (VerificationError, OSError, subprocess.CalledProcessError) as error:
            results[name] = {"status": "fail", "error": str(error)}
    statuses = [row["status"] for row in results.values()]
    status = ("fail" if "fail" in statuses else
              "incomplete" if "incomplete" in statuses else "pass")
    return {"schema": "litchi-0532-measurement-verification-v1", "status": status,
            "scope": selected,
            "performance_claim": "none; baseline attribution and diagnostics only",
            "components": results}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=(
        "all", "source", "build", "builds", "assembly", "captures", "native", "allocation",
        "analysis", "profile", "mechanism", "hardware", "variation", "quality", "cleanup", "seal",
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
