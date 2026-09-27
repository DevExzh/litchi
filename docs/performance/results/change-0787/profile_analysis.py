#!/usr/bin/env python3
"""Offline analysis of the retained 0787 profile and thread-trace packet.

The profile lane is a Callgrind guest-instruction diagnostic for one ordered
Part read.  The trace lane is a separate whole-child syscall census.  This
module never starts a benchmark, invokes a profiler, invokes ``git``, or
interprets a trace duration as native timing.  It verifies the retained
receipts, parses the raw Callgrind graph through the previously qualified
0784 parser, and writes a relocation-stable JSON/Markdown report.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
import shlex
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable, Iterator


HERE = Path(__file__).resolve().parent
OWNER = "cached_part_profile::cached_part_region_0787"
CHILD = "litchi_opc::source_backed::batch::read_parts_ordered"
TRACE_STATES = ("fresh", "primed")
WIDTHS = (1, 8, 32)
PROFILE_ORDER = ((1, 8, 32), (32, 8, 1))
TRACE_ORDER = ((1, 8, 32), (32, 8, 1))
PROFILE_ANALYSIS_JSON = HERE / "profile-analysis.json"
PROFILE_ANALYSIS_MD = HERE / "profile-analysis.md"
REPORT_SCHEMA = "litchi.execution-baseline.v1"
ANALYSIS_SCHEMA = "litchi.cached-part-profile-analysis.0787.v1"

# The parser is deliberately reused from the already reviewed 0784 packet.
# Both hashes are checked before loading it: SHA-256 gives a convenient file
# identity while the Git blob SHA-1 binds the exact original committed blob.
PARSER_RELATIVE = Path("../change-0784/profile_analysis.py")
PARSER_SHA256 = "70b0bb3fc665860a9aa586725d72ced75e41de06f95e3b1c5408c70998e926b6"
PARSER_BLOB_SHA1 = "1c42922ea9ab5f3624f294f4f0e5f615775c234e"

HEX = frozenset("0123456789abcdefABCDEF")


class EvidenceError(RuntimeError):
    """A retained artifact is absent, stale, malformed, or contradictory."""


def fail(message: str) -> None:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def git_blob_sha1(path: Path) -> str:
    try:
        data = path.read_bytes()
    except OSError as error:
        fail(f"cannot read parser source {path}: {error}")
    header = f"blob {len(data)}\0".encode()
    return hashlib.sha1(header + data).hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def read_text(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing text evidence: {path}")
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read text evidence {path}: {error}")


def rel(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(HERE.resolve()))
    except ValueError:
        return str(path)


def json_digest(value: Any) -> str:
    encoded = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def file_identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    marker = "/docs/performance/results/change-0787/"
    text = raw.replace("\\", "/")
    if marker in text:
        candidate = HERE / text.split(marker, 1)[1]
    elif text.startswith("docs/performance/results/change-0787/"):
        candidate = HERE / text.split("change-0787/", 1)[1]
    else:
        candidate = Path(raw) if Path(raw).is_absolute() else HERE / raw
    candidate = candidate.resolve(strict=False)
    try:
        candidate.relative_to(HERE.resolve())
    except ValueError:
        fail(f"{label} escapes packet: {raw}")
    return candidate


def artifact(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    expected_bytes = value.get("bytes", value.get("size"))
    expected_sha = value.get("sha256", value.get("digest"))
    require(type(expected_bytes) is int and expected_bytes >= 0,
            f"{label}.bytes is invalid")
    require(is_sha(expected_sha), f"{label}.sha256 is invalid")
    path = packet_path(value.get("path"), label)
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    require(path.stat().st_size == expected_bytes, f"{label}.bytes changed")
    require(sha256(path) == expected_sha, f"{label}.sha256 changed")
    return {"path": rel(path), "bytes": expected_bytes, "sha256": expected_sha}


def _walk(value: Any, path: tuple[str, ...] = ()) -> Iterator[tuple[tuple[str, ...], Any]]:
    yield path, value
    if isinstance(value, dict):
        for key, child in value.items():
            yield from _walk(child, path + (str(key),))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _walk(child, path + (str(index),))


def load_cleanup() -> Any | None:
    path = HERE / "cleanup.json"
    if not path.is_file() or path.is_symlink():
        return None
    value = read_json(path)
    require(isinstance(value, dict), "cleanup.json is malformed")
    require(value.get("verified_before_removal") is True
            or value.get("executables_verified_before_removal") is True,
            "cleanup witness was not verified before removal")
    return value


def _cleanup_match(cleanup: Any, expected: dict[str, Any], name: str) -> bool:
    for path, value in _walk(cleanup):
        if not isinstance(value, dict):
            continue
        digest = value.get("sha256", value.get("digest"))
        size = value.get("bytes", value.get("size"))
        raw_path = value.get("path")
        if digest != expected["sha256"] or size != expected["bytes"]:
            continue
        if raw_path is None or Path(str(raw_path)).name == name:
            return True
    return False


def external_identity(value: Any, label: str) -> dict[str, Any]:
    """Verify a live executable or its exact post-cleanup witness.

    The path is intentionally reduced to its basename in the report.  This
    keeps replay stable when the target directory is removed or relocated.
    """

    require(isinstance(value, dict), f"{label} binary receipt is malformed")
    raw = value.get("path")
    expected = {
        "bytes": value.get("bytes", value.get("size")),
        "sha256": value.get("sha256", value.get("digest")),
    }
    require(isinstance(raw, str) and raw, f"{label} binary path is missing")
    require(type(expected["bytes"]) is int and expected["bytes"] > 0,
            f"{label} binary bytes are invalid")
    require(is_sha(expected["sha256"]), f"{label} binary SHA is invalid")
    path = Path(raw)
    name = path.name
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == expected["bytes"], f"{label} binary size changed")
        require(sha256(path) == expected["sha256"], f"{label} binary hash changed")
    else:
        cleanup = load_cleanup()
        require(cleanup is not None and _cleanup_match(cleanup, expected, name),
                f"{label} binary is missing without an exact cleanup witness")
    return {"name": name, **expected, "custody_verified": True}


def _normalise_path_token(token: str) -> str:
    marker = "/docs/performance/results/change-0787/"
    text = token.replace("\\", "/")
    prefix = ""
    body = text
    if "=" in text and text.split("=", 1)[0].startswith("--"):
        prefix, body = text.split("=", 1)
        prefix += "="
    if marker in body:
        return prefix + body.split(marker, 1)[1]
    if body.startswith("docs/performance/results/change-0787/"):
        return prefix + body.split("change-0787/", 1)[1]
    return token


def canonical_command(command: Any) -> list[str]:
    require(isinstance(command, list) and all(isinstance(x, str) for x in command),
            "command receipt is malformed")
    result: list[str] = []
    binary_names = {"baseline-profile", "baseline-profile-compat",
                    "before-observer", "after-observer", "before-native", "after-native"}
    for token in command:
        value = _normalise_path_token(token)
        if value == token and "/" in token and Path(token).name in binary_names:
            value = f"<binary:{Path(token).name}>"
        result.append(value)
    return result


def canonical_binary_command(command: Any, binary_name: str) -> list[str]:
    result = canonical_command(command)
    result = [f"<binary:{binary_name}>" if token.endswith("/" + binary_name)
              or token == binary_name else token for token in result]
    return result


def load_parser() -> tuple[Any, dict[str, str]]:
    parser_path = (HERE / PARSER_RELATIVE).resolve()
    require(parser_path.is_file() and not parser_path.is_symlink(),
            f"reused Callgrind parser is missing: {parser_path}")
    actual_sha = sha256(parser_path)
    actual_blob = git_blob_sha1(parser_path)
    require(actual_sha == PARSER_SHA256, "reused Callgrind parser SHA-256 changed")
    require(actual_blob == PARSER_BLOB_SHA1, "reused Callgrind parser Git blob changed")
    spec = importlib.util.spec_from_file_location("profile_analysis_0784_parser", parser_path)
    require(spec is not None and spec.loader is not None, "cannot load reused parser")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The parser's generic relative() helper is bound to its module global.
    module.HERE = HERE
    return module, {"path": "../change-0784/profile_analysis.py", "sha256": actual_sha,
                    "git_blob_sha1": actual_blob}


def require_hex_revision(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 40
            and all(c in HEX for c in value), f"{label} revision is invalid")


def load_origin() -> dict[str, Any]:
    value = read_json(HERE / "origin.json")
    require(isinstance(value, dict), "origin.json is malformed")
    require_hex_revision(value.get("base"), "origin")
    return {"base_revision": value["base"],
            "owned_worktree_recorded": isinstance(value.get("owned_worktree"), str),
            "target_name": Path(str(value.get("target", ""))).name}


def load_plans() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    profile = read_json(HERE / "profile-plan.json")
    trace = read_json(HERE / "trace-plan.json")
    require(isinstance(profile, dict) and profile.get("schema") ==
            "litchi.cached-part-profile.0787.v1", "profile plan schema changed")
    require(profile.get("events") == ["Ir"] and profile.get("collect_at_start") is False,
            "profile event/collection contract changed")
    require(profile.get("orders") == [list(x) for x in PROFILE_ORDER]
            and profile.get("repeats") == 2 and profile.get("samples") == 1
            and profile.get("warmup") == 0 and profile.get("widths") == list(WIDTHS),
            "profile order/sample contract changed")
    require(profile.get("shape") == "large" and profile.get("state") == "primed"
            and profile.get("task_floor") == 0
            and profile.get("owner_suffix") == "cached_part_region_0787",
            "profile case contract changed")
    require(isinstance(trace, dict) and trace.get("schema") ==
            "litchi.cached-part-thread-trace.0787.v1", "trace plan schema changed")
    require(trace.get("repeats") == 2 and trace.get("samples") == 1
            and trace.get("warmup") == 0 and trace.get("widths") == list(WIDTHS)
            and trace.get("states") == list(TRACE_STATES),
            "trace order/sample contract changed")
    require(trace.get("shape") == "large" and trace.get("task_floor") == 0,
            "trace case contract changed")
    plans = {
        "profile": {"file": file_identity(HERE / "profile-plan.json"),
                    "sha256": sha256(HERE / "profile-plan.json"),
                    "value": profile},
        "trace": {"file": file_identity(HERE / "trace-plan.json"),
                   "sha256": sha256(HERE / "trace-plan.json"),
                   "value": trace},
    }
    return profile, trace, plans


def _production_identity() -> dict[str, Any]:
    source_path = HERE / "build-before" / "source.json"
    source = read_json(source_path)
    require(isinstance(source, dict), "build-before/source.json is malformed")
    production = source.get("production")
    require(isinstance(production, dict), "production source census is missing")
    require_hex_revision(production.get("revision"), "production source")
    files = production.get("files")
    require(isinstance(files, dict) and len(files) == 9196,
            "production source census cardinality changed")
    tool = source.get("tool")
    require(isinstance(tool, dict) and len(tool) == 4,
            "perf tool source census changed")
    build_path = HERE / "build-before" / "build.json"
    build = read_json(build_path)
    require(isinstance(build, dict), "build-before/build.json is malformed")
    source_ref = artifact(build.get("source"), "build-before source")
    require(source_ref["path"] == "build-before/source.json",
            "build-before source path changed")
    rows = build.get("rows")
    require(isinstance(rows, list) and len(rows) == 2,
            "build-before build row count changed")
    logs = []
    for row in rows:
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                "build-before build command failed")
        logs.append(artifact(row.get("log"), "build-before build log"))
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "observer"},
            "build-before binary matrix changed")
    checked = {name: external_identity(value, f"build-before {name}")
               for name, value in binaries.items()}
    return {
        "manifest": file_identity(build_path),
        "source": source_ref,
        "production_revision": production["revision"],
        "production_file_count": len(files),
        "production_manifest_sha256": sha256(source_path),
        "tool_file_count": len(tool),
        "tool_manifest_sha256": json_digest(tool),
        "logs": logs,
        "binaries": checked,
    }, production


def _archive_hashes(directory: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    require(directory.is_dir() and not directory.is_symlink(),
            f"profile source archive is missing: {directory}")
    for path in sorted(directory.rglob("*")):
        if path.is_file() and not path.is_symlink():
            result[str(path.relative_to(HERE))] = sha256(path)
    require(set(result) == {
        rel(directory / "Cargo.lock"), rel(directory / "Cargo.toml"),
        rel(directory / "Cargo.toml.template"), rel(directory / "README.md"),
        rel(directory / "src" / "main.rs")
    }, f"{rel(directory)} source file set changed")
    return result


def validate_profile_build(name: str, production: dict[str, Any], profile_sha: str,
                           archive_directory: str) -> dict[str, Any]:
    directory = HERE / name
    receipt_path = directory / "receipt.json"
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{name} receipt is malformed or failed")
    command = canonical_command(receipt.get("command"))
    expected_command = ["cargo", "build", "--offline", "--locked", "--release",
                        "--manifest-path", "profile-src/Cargo.toml"]
    require(command == expected_command, f"{name} build command changed")
    require(receipt.get("plan_sha256") == profile_sha, f"{name} plan hash changed")
    environment = receipt.get("environment")
    require(isinstance(environment, dict) and environment.get("CARGO_BUILD_JOBS") == "2"
            and environment.get("CARGO_INCREMENTAL") == "0",
            f"{name} build environment changed")
    source_ref = artifact(receipt.get("source"), f"{name} source")
    require(source_ref["path"] == f"{name}/source.json", f"{name} source path changed")
    source = read_json(HERE / source_ref["path"])
    require(source == production, f"{name} production source differs from baseline 9196 census")
    probe_ref = artifact(receipt.get("probe"), f"{name} probe")
    require(probe_ref["path"] == f"{name}/probe.json", f"{name} probe path changed")
    probe = read_json(HERE / probe_ref["path"])
    require(isinstance(probe, dict), f"{name} probe is malformed")
    archive_root = HERE / archive_directory
    hashes = _archive_hashes(archive_root)
    expected_probe = {key.replace(f"{archive_directory}/", "profile-src/", 1): value
                      for key, value in hashes.items()}
    require(probe == expected_probe, f"{name} probe does not bind {archive_directory}")
    symbol_ref = artifact(receipt.get("owner_symbol"), f"{name} owner symbol")
    symbol_path = HERE / symbol_ref["path"]
    symbol_lines = read_text(symbol_path).splitlines()
    require(len(symbol_lines) == 1 and symbol_lines[0].endswith(OWNER),
            f"{name} owner symbol is not exact")
    require(receipt.get("owner") == OWNER, f"{name} owner name changed")
    binary = external_identity(receipt.get("binary"), f"{name} profile")
    return {
        "receipt": file_identity(receipt_path),
        "command": expected_command,
        "source": source_ref,
        "probe": probe_ref,
        "probe_archive": {"directory": archive_directory,
                           "files": [{"path": key, "sha256": value}
                                     for key, value in sorted(hashes.items())]},
        "owner": OWNER,
        "owner_symbol": symbol_ref,
        "binary": binary,
        "production_source_equal": True,
    }


def _expected_profile_command(directory: str, stem: str, width: int,
                              binary_name: str) -> list[str]:
    affinity = ",".join(str(index) for index in range(32))
    return ["taskset", "-c", affinity, "valgrind", "--tool=callgrind",
            "--collect-atstart=no", f"--toggle-collect={OWNER}",
            f"--zero-before={OWNER}", f"--dump-after={OWNER}",
            f"--callgrind-out-file={directory}/{stem}.callgrind",
            f"<binary:{binary_name}>", "--route", "parts", "--shape", "large",
            "--workers", str(width), "--task-floor", "0", "--state", "primed",
            "--samples", "1", "--warmup", "0", "--output",
            f"{directory}/{stem}.json"]


def _expected_trace_command(directory: str, stem: str, width: int, state: str,
                            binary_name: str) -> list[str]:
    affinity = ",".join(str(index) for index in range(32))
    return ["taskset", "-c", affinity, "strace", "-f", "-c", "-e",
            "trace=clone,clone3,futex", "-o", f"{directory}/{stem}.strace",
            f"<binary:{binary_name}>", "--route", "parts", "--shape", "large",
            "--state", state, "--workers", str(width), "--task-floor", "0",
            "--samples", "1", "--warmup", "0", "--output",
            f"{directory}/{stem}.json"]


def _artifact_map(row: dict[str, Any], expected: set[str], label: str) -> dict[str, dict[str, Any]]:
    values = row.get("artifacts")
    require(isinstance(values, dict) and set(values) == expected,
            f"{label} artifact set changed")
    result: dict[str, dict[str, Any]] = {}
    for name in sorted(expected):
        require(Path(name).name == name, f"{label} artifact key escapes packet")
        result[name] = artifact(values[name], f"{label} {name}")
        require(result[name]["path"] == f"{label}/{name}",
                f"{label} {name} path changed")
    return result


def _report_config(report: dict[str, Any], width: int, state: str, label: str) -> None:
    require(report.get("schema") == REPORT_SCHEMA, f"{label} schema changed")
    config = report.get("config")
    require(isinstance(config, dict), f"{label} config is missing")
    expected = {"route": "parts", "shape": "large", "workers": width,
                "task_floor": 0, "aggregate_parallel_bytes": 65536,
                "state": state, "samples": 1, "warmup": 0,
                "cpu_task_limit": 1_000_000}
    for key, value in expected.items():
        require(config.get(key) == value, f"{label} config {key} changed")


def _verification(report: dict[str, Any], label: str) -> dict[str, Any]:
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1,
            f"{label} sample count changed")
    sample = samples[0]
    require(isinstance(sample, dict) and sample.get("sample") == 0,
            f"{label} sample is malformed")
    verification = sample.get("verification")
    require(isinstance(verification, dict)
            and verification.get("ordered") is True
            and verification.get("all_member_sha256_match") is True
            and type(verification.get("members")) is int
            and type(verification.get("logical_bytes")) is int
            and is_sha(verification.get("sequence_sha256")),
            f"{label} output verification failed")
    require(verification["members"] == 32 and verification["logical_bytes"] == 8 * 1024 * 1024,
            f"{label} output dimensions changed")
    resources = sample.get("resources")
    require(isinstance(resources, dict), f"{label} resource snapshot is missing")
    for marker in ("before_operation", "after_operation", "after_drop"):
        snapshot = resources.get(marker)
        require(isinstance(snapshot, dict), f"{label} resource {marker} is missing")
        for key in ("workers", "io_concurrency", "cpu_tasks"):
            require(type(snapshot.get(key)) is int and snapshot[key] >= 0,
                    f"{label} resource {marker}.{key} is invalid")
    require(resources.get("worker_and_io_released") is True
            and resources.get("cpu_tasks_within_limit") is True,
            f"{label} resource release/budget witness failed")
    return sample


def _output_fingerprint(report: dict[str, Any]) -> str:
    """Fingerprint deterministic output/corpus identities, excluding timings."""

    selected: list[tuple[tuple[str, ...], Any]] = []
    for path, value in _walk(report):
        if not path:
            continue
        if path[0] == "samples" and (len(path) < 2 or path[1] != "0"):
            continue
        key = path[-1].lower().replace("_", "")
        if isinstance(value, (str, int, list, dict)) and (
                "output" in key or "corpus" in key or "member" in key
                or key.endswith("sha256") or key.endswith("digest")):
            selected.append((path, value))
    require(selected, "report has no deterministic output identity")
    return json_digest(selected)


def _validate_report(path: Path, width: int, state: str, mode: str,
                     qualification: Path | None = None) -> dict[str, Any]:
    label = rel(path)
    report = read_json(path)
    require(isinstance(report, dict), f"{label} is malformed")
    _report_config(report, width, state, label)
    corpus = report.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("selected_payload_member_count") == 32,
            f"{label} corpus is incomplete")
    sample = _verification(report, label)
    metrics = report.get("metrics")
    require(isinstance(metrics, dict), f"{label} metrics are missing")
    if mode == "profile":
        require(metrics.get("source_metrics_feature") is False
                and "unavailable" in str(metrics.get("cpu_clock", "")).lower(),
                f"{label} profile metric availability changed")
        require(sample.get("cpu_ns") is None, f"{label} profile carries CPU time")
        source = sample.get("source_metrics")
        require(isinstance(source, dict)
                and source.get("availability") == "unavailable-profile-build"
                and source.get("logical_calls") is None,
                f"{label} profile source metrics changed")
    else:
        require(metrics.get("source_metrics_feature") is True
                and "ProcessCPUTime" in str(metrics.get("cpu_clock", "")),
                f"{label} observer metric availability changed")
        require(type(sample.get("cpu_ns")) is int and sample["cpu_ns"] >= 0,
                f"{label} observer CPU witness is invalid")
        source = sample.get("source_metrics")
        require(isinstance(source, dict)
                and source.get("availability") == "source-metrics-feature",
                f"{label} observer source metrics changed")
        expected_calls = 64 if state == "fresh" else 0
        require(source.get("logical_calls") == expected_calls,
                f"{label} paired source call counter changed")
        if state == "fresh":
            require(source.get("requested_bytes") == 146041
                    and source.get("returned_bytes") == 146041
                    and source.get("short_reads") == 0
                    and source.get("active_reads_after_operation") == 0
                    and source.get("max_simultaneous_reads") == 1,
                    f"{label} fresh source counters changed")
        else:
            require(source.get("requested_bytes") == 0
                    and source.get("returned_bytes") == 0
                    and source.get("short_reads") == 0
                    and source.get("active_reads_after_operation") == 0
                    and source.get("max_simultaneous_reads") == 0,
                    f"{label} primed source counters changed")
    require(sample.get("primed_cache_hit_control") is (state == "primed"),
            f"{label} cache state witness changed")
    fingerprint = _output_fingerprint(report)
    qualification_identity: dict[str, Any] | None = None
    if qualification is not None:
        oracle = read_json(qualification)
        require(isinstance(oracle, dict), f"{rel(qualification)} is malformed")
        oracle_fingerprint = _output_fingerprint(oracle)
        require(oracle_fingerprint == fingerprint,
                f"{label} output identity differs from qualification-before")
        require(oracle.get("corpus") == corpus,
                f"{label} corpus identity differs from qualification-before")
        qualification_identity = {"file": file_identity(qualification),
                                  "output_fingerprint": oracle_fingerprint}
    return {
        "file": file_identity(path),
        "output_fingerprint": fingerprint,
        "qualification_before": qualification_identity,
        "verification": {
            "sequence_sha256": sample["verification"]["sequence_sha256"],
            "members": sample["verification"]["members"],
            "logical_bytes": sample["verification"]["logical_bytes"],
        },
        "source_metrics": {
            "logical_calls": sample["source_metrics"].get("logical_calls"),
            "requested_bytes": sample["source_metrics"].get("requested_bytes"),
            "returned_bytes": sample["source_metrics"].get("returned_bytes"),
            "short_reads": sample["source_metrics"].get("short_reads"),
            "max_simultaneous_reads": sample["source_metrics"].get("max_simultaneous_reads"),
        },
    }


def _category(name: str) -> str:
    lowered = name.lower()
    if any(word in lowered for word in ("thread", "pthread", "clone", "spawn", "join",
                                        "tls", "scheduler", "park", "waker")):
        return "threading"
    if any(word in lowered for word in ("malloc", "calloc", "realloc", "free", "alloc",
                                        "dealloc", "rawvec", "memalign", "box")):
        return "allocation"
    if any(word in lowered for word in ("channel", "sender", "receiver", "send", "recv",
                                        "queue", "deque", "mutex", "condvar")):
        return "channel"
    if any(word in lowered for word in ("cache", "hash", "hasher", "hashmap", "memo")):
        return "cache/hash"
    return "other"


def _diagnostics(parsed: dict[str, Any], parser: Any) -> dict[str, Any]:
    rows = [{"name": function["name"] or "<unnamed>", "self_ir": function["self_ir"]}
            for function in parsed["functions"].values() if function["self_ir"] > 0]
    rows.sort(key=lambda row: (-row["self_ir"], row["name"]))
    categories: dict[str, int] = {name: 0 for name in
                                  ("threading", "allocation", "channel", "cache/hash", "other")}
    tops: dict[str, list[dict[str, Any]]] = {name: [] for name in categories}
    for row in rows:
        key = _category(row["name"])
        categories[key] += row["self_ir"]
        if len(tops[key]) < 5:
            tops[key].append(row)
    return {"all_self_ir": sum(row["self_ir"] for row in rows),
            "top_self_functions": rows[:12],
            "categories": {key: {"self_ir": categories[key], "top_functions": tops[key]}
                           for key in categories}}


def _profile_row(parsed: dict[str, Any], parser: Any, raw: dict[str, Any],
                 report: dict[str, Any], repeat: int, width: int) -> dict[str, Any]:
    names = parsed["names"]
    owner_ids = [function_id for function_id, name in names.items() if name == OWNER]
    require(len(owner_ids) == 1, f"{raw['label']} owner name is not unique")
    owner_id = owner_ids[0]
    owner_function = parsed["functions"].get(owner_id)
    require(isinstance(owner_function, dict) and owner_function.get("records") == 1,
            f"{raw['label']} owner function record count changed")
    incoming = parser.incoming_edges(parsed, owner_id)
    require(len(incoming) == 1 and incoming[0]["calls"] == 1
            and incoming[0]["caller"] == "cached_part_profile::run_sample",
            f"{raw['label']} owner call edge is not exactly one")
    owner_inclusive = incoming[0]["inclusive_ir"]
    view = parser.function_view(owner_function, names)
    require(view["name"] == OWNER and view["self_ir"] > 0,
            f"{raw['label']} owner self Ir is missing")
    require(view["direct_children"] and any(item["callee"] == CHILD
                                              for item in view["direct_children"]),
            f"{raw['label']} operation child is missing")
    require(view["self_ir"] + view["direct_children_ir"] == owner_inclusive,
            f"{raw['label']} owner immediate-child partition changed")
    require(owner_inclusive == parsed["header"]["summary_ir"],
            f"{raw['label']} owner edge differs from region summary")
    all_self = sum(function["self_ir"] for function in parsed["functions"].values())
    require(all_self == parsed["header"]["summary_ir"],
            f"{raw['label']} function self-Ir partition differs from summary")
    path = parser.dominant_path(parsed, owner_id, 8)
    reachable: set[int] = set()
    pending = [owner_id]
    while pending:
        current = pending.pop()
        if current in reachable or current not in parsed["functions"]:
            continue
        reachable.add(current)
        pending.extend(edge["callee_id"] for edge in parsed["functions"][current]["edges"]
                       if edge["calls"] > 0 and edge["inclusive_ir"] > 0)
    disconnected = [function for function_id, function in parsed["functions"].items()
                    if function_id not in reachable and function["self_ir"] > 0]
    disconnected.sort(key=lambda function: (-function["self_ir"], function["name"]))
    diagnostic = _diagnostics(parsed, parser)
    diagnostic["disconnected_positive_self_functions"] = len(disconnected)
    diagnostic["disconnected_top_self_functions"] = [
        {"name": function["name"] or "<unnamed>", "self_ir": function["self_ir"]}
        for function in disconnected[:8]
    ]
    return {
        "repeat": repeat,
        "workers": width,
        "raw": raw["numbered"],
        "report": report,
        "events": parsed["header"]["events"],
        "summary_ir": parsed["header"]["summary_ir"],
        "totals_ir": parsed["header"]["totals_ir"],
        "owner": {
            "name": OWNER, "calls": incoming[0]["calls"],
            "caller": incoming[0]["caller"], "inclusive_ir": owner_inclusive,
            "self_ir": view["self_ir"],
            "direct_children_ir": view["direct_children_ir"],
            "direct_children": view["direct_children"],
            "partition_exact": True,
        },
        "ancestry": parser.ancestry_to_owner(parsed, owner_id),
        "dominant_path": path,
        "statistics": parsed["statistics"],
        "diagnostic": diagnostic,
        "thread_coverage": {
            "owner_graph_only": True,
            "owner_inclusive_all_threads": False,
            "physical_worker_fraction": False,
            "note": "Callgrind worker roots may be disconnected from the owner ancestry; this is a guest-Ir cost diagnostic.",
        },
    }


def _validate_profile_final(parser: Any, path: Path, raw: dict[str, Any]) -> dict[str, Any]:
    parsed = parser.parse_raw(path, allow_empty=True)
    header = parsed["header"]
    require(header["events"] == ["Ir"] and header["part"] == 2
            and header["trigger"] == "Program termination"
            and header["summary_ir"] == 0 and header["totals_ir"] == 0
            and not parsed["functions"], f"{raw['label']} termination dump is not zero")
    return {"file": file_identity(path), "events": header["events"],
            "part": header["part"], "trigger": header["trigger"],
            "summary_ir": header["summary_ir"], "totals_ir": header["totals_ir"]}


def validate_failed_profile(parser: Any, build: dict[str, Any], plan_sha: str) -> dict[str, Any]:
    directory = HERE / "profiles"
    receipt_path = directory / "receipts.json"
    rows = read_json(receipt_path)
    require(isinstance(rows, list) and len(rows) == 1, "failed profile receipt cardinality changed")
    row = rows[0]
    require(isinstance(row, dict) and row.get("repeat") == 0 and row.get("workers") == 1
            and row.get("exit_code") == -11 and row.get("plan_sha256") == plan_sha,
            "failed profile receipt changed")
    require(row.get("binary", {}).get("sha256") == build["binary"]["sha256"]
            and row.get("binary", {}).get("bytes") == build["binary"]["bytes"],
            "failed profile binary identity changed")
    require(row.get("driver_sha256") == sha256(HERE / "profile.py"),
            "failed profile driver hash changed")
    command = canonical_binary_command(row.get("command"), build["binary"]["name"])
    expected = _expected_profile_command("profiles", "0-1", 1, build["binary"]["name"])
    require(command == expected, "failed profile command changed")
    artifacts = _artifact_map(row, {"0-1.callgrind", "0-1.log"}, "profiles")
    raw_path = HERE / artifacts["0-1.callgrind"]["path"]
    parsed = parser.parse_raw(raw_path, allow_empty=True)
    header = parsed["header"]
    require(header["events"] == ["Ir"] and header["part"] == 1
            and header["trigger"] == "Program termination"
            and header["summary_ir"] == 0 and not parsed["functions"],
            "failed profile zero dump changed")
    header_command = canonical_binary_command(shlex.split(header["command"] or ""),
                                               build["binary"]["name"])
    require(header_command == expected[10:], "failed profile raw command changed")
    log = read_text(HERE / artifacts["0-1.log"]["path"])
    require("SIGSEGV" in log and "rustix::backend::param::auxv" in log,
            "failed profile SIGSEGV witness changed")
    fallback = read_json(HERE / "profile-clock-fallback.json")
    require(fallback == {
        "reason": "Callgrind SIGSEGV before region in rustix linux_raw auxv check_elf_base clock initialization;zero collectedIr",
        "failed_lane": "profiles", "failed_build": "profile-build",
        "original_probe_archive": "profile-src-before-clock-fallback",
        "compatibility_build": "profile-build-compat",
        "compatibility_lane": "profiles-compat",
        "change": "profile-only cpu_now_ns returnsNone; native tool remains unchanged",
        "native_measurements_affected": False,
    }, "profile-clock-fallback contract changed")
    return {
        "receipt": file_identity(receipt_path),
        "row": {"repeat": 0, "workers": 1, "exit_code": -11,
                "binary": build["binary"]},
        "artifacts": artifacts,
        "raw": {"events": header["events"], "part": header["part"],
                "trigger": header["trigger"], "summary_ir": header["summary_ir"],
                "totals_ir": header["totals_ir"]},
        "log_contains_sigsegv_before_owner": True,
        "qualified": False,
        "reason": fallback["reason"],
        "native_measurements_affected": False,
        "fallback": file_identity(HERE / "profile-clock-fallback.json"),
    }


def validate_profile_lane(parser: Any, build: dict[str, Any], profile_plan: dict[str, Any],
                          profile_plan_sha: str) -> dict[str, Any]:
    directory = HERE / "profiles-compat"
    receipt_path = directory / "receipts.json"
    rows = read_json(receipt_path)
    require(isinstance(rows, list) and len(rows) == 6,
            "compatibility profile receipt cardinality changed")
    complete = read_json(directory / "complete.json")
    require(isinstance(complete, dict) and complete.get("processes") == 6,
            "compatibility profile completion count changed")
    complete_receipts = artifact(complete.get("receipts"), "profiles-compat completion receipts")
    require(complete_receipts["path"] == "profiles-compat/receipts.json",
            "profiles-compat completion receipt path changed")
    binary = build["binary"]
    entries: list[dict[str, Any]] = []
    order = [(0, 1), (0, 8), (0, 32), (1, 32), (1, 8), (1, 1)]
    for row, (repeat, width) in zip(rows, order):
        stem = f"{repeat}-{width}"
        label = f"profiles-compat/{stem}"
        require(isinstance(row, dict) and row.get("repeat") == repeat
                and row.get("workers") == width and row.get("exit_code") == 0
                and row.get("plan_sha256") == profile_plan_sha,
                f"{label} receipt changed")
        require(row.get("binary", {}).get("sha256") == binary["sha256"]
                and row.get("binary", {}).get("bytes") == binary["bytes"],
                f"{label} binary identity changed")
        require(row.get("driver_sha256") == sha256(HERE / "profile_compat.py"),
                f"{label} driver hash changed")
        command = canonical_binary_command(row.get("command"), binary["name"])
        require(command == _expected_profile_command("profiles-compat", stem, width,
                                                     binary["name"]),
                f"{label} Callgrind command changed")
        artifacts = _artifact_map(row, {f"{stem}.callgrind", f"{stem}.callgrind.1",
                                        f"{stem}.json", f"{stem}.log"}, "profiles-compat")
        numbered_path = HERE / artifacts[f"{stem}.callgrind.1"]["path"]
        final_path = HERE / artifacts[f"{stem}.callgrind"]["path"]
        raw = {"label": label, "numbered": artifacts[f"{stem}.callgrind.1"]["path"]}
        numbered = parser.parse_raw(numbered_path)
        header = numbered["header"]
        require(header["events"] == ["Ir"] and header["part"] == 1
                and header["trigger"] == f"--dump-after={OWNER}"
                and type(header["summary_ir"]) is int and header["summary_ir"] > 0
                and header["totals_ir"] == header["summary_ir"],
                f"{label} numbered Callgrind dump contract changed")
        header_command = canonical_binary_command(shlex.split(header["command"] or ""),
                                                   binary["name"])
        expected_command = _expected_profile_command("profiles-compat", stem, width,
                                                     binary["name"])
        require(header_command == expected_command[10:],
                f"{label} raw command changed")
        final = _validate_profile_final(parser, final_path, raw)
        numbered_parts = sorted(directory.glob(f"{stem}.callgrind.*"))
        require([path.name for path in numbered_parts] == [f"{stem}.callgrind.1"],
                f"{label} has more than one numbered Callgrind dump")
        report_path = HERE / artifacts[f"{stem}.json"]["path"]
        qualification = HERE / "qualification-before" / f"0-large-primed-0-{width}-before.json"
        report = _validate_report(report_path, width, "primed", "profile", qualification)
        parsed_row = _profile_row(numbered, parser, raw, report, repeat, width)
        parsed_row["final"] = final
        parsed_row["numbered"] = {"file": file_identity(numbered_path),
                                   "part": header["part"], "trigger": header["trigger"]}
        entries.append(parsed_row)
    summary_by_width: dict[str, Any] = {}
    for width in WIDTHS:
        selected = [entry for entry in entries if entry["workers"] == width]
        values = [entry["summary_ir"] for entry in selected]
        categories: dict[str, Any] = {}
        for category in ("threading", "allocation", "channel", "cache/hash"):
            values_by_repeat = [entry["diagnostic"]["categories"][category]["self_ir"]
                                for entry in selected]
            categories[category] = {"repeat_values": values_by_repeat,
                                    "mean_self_ir": statistics.fmean(values_by_repeat)}
        top_names: dict[str, list[int]] = {}
        for entry in selected:
            for row in entry["diagnostic"]["top_self_functions"]:
                top_names.setdefault(row["name"], []).append(row["self_ir"])
        top_functions = [{"name": name, "mean_self_ir": statistics.fmean(ir_values),
                          "repeat_values": ir_values}
                         for name, ir_values in top_names.items()]
        top_functions.sort(key=lambda row: (-row["mean_self_ir"], row["name"]))
        summary_by_width[str(width)] = {
            "repeat_values": values, "mean_summary_ir": statistics.fmean(values),
            "owner_self_ir": [entry["owner"]["self_ir"] for entry in selected],
            "owner_direct_children_ir": [entry["owner"]["direct_children_ir"] for entry in selected],
            "diagnostic_categories": categories,
            "top_self_functions": top_functions[:12],
        }
    return {
        "directory": "profiles-compat", "receipts": file_identity(receipt_path),
        "complete": file_identity(directory / "complete.json"),
        "binary": binary, "entries": entries, "summary_by_width": summary_by_width,
        "qualified": True,
        "scope": "operation-only Ir inside one exact Callgrind owner region; worker roots may be disconnected",
    }


def _parse_strace(path: Path) -> dict[str, dict[str, int]]:
    text = read_text(path)
    lines = text.splitlines()
    require(lines and lines[0].startswith("% time")
            and any(line.strip().endswith(" total") for line in lines),
            f"{rel(path)} strace summary is malformed")
    result: dict[str, dict[str, int]] = {}
    total: tuple[int, int] | None = None
    for line in lines:
        fields = line.split()
        if len(fields) < 5:
            continue
        if fields[-1] == "total":
            try:
                total = (int(fields[3]), int(fields[4]) if fields[4].isdigit() else 0)
            except ValueError:
                fail(f"{rel(path)} total syscall counts are malformed")
            continue
        syscall = fields[-1]
        if syscall not in {"clone", "clone3", "futex"}:
            continue
        require(len(fields) >= 5, f"{rel(path)} syscall row is malformed")
        try:
            calls = int(fields[3])
            errors = int(fields[4]) if fields[4].isdigit() else 0
        except ValueError:
            fail(f"{rel(path)} syscall counts are malformed")
        require(calls >= 0 and errors >= 0 and errors <= calls,
                f"{rel(path)} syscall counts are invalid")
        require(syscall not in result, f"{rel(path)} duplicate {syscall} row")
        result[syscall] = {"attempts": calls, "errors": errors,
                           "successes": calls - errors}
    for syscall in ("clone", "clone3", "futex"):
        result.setdefault(syscall, {"attempts": 0, "errors": 0, "successes": 0})
    require(total is not None
            and total[0] == sum(row["attempts"] for row in result.values())
            and total[1] == sum(row["errors"] for row in result.values()),
            f"{rel(path)} total syscall counts do not match parsed rows")
    return result


def _trace_expected_clone3(width: int, state: str, after: bool) -> int:
    if width == 1:
        return 0
    if after:
        return width
    return width if state == "fresh" else 2 * width


def validate_trace_lane(directory_name: str, trace_plan: dict[str, Any], trace_plan_sha: str,
                        baseline_binary: dict[str, Any] | None, after: bool = False) -> dict[str, Any]:
    directory = HERE / directory_name
    if not directory.is_dir():
        return {"available": False, "directory": directory_name,
                "note": "optional after observer trace lane has not been retained"}
    receipt_path = directory / "receipts.json"
    rows = read_json(receipt_path)
    require(isinstance(rows, list) and len(rows) == 12,
            f"{directory_name} receipt cardinality changed")
    complete = read_json(directory / "complete.json")
    require(isinstance(complete, dict) and complete.get("processes") == 12,
            f"{directory_name} completion count changed")
    complete_receipts = artifact(complete.get("receipts"), f"{directory_name} completion receipts")
    require(complete_receipts["path"] == f"{directory_name}/receipts.json",
            f"{directory_name} completion receipt path changed")
    order = [(repeat, width, state)
             for repeat in range(2)
             for width in (WIDTHS if repeat == 0 else tuple(reversed(WIDTHS)))
             for state in TRACE_STATES]
    entries: list[dict[str, Any]] = []
    binary: dict[str, Any] | None = None
    for row, (repeat, width, state) in zip(rows, order):
        stem = f"{repeat}-{width}-{state}"
        label = f"{directory_name}/{stem}"
        require(isinstance(row, dict) and row.get("repeat") == repeat
                and row.get("workers") == width and row.get("state") == state
                and row.get("exit_code") == 0 and row.get("plan_sha256") == trace_plan_sha,
                f"{label} receipt changed")
        current_binary = external_identity(row.get("binary"), f"{label}")
        if binary is None:
            binary = current_binary
        else:
            require(current_binary["name"] == binary["name"]
                    and current_binary["bytes"] == binary["bytes"]
                    and current_binary["sha256"] == binary["sha256"],
                    f"{label} binary identity differs within trace lane")
        if baseline_binary is not None and not after:
            require(current_binary["name"] == baseline_binary["name"]
                    and current_binary["bytes"] == baseline_binary["bytes"]
                    and current_binary["sha256"] == baseline_binary["sha256"],
                    f"{label} observer binary differs from build-before")
        command = canonical_binary_command(row.get("command"), current_binary["name"])
        require(command == _expected_trace_command(directory_name, stem, width, state,
                                                   current_binary["name"]),
                f"{label} strace command changed")
        artifacts = _artifact_map(row, {f"{stem}.json", f"{stem}.log", f"{stem}.strace"},
                                  directory_name)
        report_path = HERE / artifacts[f"{stem}.json"]["path"]
        qualification = HERE / "qualification-before" / f"0-large-{state}-0-{width}-before.json"
        report = _validate_report(report_path, width, state, "observer", qualification)
        strace_path = HERE / artifacts[f"{stem}.strace"]["path"]
        syscall = _parse_strace(strace_path)
        expected_clone3 = _trace_expected_clone3(width, state, after)
        require(syscall["clone3"]["successes"] == expected_clone3,
                f"{label} clone3 success count differs from expected {'after' if after else 'before'} lane")
        require(syscall["clone"]["successes"] == 0,
                f"{label} clone success count is unexpectedly nonzero")
        source = report["source_metrics"]
        entries.append({
            "repeat": repeat, "workers": width, "state": state,
            "report": report, "strace": {"file": file_identity(strace_path),
                                          "syscalls": syscall},
            "paired_source_counters": source,
            "expected_clone3_successes": expected_clone3,
            "trace_wall_time_pooled": False,
        })
    assert binary is not None
    pairs: list[dict[str, Any]] = []
    for repeat in range(2):
        for width in WIDTHS:
            fresh = next(entry for entry in entries
                         if entry["repeat"] == repeat and entry["workers"] == width
                         and entry["state"] == "fresh")
            primed = next(entry for entry in entries
                          if entry["repeat"] == repeat and entry["workers"] == width
                          and entry["state"] == "primed")
            fresh_source = fresh["paired_source_counters"]
            primed_source = primed["paired_source_counters"]
            require(fresh_source["logical_calls"] == 64
                    and primed_source["logical_calls"] == 0,
                    f"{directory_name} paired source counter changed")
            require(fresh["report"]["verification"]["sequence_sha256"] ==
                    primed["report"]["verification"]["sequence_sha256"],
                    f"{directory_name} fresh/primed output sequence changed")
            pairs.append({
                "repeat": repeat, "workers": width,
                "fresh_one_operation_clone3_successes": fresh["strace"]["syscalls"]["clone3"]["successes"],
                "primed_preload_plus_one_operation_clone3_successes": primed["strace"]["syscalls"]["clone3"]["successes"],
                "expected_fresh": _trace_expected_clone3(width, "fresh", after),
                "expected_primed": _trace_expected_clone3(width, "primed", after),
                "fresh_source_calls": fresh_source["logical_calls"],
                "primed_source_calls": primed_source["logical_calls"],
                "source_counter_pair_exact": True,
                "clone3_pair_exact": True,
            })
    return {
        "available": True, "directory": directory_name,
        "receipts": file_identity(receipt_path), "complete": file_identity(directory / "complete.json"),
        "binary": binary, "entries": entries, "pairs": pairs,
        "trace_wall_time_pooled": False,
        "scope": "whole-child clone/clone3/futex syscall summary; attempts, errors, and successes retained",
        "expected_clone3_model": "W fresh and 2W primed before; W fresh and W primed after; W=1 is zero",
    }


def _aggregate_profile_widths(entries: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for width in WIDTHS:
        selected = [entry for entry in entries if entry["workers"] == width]
        values = [entry["summary_ir"] for entry in selected]
        result[str(width)] = {
            "repeat_values": values,
            "mean_summary_ir": statistics.fmean(values),
            "owner_self_ir": [entry["owner"]["self_ir"] for entry in selected],
            "owner_direct_children_ir": [entry["owner"]["direct_children_ir"] for entry in selected],
        }
    return result


def build_analysis() -> dict[str, Any]:
    parser, parser_identity = load_parser()
    origin = load_origin()
    profile_plan, trace_plan, plans = load_plans()
    require(profile_plan.get("owner_suffix") in OWNER,
            "profile owner suffix does not bind exact owner")
    baseline, production = _production_identity()
    profile_sha = plans["profile"]["sha256"]
    trace_sha = plans["trace"]["sha256"]
    failed_build = validate_profile_build("profile-build", production, profile_sha,
                                          "profile-src-before-clock-fallback")
    compatibility_build = validate_profile_build("profile-build-compat", production, profile_sha,
                                                 "profile-src")
    failed_profile = validate_failed_profile(parser, failed_build, profile_sha)
    profile = validate_profile_lane(parser, compatibility_build, profile_plan, profile_sha)
    trace_before = validate_trace_lane("traces", trace_plan, trace_sha,
                                       baseline["binaries"]["observer"], after=False)
    trace_after = validate_trace_lane("traces-after", trace_plan, trace_sha,
                                      None, after=True)
    return {
        "schema": ANALYSIS_SCHEMA,
        "packet": "change-0787",
        "scope": {
            "profile": "operation-only Callgrind Ir for cached_part_region_0787; exact owner graph and immediate-child partition",
            "trace": "whole-child clone/clone3/futex syscall counts with source-counter pairs",
            "claims_excluded": [
                "native wall-time or latency claims",
                "physical worker fractions or serial-fraction claims",
                "owner-inclusive-all-threads claim when worker roots are disconnected",
                "trace wall time pooled with native or profile measurements",
            ],
        },
        "origin": origin,
        "parser": parser_identity,
        "plans": plans,
        "source": baseline,
        "profile_builds": {
            "failed_baseline": failed_build,
            "compatibility": compatibility_build,
        },
        "failed_attempt": failed_profile,
        "profile": {
            "lane": "profiles-compat",
            "qualified": profile["qualified"],
            "scope": profile["scope"],
            "receipts": profile["receipts"],
            "complete": profile["complete"],
            "binary": profile["binary"],
            "entries": profile["entries"],
            "summary_by_width": _aggregate_profile_widths(profile["entries"]),
            "diagnostic_summary_by_width": profile["summary_by_width"],
        },
        "traces": {"before": trace_before, "after": trace_after},
        "verification": {
            "profile_exact_owner_call": True,
            "profile_one_numbered_dump_per_case": True,
            "profile_termination_dump_zero": True,
            "profile_events_ir": True,
            "profile_summary_equals_all_self_ir": True,
            "profile_owner_immediate_partition_exact": True,
            "profile_output_identity_matches_qualification_before": True,
            "trace_source_counter_pairs_exact": True,
            "trace_clone3_counts_exact": True,
            "trace_wall_time_pooled": False,
        },
    }


def _fmt(value: Any) -> str:
    if value is None:
        return "n/a"
    if isinstance(value, float):
        return f"{value:.1f}"
    return str(value)


def render_markdown(analysis: dict[str, Any]) -> str:
    lines = [
        "# 0787 cached Part profile analysis",
        "",
        "This report replays retained receipts, reports, raw Callgrind files, and strace summaries offline.",
        "Callgrind values are guest instruction counts for the exact operation owner. The trace lane is a",
        "separate whole-child syscall census. No native wall-time claim, physical worker fraction, or",
        "trace-time pooling is made.",
        "",
        f"- Reused parser: `{analysis['parser']['path']}` (Git blob `{analysis['parser']['git_blob_sha1']}`)",
        f"- Production source census: {analysis['source']['production_file_count']} files at revision `{analysis['source']['production_revision']}`",
        f"- Profile lane: `{analysis['profile']['lane']}`, six qualified regions",
        f"- Failed baseline attempt retained: SIGSEGV before owner, zero Ir, qualified `{analysis['failed_attempt']['qualified']}`",
        "",
        "## Callgrind diagnostic by width",
        "",
        "| Width | Repeat Ir | Mean Ir | Owner self Ir | Immediate child Ir | Threading self Ir | Allocation self Ir | Channel self Ir | Cache/hash self Ir |",
        "|---:|---|---:|---|---|---:|---:|---:|---:|",
    ]
    summary = analysis["profile"]["diagnostic_summary_by_width"]
    for width in WIDTHS:
        row = summary[str(width)]
        categories = row["diagnostic_categories"]
        lines.append("| " + " | ".join([
            str(width), ", ".join(map(str, row["repeat_values"])), _fmt(row["mean_summary_ir"]),
            ", ".join(map(str, row["owner_self_ir"])),
            ", ".join(map(str, row["owner_direct_children_ir"])),
            _fmt(categories["threading"]["mean_self_ir"]),
            _fmt(categories["allocation"]["mean_self_ir"]),
            _fmt(categories["channel"]["mean_self_ir"]),
            _fmt(categories["cache/hash"]["mean_self_ir"]),
        ]) + " |")
    lines.extend([
        "",
        "The owner self row plus its immediate child inclusive rows equals the owner edge and the numbered region summary.",
        "Descendant inclusive rows are retained in each dominant path and are not added to that partition.",
        "Worker roots may be disconnected in a threaded Callgrind graph, so owner inclusive is not treated as an all-thread total.",
        "",
        "## Top self-Ir functions",
        "",
        "| Width | Repeat | Top self-Ir functions (name: Ir) |",
        "|---:|---:|---|",
    ])
    for entry in analysis["profile"]["entries"]:
        top = "; ".join(f"{row['name']}: {row['self_ir']}"
                         for row in entry["diagnostic"]["top_self_functions"][:5])
        lines.append(f"| {entry['workers']} | {entry['repeat']} | {top} |")
    lines.extend([
        "",
        "## Dominant paths",
        "",
        "| Width | Repeat | Path (owner to dominant child) |",
        "|---:|---:|---|",
    ])
    for entry in analysis["profile"]["entries"]:
        names = []
        for node in entry["dominant_path"]:
            names.append(str(node["name"]))
        lines.append(f"| {entry['workers']} | {entry['repeat']} | " + " → ".join(names) + " |")
    lines.extend(["", "## Thread syscall counts", ""])
    lines.append("Trace wall time is deliberately absent from this table and is never pooled with profile or native results.")
    lines.extend(["", "| Lane | Repeat | Width | State | clone attempts/errors/successes | clone3 attempts/errors/successes | futex attempts/errors/successes |", "|---|---:|---:|---|---|---|---|"])
    for lane_name in ("before", "after"):
        lane = analysis["traces"][lane_name]
        if not lane.get("available"):
            lines.append(f"| {lane_name} | — | — | — | pending | pending | pending |")
            continue
        for entry in lane["entries"]:
            syscalls = entry["strace"]["syscalls"]
            def counts(name: str) -> str:
                row = syscalls[name]
                return f"{row['attempts']}/{row['errors']}/{row['successes']}"
            lines.append(f"| {lane_name} | {entry['repeat']} | {entry['workers']} | {entry['state']} | {counts('clone')} | {counts('clone3')} | {counts('futex')} |")
    lines.extend([
        "",
        "For the before lane, fresh one-operation clone3 successes are W (W>1) and primed preload-plus-operation successes are 2W; W=1 is zero. The optional after lane expects W in both states because only the preload creates the extra worker set.",
        "Each fresh/primed pair retains source calls 64/0 and matching output sequence identity.",
        "",
        "## Custody and limits",
        "",
        "All packet artifacts are checked by recorded size and SHA-256. Profile executables are accepted only while live or when an exact verified cleanup witness binds basename, size, and hash. The initial baseline profile attempt remains disclosed and unqualified; the compatibility profile is the qualified Callgrind lane.",
        "",
    ])
    return "\n".join(lines)


def write_outputs(analysis: dict[str, Any]) -> None:
    PROFILE_ANALYSIS_JSON.write_text(
        json.dumps(analysis, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    PROFILE_ANALYSIS_MD.write_text(render_markdown(analysis), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    argument_parser = argparse.ArgumentParser(description=__doc__)
    mode = argument_parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write profile-analysis.json and .md")
    mode.add_argument("--check", action="store_true", help="replay and compare generated outputs")
    args = argument_parser.parse_args(argv)
    try:
        analysis = build_analysis()
        if args.check:
            expected_json = json.dumps(analysis, indent=2, sort_keys=True, ensure_ascii=False) + "\n"
            require(PROFILE_ANALYSIS_JSON.is_file(), "profile-analysis.json is missing")
            require(PROFILE_ANALYSIS_JSON.read_text(encoding="utf-8") == expected_json,
                    "profile-analysis.json differs; rerun with --write")
            expected_md = render_markdown(analysis)
            require(PROFILE_ANALYSIS_MD.is_file(), "profile-analysis.md is missing")
            require(PROFILE_ANALYSIS_MD.read_text(encoding="utf-8") == expected_md,
                    "profile-analysis.md differs; rerun with --write")
            print("0787 profile analysis check passed")
        else:
            write_outputs(analysis)
            print(f"wrote {PROFILE_ANALYSIS_JSON} and {PROFILE_ANALYSIS_MD}")
        return 0
    except EvidenceError as error:
        print(f"profile analysis failed: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
