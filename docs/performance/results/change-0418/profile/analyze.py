#!/usr/bin/env python3
"""Create deterministic, symbolized views of the retained 0418 perf data.

This is a postprocessor for ``litchi-goal-0418-profile.py``.  It never starts
the workload or Cargo.  ``--write`` runs only perf/readelf over already
retained artifacts and writes the no-inline views plus their provenance;
``--replay`` runs the same postprocessing in a temporary directory and checks
that every byte and identity binding is reproducible.  A ``.zst`` (or ``.gz``)
sidecar may stand in for an intentionally compressed raw artifact.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import json
import math
import os
import subprocess
import sys
import tempfile
from pathlib import Path, PureWindowsPath
from typing import Any


PROFILE_DIR = Path(__file__).resolve().parent
DEFAULT_ROOT = PROFILE_DIR.parent
EXPECTED_CHANGE = 418
ROLES = ("control", "candidate")
MODES = ("normal", "allocator")


class AnalysisError(ValueError):
    """A retained input or postprocessing result failed the 0418 contract."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            fail(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def _constant(value: str) -> None:
    fail(f"non-finite JSON number {value!r}")


def _finite(value: Any, path: str = "JSON") -> None:
    if value is None or isinstance(value, (str, bool)):
        return
    if isinstance(value, (int, float)):
        if not math.isfinite(float(value)):
            fail(f"{path} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _finite(item, f"{path}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            _finite(item, f"{path}.{key}")
        return
    fail(f"{path} contains unsupported JSON type {type(value).__name__}")


def load_json(path: Path) -> Any:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=_constant,
        )
    except AnalysisError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot read {path}: {error}")
    _finite(value, str(path))
    return value


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(
            value, sort_keys=True, separators=(",", ":"),
            ensure_ascii=False, allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError) as error:
        fail(f"cannot canonicalize JSON: {error}")
    raise AssertionError("unreachable")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def relpath(path: Path, root: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return str(path.resolve())


def _relative(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a non-empty path")
    if Path(value).is_absolute() or PureWindowsPath(value).is_absolute():
        fail(f"{label} must be relative")
    if "\\" in value or ".." in Path(value).parts:
        fail(f"{label} must use normalized relative POSIX paths")
    if Path(value).as_posix() != value:
        fail(f"{label} is not normalized")
    return value


def _inside(root: Path, relative: str, label: str) -> Path:
    relative = _relative(relative, label)
    path = (root / relative).resolve()
    try:
        path.relative_to(root.resolve())
    except ValueError:
        fail(f"{label} escapes {root}")
    return path


def _sidecar(path: Path) -> tuple[Path, str] | None:
    if path.is_file():
        return path, "plain"
    for suffix, kind in ((".zst", "zstd"), (".zstd", "zstd"), (".gz", "gzip")):
        candidate = Path(str(path) + suffix)
        if candidate.is_file():
            return candidate, kind
    return None


MAX_DECOMPRESSED_ARTIFACT = 512 * 1024 * 1024


def _bounded_gzip(path: Path, label: str) -> bytes:
    with gzip.open(path, "rb") as stream:
        data = stream.read(MAX_DECOMPRESSED_ARTIFACT + 1)
    if len(data) > MAX_DECOMPRESSED_ARTIFACT:
        fail(f"{label}: decompressed artifact exceeds 512 MiB")
    return data


def _bounded_zstd(path: Path, label: str) -> bytes:
    process = subprocess.Popen(
        ["zstd", "-q", "-d", "-c", str(path)],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        env={**os.environ, "DEBUGINFOD_URLS": ""},
    )
    assert process.stdout is not None
    chunks: list[bytes] = []
    size = 0
    while True:
        block = process.stdout.read(1024 * 1024)
        if not block:
            break
        size += len(block)
        if size > MAX_DECOMPRESSED_ARTIFACT:
            process.kill()
            process.wait()
            fail(f"{label}: decompressed artifact exceeds 512 MiB")
        chunks.append(block)
    stderr = process.stderr.read() if process.stderr is not None else b""
    exit_code = process.wait()
    if exit_code:
        fail(f"{label}: zstd failed: {stderr.decode(errors='replace')}")
    return b"".join(chunks)


def read_artifact(root: Path, relative: str, label: str = "artifact") -> tuple[bytes, dict[str, Any]]:
    """Read a plain artifact or its compressed sidecar and return both hashes."""
    declared = _relative(relative, f"{label}.path")
    path = _inside(root, declared, f"{label}.path")
    found = _sidecar(path)
    if found is None:
        fail(f"{label}: missing {declared} and supported compressed sidecars")
    source, kind = found
    try:
        if kind == "plain":
            if source.stat().st_size > MAX_DECOMPRESSED_ARTIFACT:
                fail(f"{label}: artifact exceeds 512 MiB")
            data = source.read_bytes()
        elif kind == "gzip":
            data = _bounded_gzip(source, label)
        else:
            data = _bounded_zstd(source, label)
    except (OSError, EOFError) as error:
        fail(f"{label}: cannot read {source}: {error}")
    return data, {
        "declared_path": declared,
        "source_path": relpath(source, root),
        "compression": kind,
        "source_bytes": source.stat().st_size,
        "source_sha256": sha256_file(source),
        "bytes": len(data),
        "sha256": sha256_bytes(data),
    }


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(canonical(value) + b"\n")


def _require_mapping(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def _require_string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a non-empty string")
    return value


def _require_integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label} must be an integer >= {minimum}")
    return value


def _hash_descriptor(path: Path, descriptor: dict[str, Any], label: str) -> dict[str, Any]:
    if not path.is_file():
        fail(f"{label}: binary is unavailable before postprocessing: {path}")
    actual_bytes = path.stat().st_size
    actual_hash = sha256_file(path)
    expected_hash = descriptor.get("sha256")
    expected_bytes = descriptor.get("bytes")
    if expected_hash != actual_hash:
        fail(f"{label}: binary hash differs from build identity")
    if expected_bytes != actual_bytes:
        fail(f"{label}: binary size differs from build identity")
    mode = path.stat().st_mode
    if not mode & 0o111:
        fail(f"{label}: binary is not executable")
    return {
        "path": str(path),
        "sha256": actual_hash,
        "bytes": actual_bytes,
        "mode_bits": oct(mode & 0o777),
    }


def _load_context(root: Path, *, require_binaries: bool = True) -> dict[str, Any]:
    protocol_path = root / "protocol.json"
    build_path = root / "build-identity.json"
    capture_path = root / "capture.json"
    profile_path = root / "profile.json"
    protocol = _require_mapping(load_json(protocol_path), "protocol")
    build = _require_mapping(load_json(build_path), "build identity")
    capture = _require_mapping(load_json(capture_path), "capture")
    profile = _require_mapping(load_json(profile_path), "profile")
    if protocol.get("change") != EXPECTED_CHANGE:
        fail("protocol change is not 0418")
    if build.get("change") != EXPECTED_CHANGE or build.get("status") != "complete":
        fail("build identity is not a complete 0418 build")
    if capture.get("change") != EXPECTED_CHANGE or capture.get("status") != "complete":
        fail("capture is not a complete 0418 capture")
    if profile.get("change") != EXPECTED_CHANGE or profile.get("status") != "complete":
        fail("profile is not a complete 0418 profile")
    protocol_hash = sha256_file(protocol_path)
    build_hash = sha256_file(build_path)
    capture_hash = sha256_file(capture_path)
    profile_hash = sha256_file(profile_path)
    if profile.get("protocol", {}).get("sha256") != protocol_hash:
        fail("profile protocol hash does not match protocol.json")
    if profile.get("build_identity", {}).get("sha256") != build_hash:
        fail("profile build identity hash does not match build-identity.json")
    if profile.get("capture", {}).get("sha256") != capture_hash:
        fail("profile capture hash does not match capture.json")
    if protocol.get("profile", {}).get("samples") != 20 or protocol.get("profile", {}).get("warmups") != 3:
        fail("protocol profile boundary must be 20 samples and 3 warmups")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol profile must be CPU 2 with one worker")
    profile_execution = _require_mapping(profile.get("execution"), "profile.execution")
    selectors = profile_execution.get("selectors")
    if not isinstance(selectors, list) or not selectors:
        fail("profile.execution.selectors must be a non-empty list")
    normalized_selectors = protocol.get("profile_selectors")
    if normalized_selectors is None:
        configured = protocol.get("profile", {}).get("selectors")
        if configured is None:
            jobs = protocol.get("jobs")
            if not isinstance(jobs, list) or not jobs or not isinstance(jobs[0], dict):
                fail("protocol has no default profile selector")
            configured = jobs[0].get("selector")
        if isinstance(configured, str):
            normalized_selectors = [configured]
        elif isinstance(configured, list):
            normalized_selectors = configured
    if not isinstance(normalized_selectors, list) or selectors != normalized_selectors:
        fail("profile selectors differ from protocol profile_selectors")
    if profile_execution.get("cpu") != 2 or profile_execution.get("workers") != 1:
        fail("profile execution CPU/workers differ from protocol")
    if profile_execution.get("profile", {}).get("samples") != 20 or profile_execution.get("profile", {}).get("warmups") != 3:
        fail("profile execution boundary differs from protocol")
    roles = _require_mapping(build.get("roles"), "build.roles")
    if set(roles) != set(ROLES):
        fail("build roles must contain control and candidate")
    binaries: dict[str, dict[str, dict[str, Any]]] = {}
    for role in ROLES:
        role_map = _require_mapping(roles[role], f"build.roles.{role}")
        mode_map = _require_mapping(role_map.get("binaries"), f"build.roles.{role}.binaries")
        binaries[role] = {}
        for mode in MODES:
            descriptor = _require_mapping(mode_map.get(mode), f"build.roles.{role}.binaries.{mode}")
            binary_path = Path(_require_string(descriptor.get("path"), f"{role}/{mode}.path"))
            if require_binaries:
                identity = _hash_descriptor(binary_path, descriptor, f"{role}/{mode}")
            else:
                expected_hash = descriptor.get("sha256")
                expected_bytes = descriptor.get("bytes")
                if not isinstance(expected_hash, str) or len(expected_hash) != 64:
                    fail(f"{role}/{mode}: invalid recorded binary hash")
                if not isinstance(expected_bytes, int) or expected_bytes < 1:
                    fail(f"{role}/{mode}: invalid recorded binary size")
                identity = {
                    "path": str(binary_path), "sha256": expected_hash,
                    "bytes": expected_bytes, "mode_bits": descriptor.get("mode_bits"),
                }
            binaries[role][mode] = {"descriptor": descriptor, "identity": identity}
    gate_path = root / "checks" / "preflight-validation.json"
    if not gate_path.is_file():
        fail("checks/preflight-validation.json is required")
    gate = _require_mapping(load_json(gate_path), "preflight validation")
    if gate.get("status") != "pass":
        fail("preflight validation gate did not pass")
    return {
        "root": root,
        "protocol": protocol,
        "build": build,
        "capture": capture,
        "profile": profile,
        "paths": {
            "protocol": protocol_path,
            "build": build_path,
            "capture": capture_path,
            "profile": profile_path,
            "gate": gate_path,
        },
        "hashes": {
            "protocol": protocol_hash,
            "build": build_hash,
            "capture": capture_hash,
            "profile": profile_hash,
            "gate": sha256_file(gate_path),
        },
        "binaries": binaries,
        "selectors": selectors,
    }


def _record_for(context: dict[str, Any], role: str, phase: str, selector: str) -> dict[str, Any]:
    records = [
        item for item in context["profile"].get("records", [])
        if isinstance(item, dict)
        and item.get("stage") == "complete"
        and item.get("phase") == phase
        and item.get("role") == role
        and item.get("selector") == selector
    ]
    if len(records) != 1:
        fail(f"profile must contain exactly one complete {phase} record for {role}/{selector}")
    return records[0]


def _capture_record(context: dict[str, Any], role: str, selector: str) -> dict[str, Any]:
    # A1/B1 are the first role occurrence and make the binding deterministic;
    # the capture validator has already checked all ABBA legs.
    leg = "A1" if role == "control" else "B1"
    records = [
        item for item in context["capture"].get("runs", [])
        if isinstance(item, dict)
        and item.get("phase") == "normal"
        and item.get("preflight") is False
        and item.get("leg") == leg
        and item.get("role") == role
        and item.get("selector") == selector
        and item.get("exit_code") == 0
    ]
    if len(records) != 1:
        fail(f"capture must contain one successful formal {leg} record for {role}/{selector}")
    return records[0]


def _contains_subsequence(values: list[Any], expected: list[Any]) -> bool:
    if len(expected) > len(values):
        return False
    return any(values[index:index + len(expected)] == expected for index in range(len(values) - len(expected) + 1))


def _validate_run_identity(context: dict[str, Any], record: dict[str, Any], role: str, mode: str, label: str) -> None:
    expected_source = context["build"]["roles"][role]["source"]
    expected_binary = context["binaries"][role][mode]["identity"]
    if record.get("role") not in (None, role):
        fail(f"{label}: record role differs from {role}")
    if record.get("binary_sha256") != expected_binary["sha256"] or record.get("binary_bytes") != expected_binary["bytes"]:
        fail(f"{label}: record binary identity differs from build")
    if record.get("source_revision", record.get("source", {}).get("revision")) != expected_source["revision"]:
        fail(f"{label}: record source revision differs from build")
    source = record.get("source")
    if not isinstance(source, dict):
        source = {
            "revision": record.get("source_revision"),
            "clean": record.get("source_after", {}).get("clean", True),
            "source_files": record.get("source_files"),
        }
    if source.get("clean") is not True:
        fail(f"{label}: record source is not clean")
    if source.get("source_files") != expected_source.get("source_files"):
        fail(f"{label}: record source file hash manifest differs from build")
    after = record.get("source_after")
    if isinstance(after, dict):
        if after.get("revision") != expected_source["revision"] or after.get("clean") is not True:
            fail(f"{label}: post-run source binding differs from build")
        if after.get("source_files") != expected_source.get("source_files"):
            fail(f"{label}: post-run source file hash manifest differs from build")
    argv = record.get("argv")
    if not isinstance(argv, list) or not all(isinstance(value, str) for value in argv):
        fail(f"{label}: record argv is missing or malformed")
    common_flags = context["protocol"].get("common_flags")
    if not isinstance(common_flags, list) or not all(isinstance(value, str) for value in common_flags):
        fail("protocol.common_flags is missing or malformed")
    if not _contains_subsequence(argv, common_flags):
        fail(f"{label}: argv does not contain the protocol common flags")
    samples = record.get("samples")
    warmups = record.get("warmups")
    if isinstance(samples, int) and isinstance(warmups, int):
        if "--samples" not in argv or "--warmup" not in argv:
            fail(f"{label}: argv lacks sample boundary flags")
        if argv[argv.index("--samples") + 1] != str(samples) or argv[argv.index("--warmup") + 1] != str(warmups):
            fail(f"{label}: argv sample boundary differs from record")


def _artifact_descriptor(record: dict[str, Any], name: str, label: str) -> dict[str, Any]:
    artifacts = record.get("artifacts")
    if isinstance(artifacts, dict) and isinstance(artifacts.get(name), dict):
        return artifacts[name]
    # Formal capture records use top-level report/catalog fields and keep
    # their digests in the artifact_sha256 object.  Normalize that shape to
    # the profile driver's nested artifact descriptor here.
    relative = record.get(name)
    digests = record.get("artifact_sha256")
    if isinstance(relative, str) and isinstance(digests, dict):
        digest = digests.get(name)
        if isinstance(digest, str):
            path = Path(_relative(relative, f"{label}.{name}.path"))
            return {"path": path.as_posix(), "bytes": None, "sha256": digest, "formal_bytes_unknown": True}
    fail(f"{label} is missing artifact descriptor {name!r}")
    raise AssertionError("unreachable")


def _artifact_from_record(root: Path, record: dict[str, Any], name: str, label: str) -> tuple[bytes, dict[str, Any]]:
    descriptor = _artifact_descriptor(record, name, label)
    data, info = read_artifact(root, descriptor.get("path"), label)
    if descriptor.get("sha256") != info["sha256"]:
        fail(f"{label}.{name} does not match its manifest hash/size")
    if not descriptor.get("formal_bytes_unknown") and descriptor.get("bytes") != info["bytes"]:
        fail(f"{label}.{name} does not match its manifest hash/size")
    return data, info


def _stable_catalog(value: Any) -> Any:
    if not isinstance(value, dict):
        return value
    # Paths and build metadata may differ between a profile's fresh output and
    # the formal capture.  These fields are the immutable corpus identity.
    keys = (
        "manifest_version", "manifest_kind", "catalog_id", "canonicalization",
        "catalog_sha256", "content_set_sha256", "corpora", "case_bindings",
    )
    return {key: value[key] for key in keys if key in value}


def _report_identity(report: dict[str, Any], role: str, selector: str, samples: int, warmups: int,
                     expected_binary: dict[str, Any], protocol: dict[str, Any], label: str) -> dict[str, Any]:
    if report.get("tool", {}).get("binary") not in ("litchi-perf-baseline", "litchi-perf-baseline-alloc"):
        fail(f"{label}: unexpected harness binary")
    if report.get("configuration", {}).get("samples_per_case") != samples:
        fail(f"{label}: samples_per_case does not match boundary")
    if report.get("configuration", {}).get("warmup_iterations_per_case") != warmups:
        fail(f"{label}: warmup_iterations_per_case does not match boundary")
    cases = report.get("configuration", {}).get("cases")
    if cases != [selector]:
        fail(f"{label}: report cases do not identify {selector!r}")
    environment = _require_mapping(report.get("environment"), f"{label}.environment")
    binary = _require_mapping(report.get("binary_identity"), f"{label}.binary_identity")
    source_role = _require_mapping(expected_binary, f"{label}.expected_binary")
    if binary.get("binary_sha256") != source_role.get("sha256"):
        fail(f"{label}: report binary hash differs from profile build")
    if binary.get("binary_bytes") != source_role.get("bytes"):
        fail(f"{label}: report binary size differs from profile build")
    if environment.get("git_worktree_dirty") is not False:
        fail(f"{label}: report source tree must be clean")
    revision = environment.get("git_revision")
    if not isinstance(revision, str):
        fail(f"{label}: report has no source revision")
    # The build identity is authoritative; protocol roles are optional.
    return {
        "revision": revision,
        "binary_sha256": binary.get("binary_sha256"),
        "binary_bytes": binary.get("binary_bytes"),
        "cpu_affinity": environment.get("cpu_affinity"),
        "workers": report.get("configuration", {}).get("execution_workers"),
        "instrumentation": report.get("tool", {}).get("instrumentation"),
    }


def _validate_report_pair(context: dict[str, Any], role: str, selector: str,
                         profile_report: dict[str, Any], profile_catalog: dict[str, Any],
                         formal_report: dict[str, Any], formal_catalog: dict[str, Any],
                         profile_samples: int, profile_warmups: int, label: str) -> dict[str, Any]:
    expected = context["binaries"][role]["normal"]["identity"]
    profile_identity = _report_identity(
        profile_report, role, selector, profile_samples, profile_warmups, expected,
        context["protocol"], f"{label}.profile",
    )
    formal_identity = _report_identity(
        formal_report, role, selector, formal_report.get("configuration", {}).get("samples_per_case", -1),
        formal_report.get("configuration", {}).get("warmup_iterations_per_case", -1), expected,
        context["protocol"], f"{label}.formal",
    )
    expected_revision = context["build"]["roles"][role]["source"]["revision"]
    if profile_identity["revision"] != expected_revision or formal_identity["revision"] != expected_revision:
        fail(f"{label}: report source revision differs from build identity")
    if formal_report.get("configuration", {}).get("samples_per_case") != 500 or formal_report.get("configuration", {}).get("warmup_iterations_per_case") != 20:
        fail(f"{label}.formal: formal comparison must be the 500/20 normal run")
    if profile_identity["binary_sha256"] != formal_identity["binary_sha256"]:
        fail(f"{label}: profile and formal binary identities differ")
    if profile_identity["revision"] != formal_identity["revision"]:
        fail(f"{label}: profile and formal source revisions differ")
    profile_rows = profile_report.get("results")
    formal_rows = formal_report.get("results")
    if not isinstance(profile_rows, list) or len(profile_rows) != 1 or not isinstance(formal_rows, list) or len(formal_rows) != 1:
        fail(f"{label}: profile/formal reports must contain one result row")
    profile_row, formal_row = profile_rows[0], formal_rows[0]
    if profile_row.get("case") != selector or formal_row.get("case") != selector:
        fail(f"{label}: result case does not match selector")
    if profile_row.get("output_sha256") != formal_row.get("output_sha256"):
        fail(f"{label}: profile/formal output hashes differ")
    if _stable_catalog(profile_catalog) != _stable_catalog(formal_catalog):
        fail(f"{label}: profile/formal corpus catalogs differ")
    if profile_row.get("corpus") != formal_row.get("corpus"):
        fail(f"{label}: profile/formal corpus bindings differ")
    # Common flags are captured as part of the report's input identity.  Keep
    # this check explicit so a future profile cannot silently change shape.
    if profile_report.get("configuration", {}).get("execution_workers") != [1]:
        fail(f"{label}: profile execution_workers must be [1]")
    if formal_report.get("configuration", {}).get("execution_workers") != [1]:
        fail(f"{label}: formal execution_workers must be [1]")
    return {
        "selector": selector,
        "role": role,
        "profile": profile_identity,
        "formal": formal_identity,
        "corpus": profile_row.get("corpus"),
        "output_sha256": profile_row.get("output_sha256"),
        "catalog_identity": _stable_catalog(profile_catalog),
        "formal_report_sha256": sha256_bytes(canonical(formal_report)),
        "formal_catalog_sha256": sha256_bytes(canonical(formal_catalog)),
    }


def _formal_binding(context: dict[str, Any], role: str, selector: str, profile_record: dict[str, Any]) -> dict[str, Any]:
    _validate_run_identity(context, profile_record, role, "normal", f"profile/{role}/{selector}")
    formal = _capture_record(context, role, selector)
    _validate_run_identity(context, formal, role, "normal", f"formal/{role}/{selector}")
    profile_report_data, profile_report_info = _artifact_from_record(
        context["root"], profile_record, "workload_report", f"profile/{role}/{selector}"
    )
    profile_catalog_data, profile_catalog_info = _artifact_from_record(
        context["root"], profile_record, "workload_catalog", f"profile/{role}/{selector}"
    )
    formal_report_data, formal_report_info = _artifact_from_record(
        context["root"], formal, "report", f"formal/{role}/{selector}"
    )
    formal_catalog_data, formal_catalog_info = _artifact_from_record(
        context["root"], formal, "catalog", f"formal/{role}/{selector}"
    )
    profile_report = _require_mapping(json.loads(profile_report_data.decode("utf-8"), object_pairs_hook=_pairs, parse_constant=_constant), "profile report")
    profile_catalog = _require_mapping(json.loads(profile_catalog_data.decode("utf-8"), object_pairs_hook=_pairs, parse_constant=_constant), "profile catalog")
    formal_report = _require_mapping(json.loads(formal_report_data.decode("utf-8"), object_pairs_hook=_pairs, parse_constant=_constant), "formal report")
    formal_catalog = _require_mapping(json.loads(formal_catalog_data.decode("utf-8"), object_pairs_hook=_pairs, parse_constant=_constant), "formal catalog")
    if profile_record.get("samples") != 20 or profile_record.get("warmups") != 3:
        fail(f"profile/{role}/{selector}: record boundary must be 20/3")
    if profile_report.get("tool", {}).get("instrumentation") != "none":
        fail(f"profile/{role}/{selector}: CPU record must use normal instrumentation")
    validation = _validate_report_pair(
        context, role, selector, profile_report, profile_catalog,
        formal_report, formal_catalog, 20, 3, f"binding/{role}/{selector}"
    )
    return {
        **validation,
        "profile_report": profile_report_info,
        "profile_catalog": profile_catalog_info,
        "formal_report": formal_report_info,
        "formal_catalog": formal_catalog_info,
        "formal_leg": formal.get("leg"),
        "formal_key": formal.get("key"),
    }


def _elf_comments(context: dict[str, Any], *, execute: bool, existing: Any = None) -> list[dict[str, Any]]:
    expected: list[dict[str, Any]] = []
    for role in ROLES:
        for mode in MODES:
            identity = context["binaries"][role][mode]["identity"]
            binary = Path(identity["path"])
            argv = ["readelf", "--string-dump=.comment", str(binary)]
            if execute:
                process = subprocess.run(
                    argv, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
                    env={**os.environ, "DEBUGINFOD_URLS": ""},
                )
                stdout, stderr, exit_code = process.stdout, process.stderr, process.returncode
            else:
                fail("internal error: ELF comment replay requested without execution")
            if exit_code != 0:
                fail(f"readelf .comment failed for {role}/{mode}: {stderr.decode(errors='replace')}")
            expected.append({
                "role": role, "mode": mode, "argv": argv,
                "binary": identity, "exit_code": exit_code,
                "stdout": stdout.decode("utf-8", errors="replace"),
                "stderr": stderr.decode("utf-8", errors="replace"),
                "stdout_sha256": sha256_bytes(stdout), "stderr_sha256": sha256_bytes(stderr),
            })
    if existing is not None and canonical(existing) != canonical(expected):
        fail("ELF .comment replay differs from retained elf-comments.json")
    return expected


def _command_specs(data_path: Path, stem: str) -> list[tuple[str, list[str], str]]:
    return [
        ("perf-header", ["perf", "report", "--header-only", "-i", str(data_path)], f"{stem}.header.txt"),
        ("perf-self", ["perf", "report", "--stdio", "--no-inline", "--no-children", "--percent-limit", "0", "--sort", "symbol,dso", "-i", str(data_path)], f"{stem}.self.txt"),
        ("perf-children", ["perf", "report", "--stdio", "--no-inline", "--children", "--percent-limit", "0", "--sort", "symbol,dso", "-i", str(data_path)], f"{stem}.children.txt"),
        ("perf-script", ["perf", "script", "--no-inline", "-F", "comm,pid,tid,time,period,event,ip,sym,dso", "-i", str(data_path)], f"{stem}.script.txt"),
    ]


def _run_postprocess(context: dict[str, Any], role: str, selector: str, record: dict[str, Any],
                     command_records: list[dict[str, Any]], temporary: Path | None = None) -> None:
    root = context["root"]
    profile_record_artifact = _artifact_descriptor(record, "data", f"profile/{role}/{selector}")
    data, info = read_artifact(root, profile_record_artifact.get("path"), f"profile/{role}/{selector}.data")
    expected = profile_record_artifact.get("sha256")
    if expected != info["sha256"] or profile_record_artifact.get("bytes") != info["bytes"]:
        fail(f"profile/{role}/{selector}.data does not match profile.json")
    with tempfile.TemporaryDirectory(prefix="litchi-0418-perf-") as temp_dir:
        data_path = Path(temp_dir) / "record.data"
        data_path.write_bytes(data)
        stem = f"{role}-{selector}"
        output_root = temporary if temporary is not None else root / "profile"
        output_root.mkdir(parents=True, exist_ok=True)
        for kind, argv, output_name in _command_specs(data_path, stem):
            output = output_root / output_name
            stderr_path = output_root / f"{output_name}.stderr"
            if temporary is None and output.exists():
                fail(f"refusing to overwrite existing analysis artifact {output}")
            started = utc_now()
            with output.open("wb") as stdout, stderr_path.open("wb") as stderr:
                process = subprocess.run(
                    argv, cwd=root, stdout=stdout, stderr=stderr,
                    env={**os.environ, "DEBUGINFOD_URLS": ""}, check=False,
                )
            finished = utc_now()
            stdout_data = output.read_bytes() if output.is_file() else b""
            stderr_data = stderr_path.read_bytes() if stderr_path.is_file() else b""
            command = {
                "role": role, "selector": selector, "kind": kind,
                "argv": argv, "cwd": str(root),
                "environment_overrides": {"DEBUGINFOD_URLS": ""},
                "input": info, "output": relpath(output, root),
                "stderr": relpath(stderr_path, root),
                "started_utc": started, "finished_utc": finished,
                "exit_code": process.returncode,
                "output_bytes": len(stdout_data), "output_sha256": sha256_bytes(stdout_data),
                "stderr_bytes": len(stderr_data), "stderr_sha256": sha256_bytes(stderr_data),
            }
            command_records.append(command)
            if process.returncode != 0:
                fail(f"{kind} failed for {role}/{selector}: {stderr_data.decode(errors='replace')}")


def _output_bytes(root: Path, command: dict[str, Any], label: str) -> tuple[bytes, dict[str, Any]]:
    data, info = read_artifact(root, command.get("output"), label)
    if command.get("output_sha256") != info["sha256"] or command.get("output_bytes") != info["bytes"]:
        fail(f"{label}: output hash/size differs from analysis command record")
    return data, info


def _replay_commands(context: dict[str, Any], commands: dict[str, Any]) -> None:
    records = commands.get("commands")
    if not isinstance(records, list) or not records:
        fail("analysis-commands.json has no command records")
    by_pair: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for command in records:
        if isinstance(command, dict):
            by_pair.setdefault((command.get("role"), command.get("selector")), []).append(command)
    for role in ROLES:
        for selector in context["selectors"]:
            key_records = by_pair.get((role, selector), [])
            if {item.get("kind") for item in key_records} != {"perf-header", "perf-self", "perf-children", "perf-script"}:
                fail(f"analysis commands incomplete for {role}/{selector}")
            data_descriptor = _artifact_descriptor(_record_for(context, role, "cpu-record", selector), "data", "profile/{role}/{selector}")
            data, info = read_artifact(context["root"], data_descriptor.get("path"), f"profile/{role}/{selector}.data")
            if data_descriptor.get("sha256") != info["sha256"] or data_descriptor.get("bytes") != info["bytes"]:
                fail(f"profile/{role}/{selector}.data changed since profile capture")
            with tempfile.TemporaryDirectory(prefix="litchi-0418-replay-") as temp_dir:
                data_path = Path(temp_dir) / "record.data"
                data_path.write_bytes(data)
                for command in sorted(key_records, key=lambda item: item.get("kind", "")):
                    argv = list(command.get("argv", []))
                    if not argv or argv[-2:] == []:
                        fail("analysis command has no argv")
                    if "-i" not in argv:
                        fail(f"{command.get('kind')} command has no input")
                    input_index = max(index for index, value in enumerate(argv) if value == "-i")
                    if input_index + 1 >= len(argv):
                        fail("analysis command has incomplete -i")
                    argv[input_index + 1] = str(data_path)
                    stdout_path = Path(temp_dir) / f"{command['kind']}.out"
                    stderr_path = Path(temp_dir) / f"{command['kind']}.err"
                    process = subprocess.run(
                        argv, cwd=context["root"], stdout=stdout_path.open("wb"), stderr=stderr_path.open("wb"),
                        env={**os.environ, "DEBUGINFOD_URLS": ""}, check=False,
                    )
                    actual = stdout_path.read_bytes()
                    actual_error = stderr_path.read_bytes()
                    if process.returncode != command.get("exit_code"):
                        fail(f"replay exit differs for {command.get('kind')}")
                    if sha256_bytes(actual) != command.get("output_sha256") or len(actual) != command.get("output_bytes"):
                        fail(f"replay output differs for {role}/{selector}/{command.get('kind')}")
                    if sha256_bytes(actual_error) != command.get("stderr_sha256") or len(actual_error) != command.get("stderr_bytes"):
                        fail(f"replay stderr differs for {role}/{selector}/{command.get('kind')}")


def _build_command_manifest(context: dict[str, Any], command_records: list[dict[str, Any]], comments: list[dict[str, Any]], bindings: list[dict[str, Any]]) -> dict[str, Any]:
    root = context["root"]
    profile_path = context["paths"]["profile"]
    return {
        "schema_version": 1,
        "change": EXPECTED_CHANGE,
        "status": "complete",
        "created_utc": utc_now(),
        "protocol": {"path": relpath(context["paths"]["protocol"], root), "sha256": context["hashes"]["protocol"]},
        "build_identity": {"path": relpath(context["paths"]["build"], root), "sha256": context["hashes"]["build"]},
        "capture": {"path": relpath(context["paths"]["capture"], root), "sha256": context["hashes"]["capture"]},
        "profile": {"path": relpath(profile_path, root), "sha256": context["hashes"]["profile"]},
        "gates": {"preflight_validation": {"path": relpath(context["paths"]["gate"], root), "sha256": context["hashes"]["gate"], "status": "pass"}},
        "boundary": {"samples": 20, "warmups": 3, "cpu": 2, "workers": 1, "event": "cycles:u", "call_graph": "fp,127"},
        "roles": {
            role: {
                mode: {"path": context["binaries"][role][mode]["identity"]["path"], "sha256": context["binaries"][role][mode]["identity"]["sha256"], "bytes": context["binaries"][role][mode]["identity"]["bytes"]}
                for mode in MODES
            }
            for role in ROLES
        },
        "formal_bindings": bindings,
        "elf_comments": comments,
        "commands": command_records,
        "scope": "perf report/script are whole-command CPU period views; setup, corpus generation, warmups, timed children, verification and output/report writing remain in scope; no phase or speedup claim",
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT, help="change-0418 result directory")
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--write", action="store_true", help="run postprocessing and retain its outputs")
    action.add_argument("--replay", action="store_true", help="re-run postprocessing and compare retained bytes")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root = args.root.resolve()
    try:
        context = _load_context(root)
        commands_path = root / "profile" / "analysis-commands.json"
        comments_path = root / "profile" / "elf-comments.json"
        if args.replay:
            commands = _require_mapping(load_json(commands_path), "analysis commands")
            if commands.get("status") != "complete":
                fail("analysis command manifest is not complete")
            for key in ("protocol", "build_identity", "capture", "profile"):
                if commands.get(key, {}).get("sha256") != context["hashes"][{"protocol": "protocol", "build_identity": "build", "capture": "capture", "profile": "profile"}[key]]:
                    fail(f"analysis command {key} hash is stale")
            _replay_commands(context, commands)
            comments = _require_mapping(load_json(comments_path), "ELF comments")
            _elf_comments(context, execute=True, existing=comments.get("entries"))
            print("0418 profile analysis replay passed")
            return 0
        if commands_path.exists() or comments_path.exists():
            fail("refusing --write because analysis artifacts already exist; use --replay")
        command_records: list[dict[str, Any]] = []
        bindings: list[dict[str, Any]] = []
        for role in ROLES:
            for selector in context["selectors"]:
                record = _record_for(context, role, "cpu-record", selector)
                bindings.append(_formal_binding(context, role, selector, record))
                _run_postprocess(context, role, selector, record, command_records)
        comments = _elf_comments(context, execute=True)
        manifest = _build_command_manifest(context, command_records, comments, bindings)
        write_json(commands_path, manifest)
        write_json(comments_path, {"schema_version": 1, "change": EXPECTED_CHANGE, "entries": comments})
        print(f"0418 profile analysis complete: {commands_path}")
        return 0
    except (AnalysisError, OSError, subprocess.SubprocessError) as error:
        print(f"0418 profile analysis failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
