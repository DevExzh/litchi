#!/usr/bin/env python3
"""Portable, fail-closed verifier for the 0436 ODT streaming ABBA matrix.

This verifier consumes retained reports, receipts, build descriptors, and the
copied 0436 report oracle.  It does not invoke Cargo, Git, a workload, or a
binary.  Build binaries may be removed after their immutable path/SHA/size
descriptor has been retained; ``--require-binaries`` is available for the
pre-cleanup gate.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
# Formal descriptors use absolute binary paths.  Keep relative-path support
# for an in-tree invocation, while allowing a portable export whose parent
# chain is shorter and contains no live checkout.
REPO = ROOT.parents[3] if len(ROOT.parents) > 3 else ROOT
PHASES = {
    "A1": ("before-streaming", 0, 6, "R1"),
    "B1": ("after-streaming", 0, 6, "R1"),
    "B2": ("after-streaming", 6, 12, "R2"),
    "A2": ("before-streaming", 6, 12, "R2"),
}
PHASE_ORDER = tuple(PHASES)
BUILD_DIR = {"before-streaming": "before", "after-streaming": "after"}
PROFILE_ROLES = ("before-streaming", "after-streaming")
PROFILE_KINDS = ("stat", "record")
PROFILE_EVENTS = ["cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u"]
ODT_MEMBERS = ("mimetype", "content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml")
ODT_MIMETYPE = "application/vnd.oasis.opendocument.text"
STAGES = ("precleanup", "aftercleanup", "final")
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
RSS_RE = re.compile(r"^Maximum resident set size \(kbytes\):\s*(\d+)\s*$")
MAX_JSON_BYTES = 512 * 1024 * 1024


class VerificationError(ValueError):
    pass


def fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")


def duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> Any:
    raise VerificationError(f"non-finite JSON value {value!r}")


def load(path: Path, label: str | None = None) -> Any:
    label = label or str(path)
    if not path.is_file():
        fail(label, "file is missing")
    if path.stat().st_size > MAX_JSON_BYTES:
        fail(label, "JSON file is too large")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=duplicate_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, VerificationError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def bundle_path(value: str, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute():
        fail(label, "must be a non-empty bundle-relative path")
    path = (ROOT / value).resolve()
    if not path.is_relative_to(ROOT.resolve()):
        fail(label, "path escapes the evidence bundle")
    return path


def timestamp(value: Any, label: str) -> dt.datetime:
    if not isinstance(value, str) or not value:
        fail(label, "must be a non-empty ISO-8601 timestamp")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(label, f"invalid ISO-8601 timestamp: {error}")
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        fail(label, "timestamp must include a timezone")
    return parsed


def digest(value: Any, path: str) -> str:
    if not isinstance(value, str) or HEX64.fullmatch(value) is None:
        fail(path, "expected a SHA-256 hexadecimal digest")
    return value.lower()


def u64(value: Any, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0 or value > (1 << 64) - 1:
        fail(path, "expected a u64")
    return value


def obj(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    return value


def array(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "expected an array")
    return value


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()


def load_oracle():
    path = ROOT / "oracle" / "verify-report.py"
    spec = importlib.util.spec_from_file_location("change0436_oracle", path)
    if spec is None or spec.loader is None:
        raise VerificationError("cannot load copied oracle verifier")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def oracle_role(protocol: dict[str, Any], role: str) -> str:
    """Resolve a role-specific copied-oracle role without guessing candidates."""

    spec = obj(protocol["roles"][role], f"protocol.roles.{role}")
    if isinstance(spec.get("oracle_role"), str):
        return spec["oracle_role"]
    role_oracle = spec.get("oracle")
    if isinstance(role_oracle, dict) and isinstance(role_oracle.get("role"), str):
        return role_oracle["role"]
    mapped = protocol.get("oracle_roles")
    if isinstance(mapped, dict) and isinstance(mapped.get(role), str):
        return mapped[role]
    global_oracle = protocol.get("oracle")
    if isinstance(global_oracle, dict):
        mapped = global_oracle.get("roles")
        if isinstance(mapped, dict) and isinstance(mapped.get(role), str):
            return mapped[role]
    return role


def validate_protocol(protocol: dict[str, Any]) -> None:
    if protocol.get("change") != 436:
        fail("protocol.change", "must be 436")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol", "must bind CPU 2 and one worker")
    if protocol.get("samples") != 30 or protocol.get("warmups") != 3 or protocol.get("repeats") != 2:
        fail("protocol", "must retain samples=30, warmups=3, repeats=2")
    if protocol.get("modes") != ["normal", "allocator"]:
        fail("protocol.modes", "must be normal then allocator")
    shapes = protocol.get("shapes")
    if shapes != {"tiny": 64, "medium": 8192, "large": 32768}:
        fail("protocol.shapes", "does not match the frozen three-shape matrix")
    order = array(protocol.get("order"), "protocol.order")
    if len(order) != 12:
        fail("protocol.order", "must contain 12 lanes")
    expected_order = [
        ("normal", "tiny", "R1"), ("normal", "medium", "R1"), ("normal", "large", "R1"),
        ("allocator", "tiny", "R1"), ("allocator", "medium", "R1"), ("allocator", "large", "R1"),
        ("allocator", "large", "R2"), ("allocator", "medium", "R2"), ("allocator", "tiny", "R2"),
        ("normal", "large", "R2"), ("normal", "medium", "R2"), ("normal", "tiny", "R2"),
    ]
    actual = [(obj(row, f"protocol.order[{i}]").get("mode"), row.get("shape"), row.get("repeat")) for i, row in enumerate(order)]
    if actual != expected_order:
        fail("protocol.order", f"expected {expected_order!r}, got {actual!r}")
    roles = obj(protocol.get("roles"), "protocol.roles")
    expected_roles = {
        "before-streaming": ("odt_streaming_create", "streaming", "litchi_odt::streaming::stream_plain_paragraphs_to", "fixed_explicit_window"),
        "after-streaming": ("odt_streaming_create", "streaming", "litchi_odt::streaming::stream_plain_paragraphs_to", "fixed_explicit_window"),
    }
    if set(roles) != set(expected_roles):
        fail("protocol.roles", "must contain exactly the before and after streaming roles")
    for role, (selector, source_role, implementation, retention) in expected_roles.items():
        spec = obj(roles[role], f"protocol.roles.{role}")
        if (spec.get("selector"), spec.get("source_role"), spec.get("implementation"), spec.get("retention")) != (selector, source_role, implementation, retention):
            fail(f"protocol.roles.{role}", "selector, source role, implementation, or retention differs")
        if spec.get("build_directory") != BUILD_DIR[role]:
            fail(f"protocol.roles.{role}.build_directory", "does not match the role build")
        if spec.get("source_field") != "odt_paragraphs":
            fail(f"protocol.roles.{role}.source_field", "must be odt_paragraphs")
        if spec.get("retained_authoring_window_bytes") != 4096:
            fail(f"protocol.roles.{role}.retained_authoring_window_bytes", "must bind the fixed 4096-byte streaming window")
    corpus = obj(protocol.get("corpus"), "protocol.corpus")
    if corpus.get("package_format") != "ODT/ODF/ZIP" or corpus.get("compression") != "mimetype=stored;xml=deflate" or corpus.get("mimetype") != "application/vnd.oasis.opendocument.text":
        fail("protocol.corpus", "does not identify the ODT corpus")
    if corpus.get("member_names") != ["mimetype", "content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml"] or corpus.get("manifest_entry_count") != 4 or corpus.get("target_entry") != "content.xml":
        fail("protocol.corpus", "member topology differs from the ODT contract")
    contract = obj(protocol.get("source_contract"), "protocol.source_contract")
    required_source_fields = contract.get("required_fields")
    if not isinstance(required_source_fields, list) or "semantic_sha256" not in required_source_fields or "styles_xml_sha256" not in required_source_fields or "meta_xml_sha256" not in required_source_fields:
        fail("protocol.source_contract.required_fields", "does not bind semantic, style, and metadata identities")
    if contract.get("all_boolean_gates_required_true") is not True:
        fail("protocol.source_contract.all_boolean_gates_required_true", "must require every correctness gate")
    identity = obj(protocol.get("identity_scope"), "protocol.identity_scope")
    identity_fields = {
        "archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256",
        "content_xml_sha256", "styles_xml_sha256", "meta_xml_sha256", "paragraph_count",
        "member_names", "mimetype", "manifest_entry_count", "semantic_sha256", "output_sha256",
        "sink_accepted_bytes", "sink_write_calls",
    }
    for name in ("same_role_required", "cross_role_required"):
        fields = identity.get(name)
        if not isinstance(fields, list) or len(fields) != len(identity_fields) or set(fields) != identity_fields:
            fail(f"protocol.identity_scope.{name}", "must bind the complete archive, content, semantic, output, and sink identity")
    if contract.get("same_role_content_identity_required") is not True:
        fail("protocol.source_contract.same_role_content_identity_required", "same-role determinism is required")
    report_contract = obj(protocol.get("report_contract"), "protocol.report_contract")
    if report_contract.get("required_reports") != 24 or report_contract.get("reports_per_role") != 12 or report_contract.get("retained_samples") != 720:
        fail("protocol.report_contract", "must bind 24 reports, 12 reports per role, and 720 retained samples")
    profile = protocol.get("profile_scope")
    if profile is not None:
        profile = obj(profile, "protocol.profile_scope")
        if profile.get("roles") != list(PROFILE_ROLES) or profile.get("mode") != "normal" or profile.get("shape") != "large":
            fail("protocol.profile_scope", "profile role/mode/shape scope differs")
        if profile.get("captures") != ["perf-stat", "perf-record"] or profile.get("events") != PROFILE_EVENTS or profile.get("call_graph") != "fp":
            fail("protocol.profile_scope", "profile event or call-graph scope differs")
        if profile.get("required") is not True:
            fail("protocol.profile_scope.required", "all four profiles must be declared")
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    oracle_protocol = bundle_path(oracle.get("path"), "protocol.oracle.path")
    oracle_verifier = bundle_path(oracle.get("verifier_path"), "protocol.oracle.verifier_path")
    if oracle.get("sha256") != sha(oracle_protocol) or oracle.get("verifier_sha256") != sha(oracle_verifier):
        fail("protocol.oracle", "does not match the retained oracle files")


def check_source_manifest(manifest: dict[str, Any], label: str) -> None:
    path_value = manifest.get("path")
    expected_sha = manifest.get("sha256")
    files = manifest.get("files")
    if not isinstance(path_value, str) or Path(path_value).is_absolute() or not isinstance(files, int) or files < 1:
        fail(label, "malformed source manifest descriptor")
    if not isinstance(expected_sha, str) or HEX64.fullmatch(expected_sha) is None:
        fail(f"{label}.sha256", "malformed source manifest SHA")
    path = bundle_path(path_value, f"{label}.path")
    if not path.is_file() or sha(path) != expected_sha.lower():
        fail(label, "retained source manifest hash is stale")
    content = obj(load(path, str(path)), str(path))
    if len(content) != files:
        fail(label, "source manifest file count differs from descriptor")
    for name, file_sha in content.items():
        if not isinstance(name, str) or Path(name).is_absolute() or not isinstance(file_sha, str) or HEX64.fullmatch(file_sha) is None:
            fail(str(path), "malformed source manifest entry")
        bundle_path(name, f"{path}.{name}")
        # The manifest is the immutable build binding.  The current checkout
        # may intentionally contain the other side of this before/after
        # comparison, and a portable post-cleanup export may contain no
        # checkout at all, so do not substitute current files for that binding.


def check_build(role: str, protocol_sha: str, require_binaries: bool) -> tuple[Path, dict[str, Any]]:
    directory = ROOT / BUILD_DIR[role]
    path = directory / "build.json"
    build = obj(load(path, str(path)), str(path))
    if build.get("change") != 436:
        fail(str(path), "wrong change")
    if build.get("role") not in {None, directory.name, role}:
        fail(str(path), "wrong role")
    revision = build.get("revision")
    if not isinstance(revision, str) or HEX40.fullmatch(revision) is None:
        fail(str(path), "malformed executable source revision")
    if build.get("protocol_sha256") != protocol_sha:
        fail(str(path), "protocol binding is stale")
    outer_verifier_sha = build.get("outer_verifier_sha256", build.get("verifier_sha256"))
    if outer_verifier_sha != sha(ROOT / "verify.py") and build.get("verifier_sha256") != sha(ROOT / "oracle" / "verify-report.py"):
        fail(str(path), "outer verifier binding is stale")
    if build.get("capture_driver_sha256") != sha(ROOT / "capture.py"):
        fail(str(path), "capture driver binding is stale")
    protocol = obj(load(ROOT / "protocol.json"), "protocol")
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    if build.get("oracle_protocol_sha256", oracle["sha256"]) != oracle["sha256"] or build.get("oracle_verifier_sha256", oracle["verifier_sha256"]) != oracle["verifier_sha256"]:
        fail(str(path), "oracle binding is stale")
    check_source_manifest(obj(build.get("source_manifest"), f"{path}.source_manifest"), f"{path}.source_manifest")
    build_receipt_path = bundle_path(build.get("build_receipt"), f"{path}.build_receipt")
    if build.get("build_receipt_sha256") != sha(build_receipt_path):
        fail(f"{path}.build_receipt_sha256", "retained build receipt hash is stale")
    build_receipt = obj(load(build_receipt_path, str(build_receipt_path)), str(build_receipt_path))
    if (build_receipt.get("change") != 436 or build_receipt.get("status") != "pass" or
            build_receipt.get("source_unchanged") is not True or
            build_receipt.get("revision") != build.get("revision") or
            build_receipt.get("source_before") != build_receipt.get("source_after") or
            build_receipt.get("source_before") != build.get("source_manifest")):
        fail(str(path), "retained build receipt does not bind the descriptor")
    binary_copies_receipt_path = bundle_path(build.get("binary_copies_receipt"), f"{path}.binary_copies_receipt")
    if build.get("binary_copies_receipt_sha256") != sha(binary_copies_receipt_path):
        fail(f"{path}.binary_copies_receipt_sha256", "retained binary-copy receipt hash is stale")
    binaries = obj(build.get("binaries"), f"{path}.binaries")
    if set(binaries) != {"normal", "allocator"}:
        fail(str(path), "normal and allocator identities are required")
    copies_path = binary_copies_receipt_path
    copies = obj(load(copies_path, str(copies_path)), str(copies_path))
    if copies != binaries:
        fail(str(path), "binary descriptor differs from the retained binary-copy receipt")
    for mode in ("normal", "allocator"):
        identity = obj(binaries[mode], f"{path}.binaries.{mode}")
        binary_path = Path(identity.get("path", ""))
        if not binary_path.is_absolute():
            binary_path = REPO / binary_path
        if not isinstance(identity.get("sha256"), str) or HEX64.fullmatch(identity["sha256"]) is None or not isinstance(identity.get("bytes"), int) or identity["bytes"] <= 0:
            fail(f"{path}.binaries.{mode}", "malformed binary identity")
        # A portable post-cleanup replay must not probe the original checkout
        # just because the descriptor retains an absolute build path.  Binary
        # bytes are required only for the explicit pre-cleanup gate.
        if require_binaries:
            if not binary_path.is_file():
                fail(f"{path}.binaries.{mode}", "binary is missing before cleanup")
            if binary_path.stat().st_size != identity["bytes"] or sha(binary_path) != identity["sha256"].lower():
                fail(f"{path}.binaries.{mode}", "binary hash/size differs from descriptor")
    return directory, build


def artifact_path(receipt_path: Path, value: Any, label: str) -> Path:
    record = obj(value, label)
    relative = record.get("path")
    path = bundle_path(relative, f"{label}.path")
    selected = path
    if not selected.is_file():
        compressed = Path(str(path) + ".gz")
        if compressed.is_file():
            selected = compressed
        else:
            fail(label, "artifact is missing in raw and gzip form")
    try:
        raw = gzip.decompress(selected.read_bytes()) if selected.name.endswith(".gz") else selected.read_bytes()
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(label, f"invalid compressed artifact: {error}")
    if not isinstance(record.get("bytes"), int) or isinstance(record["bytes"], bool) or record["bytes"] < 0 or record["bytes"] != len(raw):
        fail(label, "artifact byte count is stale")
    if record.get("sha256") != sha_bytes(raw):
        fail(label, "artifact hash is stale")
    return selected


def artifact_bytes(path: Path, label: str) -> bytes:
    try:
        raw = path.read_bytes()
        return gzip.decompress(raw) if path.name.endswith(".gz") else raw
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(label, f"cannot read retained artifact: {error}")
    raise AssertionError("unreachable")


def check_argv_path(value: Any, artifact: Path, label: str) -> None:
    if not isinstance(value, str) or not Path(value).is_absolute():
        fail(label, "argv path must be absolute")
    actual = Path(value).parts
    logical = Path(str(artifact)[:-3]) if artifact.name.endswith(".gz") else artifact
    expected = logical.relative_to(ROOT).parts
    if len(actual) < len(expected) or actual[-len(expected):] != expected:
        fail(label, "argv path does not bind the retained artifact")


def check_binary_argv(value: Any, binary: dict[str, Any], label: str) -> None:
    if not isinstance(value, str) or not Path(value).is_absolute():
        fail(label, "binary argv path must be absolute")
    actual = Path(value)
    descriptor = Path(binary["path"])
    if descriptor.is_absolute():
        if actual != descriptor:
            fail(label, "binary argv path differs from the build descriptor")
    elif actual.name != descriptor.name:
        fail(label, "binary argv basename differs from the build descriptor")


def check_capture_argv(
    receipt: dict[str, Any],
    protocol: dict[str, Any],
    lane: dict[str, Any],
    binary: dict[str, Any],
    artifacts: dict[str, Path],
    label: str,
) -> None:
    argv = array(receipt.get("argv"), f"{label}.argv")
    if len(argv) != 22 or any(not isinstance(value, str) or not value for value in argv):
        fail(f"{label}.argv", "capture argv does not match the frozen command shape")
    fixed = {
        0: "taskset", 1: "-c", 2: str(protocol["cpu"]), 3: "/usr/bin/time", 4: "-v", 5: "-o",
        8: "--case", 9: protocol["roles"][receipt["role"]]["selector"], 10: "--semantic-shape", 11: lane["shape"],
        12: "--workers", 13: str(protocol["workers"]), 14: "--samples", 15: str(protocol["samples"]),
        16: "--warmup", 17: str(protocol["warmups"]), 18: "--json", 20: "--corpus-manifest",
    }
    for index, expected in fixed.items():
        if argv[index] != expected:
            fail(f"{label}.argv[{index}]", f"must be {expected!r}")
    check_argv_path(argv[6], artifacts["resource_log"], f"{label}.argv[6]")
    check_binary_argv(argv[7], binary, f"{label}.argv[7]")
    check_argv_path(argv[19], artifacts["report"], f"{label}.argv[19]")
    check_argv_path(argv[21], artifacts["catalog"], f"{label}.argv[21]")


def check_profile_argv(
    receipt: dict[str, Any],
    protocol: dict[str, Any],
    binary: dict[str, Any],
    artifacts: dict[str, Path],
    kind: str,
    label: str,
) -> None:
    argv = array(receipt.get("argv"), f"{label}.argv")
    expected_length = 31 if kind == "stat" else 34
    if len(argv) != expected_length or any(not isinstance(value, str) or not value for value in argv):
        fail(f"{label}.argv", "profile argv contains malformed arguments")
    prefix = ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o"]
    if argv[:6] != prefix:
        fail(f"{label}.argv", "profile argv prefix differs from protocol")
    check_argv_path(argv[6], artifacts["resource"], f"{label}.argv[6]")
    if kind == "stat":
        expected = ["perf", "stat", "--no-big-num", "-x,", "-e", ",".join(PROFILE_EVENTS), "-o"]
        if argv[7:14] != expected:
            fail(f"{label}.argv", "perf-stat argv differs from the frozen event command")
        check_argv_path(argv[14], artifacts["perf_stat"], f"{label}.argv[14]")
        if argv[15] != "--":
            fail(f"{label}.argv[15]", "perf-stat must terminate its profiler options with --")
        tail = 16
    else:
        expected = ["perf", "record", "--no-buildid-cache", "-o"]
        if argv[7:11] != expected:
            fail(f"{label}.argv", "perf-record argv prefix differs from the frozen command")
        check_argv_path(argv[11], artifacts["perf_data"], f"{label}.argv[11]")
        if argv[12:19] != ["-F", "999", "-e", "cycles:u", "--call-graph", "fp,127", "--"]:
            fail(f"{label}.argv", "perf-record sampling/call-graph argv differs")
        tail = 19
    base = ["--case", receipt["selector"], "--semantic-shape", "large", "--workers", str(protocol["workers"]), "--samples", str(protocol["samples"]), "--warmup", str(protocol["warmups"]), "--json"]
    if argv[tail + 1:tail + 1 + len(base)] != base:
        fail(f"{label}.argv", "profile workload arguments differ from the protocol")
    check_binary_argv(argv[tail], binary, f"{label}.argv[{tail}]")
    report_index = tail + 1 + len(base)
    check_argv_path(argv[report_index], artifacts["report"], f"{label}.argv[{report_index}]")
    if len(argv) != report_index + 3 or argv[report_index + 1] != "--corpus-manifest":
        fail(f"{label}.argv", "profile corpus-manifest arguments are malformed")
    check_argv_path(argv[report_index + 2], artifacts["catalog"], f"{label}.argv[{report_index + 2}]")


def check_capture_state(
    state: dict[str, Any],
    phase: str,
    role: str,
    build: dict[str, Any],
    ambient_build: dict[str, Any],
    protocol_sha: str,
    indexed: list[Any],
    label: str,
) -> None:
    if state.get("schema") != "litchi-0436-capture-state-v1" or state.get("change") != 436:
        fail(label, "capture-state schema or change differs")
    if state.get("phase") != phase or state.get("attempt") != "formal" or state.get("role") != role:
        fail(label, "capture-state phase, role, or attempt differs")
    if state.get("expected_lanes") != 6 or state.get("completed_lanes") != 6 or state.get("status") != "pass":
        fail(label, "capture-state does not prove all six formal lanes passed")
    if state.get("source_manifest") != build["source_manifest"] or state.get("build_source_manifest") != build["source_manifest"]:
        fail(label, "capture-state build source binding differs")
    if state.get("ambient_source_manifest") != ambient_build["source_manifest"]:
        fail(label, "capture-state ambient source binding differs")
    if state.get("source_before") != state.get("source_after") or state.get("source_before") != ambient_build["source_manifest"]:
        fail(label, "capture-state source custody differs from ambient build")
    for key in ("status_before", "status_after"):
        statuses = state.get(key)
        if not isinstance(statuses, list) or any(not isinstance(value, str) for value in statuses):
            fail(f"{label}.{key}", "outside-bundle custody status must be a string list")
    if state["status_before"] != state["status_after"]:
        fail(label, "outside-bundle custody changed during phase")
    if state.get("protocol_sha256") != protocol_sha:
        fail(label, "capture-state protocol hash is stale")
    if state.get("oracle_protocol_sha256") != obj(load(ROOT / "protocol.json").get("oracle"), "protocol.oracle")["sha256"] or state.get("oracle_verifier_sha256") != obj(load(ROOT / "protocol.json").get("oracle"), "protocol.oracle")["verifier_sha256"]:
        fail(label, "capture-state oracle hashes are stale")
    if state.get("driver_sha256") != sha(ROOT / "capture.py"):
        fail(label, "capture-state driver hash is stale")
    if state.get("index") != indexed:
        fail(label, "capture-state index differs from capture-index.json")
    timestamp(state.get("finished_utc"), f"{label}.finished_utc")


def check_capture_receipt(
    receipt: dict[str, Any],
    phase: str,
    role: str,
    lane: dict[str, Any],
    name: str,
    build: dict[str, Any],
    ambient_build: dict[str, Any],
    protocol: dict[str, Any],
    protocol_sha: str,
    state: dict[str, Any],
    label: str,
) -> tuple[dt.datetime, dt.datetime]:
    if receipt.get("schema") != "litchi-0436-capture-receipt-v1" or receipt.get("change") != 436:
        fail(label, "capture receipt schema or change differs")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
        fail(label, "lane is not a passing workload/oracle invocation")
    if receipt.get("phase") != phase or receipt.get("attempt") != "formal" or receipt.get("role") != role or receipt.get("build_directory") != BUILD_DIR[role]:
        fail(label, "receipt phase, attempt, role, or build directory differs")
    if receipt.get("selector") != protocol["roles"][role]["selector"] or receipt.get("name") != name or receipt.get("lane") != lane:
        fail(label, "receipt does not bind the frozen lane")
    if receipt.get("revision") != build.get("revision") or receipt.get("source_manifest") != build.get("source_manifest") or receipt.get("binary") != build["binaries"][lane["mode"]]:
        fail(label, "receipt executable source binding differs from role build")
    if receipt.get("ambient_source_manifest") != ambient_build["source_manifest"]:
        fail(label, "receipt ambient source binding differs from after build")
    if receipt.get("protocol_sha256") != protocol_sha or receipt.get("oracle_protocol_sha256") != protocol["oracle"]["sha256"] or receipt.get("oracle_verifier_sha256") != protocol["oracle"]["verifier_sha256"]:
        fail(label, "receipt protocol/oracle binding is stale")
    if receipt.get("driver_sha256") != sha(ROOT / "capture.py") or receipt.get("verifier_sha256") != sha(ROOT / "verify.py"):
        fail(label, "receipt capture/verifier driver binding is stale")
    if receipt.get("source_before") != state.get("source_before") or receipt.get("source_after") != state.get("source_after"):
        fail(label, "receipt source custody differs from phase custody")
    if receipt.get("status_before") != state.get("status_before") or receipt.get("status_after") != state.get("status_after"):
        fail(label, "receipt outside-bundle custody differs from phase custody")
    if receipt.get("source_unchanged") is not True or receipt.get("outside_bundle_status_unchanged") is not True:
        fail(label, "receipt does not prove unchanged custody")
    cwd = receipt.get("cwd")
    if not isinstance(cwd, str) or not cwd or not Path(cwd).is_absolute():
        fail(f"{label}.cwd", "capture cwd must be an absolute path")
    started = timestamp(receipt.get("started_utc"), f"{label}.started_utc")
    finished = timestamp(receipt.get("finished_utc"), f"{label}.finished_utc")
    if finished < started:
        fail(label, "lane finished before it started")
    return started, finished


def read_resource(path: Path, label: str) -> dict[str, Any]:
    selected = path
    if not selected.is_file() and path.with_suffix(path.suffix + ".gz").is_file():
        selected = path.with_suffix(path.suffix + ".gz")
    try:
        if selected.suffix == ".gz":
            text = gzip.open(selected, "rt", encoding="utf-8").read()
        else:
            text = selected.read_text(encoding="utf-8")
    except OSError as error:
        fail(label, f"cannot read resource log: {error}")
    values = [int(match.group(1)) for line in text.splitlines() if (match := RSS_RE.match(line.lstrip()))]
    if len(values) != 1:
        fail(label, "must contain exactly one GNU time RSS line")
    return {"status": "measured", "scope": "gnu_time_v_verbose_whole_fresh_process", "kib": values[0], "bytes": values[0] * 1024, "path": str(selected.relative_to(ROOT))}


def verify_frozen_inputs() -> None:
    path = ROOT / "frozen-inputs.json"
    frozen = obj(load(path, str(path)), str(path))
    if frozen.get("status") != "frozen":
        fail(str(path), "frozen-inputs status is not frozen")
    files = obj(frozen.get("files"), f"{path}.files")
    expected = {
        "protocol.json",
        "oracle/verify-report.py",
        "oracle/protocol.json",
        "check.py",
    }
    if set(files) != expected:
        fail(f"{path}.files", "does not enumerate the frozen protocol, oracle, and custody driver")
    for name in sorted(expected):
        target = bundle_path(name, f"{path}.files.{name}")
        if digest(files.get(name), f"{path}.files.{name}") != sha(target):
            fail(f"{path}.files.{name}", "frozen input digest differs")


def verify_compression() -> int:
    path = ROOT / "compression.json"
    if not path.is_file():
        return 0
    records = obj(load(path, str(path)), str(path))
    seen: set[str] = set()
    for stored_name, value in records.items():
        if not isinstance(stored_name, str) or not stored_name.endswith(".gz"):
            fail(f"{path}.records", "stored paths must end in .gz")
        if stored_name in seen:
            fail(str(path), f"duplicate stored path {stored_name}")
        seen.add(stored_name)
        row = obj(value, f"{path}.{stored_name}")
        if set(row) != {"original_path", "original_sha256", "original_bytes", "stored_sha256", "stored_bytes"}:
            fail(f"{path}.{stored_name}", "compression record fields differ")
        original_name = row["original_path"]
        if not isinstance(original_name, str) or stored_name[:-3] != original_name:
            fail(f"{path}.{stored_name}.original_path", "does not match the stored .gz path")
        stored = bundle_path(stored_name, f"{path}.{stored_name}")
        if not stored.is_file():
            fail(f"{path}.{stored_name}", "stored compressed artifact is missing")
        stored_raw = stored.read_bytes()
        if len(stored_raw) != u64(row["stored_bytes"], f"{path}.{stored_name}.stored_bytes") or sha_bytes(stored_raw) != digest(row["stored_sha256"], f"{path}.{stored_name}.stored_sha256"):
            fail(f"{path}.{stored_name}", "stored compressed artifact binding differs")
        try:
            original_raw = gzip.decompress(stored_raw)
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            fail(f"{path}.{stored_name}", f"invalid gzip stream: {error}")
        original = bundle_path(original_name, f"{path}.{stored_name}.original_path")
        if original.is_file() and original.read_bytes() != original_raw:
            fail(f"{path}.{stored_name}", "raw and compressed artifact bytes differ")
        if len(original_raw) != u64(row["original_bytes"], f"{path}.{stored_name}.original_bytes") or sha_bytes(original_raw) != digest(row["original_sha256"], f"{path}.{stored_name}.original_sha256"):
            fail(f"{path}.{stored_name}", "original artifact binding differs")
    actual = {str(value.relative_to(ROOT)) for value in ROOT.rglob("*.gz") if value.is_file()}
    if actual != seen:
        fail(str(path), "compressed artifact inventory differs from compression.json")
    return len(seen)


def verify_inventory(required: bool) -> int:
    path = ROOT / "SHA256SUMS"
    if not path.is_file():
        if required:
            fail("SHA256SUMS", "sealed inventory is required")
        return 0
    rows: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail("SHA256SUMS", f"cannot read inventory: {error}")
    for line in lines:
        if "  " not in line:
            fail("SHA256SUMS", "malformed inventory line")
        value, name = line.split("  ", 1)
        if name == "SHA256SUMS" or name in rows:
            fail("SHA256SUMS", "duplicate or self-referential inventory entry")
        rows[name] = digest(value, f"SHA256SUMS.{name}")
        target = bundle_path(name, f"SHA256SUMS.{name}")
        if not target.is_file() or sha(target) != rows[name]:
            fail(f"SHA256SUMS.{name}", "inventory digest differs")
    expected = {str(value.relative_to(ROOT)) for value in ROOT.rglob("*") if value.is_file() and value.name != "SHA256SUMS"}
    if set(rows) != expected:
        fail("SHA256SUMS", "inventory does not cover the bundle exactly")
    return len(rows)


def verify_cleanup(stage: str) -> None:
    if stage not in STAGES:
        fail("stage", f"must be one of {STAGES!r}")
    if stage == "precleanup":
        return
    path = ROOT / "checks" / "cleanup-inventory.json"
    inventory = obj(load(path, str(path)), str(path))
    if inventory.get("status") != "pass":
        fail(str(path), "cleanup inventory does not record a passing cleanup")
    removed = array(inventory.get("removed"), f"{path}.removed")
    expected_paths = {
        "/tmp/litchi-goal-0436-binaries",
        "/tmp/litchi-goal-0436-odt",
        "/tmp/litchi-goal-0436-odt-tests",
        "/tmp/litchi-goal-0436-odt-harness",
        "/tmp/litchi-goal-0436-common-tests",
    }
    seen: set[str] = set()
    for index, value in enumerate(removed):
        row = obj(value, f"{path}.removed[{index}]")
        name = row.get("path")
        if not isinstance(name, str) or name not in expected_paths or name in seen:
            fail(f"{path}.removed[{index}].path", "removed directory is not the frozen 0436 scratch paths")
        seen.add(name)
        u64(row.get("regular_files"), f"{path}.removed[{index}].regular_files")
        u64(row.get("regular_file_bytes"), f"{path}.removed[{index}].regular_file_bytes")
    if seen != expected_paths:
        fail(str(path), "cleanup inventory does not enumerate all five scratch directories")
    preserved = obj(inventory.get("preserved_target_directory_identity"), f"{path}.preserved_target_directory_identity")
    if not preserved:
        fail(str(path), "cleanup inventory has no preserved target directory identities")
    for name, identity in preserved.items():
        if not isinstance(name, str) or not name or not isinstance(identity, list) or len(identity) != 2:
            fail(f"{path}.preserved_target_directory_identity", "preserved directory identity is malformed")
        u64(identity[0], f"{path}.preserved_target_directory_identity.{name}[0]")
        u64(identity[1], f"{path}.preserved_target_directory_identity.{name}[1]")
    if inventory.get("user_goal_sha256") != "bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1":
        fail(str(path), "user goal hash differs from the pinned original")


def report_identity(verified: dict[str, Any], label: str) -> dict[str, Any]:
    report = obj(verified["report"], f"{label}.report")
    result = obj(verified["result"], f"{label}.result")
    corpus = obj(result["corpus"], f"{label}.corpus")
    source = obj(result["source"], f"{label}.source")
    paragraphs = obj(source.get("odt_paragraphs"), f"{label}.source.odt_paragraphs")
    sink = obj(result.get("sink"), f"{label}.sink")
    return {
        "archive_bytes": u64(corpus.get("archive_bytes"), f"{label}.archive_bytes"),
        "archive_sha256": digest(corpus.get("archive_sha256"), f"{label}.archive_sha256"),
        "target_payload_bytes": u64(corpus.get("target_payload_bytes"), f"{label}.target_payload_bytes"),
        "target_payload_sha256": digest(corpus.get("target_payload_sha256"), f"{label}.target_payload_sha256"),
        "output_sha256": digest(result.get("output_sha256"), f"{label}.output_sha256"),
        "semantic_sha256": digest(paragraphs.get("semantic_sha256"), f"{label}.semantic_sha256"),
        "content_xml_sha256": digest(paragraphs.get("content_xml_sha256"), f"{label}.content_xml_sha256"),
        "styles_xml_sha256": digest(paragraphs.get("styles_xml_sha256"), f"{label}.styles_xml_sha256"),
        "meta_xml_sha256": digest(paragraphs.get("meta_xml_sha256"), f"{label}.meta_xml_sha256"),
        "paragraph_count": u64(paragraphs.get("paragraph_count"), f"{label}.paragraph_count"),
        "member_names": list(ODT_MEMBERS),
        "mimetype": ODT_MIMETYPE,
        "manifest_entry_count": 4,
        "sink_accepted_bytes": u64(sink.get("accepted_bytes"), f"{label}.sink.accepted_bytes"),
        "sink_write_calls": u64(sink.get("write_calls"), f"{label}.sink.write_calls"),
        "environment_revision": report["environment"].get("git_revision"),
        "environment_dirty": report["environment"].get("git_worktree_dirty"),
    }


def verify_profiles(protocol: dict[str, Any], builds: dict[str, tuple[Path, dict[str, Any]]], ambient: set[tuple[str, bool]], oracle: Any) -> list[dict[str, Any]]:
    """Validate the four required large-normal profile receipts/artifacts."""

    rows: list[dict[str, Any]] = []
    expected_receipts: set[Path] = set()
    protocol_sha = sha(ROOT / "protocol.json")
    ambient_source = builds["after-streaming"][1]["source_manifest"]
    oracle_spec = obj(protocol.get("oracle"), "protocol.oracle")
    for role in PROFILE_ROLES:
        for kind in PROFILE_KINDS:
            directory = ROOT / "profiles" / BUILD_DIR[role] / kind
            receipt_path = directory / "receipt.json"
            expected_receipts.add(receipt_path)
            receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
            label = str(receipt_path)
            if receipt.get("schema") != "litchi-0436-profile-receipt-v1" or receipt.get("change") != 436:
                fail(label, "profile schema or change differs")
            if receipt.get("status") != "pass" or receipt.get("role") != role or receipt.get("kind") != kind:
                fail(label, "required profile is not passing")
            if receipt.get("attempt") != "formal" or receipt.get("preparatory") is not False:
                fail(label, "formal profile must be explicitly non-preparatory")
            if receipt.get("selector") != protocol["roles"][role]["selector"] or receipt.get("shape") != "large":
                fail(label, "profile case identity differs from the frozen protocol")
            if receipt.get("protocol_path") != "protocol.json" or receipt.get("protocol_sha256") != protocol_sha:
                fail(label, "profile protocol binding is stale")
            if receipt.get("source_manifest") != builds[role][1]["source_manifest"] or receipt.get("ambient_source_manifest") != ambient_source:
                fail(label, "profile source/ambient binding is stale")
            if receipt.get("ambient_build_directory") != "after":
                fail(label, "formal profiles must use the after candidate as ambient checkout")
            if receipt.get("source_before") != receipt.get("source_after") or receipt.get("source_unchanged") is not True:
                fail(label, "profile source custody changed")
            if receipt.get("driver_sha256") != sha(ROOT / "profile.py"):
                fail(label, "profile driver hash is stale")
            if receipt.get("oracle_verifier_path") != "oracle/verify-report.py" or receipt.get("oracle_verifier_sha256") != oracle_spec["verifier_sha256"]:
                fail(label, "profile oracle verifier binding is stale")
            if receipt.get("oracle_role") != oracle_role(protocol, role):
                fail(label, "profile oracle role differs from the protocol")
            started = timestamp(receipt.get("started_utc"), f"{label}.started_utc")
            finished = timestamp(receipt.get("finished_utc"), f"{label}.finished_utc")
            if finished < started:
                fail(label, "profile finished before it started")
            if receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
                fail(label, "profile workload or oracle did not pass")
            if kind == "stat":
                if receipt.get("stat_events") != PROFILE_EVENTS:
                    fail(label, "perf-stat event set is not the frozen user-mode set")
            elif receipt.get("record_event") != "cycles:u" or receipt.get("record_frequency_hz") != 999 or receipt.get("call_graph") != "fp,127":
                fail(label, "perf-record sampling/call-graph binding is stale")
            if receipt.get("binary") != builds[role][1]["binaries"]["normal"]:
                fail(label, "profile binary binding differs from role build")
            artifacts = obj(receipt.get("artifacts"), f"{label}.artifacts")
            expected_artifacts = {"report", "catalog", "resource", "workload_log", "oracle_log"}
            expected_artifacts.add("perf_stat" if kind == "stat" else "perf_data")
            if kind == "record":
                expected_artifacts.update({"perf_script", "perf_report"})
            if set(artifacts) != expected_artifacts:
                fail(f"{label}.artifacts", "profile artifact inventory differs")
            artifact_paths = {
                key: artifact_path(receipt_path, artifacts.get(key), f"{label}.artifacts.{key}")
                for key in expected_artifacts
            }
            check_profile_argv(receipt, protocol, builds[role][1]["binaries"]["normal"], artifact_paths, kind, label)
            oracle_log = artifact_bytes(artifact_paths["oracle_log"], f"{label}.artifacts.oracle_log")
            if not oracle_log.startswith(b"stdout=VALID\n") or b"stderr=" not in oracle_log:
                fail(f"{label}.artifacts.oracle_log", "profile oracle log does not retain VALID output")
            if kind == "stat":
                if not artifact_bytes(artifact_paths["perf_stat"], f"{label}.artifacts.perf_stat"):
                    fail(f"{label}.artifacts.perf_stat", "perf stat output is empty")
            else:
                for key in ("perf_data", "perf_script", "perf_report"):
                    if not artifact_bytes(artifact_paths[key], f"{label}.artifacts.{key}"):
                        fail(f"{label}.artifacts.{key}", "retained perf artifact is empty")
            report_path = artifact_paths["report"]
            verified = oracle.validate_report(report_path, "normal", "large", oracle_role(protocol, role))
            report = verified["report"]
            binary = builds[role][1]["binaries"]["normal"]
            binary_report = obj(report.get("binary_identity"), f"{report_path}.binary_identity")
            if binary_report.get("path") != binary["path"] or binary_report.get("binary_sha256") != binary["sha256"] or binary_report.get("binary_bytes") != binary["bytes"]:
                fail(str(report_path), "profile report binary identity differs from build")
            env = obj(report.get("environment"), f"{report_path}.environment")
            identity = (env.get("git_revision"), env.get("git_worktree_dirty"))
            if identity not in ambient:
                fail(str(report_path), "profile ambient checkout differs from formal matrix")
            rows.append({"role": role, "kind": kind, "receipt": receipt, "report": str(report_path.relative_to(ROOT)), "resource": read_resource(artifact_paths["resource"], f"{role}.{kind}.resource")})
    profile_root = ROOT / "profiles"
    for receipt_path in (profile_root.rglob("receipt.json") if profile_root.is_dir() else ()):
        if receipt_path not in expected_receipts:
            receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
            if receipt.get("status") not in {"pass", "failed"}:
                fail(str(receipt_path), "retained extra profile receipt is not terminal")
            if receipt.get("source_before") != receipt.get("source_after") or receipt.get("source_unchanged") is not True:
                fail(str(receipt_path), "retained extra profile receipt does not prove source custody")
    return rows


def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    order = array(protocol["order"], "protocol.order")
    _, first, last, repeat = PHASES[phase]
    lanes = [obj(value, f"protocol.order[{i}]") for i, value in enumerate(order[first:last], first)]
    if any(lane.get("repeat") != repeat for lane in lanes):
        fail(f"protocol.phase.{phase}", "lane repeat does not match ABBA phase")
    return lanes


def verify_extra_attempts(
    protocol: dict[str, Any],
    protocol_sha: str,
    builds: dict[str, tuple[Path, dict[str, Any]]],
) -> list[dict[str, Any]]:
    """Validate retained nonformal attempts instead of silently ignoring them."""

    runs_root = ROOT / "runs"
    if not runs_root.is_dir():
        return []
    extras: list[dict[str, Any]] = []
    for state_path in sorted(runs_root.glob("*/ */capture-state.json".replace(" ", ""))):
        if state_path.parent.name == "formal":
            continue
        state = obj(load(state_path, str(state_path)), str(state_path))
        phase = state.get("phase")
        if phase not in PHASES:
            fail(str(state_path), "retained attempt has an unknown phase")
        role = PHASES[phase][0]
        attempt = state.get("attempt")
        if not isinstance(attempt, str) or not attempt or attempt == "formal":
            fail(str(state_path), "retained attempt tag is malformed")
        if state.get("change") != 436 or state.get("schema") != "litchi-0436-capture-state-v1" or state.get("role") != role:
            fail(str(state_path), "retained attempt state identity differs")
        if state.get("status") not in {"pass", "failed"}:
            fail(str(state_path), "retained attempt is not terminal")
        if state.get("source_before") != state.get("source_after") or state.get("status_before") != state.get("status_after"):
            fail(str(state_path), "retained attempt changed source or outside-bundle custody")
        if state.get("source_manifest") != builds[role][1]["source_manifest"] or state.get("ambient_source_manifest") != builds["after-streaming"][1]["source_manifest"]:
            fail(str(state_path), "retained attempt source binding differs")
        if state.get("protocol_sha256") != protocol_sha or state.get("driver_sha256") != sha(ROOT / "capture.py"):
            fail(str(state_path), "retained attempt driver binding is stale")
        index = array(state.get("index"), f"{state_path}.index")
        if state.get("completed_lanes") != len(index) or len(index) > 6:
            fail(str(state_path), "retained attempt completion count differs from its index")
        lanes = expected_lanes(protocol, phase)
        for position, receipt_name in enumerate(index):
            if not isinstance(receipt_name, str) or Path(receipt_name).is_absolute() or not receipt_name.endswith("-receipt.json"):
                fail(f"{state_path}.index[{position}]", "retained receipt path is malformed")
            expected_path = state_path.parent / Path(receipt_name).name
            receipt = obj(load(expected_path, str(expected_path)), str(expected_path))
            if receipt.get("phase") != phase or receipt.get("attempt") != attempt or receipt.get("role") != role or receipt.get("lane") != lanes[position]:
                fail(str(expected_path), "retained attempt receipt does not bind its phase lane")
            if receipt.get("status") not in {"pass", "failed"}:
                fail(str(expected_path), "retained attempt receipt is not terminal")
            if receipt.get("source_before") != receipt.get("source_after") or receipt.get("source_unchanged") is not True:
                fail(str(expected_path), "retained attempt receipt does not prove source custody")
            if receipt.get("protocol_sha256") != protocol_sha or receipt.get("driver_sha256") != sha(ROOT / "capture.py") or receipt.get("verifier_sha256") != sha(ROOT / "verify.py"):
                fail(str(expected_path), "retained attempt receipt driver binding is stale")
            started = timestamp(receipt.get("started_utc"), f"{expected_path}.started_utc")
            finished = timestamp(receipt.get("finished_utc"), f"{expected_path}.finished_utc")
            if finished < started:
                fail(str(expected_path), "retained attempt receipt finished before it started")
            artifacts = obj(receipt.get("artifacts"), f"{expected_path}.artifacts")
            for artifact_name, artifact in artifacts.items():
                artifact_path(expected_path, artifact, f"{expected_path}.artifacts.{artifact_name}")
            if receipt["status"] == "pass" and (receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0):
                fail(str(expected_path), "passing retained attempt has a nonzero workload/oracle result")
        if state.get("status") == "pass" and len(index) != 6:
            fail(str(state_path), "passing retained attempt does not contain all six lanes")
        extras.append({"path": str(state_path.relative_to(ROOT)), "status": state["status"], "phase": phase, "attempt": attempt})
    return extras


def verify_matrix(*, require_binaries: bool = False) -> dict[str, Any]:
    protocol_path = ROOT / "protocol.json"
    verify_frozen_inputs()
    protocol = obj(load(protocol_path, str(protocol_path)), "protocol")
    validate_protocol(protocol)
    protocol_sha = sha(protocol_path)
    oracle = load_oracle()
    builds: dict[str, tuple[Path, dict[str, Any]]] = {}
    for role in PROFILE_ROLES:
        builds[role] = check_build(role, protocol_sha, require_binaries)
    if builds["before-streaming"][1]["revision"] == builds["after-streaming"][1]["revision"]:
        fail("builds", "before and after executable revisions must remain distinct")
    if builds["before-streaming"][1]["source_manifest"] == builds["after-streaming"][1]["source_manifest"]:
        fail("builds", "before and after source manifests must remain distinct")
    rows: list[dict[str, Any]] = []
    formal_dirs: dict[str, Path] = {}
    phase_bounds: dict[str, tuple[dt.datetime, dt.datetime]] = {}
    for phase, (role, _, _, _) in PHASES.items():
        directory = ROOT / "runs" / phase / "formal"
        formal_dirs[phase] = directory
        state_path = directory / "capture-state.json"
        state = obj(load(state_path, str(state_path)), str(state_path))
        index_path = directory / "capture-index.json"
        indexed = array(load(index_path, str(index_path)), str(index_path))
        lanes = expected_lanes(protocol, phase)
        if len(indexed) != len(lanes):
            fail(str(index_path), "does not list exactly the phase lanes")
        check_capture_state(
            state,
            phase,
            role,
            builds[role][1],
            builds["after-streaming"][1],
            protocol_sha,
            indexed,
            str(state_path),
        )
        build_dir, build = builds[role]
        previous_finished: dt.datetime | None = None
        phase_started: dt.datetime | None = None
        phase_finished: dt.datetime | None = None
        for position, lane in enumerate(lanes):
            mode, shape, repeat = lane["mode"], lane["shape"], lane["repeat"]
            name = f"{phase}-{mode}-{shape}-{str(repeat).lower()}"
            expected_receipt = directory / f"{name}-receipt.json"
            if indexed[position] != str(expected_receipt.relative_to(ROOT)):
                fail(f"{index_path}[{position}]", "receipt ordering/path differs from protocol")
            receipt = obj(load(expected_receipt, str(expected_receipt)), str(expected_receipt))
            artifacts = obj(receipt.get("artifacts"), f"{expected_receipt}.artifacts")
            if set(artifacts) != {"report", "catalog", "workload_log", "resource_log", "oracle_log"}:
                fail(f"{expected_receipt}.artifacts", "capture artifact inventory differs")
            artifact_paths = {
                key: artifact_path(expected_receipt, artifacts.get(key), f"{expected_receipt}.artifacts.{key}")
                for key in artifacts
            }
            started, finished = check_capture_receipt(
                receipt,
                phase,
                role,
                lane,
                name,
                build,
                builds["after-streaming"][1],
                protocol,
                protocol_sha,
                state,
                str(expected_receipt),
            )
            if previous_finished is not None and started < previous_finished:
                fail(str(expected_receipt), "formal lanes overlap or are out of capture order")
            previous_finished = finished
            phase_started = started if phase_started is None else min(phase_started, started)
            phase_finished = finished if phase_finished is None else max(phase_finished, finished)
            check_capture_argv(receipt, protocol, lane, build["binaries"][mode], artifact_paths, str(expected_receipt))
            oracle_log = artifact_bytes(artifact_paths["oracle_log"], f"{expected_receipt}.artifacts.oracle_log")
            if b"stdout=VALID" not in oracle_log or b"stderr=" not in oracle_log:
                fail(f"{expected_receipt}.artifacts.oracle_log", "oracle log does not retain VALID output")
            if receipt.get("oracle_stdout_sha256") != sha_bytes(oracle_log):
                fail(str(expected_receipt), "oracle log hash differs from receipt")
            report_path = artifact_paths["report"]
            catalog_path = artifact_paths["catalog"]
            resource_path = artifact_paths["resource_log"]
            # The copied oracle checks the adjacent catalog sidecar itself.
            verified = oracle.validate_report(report_path, mode, shape, oracle_role(protocol, role))
            report = verified["report"]
            binary_report = obj(report.get("binary_identity"), f"{report_path}.binary_identity")
            binary = build["binaries"][mode]
            if binary_report.get("path") != binary["path"] or binary_report.get("binary_sha256") != binary["sha256"] or binary_report.get("binary_bytes") != binary["bytes"]:
                fail(str(report_path), "report binary identity differs from build descriptor")
            env = obj(report.get("environment"), f"{report_path}.environment")
            if not isinstance(env.get("git_revision"), str) or HEX40.fullmatch(env["git_revision"]) is None or not isinstance(env.get("git_worktree_dirty"), bool):
                fail(str(report_path), "runtime checkout identity is malformed")
            rows.append({
                "phase": phase, "role": role, "lane": lane, "name": name, "report_path": str(report_path.relative_to(ROOT)),
                "catalog_path": str(catalog_path.relative_to(ROOT)), "resource": read_resource(resource_path, f"{name}.resource"),
                "receipt": receipt, "build": build, "verified": verified,
                "identity": report_identity(verified, name),
                "started": started, "finished": finished,
            })
        if phase_started is None or phase_finished is None:
            fail(str(state_path), "formal phase has no lane timestamps")
        state_finished = timestamp(state.get("finished_utc"), f"{state_path}.finished_utc")
        if state_finished < phase_finished:
            fail(str(state_path), "phase finished before its final lane")
        phase_bounds[phase] = (phase_started, phase_finished)
    if len(rows) != 24:
        fail("matrix", "formal matrix must contain 24 reports")
    for previous, current in zip(PHASE_ORDER, PHASE_ORDER[1:]):
        if phase_bounds[previous][1] > phase_bounds[current][0]:
            fail("phase.order", f"{previous} overlaps or follows {current} incorrectly")
    ambient = {(row["identity"]["environment_revision"], row["identity"]["environment_dirty"]) for row in rows}
    if len(ambient) != 1:
        fail("ambient_identity", "runtime checkout revision/dirty state differs across formal reports")
    identity_pairs: list[dict[str, Any]] = []
    by_key: dict[tuple[str, str, str], dict[str, dict[str, Any]]] = {}
    for row in rows:
        key = (row["lane"]["mode"], row["lane"]["shape"], row["lane"]["repeat"])
        by_key.setdefault(key, {})[row["role"]] = row
    identity_fields = [
        "archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256",
        "output_sha256", "content_xml_sha256", "semantic_sha256", "styles_xml_sha256",
        "meta_xml_sha256", "paragraph_count", "member_names", "mimetype", "manifest_entry_count",
        "sink_accepted_bytes", "sink_write_calls",
    ]
    cross_fields = list(identity_fields)
    cross_spec = protocol.get("cross_phase_identity")
    if cross_spec is not None:
        cross_spec = obj(cross_spec, "protocol.cross_phase_identity")
        required = cross_spec.get("required_fields")
        if required != cross_fields:
            fail("protocol.cross_phase_identity.required_fields", "must pin exact archive, content, semantic, output, and sink identities")
    for key, pair in sorted(by_key.items()):
        if set(pair) != set(PROFILE_ROLES):
            fail(f"cross_role.{key}", "must contain one report for each ODT role")
        before = pair["before-streaming"]["identity"]
        after = pair["after-streaming"]["identity"]
        mismatches = {
            field: {"before": before[field], "after": after[field]}
            for field in cross_fields if before[field] != after[field]
        }
        if mismatches:
            fail(f"cross_role.{key}", f"exact archive/content/semantic identity mismatch: {mismatches}")
        identity_pairs.append({
            "mode": key[0], "shape": key[1], "repeat": key[2],
            "identity": {field: before[field] for field in cross_fields},
            "status": "pass",
        })
    for role in PROFILE_ROLES:
        for mode in ("normal", "allocator"):
            for shape in ("tiny", "medium", "large"):
                repeat_rows = [row for row in rows if row["role"] == role and row["lane"]["mode"] == mode and row["lane"]["shape"] == shape]
                if len(repeat_rows) != 2:
                    fail(f"repeat.{role}.{mode}.{shape}", "must retain exactly R1 and R2")
                first, second = sorted(repeat_rows, key=lambda row: row["lane"]["repeat"])
                if any(first["identity"][field] != second["identity"][field] for field in identity_fields):
                    fail(f"repeat.{role}.{mode}.{shape}", "archive/content/semantic identity changed between repeats")
    profiles = verify_profiles(protocol, builds, ambient, oracle)
    extras = verify_extra_attempts(protocol, protocol_sha, builds)
    return {
        "schema_version": 1,
        "change": 436,
        "protocol_sha256": protocol_sha,
        "oracle": {"protocol_sha256": protocol["oracle"]["sha256"], "verifier_sha256": protocol["oracle"]["verifier_sha256"]},
        "builds": {role: {"directory": build_dir.name, "revision": build["revision"], "source_manifest": build["source_manifest"], "binaries": build["binaries"]} for role, (build_dir, build) in builds.items()},
        "ambient": {"git_revision": next(iter(ambient))[0], "git_worktree_dirty": next(iter(ambient))[1]},
        "matrix": {"reports": len(rows), "samples_per_report": protocol["samples"], "warmups": protocol["warmups"], "phases": list(PHASE_ORDER), "roles": list(PROFILE_ROLES)},
        "cross_phase_identity": identity_pairs,
        "profiles": profiles,
        "rows": rows,
        "extra_attempts": extras,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-binaries", action="store_true")
    parser.add_argument("--portable-check", action="store_true", help="run without probing retained build binaries")
    parser.add_argument("--require-inventory", action="store_true", help="require and validate the sealed SHA256SUMS inventory")
    parser.add_argument("--stage", choices=STAGES, default="final")
    parser.add_argument("--json", action="store_true", help="emit the verified matrix envelope")
    args = parser.parse_args()
    try:
        if args.portable_check:
            args.require_binaries = False
        result = verify_matrix(require_binaries=args.require_binaries)
        verify_cleanup(args.stage)
        if args.stage != "precleanup" and not (ROOT / "compression.json").is_file():
            fail("compression.json", "post-cleanup replay requires the sealed compression inventory")
        compressed = verify_compression()
        inventory = verify_inventory(args.require_inventory)
        result["stage"] = args.stage
        result["compressed_artifacts"] = compressed
        result["inventory_files"] = inventory
    except (OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True, allow_nan=False))
    else:
        print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
