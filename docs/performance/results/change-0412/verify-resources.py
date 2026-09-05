#!/usr/bin/env python3
"""Fail-closed verifier for the bounded 0412 resource capture.

The capture is deliberately split into the original six-case normal and
allocator runs, plain ``OwnedSource`` runs, one ``perf record`` run, and three
``perf stat`` runs.  This script only reads those artifacts.  It never invokes
Cargo, the benchmark, ``perf``, or a repository command, and writes a compact
observation only after every identity and shape check succeeds.

The JSON reports are schema 1 reports with schema 2 corpus-catalog sidecars.
Published bundles may losslessly compress any input artifact as ``.zst``;
plain files take precedence only when the compressed counterpart is absent so
an ambiguous pair cannot silently hide a changed artifact.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import re
import struct
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SCRIPT_VERSION = "0412-resource-verifier-v1"
DEFAULT_ROOT = Path("/tmp/litchi-goal-0412-capture")
DEFAULT_OUTPUT_NAME = "resources.json"
CAPTURE_ROOT = "/tmp/litchi-goal-0412-capture"
RUSTFLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
RUSTC_PREFIX = "rustc 1.98.1 "
CPU = "2"

CONTROL_SCOPE = "caller-provided ReadAt logical ranges: actual CFB metadata versus opaque payload"
CANDIDATE_SCOPE = (
    "caller-provided ReadAt logical ranges: XLS classification catalog v2 "
    "(sorted, exact-adjacent coalescing only; overlaps/duplicates preserved; "
    "repeated-read union disabled); actual CFB metadata versus opaque payload"
)

NORMAL_LABELS = tuple(f"normal-{index}" for index in range(1, 5))
ALLOCATOR_LABELS = tuple(f"allocator-{index}" for index in range(1, 5))
OWNED_LABELS = tuple(f"owned-{index}" for index in range(1, 5))
OWNED_ALLOCATOR_LABELS = tuple(f"owned-allocator-{index}" for index in range(1, 3))
STAT_LABELS = tuple(f"owned-stat-{index}" for index in range(1, 4))
PROFILE_LABEL = "owned-profile"
RSS_LABELS = NORMAL_LABELS + ALLOCATOR_LABELS + OWNED_LABELS + OWNED_ALLOCATOR_LABELS

EVENTS = (
    "task-clock",
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "page-faults",
    "context-switches",
    "cpu-migrations",
    "l2_cache_req_stat.dc_access_in_l2",
    "l2_cache_req_stat.dc_hit_in_l2",
)
NONZERO_EVENTS = set(EVENTS) - {"context-switches", "cpu-migrations"}

SHA256 = re.compile(r"^[0-9a-f]{64}$")
REVISION = re.compile(r"^[0-9a-f]{40}$")


class VerificationError(ValueError):
    """The capture is incomplete, malformed, or fails an identity gate."""


def reject_constant(token: str) -> Any:
    raise VerificationError(f"non-finite JSON constant {token!r}")


def no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def canonical_json(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise VerificationError(f"value is not canonically serializable: {error}") from error


def digest(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def expect(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def obj(value: Any, context: str) -> dict[str, Any]:
    expect(isinstance(value, dict), f"{context} must be an object")
    return value


def list_value(value: Any, context: str) -> list[Any]:
    expect(isinstance(value, list), f"{context} must be a list")
    return value


def string(value: Any, context: str, *, nonempty: bool = True) -> str:
    expect(isinstance(value, str), f"{context} must be a string")
    expect(not nonempty or value != "", f"{context} must not be empty")
    return value


def integer(value: Any, context: str, *, minimum: int = 0) -> int:
    expect(isinstance(value, int) and not isinstance(value, bool), f"{context} must be an integer")
    expect(value >= minimum, f"{context} must be >= {minimum}")
    return value


def finite_number(value: Any, context: str, *, minimum: float | None = None) -> float:
    expect(isinstance(value, (int, float)) and not isinstance(value, bool), f"{context} must be numeric")
    result = float(value)
    expect(math.isfinite(result), f"{context} must be finite")
    if minimum is not None:
        expect(result >= minimum, f"{context} must be >= {minimum}")
    return result


def artifact(root: Path, relative: str | Path) -> tuple[bytes, dict[str, Any]]:
    """Read a plain artifact or its optional zstd representation.

    A compressed fallback is intentionally optional.  If both forms are
    present, fail closed: otherwise a verifier could accidentally validate a
    stale plain file while a published compressed file had changed.
    """

    plain = root / relative
    compressed = Path(str(plain) + ".zst")
    if plain.exists() and compressed.exists():
        raise VerificationError(f"ambiguous artifact pair: {plain} and {compressed}")
    if plain.exists():
        try:
            raw = plain.read_bytes()
        except OSError as error:
            raise VerificationError(f"cannot read {plain}: {error}") from error
        return raw, {
            "path": str(plain),
            "compressed": False,
            "bytes": len(raw),
            "sha256": hashlib.sha256(raw).hexdigest(),
        }
    if not compressed.exists():
        raise VerificationError(f"missing artifact: {plain} (or {compressed})")
    try:
        raw = subprocess.check_output(["zstd", "-q", "-dc", str(compressed)])
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError(f"cannot decompress {compressed}: {error}") from error
    return raw, {
        "path": str(compressed),
        "compressed": True,
        "artifact_bytes": compressed.stat().st_size,
        "artifact_sha256": hashlib.sha256(compressed.read_bytes()).hexdigest(),
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
    }


def read_text(root: Path, relative: str | Path) -> tuple[str, dict[str, Any]]:
    raw, info = artifact(root, relative)
    try:
        return raw.decode("utf-8"), info
    except UnicodeDecodeError as error:
        raise VerificationError(f"{info['path']} is not UTF-8: {error}") from error


def read_json(root: Path, relative: str | Path) -> tuple[Any, dict[str, Any]]:
    text, info = read_text(root, relative)
    try:
        value = json.loads(
            text,
            object_pairs_hook=no_duplicate_pairs,
            parse_constant=reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, VerificationError) as error:
        raise VerificationError(f"cannot parse {info['path']}: {error}") from error
    return value, info


def require_sha256(value: Any, context: str) -> str:
    value = string(value, context)
    expect(SHA256.fullmatch(value) is not None, f"{context} must be a lowercase SHA-256 digest")
    return value


def require_revision(value: Any, context: str) -> str:
    value = string(value, context)
    expect(REVISION.fullmatch(value) is not None, f"{context} must be a 40-character revision")
    return value


def expected_protocol(root: Path) -> dict[str, Any]:
    value, _ = read_json(root, "protocol.json")
    protocol = obj(value, "protocol.json")
    expect(protocol.get("change") == "0412", "protocol change must be 0412")
    expect(isinstance(protocol.get("purpose"), str) and protocol["purpose"], "protocol purpose missing")
    control_revision = require_revision(protocol.get("control_revision"), "protocol.control_revision")
    candidate_revision = require_revision(protocol.get("candidate_revision"), "protocol.candidate_revision")
    cases = tuple(string(item, "protocol.cases[]") for item in list_value(protocol.get("cases"), "protocol.cases"))
    expect(cases == (
        "xls_semantic_open",
        "xls_source_backed_open",
        "xls_eager_open_list_worksheets",
        "xls_source_backed_open_list_worksheets",
        "xls_eager_open_one_cell",
        "xls_source_backed_open_one_cell",
    ), "protocol cases do not match the original six-case contract")
    normal = obj(protocol.get("normal"), "protocol.normal")
    allocator = obj(protocol.get("allocator"), "protocol.allocator")
    owned = obj(protocol.get("owned_source"), "protocol.owned_source")
    profile = obj(protocol.get("profile"), "protocol.profile")
    counters = obj(protocol.get("counters"), "protocol.counters")
    expect(normal.get("order") == ["control", "candidate", "candidate", "control"], "protocol normal order mismatch")
    expect(allocator.get("order") == ["control", "candidate", "candidate", "control"], "protocol allocator order mismatch")
    expect(normal.get("samples") == 500 and normal.get("warmup") == 20, "protocol normal sample contract mismatch")
    expect(allocator.get("samples") == 30 and allocator.get("warmup") == 3, "protocol allocator sample contract mismatch")
    expect(owned.get("repetitions") == 4 and owned.get("samples") == 500 and owned.get("warmup") == 20, "protocol owned sample contract mismatch")
    expect(owned.get("allocation_repetitions") == 2 and owned.get("allocation_samples") == 30 and owned.get("allocation_warmup") == 3, "protocol owned allocator sample contract mismatch")
    expect(profile.get("case") == "xls_owned_source_open_one_cell", "protocol profile case mismatch")
    expect(profile.get("samples") == 10000 and profile.get("warmup") == 20, "protocol profile sample contract mismatch")
    expect(profile.get("event") == "cycles:u" and profile.get("frequency") == 999 and profile.get("call_graph") == "fp,127", "protocol profile configuration mismatch")
    expect(counters.get("case") == "xls_owned_source_open_one_cell", "protocol counter case mismatch")
    expect(counters.get("repetitions") == 3 and counters.get("samples") == 3000 and counters.get("warmup") == 10, "protocol counter sample contract mismatch")
    expect(counters.get("events") == ",".join(EVENTS), "protocol counter event contract mismatch")
    expect(isinstance(counters.get("scope"), str) and counters["scope"], "protocol counter scope missing")
    expect(protocol.get("cpu") == 2 and protocol.get("filesystem_cache") == "warm", "protocol CPU/cache contract mismatch")
    return {
        "raw": protocol,
        "control_revision": control_revision,
        "candidate_revision": candidate_revision,
        "cases": cases,
        "owned_cases": (
            "xls_owned_source_open",
            "xls_owned_source_open_list_worksheets",
            "xls_owned_source_open_one_cell",
        ),
    }


def load_identities(root: Path, protocol: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    identities: dict[str, dict[str, Any]] = {}
    metadata: dict[str, Any] = {}
    for role, revision in (("control", protocol["control_revision"]), ("candidate", protocol["candidate_revision"])):
        value, info = read_json(root, f"{role}-build-identity.json")
        identity = obj(value, f"{role}-build-identity.json")
        expect(identity.get("revision") == revision, f"{role} build identity revision mismatch")
        status = identity.get("source_status", identity.get("status"))
        expect(status == "", f"{role} build identity is not clean")
        if "source_status" in identity:
            expect(identity.get("source_status") == "", f"{role}.source_status is not clean")
        if "status" in identity:
            expect(identity.get("status") == "", f"{role}.status is not clean")
        expect(identity.get("exit_code") == 0, f"{role} build did not exit 0")
        binaries = obj(identity.get("binaries"), f"{role}.binaries")
        binary_metadata: dict[str, Any] = {}
        for binary in ("litchi-perf-baseline", "litchi-perf-baseline-alloc"):
            entry = obj(binaries.get(binary), f"{role}.binaries.{binary}")
            integer(entry.get("bytes"), f"{role}.binaries.{binary}.bytes", minimum=1)
            binary_metadata[binary] = {
                "bytes": entry["bytes"],
                "sha256": require_sha256(entry.get("sha256"), f"{role}.binaries.{binary}.sha256"),
            }
        identity["_status"] = status
        identity["_binary_metadata"] = binary_metadata
        identities[role] = identity
        metadata[role] = {
            "artifact": info,
            "revision": revision,
            "source_status": status,
            "binaries": binary_metadata,
        }
    expect(identities["control"]["revision"] != identities["candidate"]["revision"], "control and candidate revisions must differ")
    return identities, metadata


def parse_timestamp(value: Any, context: str) -> datetime:
    text = string(value, context)
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as error:
        raise VerificationError(f"{context} is not an ISO-8601 timestamp: {error}") from error
    expect(parsed.tzinfo is not None, f"{context} must include a timezone")
    return parsed.astimezone(timezone.utc)


def expected_command(label: str, protocol: dict[str, Any], variant: str) -> list[str]:
    expected_binary = "litchi-perf-baseline-alloc" if "allocator" in label else "litchi-perf-baseline"
    binary = f"/tmp/litchi-goal-0412-{variant}-binaries/{expected_binary}"
    if label in NORMAL_LABELS:
        cases = ",".join(protocol["cases"])
        samples, warmup = 500, 20
    elif label in ALLOCATOR_LABELS:
        cases = ",".join(protocol["cases"])
        samples, warmup = 30, 3
    elif label in OWNED_LABELS:
        cases = ",".join(protocol["owned_cases"])
        samples, warmup = 500, 20
    elif label in OWNED_ALLOCATOR_LABELS:
        cases = ",".join(protocol["owned_cases"])
        samples, warmup = 30, 3
    elif label == PROFILE_LABEL:
        cases, samples, warmup = protocol["owned_cases"][-1], 10000, 20
    elif label in STAT_LABELS:
        cases, samples, warmup = protocol["owned_cases"][-1], 3000, 10
    else:
        raise VerificationError(f"unexpected command label {label!r}")
    benchmark = [
        binary,
        "--filesystem-cache", "warm",
        "--case", cases,
        "--samples", str(samples),
        "--warmup", str(warmup),
        "--json", f"{CAPTURE_ROOT}/{label}.json",
        "--corpus-manifest", f"{CAPTURE_ROOT}/{label}.catalog.json",
    ]
    if label == PROFILE_LABEL:
        command = [
            "perf", "record", "--no-buildid-cache", "-e", "cycles:u", "-F", "999",
            "--call-graph", "fp,127", "-o", f"{CAPTURE_ROOT}/{label}.data", "--", *benchmark,
        ]
    elif label in STAT_LABELS:
        command = [
            "perf", "stat", "--no-big-num", "-x,", "-e", ",".join(EVENTS),
            "-o", f"{CAPTURE_ROOT}/{label}.csv", "--", *benchmark,
        ]
    else:
        command = ["/usr/bin/time", "-v", "-o", f"{CAPTURE_ROOT}/{label}.time.txt", *benchmark]
    return ["taskset", "-c", CPU, *command]


def expected_labels() -> tuple[str, ...]:
    return NORMAL_LABELS + ALLOCATOR_LABELS + OWNED_LABELS + OWNED_ALLOCATOR_LABELS + (PROFILE_LABEL,) + STAT_LABELS


def validate_commands(root: Path, protocol: dict[str, Any], identities: dict[str, dict[str, Any]]) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    value, info = read_json(root, "commands.json")
    entries = list_value(value, "commands.json")
    expected = expected_labels()
    expect(len(entries) == len(expected), f"commands.json has {len(entries)} entries; expected {len(expected)}")
    by_label: dict[str, dict[str, Any]] = {}
    starts: list[tuple[datetime, float, str]] = []
    for index, value in enumerate(entries):
        item = obj(value, f"commands.json[{index}]")
        label = string(item.get("label"), f"commands.json[{index}].label")
        expect(label not in by_label, f"commands.json has duplicate label {label!r}")
        expect(label in expected, f"commands.json has unexpected label {label!r}")
        by_label[label] = item
        variant = string(item.get("variant"), f"commands.json.{label}.variant")
        expected_variant = "control" if label in (NORMAL_LABELS[0], NORMAL_LABELS[3], ALLOCATOR_LABELS[0], ALLOCATOR_LABELS[3]) else "candidate"
        expect(variant == expected_variant, f"{label}: variant must be {expected_variant}")
        identity = identities[variant]
        expect(item.get("revision") == identity["revision"], f"{label}: journal revision mismatch")
        expect(item.get("source_status") == "", f"{label}: journal source_status must be empty")
        binary_name = "litchi-perf-baseline-alloc" if "allocator" in label else "litchi-perf-baseline"
        expected_hash = identity["_binary_metadata"][binary_name]["sha256"]
        expect(item.get("binary_sha256") == expected_hash, f"{label}: journal binary hash mismatch")
        argv = list_value(item.get("argv"), f"commands.json.{label}.argv")
        expect(all(isinstance(token, str) for token in argv), f"{label}: argv tokens must be strings")
        expect(argv == expected_command(label, protocol, variant), f"{label}: argv differs from capture protocol")
        expected_cwd = f"/tmp/litchi-goal-0412-{variant}-worktree"
        expect(item.get("cwd") == expected_cwd, f"{label}: cwd differs from capture protocol")
        expect(item.get("exit_code") == 0, f"{label}: command exit code is not zero")
        integer(item.get("launcher_process_id"), f"{label}: launcher_process_id", minimum=1)
        wall = finite_number(item.get("wall_seconds"), f"{label}: wall_seconds", minimum=0.0)
        expect(wall > 0.0, f"{label}: wall_seconds must be positive")
        started = parse_timestamp(item.get("started_utc"), f"{label}: started_utc")
        starts.append((started, wall, label))
    expect(set(by_label) == set(expected), "commands.json label set does not match the declared capture")
    for previous, current in zip(starts, starts[1:]):
        expect(current[0] >= previous[0], f"commands.json is not chronologically ordered at {current[2]}")
        # capture.py waits for each child before appending the next journal
        # entry.  Timestamp and monotonic wall clocks are different clocks, so
        # tolerate only sub-second journal/write scheduling jitter.
        expect(
            (current[0] - previous[0]).total_seconds() + 0.5 >= previous[1],
            f"capture commands overlap: {previous[2]} and {current[2]}",
        )
    journal = {
        "artifact": info,
        "labels": list(expected),
        "entries": {
            label: {
                "variant": by_label[label]["variant"],
                "revision": by_label[label]["revision"],
                "binary_sha256": by_label[label]["binary_sha256"],
                "cwd": by_label[label]["cwd"],
                "argv": by_label[label]["argv"],
                "started_utc": by_label[label]["started_utc"],
                "wall_seconds": by_label[label]["wall_seconds"],
                "exit_code": by_label[label]["exit_code"],
            }
            for label in expected
        },
        "sequential_nonoverlap_verified": True,
    }
    return by_label, journal


def canonical_corpus_id(corpus: dict[str, Any]) -> str:
    package_format = string(corpus.get("package_format"), "corpus.package_format")
    archive_sha256 = require_sha256(corpus.get("archive_sha256"), "corpus.archive_sha256")
    slug = ""
    for character in package_format:
        if character.isascii() and character.isalnum():
            slug += character.lower()
        elif not slug or not slug.endswith("-"):
            slug += "-"
    slug = slug.strip("-")
    expect(slug != "", "corpus package_format has no identifier characters")
    return f"{slug}:sha256:{archive_sha256}"


def validate_catalog(catalog: dict[str, Any], report: dict[str, Any], cases: tuple[str, ...], revision: str, context: str) -> None:
    expect(catalog.get("manifest_version") == 2, f"{context}: manifest_version must be 2")
    expect(catalog.get("manifest_kind") == "corpus-catalog", f"{context}: manifest_kind mismatch")
    expect(catalog.get("catalog_id") == "litchi-perf-corpus-v2", f"{context}: catalog_id mismatch")
    canonicalization = obj(catalog.get("canonicalization"), f"{context}.canonicalization")
    expect(canonicalization.get("algorithm") == "sorted-json-utf8-compact-v1", f"{context}: canonicalization algorithm mismatch")
    expect(canonicalization.get("hash") == "sha256", f"{context}: canonicalization hash mismatch")
    catalog_hash = require_sha256(catalog.get("catalog_sha256"), f"{context}.catalog_sha256")
    without_hash = dict(catalog)
    without_hash.pop("catalog_sha256", None)
    expect(digest(without_hash) == catalog_hash, f"{context}: catalog_sha256 does not bind catalog content")
    content_hash = require_sha256(catalog.get("content_set_sha256"), f"{context}.content_set_sha256")
    build = obj(catalog.get("build"), f"{context}.build")
    expect(build.get("git_revision") == revision, f"{context}: build revision mismatch")
    expect(build.get("git_worktree_dirty") is False, f"{context}: catalog build is dirty")
    corpora = list_value(catalog.get("corpora"), f"{context}.corpora")
    expect(corpora, f"{context}: corpus list is empty")
    report_corpus = obj(report["results"][0].get("corpus"), f"{context}.report_corpus")
    corpus_id = canonical_corpus_id(report_corpus)
    matching = [obj(item, f"{context}.corpora[]") for item in corpora if item.get("id") == corpus_id]
    expect(len(matching) == 1, f"{context}: report corpus is absent or duplicated in catalog")
    expect(matching[0].get("legacy_v1") == report_corpus, f"{context}: catalog legacy corpus differs from report")
    bindings = list_value(catalog.get("case_bindings"), f"{context}.case_bindings")
    expect(len(bindings) == len(cases), f"{context}: case binding count mismatch")
    actual_cases: list[str] = []
    projected_bindings: list[dict[str, Any]] = []
    for index, item_value in enumerate(bindings):
        item = obj(item_value, f"{context}.case_bindings[{index}]")
        case = string(item.get("case"), f"{context}.case_bindings[{index}].case")
        expect(case not in actual_cases, f"{context}: duplicate case binding {case!r}")
        actual_cases.append(case)
        expect(item.get("corpus_id") == corpus_id, f"{context}: binding corpus id mismatch")
        expect(item.get("legacy_name") == report_corpus.get("name"), f"{context}: binding legacy name mismatch")
        expect(item.get("legacy_archive_sha256") == report_corpus.get("archive_sha256"), f"{context}: binding archive hash mismatch")
        expect(item.get("role") == "timed", f"{context}: non-timed binding present")
        projected_bindings.append({"case": case, "corpus_id": item["corpus_id"], "role": item["role"]})
    expect(set(actual_cases) == set(cases), f"{context}: catalog cases differ from report cases")
    projected_corpora = []
    for item in corpora:
        corpus_item = obj(item, f"{context}.corpora[]")
        members = obj(corpus_item.get("members"), f"{context}.corpora[].members")
        member_projection = []
        for member in list_value(members.get("items"), f"{context}.corpora[].members.items"):
            member = obj(member, f"{context}.corpora[].members.items[]")
            member_projection.append({"ordinal": member.get("ordinal"), "name": member.get("name"), "sha256": member.get("sha256")})
        projected_corpora.append({"id": corpus_item.get("id"), "archive_sha256": obj(corpus_item.get("bytes"), f"{context}.corpus.bytes").get("archive_sha256"), "members": member_projection})
    expect(digest({"corpora": projected_corpora, "case_bindings": projected_bindings}) == content_hash, f"{context}: content_set_sha256 does not bind catalog content")
    reference = obj(report.get("corpus_catalog"), f"{context}.report_reference")
    expect(reference == {
        "manifest_version": 2,
        "catalog_id": catalog["catalog_id"],
        "catalog_sha256": catalog_hash,
        "content_set_sha256": content_hash,
    }, f"{context}: report corpus catalog reference mismatch")


def expected_output(corpus_hash: str, case: str) -> str:
    if case in ("xls_semantic_open", "xls_source_backed_open", "xls_owned_source_open"):
        return corpus_hash
    if case.endswith("list_worksheets"):
        names = ("Comments", "Untouched")
        raw = b"".join(len(name.encode("utf-8")).to_bytes(8, "little") + name.encode("utf-8") for name in names)
        return hashlib.sha256(raw).hexdigest()
    if case.endswith("one_cell"):
        return hashlib.sha256(struct.pack("<d", 42.0)).hexdigest()
    raise VerificationError(f"cannot derive output oracle for {case!r}")


def validate_report(root: Path, label: str, protocol: dict[str, Any], identities: dict[str, dict[str, Any]], corpus_state: dict[str, Any]) -> dict[str, Any]:
    report_value, report_info = read_json(root, f"{label}.json")
    report = obj(report_value, label)
    if label in NORMAL_LABELS:
        cases, samples, warmup, allocator = protocol["cases"], 500, 20, False
    elif label in ALLOCATOR_LABELS:
        cases, samples, warmup, allocator = protocol["cases"], 30, 3, True
    elif label in OWNED_LABELS:
        cases, samples, warmup, allocator = protocol["owned_cases"], 500, 20, False
    elif label in OWNED_ALLOCATOR_LABELS:
        cases, samples, warmup, allocator = protocol["owned_cases"], 30, 3, True
    elif label == PROFILE_LABEL:
        cases, samples, warmup, allocator = (protocol["owned_cases"][-1],), 10000, 20, False
    elif label in STAT_LABELS:
        cases, samples, warmup, allocator = (protocol["owned_cases"][-1],), 3000, 10, False
    else:
        raise VerificationError(f"unknown report label {label!r}")
    variant = "control" if label in (NORMAL_LABELS[0], NORMAL_LABELS[3], ALLOCATOR_LABELS[0], ALLOCATOR_LABELS[3]) else "candidate"
    identity = identities[variant]
    expected_binary = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    expect(report.get("schema_version") == 1, f"{label}: schema_version must be 1")
    tool = obj(report.get("tool"), f"{label}.tool")
    expect(tool.get("name") == "litchi-perf-baseline", f"{label}: tool name mismatch")
    expect(tool.get("binary") == expected_binary, f"{label}: tool binary mismatch")
    expect(tool.get("profile") == "release", f"{label}: tool profile mismatch")
    expected_instrumentation = "system_allocator_operation_scoped" if allocator else "none"
    expect(tool.get("instrumentation") == expected_instrumentation, f"{label}: instrumentation mismatch")
    binary_identity = obj(report.get("binary_identity"), f"{label}.binary_identity")
    expected_binary_meta = identity["_binary_metadata"][expected_binary]
    expect(binary_identity.get("binary_sha256") == expected_binary_meta["sha256"], f"{label}: report binary hash mismatch")
    expect(binary_identity.get("binary_bytes") == expected_binary_meta["bytes"], f"{label}: report binary size mismatch")
    expect(binary_identity.get("profile") == "release" and binary_identity.get("executable") is True, f"{label}: executable identity mismatch")
    environment = obj(report.get("environment"), f"{label}.environment")
    expect(environment.get("git_revision") == identity["revision"], f"{label}: report revision mismatch")
    expect(environment.get("git_worktree_dirty") is False, f"{label}: report worktree is dirty")
    expect(environment.get("cpu_affinity") == CPU, f"{label}: CPU affinity mismatch")
    expect(environment.get("rustflags") == RUSTFLAGS, f"{label}: RUSTFLAGS mismatch")
    expect(string(environment.get("rustc_version"), f"{label}.environment.rustc_version").startswith(RUSTC_PREFIX), f"{label}: rustc version mismatch")
    expected_allocator = "CountingSystemAllocator(std::alloc::System)" if allocator else "Rust system allocator"
    expect(environment.get("allocator") == expected_allocator, f"{label}: allocator identity mismatch")
    configuration = obj(report.get("configuration"), f"{label}.configuration")
    expect(configuration.get("samples_per_case") == samples, f"{label}: samples_per_case mismatch")
    expect(configuration.get("warmup_iterations_per_case") == warmup, f"{label}: warmup mismatch")
    expect(configuration.get("filesystem_cache_states") == ["warm"], f"{label}: filesystem cache mismatch")
    expect(configuration.get("filesystem_fresh_child_per_sample") is True, f"{label}: fresh-child contract missing")
    expect(configuration.get("filesystem_process_isolated") is True, f"{label}: process isolation missing")
    expect(configuration.get("filesystem_root_selected") is False, f"{label}: filesystem root must not be selected")
    expect(configuration.get("cases") == list(cases), f"{label}: report case set mismatch")
    expect(configuration.get("execution_workers") == [1], f"{label}: execution worker contract mismatch")
    results = list_value(report.get("results"), f"{label}.results")
    expect(len(results) == len(cases), f"{label}: result count mismatch")
    result_map: dict[str, dict[str, Any]] = {}
    source_scopes: set[str] = set()
    for index, value in enumerate(results):
        result = obj(value, f"{label}.results[{index}]")
        case = string(result.get("case"), f"{label}.results[{index}].case")
        expect(case in cases and case not in result_map, f"{label}: unexpected or duplicate case {case!r}")
        result_map[case] = result
        corpus = obj(result.get("corpus"), f"{label}.{case}.corpus")
        corpus_hash = require_sha256(corpus.get("archive_sha256"), f"{label}.{case}.corpus.archive_sha256")
        integer(corpus.get("archive_bytes"), f"{label}.{case}.corpus.archive_bytes", minimum=1)
        integer(corpus.get("archive_member_count"), f"{label}.{case}.corpus.archive_member_count", minimum=1)
        if "corpus" not in corpus_state:
            corpus_state["corpus"] = corpus
            corpus_state["hash"] = corpus_hash
        expect(corpus == corpus_state["corpus"], f"{label}.{case}: corpus identity differs across reports")
        elapsed = obj(result.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
        expect(elapsed.get("unit") == "ns", f"{label}.{case}: elapsed unit mismatch")
        elapsed_samples = list_value(elapsed.get("samples"), f"{label}.{case}.elapsed_ns.samples")
        expect(len(elapsed_samples) == samples, f"{label}.{case}: elapsed sample count mismatch")
        expect(all(isinstance(item, int) and not isinstance(item, bool) and item > 0 for item in elapsed_samples), f"{label}.{case}: elapsed samples must be positive integers")
        order = list_value(elapsed.get("sample_order"), f"{label}.{case}.elapsed_ns.sample_order")
        expect(len(order) == samples and sorted(order) == list(range(samples)), f"{label}.{case}: sample order is not a complete permutation")
        expect(result.get("output_sha256") == expected_output(corpus_hash, case), f"{label}.{case}: output hash differs from oracle")
        operation_metrics = obj(result.get("operation_metrics"), f"{label}.{case}.operation_metrics")
        expect(operation_metrics.get("sample_count") == samples, f"{label}.{case}: operation sample count mismatch")
        operation_indices = list_value(operation_metrics.get("sample_indices"), f"{label}.{case}.operation_metrics.sample_indices")
        expect(len(operation_indices) == samples and sorted(operation_indices) == list(range(samples)), f"{label}.{case}: operation sample indices are not a complete permutation")
        allocation = obj(operation_metrics.get("allocation"), f"{label}.{case}.operation_metrics.allocation")
        expect(allocation.get("status") == ("measured" if allocator else "unavailable"), f"{label}.{case}: allocation status mismatch")
        if allocator:
            for metric in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes"):
                metric_value = obj(allocation.get(metric), f"{label}.{case}.allocation.{metric}")
                expect(metric_value.get("status") == "measured", f"{label}.{case}: allocation metric {metric} is not measured")
                values = list_value(metric_value.get("values"), f"{label}.{case}.allocation.{metric}.values")
                expect(len(values) == samples and all(isinstance(item, int) and not isinstance(item, bool) and item >= 0 for item in values), f"{label}.{case}: allocation vector {metric} malformed")
            failed = allocation["failed_allocation_calls"]["values"]
            expect(failed == [0] * samples, f"{label}.{case}: failed allocations are nonzero")
        if case.startswith("xls_source_backed_"):
            source = obj(result.get("source"), f"{label}.{case}.source")
            xls = obj(source.get("xls"), f"{label}.{case}.source.xls")
            scope = string(xls.get("source_counter_scope"), f"{label}.{case}.source.xls.source_counter_scope")
            source_scopes.add(scope)
            expect(xls.get("implementation") == "source-backed", f"{label}.{case}: source implementation mismatch")
            expect(xls.get("archive_sha256") == corpus_hash, f"{label}.{case}: source archive hash mismatch")
            for field in ("read_calls", "read_bytes", "max_in_flight_reads"):
                values = list_value(source.get(field), f"{label}.{case}.source.{field}")
                expect(len(values) == samples, f"{label}.{case}: source {field} length mismatch")
            stability = list_value(xls.get("source_version_stability_verified"), f"{label}.{case}.source.xls.source_version_stability_verified")
            expect(stability == [True] * samples, f"{label}.{case}: source version stability failed")
        else:
            expect(result.get("source") is None, f"{label}.{case}: eager/plain case must have source=None")
    expect(set(result_map) == set(cases), f"{label}: result case set is incomplete")
    catalog_value, catalog_info = read_json(root, f"{label}.catalog.json")
    validate_catalog(obj(catalog_value, f"{label}.catalog"), report, tuple(cases), identity["revision"], f"{label}.catalog")
    return {
        "artifact": report_info,
        "catalog_artifact": catalog_info,
        "variant": variant,
        "revision": identity["revision"],
        "binary": expected_binary,
        "instrumentation": expected_instrumentation,
        "cases": list(cases),
        "samples": samples,
        "warmup": warmup,
        "source_scopes": sorted(source_scopes),
    }


def parse_rss(root: Path, label: str) -> tuple[int, dict[str, Any]]:
    text, info = read_text(root, f"{label}.time.txt")
    expect("Exit status: 0" in text, f"{label}: time wrapper did not record exit status 0")
    matches = re.findall(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", text, re.MULTILINE)
    expect(len(matches) == 1, f"{label}: expected one maximum RSS line")
    value = int(matches[0])
    expect(value > 0, f"{label}: RSS must be positive")
    return value, info


def parse_counter_csv(root: Path, label: str) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    text, info = read_text(root, f"{label}.csv")
    rows: list[list[str]] = []
    for row in csv.reader(line for line in text.splitlines() if line.strip() and not line.lstrip().startswith("#")):
        expect(len(row) >= 5, f"{label}: malformed perf stat row {row!r}")
        rows.append(row)
    expect(len(rows) == len(EVENTS), f"{label}: expected {len(EVENTS)} PMU rows, got {len(rows)}")
    values: dict[str, dict[str, Any]] = {}
    for row in rows:
        count_text, unit, event, running_text, percent_text = row[:5]
        expect(event in EVENTS and event not in values, f"{label}: unknown or duplicate event {event!r}")
        try:
            count = float(count_text)
            running_ns = int(running_text)
            scheduled_percent = float(percent_text)
        except ValueError as error:
            raise VerificationError(f"{label}: non-numeric PMU field in {row!r}") from error
        expect(math.isfinite(count) and count >= 0, f"{label}.{event}: count must be finite and nonnegative")
        expect(running_ns > 0, f"{label}.{event}: running time must be positive")
        expect(math.isfinite(scheduled_percent) and 0.0 < scheduled_percent <= 100.0, f"{label}.{event}: scheduling percentage must be in (0, 100]")
        if event in NONZERO_EVENTS:
            expect(count > 0.0, f"{label}.{event}: PMU count must be positive")
        values[event] = {
            "count": count,
            "unit": unit,
            "running_ns": running_ns,
            "scheduled_percent": scheduled_percent,
        }
    expect(list(values) == list(EVENTS), f"{label}: PMU event order differs from protocol")
    return values, info


def parse_profile(root: Path, protocol: dict[str, Any]) -> dict[str, Any]:
    data, data_info = artifact(root, "owned-profile.data")
    expect(data, "owned-profile.data is empty")
    text, self_info = read_text(root, "owned-profile-self.stdout")
    expect("# Total Lost Samples: 0" in text, "owned-profile: perf report recorded lost samples")
    match = re.search(r"^# Event count \(approx\.\):\s*(\d+)\s*$", text, re.MULTILINE)
    expect(match is not None, "owned-profile: missing whole-process event count")
    event_count = int(match.group(1))
    expect(event_count > 0, "owned-profile: whole-process event count must be positive")
    return {
        "event": protocol["profile"]["event"],
        "frequency": protocol["profile"]["frequency"],
        "call_graph": protocol["profile"]["call_graph"],
        "lost_samples": 0,
        "whole_process_event_count": event_count,
        "data_artifact": data_info,
        "self_report_artifact": self_info,
    }


def verify(root: Path) -> dict[str, Any]:
    expect(root.is_dir(), f"capture directory does not exist: {root}")
    protocol = expected_protocol(root)
    identities, identity_output = load_identities(root, protocol)
    command_map, command_output = validate_commands(root, protocol, identities)
    corpus_state: dict[str, Any] = {}
    reports: dict[str, Any] = {}
    for label in NORMAL_LABELS + ALLOCATOR_LABELS + OWNED_LABELS + OWNED_ALLOCATOR_LABELS + (PROFILE_LABEL,) + STAT_LABELS:
        reports[label] = validate_report(root, label, protocol, identities, corpus_state)
    expect(set(reports["normal-1"]["source_scopes"]) == {CONTROL_SCOPE}, "normal control observer scope marker mismatch")
    expect(set(reports["normal-4"]["source_scopes"]) == {CONTROL_SCOPE}, "normal control observer scope marker mismatch")
    expect(set(reports["normal-2"]["source_scopes"]) == {CANDIDATE_SCOPE}, "normal candidate observer scope marker mismatch")
    expect(set(reports["normal-3"]["source_scopes"]) == {CANDIDATE_SCOPE}, "normal candidate observer scope marker mismatch")
    expect(set(reports["allocator-1"]["source_scopes"]) == {CONTROL_SCOPE}, "allocator control observer scope marker mismatch")
    expect(set(reports["allocator-4"]["source_scopes"]) == {CONTROL_SCOPE}, "allocator control observer scope marker mismatch")
    expect(set(reports["allocator-2"]["source_scopes"]) == {CANDIDATE_SCOPE}, "allocator candidate observer scope marker mismatch")
    expect(set(reports["allocator-3"]["source_scopes"]) == {CANDIDATE_SCOPE}, "allocator candidate observer scope marker mismatch")
    for label in OWNED_LABELS + OWNED_ALLOCATOR_LABELS + (PROFILE_LABEL,) + STAT_LABELS:
        expect(reports[label]["source_scopes"] == [], f"{label}: plain OwnedSource report unexpectedly carries source observer scope")
    rss: dict[str, int] = {}
    rss_artifacts: dict[str, Any] = {}
    for label in RSS_LABELS:
        rss[label], rss_artifacts[label] = parse_rss(root, label)
    counters: dict[str, Any] = {}
    counter_artifacts: dict[str, Any] = {}
    for label in STAT_LABELS:
        events, info = parse_counter_csv(root, label)
        cycles = events["cycles"]["count"]
        instructions = events["instructions"]["count"]
        expect(cycles > 0.0, f"{label}: cycles count must be positive")
        counters[label] = {
            "scope": protocol["raw"]["counters"]["scope"],
            "events": events,
            "ipc_from_scaled_whole_process_counts": instructions / cycles,
            "branch_miss_percent_from_scaled_whole_process_counts": 100.0 * events["branch-misses"]["count"] / events["branches"]["count"],
        }
        counter_artifacts[label] = info
    profile = parse_profile(root, protocol["raw"])
    expect(reports[PROFILE_LABEL]["cases"] == ["xls_owned_source_open_one_cell"], "owned profile case mismatch")
    return {
        "schema_version": 1,
        "verifier": SCRIPT_VERSION,
        "performance_claim": "none",
        "scope": "whole process: corpus generation, layout/oracle, warmups, operations, cloning, validation and serialization; not operation-local",
        "protocol": {
            "change": protocol["raw"]["change"],
            "purpose": protocol["raw"]["purpose"],
            "control_revision": protocol["control_revision"],
            "candidate_revision": protocol["candidate_revision"],
            "normal": protocol["raw"]["normal"],
            "allocator": protocol["raw"]["allocator"],
            "owned_source": protocol["raw"]["owned_source"],
            "profile": protocol["raw"]["profile"],
            "counters": protocol["raw"]["counters"],
            "cpu": protocol["raw"]["cpu"],
            "filesystem_cache": protocol["raw"]["filesystem_cache"],
        },
        "binary_identity": identity_output,
        "commands": command_output,
        "reports": reports,
        "corpus": {
            "name": corpus_state["corpus"].get("name"),
            "package_format": corpus_state["corpus"].get("package_format"),
            "archive_bytes": corpus_state["corpus"].get("archive_bytes"),
            "archive_sha256": corpus_state["hash"],
            "target_entry": corpus_state["corpus"].get("target_entry"),
            "target_payload_bytes": corpus_state["corpus"].get("target_payload_bytes"),
            "target_payload_sha256": corpus_state["corpus"].get("target_payload_sha256"),
        },
        "rss_kib": rss,
        "rss_artifacts": rss_artifacts,
        "profiles": {"owned_source_open_one_cell": profile},
        "counters": counters,
        "counter_artifacts": counter_artifacts,
        "observer_scope": {
            "control": CONTROL_SCOPE,
            "candidate": CANDIDATE_SCOPE,
            "differ": CONTROL_SCOPE != CANDIDATE_SCOPE,
            "normal_abba_verified": True,
            "allocator_abba_verified": True,
            "plain_owned_source_has_no_scope": True,
        },
        "validation": {
            "expected_rss_capture_count": 14,
            "observed_rss_capture_count": len(rss),
            "expected_profile_capture_count": 1,
            "expected_counter_capture_count": 3,
            "exact_command_argv_verified": True,
            "clean_revisions_verified": True,
            "binary_hashes_verified": True,
            "positive_pmu_counts_and_scheduling_verified": True,
            "sequential_capture_verified": True,
            "original_normal_allocator_case_sets_separate": True,
            "owned_source_reports_source_none": True,
        },
        "limitations": [
            "Perf scales multiplexed hardware events; retain scheduled percentages and do not infer operation-local IPC.",
            "Native L2 requests/hits are not exact L1 or LLC metrics.",
            "RSS is a child high-water mark for each whole-process dispatch, not an operation peak.",
            "No production before/after, cold-cache, physical-I/O, remote-source, scaling, or exact request-size claim is inferred.",
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", "--out", dest="root", type=Path, default=DEFAULT_ROOT, help="0412 capture root")
    parser.add_argument("--output", "--json-out", dest="output", type=Path, help="verification JSON path (default: ROOT/resources.json)")
    args = parser.parse_args(argv)
    output = args.output or args.root / DEFAULT_OUTPUT_NAME
    try:
        result = verify(args.root)
        output.write_text(json.dumps(result, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    except (OSError, VerificationError, ValueError, KeyError, TypeError) as error:
        print(f"litchi-goal-0412-resources: FAIL: {error}", file=sys.stderr)
        return 2
    print(f"litchi-goal-0412-resources: PASS: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
