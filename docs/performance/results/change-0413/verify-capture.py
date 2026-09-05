#!/usr/bin/env python3
"""Fail-closed validator for the bounded 0413 CFB scratch capture.

This is deliberately a read-only evidence checker.  It reads the declared
protocol, build identities, journal, reports, catalogs, and wrapper artifacts;
it never runs Cargo, the benchmark, perf, or a repository command.  JSON and
catalog checks reuse the repository's existing schema validators where they
apply, while this script binds the 0413-specific ABBA, oracle, observer, and
allocation requirements.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


VERSION = "0413-cfb-scratch-verifier-v1"
DEFAULT_ROOT = Path("/tmp/litchi-goal-0413-capture")
CONTROL_REVISION = "ceba0345220c1ca6a7f61f3fac86145b5afc55ca"
SHA256 = re.compile(r"^[0-9a-f]{64}$")
REVISION = re.compile(r"^[0-9a-f]{40}$")
RUSTFLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
RUSTC_PREFIX = "rustc 1.98.1 "
CPU = "2"
ALLOCATOR_SCOPE = "operation_global_system_allocator"
ALLOCATOR_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
)
ALL_ALLOCATOR_FIELDS = ALLOCATOR_FIELDS + (
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
)
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

XLS_CASES = (
    "xls_semantic_open",
    "xls_source_backed_open",
    "xls_eager_open_list_worksheets",
    "xls_source_backed_open_list_worksheets",
    "xls_eager_open_one_cell",
    "xls_source_backed_open_one_cell",
    "xls_owned_source_open",
    "xls_owned_source_open_list_worksheets",
    "xls_owned_source_open_one_cell",
)
OWNED_CASES = XLS_CASES[-3:]
ABBA = ("control", "candidate", "candidate", "control")
GROUPS = (
    ("normal", tuple(f"normal-{i}" for i in range(1, 5))),
    ("allocator", tuple(f"allocator-{i}" for i in range(1, 5))),
    ("guard-normal", tuple(f"guard-normal-{i}" for i in range(1, 5))),
    ("guard-allocator", tuple(f"guard-allocator-{i}" for i in range(1, 5))),
    ("profile", ("control-profile", "candidate-profile")),
    ("stat", tuple(f"stat-{i}" for i in range(1, 5))),
)
LABELS = tuple(label for _, labels in GROUPS for label in labels)

XLS_CORPUS = {
    "name": "xls-comments-opaque-heavy",
    "generator": "litchi-xls-comments-opaque-heavy-v1",
    "package_format": "XLS/CFB",
    "shape": "256-comments-opaque-heavy",
    "payload_kind": "incompressible",
    "compression": "none",
    "entry_count": 257,
    "archive_member_count": 10,
    "entry_bytes": 2097152,
    "uncompressed_payload_bytes": 16858200,
    "archive_bytes": 16995840,
    "archive_sha256": "6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53",
    "target_entry": "Workbook",
    "target_payload_bytes": 80946,
    "target_payload_sha256": "c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041",
}
CFB_CORPORA = {
    "tiny": {
        "name": "cfb-tiny-incompressible",
        "generator": "litchi-cfb-synthetic-v1",
        "package_format": "CFB/OLE2",
        "shape": "tiny",
        "payload_kind": "incompressible",
        "compression": "none",
        "entry_count": 3,
        "archive_member_count": 3,
        "entry_bytes": 512,
        "uncompressed_payload_bytes": 1536,
        "archive_bytes": 3584,
        "archive_sha256": "186750b66895472e6c4b61bdd6e89d7cfd066baec1000cc0e9aab5d86457c0e8",
        "target_entry": "benchmark_stream_00002.bin",
        "target_payload_bytes": 512,
        "target_payload_sha256": "df355b60021f82a84ec1ca06edcf7aea64a5272388c4eed239562fd63d3fceb3",
    },
    "few-large": {
        "name": "cfb-few-large-incompressible",
        "generator": "litchi-cfb-synthetic-v1",
        "package_format": "CFB/OLE2",
        "shape": "few-large",
        "payload_kind": "incompressible",
        "compression": "none",
        "entry_count": 4,
        "archive_member_count": 4,
        "entry_bytes": 4194304,
        "uncompressed_payload_bytes": 16777216,
        "archive_bytes": 16912384,
        "archive_sha256": "4b732058fb9f06fa9166208b207dff2339e2c9caf6e6091a06648a27b3ac4cfa",
        "target_entry": "benchmark_stream_00003.bin",
        "target_payload_bytes": 4194304,
        "target_payload_sha256": "57d0fc8d7b94f2ef821acdb06ac16a3cd450572b825d4e26e0d400df5e838706",
    },
}


class VerificationError(ValueError):
    pass


def fail(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def no_constants(token: str) -> Any:
    raise VerificationError(f"non-finite JSON constant {token!r}")


def no_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        fail(key not in result, f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise VerificationError(f"cannot canonicalize JSON value: {error}") from error


def digest(value: Any) -> str:
    return hashlib.sha256(canonical(value)).hexdigest()


def string(value: Any, context: str) -> str:
    fail(isinstance(value, str) and value != "", f"{context} must be a non-empty string")
    return value


def sha256(value: Any, context: str) -> str:
    value = string(value, context)
    fail(SHA256.fullmatch(value) is not None, f"{context} must be lowercase SHA-256")
    return value


def revision(value: Any, context: str) -> str:
    value = string(value, context)
    fail(REVISION.fullmatch(value) is not None, f"{context} must be a 40-character revision")
    return value


def integer(value: Any, context: str, minimum: int = 0) -> int:
    fail(isinstance(value, int) and not isinstance(value, bool), f"{context} must be an integer")
    fail(value >= minimum, f"{context} must be >= {minimum}")
    return value


def obj(value: Any, context: str) -> dict[str, Any]:
    fail(isinstance(value, dict), f"{context} must be an object")
    return value


def array(value: Any, context: str) -> list[Any]:
    fail(isinstance(value, list), f"{context} must be an array")
    return value


def artifact(root: Path, relative: str | Path) -> tuple[bytes, dict[str, Any]]:
    plain = root / relative
    compressed = Path(str(plain) + ".zst")
    fail(not (plain.exists() and compressed.exists()), f"ambiguous artifact pair: {plain}")
    if plain.exists():
        raw = plain.read_bytes()
        return raw, {"path": str(plain), "compressed": False, "bytes": len(raw), "sha256": hashlib.sha256(raw).hexdigest()}
    fail(compressed.exists(), f"missing artifact: {plain} or {compressed}")
    try:
        raw = subprocess.check_output(["zstd", "-q", "-dc", str(compressed)])
    except (OSError, subprocess.CalledProcessError) as error:
        raise VerificationError(f"cannot decompress {compressed}: {error}") from error
    return raw, {
        "path": str(compressed),
        "compressed": True,
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "artifact_bytes": compressed.stat().st_size,
        "artifact_sha256": hashlib.sha256(compressed.read_bytes()).hexdigest(),
    }


def read_json(root: Path, relative: str | Path) -> tuple[Any, dict[str, Any]]:
    raw, info = artifact(root, relative)
    try:
        value = json.loads(raw.decode("utf-8"), object_pairs_hook=no_duplicates, parse_constant=no_constants)
    except (UnicodeDecodeError, json.JSONDecodeError, VerificationError) as error:
        raise VerificationError(f"cannot parse {info['path']}: {error}") from error
    return value, info


def read_absolute_json(path: Path) -> tuple[Any, dict[str, Any]]:
    fail(path.exists(), f"missing identity artifact: {path}")
    raw = path.read_bytes()
    try:
        value = json.loads(raw.decode("utf-8"), object_pairs_hook=no_duplicates, parse_constant=no_constants)
    except (UnicodeDecodeError, json.JSONDecodeError, VerificationError) as error:
        raise VerificationError(f"cannot parse {path}: {error}") from error
    return value, {"path": str(path), "compressed": False, "bytes": len(raw), "sha256": hashlib.sha256(raw).hexdigest()}


def load_protocol(root: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    value, info = read_json(root, "protocol.json")
    protocol = obj(value, "protocol.json")
    fail(protocol.get("change") == "0413", "protocol change must be 0413")
    fail(string(protocol.get("purpose"), "protocol.purpose"), "protocol purpose missing")
    fail(revision(protocol.get("control_revision"), "protocol.control_revision") == CONTROL_REVISION, "control revision is not the frozen 0413 control")
    revision(protocol.get("candidate_revision"), "protocol.candidate_revision")
    fail(protocol.get("cases") == list(XLS_CASES), "protocol XLS case list mismatch")
    for name, expected in (("normal", {"order": list(ABBA), "samples": 1000, "warmup": 50}), ("allocator", {"order": list(ABBA), "samples": 30, "warmup": 3})):
        group = obj(protocol.get(name), f"protocol.{name}")
        for key, wanted in expected.items():
            fail(group.get(key) == wanted, f"protocol.{name}.{key} mismatch")
    profile = obj(protocol.get("profile"), "protocol.profile")
    fail(profile.get("case") == OWNED_CASES[-1] and profile.get("samples") == 10000 and profile.get("warmup") == 20, "protocol profile mismatch")
    fail(profile.get("event") == "cycles:u" and profile.get("frequency") == 999 and profile.get("call_graph") == "fp,127", "protocol profile configuration mismatch")
    counters = obj(protocol.get("counters"), "protocol.counters")
    fail(counters.get("case") == OWNED_CASES[-1] and counters.get("repetitions") == 2 and counters.get("samples") == 3000 and counters.get("warmup") == 10, "protocol counter sample mismatch")
    fail(counters.get("events") == ",".join(EVENTS), "protocol counter events mismatch")
    fail(protocol.get("cpu") == 2 and protocol.get("filesystem_cache") == "warm", "protocol CPU/cache mismatch")
    guards = obj(protocol.get("guards"), "protocol.guards")
    fail(guards.get("cases") == ["cfb_open"] and guards.get("shapes") == ["tiny", "few-large"] and guards.get("payloads") == ["incompressible"], "protocol guard corpus mismatch")
    for name, expected in (("normal", {"order": list(ABBA), "samples": 1000, "warmup": 50}), ("allocator", {"order": list(ABBA), "samples": 30, "warmup": 3})):
        group = obj(guards.get(name), f"protocol.guards.{name}")
        for key, wanted in expected.items():
            fail(group.get(key) == wanted, f"protocol.guards.{name}.{key} mismatch")
    return protocol, info


def load_identities(root: Path, protocol: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    identities: dict[str, dict[str, Any]] = {}
    output: dict[str, Any] = {}
    for role in ("control", "candidate"):
        # The capture is live under /tmp; the bundle step copies these same
        # sidecars into the capture root.  Prefer the bundled sidecar and
        # retain the live fallback so this checker can validate before copy.
        candidates = [root / f"{role}-build-identity.json", root.parent / "checks" / f"{role}-build-identity.json"]
        if root.resolve() == DEFAULT_ROOT:
            candidates.append(Path(f"/tmp/litchi-goal-0413-{role}-binaries/identity.json"))
        path = next((candidate for candidate in candidates if candidate.exists()), None)
        fail(path is not None, f"missing {role} build identity sidecar")
        value, info = read_absolute_json(path)
        identity = obj(value, f"{role} build identity")
        expected_revision = protocol[f"{role}_revision"]
        fail(identity.get("revision") == expected_revision, f"{role} build revision mismatch")
        fail(identity.get("source_status", identity.get("status")) == "", f"{role} source is dirty")
        if "source_status" in identity:
            fail(identity["source_status"] == "", f"{role}.source_status is not clean")
        if "status" in identity:
            fail(identity["status"] == "", f"{role}.status is not clean")
        fail(identity.get("exit_code") == 0, f"{role} build did not exit 0")
        binaries = obj(identity.get("binaries"), f"{role}.binaries")
        metadata: dict[str, dict[str, Any]] = {}
        for binary in ("litchi-perf-baseline", "litchi-perf-baseline-alloc"):
            entry = obj(binaries.get(binary), f"{role}.binaries.{binary}")
            metadata[binary] = {"bytes": integer(entry.get("bytes"), f"{role}.{binary}.bytes", 1), "sha256": sha256(entry.get("sha256"), f"{role}.{binary}.sha256")}
        identity["_binaries"] = metadata
        identities[role] = identity
        output[role] = {"artifact": info, "revision": identity["revision"], "source_status": "", "binaries": metadata}
    fail(identities["control"]["revision"] != identities["candidate"]["revision"], "control and candidate revisions must differ")
    return identities, output


def parse_time(value: Any, context: str) -> datetime:
    value = string(value, context)
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise VerificationError(f"{context} is not ISO-8601: {error}") from error
    fail(parsed.tzinfo is not None, f"{context} lacks timezone")
    return parsed.astimezone(timezone.utc)


def role_for(label: str) -> str:
    for _, labels in GROUPS:
        if label in labels:
            if labels == ("control-profile", "candidate-profile"):
                return label.split("-", 1)[0]
            return ABBA[labels.index(label)]
    raise VerificationError(f"unknown label {label!r}")


def report_kind(label: str) -> tuple[str, tuple[str, ...], int, int, bool]:
    if label.startswith("normal-"):
        return "xls", XLS_CASES, 1000, 50, False
    if label.startswith("allocator-"):
        return "xls", XLS_CASES, 30, 3, True
    if label.startswith("guard-normal-"):
        return "cfb", ("cfb_open",), 1000, 50, False
    if label.startswith("guard-allocator-"):
        return "cfb", ("cfb_open",), 30, 3, True
    if label.endswith("-profile"):
        return "xls", (OWNED_CASES[-1],), 10000, 20, False
    if label.startswith("stat-"):
        return "xls", (OWNED_CASES[-1],), 3000, 10, False
    raise VerificationError(f"unknown report label {label!r}")


def expected_command(root: Path, protocol: dict[str, Any], label: str, role: str) -> list[str]:
    root = DEFAULT_ROOT  # Recorded paths are provenance, independent of replay location.
    kind, cases, samples, warmup, allocator = report_kind(label)
    binary_name = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    binary = f"/tmp/litchi-goal-0413-{role}-binaries/{binary_name}"
    benchmark = [binary, "--filesystem-cache", "warm", "--case", ",".join(cases), "--samples", str(samples), "--warmup", str(warmup), "--json", str(root / f"{label}.json"), "--corpus-manifest", str(root / f"{label}.catalog.json")]
    if kind == "cfb":
        benchmark += ["--shape", "tiny,few-large", "--payload", "incompressible"]
    if label.startswith("stat-"):
        command = ["perf", "stat", "--no-big-num", "-x,", "-e", protocol["counters"]["events"], "-o", str(root / f"{label}.csv"), "--", *benchmark]
    elif label.endswith("-profile"):
        profile = protocol["profile"]
        command = ["perf", "record", "--no-buildid-cache", "-e", profile["event"], "-F", str(profile["frequency"]), "--call-graph", profile["call_graph"], "-o", str(root / f"{label}.data"), "--", *benchmark]
    else:
        command = ["/usr/bin/time", "-v", "-o", str(root / f"{label}.time.txt"), *benchmark]
    return ["taskset", "-c", CPU, *command]


def validate_journal(root: Path, protocol: dict[str, Any], protocol_info: dict[str, Any], identities: dict[str, dict[str, Any]]) -> dict[str, Any]:
    value, info = read_json(root, "commands.json")
    entries = array(value, "commands.json")
    fail(len(entries) == len(LABELS), f"commands.json has {len(entries)} entries; expected {len(LABELS)}")
    by_label: dict[str, dict[str, Any]] = {}
    starts: list[tuple[datetime, float, str]] = []
    for index, raw in enumerate(entries):
        item = obj(raw, f"commands.json[{index}]")
        label = string(item.get("label"), f"commands.json[{index}].label")
        fail(label in LABELS and label not in by_label, f"unexpected or duplicate journal label {label!r}")
        by_label[label] = item
        role = string(item.get("variant"), f"commands.{label}.variant")
        fail(role == role_for(label), f"{label}: journal role mismatch")
        fail(item.get("protocol_sha256") == protocol_info["sha256"], f"{label}: protocol hash mismatch")
        identity = identities[role]
        fail(item.get("revision") == identity["revision"], f"{label}: journal revision mismatch")
        fail(item.get("source_status") == "", f"{label}: journal source is dirty")
        _, _, _, _, allocator = report_kind(label)
        binary_name = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
        fail(item.get("binary_sha256") == identity["_binaries"][binary_name]["sha256"], f"{label}: journal binary hash mismatch")
        fail(array(item.get("argv"), f"commands.{label}.argv") == expected_command(root, protocol, label, role), f"{label}: argv differs from protocol")
        fail(item.get("cwd") == f"/tmp/litchi-goal-0413-{role}-worktree", f"{label}: cwd mismatch")
        fail(item.get("exit_code") == 0, f"{label}: exit code is not zero")
        integer(item.get("launcher_process_id"), f"{label}: launcher_process_id", 1)
        wall = item.get("wall_seconds")
        fail(isinstance(wall, (int, float)) and not isinstance(wall, bool) and math.isfinite(float(wall)) and wall > 0, f"{label}: wall_seconds must be positive and finite")
        starts.append((parse_time(item.get("started_utc"), f"{label}.started_utc"), float(wall), label))
    fail(tuple(by_label) == LABELS, "journal labels/order do not match capture protocol")
    for previous, current in zip(starts, starts[1:]):
        fail(current[0] >= previous[0], f"journal timestamps are out of order at {current[2]}")
        fail((current[0] - previous[0]).total_seconds() + 0.5 >= previous[1], f"journal commands overlap: {previous[2]}, {current[2]}")
    return {"artifact": info, "protocol_sha256": protocol_info["sha256"], "labels": list(LABELS), "sequential_nonoverlap_verified": True}


def expected_output(case: str, archive: str) -> str:
    if case in ("xls_semantic_open", "xls_source_backed_open", "xls_owned_source_open"):
        return archive
    if case.endswith("list_worksheets"):
        raw = b"".join(len(name.encode()).to_bytes(8, "little") + name.encode() for name in ("Comments", "Untouched"))
        return hashlib.sha256(raw).hexdigest()
    if case.endswith("one_cell"):
        return "e726a50d216e6d71d7c53aabd23ab5e0d4677c3ef1f41fc35410143ebe6381c1"
    raise VerificationError(f"no XLS oracle for {case!r}")


def expected_corpus(corpus: dict[str, Any], kind: str) -> None:
    wanted = XLS_CORPUS if kind == "xls" else CFB_CORPORA.get(corpus.get("shape"))
    fail(wanted is not None, f"unknown {kind} corpus shape {corpus.get('shape')!r}")
    for key, value in wanted.items():
        fail(corpus.get(key) == value, f"corpus.{key} does not match frozen {kind} oracle")
    fail(corpus.get("xlsx") is None, "corpus.xlsx must be null for this capture")
    sha256(corpus.get("archive_sha256"), "corpus.archive_sha256")
    sha256(corpus.get("target_payload_sha256"), "corpus.target_payload_sha256")
    integer(corpus.get("archive_bytes"), "corpus.archive_bytes", 1)
    integer(corpus.get("archive_member_count"), "corpus.archive_member_count", 1)


def validate_catalog(catalog: dict[str, Any], report: dict[str, Any], revision_value: str, context: str) -> None:
    fail(catalog.get("manifest_version") == 2 and catalog.get("manifest_kind") == "corpus-catalog" and catalog.get("catalog_id") == "litchi-perf-corpus-v2", f"{context}: catalog identity mismatch")
    canonicalization = obj(catalog.get("canonicalization"), f"{context}.canonicalization")
    fail(canonicalization == {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}, f"{context}: canonicalization mismatch")
    catalog_hash = sha256(catalog.get("catalog_sha256"), f"{context}.catalog_sha256")
    without_hash = dict(catalog)
    without_hash.pop("catalog_sha256", None)
    fail(digest(without_hash) == catalog_hash, f"{context}: catalog hash does not bind content")
    content_hash = sha256(catalog.get("content_set_sha256"), f"{context}.content_set_sha256")
    build = obj(catalog.get("build"), f"{context}.build")
    fail(build.get("tool") == "litchi-perf-baseline" and build.get("tool_version") == "0.1.0" and build.get("git_revision") == revision_value and build.get("git_worktree_dirty") is False and build.get("source_files") == [], f"{context}: catalog build identity mismatch")
    results = [obj(item, f"{context}.report.results[]") for item in array(report.get("results"), f"{context}.report.results")]
    result_corpora = {f"{item['case']}\0{item['corpus']['package_format']}:{item['corpus']['archive_sha256']}": item["corpus"] for item in results}
    corpora = array(catalog.get("corpora"), f"{context}.corpora")
    ids: dict[str, dict[str, Any]] = {}
    for index, raw in enumerate(corpora):
        item = obj(raw, f"{context}.corpora[{index}]")
        cid = string(item.get("id"), f"{context}.corpora[{index}].id")
        fail(cid not in ids, f"{context}: duplicate corpus id")
        legacy = obj(item.get("legacy_v1"), f"{context}.corpora[{index}].legacy_v1")
        archive = sha256(legacy.get("archive_sha256"), f"{context}.corpora[{index}].legacy_v1.archive_sha256")
        ids[cid] = {"legacy": legacy, "archive": archive, "item": item}
    expected_ids = {f"{corpus['package_format'].lower().replace('/', '-')}:sha256:{corpus['archive_sha256']}" for corpus in (item["corpus"] for item in results)}
    fail(set(ids) == expected_ids, f"{context}: catalog corpus set differs from report")
    for item in results:
        corpus = item["corpus"]
        cid = f"{corpus['package_format'].lower().replace('/', '-')}:sha256:{corpus['archive_sha256']}"
        # The Rust formatter collapses repeated separators; these package
        # formats contain only one separator, so the compact form is exact.
        if corpus["package_format"] == "XLS/CFB":
            cid = f"xls-cfb:sha256:{corpus['archive_sha256']}"
        elif corpus["package_format"] == "CFB/OLE2":
            cid = f"cfb-ole2:sha256:{corpus['archive_sha256']}"
        fail(ids[cid]["legacy"] == corpus, f"{context}: catalog legacy corpus differs from report")
    bindings = array(catalog.get("case_bindings"), f"{context}.case_bindings")
    expected_pairs = set()
    for item in results:
        corpus = item["corpus"]
        cid = "xls-cfb" if corpus["package_format"] == "XLS/CFB" else "cfb-ole2"
        expected_pairs.add((item["case"], f"{cid}:sha256:{corpus['archive_sha256']}"))
    actual_pairs = set()
    projected_bindings: list[dict[str, Any]] = []
    for index, raw in enumerate(bindings):
        item = obj(raw, f"{context}.case_bindings[{index}]")
        pair = (string(item.get("case"), f"{context}.case_bindings[{index}].case"), string(item.get("corpus_id"), f"{context}.case_bindings[{index}].corpus_id"))
        fail(pair not in actual_pairs, f"{context}: duplicate case/corpus binding")
        actual_pairs.add(pair)
        fail(pair in expected_pairs and item.get("role") == "timed", f"{context}: unexpected case binding")
        corpus = next(r for r in results if (r["case"], ("xls-cfb" if r["corpus"]["package_format"] == "XLS/CFB" else "cfb-ole2") + ":sha256:" + r["corpus"]["archive_sha256"]) == pair)["corpus"]
        fail(item.get("legacy_name") == corpus["name"] and item.get("legacy_archive_sha256") == corpus["archive_sha256"], f"{context}: binding legacy identity mismatch")
        projected_bindings.append({"case": item["case"], "corpus_id": item["corpus_id"], "role": item["role"]})
    fail(actual_pairs == expected_pairs, f"{context}: catalog bindings differ from report")
    projected_corpora = []
    for item in corpora:
        members = obj(item.get("members"), f"{context}.corpora[].members")
        member_projection = []
        for member in array(members.get("items"), f"{context}.corpora[].members.items"):
            member = obj(member, f"{context}.corpora[].members.items[]")
            member_projection.append({"ordinal": member.get("ordinal"), "name": member.get("name"), "sha256": member.get("sha256")})
        projected_corpora.append({"id": item["id"], "archive_sha256": obj(item.get("bytes"), f"{context}.corpora[].bytes").get("archive_sha256"), "members": member_projection})
    fail(digest({"corpora": projected_corpora, "case_bindings": projected_bindings}) == content_hash, f"{context}: content-set hash mismatch")
    fail(report.get("corpus_catalog") == {"manifest_version": 2, "catalog_id": catalog["catalog_id"], "catalog_sha256": catalog_hash, "content_set_sha256": content_hash}, f"{context}: report catalog reference mismatch")


def validate_allocation(operation: dict[str, Any], samples: int, measured: bool, context: str) -> dict[str, Any]:
    allocation = obj(operation.get("allocation"), f"{context}.allocation")
    status = "measured" if measured else "unavailable"
    fail(allocation.get("status") == status and allocation.get("scope") == ALLOCATOR_SCOPE, f"{context}: allocation envelope mismatch")
    projection: dict[str, Any] = {"status": allocation["status"], "scope": allocation["scope"]}
    for field in ALL_ALLOCATOR_FIELDS:
        metric = obj(allocation.get(field), f"{context}.allocation.{field}")
        fail(metric.get("status") == status and metric.get("scope") == ALLOCATOR_SCOPE, f"{context}: allocation {field} status/scope mismatch")
        if measured:
            values = array(metric.get("values"), f"{context}.allocation.{field}.values")
            fail(len(values) == samples and all(isinstance(v, int) and not isinstance(v, bool) and v >= 0 for v in values), f"{context}: allocation {field} vector malformed")
            if field in ALLOCATOR_FIELDS:
                projection[field] = values
        else:
            fail("values" not in metric, f"{context}: unavailable allocation {field} exposes values")
    if measured:
        fail(projection["failed_allocation_calls"] == [0] * samples, f"{context}: failed allocations are nonzero")
    return projection


def validate_source(source: dict[str, Any], samples: int, context: str) -> None:
    for key in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads"):
        values = array(source.get(key), f"{context}.{key}")
        fail(len(values) == samples and all(isinstance(v, int) and not isinstance(v, bool) and v >= 0 for v in values), f"{context}.{key} vector malformed")
    xls = obj(source.get("xls"), f"{context}.xls")
    operation = "open+one-cell" if "_one_cell." in context else "open+list" if "_list_worksheets." in context else "open"
    cell = operation == "open+one-cell"
    scope = "caller-provided ReadAt logical ranges: XLS classification catalog v2 (sorted, exact-adjacent coalescing only; overlaps/duplicates preserved; repeated-read union disabled); actual CFB metadata versus opaque payload"
    fail(xls.get("source_counter_scope") == scope, f"{context}: observer scope mismatch")
    fail(xls.get("operation") == operation and xls.get("timing_scope") == operation, f"{context}: operation scope mismatch")
    expected_root = dict(read_calls=362 if cell else 334, read_bytes=138593 if cell else 138459, ordinary_payload_read_calls=0, ordinary_payload_read_bytes=0, max_in_flight_reads=1)
    expected_xls = dict(source_retained_bytes=16995840, complete_archive_materialized_bytes=0, parsed_sheet_counts=2, parsed_cell_counts=int(cell), source_version_checks={"open":168,"open+list":162,"open+one-cell":224}[operation], cfb_structural_read_calls=265, cfb_structural_read_bytes=136704, workbook_global_read_calls=69, workbook_global_read_bytes=1755, selected_worksheet_read_calls=28 if cell else 0, selected_worksheet_read_bytes=134 if cell else 0, unselected_worksheet_read_calls=0, unselected_worksheet_read_bytes=0, opaque_payload_read_calls=0, opaque_payload_read_bytes=0, open_reads_zero_worksheet_payload=not cell, selected_query_reads_only_selected_worksheet=cell)
    for owner, values in [(source,expected_root),(xls,expected_xls)]:
        for key, value in values.items():
            fail(owner.get(key) == [value]*samples, f"{context}: frozen locality counter mismatch: {key}")
    fail(xls.get("implementation") == "source-backed" and "v2" in string(xls.get("source_counter_scope"), f"{context}.xls.source_counter_scope"), f"{context}: XLS observer is not v2")
    fail(xls.get("archive_sha256") == XLS_CORPUS["archive_sha256"] and xls.get("workbook_stream_sha256") == XLS_CORPUS["target_payload_sha256"], f"{context}: source archive/workbook oracle mismatch")
    for key, value in xls.items():
        if isinstance(value, list):
            fail(len(value) == samples, f"{context}.xls.{key} vector length mismatch")
    fail(xls.get("source_version_stability_verified") == [True] * samples, f"{context}: source version stability failed")


def validate_report(root: Path, label: str, protocol: dict[str, Any], identities: dict[str, dict[str, Any]]) -> tuple[dict[str, Any], dict[str, Any]]:
    kind, cases, samples, warmup, allocator = report_kind(label)
    value, report_info = read_json(root, f"{label}.json")
    report = obj(value, label)
    role = role_for(label)
    binary_name = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    tool = obj(report.get("tool"), f"{label}.tool")
    fail(report.get("schema_version") == 1 and tool.get("name") == "litchi-perf-baseline" and tool.get("binary") == binary_name and tool.get("profile") == "release", f"{label}: tool identity mismatch")
    fail(tool.get("instrumentation") == ("system_allocator_operation_scoped" if allocator else "none"), f"{label}: instrumentation mismatch")
    binary = obj(report.get("binary_identity"), f"{label}.binary_identity")
    expected_binary = identities[role]["_binaries"][binary_name]
    fail(binary.get("binary_sha256") == expected_binary["sha256"] and binary.get("binary_bytes") == expected_binary["bytes"] and binary.get("profile") == "release" and binary.get("executable") is True, f"{label}: binary identity mismatch")
    environment = obj(report.get("environment"), f"{label}.environment")
    fail(environment.get("git_revision") == identities[role]["revision"] and environment.get("git_worktree_dirty") is False and environment.get("cpu_affinity") == CPU and environment.get("rustflags") == RUSTFLAGS, f"{label}: environment identity mismatch")
    fail(string(environment.get("rustc_version"), f"{label}.environment.rustc_version").startswith(RUSTC_PREFIX), f"{label}: rustc version mismatch")
    fail(environment.get("allocator") == ("CountingSystemAllocator(std::alloc::System)" if allocator else "Rust system allocator"), f"{label}: allocator identity mismatch")
    config = obj(report.get("configuration"), f"{label}.configuration")
    for key, wanted in (("samples_per_case", samples), ("warmup_iterations_per_case", warmup), ("filesystem_cache_states", ["warm"]), ("filesystem_fresh_child_per_sample", True), ("filesystem_process_isolated", True), ("filesystem_root_selected", False), ("cases", list(cases)), ("execution_workers", [1])):
        fail(config.get(key) == wanted, f"{label}: configuration.{key} mismatch")
    if kind == "cfb":
        fail(config.get("corpus_shapes") == ["tiny", "few-large"] and config.get("payload_kinds") == ["incompressible"], f"{label}: CFB guard shapes/payload mismatch")
    results = array(report.get("results"), f"{label}.results")
    fail(len(results) == (2 if kind == "cfb" else len(cases)), f"{label}: result count mismatch")
    seen_cases: list[str] = []
    observer: dict[str, Any] = {}
    corpus_summaries: list[dict[str, Any]] = []
    for index, raw in enumerate(results):
        result = obj(raw, f"{label}.results[{index}]")
        case = string(result.get("case"), f"{label}.results[{index}].case")
        fail(case in cases and (kind == "cfb" or case not in seen_cases), f"{label}: unexpected/duplicate case {case!r}")
        seen_cases.append(case)
        corpus = obj(result.get("corpus"), f"{label}.{case}.corpus")
        expected_corpus(corpus, kind)
        elapsed = obj(result.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
        elapsed_values = array(elapsed.get("samples"), f"{label}.{case}.elapsed_ns.samples")
        fail(elapsed.get("unit") == "ns" and len(elapsed_values) == samples and elapsed_values == sorted(elapsed_values) and all(isinstance(v, int) and not isinstance(v, bool) and v > 0 for v in elapsed_values), f"{label}.{case}: elapsed samples malformed")
        order = array(elapsed.get("sample_order"), f"{label}.{case}.elapsed_ns.sample_order")
        fail(len(order) == samples and sorted(order) == list(range(samples)), f"{label}.{case}: sample order is not a complete permutation")
        if kind == "xls":
            fail(result.get("output_sha256") == expected_output(case, corpus["archive_sha256"]), f"{label}.{case}: output oracle mismatch")
            if case.startswith("xls_source_backed_"):
                source = obj(result.get("source"), f"{label}.{case}.source")
                validate_source(source, samples, f"{label}.{case}.source")
            else:
                # None covers both serde omission and an explicit JSON null.
                fail(result.get("source") is None, f"{label}.{case}: plain/eager result fabricated source observer")
            operation = obj(result.get("operation_metrics"), f"{label}.{case}.operation_metrics")
            fail(operation.get("sample_count") == samples and operation.get("sample_indices") == order, f"{label}.{case}: operation sample order mismatch")
            allocation_projection = validate_allocation(operation, samples, allocator, f"{label}.{case}.operation_metrics")
            observer[case] = {"source": result.get("source"), "operation_source": operation.get("source"), "allocation": allocation_projection}
            corpus_summaries.append({"case": case, "archive_sha256": corpus["archive_sha256"], "output_sha256": result["output_sha256"]})
        else:
            fail(result.get("source") is None and result.get("output_sha256") is None and result.get("sink") is None and "operation_metrics" not in result, f"{label}.{case}: CFB open result carries unexpected operation/output evidence")
            observer[f"{case}\0{corpus['shape']}"] = {"source": None, "operation_source": None, "allocation": None}
            corpus_summaries.append({"case": case, "shape": corpus["shape"], "archive_sha256": corpus["archive_sha256"]})
    if kind == "xls":
        fail(tuple(seen_cases) == cases, f"{label}: result case order mismatch")
    else:
        fail({item["shape"] for item in corpus_summaries} == {"tiny", "few-large"}, f"{label}: CFB guard result shapes incomplete")
    # Existing repository validators check the full operation and parallel
    # metric trees.  Their exceptions remain evidence failures here.
    try:
        from tools import perf_compare
        perf_compare.validate_parallel_metrics(report)
        for result in results:
            if "operation_metrics" in result:
                elapsed = result["elapsed_ns"]
                perf_compare._validate_operation_metrics(result["operation_metrics"], label, elapsed["samples"], report["schema_version"], elapsed_sample_order=elapsed["sample_order"])
    except (ImportError, AttributeError) as error:
        raise VerificationError(f"repository schema validator unavailable: {error}") from error
    catalog_value, catalog_info = read_json(root, f"{label}.catalog.json")
    validate_catalog(obj(catalog_value, f"{label}.catalog"), report, identities[role]["revision"], f"{label}.catalog")
    return {"artifact": report_info, "catalog_artifact": catalog_info, "role": role, "kind": kind, "cases": list(cases), "samples": samples, "warmup": warmup, "corpus": corpus_summaries, "observer": observer}, report


def validate_wrappers(root: Path, protocol: dict[str, Any]) -> dict[str, Any]:
    time_labels = LABELS[:16]
    for label in time_labels:
        raw, _ = artifact(root, f"{label}.time.txt")
        text = raw.decode("utf-8")
        fail("Exit status: 0" in text, f"{label}: time wrapper did not record exit 0")
        fail(re.search(r"^\s*Maximum resident set size \(kbytes\):\s*\d+\s*$", text, re.MULTILINE) is not None, f"{label}: missing RSS evidence")
    for label in ("stat-1", "stat-2", "stat-3", "stat-4"):
        raw, _ = artifact(root, f"{label}.csv")
        rows = []
        for row in csv.reader(line for line in raw.decode("utf-8").splitlines() if line.strip() and not line.lstrip().startswith("#")):
            fail(len(row) >= 5, f"{label}: malformed perf stat row")
            rows.append(row)
        fail([row[2] for row in rows] == list(EVENTS), f"{label}: PMU event order differs from protocol")
        for row in rows:
            try:
                count, running, percent = float(row[0]), int(row[3]), float(row[4])
            except ValueError as error:
                raise VerificationError(f"{label}: malformed PMU numeric field") from error
            fail(math.isfinite(count) and count >= 0 and running > 0 and math.isfinite(percent) and 0 < percent <= 100, f"{label}: invalid PMU evidence")
    for label in ("control-profile", "candidate-profile"):
        raw, _ = artifact(root, f"{label}.data")
        fail(bool(raw), f"{label}: empty perf data")
    return {"time_reports": len(time_labels), "stat_reports": 4, "profile_data": 2}


def compare_observers(reports: dict[str, dict[str, Any]]) -> dict[str, Any]:
    checks: dict[str, Any] = {}
    for group, labels in (("normal", GROUPS[0][1]), ("allocator", GROUPS[1][1]), ("stat/profile", ("stat-1", "stat-2", "stat-3", "stat-4", "control-profile", "candidate-profile"))):
        base = reports[labels[0]]["observer"]
        for label in labels[1:]:
            fail(reports[label]["observer"] == base, f"{group}: observer/source/allocation evidence differs at {label}")
        checks[group] = {"labels": list(labels), "exact_observer_equality": True}
    for group in ("guard-normal", "guard-allocator"):
        labels = dict(GROUPS)[group]
        base = reports[labels[0]]["observer"]
        for label in labels[1:]:
            fail(reports[label]["observer"] == base, f"{group}: CFB observer envelope differs at {label}")
        checks[group] = {"labels": list(labels), "exact_observer_equality": True}
    return checks


def verify(root: Path) -> dict[str, Any]:
    fail(root.is_dir(), f"capture directory does not exist: {root}")
    protocol, protocol_info = load_protocol(root)
    identities, identity_output = load_identities(root, protocol)
    journal = validate_journal(root, protocol, protocol_info, identities)
    reports: dict[str, dict[str, Any]] = {}
    report_values: dict[str, dict[str, Any]] = {}
    for label in LABELS:
        summary, report = validate_report(root, label, protocol, identities)
        reports[label] = summary
        report_values[label] = report
    # Every XLS report is bound to the same generated archive; guard reports
    # are independently bound to the two frozen CFB archive identities.
    for label in LABELS:
        expected = XLS_CORPUS if reports[label]["kind"] == "xls" else CFB_CORPORA
        if reports[label]["kind"] == "xls":
            fail({item["archive_sha256"] for item in reports[label]["corpus"]} == {expected["archive_sha256"]}, f"{label}: XLS corpus identity drift")
        else:
            fail({item["archive_sha256"] for item in reports[label]["corpus"]} == {value["archive_sha256"] for value in expected.values()}, f"{label}: CFB corpus identity drift")
    observers = compare_observers(reports)
    wrappers = validate_wrappers(root, protocol)
    return {
        "schema_version": 1,
        "verifier": VERSION,
        "performance_claim": "none",
        "protocol": {"change": protocol["change"], "purpose": protocol["purpose"], "sha256": protocol_info["sha256"], "control_revision": protocol["control_revision"], "candidate_revision": protocol["candidate_revision"], "normal": protocol["normal"], "allocator": protocol["allocator"], "guards": protocol["guards"]},
        "identities": identity_output,
        "journal": journal,
        "reports": {label: {key: value for key, value in summary.items() if key != "observer"} for label, summary in reports.items()},
        "observer_equality": observers,
        "wrapper_artifacts": wrappers,
        "validation": {"reports": len(reports), "all_22_reports_verified": len(reports) == 22, "oracle_and_catalog_binding_verified": True, "sample_counts_and_permutations_verified": True, "v2_observers_equal_across_roles": True, "plain_source_observers_absent_or_null": True, "allocation_calls_and_bytes_equal_across_roles": True, "live_peak_snapshots_excluded_from_equality": True},
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--repo-root", type=Path, default=Path.cwd(), help="workspace containing the tools schema validators")
    parser.add_argument("--output", type=Path, help="write verification JSON after all checks pass")
    args = parser.parse_args()
    output = args.output or args.root / "0413-verification.json"
    try:
        sys.path.insert(0, str(args.repo_root.resolve()))
        result = verify(args.root)
        output.write_text(json.dumps(result, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    except (OSError, VerificationError, ValueError, KeyError, TypeError) as error:
        print(f"litchi-goal-0413-verify: FAIL: {error}", file=sys.stderr)
        return 2
    print(f"litchi-goal-0413-verify: PASS: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
