#!/usr/bin/env python3
"""Fail-closed verifier for the portable 0471 memory evidence bundle."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Mapping, Sequence

import analyze


ROOT = Path(__file__).resolve().parent
HEX = re.compile(r"[0-9a-f]{64}\Z")
LANES = ("A1", "B1", "B2", "A2", "A-full", "B-full", "A-heap", "B-heap")
ROLES = {"A1": "control", "A2": "control", "A-full": "control", "A-heap": "control",
         "B1": "candidate", "B2": "candidate", "B-full": "candidate", "B-heap": "candidate"}
FIXTURES = {
    "test-data/poi/test-data/spreadsheet/54016.xls": "2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a",
    "test-data/rtf/watermark.rtf": "48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb",
}
EXPECTED_CHANGED_FILE = "crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs"
PRIOR_CHANGED_FILES = {
    "crates/litchi-xlsx/src/raw/compact.rs",
    "crates/litchi-xlsx/src/raw/web.rs",
    "crates/litchi-xlsx/src/raw/web/check.rs",
    EXPECTED_CHANGED_FILE,
}
GATES = ("xlsx-fmt", "xlsx-tests", "workspace-check", "xlsx-clippy", "xlsx-rustdoc", "boundaries")


class VerificationError(ValueError):
    """Raised when an identity, artifact, or replay check fails."""


def fail(message: str) -> None:
    raise VerificationError(message)


def read_json(path: Path, label: str = "JSON") -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(
                stream,
                object_pairs_hook=_object_pairs,
                parse_constant=lambda value: (_ for _ in ()).throw(
                    VerificationError(f"{label}: non-finite JSON value {value!r}")
                ),
            )
    except (OSError, UnicodeError, json.JSONDecodeError, VerificationError) as error:
        raise VerificationError(f"{label}: cannot read {path}: {error}") from error


def _object_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8")
    except (TypeError, ValueError) as error:
        raise VerificationError(f"cannot canonicalize JSON: {error}") from error


def canonical_equal(left: Any, right: Any) -> bool:
    return canonical(left) == canonical(right)


def sha256_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            size = 0
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        raise VerificationError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), size


def _regular(path: Path, label: str) -> Path:
    if not path.is_file() or path.is_symlink():
        fail(f"{label}: regular file required: {path}")
    return path


def _relative(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or "\\" in value:
        fail(f"{label}: expected a relative POSIX path")
    path = Path(value)
    if path.is_absolute() or path.as_posix() != value or ".." in path.parts:
        fail(f"{label}: path is not relative and traversal-free")
    return path


def bundle_file(root: Path, value: Any, label: str) -> Path:
    path = root / _relative(value, label)
    try:
        if not path.resolve(strict=True).is_relative_to(root.resolve()):
            fail(f"{label}: path escapes the bundle")
    except OSError as error:
        fail(f"{label}: cannot resolve path: {error}")
    return _regular(path, label)


def binding(path: Path, root: Path) -> dict[str, Any]:
    digest, size = sha256_file(path)
    try:
        name = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        name = path.name
    return {"path": name, "sha256": digest, "bytes": size}


def _digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or HEX.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def verify_protocol(root: Path = ROOT) -> dict[str, Any]:
    protocol = read_json(root / "protocol.json", "protocol.json")
    if not isinstance(protocol, dict):
        fail("protocol.json: expected an object")
    if protocol.get("schema") != "litchi-0471-protocol-v1":
        fail("protocol.json: schema mismatch")
    checks = {
        "probe_order": ["A1", "B1", "B2", "A2"],
        "probe_cases": list(analyze.PROBE_CASES),
        "probe_shapes": list(analyze.PROBE_SHAPES),
        "probe_samples": analyze.PROBE_SAMPLES,
        "probe_warmups": analyze.PROBE_WARMUPS,
        "full_guard_order": ["A-full", "B-full"],
        "full_guard_samples": analyze.FULL_SAMPLES,
        "full_guard_warmups": analyze.FULL_WARMUPS,
        "heap_order": ["A-heap", "B-heap"],
        "heap_samples": analyze.HEAP_SAMPLES,
        "heap_warmups": analyze.HEAP_WARMUPS,
        "heap_case": "xlsx_one_percent_commit_save",
        "heap_shape": "dense-wide",
        "workers": 1,
        "writer_shapes": ["payload-heavy"],
    }
    for key, expected in checks.items():
        if protocol.get(key) != expected:
            fail(f"protocol.json: {key} differs from the frozen experiment")
    _digest(protocol.get("capture_driver_sha256"), "protocol.capture_driver_sha256")
    if not isinstance(protocol.get("build_path"), str) or not protocol["build_path"].startswith("/tmp/"):
        fail("protocol.build_path: expected the recorded absolute temporary path")
    return protocol


def _role_binding(root: Path, role: str) -> tuple[dict[str, Any], Path]:
    path = root / f"{role}-binding.json"
    value = read_json(path, f"{role}-binding.json")
    if not isinstance(value, dict):
        fail(f"{role}-binding.json: expected an object")
    expected_schema = "litchi-0470-role-binding-v1" if role == "control" else "litchi-0471-role-binding-v1"
    if value.get("schema") != expected_schema:
        fail(f"{role}: role binding schema differs")
    if value.get("role") != role or value.get("clean_build") is not True:
        fail(f"{role}: role or clean_build is invalid")
    for key in ("revision", "binary_sha256", "bytes", "binary_path", "build_path", "build_receipt_path",
                "build_receipt_sha256", "source_manifest", "source_manifest_sha256", "included_fixtures",
                "source_binding_sha256"):
        if key not in value:
            fail(f"{role}: missing binding field {key}")
    _digest(value["binary_sha256"], f"{role}.binary_sha256")
    _digest(value["build_receipt_sha256"], f"{role}.build_receipt_sha256")
    _digest(value["source_manifest_sha256"], f"{role}.source_manifest_sha256")
    _digest(value["source_binding_sha256"], f"{role}.source_binding_sha256")
    if not isinstance(value["revision"], str) or len(value["revision"]) < 7:
        fail(f"{role}.revision: malformed revision")
    if isinstance(value["bytes"], bool) or not isinstance(value["bytes"], int) or value["bytes"] <= 0:
        fail(f"{role}.bytes: malformed binary size")
    if value["included_fixtures"] != FIXTURES:
        fail(f"{role}.included_fixtures differs")
    manifest = bundle_file(root, value["source_manifest"], f"{role}.source_manifest")
    digest, _ = sha256_file(manifest)
    if digest != value["source_manifest_sha256"]:
        fail(f"{role}: source manifest digest differs")
    source_binding_path = root / ("prior/candidate-source-binding.json" if role == "control" else "candidate-source-binding.json")
    source_binding = _regular(source_binding_path, f"{role} source binding")
    digest, _ = sha256_file(source_binding)
    if digest != value["source_binding_sha256"]:
        fail(f"{role}: source binding digest differs")
    source = read_json(source_binding, f"{role} source binding")
    if not isinstance(source, dict) or source.get("source_manifest_sha256") != value["source_manifest_sha256"]:
        fail(f"{role}: source binding does not bind its manifest")
    if source.get("revision") != value["revision"] or source.get("source_count") != len(read_json(manifest, f"{role} manifest")):
        fail(f"{role}: source binding revision/count differs")
    if source.get("included_fixtures") != FIXTURES:
        fail(f"{role}: source binding fixtures differ")
    expected_source_schema = "litchi-0470-source-binding-v1" if role == "control" else "litchi-0471-source-binding-v1"
    if source.get("schema") != expected_source_schema or source.get("build_path") != read_json(root / "protocol.json")["build_path"]:
        fail(f"{role}: source binding schema/build path differs")
    changed_files = source.get("changed_files")
    if not isinstance(changed_files, list) or any(not isinstance(path, str) for path in changed_files):
        fail(f"{role}: source binding changed_files is malformed")
    expected_changed = sorted(PRIOR_CHANGED_FILES) if role == "control" else [EXPECTED_CHANGED_FILE]
    if sorted(changed_files) != expected_changed:
        fail(f"{role}: source binding changed_files differs")
    build_path = _relative(value["build_receipt_path"], f"{role}.build_receipt_path")
    receipt_path = root / build_path
    receipt = _regular(receipt_path, f"{role} build receipt")
    digest, _ = sha256_file(receipt)
    if digest != value["build_receipt_sha256"]:
        fail(f"{role}: build receipt digest differs")
    verify_build_receipt(root, receipt_path, protocol=read_json(root / "protocol.json"), prior=(role == "control"))
    if role == "control":
        reused_path = bundle_file(root, value.get("reused_binding_path"), "control.reused_binding_path")
        reused_digest, _ = sha256_file(reused_path)
        if reused_digest != value.get("reused_binding_sha256"):
            fail("control: reused binding digest differs")
        reused = read_json(reused_path, "prior candidate binding")
        if reused.get("schema") != "litchi-0470-role-binding-v1" or reused.get("role") != "candidate":
            fail("control: reused binding schema/role differs")
        for field in ("revision", "binary_sha256", "bytes", "build_path", "build_receipt_sha256", "included_fixtures", "source_binding_sha256"):
            if reused.get(field) != value.get(field):
                fail(f"control: reused binding field {field} differs")
    return value, source


def verify_build_receipt(root: Path, path: Path, *, protocol: Mapping[str, Any], prior: bool) -> dict[str, Any]:
    value = read_json(path, str(path))
    if not isinstance(value, dict) or value.get("exit_code") != 0:
        fail(f"{path}: build did not succeed")
    if value.get("schema") not in {"litchi-0470-command-v1", "litchi-0471-command-v1"}:
        fail(f"{path}: unsupported build receipt schema")
    if value.get("cwd") != protocol.get("build_path"):
        fail(f"{path}: build path differs from protocol")
    environment = value.get("environment")
    if not isinstance(environment, dict) or environment.get("RUSTUP_TOOLCHAIN") != "1.98.1" or environment.get("CARGO_PROFILE_RELEASE_DEBUG") != "1":
        fail(f"{path}: build environment is not the frozen release configuration")
    outputs = value.get("outputs")
    if not isinstance(outputs, dict) or set(outputs) != {".stdout", ".stderr"}:
        fail(f"{path}: build outputs are incomplete")
    for suffix, item in outputs.items():
        if not isinstance(item, dict) or set(item) != {"sha256", "bytes"}:
            fail(f"{path}: malformed build output binding {suffix}")
        _digest(item.get("sha256"), f"{path}.{suffix}.sha256")
        artifact = path.with_name(path.stem + suffix)
        digest, size = sha256_file(_regular(artifact, f"{path}.{suffix}"))
        if digest != item["sha256"] or size != item["bytes"]:
            fail(f"{path}: build output {suffix} differs")
    return value


def verify_source_pair(root: Path, bindings: Mapping[str, Mapping[str, Any]]) -> list[str]:
    manifests: dict[str, dict[str, str]] = {}
    for role, value in bindings.items():
        manifest = read_json(bundle_file(root, value["source_manifest"], f"{role}.source_manifest"), f"{role} manifest")
        if not isinstance(manifest, dict) or not manifest or any(not isinstance(k, str) or not isinstance(v, str) or HEX.fullmatch(v) is None for k, v in manifest.items()):
            fail(f"{role}: malformed source manifest")
        manifests[role] = manifest
    control, candidate = manifests["control"], manifests["candidate"]
    if len(control) != 6993 or len(candidate) != 6993:
        fail("source manifests: expected 6,993 files in both roles")
    if set(control) != set(candidate):
        fail("source manifests: path sets differ")
    changed = [path for path in sorted(control) if control[path] != candidate[path]]
    if changed != [EXPECTED_CHANGED_FILE]:
        fail(f"source manifests: expected exactly {EXPECTED_CHANGED_FILE}, got {changed}")
    for role, value in bindings.items():
        if value["included_fixtures"] != FIXTURES:
            fail(f"{role}: binding fixtures differ")
    return changed


def _expected_artifacts(heap: bool) -> set[str]:
    common = {"started.json", "corpus-catalog.json", "stdout.log", "report.json", "resource.log", "stderr.log"}
    if heap:
        # capture.py records the compressed trace.  The three heaptrack_print
        # artifacts are produced by the later exporter and are bound below.
        common.add("heaptrack.zst")
    return common


def _report_config(report: Mapping[str, Any], lane: str, protocol: Mapping[str, Any], samples: int, warmups: int) -> None:
    if report.get("schema_version") != 1:
        fail(f"{lane}: report schema differs")
    configuration = report.get("configuration")
    if not isinstance(configuration, dict) or configuration.get("samples_per_case") != samples or configuration.get("warmup_iterations_per_case") != warmups:
        fail(f"{lane}: samples/warmups differ")
    if samples == analyze.PROBE_SAMPLES:
        if tuple(configuration.get("cases", ())) != tuple(protocol["probe_cases"]):
            fail(f"{lane}: normal case selector differs")
        if tuple(configuration.get("writer_shapes", ())) != tuple(protocol["writer_shapes"]):
            fail(f"{lane}: normal writer selector differs")
    elif samples == analyze.HEAP_SAMPLES:
        if configuration.get("cases") != [protocol["heap_case"]] or configuration.get("xlsx_shapes") != [protocol["heap_shape"]]:
            fail(f"{lane}: heap selector differs")


def _verify_corpus_catalog(path: Path, label: str) -> dict[str, Any]:
    catalog = read_json(path, label)
    if not isinstance(catalog, dict):
        fail(f"{label}: expected an object")
    if catalog.get("manifest_version") != 2 or catalog.get("manifest_kind") != "corpus-catalog" or catalog.get("catalog_id") != "litchi-perf-corpus-v2":
        fail(f"{label}: catalog identity differs")
    if catalog.get("canonicalization") != {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}:
        fail(f"{label}: catalog canonicalization differs")
    catalog_digest = _digest(catalog.get("catalog_sha256"), f"{label}.catalog_sha256")
    content_digest = _digest(catalog.get("content_set_sha256"), f"{label}.content_set_sha256")
    without_catalog_hash = dict(catalog)
    without_catalog_hash.pop("catalog_sha256", None)
    computed_catalog = hashlib.sha256(canonical(without_catalog_hash)).hexdigest()
    if computed_catalog != catalog_digest:
        fail(f"{label}: catalog_sha256 does not match canonical catalog")
    corpora = catalog.get("corpora")
    bindings = catalog.get("case_bindings")
    if not isinstance(corpora, list) or not isinstance(bindings, list):
        fail(f"{label}: catalog corpora/bindings are malformed")
    content_corpora = []
    for index, corpus in enumerate(corpora):
        if not isinstance(corpus, dict) or not isinstance(corpus.get("id"), str):
            fail(f"{label}.corpora[{index}]: malformed corpus")
        bytes_value = corpus.get("bytes")
        members = corpus.get("members")
        if not isinstance(bytes_value, dict) or not isinstance(bytes_value.get("archive_sha256"), str) or not isinstance(members, dict) or not isinstance(members.get("items"), list):
            fail(f"{label}.corpora[{index}]: missing content identity")
        content_members = []
        for member in members["items"]:
            if not isinstance(member, dict) or not isinstance(member.get("ordinal"), int) or not isinstance(member.get("name"), str):
                fail(f"{label}.corpora[{index}]: malformed member identity")
            content_members.append({"ordinal": member["ordinal"], "name": member["name"], "sha256": member.get("sha256")})
        content_corpora.append({"id": corpus["id"], "archive_sha256": bytes_value["archive_sha256"], "members": content_members})
    content_bindings = []
    for index, item in enumerate(bindings):
        if not isinstance(item, dict) or not all(isinstance(item.get(key), str) for key in ("case", "corpus_id", "role")):
            fail(f"{label}.case_bindings[{index}]: malformed binding")
        content_bindings.append({"case": item["case"], "corpus_id": item["corpus_id"], "role": item["role"]})
    computed_content = hashlib.sha256(canonical({"corpora": content_corpora, "case_bindings": content_bindings})).hexdigest()
    if computed_content != content_digest:
        fail(f"{label}: content_set_sha256 does not match catalog")
    return {
        "manifest_version": catalog["manifest_version"],
        "catalog_id": catalog["catalog_id"],
        "catalog_sha256": catalog_digest,
        "content_set_sha256": content_digest,
    }


def verify_lane(root: Path, lane: str, bindings: Mapping[str, Mapping[str, Any]], protocol: Mapping[str, Any]) -> None:
    role = ROLES[lane]
    receipt_path = root / lane / "receipt.json"
    receipt = read_json(receipt_path, f"{lane}/receipt.json")
    if not isinstance(receipt, dict) or receipt.get("schema") != "litchi-0471-capture-v1":
        fail(f"{lane}: capture receipt schema differs")
    binding_path = root / f"{role}-binding.json"
    expected_binding_sha, _ = sha256_file(binding_path)
    if receipt.get("role") != role or receipt.get("lane") != lane or receipt.get("binding_sha256") != expected_binding_sha:
        fail(f"{lane}: role or binding identity differs")
    if receipt.get("protocol_sha256") != sha256_file(root / "protocol.json")[0] or receipt.get("driver_sha256") != protocol["capture_driver_sha256"]:
        fail(f"{lane}: protocol/driver identity differs")
    if receipt.get("revision") != bindings[role]["revision"] or receipt.get("binary_sha256") != bindings[role]["binary_sha256"]:
        fail(f"{lane}: receipt role identity differs")
    if receipt.get("cwd") != protocol["build_path"] or receipt.get("clean_before") is not True or receipt.get("clean_after") is not True or receipt.get("binary_unchanged") is not True or receipt.get("exit_code") != 0 or receipt.get("report_metadata_matches_clean_role") is not True:
        fail(f"{lane}: capture did not finish with clean identity")
    samples, warmups = ((analyze.PROBE_SAMPLES, analyze.PROBE_WARMUPS) if lane in analyze.LANES else (analyze.FULL_SAMPLES, analyze.FULL_WARMUPS) if lane in analyze.FULL_LANES else (analyze.HEAP_SAMPLES, analyze.HEAP_WARMUPS))
    if receipt.get("samples") != samples or receipt.get("warmups") != warmups:
        fail(f"{lane}: receipt samples/warmups differ")
    directory = root / lane
    expected = _expected_artifacts(lane in analyze.HEAP_LANES)
    artifacts = receipt.get("artifacts")
    if not isinstance(artifacts, dict) or set(artifacts) != expected:
        fail(f"{lane}: artifact set differs")
    for name, descriptor in artifacts.items():
        path = _regular(directory / name, f"{lane}/{name}")
        if not isinstance(descriptor, dict) or set(descriptor) != {"sha256", "bytes"}:
            fail(f"{lane}/{name}: malformed artifact descriptor")
        digest, size = sha256_file(path)
        if digest != descriptor.get("sha256") or size != descriptor.get("bytes"):
            fail(f"{lane}/{name}: receipt hash/size differs")
    catalog_reference = _verify_corpus_catalog(directory / "corpus-catalog.json", f"{lane}/corpus-catalog.json")
    report = read_json(directory / "report.json", f"{lane}/report.json")
    if report.get("corpus_catalog") != catalog_reference:
        fail(f"{lane}: report corpus catalog reference differs")
    _report_config(report, lane, protocol, samples, warmups)
    if report.get("tool") != {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": "litchi-perf-baseline",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": "none",
    }:
        fail(f"{lane}: benchmark tool identity differs")
    binary = report.get("binary_identity")
    environment = report.get("environment")
    if not isinstance(binary, dict) or binary.get("path") != bindings[role]["binary_path"] or binary.get("binary_sha256") != bindings[role]["binary_sha256"] or binary.get("binary_bytes") != bindings[role]["bytes"] or binary.get("profile") != "release" or binary.get("executable") is not True:
        fail(f"{lane}: report binary identity differs")
    if not isinstance(environment, dict) or environment.get("git_revision") != bindings[role]["revision"] or environment.get("git_worktree_dirty") is not False:
        fail(f"{lane}: report source identity differs")
    if lane in analyze.HEAP_LANES:
        export = read_json(directory / "heaptrack-print.json", f"{lane}/heaptrack-print.json")
        if not isinstance(export, dict) or not str(export.get("schema", "")).endswith("-heap-export-v1") or export.get("exit_code") != 0 or export.get("input_path") != "heaptrack.zst":
            fail(f"{lane}: Heaptrack export receipt is invalid")
        if export.get("input_sha256") != artifacts["heaptrack.zst"]["sha256"]:
            fail(f"{lane}: Heaptrack input hash differs")
        outputs = export.get("artifacts")
        if not isinstance(outputs, dict) or set(outputs) != {"heaptrack-print.stdout", "heaptrack-print.stderr"}:
            fail(f"{lane}: Heaptrack export artifacts are incomplete")
        for name, descriptor in outputs.items():
            path = _regular(directory / name, f"{lane}/{name}")
            digest, size = sha256_file(path)
            if not isinstance(descriptor, dict) or descriptor.get("sha256") != digest or descriptor.get("bytes") != size:
                fail(f"{lane}/{name}: export artifact binding differs")


def verify_gate_receipt(root: Path, candidate_manifest: Mapping[str, str]) -> None:
    receipt = read_json(root / "validation/receipt.json", "validation/receipt.json")
    if not isinstance(receipt, dict) or not isinstance(receipt.get("checks"), list):
        fail("validation/receipt.json: malformed gate receipt")
    checks = receipt["checks"]
    if {item.get("name") for item in checks if isinstance(item, dict)} != set(GATES) or len(checks) != len(GATES):
        fail("validation/receipt.json: gate set differs")
    for item in checks:
        if not isinstance(item, dict) or item.get("exit_code") != 0:
            fail("validation/receipt.json: a correctness gate failed")
        if item.get("source_hashes") != candidate_manifest:
            fail(f"validation/receipt.json: source hash inventory differs for {item.get('name')}")
        log = bundle_file(root / "validation", item.get("log"), f"validation/{item.get('name')}.log")
        digest, _ = sha256_file(log)
        if digest != item.get("log_sha256"):
            fail(f"validation/{item.get('name')}: log digest differs")


def verify_seal(root: Path) -> None:
    seal = _regular(root / "SHA256SUMS", "SHA256SUMS")
    entries: dict[str, str] = {}
    for number, line in enumerate(seal.read_text(encoding="utf-8").splitlines(), 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or HEX.fullmatch(fields[0]) is None:
            fail(f"SHA256SUMS:{number}: malformed line")
        name = _relative(fields[1], f"SHA256SUMS:{number}").as_posix()
        if name == "SHA256SUMS" or name in entries:
            fail(f"SHA256SUMS:{number}: duplicate/self entry")
        path = bundle_file(root, name, f"SHA256SUMS:{number}")
        digest, _ = sha256_file(path)
        if digest != fields[0]:
            fail(f"SHA256SUMS: digest differs for {name}")
        entries[name] = digest
    actual = {
        path.relative_to(root).as_posix()
        for path in root.rglob("*")
        if path.is_file() and not path.is_symlink() and path.name != "SHA256SUMS"
    }
    if set(entries) != actual:
        fail(f"SHA256SUMS: exact coverage differs (missing={sorted(actual - set(entries))}, extra={sorted(set(entries) - actual)})")


def verify_live_files(root: Path, bindings: Mapping[str, Mapping[str, Any]], protocol: Mapping[str, Any]) -> None:
    # The captures deliberately reuse one absolute checkout.  At any instant
    # it can represent only one role; authenticate both binaries, then audit
    # the source tree for whichever role is currently checked out.
    tree = Path(protocol["build_path"])
    try:
        revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=tree, text=True, stderr=subprocess.STDOUT).strip()
        dirty = subprocess.check_output(["git", "status", "--porcelain"], cwd=tree, text=True, stderr=subprocess.STDOUT).strip()
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"live build tree: cannot inspect checkout: {error}")
    if dirty:
        fail("live build tree is dirty")
    live_roles = [role for role, value in bindings.items() if value["revision"] == revision]
    if len(live_roles) != 1:
        fail("live build tree revision matches neither or both bound roles")
    live_role = live_roles[0]
    for role, value in bindings.items():
        binary = Path(value["binary_path"])
        digest, size = sha256_file(_regular(binary, f"{role}.binary_path"))
        if digest != value["binary_sha256"] or size != value["bytes"]:
            fail(f"{role}: live binary identity differs")
        if role != live_role:
            continue
        manifest = read_json(bundle_file(root, value["source_manifest"], f"{role}.source_manifest"), f"{role} manifest")
        for path, expected in manifest.items():
            digest, _ = sha256_file(_regular(tree / path, f"{role} source {path}"))
            if digest != expected:
                fail(f"{role}: live source differs at {path}")
        for path, expected in FIXTURES.items():
            digest, _ = sha256_file(_regular(tree / path, f"{role} fixture {path}"))
            if digest != expected:
                fail(f"{role}: live fixture differs at {path}")


def verify_bundle(root: Path = ROOT, *, live: bool = False) -> dict[str, Any]:
    root = root.resolve()
    protocol = verify_protocol(root)
    control, control_source = _role_binding(root, "control")
    candidate, candidate_source = _role_binding(root, "candidate")
    if control["revision"] == candidate["revision"] or control["binary_sha256"] == candidate["binary_sha256"]:
        fail("control and candidate identities must be distinct")
    changed = verify_source_pair(root, {"control": control, "candidate": candidate})
    bindings = {"control": control, "candidate": candidate}
    for lane in LANES:
        verify_lane(root, lane, bindings, protocol)
    verify_gate_receipt(root, read_json(bundle_file(root, candidate["source_manifest"], "candidate manifest"), "candidate manifest"))
    summary = read_json(root / "summary.json", "summary.json")
    try:
        expected = analyze.build_summary(root)
    except (analyze.AnalysisError, OSError, TypeError, ValueError, KeyError) as error:
        fail(f"summary replay failed: {error}")
    if not canonical_equal(summary, expected):
        fail("summary.json does not equal a fresh canonical replay")
    verify_seal(root)
    if live:
        verify_live_files(root, bindings, protocol)
    else:
        # A sealed bundle must be replayable without the temporary executables
        # or checkout that produced it.
        for role, value in bindings.items():
            if Path(value["binary_path"]).exists():
                fail(f"{role}: temporary binary still exists in flagless mode")
        if Path(protocol["build_path"]).exists():
            fail("temporary build tree still exists in flagless mode")
    return {"status": "pass", "changed_files": changed, "live": live}


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--live", action="store_true", help="also inspect the authenticated temporary binaries and trees")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify_bundle(args.root, live=args.live)
        print(json.dumps(result, sort_keys=True))
        return 0
    except (VerificationError, OSError, TypeError, ValueError, KeyError) as error:
        print(json.dumps({"status": "fail", "error": str(error)}, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
