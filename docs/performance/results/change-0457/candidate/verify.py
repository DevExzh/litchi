#!/usr/bin/env python3
"""Portable verifier for the 0457 source-backed ODP candidate capture.

The verifier authenticates the final candidate protocol, build/source/binary
bindings, retained source-custody receipts, all twelve successful reports, and
the independently retained source/output ZIP fixtures.  It accepts deleted
candidate binaries after capture; ``--precleanup`` additionally requires the
authenticated executable paths to still exist.  Failed attempts are retained
as evidence and never contribute samples.
"""

from __future__ import annotations

import argparse
from contextlib import contextmanager
import tempfile
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
PHASES = {"R1": (0, 6), "R2": (6, 12)}
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
OUTPUT_BINDING_SCHEMA = "litchi-0457-odp-source-tail-candidate-output-binding-v1"
NATIVE_OUTPUT_SCHEMA = "litchi-0457-native-odp-append-oracle-v1"
CAPTURE_SCHEMA = "litchi-0457-source-tail-candidate-capture-receipt-v1"


class VerifyError(AssertionError):
    pass


def fail(message: str) -> None:
    raise VerifyError(message)


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def canonical(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def integer(value: Any, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def sha_text(value: Any, label: str) -> str:
    value = text(value, label).lower()
    if len(value) != 64 or any(char not in "0123456789abcdef" for char in value):
        fail(f"{label}: expected lowercase SHA-256")
    return value


def artifact(base: Path, record: Any, label: str) -> Path:
    row = obj(record, label)
    raw_path = text(row.get("path"), f"{label}.path")
    relative = Path(raw_path)
    if relative.is_absolute() or ".." in relative.parts:
        fail(f"{label}: artifact path must be bundle-relative")
    path = (base / relative).resolve()
    if not path.is_file() or path.is_symlink() or not path.is_relative_to(base.resolve()):
        fail(f"{label}: artifact is missing or escapes the candidate bundle")
    if path.stat().st_size != integer(row.get("bytes"), f"{label}.bytes") or sha(path) != sha_text(row.get("sha256"), f"{label}.sha256"):
        fail(f"{label}: artifact identity differs")
    return path


def path_tail_matches(value: Any, relative: str, label: str) -> None:
    actual = Path(text(value, label))
    expected = Path(relative)
    if actual.name != expected.name or actual.parts[-len(expected.parts) :] != expected.parts:
        fail(f"{label}: captured absolute path does not bind {relative}")


def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    order = protocol.get("order")
    if not isinstance(order, list) or len(order) != 12:
        fail("protocol order must contain twelve lanes")
    first, last = PHASES[phase]
    lanes = [obj(item, f"protocol.order[{index}]") for index, item in enumerate(order[first:last], first)]
    expected_order = (
        [("normal", "tiny"), ("normal", "medium"), ("normal", "large"),
         ("allocator", "tiny"), ("allocator", "medium"), ("allocator", "large")]
        if phase == "R1" else
        [("allocator", "large"), ("allocator", "medium"), ("allocator", "tiny"),
         ("normal", "large"), ("normal", "medium"), ("normal", "tiny")]
    )
    if [(lane.get("mode"), lane.get("shape")) for lane in lanes] != expected_order:
        fail(f"{phase}: lane order differs from protocol")
    if any(lane.get("phase") != phase or lane.get("repeat") != phase for lane in lanes):
        fail(f"{phase}: phase/repeat identity differs")
    if any(lane.get("shape") not in SHAPES or lane.get("mode") not in {"normal", "allocator"} for lane in lanes):
        fail(f"{phase}: unsupported lane identity")
    return lanes


def ambient_empty(value: Any, label: str) -> None:
    if value != {}:
        fail(f"{label}: ambient allocator environment is not the authenticated empty observation")


def check_member(value: Any, label: str) -> dict[str, Any]:
    member = obj(value, label)
    path = text(member.get("path"), f"{label}.path")
    text(member.get("media_type"), f"{label}.media_type")
    method = text(member.get("compression_method"), f"{label}.compression_method").lower()
    if method not in {"store", "stored", "deflate", "deflated"}:
        fail(f"{label}: unsupported compression method")
    if not isinstance(member.get("data_descriptor"), bool):
        fail(f"{label}.data_descriptor: expected boolean")
    integer(member.get("crc32"), f"{label}.crc32")
    integer(member.get("decoded_bytes"), f"{label}.decoded_bytes")
    integer(member.get("compressed_bytes"), f"{label}.compressed_bytes")
    sha_text(member.get("decoded_sha256"), f"{label}.decoded_sha256")
    sha_text(member.get("compressed_sha256"), f"{label}.compressed_sha256")
    return member


def member_map(value: Any, label: str) -> dict[str, dict[str, Any]]:
    if not isinstance(value, list):
        fail(f"{label}: expected member array")
    result: dict[str, dict[str, Any]] = {}
    for index, raw in enumerate(value):
        member = check_member(raw, f"{label}[{index}]")
        path = member["path"]
        if path in result:
            fail(f"{label}: duplicate member {path}")
        result[path] = member
    expected = {"META-INF/manifest.xml", "Opaque/litchi-perf-odp-existing-append-opaque.bin", "content.xml", "meta.xml", "mimetype", "styles.xml"}
    if set(result) != expected:
        fail(f"{label}: member set differs from the six-member corpus")
    return result


@contextmanager
def import_native(path: Path):
    # Synthetic fixtures have their own bound archive/member identities. Replay
    # the authenticated oracle bytes without the unrelated native inventory,
    # whose repository-relative lookup assumes the original directory depth.
    with tempfile.TemporaryDirectory(prefix="litchi-0457-synthetic-oracle-") as directory:
        copied = Path(directory) / path.name
        copied.write_bytes(path.read_bytes())
        spec = importlib.util.spec_from_file_location("change0457_candidate_native_oracle", copied)
        if spec is None or spec.loader is None:
            fail(f"cannot load native output oracle {path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
        if not callable(getattr(module, "verify", None)):
            fail("native output oracle has no verify()")
        yield module


def title(index: int) -> str:
    variants = ("plain slide", "Unicode café Δ 中", 'entities <&> "quoted"', "mixed façade Ω <&> value")
    return f"litchi-perf-odp-buffered-title-{index:05d} {variants[index % 4]}"


def body(index: int) -> str:
    variants = ("plain slide", "Unicode café Δ 中", 'entities <&> "quoted"', "mixed façade Ω <&> value")
    return f"litchi-perf-odp-buffered-body-{index:05d} {variants[index % 4]}"


def check_output_binding() -> str:
    protocol = load(ROOT / "protocol.json", "candidate protocol")
    output_spec = obj(protocol.get("output_binding"), "protocol.output_binding")
    expected_sha = sha_text(output_spec.get("sha256"), "protocol.output_binding.sha256")
    binding_path = ROOT / text(output_spec.get("path"), "protocol.output_binding.path")
    if not binding_path.is_file() or binding_path.is_symlink() or sha(binding_path) != expected_sha:
        fail("candidate output binding is missing or differs from the final protocol hash")
    binding = obj(load(binding_path, str(binding_path)), "candidate output binding")
    if binding.get("schema") != OUTPUT_BINDING_SCHEMA:
        fail("candidate output binding schema differs")
    binder = obj(binding.get("binder"), "candidate output binding.binder")
    binder_relative = Path(text(binder.get("path"), "candidate output binding.binder.path"))
    if binder_relative.is_absolute() or ".." in binder_relative.parts:
        fail("candidate output binding helper path escapes the bundle")
    if binder.get("path") != output_spec.get("binder_path") or sha(ROOT / binder_relative) != sha_text(binder.get("sha256"), "candidate output binding.binder.sha256"):
        fail("candidate output binding helper identity differs")
    if output_spec.get("binder_sha256") != binder.get("sha256"):
        fail("candidate output binding helper does not match protocol")
    native_spec = obj(output_spec.get("native_oracle"), "protocol.output_binding.native_oracle")
    native_relative = text(native_spec.get("path"), "protocol.output_binding.native_oracle.path")
    if native_relative != "../native/verify-output.py":
        fail("candidate native output oracle path differs")
    native_path = (ROOT / native_relative).resolve()
    if not native_path.is_file() or sha(native_path) != sha_text(native_spec.get("sha256"), "protocol.output_binding.native_oracle.sha256"):
        fail("candidate native output oracle identity differs")
    binding_native = obj(binding.get("native_oracle"), "candidate output binding.native_oracle")
    if binding_native.get("path") != native_relative or binding_native.get("sha256") != native_spec.get("sha256"):
        fail("candidate output binding native oracle differs from protocol")
    with import_native(native_path) as native:
        corpus_bindings = load(ROOT / "corpus-bindings.json", "candidate corpus bindings")
        shapes = obj(binding.get("shapes"), "candidate output binding.shapes")
        control_shapes = obj(corpus_bindings.get("shapes"), "candidate corpus bindings.shapes")
        if set(shapes) != set(SHAPES):
            fail("candidate output binding must contain every shape exactly once")
        for shape, count in SHAPES.items():
            item = obj(shapes.get(shape), f"candidate output binding.shapes.{shape}")
            request = obj(item.get("request"), f"candidate output binding.shapes.{shape}.request")
            expected_request = {"title": title(count), "body": body(count), "name": f"page{count + 1}"}
            if request != expected_request:
                fail(f"{shape}: output binding request differs")
            control = obj(control_shapes.get(shape), f"candidate corpus bindings.shapes.{shape}")
            control_source = obj(control.get("source"), f"candidate corpus bindings.shapes.{shape}.source")
            source = obj(item.get("source"), f"candidate output binding.shapes.{shape}.source")
            output = obj(item.get("output"), f"candidate output binding.shapes.{shape}.output")
            if source.get("archive_sha256") != control_source.get("archive_sha256") or source.get("archive_bytes") != control_source.get("archive_bytes") or source.get("content_xml_sha256") != control_source.get("content_xml_sha256") or source.get("content_xml_bytes") != control_source.get("content_xml_bytes"):
                fail(f"{shape}: bound source archive/content differs from control")
            source_members = member_map(source.get("members"), f"candidate output binding.shapes.{shape}.source.members")
            output_members = member_map(output.get("members"), f"candidate output binding.shapes.{shape}.output.members")
            if source.get("members") != control_source.get("members"):
                fail(f"{shape}: bound source members differ from control")
            for member_name in source_members:
                if member_name != "content.xml" and source_members[member_name] != output_members[member_name]:
                    fail(f"{shape}: untouched output member differs from source: {member_name}")
            fixtures = obj(item.get("fixtures"), f"candidate output binding.shapes.{shape}.fixtures")
            fixture_paths: dict[str, Path] = {}
            for side in ("source", "output"):
                record = obj(fixtures.get(side), f"candidate output binding.shapes.{shape}.fixtures.{side}")
                fixture_paths[side] = artifact(ROOT, record, f"candidate output binding.shapes.{shape}.fixtures.{side}")
                expected_archive = source if side == "source" else output
                if record.get("sha256") != expected_archive.get("archive_sha256") or record.get("bytes") != expected_archive.get("archive_bytes"):
                    fail(f"{shape}: retained {side} fixture differs from binding")
            native_result = obj(item.get("native_result"), f"candidate output binding.shapes.{shape}.native_result")
            if native_result.get("schema") != NATIVE_OUTPUT_SCHEMA or native_result.get("status") != "validated":
                fail(f"{shape}: native output receipt is not validated")
            if sha_text(item.get("native_result_sha256"), f"{shape}.native_result_sha256") != hashlib.sha256(canonical(native_result)).hexdigest():
                fail(f"{shape}: native output receipt hash differs")
            native_source = obj(native_result.get("source"), f"{shape}.native_result.source")
            native_output = obj(native_result.get("output"), f"{shape}.native_result.output")
            native_content = obj(native_result.get("content"), f"{shape}.native_result.content")
            if native_source.get("path") != fixtures["source"].get("path") or native_output.get("path") != fixtures["output"].get("path"):
                fail(f"{shape}: native output receipt fixture paths differ")
            if native_source.get("sha256") != source.get("archive_sha256") or native_source.get("bytes") != source.get("archive_bytes") or native_output.get("sha256") != output.get("archive_sha256") or native_output.get("bytes") != output.get("archive_bytes"):
                fail(f"{shape}: native output receipt archive identities differ")
            if native_content.get("source_content_xml_sha256") != source.get("content_xml_sha256") or native_content.get("source_content_xml_bytes") != source.get("content_xml_bytes") or native_content.get("output_content_xml_sha256") != output.get("content_xml_sha256") or native_content.get("output_content_xml_bytes") != output.get("content_xml_bytes"):
                fail(f"{shape}: native output receipt content identities differ")
            try:
                replay = native.verify(fixture_paths["source"], fixture_paths["output"], request["title"], request["body"], request["name"])
            except Exception as error:
                fail(f"{shape}: native output oracle replay failed: {error}")
            replay["source"]["path"] = fixtures["source"]["path"]
            replay["output"]["path"] = fixtures["output"]["path"]
            if replay != native_result:
                fail(f"{shape}: native output receipt does not reproduce from retained fixtures")
        return expected_sha


def check_bindings(protocol_path: Path, require_binaries: bool, repo_root: Path | None) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any], str]:
    protocol = obj(load(protocol_path, str(protocol_path)), "candidate protocol")
    if protocol.get("schema") != "litchi-0457-odp-source-tail-candidate-v1" or protocol.get("change") != 457:
        fail("protocol schema/change differs")
    if protocol.get("capture_driver_sha256") != sha(ROOT / "capture.py"):
        fail("capture driver hash differs from protocol")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1 or protocol.get("samples") != 30 or protocol.get("warmups") != 3:
        fail("protocol dimensions differ")
    if protocol.get("selector") != "odp_source_tail_append_lifecycle" or protocol.get("shapes") != SHAPES:
        fail("protocol selector or shape matrix differs")
    environment = obj(protocol.get("environment"), "protocol.environment")
    ambient_empty(environment.get("allocator_ambient"), "protocol.environment.allocator_ambient")
    if environment.get("rustflags") is not None or "RUSTFLAGS" in obj(environment.get("fixed"), "protocol.environment.fixed"):
        fail("candidate protocol must make no RUSTFLAGS build claim")
    role = obj(obj(protocol.get("roles"), "protocol.roles").get("candidate"), "protocol.roles.candidate")
    build_path = ROOT / text(obj(role.get("build_receipt"), "candidate build receipt role").get("path"), "candidate build receipt path")
    source_path = ROOT / text(obj(role.get("source_manifest"), "candidate source manifest role").get("path"), "candidate source manifest path")
    binary_path = ROOT / text(obj(role.get("binary_binding"), "candidate binary binding role").get("path"), "candidate binary binding path")
    build = obj(load(build_path, str(build_path)), "candidate build receipt")
    source_map = obj(load(source_path, str(source_path)), "candidate source manifest")
    if build.get("change") != 457 or build.get("status") != "pass" or build.get("source_unchanged") is not True:
        fail("candidate build receipt is not a successful source-unchanged build")
    revision = text(build.get("revision"), "candidate build revision")
    source_digest = sha(source_path)
    source_record = {"path": f"sources/{source_digest}.json", "sha256": source_digest, "files": len(source_map)}
    if build.get("source_before") != source_record or build.get("source_after") != source_record:
        fail("candidate build/source custody records differ")
    # The frozen candidate protocol binds the source-manifest path through the
    # role and binary-binding records; unlike the older control protocol it
    # does not duplicate an expected_source_manifest object under custody.
    # Honor that optional duplicate when present, while accepting the actual
    # candidate protocol shape.
    expected_protocol_source = protocol.get("custody", {}).get("expected_source_manifest")
    if expected_protocol_source is not None and expected_protocol_source != source_record:
        fail("candidate protocol source custody record differs")
    bindings = obj(load(binary_path, str(binary_path)), "candidate binary bindings")
    if bindings.get("schema") != "litchi-0457-source-tail-candidate-binary-binding-v1":
        fail("candidate binary binding schema differs")
    expected_build_record = {"path": str(build_path.relative_to(ROOT)), "sha256": sha(build_path)}
    expected_source_record = {"path": str(source_path.relative_to(ROOT)), "sha256": source_digest, "files": len(source_map)}
    if bindings.get("build_receipt") != expected_build_record or bindings.get("source_manifest") != expected_source_record:
        fail("candidate binary binding receipt/source identity differs")
    if repo_root is None:
        repo_root = Path(build.get("cwd")) if isinstance(build.get("cwd"), str) else None
    if repo_root is not None:
        repo_root = repo_root.resolve()
    for mode in ("normal", "allocator"):
        item = obj(obj(bindings.get("binaries"), "candidate binary bindings.binaries").get(mode), f"candidate binary bindings.binaries.{mode}")
        if item.get("profile") != "release" or item.get("executable") is not True:
            fail(f"{mode} binary binding identity differs")
        integer(item.get("bytes"), f"{mode} binary bytes", minimum=1)
        sha_text(item.get("sha256"), f"{mode} binary sha256")
        source_name = text(item.get("source_path"), f"{mode} binary source_path")
        if Path(source_name).is_absolute() or ".." in Path(source_name).parts:
            fail(f"{mode} binary source_path is not repository-relative")
        copy_name = text(item.get("copy_path"), f"{mode} binary copy_path")
        if require_binaries:
            if repo_root is None:
                fail("--precleanup requires --repo-root or an absolute build cwd")
            source_binary = repo_root / source_name
            copy_binary = Path(copy_name)
            for candidate, label in ((source_binary, f"{mode} source binary"), (copy_binary, f"{mode} copied binary")):
                if not candidate.is_file() or candidate.is_symlink() or candidate.stat().st_size != item["bytes"] or sha(candidate) != item["sha256"]:
                    fail(f"{label} differs from binary binding")
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    for key, hash_key in (("verifier_path", "verifier_sha256"), ("protocol_path", "protocol_sha256"), ("corpus_bindings_path", "corpus_bindings_sha256")):
        path = ROOT / text(oracle.get(key), f"protocol.oracle.{key}")
        if not path.is_file() or sha(path) != sha_text(oracle.get(hash_key), f"protocol.oracle.{hash_key}"):
            fail(f"candidate oracle {key} differs from protocol")
    custody = obj(protocol.get("custody"), "protocol.custody")
    custody_path = (ROOT / text(custody.get("driver_path"), "protocol.custody.driver_path")).resolve()
    if not custody_path.is_file() or sha(custody_path) != sha_text(custody.get("driver_sha256"), "protocol.custody.driver_sha256"):
        fail("candidate custody driver differs from protocol")
    output_binding_sha = check_output_binding()
    build_cwd = text(build.get("cwd"), "candidate build cwd")
    cwd_path = Path(build_cwd)
    if not cwd_path.is_absolute() or cwd_path != cwd_path.resolve():
        fail("candidate build cwd must be an absolute canonical path")
    return protocol, role | {"revision": revision}, bindings, source_record, output_binding_sha


def check_report(report_path: Path, receipt: dict[str, Any], lane: dict[str, Any], protocol: dict[str, Any], revision: str) -> None:
    report = obj(load(report_path, str(report_path)), str(report_path))
    binary = obj(receipt.get("binary"), f"{report_path}.receipt.binary")
    identity = obj(report.get("binary_identity"), f"{report_path}.binary_identity")
    if identity.get("path") != binary.get("copy_path") or identity.get("binary_sha256") != binary.get("sha256") or identity.get("binary_bytes") != binary.get("bytes") or identity.get("profile") != "release" or identity.get("executable") is not True:
        fail(f"{report_path}: binary identity differs from receipt")
    environment = obj(report.get("environment"), f"{report_path}.environment")
    if environment.get("git_revision") != revision or environment.get("rustflags") is not None:
        fail(f"{report_path}: revision or null rustflags identity differs")
    configuration = obj(report.get("configuration"), f"{report_path}.configuration")
    if configuration.get("samples_per_case") != protocol["samples"] or configuration.get("warmup_iterations_per_case") != protocol["warmups"] or configuration.get("execution_workers") != [protocol["workers"]] or configuration.get("cases") != [protocol["selector"]] or configuration.get("semantic_shapes") != [lane["shape"]]:
        fail(f"{report_path}: execution configuration differs")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        fail(f"{report_path}: expected one result")
    result = obj(results[0], f"{report_path}.results[0]")
    if result.get("case") != protocol["selector"] or obj(result.get("corpus"), f"{report_path}.corpus").get("shape") != lane["shape"]:
        fail(f"{report_path}: selector or shape differs")


def path_arg_matches(value: Any, artifact_record: Any, label: str) -> None:
    record = obj(artifact_record, label)
    path_tail_matches(value, text(record.get("path"), f"{label}.path"), label)


def check_argv(receipt: dict[str, Any], lane: dict[str, Any], protocol: dict[str, Any]) -> None:
    argv = receipt.get("argv")
    if not isinstance(argv, list) or len(argv) < 8:
        fail(f"{receipt.get('name')}: workload argv is malformed")
    cli = obj(protocol.get("workload_cli"), "protocol.workload_cli")
    binary = obj(receipt.get("binary"), f"{receipt.get('name')}.binary")
    if argv[0:7] != ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o", argv[6]] or argv[7] != binary.get("copy_path"):
        fail(f"{receipt.get('name')}: workload launcher/binary identity differs")
    if not isinstance(argv[6], str):
        fail(f"{receipt.get('name')}: resource path is malformed")
    expected = [cli["case_flag"], protocol["selector"], cli["shape_flag"], lane["shape"], cli["workers_flag"], str(protocol["workers"]), cli["samples_flag"], str(protocol["samples"]), cli["warmup_flag"], str(protocol["warmups"]), cli["report_flag"], None, cli["catalog_flag"], None]
    actual = argv[8:]
    if len(actual) != len(expected):
        fail(f"{receipt.get('name')}: workload argument count differs")
    for index, (got, want) in enumerate(zip(actual, expected)):
        if want is not None and got != want:
            fail(f"{receipt.get('name')}: workload argument {index} differs")
    artifacts = obj(receipt.get("artifacts"), f"{receipt.get('name')}.artifacts")
    path_arg_matches(argv[6], artifacts.get("resource_log"), f"{receipt.get('name')}.resource argv")
    path_arg_matches(actual[11], artifacts.get("report"), f"{receipt.get('name')}.report argv")
    path_arg_matches(actual[13], artifacts.get("catalog"), f"{receipt.get('name')}.catalog argv")
    oracle_argv = receipt.get("oracle_argv")
    if not isinstance(oracle_argv, list) or len(oracle_argv) != 9:
        fail(f"{receipt.get('name')}: oracle argv is malformed")
    if not text(oracle_argv[0], f"{receipt.get('name')}.oracle_argv[0]") or oracle_argv[1] != "-B" or not Path(oracle_argv[2]).as_posix().endswith("oracle/verify-report.py"):
        fail(f"{receipt.get('name')}: oracle executable differs")
    if oracle_argv[3] != "--report" or oracle_argv[5] != "--mode" or oracle_argv[7] != "--shape" or oracle_argv[6] != lane["mode"] or oracle_argv[8] != lane["shape"]:
        fail(f"{receipt.get('name')}: oracle flags differ")
    path_arg_matches(oracle_argv[4], artifacts.get("report"), f"{receipt.get('name')}.oracle report argv")


def check_phase(protocol: dict[str, Any], protocol_sha: str, role: dict[str, Any], bindings: dict[str, Any], expected_source: dict[str, Any], output_binding_sha: str, phase: str, attempt: str) -> int:
    directory = ROOT / "runs" / phase / attempt
    state_path = directory / "capture-state.json"
    state = obj(load(state_path, str(state_path)), str(state_path))
    if state.get("status") != "pass" or state.get("phase") != phase or state.get("attempt") != attempt or state.get("completed_lanes") != 6 or state.get("protocol_sha256") != protocol_sha or state.get("driver_sha256") != sha(ROOT / "capture.py") or state.get("output_binding_sha256") != output_binding_sha:
        fail(f"{phase}: capture state is not a successful bound phase")
    if state.get("source_before") != expected_source or state.get("source_after") != expected_source or state.get("source_unchanged") is not True or state.get("outside_bundle_status_unchanged") is not True:
        fail(f"{phase}: source/outside-bundle custody differs")
    ambient_empty(state.get("allocator_environment"), f"{phase}.allocator_environment")
    ambient_empty(state.get("allocator_environment_after"), f"{phase}.allocator_environment_after")
    lanes = expected_lanes(protocol, phase)
    expected_index = [str((directory / f"{phase}-{lane['mode']}-{lane['shape']}-{lane['repeat'].lower()}-receipt.json").relative_to(ROOT)) for lane in lanes]
    if state.get("index") != expected_index:
        fail(f"{phase}: capture index differs")
    for lane in lanes:
        name = f"{phase}-{lane['mode']}-{lane['shape']}-{lane['repeat'].lower()}"
        receipt_path = directory / f"{name}-receipt.json"
        receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
        if receipt.get("schema") != CAPTURE_SCHEMA or receipt.get("change") != 457 or receipt.get("status") != "pass" or receipt.get("phase") != phase or receipt.get("attempt") != attempt or receipt.get("name") != name or receipt.get("lane") != lane or receipt.get("selector") != protocol["selector"] or receipt.get("revision") != role["revision"]:
            fail(f"{name}: receipt identity differs")
        if receipt.get("protocol_sha256") != protocol_sha or receipt.get("driver_sha256") != sha(ROOT / "capture.py") or receipt.get("output_binding_sha256") != output_binding_sha or receipt.get("binary_binding_sha256") != sha(ROOT / "binary-bindings.json"):
            fail(f"{name}: receipt binding differs")
        if receipt.get("cwd") != role["build_cwd"] or receipt.get("source_before") != expected_source or receipt.get("source_after") != expected_source or receipt.get("source_unchanged") is not True:
            fail(f"{name}: source/cwd custody differs")
        ambient_empty(receipt.get("allocator_environment"), f"{name}.allocator_environment")
        ambient_empty(receipt.get("allocator_environment_after"), f"{name}.allocator_environment_after")
        binary = obj(obj(bindings.get("binaries"), "candidate binary bindings.binaries").get(lane["mode"]), f"{name}.binary binding")
        if receipt.get("binary") != binary:
            fail(f"{name}: binary binding differs")
        artifacts = obj(receipt.get("artifacts"), f"{name}.artifacts")
        required = {"report", "catalog", "workload_log", "resource_log", "oracle_log"}
        if set(artifacts) != required:
            fail(f"{name}: artifact set differs")
        report_path = artifact(ROOT, artifacts["report"], f"{name}.report")
        for key in required - {"report"}:
            artifact(ROOT, artifacts[key], f"{name}.{key}")
        check_argv(receipt, lane, protocol)
        if receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
            fail(f"{name}: workload/oracle did not pass")
        check_report(report_path, receipt, lane, protocol, role["revision"])
        oracle_result = subprocess.run([sys.executable, "-B", str(ROOT / protocol["oracle"]["verifier_path"]), "--report", str(report_path), "--mode", lane["mode"], "--shape", lane["shape"]], capture_output=True, text=True)
        if oracle_result.returncode != 0 or oracle_result.stdout.strip() != protocol["oracle"]["success_stdout"]:
            fail(f"{name}: independent candidate oracle rejected report: {oracle_result.stderr.strip()}")
    return len(lanes) * protocol["samples"]


def inspect_failed_attempts(phase: str, selected: str) -> list[dict[str, Any]]:
    phase_root = ROOT / "runs" / phase
    if not phase_root.is_dir():
        fail(f"{phase}: runs directory is missing")
    failed: list[dict[str, Any]] = []
    for directory in sorted(path for path in phase_root.iterdir() if path.is_dir()):
        if directory.name == selected:
            continue
        state_path = directory / "capture-state.json"
        if not state_path.is_file():
            fail(f"{phase}/{directory.name}: unrecognized attempt without capture state")
        state = obj(load(state_path, str(state_path)), str(state_path))
        if state.get("status") != "failed":
            fail(f"{phase}/{directory.name}: unselected attempt is not an authenticated failure")
        if state.get("phase") != phase or state.get("attempt") != directory.name or state.get("completed_lanes", 0) >= 6:
            fail(f"{phase}/{directory.name}: failed attempt completion identity differs")
        if state.get("source_unchanged") is not True or state.get("outside_bundle_status_unchanged") is not True:
            fail(f"{phase}/{directory.name}: failed attempt custody is not unchanged")
        ambient_empty(state.get("allocator_environment"), f"{phase}/{directory.name}.allocator_environment")
        ambient_empty(state.get("allocator_environment_after"), f"{phase}/{directory.name}.allocator_environment_after")
        index = state.get("index")
        if not isinstance(index, list) or len(index) != state.get("completed_lanes"):
            fail(f"{phase}/{directory.name}: failed attempt index differs")
        for relative in index:
            receipt_path = ROOT / text(relative, f"{phase}/{directory.name}.index")
            if receipt_path.parent != directory:
                fail(f"{phase}/{directory.name}: failed receipt escapes attempt")
            receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
            if receipt.get("phase") != phase or receipt.get("attempt") != directory.name:
                fail(f"{phase}/{directory.name}: failed receipt identity differs")
            if receipt.get("status") not in {"pass", "failed"}:
                fail(f"{phase}/{directory.name}: failed receipt status is unsupported")
            if receipt.get("status") == "failed" and not text(receipt.get("error"), f"{receipt_path}.error"):
                fail(f"{phase}/{directory.name}: failed receipt has no error")
            if receipt.get("source_unchanged") is not True:
                fail(f"{phase}/{directory.name}: failed receipt source custody differs")
            artifacts = receipt.get("artifacts", {})
            if not isinstance(artifacts, dict):
                fail(f"{phase}/{directory.name}: failed receipt artifacts are malformed")
            for key, record in artifacts.items():
                artifact(ROOT, record, f"{receipt_path}.{key}")
        failed.append({"phase": phase, "attempt": directory.name, "status": "authenticated-failure", "completed_lanes": state.get("completed_lanes"), "receipt_count": len(index)})
    return failed


def verify(protocol_path: Path, attempt: str, require_binaries: bool, repo_root: Path | None) -> dict[str, Any]:
    protocol_path = protocol_path.resolve()
    if protocol_path.parent != ROOT:
        fail("portable verifier requires the protocol inside the candidate bundle")
    protocol, role, bindings, expected_source, output_binding_sha = check_bindings(protocol_path, require_binaries, repo_root)
    protocol_sha = sha(protocol_path)
    role["build_cwd"] = text(load(ROOT / protocol["roles"]["candidate"]["build_receipt"]["path"], "candidate build receipt").get("cwd"), "candidate build cwd")
    failed_attempts = inspect_failed_attempts("R1", attempt) + inspect_failed_attempts("R2", attempt)
    samples = sum(check_phase(protocol, protocol_sha, role, bindings, expected_source, output_binding_sha, phase, attempt) for phase in ("R1", "R2"))
    if samples != 360:
        fail(f"verified sample count is {samples}, expected 360")
    return {
        "status": "pass",
        "change": 457,
        "selector": protocol["selector"],
        "attempt": attempt,
        "phases": 2,
        "reports": 12,
        "samples": samples,
        "protocol_sha256": protocol_sha,
        "binary_binding_sha256": sha(ROOT / "binary-bindings.json"),
        "output_binding_sha256": output_binding_sha,
        "failed_attempts": failed_attempts,
        "runtime_proof_contract": {
            "status": "verified",
            "binding": "each retained runtime proof source-version ID/revision vector equals the aligned publication-report vector",
            "fixture_scalar": "proof_source_version_id/proof_source_version_revision is retained as untimed fixture metadata and is not compared with timed provider IDs",
        },
        "scope": "specialized source-backed ODP publication-plan candidate only; no ordinary Commit/Patch comparison claim",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--precleanup", action="store_true", help="require authenticated executable paths to still exist")
    args = parser.parse_args(argv)
    try:
        result = verify(args.protocol, args.attempt, args.precleanup, args.repo_root)
        print(json.dumps(result, ensure_ascii=False, sort_keys=True))
        return 0
    except (OSError, KeyError, TypeError, ValueError, VerifyError, subprocess.SubprocessError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
