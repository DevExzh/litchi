#!/usr/bin/env python3
"""Fail-closed verifier for the 0468 XLSX CPU attribution bundle.

The capture is diagnostic evidence.  This verifier authenticates the frozen
protocol, source/build bindings, four capture lanes, corpus/catalog identity,
counter and output envelopes, compressed perf artifacts, and the derived
frame-pointer summary.  It never builds, invokes ``perf``, or treats sampled
periods as elapsed phase time.

``--preseal`` is used while the authenticated profile binary and clean source
tree still exist.  Normal verification is the portable post-cleanup check:
the live temporary tree/binary must be absent and ``SHA256SUMS`` must cover
every bundle file exactly.
"""

from __future__ import annotations

import argparse
import copy
import gzip
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import statistics
import subprocess
import sys
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
CHANGE = 468
SCHEMA = "litchi-0468-verification-v1"
CASE = "xlsx_one_percent_commit_save"
CORPUS_SHA256 = "5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714"
CATALOG_SHA256 = "db2c2a61e9c0bdd4660860b79327db141fce42dc0a9819f2ff01c995cf10dafa"
CONTENT_SET_SHA256 = "66889d1124b43cd8ec4580ba0675509d97419df2d10e0d8a6130a3a509076936"
REVISION = "933cb6b80eaed36af21f4dce984bd6b896d3543e"
LANES = {
    "normal-r1": (30, 3),
    "counters": (30, 3),
    "samples-fp": (50, 3),
    "normal-r2": (30, 3),
}
LANE_ORDER = tuple(LANES)
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
EXPECTED_ENV = {
    "RUSTUP_TOOLCHAIN": "1.98.1",
    "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "CARGO_BUILD_JOBS": "4",
    "CARGO_INCREMENTAL": "0",
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
}
REQUIRED_COUNTERS = (
    "cycles:u",
    "instructions:u",
    "branches:u",
    "branch-misses:u",
    "cache-misses:u",
    "page-faults",
)
EXPECTED_TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}


class VerificationError(ValueError):
    """Raised for any missing, malformed, or mismatched evidence."""


def fail(message: str) -> None:
    raise VerificationError(message)


def regular(path: Path, label: str) -> Path:
    if path.is_symlink() or not path.is_file():
        fail(f"{label}: missing, symlinked, or non-regular file")
    return path


def sha256_file(path: Path, *, gzip_stream: bool = False) -> tuple[str, int]:
    digest = hashlib.sha256()
    try:
        stream = gzip.open(path, "rb") if gzip_stream else path.open("rb")
        with stream:
            size = 0
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest(), size


def read_json(path: Path, label: str) -> Any:
    regular(path, label)
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream, parse_constant=lambda value: (_ for _ in ()).throw(
                ValueError(f"non-finite JSON constant {value}")
            ))
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label}: invalid JSON ({error})")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty string")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label).lower()
    if SHA256_RE.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def relative(value: Any, label: str) -> Path:
    raw = text(value, label)
    path = Path(raw)
    if path.is_absolute() or path.as_posix() != raw or ".." in path.parts:
        fail(f"{label}: expected a relative traversal-free path")
    return path


def bundle_file(root: Path, value: Any, label: str) -> Path:
    path = root / relative(value, label)
    try:
        if not path.resolve(strict=True).is_relative_to(root.resolve()):
            fail(f"{label}: path escapes bundle")
    except OSError as error:
        fail(f"{label}: cannot resolve ({error})")
    return regular(path, label)


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, ensure_ascii=False, allow_nan=False,
                          sort_keys=True, separators=(",", ":")).encode("utf-8")
    except (TypeError, ValueError) as error:
        fail(f"cannot canonicalize JSON: {error}")
    raise AssertionError("unreachable")


def canonical_equal(left: Any, right: Any) -> bool:
    return canonical(left) == canonical(right)


def _counter_signature(value: Any, key: str = "") -> Any:
    """Compare counter meaning across lanes while allowing sample counts to differ."""
    if isinstance(value, dict):
        if "values" in value and "status" in value and "scope" in value:
            values = value.get("values")
            result = {name: _counter_signature(item, name)
                      for name, item in value.items() if name != "values"}
            if isinstance(values, list):
                result["constant_value"] = values[0] if values else None
                result["all_values_constant"] = bool(values) and len(set(values)) == 1
            return result
        return {
            name: _counter_signature(item, name)
            for name, item in value.items()
            if name not in {"sample_count", "sample_indices"}
        }
    if isinstance(value, list):
        return [_counter_signature(item, key) for item in value]
    return value


def file_binding(path: Path, root: Path, *, decompressed: bool = False) -> dict[str, Any]:
    compressed_sha, compressed_bytes = sha256_file(path)
    if decompressed:
        plain_sha, plain_bytes = sha256_file(path, gzip_stream=True)
    else:
        plain_sha, plain_bytes = compressed_sha, compressed_bytes
    try:
        name = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        if path.parent.name == "tools":
            name = f"tools/{path.name}"
        elif path.parent.name == "change-0466":
            name = f"../change-0466/{path.name}"
        else:
            name = path.name
    return {
        "path": name,
        "sha256": plain_sha,
        "bytes": plain_bytes,
        "compressed_sha256": compressed_sha,
        "compressed_bytes": compressed_bytes,
    }


def _load_perf_compare(root: Path) -> Any:
    candidates = [root / "tools" / "perf_compare.py"]
    # The normal checkout layout is useful before packaging; the first path is
    # the only one used by a portable bundle.
    candidates.append(root.parents[3] / "tools" / "perf_compare.py")
    for path in candidates:
        if path.is_file():
            spec = importlib.util.spec_from_file_location("litchi_0468_perf_compare", path)
            if spec is None or spec.loader is None:
                continue
            module = importlib.util.module_from_spec(spec)
            sys.modules[spec.name] = module
            try:
                spec.loader.exec_module(module)
            except (OSError, SyntaxError, TypeError, ValueError) as error:
                fail(f"perf_compare.py: cannot load ({error})")
            return module, path
    fail("tools/perf_compare.py: dependency missing")


def _load_analyzer(root: Path) -> tuple[Any, Path]:
    path = root / "analyze.py"
    regular(path, "analyze.py")
    spec = importlib.util.spec_from_file_location("litchi_0468_analyze_verify", path)
    if spec is None or spec.loader is None:
        fail("analyze.py: cannot load parser")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, SyntaxError, TypeError, ValueError) as error:
        fail(f"analyze.py: cannot load parser ({error})")
    return module, path


def _verify_sha256sums(root: Path) -> int:
    sums = regular(root / "SHA256SUMS", "SHA256SUMS")
    rows: dict[str, str] = {}
    try:
        lines = sums.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"SHA256SUMS: cannot read ({error})")
    for number, line in enumerate(lines, 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"SHA256SUMS: malformed line {number}")
        name = relative(fields[1], f"SHA256SUMS line {number}").as_posix()
        if name == "SHA256SUMS" or name in rows:
            fail(f"SHA256SUMS: duplicate or self entry {name}")
        member = bundle_file(root, name, f"SHA256SUMS.{name}")
        if sha256_file(member)[0] != fields[0]:
            fail(f"SHA256SUMS: hash differs for {name}")
        rows[name] = fields[0]
    actual = set()
    for member in root.rglob("*"):
        if member.is_symlink():
            fail(f"bundle contains symlink {member.relative_to(root)}")
        if member.is_file() and member != sums:
            actual.add(member.relative_to(root).as_posix())
    if set(rows) != actual:
        fail("SHA256SUMS: exact coverage differs")
    return len(rows)


def _verify_protocol(root: Path) -> dict[str, Any]:
    protocol = obj(read_json(root / "protocol.json", "protocol.json"), "protocol.json")
    if protocol.get("schema") != "litchi-0468-profile-protocol-v1" or protocol.get("status") != "frozen":
        fail("protocol.json: schema/status differs")
    if protocol.get("revision") != REVISION or protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol.json: revision or execution identity differs")
    if protocol.get("corpus_sha256") != CORPUS_SHA256 or protocol.get("order") != list(LANE_ORDER):
        fail("protocol.json: corpus or lane order differs")
    if protocol.get("capture_driver_sha256") != sha256_file(root / "capture.py")[0]:
        fail("protocol.json: capture driver hash differs")
    lanes = obj(protocol.get("lanes"), "protocol.lanes")
    if set(lanes) != set(LANES):
        fail("protocol.lanes: lane set differs")
    for lane, (samples, warmups) in LANES.items():
        config = obj(lanes.get(lane), f"protocol.lanes.{lane}")
        if config.get("samples") != samples or config.get("warmups") != warmups:
            fail(f"protocol.lanes.{lane}: sample configuration differs")
    return protocol


def _verify_source_binding(root: Path, protocol: Mapping[str, Any]) -> dict[str, Any]:
    binding = obj(read_json(root / "source-binding.json", "source-binding.json"), "source-binding.json")
    if binding.get("schema") != "litchi-0468-source-binding-v1":
        fail("source-binding.json: schema differs")
    if binding.get("revision") != protocol["revision"] or binding.get("files") != 6992:
        fail("source-binding.json: revision/file count differs")
    manifest = regular(root / "sources.json", "sources.json")
    if digest(binding.get("source_manifest_sha256"), "source-binding.source_manifest_sha256") != sha256_file(manifest)[0]:
        fail("source-binding.json: source manifest hash differs")
    source_map = obj(read_json(manifest, "sources.json"), "sources.json")
    if len(source_map) != binding["files"]:
        fail("sources.json: source file count differs from binding")
    for name, expected in source_map.items():
        # The manifest names files in the authenticated sparse source tree;
        # those files are not duplicated inside the portable evidence bundle.
        relative(name, f"sources.json.{name}")
        digest(expected, f"sources.json.{name}")
    fixtures = obj(binding.get("included_fixtures"), "source-binding.included_fixtures")
    if set(fixtures) != {
        "test-data/poi/test-data/spreadsheet/54016.xls",
        "test-data/rtf/watermark.rtf",
    }:
        fail("source-binding: compile fixture set differs")
    for name, expected in fixtures.items():
        digest(expected, f"source-binding.included_fixtures.{name}")
    return binding


def _verify_command_receipt(root: Path, path: Path, label: str, expected_driver: str, *, started: bool = False, gzip_stdout: bool = False) -> dict[str, Any]:
    receipt = obj(read_json(path, label), label)
    if receipt.get("schema") != "litchi-0468-command-v1":
        fail(f"{label}: command receipt schema differs")
    if not started and receipt.get("exit_code") != 0:
        fail(f"{label}: command did not complete successfully")
    if receipt.get("driver_sha256") != expected_driver:
        fail(f"{label}: driver identity differs")
    argv = receipt.get("argv")
    if not isinstance(argv, list) or not argv or any(not isinstance(x, str) for x in argv):
        fail(f"{label}.argv: malformed")
    if not isinstance(receipt.get("cwd"), str) or not receipt["cwd"]:
        fail(f"{label}.cwd: missing")
    environment = obj(receipt.get("environment"), f"{label}.environment")
    if environment != EXPECTED_ENV:
        fail(f"{label}.environment: build environment differs")
    if not started:
        outputs = obj(receipt.get("outputs"), f"{label}.outputs")
        prefix = path.name.removesuffix(".json")
        for suffix in (".stdout", ".stderr"):
            output = obj(outputs.get(suffix), f"{label}.outputs.{suffix}")
            output_path = path.with_name(prefix + suffix)
            compressed_output = gzip_stdout and suffix == ".stdout" and not output_path.exists()
            measured_path = path.with_name(prefix + suffix + ".gz") if compressed_output else output_path
            measured_sha, measured_bytes = sha256_file(measured_path, gzip_stream=compressed_output)
            if digest(output.get("sha256"), f"{label}.outputs.{suffix}.sha256") != measured_sha:
                fail(f"{label}: {suffix} output hash differs")
            if integer(output.get("bytes"), f"{label}.outputs.{suffix}.bytes") != (measured_bytes if compressed_output else measured_path.stat().st_size):
                fail(f"{label}: {suffix} output size differs")
    return receipt


def _verify_build(root: Path, protocol: Mapping[str, Any], binding: Mapping[str, Any]) -> dict[str, Any]:
    driver = sha256_file(root / "capture.py")[0]
    started = _verify_command_receipt(root, root / "build.started.json", "build.started.json", driver, started=True)
    finished = _verify_command_receipt(root, root / "build.json", "build.json", driver)
    if started.get("argv") != finished.get("argv") or started.get("cwd") != finished.get("cwd"):
        fail("build receipts: command identity differs")
    expected_argv = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--target-dir",
        "/home/zhuhe/code/litchi/tools/perf-baseline/target", "--bin",
        "litchi-perf-baseline",
    ]
    if finished.get("argv") != expected_argv:
        fail("build.json: build argv differs")
    build_binding = obj(read_json(root / "binding.json", "binding.json"), "binding.json")
    if build_binding.get("schema") != "litchi-0468-build-binding-v1":
        fail("binding.json: schema differs")
    if build_binding.get("revision") != protocol["revision"] or build_binding.get("clean_build") is not True:
        fail("binding.json: build provenance differs")
    if build_binding.get("source_binding_sha256") != sha256_file(root / "source-binding.json")[0]:
        fail("binding.json: source binding hash differs")
    if build_binding.get("build_receipt_sha256") != sha256_file(root / "build.json")[0]:
        fail("binding.json: build receipt hash differs")
    binary = text(build_binding.get("binary_path"), "binding.binary_path")
    if build_binding.get("binary_sha256") != digest(build_binding.get("binary_sha256"), "binding.binary_sha256"):
        fail("binding.json: malformed binary hash")
    if integer(build_binding.get("bytes"), "binding.bytes", 1) <= 0:
        fail("binding.json: empty binary")
    if started.get("cwd") != binding.get("clean_tree"):
        fail("build.json: clean tree path differs from source binding")
    return build_binding


def _require_arg(argv: list[str], values: list[str], label: str) -> None:
    position = 0
    for value in values:
        try:
            position = argv.index(value, position) + 1
        except ValueError:
            fail(f"{label}: expected argv token {value!r}")


def _verify_process_receipt(root: Path, lane_root: Path, name: str, expected_driver: str) -> dict[str, Any]:
    return _verify_command_receipt(root, lane_root / name, f"{lane_root.name}/{name}", expected_driver)


def _verify_capture_receipt(root: Path, lane: str, expected_binary: Mapping[str, Any], protocol: Mapping[str, Any], binding: Mapping[str, Any]) -> dict[str, Any]:
    lane_root = root / lane
    receipt = obj(read_json(lane_root / "receipt.json", f"{lane}.receipt.json"), f"{lane}.receipt.json")
    if receipt.get("schema") != "litchi-0468-capture-v1" or receipt.get("lane") != lane:
        fail(f"{lane}: receipt schema/lane differs")
    if receipt.get("revision") != protocol["revision"] or receipt.get("exit_code", 0) != 0:
        fail(f"{lane}: receipt command failed or revision differs")
    if receipt.get("clean_source_before_and_after") is not True or receipt.get("source_and_binary_unchanged") is not True:
        fail(f"{lane}: source/build cleanliness attestation missing")
    if receipt.get("binary_sha256") != expected_binary["binary_sha256"]:
        fail(f"{lane}: binary identity differs")
    if receipt.get("binding_sha256") != sha256_file(root / "binding.json")[0] or receipt.get("protocol_sha256") != sha256_file(root / "protocol.json")[0]:
        fail(f"{lane}: binding/protocol identity differs")
    artifacts = obj(receipt.get("artifacts"), f"{lane}.artifacts")
    common_files = {
        "capture.json", "capture.started.json", "capture.stderr", "capture.stdout",
        "corpus-catalog.json", "report.json", "receipt.json", "resource.log",
    }
    expected_files = set(common_files)
    if lane == "counters":
        expected_files.add("counters.csv")
    if lane == "samples-fp":
        expected_files.update({
            "perf.data.gz", "perf-script.json", "perf-script.started.json",
            "perf-script.stderr", "perf-script.stdout.gz", "top-symbols.json",
            "top-symbols.started.json", "top-symbols.stderr", "top-symbols.stdout",
        })
    actual_files = {path.name for path in lane_root.iterdir() if path.is_file()}
    if actual_files != expected_files:
        fail(f"{lane}: final artifact set differs")
    for name, metadata in artifacts.items():
        candidate = root / lane / name
        if not candidate.exists() and lane == "samples-fp" and name == "perf.data":
            # export() replaces this raw file with the gzip member recorded in
            # compression.json, while retaining the capture-time receipt.
            continue
        path = bundle_file(root, f"{lane}/{name}", f"{lane}.artifacts.{name}")
        metadata = obj(metadata, f"{lane}.artifacts.{name}")
        if digest(metadata.get("sha256"), f"{lane}.artifacts.{name}.sha256") != sha256_file(path)[0]:
            fail(f"{lane}: artifact hash differs for {name}")
        if path.exists() and integer(metadata.get("bytes"), f"{lane}.artifacts.{name}.bytes") != path.stat().st_size:
            fail(f"{lane}: artifact size differs for {name}")
    driver = sha256_file(root / "capture.py")[0]
    command = _verify_process_receipt(root, lane_root, "capture.json", driver)
    _require_arg(command["argv"], ["taskset", "-c", "2", "/usr/bin/time", "-v"], f"{lane}.capture.json")
    if lane in {"counters", "samples-fp"}:
        _require_arg(command["argv"], ["perf"], f"{lane}.capture.json")
    samples, warmups = LANES[lane]
    _require_arg(command["argv"], ["--case", CASE, "--xlsx-shape", "dense-wide", "--workers", "1", "--warmup", str(warmups), "--samples", str(samples)], f"{lane}.capture.json")
    return receipt


def _verify_corpus_catalog(root: Path, lane: str, report: Mapping[str, Any], protocol: Mapping[str, Any], expected_revision: str) -> dict[str, Any]:
    catalog = obj(read_json(root / lane / "corpus-catalog.json", f"{lane}/corpus-catalog.json"), f"{lane}.corpus-catalog")
    if catalog.get("catalog_id") != "litchi-perf-corpus-v2" or catalog.get("manifest_version") != 2:
        fail(f"{lane}: catalog identity differs")
    if catalog.get("catalog_sha256") != CATALOG_SHA256 or catalog.get("content_set_sha256") != CONTENT_SET_SHA256:
        fail(f"{lane}: catalog digest differs")
    build = obj(catalog.get("build"), f"{lane}.catalog.build")
    if build.get("git_revision") != expected_revision or build.get("git_worktree_dirty") is not False:
        fail(f"{lane}: catalog build cleanliness differs")
    if build.get("source_files") != []:
        fail(f"{lane}: catalog source file semantics differ")
    bindings = catalog.get("case_bindings")
    if not isinstance(bindings, list) or len(bindings) != 1:
        fail(f"{lane}: catalog case bindings differ")
    row = obj(bindings[0], f"{lane}.catalog.case_bindings[0]")
    if row.get("case") != CASE or row.get("corpus_id") != f"xlsx-opc-zip:sha256:{CORPUS_SHA256}" or row.get("role") != "timed":
        fail(f"{lane}: catalog timed case identity differs")
    corpora = catalog.get("corpora")
    if not isinstance(corpora, list) or len(corpora) != 1:
        fail(f"{lane}: catalog corpus count differs")
    corpus = obj(corpora[0], f"{lane}.catalog.corpora[0]")
    if corpus.get("id") != row.get("corpus_id") or corpus.get("name") != "xlsx-dense-wide":
        fail(f"{lane}: catalog corpus identity differs")
    return catalog


def _validate_report(root: Path, report: Mapping[str, Any], lane: str, samples: int, warmups: int, binary: Mapping[str, Any], protocol: Mapping[str, Any], perf_compare: Any) -> tuple[dict[str, Any], dict[str, Any]]:
    report = obj(report, f"{lane}.report")
    policy_path = regular(root / "report-policy.json", "report-policy.json")
    # The policy parser is reused for its strict report/environment/schema
    # checks.  This local policy narrows it to the one 0468 workload and does
    # not perform a comparison or authorize a latency claim.
    policy = read_json(policy_path, "report-policy.json")
    policy = copy.deepcopy(policy)
    policy.update({"minimum_samples": samples, "expected_result_count": 1, "expected_result_keys_sha256": None, "required_cases": [CASE], "require_distinct_revisions": False, "require_clean_worktree": True})
    policy["expected_configuration"] = {"samples_per_case": samples, "warmup_iterations_per_case": warmups}
    policy["tool_identity"] = EXPECTED_TOOL
    try:
        perf_compare.validate_policy(policy)
        perf_compare._validate_report_identity(dict(report), dict(report), policy)
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        fail(f"{lane}.report: perf_compare validation failed ({error})")
    if report.get("tool") != EXPECTED_TOOL:
        fail(f"{lane}.report.tool: identity differs")
    report_binary = obj(report.get("binary_identity"), f"{lane}.report.binary_identity")
    if report_binary.get("binary_sha256") != binary["binary_sha256"] or report_binary.get("binary_bytes") != binary["binary_bytes"] or report_binary.get("profile") != "release" or report_binary.get("executable") is not True:
        fail(f"{lane}.report.binary_identity: identity differs from build binding")
    environment = obj(report.get("environment"), f"{lane}.report.environment")
    if environment.get("git_revision") != protocol["revision"] or environment.get("git_worktree_dirty") is not False:
        fail(f"{lane}.report.environment: clean revision identity differs")
    if environment.get("cpu_affinity") != "2" or environment.get("rustflags") != EXPECTED_ENV["RUSTFLAGS"]:
        fail(f"{lane}.report.environment: CPU/build flags differ")
    config = obj(report.get("configuration"), f"{lane}.report.configuration")
    if config.get("cases") != [CASE] or config.get("xlsx_shapes") != ["dense-wide"] or config.get("execution_workers") != [1]:
        fail(f"{lane}.report.configuration: workload identity differs")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        fail(f"{lane}.report.results: expected one result")
    row = obj(results[0], f"{lane}.report.results[0]")
    corpus = obj(row.get("corpus"), f"{lane}.report.results[0].corpus")
    if row.get("case") != CASE or corpus.get("archive_sha256") != CORPUS_SHA256 or corpus.get("shape") != "dense-wide" or corpus.get("name") != "xlsx-dense-wide":
        fail(f"{lane}.report: timed corpus identity differs")
    if corpus.get("package_format") != "XLSX/OPC/ZIP" or corpus.get("generator") != "litchi-xlsx-synthetic-v1":
        fail(f"{lane}.report: corpus semantics differ")
    catalog_ref = obj(report.get("corpus_catalog"), f"{lane}.report.corpus_catalog")
    if catalog_ref.get("catalog_id") != "litchi-perf-corpus-v2" or catalog_ref.get("manifest_version") != 2 or catalog_ref.get("catalog_sha256") != CATALOG_SHA256 or catalog_ref.get("content_set_sha256") != CONTENT_SET_SHA256:
        fail(f"{lane}.report.corpus_catalog: identity differs")
    elapsed = obj(row.get("elapsed_ns"), f"{lane}.report.elapsed_ns")
    values = elapsed.get("samples")
    if not isinstance(values, list) or len(values) != samples or any(isinstance(x, bool) or not isinstance(x, int) or x <= 0 for x in values):
        fail(f"{lane}.report.elapsed_ns.samples: sample vector differs")
    if values != sorted(values):
        fail(f"{lane}.report.elapsed_ns.samples: samples are not sorted")
    try:
        stats = perf_compare._latencies(row, f"{lane}.report.results[0]", samples)
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        fail(f"{lane}.report.elapsed_ns: percentile validation failed ({error})")
    expected_min, expected_max = min(values), max(values)
    if elapsed.get("min") != expected_min or elapsed.get("max") != expected_max:
        fail(f"{lane}.report.elapsed_ns: min/max differs from samples")
    mean = elapsed.get("mean")
    if isinstance(mean, bool) or not isinstance(mean, (int, float)) or not math.isfinite(float(mean)) or abs(float(mean) - statistics.mean(values)) > max(1e-6, abs(statistics.mean(values)) * 1e-12):
        fail(f"{lane}.report.elapsed_ns.mean: differs from samples")
    metrics = obj(row.get("operation_metrics"), f"{lane}.report.operation_metrics")
    if metrics.get("alignment") != "elapsed_ns.samples_by_elapsed_then_sample_index" or metrics.get("sample_count") != samples or metrics.get("sample_indices") != list(range(samples)):
        fail(f"{lane}.report.operation_metrics: sample/counter alignment differs")
    sink = obj(metrics.get("sink"), f"{lane}.report.operation_metrics.sink")
    if sink.get("status") != "not_applicable" or sink.get("write_status") != "measured":
        fail(f"{lane}.report.operation_metrics.sink: status semantics differ")
    for name in ("accepted_bytes", "largest_write", "write_calls"):
        vector = obj(sink.get(name), f"{lane}.report.operation_metrics.sink.{name}")
        values2 = vector.get("values")
        if vector.get("status") != "measured" or not isinstance(values2, list) or len(values2) != samples or any(isinstance(x, bool) or not isinstance(x, int) or x < 0 for x in values2):
            fail(f"{lane}.report.operation_metrics.sink.{name}: vector semantics differ")
        if len(set(values2)) != 1:
            fail(f"{lane}.report.operation_metrics.sink.{name}: output counter is not stable")
    return report, metrics


def _verify_counters(root: Path) -> dict[str, Any]:
    path = regular(root / "counters" / "counters.csv", "counters/counters.csv")
    rows: dict[str, int] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        fields = [field.strip() for field in line.split(";")]
        if len(fields) < 4 or fields[2] not in REQUIRED_COUNTERS or fields[2] in rows:
            fail("counters/counters.csv: malformed or duplicate counter")
        try:
            rows[fields[2]] = int(fields[0].replace(",", ""))
        except ValueError:
            fail(f"counters/counters.csv: invalid value for {fields[2]}")
        if rows[fields[2]] < 0:
            fail(f"counters/counters.csv: negative value for {fields[2]}")
    if tuple(rows) != REQUIRED_COUNTERS or any(value == 0 for value in rows.values()):
        fail("counters/counters.csv: required counter set/value differs")
    return {"counters": rows}


def _verify_compression(root: Path) -> dict[str, Any]:
    data = obj(read_json(root / "compression.json", "compression.json"), "compression.json")
    if data.get("schema") != "litchi-0468-compression-v1":
        fail("compression.json: schema differs")
    rows = data.get("artifacts")
    if not isinstance(rows, list) or len(rows) != 2:
        fail("compression.json: expected two compressed artifacts")
    expected = {"samples-fp/perf.data", "samples-fp/perf-script.stdout"}
    observed = set()
    for raw in rows:
        row = obj(raw, "compression.artifact")
        plain_name = relative(row.get("path"), "compression.path").as_posix()
        compressed_name = relative(row.get("compressed_path"), "compression.compressed_path").as_posix()
        if plain_name not in expected or compressed_name != plain_name + ".gz" or plain_name in observed:
            fail("compression.json: artifact path set differs")
        observed.add(plain_name)
        compressed = bundle_file(root, compressed_name, f"compression.{plain_name}")
        compressed_sha, compressed_bytes = sha256_file(compressed)
        plain_sha, plain_bytes = sha256_file(compressed, gzip_stream=True)
        if digest(row.get("sha256"), f"compression.{plain_name}.sha256") != plain_sha or integer(row.get("bytes"), f"compression.{plain_name}.bytes", 1) != plain_bytes:
            fail(f"compression.json: decompressed identity differs for {plain_name}")
        if digest(row.get("compressed_sha256"), f"compression.{plain_name}.compressed_sha256") != compressed_sha or integer(row.get("compressed_bytes"), f"compression.{plain_name}.compressed_bytes", 1) != compressed_bytes:
            fail(f"compression.json: compressed identity differs for {plain_name}")
        if not plain_bytes:
            fail(f"compression.json: empty compressed artifact {plain_name}")
    if observed != expected:
        fail("compression.json: artifact set differs")
    return {"artifacts": sorted(observed)}


def _verify_profile_artifacts(root: Path, protocol: Mapping[str, Any]) -> dict[str, Any]:
    lane = root / "samples-fp"
    for name in ("perf.data", "perf-script.stdout"):
        if (lane / name).exists():
            fail(f"samples-fp: uncompressed profile artifact remains: {name}")
    for name in ("perf.data.gz", "perf-script.stdout.gz", "top-symbols.stdout", "top-symbols.stderr", "perf-script.stderr"):
        path = regular(lane / name, f"samples-fp/{name}")
        if name in {"perf.data.gz", "perf-script.stdout.gz", "top-symbols.stdout"} and not path.stat().st_size:
            fail(f"samples-fp/{name}: empty")
    script = gzip.open(lane / "perf-script.stdout.gz", "rt", encoding="utf-8", errors="replace")
    with script:
        if "cycles:u" not in script.read():
            fail("samples-fp/perf-script.stdout.gz: no cycles:u samples")
    driver = sha256_file(root / "capture.py")[0]
    script_receipt = _verify_command_receipt(root, lane / "perf-script.json", "samples-fp/perf-script.json", driver, gzip_stdout=True)
    report_receipt = _verify_process_receipt(root, lane, "top-symbols.json", driver)
    script_argv = script_receipt["argv"]
    _require_arg(script_argv, ["perf", "script", "--no-inline", "-i"], "samples-fp/perf-script.json")
    if not any(value.endswith("/samples-fp/perf.data") for value in script_argv):
        fail("samples-fp/perf-script.json: perf.data input identity differs")
    _require_arg(report_receipt["argv"], ["perf", "report", "--stdio", "--no-children", "--no-inline", "-g", "none", "-i"], "samples-fp/top-symbols.json")
    if not any(value.endswith("/samples-fp/perf.data") for value in report_receipt["argv"]):
        fail("samples-fp/top-symbols.json: perf.data input identity differs")
    return {"profile_data": "samples-fp/perf.data.gz", "script": "samples-fp/perf-script.stdout.gz"}


def _verify_additional_context(root: Path, analyzer: Any, script: Path, reports: list[Path]) -> dict[str, Any]:
    """Recompute and authenticate the compact context/Amdahl artifact."""
    path = regular(root / "additional-summary.json", "additional-summary.json")
    try:
        expected = analyzer.profile_context(script, reports, top=40)
    except (OSError, UnicodeError, ValueError, TypeError, KeyError, RuntimeError) as error:
        fail(f"additional-summary.json: analyzer recomputation failed ({error})")
    actual = obj(read_json(path, "additional-summary.json"), "additional-summary.json")
    if not canonical_equal(actual, expected):
        fail("additional-summary.json: exact analyzer recomputation differs")
    if actual.get("schema") != "litchi-0468-profile-context-v1":
        fail("additional-summary.json: context schema differs")
    return {"path": "additional-summary.json"}


def _verify_summary(root: Path) -> dict[str, Any]:
    summary_path = regular(root / "summary-fp.json", "summary-fp.json")
    analyzer, analyzer_path = _load_analyzer(root)
    reports = [root / lane / "report.json" for lane in ("samples-fp", "normal-r1", "normal-r2")]
    script = root / "samples-fp" / "perf-script.stdout.gz"
    try:
        expected = analyzer.analyze_script(script, reports, top=40)
    except (OSError, UnicodeError, ValueError, TypeError, KeyError, RuntimeError) as error:
        fail(f"summary-fp.json: analyzer recomputation failed ({error})")
    actual = obj(read_json(summary_path, "summary-fp.json"), "summary-fp.json")
    if not canonical_equal(actual, expected):
        fail("summary-fp.json: exact analyzer recomputation differs")
    if actual.get("schema") != "litchi-0468-xlsx-cpu-attribution-v1" or actual.get("timing_semantics", {}).get("phase_latency") is not False or actual.get("timing_semantics", {}).get("operation_counter") is not False:
        fail("summary-fp.json: descriptive timing semantics differ")
    shared = getattr(analyzer, "LEGACY_PATH", None)
    if not isinstance(shared, Path):
        fail("analyze.py: shared analyzer dependency is not explicit")
    regular(shared, "../change-0466/analyze.py")
    reports = [root / lane / "report.json" for lane in ("samples-fp", "normal-r1", "normal-r2")]
    additional = _verify_additional_context(root, analyzer, script, reports)
    return {
        "path": "summary-fp.json",
        "analyzer": file_binding(analyzer_path, root),
        "shared_analyzer": file_binding(shared, root),
        "additional": additional,
    }


def _verify_observations(root: Path) -> dict[str, Any]:
    """Recompute the descriptive timing/counter chronology receipt exactly."""
    path = regular(root / "observe.py", "observe.py")
    spec = importlib.util.spec_from_file_location("litchi_0468_observe_verify", path)
    if spec is None or spec.loader is None:
        fail("observe.py: cannot load recomputation helper")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
        expected = module.summarize(root)
    except (OSError, SyntaxError, TypeError, ValueError, KeyError) as error:
        fail(f"observations.json: recomputation failed ({error})")
    actual = obj(read_json(root / "observations.json", "observations.json"), "observations.json")
    if not canonical_equal(actual, expected):
        fail("observations.json: exact recomputation differs")
    if actual.get("schema") != "litchi-0468-observations-v1" or actual.get("chronology_verified") is not True:
        fail("observations.json: chronology/schema semantics differ")
    return {"path": "observations.json", "helper": file_binding(path, root)}


def _verify_live_state(root: Path, binding: Mapping[str, Any], *, preseal: bool) -> dict[str, Any]:
    tree = Path(text(binding.get("clean_tree"), "source-binding.clean_tree"))
    binary = Path(text(binding.get("binary_path"), "binding.binary_path"))
    if preseal:
        if not tree.is_dir() or binary.is_symlink() or not binary.is_file():
            fail("preseal: authenticated source tree or binary is absent")
        try:
            revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=tree, text=True).strip()
            dirty = subprocess.check_output(["git", "status", "--porcelain"], cwd=tree, text=True)
        except (OSError, subprocess.CalledProcessError) as error:
            fail(f"preseal: cannot inspect clean source tree ({error})")
        if revision != REVISION or dirty:
            fail("preseal: source tree revision/cleanliness differs")
        source_map = obj(read_json(root / "sources.json", "sources.json"), "sources.json")
        for name, expected in source_map.items():
            source = tree / relative(name, f"sources.json.{name}")
            if source.is_symlink() or not source.is_file() or sha256_file(source)[0] != expected:
                fail(f"preseal: source digest differs for {name}")
        for name, expected in obj(binding.get("included_fixtures"), "source-binding.included_fixtures").items():
            source = tree / relative(name, f"source-binding.included_fixtures.{name}")
            if source.is_symlink() or not source.is_file() or sha256_file(source)[0] != expected:
                fail(f"preseal: fixture digest differs for {name}")
        actual_sha, actual_bytes = sha256_file(binary)
        expected_sha = digest(binding.get("binary_sha256"), "binding.binary_sha256")
        expected_bytes = integer(binding.get("bytes"), "binding.bytes", 1)
        if actual_sha != expected_sha or actual_bytes != expected_bytes:
            fail("preseal: live binary identity differs from binding")
        return {"mode": "preseal", "tree": str(tree), "binary_sha256": actual_sha, "binary_bytes": actual_bytes}
    if tree.is_symlink() or binary.is_symlink() or tree.exists() or binary.exists():
        fail("post-cleanup: live profile tree or binary remains")
    return {"mode": "post-cleanup", "tree_absent": True, "binary_absent": True}


def verify(*, root: Path = ROOT, preseal: bool = False) -> dict[str, Any]:
    root = root.resolve()
    protocol = _verify_protocol(root)
    binding = _verify_source_binding(root, protocol)
    build_binding = _verify_build(root, protocol, binding)
    binary_identity = {"binary_sha256": digest(build_binding.get("binary_sha256"), "binding.binary_sha256"), "binary_bytes": integer(build_binding.get("bytes"), "binding.bytes", 1)}
    perf_compare, perf_compare_path = _load_perf_compare(root)
    reports: dict[str, dict[str, Any]] = {}
    operation_metrics: dict[str, Any] | None = None
    catalogs = {}
    receipts = {}
    for lane in LANE_ORDER:
        samples, warmups = LANES[lane]
        receipt = _verify_capture_receipt(root, lane, binary_identity, protocol, binding)
        receipts[lane] = receipt
        report = obj(read_json(root / lane / "report.json", f"{lane}/report.json"), f"{lane}/report.json")
        report, metrics = _validate_report(root, report, lane, samples, warmups, binary_identity, protocol, perf_compare)
        reports[lane] = report
        catalogs[lane] = _verify_corpus_catalog(root, lane, report, protocol, protocol["revision"])
        signature = _counter_signature(metrics)
        if operation_metrics is None:
            operation_metrics = signature
        elif not canonical_equal(operation_metrics, signature):
            fail(f"{lane}: operation/output counter envelope differs from normal-r1")
    _verify_counters(root)
    _verify_compression(root)
    _verify_profile_artifacts(root, protocol)
    observations = _verify_observations(root)
    summary = _verify_summary(root)
    live = _verify_live_state(root, {**binding, **build_binding}, preseal=preseal)
    sealed = _verify_sha256sums(root) if not preseal else None
    return {
        "schema": SCHEMA,
        "change": CHANGE,
        "status": "pass",
        "mode": "preseal" if preseal else "post-cleanup",
        "lanes": list(LANE_ORDER),
        "case": CASE,
        "corpus_archive_sha256": CORPUS_SHA256,
        "normal_samples": 30,
        "profile_samples": 50,
        "binary_sha256": binary_identity["binary_sha256"],
        "binary_bytes": binary_identity["binary_bytes"],
        "summary": summary,
        "observations": observations,
        "sealed_files": sealed,
        "live_state": live,
        "dependencies": {
            "perf_compare": file_binding(perf_compare_path, root),
            "analyze": file_binding(root / "analyze.py", root),
            "report_policy": file_binding(root / "report-policy.json", root),
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--preseal", action="store_true", help="require the live authenticated profile tree and binary")
    parser.add_argument("--root", type=Path, default=ROOT, help="bundle root (used by portable replay)")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify(root=args.root, preseal=args.preseal)
    except (OSError, KeyError, TypeError, ValueError, VerificationError, subprocess.SubprocessError) as error:
        result = {"schema": SCHEMA, "change": CHANGE, "status": "fail", "error": str(error)}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result.get("status") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
