#!/usr/bin/env python3
"""Fail-closed and portable verifier for the 0474 PPTX streaming bundle."""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Mapping

import analyze


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0474-verification-v1"
HEX = analyze.HEX


class VerificationError(ValueError):
    """Raised when a retained artifact is missing or no longer bound."""


def fail(message: str) -> None:
    raise VerificationError(message)


def read_json(path: Path, label: str = "JSON") -> Any:
    try:
        return analyze.read_json(path)
    except analyze.AnalysisError as error:
        raise VerificationError(f"{label}: {error}") from error


def canonical(value: Any) -> bytes:
    try:
        return analyze.canonical(value)
    except analyze.AnalysisError as error:
        raise VerificationError(str(error)) from error


def sha256_file(path: Path) -> tuple[str, int]:
    try:
        return analyze.sha256_file(path)
    except analyze.AnalysisError as error:
        raise VerificationError(str(error)) from error


def digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or HEX.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def relative(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or "\\" in value:
        fail(f"{label}: expected a relative POSIX path")
    path = Path(value)
    if path.is_absolute() or path.as_posix() != value or ".." in path.parts:
        fail(f"{label}: path is not relative and traversal-free")
    return path


def regular(path: Path, label: str) -> Path:
    if not path.is_file() or path.is_symlink():
        fail(f"{label}: regular file required")
    return path


def bundle_path(root: Path, value: Any, label: str) -> Path:
    path = root / relative(value, label)
    try:
        if not path.resolve(strict=True).is_relative_to(root.resolve()):
            fail(f"{label}: path escapes bundle")
    except OSError as error:
        fail(f"{label}: cannot resolve path: {error}")
    return regular(path, label)


def artifact(path: Path, metadata: Mapping[str, Any], label: str) -> None:
    if not isinstance(metadata, dict):
        fail(f"{label}: artifact metadata is malformed")
    expected_hash = digest(metadata.get("sha256"), f"{label}.sha256")
    expected_bytes = metadata.get("bytes")
    if isinstance(expected_bytes, bool) or not isinstance(expected_bytes, int) or expected_bytes < 0:
        fail(f"{label}.bytes: expected a non-negative integer")
    actual_hash, actual_bytes = sha256_file(regular(path, label))
    if actual_hash != expected_hash or actual_bytes != expected_bytes:
        fail(f"{label}: hash or byte count differs")


def binding(path: Path, root: Path) -> dict[str, Any]:
    digest_value, bytes_value = sha256_file(path)
    try:
        name = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        name = path.name
    return {"path": name, "sha256": digest_value, "bytes": bytes_value}


def verify_protocol(root: Path = ROOT) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    path = root / "protocol.json"
    protocol = read_json(path, "protocol.json")
    if not isinstance(protocol, dict):
        fail("protocol.json: expected an object")
    try:
        order = analyze.protocol_order(protocol)
    except analyze.AnalysisError as error:
        fail(str(error))
    for key in ("capture_driver_sha256", "build_driver_sha256"):
        digest(protocol.get(key), f"protocol.{key}")
    capture = root / "capture.py"
    build = root / "build.py"
    if not capture.is_file() or not build.is_file():
        fail("capture.py and build.py are required")
    capture_hash, _ = sha256_file(capture)
    build_hash, _ = sha256_file(build)
    if protocol["capture_driver_sha256"] != capture_hash or protocol["build_driver_sha256"] != build_hash:
        fail("protocol driver hashes differ")
    return protocol, order


def _timestamp(value: Any, label: str) -> datetime.datetime:
    if not isinstance(value, str) or not value:
        fail(f"{label}: timestamp is missing")
    try:
        parsed = datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp")
    if parsed.tzinfo is None:
        fail(f"{label}: timestamp must include a timezone")
    return parsed


def verify_build(root: Path, protocol: Mapping[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    path = root / "build.json"
    build = read_json(path, "build.json")
    if not isinstance(build, dict) or build.get("schema") != "litchi-0474-build-v1":
        fail("build.json schema differs")
    if build.get("protocol_sha256") != sha256_file(root / "protocol.json")[0]:
        fail("build.json protocol hash differs")
    for field in ("revision", "build_path", "source_manifest", "fixtures", "binaries", "receipt"):
        if field not in build:
            fail(f"build.json lacks {field}")
    if not isinstance(build["revision"], str) or len(build["revision"]) < 7:
        fail("build revision is malformed")
    source = build["source_manifest"]
    if not isinstance(source, dict):
        fail("build source_manifest is malformed")
    source_path = bundle_path(root, source.get("path"), "build.source_manifest.path")
    source_hash, _ = sha256_file(source_path)
    if source_hash != digest(source.get("sha256"), "build.source_manifest.sha256"):
        fail("source manifest hash differs")
    manifest = read_json(source_path, "source manifest")
    if not isinstance(manifest, dict) or not manifest:
        fail("source manifest must be a non-empty path/hash object")
    if source.get("files") != 7034 or len(manifest) != 7034:
        fail("source manifest file count differs")
    for name, value in manifest.items():
        if not isinstance(name, str) or Path(name).is_absolute() or ".." in Path(name).parts:
            fail("source manifest contains an unsafe path")
        digest(value, f"source manifest {name}")
    required_sources = {"tools/perf-baseline/src/lib.rs", "tools/perf-baseline/src/pptx_streaming_create.rs", "tools/perf-baseline/src/allocation_metrics.rs", "tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs", "crates/litchi-pptx/src/writer/streaming.rs", "tools/perf-baseline/Cargo.lock"}
    if not required_sources.issubset(manifest): fail("source inventory omits a required implementation")
    fixtures = build["fixtures"]
    if not isinstance(fixtures, dict) or not fixtures:
        fail("build fixtures are missing")
    for name, value in fixtures.items():
        if not isinstance(name, str) or Path(name).is_absolute() or ".." in Path(name).parts:
            fail("build fixtures contain an unsafe path")
        digest(value, f"build fixture {name}")
    expected_fixtures = {"test-data/poi/test-data/spreadsheet/54016.xls": "2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a", "test-data/rtf/watermark.rtf": "48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb"}
    if fixtures != expected_fixtures: fail("compile fixture identities differ")
    if build["build_path"] != "/tmp/litchi-goal-0474/tree": fail("build path differs")
    binaries = build["binaries"]
    if not isinstance(binaries, dict) or set(binaries) != set(analyze.MODES):
        fail("build must bind exactly normal and allocator binaries")
    for mode in analyze.MODES:
        item = binaries[mode]
        if not isinstance(item, dict):
            fail(f"build.binaries.{mode} is malformed")
        digest(item.get("sha256"), f"build.binaries.{mode}.sha256")
        size = item.get("bytes")
        if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
            fail(f"build.binaries.{mode}.bytes is malformed")
        if not isinstance(item.get("path"), str) or not item["path"].startswith("/tmp/"):
            fail(f"build.binaries.{mode}.path must retain the temporary build path")
    receipt_value = build["receipt"]
    if not isinstance(receipt_value, dict):
        fail("build receipt binding is malformed")
    receipt_path = bundle_path(root, receipt_value.get("path"), "build.receipt.path")
    if sha256_file(receipt_path)[0] != digest(receipt_value.get("sha256"), "build.receipt.sha256"):
        fail("build receipt hash differs")
    receipt = read_json(receipt_path, "build receipt")
    if not isinstance(receipt, dict) or receipt.get("exit_code") != 0:
        fail("build command did not pass")
    expected_argv = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--target-dir", "/home/zhuhe/code/litchi/tools/perf-baseline/target", "--features", "allocator-metrics", "--bin", "litchi-perf-baseline", "--bin", "litchi-perf-baseline-alloc"]
    expected_env = {"RUSTUP_TOOLCHAIN": "1.98.1", "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes", "CARGO_PROFILE_RELEASE_DEBUG": "1", "CARGO_BUILD_JOBS": "4", "CARGO_INCREMENTAL": "0", "DEBUGINFOD_URLS": "", "LC_ALL": "C"}
    if receipt.get("argv") != expected_argv or receipt.get("cwd") != build["build_path"] or receipt.get("environment") != expected_env:
        fail("build command/environment differs")
    if _timestamp(protocol.get("declared_utc"), "protocol declared") > _timestamp(receipt.get("started_utc"), "build start"):
        fail("protocol was not frozen before build")
    if receipt.get("driver_sha256") != protocol.get("build_driver_sha256"):
        fail("build receipt driver hash differs")
    if receipt.get("protocol_sha256") != sha256_file(root / "protocol.json")[0]:
        fail("build receipt protocol hash differs")
    source_binding_path = root / "source-binding.json"
    source_binding = read_json(source_binding_path, "source-binding.json")
    if not isinstance(source_binding, dict):
        fail("source-binding.json is malformed")
    source_binding_hash, _ = sha256_file(source_binding_path)
    if receipt.get("source_binding_sha256") != source_binding_hash:
        fail("build receipt source binding hash differs")
    outputs = receipt.get("artifacts")
    if not isinstance(outputs, dict) or set(outputs) != {"build-command.stdout", "build-command.stderr"}:
        fail("build receipt output artifacts differ")
    for name, meta in outputs.items():
        artifact(root / name, meta, f"build receipt {name}")
    if source_binding.get("revision") != build["revision"]:
        fail("source binding revision differs")
    if source_binding.get("source_manifest") != source:
        fail("source binding manifest differs from build")
    if source_binding.get("fixtures") != fixtures:
        fail("source binding fixtures differ from build")
    started_path = root / "build-command.started.json"
    started = read_json(started_path, "build-command.started.json")
    for key in ("argv", "cwd", "driver_sha256", "protocol_sha256", "source_binding_sha256", "environment", "started_utc"):
        if started.get(key) != receipt.get(key):
            fail(f"build started/final receipt {key} differs")
    _timestamp(started.get("started_utc"), "build-command.started.json.started_utc")
    _timestamp(receipt.get("finished_utc"), "build-command.json.finished_utc")
    if _timestamp(receipt.get("finished_utc"), "build-command.json.finished_utc") < _timestamp(started.get("started_utc"), "build-command.started.json.started_utc"):
        fail("build finished before it started")
    return build, manifest


def _artifact_path(root: Path, lane: str, name: str) -> Path:
    if not isinstance(name, str) or Path(name).is_absolute() or ".." in Path(name).parts:
        fail(f"{lane}: unsafe artifact name {name!r}")
    local = root / "captures" / lane / name
    return regular(local, f"{lane}/{name}")


def verify_catalog(report: Mapping[str, Any], catalog: Mapping[str, Any], label: str) -> None:
    report_catalog = report.get("corpus_catalog")
    if not isinstance(report_catalog, dict):
        return
    for field in ("catalog_sha256", "content_set_sha256"):
        if field in report_catalog:
            if catalog.get(field) != report_catalog[field]:
                fail(f"{label}: catalog {field} differs from report")
            digest(catalog.get(field), f"{label}.catalog.{field}")


def verify_capture(
    root: Path,
    lane: Mapping[str, Any],
    build: Mapping[str, Any],
    protocol: Mapping[str, Any],
    protocol_hash: str,
) -> tuple[dict[str, Any], dict[str, Any]]:
    name = lane["lane"]
    directory = root / "captures" / name
    if not directory.is_dir() or directory.is_symlink():
        fail(f"{name}: capture directory missing")
    receipt_path = directory / "receipt.json"
    receipt = read_json(receipt_path, f"{name}/receipt.json")
    if not isinstance(receipt, dict) or receipt.get("schema") != "litchi-0474-capture-v1":
        fail(f"{name}: receipt schema differs")
    for key, expected in (("lane", name), ("mode", lane["mode"]), ("shape", lane["shape"]),
                          ("repeat", lane["repeat"]), ("revision", build["revision"]),
                          ("samples", analyze.SAMPLES), ("warmups", analyze.WARMUPS),
                          ("protocol_sha256", protocol_hash), ("driver_sha256", protocol["capture_driver_sha256"]),
                          ("binary_sha256", build["binaries"][lane["mode"]]["sha256"])):
        if receipt.get(key) != expected:
            fail(f"{name}: receipt {key} differs")
    if receipt.get("build_sha256") != sha256_file(root / "build.json")[0]:
        fail(f"{name}: build hash differs")
    for key in ("exit_code",):
        if receipt.get(key) != 0:
            fail(f"{name}: capture command did not pass")
    for key in ("clean_before", "clean_after", "binary_unchanged"):
        if receipt.get(key) is not True:
            fail(f"{name}: receipt {key} is not true")
    artifacts = receipt.get("artifacts")
    expected_names = {"started.json", "stdout.log", "stderr.log", "report.json", "corpus-catalog.json", "resource.log"}
    if not isinstance(artifacts, dict) or set(artifacts) != expected_names:
        fail(f"{name}: capture artifact set differs")
    for artifact_name, metadata in artifacts.items():
        artifact(_artifact_path(root, name, artifact_name), metadata, f"{name}/{artifact_name}")
    started_path = directory / "started.json"
    started = read_json(started_path, f"{name}/started.json")
    for key in ("lane", "mode", "shape", "repeat", "revision", "binary_sha256", "build_sha256", "driver_sha256", "protocol_sha256", "argv", "cwd", "samples", "warmups", "environment", "started_utc", "clean_before"):
        if started.get(key) != receipt.get(key):
            fail(f"{name}: started/final receipt {key} differs")
    # Capture receipts intentionally retain the original absolute output
    # paths so an exported bundle can verify those command identities without
    # rewriting the evidence.  The live bundle is created at this fixed
    # workspace path.
    original = Path("/home/zhuhe/code/litchi/docs/performance/results/change-0474/captures") / name
    expected_argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(original / "resource.log"), build["binaries"][lane["mode"]]["path"],
                     "--workers", "1", "--warmup", "3", "--samples", "30", "--case", analyze.CASE,
                     "--semantic-shape", lane["shape"], "--json", str(original / "report.json"), "--corpus-manifest", str(original / "corpus-catalog.json")]
    if receipt.get("argv") != expected_argv or receipt.get("cwd") != build["build_path"]:
        fail(f"{name}: capture command differs")
    expected_env = {"RUSTUP_TOOLCHAIN": "1.98.1", "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
                    "CARGO_PROFILE_RELEASE_DEBUG": "1", "DEBUGINFOD_URLS": "", "LC_ALL": "C"}
    if receipt.get("environment") != expected_env: fail(f"{name}: capture environment differs")
    report_path = directory / "report.json"
    report = read_json(report_path, f"{name}/report.json")
    if not isinstance(report, dict):
        fail(f"{name}: report is malformed")
    try:
        row = analyze.validate_report(report, lane, name)
    except analyze.AnalysisError as error:
        fail(str(error))
    environment = report.get("environment")
    if not isinstance(environment, dict) or environment.get("git_worktree_dirty") is not False or environment.get("git_revision") != build["revision"]:
        fail(f"{name}: report revision differs")
    binary_identity = report.get("binary_identity")
    if not isinstance(binary_identity, dict) or binary_identity.get("binary_sha256") != receipt["binary_sha256"] or binary_identity.get("path") != build["binaries"][lane["mode"]]["path"] or binary_identity.get("binary_bytes") != build["binaries"][lane["mode"]]["bytes"]:
        fail(f"{name}: report binary identity differs")
    catalog = read_json(directory / "corpus-catalog.json", f"{name}/corpus-catalog.json")
    if not isinstance(catalog, dict):
        fail(f"{name}: corpus catalog is malformed")
    verify_catalog(report, catalog, name)
    return report, receipt


def _parse_iso(value: Any, label: str) -> datetime.datetime:
    return _timestamp(value, label)


def verify_chronology(build_receipt: Mapping[str, Any], captures: list[Mapping[str, Any]]) -> None:
    build_start = _parse_iso(build_receipt.get("started_utc"), "build.started_utc")
    build_finish = _parse_iso(build_receipt.get("finished_utc"), "build.finished_utc")
    if build_finish < build_start:
        fail("build finished before it started")
    previous = build_finish
    for index, receipt in enumerate(captures):
        start = _parse_iso(receipt.get("started_utc"), f"capture[{index}].started_utc")
        finish = _parse_iso(receipt.get("finished_utc"), f"capture[{index}].finished_utc")
        if start < build_start or start < build_finish or finish < start or start < previous:
            fail("build/capture chronology is invalid")
        previous = finish


def verify_validation(root: Path) -> int:
    directory = root / "validation"
    if not directory.is_dir():
        fail("validation directory missing")
    receipts = sorted(directory.glob("*.json"))
    passed = 0
    for path in receipts:
        value = read_json(path, str(path))
        if not isinstance(value, dict) or not isinstance(value.get("exit_code"), int):
            fail(f"validation receipt is malformed: {path.name}")
        if value["exit_code"] == 0:
            passed += 1
        artifacts = value.get("artifacts")
        if not isinstance(artifacts, dict):
            fail(f"validation receipt lacks artifacts: {path.name}")
        for name, metadata in artifacts.items():
            target = root / name
            if not target.is_file():
                target = directory / name
            artifact(target, metadata, f"validation {name}")
        before = value.get("source_hashes")
        after = value.get("source_hashes_after", before)
        if before != after:
            fail(f"validation source custody changed: {path.name}")
    if receipts and passed == 0:
        fail("validation retained no passing receipt")
    required = ("harness-tests-formatted", "pptx-streaming-tests", "pptx-unit-all", "registry-strict-final", "format-final", "rustdoc", "clippy", "boundaries")
    for name in required:
        record = read_json(directory / (name + ".json"), name)
        if record.get("exit_code") != 0: fail(f"required validation failed: {name}")
    expected_commands = {
        "harness-tests-formatted": ["cargo", "test", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics", "--lib", "--bin", "litchi-perf-baseline-alloc"],
        "pptx-streaming-tests": ["cargo", "test", "--locked", "-p", "litchi-pptx", "--test", "streaming_authoring"],
        "pptx-unit-all": ["cargo", "test", "--locked", "-p", "litchi-pptx", "--lib"],
        "registry-strict-final": ["python3", "-B", "tools/check_perf_claims.py", "--registry", "docs/performance/claim-registry-v1.json", "--repo-root", ".", "--evidence-root", ".", "--mode", "strict"],
        "format-final": ["cargo", "fmt", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--", "--check"],
        "rustdoc": ["cargo", "doc", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics", "--no-deps"],
        "clippy": ["cargo", "clippy", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics", "--lib", "--no-deps", "--message-format=json", "--", "-D", "warnings"],
        "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
    }
    for name, argv in expected_commands.items():
        record = read_json(directory / (name + ".json"), name)
        if record.get("argv") != argv: fail(f"validation command differs: {name}")
        if name not in ("boundaries", "registry-strict-final") and record.get("environment", {}).get("RUSTUP_TOOLCHAIN") != "1.98.1":
            fail(f"validation toolchain differs: {name}")
    if read_json(directory / "rustdoc.json", "rustdoc")["environment"].get("RUSTDOCFLAGS") != "-D warnings":
        fail("rustdoc warnings were not denied")
    manifest = read_json(root / "sources/source.json", "manifest")
    for name in ("harness-tests-formatted", "format-final", "rustdoc", "clippy", "boundaries"):
        record = read_json(directory / (name + ".json"), name)
        if record.get("source_hashes") != manifest: fail(f"final gate source differs: {name}")
    return len(receipts)


def verify_seal(root: Path = ROOT) -> dict[str, Any]:
    path = root / "SHA256SUMS"
    if not path.is_file():
        fail("SHA256SUMS is missing")
    entries: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        parts = line.split("  ", 1)
        if len(parts) != 2:
            fail("malformed SHA256SUMS line")
        value, name = parts
        digest(value, f"SHA256SUMS {name}")
        if name in entries:
            fail(f"duplicate SHA256SUMS entry {name}")
        target = root / relative(name, f"SHA256SUMS {name}")
        artifact(target, {"sha256": value, "bytes": target.stat().st_size}, f"SHA256SUMS {name}")
        entries[name] = value
    actual = {path.relative_to(root).as_posix() for path in root.rglob("*") if path.is_file() and not path.is_symlink() and path.name != "SHA256SUMS"}
    if set(entries) != actual:
        fail("SHA256SUMS coverage differs from regular bundle files")
    return {"status": "pass", "files": len(entries)}


def verify_live_files(root: Path, build: Mapping[str, Any]) -> dict[str, Any]:
    tree = Path(build["build_path"])
    if not tree.is_dir():
        fail(f"live build tree is missing: {tree}")
    status = subprocess.check_output(["git", "status", "--porcelain"], cwd=tree, text=True)
    if status.strip():
        fail("live build tree is dirty")
    for mode in analyze.MODES:
        item = build["binaries"][mode]
        binary = Path(item["path"])
        if not binary.is_file() or binary.is_symlink():
            fail(f"live {mode} binary is missing")
        actual_hash, actual_size = sha256_file(binary)
        if actual_hash != item["sha256"] or actual_size != item["bytes"]:
            fail(f"live {mode} binary binding differs")
    if subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=tree, text=True).strip() != build["revision"]:
        fail("live tree revision differs")
    for name, expected in read_json(root / build["source_manifest"]["path"], "source manifest").items():
        if sha256_file(regular(tree / name, name))[0] != expected: fail(f"live source differs: {name}")
    for name, expected in build["fixtures"].items():
        if sha256_file(regular(tree / name, name))[0] != expected: fail(f"live fixture differs: {name}")
    return {"status": "pass", "tree": str(tree), "binaries": list(analyze.MODES)}


def verify_bundle(root: Path = ROOT, *, sealed: bool = False, live: bool = False) -> dict[str, Any]:
    root = root.resolve()
    protocol, order = verify_protocol(root)
    protocol_hash, _ = sha256_file(root / "protocol.json")
    build, _manifest = verify_build(root, protocol)
    reports: dict[tuple[str, str, str], Mapping[str, Any]] = {}
    receipts: list[Mapping[str, Any]] = []
    for lane in order:
        report, receipt = verify_capture(root, lane, build, protocol, protocol_hash)
        reports[(lane["mode"], lane["repeat"], lane["shape"])] = report["results"][0]
        receipts.append(receipt)
    try:
        identities = analyze.validate_identity_matrix(reports)
    except analyze.AnalysisError as error:
        fail(str(error))
    validation_count = verify_validation(root)
    summary_path = root / "summary.json"
    if not summary_path.is_file():
        fail("summary.json is missing")
    try:
        derived = analyze.build_summary(root)
    except analyze.AnalysisError as error:
        fail(str(error))
    summary = read_json(summary_path, "summary.json")
    if not analyze.canonical_equal(summary, derived):
        fail("summary.json is not the current derived summary")
    build_receipt = read_json(root / build["receipt"]["path"], "build receipt")
    verify_chronology(build_receipt, receipts)
    if _timestamp(build_receipt["finished_utc"], "build finish") > _timestamp(receipts[0]["started_utc"], "first capture"):
        fail("build finished after the first capture started")
    seal_result = verify_seal(root) if sealed else {"status": "not_checked"}
    live_result = verify_live_files(root, build) if live else {"status": "not_checked"}
    if sealed and not live:
        if Path(build["build_path"]).exists() or any(Path(v["path"]).exists() for v in build["binaries"].values()):
            fail("post-cleanup sealed replay requires temporary tree and binaries absent")
    return {
        "schema": SCHEMA,
        "status": "pass",
        "lanes": len(order),
        "samples": len(order) * analyze.SAMPLES,
        "shapes": list(analyze.SHAPES),
        "identity_shapes": len(identities),
        "validation_receipts": validation_count,
        "sealed": seal_result,
        "live": live_result,
    }


def portable_check(root: Path = ROOT, *, sealed: bool = False) -> list[str]:
    with tempfile.TemporaryDirectory(prefix="litchi-0474-portable-") as directory:
        exported = Path(directory) / "bundle"
        shutil.copytree(root, exported, symlinks=False)
        verify_bundle(exported, sealed=sealed, live=False)
        report = exported / "captures" / "R1-normal-tiny" / "report.json"
        original = report.read_bytes()
        value = json.loads(original)
        value["results"][0]["output_sha256"] = "0" * 64
        report.write_text(json.dumps(value), encoding="utf-8")
        command = [sys.executable, "-B", str(exported / "verify.py")]
        if sealed:
            command.append("--sealed")
        result = subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if result.returncode == 0:
            fail("portable output mutation was accepted")
        report.write_bytes(original)
        return ["isolated export passed", "output identity mutation rejected"]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sealed", action="store_true")
    parser.add_argument("--live", action="store_true")
    parser.add_argument("--portable-check", action="store_true")
    args = parser.parse_args()
    try:
        result = verify_bundle(ROOT, sealed=args.sealed, live=args.live)
        if args.portable_check:
            result["portable"] = portable_check(ROOT, sealed=args.sealed)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (OSError, TypeError, ValueError, KeyError, VerificationError, analyze.AnalysisError) as error:
        print(json.dumps({"status": "fail", "error": str(error)}, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
