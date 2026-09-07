#!/usr/bin/env python3
"""Authenticate the retained change-0463 ODP comparison bundle.

The verifier only reads the bundle for ordinary and portable runs.  The
optional ``--precleanup`` pass additionally authenticates the retained
executables and the current Rust source tree, selecting the source epoch from
the final decision.  It never builds, captures, or changes evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
# A portable copy can be nested more deeply than the repository checkout.
REPO = ROOT.parents[3] if len(ROOT.parents) > 3 else ROOT
TASK = Path("/tmp/litchi-goal-0463")
CHANGE = 463
SCHEMA = "litchi-0463-verification-v1"
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
RETRY_RE = re.compile(r"^(.*)-r([0-9]+)$")
ASSEMBLY_TARGETS = (
    ("transaction-commit", "<litchi_odp::authoring::edit::Transaction>::commit"),
    ("to-bytes-bounded", "<litchi_odp::authoring::mutable::MutablePresentation>::to_bytes_bounded"),
    ("validate-compact-xml-parts-against", "litchi_odp::authoring::edit::validate_compact_xml_parts_against"),
)
ASSEMBLY_INSTRUCTION_RE = re.compile(r"^\s*[0-9a-f]+:\s+\S")
ASSEMBLY_STACK_RE = re.compile(r"\bsub\s+\$0x([0-9a-f]+),%rsp\b")
ASSEMBLY_CALL_RE = re.compile(r"\bcall(?:q)?\s+[^<\n]*<(.+)>")
ASSEMBLY_HEADER_RE = re.compile(r"^\s*[0-9a-f]+\s+<.*>:\s*$")


class VerificationError(AssertionError):
    pass


def fail(message: str) -> None:
    raise VerificationError(message)


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON ({error})")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected JSON object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    if value != value.lower():
        fail(f"{label}: expected lowercase SHA-256")
    if SHA256_RE.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def sha_file(path: Path) -> str:
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return value.hexdigest()


def regular(path: Path, label: str) -> Path:
    if path.is_symlink() or not path.is_file():
        fail(f"{label}: missing, symlink, or non-regular file: {path}")
    return path


def safe_path(base: Path, value: Any, label: str, *, must_exist: bool = True) -> Path:
    raw = text(value, label)
    relative = Path(raw)
    if relative.is_absolute() or ".." in relative.parts:
        fail(f"{label}: path must be relative and traversal-free")
    if base.is_symlink():
        fail(f"{label}: base is a symlink")
    candidate = base / relative
    for ancestor in (base, *candidate.parents):
        if ancestor.is_symlink():
            fail(f"{label}: path has a symlinked ancestor")
    try:
        resolved = candidate.resolve(strict=must_exist)
        base_resolved = base.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve path ({error})")
    if not resolved.is_relative_to(base_resolved):
        fail(f"{label}: path escapes its base")
    if must_exist:
        regular(candidate, label)
    elif candidate.exists() or candidate.is_symlink():
        if candidate.is_symlink() or not candidate.is_file():
            fail(f"{label}: existing path is unsafe")
    return candidate


def artifact(base: Path, value: Any, label: str) -> Path:
    row = obj(value, label)
    path = safe_path(base, row.get("path"), f"{label}.path")
    expected_bytes = integer(row.get("bytes"), f"{label}.bytes")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    if path.stat().st_size != expected_bytes or sha_file(path) != expected_sha:
        fail(f"{label}: artifact identity differs")
    return path


def assembly_observations(raw: bytes, symbol: str) -> dict[str, Any]:
    """Recompute status and generated-code observations from retained objdump."""
    lines = raw.decode("utf-8", errors="replace").splitlines()
    marker = symbol + ">:"
    try:
        start = next(index for index, line in enumerate(lines) if marker in line)
    except StopIteration:
        return {"status": "missing_or_inlined", "instruction_count": 0, "stack_frame_bytes": None, "calls": []}
    body: list[str] = []
    for line in lines[start + 1 :]:
        if ASSEMBLY_HEADER_RE.match(line):
            break
        body.append(line)
    instructions = [line for line in body if ASSEMBLY_INSTRUCTION_RE.search(line)]
    frames = [int(match.group(1), 16) for line in body if (match := ASSEMBLY_STACK_RE.search(line))]
    calls = sorted({match.group(1) for line in body if (match := ASSEMBLY_CALL_RE.search(line))})
    return {
        "status": "present" if instructions else "missing_or_inlined",
        "instruction_count": len(instructions),
        "stack_frame_bytes": max(frames, default=0),
        "calls": calls,
    }


def source_record(value: Any, label: str) -> tuple[Path, dict[str, str], dict[str, Any]]:
    row = obj(value, label)
    path = safe_path(ROOT, row.get("path"), f"{label}.path")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    expected_files = integer(row.get("files"), f"{label}.files")
    if sha_file(path) != expected_sha:
        fail(f"{label}: source manifest hash differs")
    manifest = obj(load(path, label), label)
    if len(manifest) != expected_files:
        fail(f"{label}: source manifest file count differs")
    normalized: dict[str, str] = {}
    for name, value in manifest.items():
        rel = text(name, f"{label}.files path")
        candidate = Path(rel)
        if candidate.is_absolute() or ".." in candidate.parts or candidate.as_posix() != rel:
            fail(f"{label}: source path escapes repository: {rel}")
        if not rel.endswith((".rs", ".toml", ".lock")):
            fail(f"{label}: non-Rust source entry {rel}")
        normalized[rel] = digest(value, f"{label}.files[{rel}]")
    if not normalized:
        fail(f"{label}: source manifest is empty")
    return path, normalized, row


def source_equal(left: Mapping[str, Any], right: Mapping[str, Any], label: str) -> None:
    if (
        left.get("path") != right.get("path")
        or left.get("sha256") != right.get("sha256")
        or left.get("files") != right.get("files")
    ):
        fail(f"{label}: source identity differs")


def verify_sum_file(path: Path, base: Path) -> int:
    """Verify a complete checksum inventory and reject every symlink."""
    regular(path, path.name)
    rows: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"{path}: cannot read checksum file ({error})")
    for index, line in enumerate(lines, 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"{path}: malformed checksum line {index}")
        name = fields[1]
        relative = Path(name)
        if relative.is_absolute() or ".." in relative.parts or not name:
            fail(f"{path}: unsafe checksum member {name!r}")
        if name == path.name or name in rows:
            fail(f"{path}: duplicate or self checksum member {name}")
        member = safe_path(base, name, f"{path.name}.{name}")
        if sha_file(member) != fields[0]:
            fail(f"{path}: hash differs for {name}")
        rows[name] = fields[0]
    actual: set[str] = set()
    for member in base.rglob("*"):
        if member.is_symlink():
            fail(f"{base}: bundle contains symlink {member.relative_to(base)}")
        if member.is_file() and member != path:
            actual.add(member.relative_to(base).as_posix())
    if set(rows) != actual:
        missing = sorted(actual - set(rows))
        extra = sorted(set(rows) - actual)
        fail(f"{path}: checksum coverage differs (missing={missing[:3]}, extra={extra[:3]})")
    return len(rows)


def check_hash_binding(value: Mapping[str, Any], key: str, path: Path, label: str) -> None:
    expected = digest(value.get(key), f"{label}.{key}")
    if expected != sha_file(path):
        fail(f"{label}.{key}: hash differs for {path.name}")


def expected_lanes() -> set[tuple[str, str, str, str]]:
    return {
        (repeat, instrumentation, shape, "lifecycle")
        for repeat in ("R1", "R2")
        for instrumentation in ("normal", "allocator")
        for shape in ("tiny", "medium", "large")
    }


def check_protocol() -> tuple[dict[str, Any], dict[str, Any]]:
    """Validate the frozen A1/B1/B2/A2 matrix and its two binary epochs."""
    protocol = obj(load(ROOT / "protocol.json", "protocol.json"), "protocol.json")
    baseline = obj(load(ROOT / "baseline-binding.json", "baseline-binding.json"), "baseline-binding.json")
    candidate = obj(load(ROOT / "candidate-binding.json", "candidate-binding.json"), "candidate-binding.json")
    if protocol.get("change") != CHANGE or protocol.get("schema") != "litchi-0463-protocol-v1":
        fail("protocol schema/change differs")
    revision = text(protocol.get("revision") or baseline.get("revision"), "protocol.revision")
    protocol.setdefault("revision", revision)
    order = protocol.get("order")
    if not isinstance(order, list) or len(order) != 12:
        fail("protocol.order must contain exactly 12 lifecycle lanes")
    observed: set[tuple[str, str, str, str]] = set()
    seen_ids: set[str] = set()
    for index, raw in enumerate(order):
        lane = obj(raw, f"protocol.order[{index}]")
        if set(lane) != {"id", "repeat", "instrumentation", "shape", "scope"}:
            fail(f"protocol.order[{index}] lane keys differ")
        dimensions = (lane["repeat"], lane["instrumentation"], lane["shape"], lane["scope"])
        if dimensions not in expected_lanes():
            fail(f"protocol.order[{index}] lane dimensions differ")
        lane_id = text(lane["id"], f"protocol.order[{index}].id")
        if lane_id in seen_ids:
            fail("protocol.order contains duplicate lane ids")
        seen_ids.add(lane_id)
        observed.add(dimensions)
    if observed != expected_lanes():
        fail("protocol.order does not cover the exact 12-lane variant matrix")
    for key, expected in {"cpu": 2, "workers": 1, "samples": 30, "warmup": 3, "reports_per_variant": 12, "retained_samples": 720}.items():
        if protocol.get(key) != expected:
            fail(f"protocol.{key} differs")
    if protocol.get("sequence") != ["baseline-R1", "candidate-R1", "candidate-R2", "baseline-R2"]:
        fail("protocol sequence differs")
    argv = protocol.get("argv")
    if not isinstance(argv, list) or not argv or not all(isinstance(item, str) and item for item in argv):
        fail("protocol argv is malformed")
    expected_argv = ["{binary}", "odp-append-attribution", "--mode", "{scope}", "--shape", "{shape}", "--warmup", "3", "--samples", "30", "--repeat", "{repeat}", "--output", "{report}"]
    if argv != expected_argv:
        fail("protocol argv differs")
    if not isinstance(protocol.get("environment"), dict):
        fail("protocol.environment must be an object")
    claims = text(protocol.get("claims"), "protocol.claims").lower()
    if "cold" not in claims or "scaling" not in claims:
        fail("protocol claims do not disclose cold/range/worker-scaling limits")
    for key, path in (("capture_sha256", ROOT / "capture.py"), ("oracle_sha256", ROOT / "oracle.py"), ("host_sha256", ROOT / "host.json"), ("baseline_binding_sha256", ROOT / "baseline-binding.json"), ("prior_control_binding_sha256", ROOT / "prior-control-bindings.json")):
        check_hash_binding(protocol, key, path, "protocol")
    if "candidate_binding_sha256" in protocol:
        check_hash_binding(protocol, "candidate_binding_sha256", ROOT / "candidate-binding.json", "protocol")
    for name in ("compare_sha256", "counters_sha256"):
        if name in protocol:
            check_hash_binding(protocol, name, ROOT / (name.removesuffix("_sha256") + ".py"), "protocol")
    for name, binding in (("baseline", baseline), ("candidate", candidate)):
        if binding.get("change") != CHANGE or binding.get("schema") != "litchi-0463-binary-binding-v1" or binding.get("variant") != name:
            fail(f"{name} binding schema/change/variant differs")
        if binding.get("revision") != revision:
            fail(f"{name} binding revision differs")
        build_ref = obj(binding.get("build_receipt"), f"{name}.build_receipt")
        build_path = artifact(ROOT, build_ref, f"{name}.build_receipt")
        if build_ref.get("path") != f"checks/{name}-build.json":
            fail(f"{name} build receipt path differs")
        build = obj(load(build_path, f"checks/{name}-build.json"), f"checks/{name}-build.json")
        if build.get("status") != "pass" or build.get("exit_code") != 0 or build.get("revision") != revision or build.get("source_unchanged") is not True:
            fail(f"{name} build receipt is not a successful immutable build")
        source_ref = obj(binding.get("source_manifest"), f"{name}.source_manifest")
        source_record(source_ref, f"{name}.source_manifest")
        after = obj(build.get("source_after"), f"{name}.build.source_after")
        source_equal(source_ref, after, f"{name} build/source manifest")
        source_record(after, f"{name}.build.source_after")
        binaries = obj(binding.get("binaries"), f"{name}.binaries")
        if set(binaries) != {"normal", "allocator"}:
            fail(f"{name} binding must contain normal and allocator binaries")
        for mode, raw in binaries.items():
            row = obj(raw, f"{name}.binaries.{mode}")
            binary_path = Path(text(row.get("path"), f"{name}.binaries.{mode}.path"))
            if not binary_path.is_absolute() or not binary_path.is_relative_to(TASK):
                fail(f"{name} binary path is not task-owned")
            source = Path(text(row.get("source"), f"{name}.binaries.{mode}.source"))
            if source.is_absolute() or ".." in source.parts:
                fail(f"{name} binary source escapes repository")
            integer(row.get("bytes"), f"{name}.binaries.{mode}.bytes", 1)
            digest(row.get("sha256"), f"{name}.binaries.{mode}.sha256")
    if baseline["source_manifest"] == candidate["source_manifest"]:
        fail("baseline and candidate source epochs must differ")
    return protocol, baseline, candidate


def check_current_sources(source_ref: Mapping[str, Any]) -> None:
    _, manifest, _ = source_record(source_ref, "final source manifest")
    try:
        names = set(
            subprocess.check_output(
                ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z"],
                cwd=REPO,
            ).decode().split("\0")
        )
    except (OSError, subprocess.CalledProcessError, UnicodeError) as error:
        fail(f"cannot enumerate current source files: {error}")
    names.add("Cargo.lock")
    current = {
        name
        for name in names
        if name.endswith((".rs", ".toml", ".lock")) and (REPO / name).is_file()
    }
    if current != set(manifest):
        fail("current Rust source manifest coverage differs")
    for relative, expected in manifest.items():
        path = REPO / relative
        regular(path, f"current source {relative}")
        if sha_file(path) != expected:
            fail(f"current source changed: {relative}")


def check_live_file(path_value: Any, expected_sha: Any, expected_bytes: Any, label: str) -> None:
    path = Path(text(path_value, f"{label}.path"))
    if path.is_symlink() or not path.is_file() or not path.is_relative_to(TASK):
        fail(f"{label}: live file is missing, unsafe, or outside task directory")
    if path.stat().st_size != integer(expected_bytes, f"{label}.bytes", 1) or sha_file(path) != digest(expected_sha, f"{label}.sha256"):
        fail(f"{label}: live identity differs")


def load_oracle() -> Any:
    path = ROOT / "oracle.py"
    spec = importlib.util.spec_from_file_location("litchi0463_oracle", path)
    if spec is None or spec.loader is None:
        fail("cannot load oracle.py")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        fail(f"oracle.py failed to import: {error}")
    if not callable(getattr(module, "validate", None)):
        fail("oracle.py has no callable validate")
    return module


def check_prior_control(protocol: Mapping[str, Any]) -> None:
    path = ROOT / "prior-control-bindings.json"
    value = obj(load(path, path.name), path.name)
    if value.get("schema") != "litchi-0458-prior-control-bindings-v1" or value.get("change") != 458:
        fail("prior-control-bindings schema/change differs")
    shapes = obj(value.get("shapes"), "prior-control-bindings.shapes")
    if set(shapes) != {"tiny", "medium", "large"}:
        fail("prior-control-bindings must cover tiny, medium, and large")
    for shape, raw in shapes.items():
        row = obj(raw, f"prior-control-bindings.shapes.{shape}")
        if row.get("path") != f"prior-control/{shape}.json":
            fail(f"prior-control {shape} path differs")
        artifact(ROOT, row, f"prior-control-bindings.shapes.{shape}")
        prior = obj(load(ROOT / row["path"], f"prior-control/{shape}.json"), f"prior-control/{shape}.json")
        results = prior.get("results")
        if not isinstance(results, list) or len(results) != 1:
            fail(f"prior-control/{shape}.json must contain one result")
        result = obj(results[0], f"prior-control/{shape}.json.results[0]")
        corpus = obj(result.get("corpus"), f"prior-control/{shape}.json corpus")
        source = obj(result.get("source"), f"prior-control/{shape}.json source")
        odp = obj(source.get("odp_append"), f"prior-control/{shape}.json source.odp_append")
        if corpus.get("shape") != shape or odp.get("shape") != shape:
            fail(f"prior-control/{shape}.json shape differs")
    if digest(protocol.get("prior_control_binding_sha256"), "protocol.prior_control_binding_sha256") != sha_file(path):
        fail("protocol prior-control binding hash differs")


def _source_manifest(row: Mapping[str, Any], label: str) -> dict[str, str]:
    """Load a source manifest and return its path-to-digest mapping."""
    _, manifest, _ = source_record(row, label)
    return manifest


def _delta_artifact(value: Any, label: str, expected_digest: str | None) -> None:
    """Check a source-delta artifact, allowing null for an absent epoch file."""
    if value is None:
        if expected_digest is not None:
            fail(f"{label}: source artifact is absent for an existing epoch file")
        return
    row = obj(value, label)
    relative = Path(text(row.get("path"), f"{label}.path"))
    if relative.parts[:1] != ("source-artifacts",):
        fail(f"{label}.path: source artifact must be under source-artifacts/")
    path = safe_path(ROOT, relative.as_posix(), f"{label}.path")
    if path.suffix != ".txt":
        fail(f"{label}.path: source artifact must have .txt suffix")
    expected_bytes = integer(row.get("bytes"), f"{label}.bytes")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    if path.stat().st_size != expected_bytes or sha_file(path) != expected_sha:
        fail(f"{label}: artifact identity differs")
    if expected_digest is None or expected_sha != expected_digest:
        fail(f"{label}: artifact does not match compiled source manifest")


def check_source_delta(baseline: Mapping[str, Any], candidate: Mapping[str, Any]) -> None:
    """Require exact source-epoch delta coverage and custody of each changed file."""
    delta = obj(load(ROOT / "source-delta.json", "source-delta.json"), "source-delta.json")
    if delta.get("change") != CHANGE:
        fail("source-delta change differs")
    schema = delta.get("schema")
    if schema is not None and schema != "litchi-0463-source-delta-v1":
        fail("source-delta schema differs")
    before = _source_manifest(baseline["source_manifest"], "baseline source manifest")
    after = _source_manifest(candidate["source_manifest"], "candidate source manifest")
    for key, expected in (("baseline_source_manifest", baseline["source_manifest"]), ("candidate_source_manifest", candidate["source_manifest"])):
        if key in delta:
            source_equal(obj(delta[key], f"source-delta.{key}"), expected, f"source-delta.{key}")
    changed = {name for name in set(before) | set(after) if before.get(name) != after.get(name)}
    raw_files = delta.get("files", delta.get("delta"))
    rows: dict[str, Any] = {}
    if isinstance(raw_files, dict):
        rows = raw_files
    elif isinstance(raw_files, list):
        for index, raw in enumerate(raw_files):
            entry = obj(raw, f"source-delta.files[{index}]")
            name = text(entry.get("path"), f"source-delta.files[{index}].path")
            if name in rows:
                fail(f"source-delta.files: duplicate {name}")
            rows[name] = entry
    else:
        fail("source-delta.files must be an object or array")
    if set(rows) != changed:
        fail(f"source-delta exact coverage differs (expected={sorted(changed)}, got={sorted(rows)})")
    referenced: list[str] = []
    for name in sorted(changed):
        relative = Path(name)
        if relative.is_absolute() or ".." in relative.parts or relative.as_posix() != name:
            fail(f"source-delta file path is unsafe: {name}")
        entry = obj(rows[name], f"source-delta.files[{name}]")
        if "path" in entry and entry.get("path") != name:
            fail(f"source-delta.files[{name}].path repeats a different name")
        baseline_value = entry.get("baseline", entry.get("before"))
        candidate_value = entry.get("candidate", entry.get("after"))
        _delta_artifact(baseline_value, f"source-delta.files[{name}].baseline", before.get(name))
        _delta_artifact(candidate_value, f"source-delta.files[{name}].candidate", after.get(name))
        for side, value in (("baseline", baseline_value), ("candidate", candidate_value)):
            if value is not None:
                referenced.append(text(obj(value, f"source-delta.files[{name}].{side}").get("path"), f"source-delta.files[{name}].{side}.path"))
    if len(referenced) != len(set(referenced)):
        fail("source-delta references one source artifact more than once")
    source_root = ROOT / "source-artifacts"
    if not source_root.is_dir() or source_root.is_symlink():
        fail("source-artifacts directory is missing or unsafe")
    members = [path for path in source_root.rglob("*") if path.is_file() or path.is_symlink()]
    if any(path.is_symlink() for path in members):
        fail("source-artifacts contains a symlink")
    non_txt = [path.relative_to(ROOT).as_posix() for path in members if path.suffix != ".txt"]
    if non_txt:
        fail(f"source-artifacts contains non-text members: {sorted(non_txt)[:3]}")
    artifact_paths = {path.relative_to(ROOT).as_posix() for path in members}
    if set(referenced) != artifact_paths:
        unbound = sorted(artifact_paths - set(referenced))
        dangling = sorted(set(referenced) - artifact_paths)
        fail(f"source-delta source-artifact coverage differs (unbound={unbound[:3]}, dangling={dangling[:3]})")


def check_check_receipts(protocol: Mapping[str, Any], bindings: tuple[Mapping[str, Any], Mapping[str, Any]]) -> dict[str, dict[str, Any]]:
    directory = ROOT / "checks"
    if not directory.is_dir() or directory.is_symlink():
        fail("checks directory is missing or unsafe")
    receipts: dict[str, dict[str, Any]] = {}
    referenced_logs: set[str] = set()
    referenced_sources: set[str] = set()
    for path in sorted(directory.glob("*.json")):
        label = f"checks/{path.name}"
        receipt = obj(load(path, label), label)
        if receipt.get("change") != CHANGE or digest(receipt.get("driver_sha256"), f"{label}.driver_sha256") != sha_file(ROOT / "check.py"):
            fail(f"{label}: change or driver hash differs")
        if receipt.get("revision") != protocol.get("revision"):
            fail(f"{label}: revision differs")
        status = receipt.get("status")
        if status not in {"running", "pass", "failed"}:
            fail(f"{label}: unsupported status {status!r}")
        argv = receipt.get("argv")
        if not isinstance(argv, list) or not argv or not all(isinstance(item, str) and item for item in argv):
            fail(f"{label}: argv is malformed")
        cwd = text(receipt.get("cwd"), f"{label}.cwd")
        if Path(cwd) != Path(cwd).resolve() or not Path(cwd).is_absolute():
            fail(f"{label}: cwd is not canonical absolute path")
        before = obj(receipt.get("source_before"), f"{label}.source_before")
        _, _, before_row = source_record(before, f"{label}.source_before")
        referenced_sources.add(text(before_row.get("path"), f"{label}.source_before.path"))
        if status == "running":
            fail(f"{label}: check is still running")
        after = obj(receipt.get("source_after"), f"{label}.source_after")
        _, _, after_row = source_record(after, f"{label}.source_after")
        referenced_sources.add(text(after_row.get("path"), f"{label}.source_after.path"))
        unchanged = receipt.get("source_unchanged")
        if type(unchanged) is not bool or unchanged != (before == after):
            fail(f"{label}: source_unchanged is inconsistent")
        exit_code = receipt.get("exit_code")
        if isinstance(exit_code, bool) or not isinstance(exit_code, int):
            fail(f"{label}: completed check has no integer exit code")
        if status == "pass" and (exit_code != 0 or unchanged is not True):
            fail(f"{label}: passing check is not successful and source-stable")
        if status == "failed" and exit_code == 0:
            fail(f"{label}: failed check has zero exit code")
        log_path = artifact(ROOT, receipt.get("log"), f"{label}.log")
        referenced_logs.add(log_path.relative_to(ROOT).as_posix())
        receipts[path.stem] = receipt
    if not receipts:
        fail("checks directory contains no receipts")
    actual_logs = {
        path.relative_to(ROOT).as_posix()
        for path in directory.iterdir()
        if path.is_file() and path.suffix in {".log", ".gz"}
    }
    if actual_logs != referenced_logs:
        fail(f"checks log coverage differs (unbound={sorted(actual_logs - referenced_logs)[:3]})")
    source_directory = ROOT / "sources"
    if not source_directory.is_dir() or source_directory.is_symlink():
        fail("sources directory is missing or unsafe")
    for binding in bindings:
        source_ref = binding.get("source_manifest", binding.get("source"))
        row = obj(source_ref, "binding source manifest")
        referenced_sources.add(text(row.get("path"), "binding source path"))
    actual_sources = {
        path.relative_to(ROOT).as_posix()
        for path in source_directory.iterdir()
        if path.is_file() and not path.is_symlink()
    }
    if actual_sources != referenced_sources:
        fail(f"source manifest coverage differs (unbound={sorted(actual_sources - referenced_sources)[:3]})")
    return receipts


DEFAULT_GATES: dict[str, tuple[str, ...]] = {
    name: (name,)
    for name in (
        "baseline-build",
        "baseline-bind",
        "baseline-assembly",
        "baseline-r1",
        "baseline-r2",
        "candidate-build",
        "candidate-bind",
        "candidate-assembly",
        "candidate-r1",
        "candidate-r2",
        "compare",
        "counters",
        "phase-diagnostic",
        "derive-phases",
        "owner-tests",
        "owner-clippy",
        "format",
        "doc",
        "boundaries",
        "harness-tests",
    )
}


def gate_matches(name: str, pattern: str) -> bool:
    return name == pattern or bool(re.fullmatch(re.escape(pattern) + r"-r[0-9]+", name))


def latest_passing(names: list[str], receipts: Mapping[str, Mapping[str, Any]]) -> str | None:
    passing = [name for name in names if receipts[name].get("status") == "pass"]
    if not passing:
        return None

    def rank(name: str) -> tuple[int, str]:
        match = RETRY_RE.fullmatch(name)
        return (int(match.group(2)) if match else 0, name)

    return max(passing, key=rank)


def check_required_gates(protocol: Mapping[str, Any], receipts: Mapping[str, Mapping[str, Any]]) -> dict[str, str]:
    configured: Any = protocol.get("required_gates", {name: list(patterns) for name, patterns in DEFAULT_GATES.items()})
    if isinstance(configured, list):
        configured = {name: [name] for name in configured}
    configured = obj(configured, "protocol.required_gates")
    missing_groups = sorted(set(DEFAULT_GATES) - set(configured))
    if missing_groups:
        fail(f"protocol.required_gates omits required groups: {missing_groups}")
    selected: dict[str, str] = {}
    for group, raw in configured.items():
        patterns = [raw] if isinstance(raw, str) else raw
        if not isinstance(patterns, list) or not patterns or not all(isinstance(item, str) and item for item in patterns):
            fail(f"protocol.required_gates.{group}: expected names")
        candidates = sorted(name for name in receipts if any(gate_matches(name, pattern) for pattern in patterns))
        chosen = latest_passing(candidates, receipts)
        if chosen is None:
            fail(f"required gate {group} has no passing receipt")
        selected[group] = chosen
    return selected


def path_suffix(value: Any, suffix: str, label: str) -> str:
    raw = text(value, label)
    candidate = Path(raw)
    if not candidate.is_absolute() or candidate != candidate.resolve() or not candidate.as_posix().endswith("/" + suffix):
        fail(f"{label}: expected canonical absolute path ending in {suffix}")
    return raw


def binary_binding_path(value: Any, bound_path: Any, label: str) -> str:
    """Bind archived absolute argv values to the retained binary path.

    A portable bundle preserves the captured command line verbatim.  The
    binding therefore remains the original retained ``/tmp`` path; it is not
    rewritten to the temporary directory used for the portable copy.
    """
    raw = text(value, label)
    candidate = Path(raw)
    if not candidate.is_absolute() or candidate != candidate.resolve():
        fail(f"{label}: expected canonical absolute path")
    expected = text(bound_path, f"{label}.bound_path")
    expected_path = Path(expected)
    if not expected_path.is_absolute() or expected_path != expected_path.resolve():
        fail(f"{label}.bound_path: expected canonical absolute path")
    if raw != expected:
        fail(f"{label}: captured path differs from retained binding path")
    return raw


def expected_workload(protocol: Mapping[str, Any], lane: Mapping[str, Any], binary: str, report: str) -> list[str]:
    try:
        return [part.format(binary=binary, report=report, **lane) for part in protocol["argv"]]
    except (KeyError, ValueError) as error:
        fail(f"protocol argv cannot be expanded: {error}")
    raise AssertionError("unreachable")


def check_run_artifacts(run_root: Path, receipt: Mapping[str, Any], label: str, expected: set[str]) -> None:
    rows = receipt.get("artifacts")
    if not isinstance(rows, list) or not rows:
        fail(f"{label}: artifacts list is missing")
    seen: set[str] = set()
    for index, raw in enumerate(rows):
        row = obj(raw, f"{label}.artifacts[{index}]")
        relative = text(row.get("path"), f"{label}.artifacts[{index}].path")
        if relative in seen:
            fail(f"{label}: duplicate artifact {relative}")
        seen.add(relative)
        path = artifact(ROOT, row, f"{label}.artifacts[{index}]")
        if not path.is_relative_to(run_root):
            fail(f"{label}: artifact escapes its run directory")
    if seen != expected:
        fail(f"{label}: artifact set differs (expected={sorted(expected)}, got={sorted(seen)})")
    actual = {
        path.relative_to(ROOT).as_posix()
        for path in run_root.iterdir()
        if path.is_file() and path.name != "receipt.json"
    }
    if actual != seen:
        fail(f"{label}: run directory contains unbound artifacts")
    for entry in run_root.iterdir():
        if entry.name != "receipt.json" and (entry.is_symlink() or not entry.is_file()):
            fail(f"{label}: unexpected non-file entry {entry.name}")


def check_runs(protocol: Mapping[str, Any], bindings: Mapping[str, Mapping[str, Any]], oracle: Any) -> None:
    runs_root = ROOT / "runs"
    if not runs_root.is_dir() or runs_root.is_symlink():
        fail("runs directory is missing or unsafe")
    if {entry.name for entry in runs_root.iterdir()} != {"baseline", "candidate"}:
        fail("runs directory must contain exactly baseline and candidate")
    for variant, binding in bindings.items():
        variant_root = runs_root / variant
        if not variant_root.is_dir() or variant_root.is_symlink():
            fail(f"runs/{variant}: variant directory is missing or unsafe")
        expected_ids = {lane["id"] for lane in protocol["order"]}
        actual_ids = {entry.name for entry in variant_root.iterdir() if entry.is_dir() and not entry.is_symlink()}
        if actual_ids != expected_ids:
            fail(f"runs/{variant}: exact 12 lane directories are required")
        binaries = obj(binding.get("binaries"), f"{variant}.binaries")
        for raw_lane in protocol["order"]:
            lane = obj(raw_lane, "protocol lane")
            lane_id = lane["id"]
            run_root = variant_root / lane_id
            receipt_path = run_root / "receipt.json"
            report_path = run_root / "report.json"
            receipt = obj(load(receipt_path, f"runs/{variant}/{lane_id}/receipt.json"), f"runs/{variant}/{lane_id}/receipt.json")
            regular(report_path, f"runs/{variant}/{lane_id}/report.json")
            label = f"runs/{variant}/{lane_id}"
            if receipt.get("schema") != "litchi-0463-capture-v1" or receipt.get("change") != CHANGE or receipt.get("variant") != variant or receipt.get("lane") != lane:
                fail(f"{label}: receipt identity differs")
            if receipt.get("protocol_sha256") != sha_file(ROOT / "protocol.json") or receipt.get("binary_binding_sha256") != sha_file(ROOT / (variant + "-binding.json")):
                fail(f"{label}: protocol or binding hash differs")
            if receipt.get("environment_fixed") != protocol.get("environment") or not isinstance(receipt.get("allocator_environment"), dict):
                fail(f"{label}: environment metadata differs")
            if receipt.get("status") != "pass" or receipt.get("exit_code") != 0:
                fail(f"{label}: capture did not pass")
            command = receipt.get("argv")
            if not isinstance(command, list) or not all(isinstance(item, str) for item in command):
                fail(f"{label}: argv is malformed")
            if len(command) < 8 or command[:2] != ["/usr/bin/time", "-v"] or command[2] != "-o" or command[4:6] != ["taskset", "-c"] or command[6] != str(protocol["cpu"]):
                fail(f"{label}: capture argv prefix differs")
            path_suffix(command[3], f"runs/{variant}/{lane_id}/resource.log", f"{label}.argv.resource")
            mode_binary = obj(binaries.get(lane["instrumentation"]), f"{label}.binary")
            expected_binary = binary_binding_path(command[7], mode_binary.get("path"), f"{label}.binary.path")
            workload = command[7:]
            try:
                output_index = workload.index("--output")
                report_arg = workload[output_index + 1]
            except (ValueError, IndexError):
                fail(f"{label}: capture argv has no output report")
            path_suffix(report_arg, f"runs/{variant}/{lane_id}/report.json", f"{label}.argv.report")
            if workload != expected_workload(protocol, lane, expected_binary, report_arg):
                fail(f"{label}: workload argv differs from protocol")
            proof = oracle.validate(report_path, lane, protocol)
            if receipt.get("oracle") != proof:
                fail(f"{label}: retained oracle proof differs from recomputation")
            expected = {
                f"runs/{variant}/{lane_id}/report.json",
                f"runs/{variant}/{lane_id}/stdout.log",
                f"runs/{variant}/{lane_id}/stderr.log",
                f"runs/{variant}/{lane_id}/resource.log",
            }
            check_run_artifacts(run_root, receipt, label, expected)


def check_counters(protocol: Mapping[str, Any], bindings: Mapping[str, Mapping[str, Any]], oracle: Any) -> None:
    root = ROOT / "counters"
    if not root.exists():
        return
    if not root.is_dir() or root.is_symlink() or {entry.name for entry in root.iterdir()} != {"baseline", "candidate"}:
        fail("counters directory must contain exactly baseline and candidate")
    diagnostic_protocol = dict(protocol)
    diagnostic_protocol["samples"] = 100
    for variant in ("baseline", "candidate"):
        run_root = root / variant
        receipt = obj(load(run_root / "receipt.json", f"counters/{variant}/receipt.json"), f"counters/{variant}/receipt.json")
        binding_path = ROOT / f"{variant}-binding.json"
        binding = bindings[variant]
        binary = text(obj(binding["binaries"], f"{variant}.binaries")["normal"]["path"], f"{variant}.normal.path")
        label = f"counters/{variant}"
        if receipt.get("schema") != "litchi-0463-counters-v1" or receipt.get("variant") != variant or receipt.get("status") != "pass":
            fail(f"{label}: receipt identity/status differs")
        if receipt.get("binding_sha256") != sha_file(binding_path) or receipt.get("driver_sha256") != sha_file(ROOT / "counters.py"):
            fail(f"{label}: binding or driver hash differs")
        command = receipt.get("argv")
        if not isinstance(command, list) or receipt.get("exit_code") != 0:
            fail(f"{label}: counters workload did not pass")
        expected_workload = [binary, "odp-append-attribution", "--mode", "lifecycle", "--shape", "large", "--warmup", "3", "--samples", "100", "--repeat", "diagnostic", "--output", "REPORT"]
        expected = ["/usr/bin/time", "-v", "-o", "RESOURCE", "taskset", "-c", str(protocol["cpu"]), "perf", "stat", "-x,", "-o", "COUNTERS", "-e", "cycles,instructions,branches,branch-misses,cache-misses,page-faults,context-switches", "--", *expected_workload]
        if len(command) != len(expected):
            fail(f"{label}: counters argv length differs")
        for index, expected_value in enumerate(expected):
            if index == 3:
                path_suffix(command[index], f"counters/{variant}/resource.log", f"{label}.argv.resource")
            elif index == 11:
                path_suffix(command[index], f"counters/{variant}/counters.csv", f"{label}.argv.counters")
            elif index == 15:
                binary_binding_path(command[index], obj(binding["binaries"], f"{variant}.binaries")["normal"].get("path"), f"{label}.argv.binary")
            elif index == len(expected) - 1:
                path_suffix(command[index], f"counters/{variant}/report.json", f"{label}.argv.report")
            elif command[index] != expected_value:
                fail(f"{label}: counters argv differs")
        proof = oracle.validate(run_root / "report.json", {"id": variant, "repeat": "diagnostic", "instrumentation": "normal", "shape": "large", "scope": "lifecycle"}, diagnostic_protocol)
        if receipt.get("oracle") != proof:
            fail(f"{label}: retained oracle proof differs from recomputation")
        expected_artifacts = {
            f"counters/{variant}/report.json",
            f"counters/{variant}/resource.log",
            f"counters/{variant}/counters.csv",
            f"counters/{variant}/stdout.log",
            f"counters/{variant}/stderr.log",
        }
        check_run_artifacts(run_root, receipt, label, expected_artifacts)


def check_supplementary_protocol(protocol: Mapping[str, Any]) -> dict[str, Any]:
    path = ROOT / "supplementary-protocol.json"
    supplementary = obj(load(path, path.name), path.name)
    if supplementary.get("schema") != "litchi-0463-supplementary-protocol-v1":
        fail("supplementary protocol schema differs")
    if supplementary.get("protocol_sha256") != sha_file(ROOT / "protocol.json"):
        fail("supplementary protocol does not bind protocol.json")
    check_hash_binding(supplementary, "phase_driver_sha256", ROOT / "phase-diagnostic.py", "supplementary protocol")
    check_hash_binding(supplementary, "counter_driver_sha256", ROOT / "counters.py", "supplementary protocol")
    if supplementary.get("phase_sequence") != ["baseline-R1", "candidate-R1", "candidate-R2", "baseline-R2"]:
        fail("supplementary phase sequence differs")
    for key, expected in (("phase_samples_per_report", 30), ("phase_shape", "large"), ("counter_sequence", ["baseline", "candidate"]), ("counter_samples_per_report", 100)):
        if supplementary.get(key) != expected:
            fail(f"supplementary protocol.{key} differs")
    return supplementary


def check_phase_diagnostic(protocol: Mapping[str, Any], bindings: Mapping[str, Mapping[str, Any]], oracle: Any, supplementary: Mapping[str, Any]) -> None:
    root = ROOT / "phase-diagnostic"
    expected_ids = {"baseline-R1", "candidate-R1", "candidate-R2", "baseline-R2"}
    if not root.is_dir() or root.is_symlink():
        fail("phase-diagnostic directory is missing or unsafe")
    actual_ids = {entry.name for entry in root.iterdir() if entry.is_dir() and not entry.is_symlink()}
    if actual_ids != expected_ids:
        fail("phase-diagnostic must contain exactly four matched runs")
    for lane_id in ("baseline-R1", "candidate-R1", "candidate-R2", "baseline-R2"):
        variant, repeat = lane_id.split("-")
        lane = {"id": lane_id, "repeat": repeat, "instrumentation": "normal", "shape": "large", "scope": "phases"}
        run_root = root / lane_id
        label = f"phase-diagnostic/{lane_id}"
        receipt = obj(load(run_root / "receipt.json", f"{label}/receipt.json"), f"{label}/receipt.json")
        report = run_root / "report.json"
        regular(report, f"{label}/report.json")
        if receipt.get("schema") != "litchi-0463-phase-diagnostic-v1" or receipt.get("variant") != variant or receipt.get("lane") != lane or receipt.get("status") != "pass" or receipt.get("exit_code") != 0:
            fail(f"{label}: receipt identity/status differs")
        if receipt.get("binding_sha256") != sha_file(ROOT / f"{variant}-binding.json") or receipt.get("driver_sha256") != sha_file(ROOT / "phase-diagnostic.py") or receipt.get("protocol_sha256") != sha_file(ROOT / "protocol.json"):
            fail(f"{label}: receipt binding differs")
        command = receipt.get("argv")
        if not isinstance(command, list) or not all(isinstance(item, str) for item in command):
            fail(f"{label}: argv is malformed")
        binary = obj(bindings[variant]["binaries"], f"{variant}.binaries")["normal"]
        expected = ["/usr/bin/time", "-v", "-o", "RESOURCE", "taskset", "-c", str(protocol["cpu"]), "BINARY", "odp-append-attribution", "--mode", "phases", "--shape", "large", "--warmup", "3", "--samples", "30", "--repeat", repeat, "--output", "REPORT"]
        if len(command) != len(expected):
            fail(f"{label}: argv length differs")
        for index, expected_value in enumerate(expected):
            if index == 3:
                path_suffix(command[index], f"phase-diagnostic/{lane_id}/resource.log", f"{label}.argv.resource")
            elif index == 7:
                binary_binding_path(command[index], binary.get("path"), f"{label}.argv.binary")
            elif index == len(expected) - 1:
                path_suffix(command[index], f"phase-diagnostic/{lane_id}/report.json", f"{label}.argv.report")
            elif expected_value != command[index]:
                fail(f"{label}: argv differs")
        proof = oracle.validate(report, lane, protocol)
        if receipt.get("oracle") != proof:
            fail(f"{label}: retained oracle proof differs from recomputation")
        check_run_artifacts(run_root, receipt, label, {
            f"phase-diagnostic/{lane_id}/report.json",
            f"phase-diagnostic/{lane_id}/resource.log",
            f"phase-diagnostic/{lane_id}/stdout.log",
            f"phase-diagnostic/{lane_id}/stderr.log",
        })


def recompute_json(path: Path, script: Path, label: str) -> None:
    regular(path, label)
    environment = os.environ.copy()
    environment["PYTHONPATH"] = ""
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    result = subprocess.run([sys.executable, "-B", str(script)], cwd=ROOT, env=environment, capture_output=True, text=True)
    if result.returncode != 0:
        fail(f"{label}: recomputation failed ({result.stderr.strip() or result.stdout.strip()})")
    try:
        actual = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        fail(f"{label}: recomputation did not return JSON ({error})")
    expected = load(path, label)
    if actual != expected:
        fail(f"{label}: recomputed JSON differs from retained file")


def candidate_kept(decision: Mapping[str, Any]) -> bool:
    if isinstance(decision.get("candidate_kept"), bool):
        return bool(decision["candidate_kept"])
    if isinstance(decision.get("keep_candidate"), bool):
        return bool(decision["keep_candidate"])
    if isinstance(decision.get("production_change_retained"), bool):
        return bool(decision["production_change_retained"])
    value = decision.get("decision", decision.get("outcome"))
    if value is None:
        value = decision.get("status")
    if not isinstance(value, str):
        fail("decision must disclose candidate_kept or a decision/outcome string")
    normalized = value.lower().replace("-", "_").replace(" ", "_")
    if normalized in {"keep", "kept", "keep_candidate", "candidate", "accept", "accepted", "accepted_scoped", "retain_candidate", "retained"}:
        return True
    if normalized in {"baseline", "revert", "reverted", "reject", "rejected", "reject_candidate", "revert_candidate", "baseline_only", "retain_baseline"}:
        return False
    fail(f"decision outcome is not a recognized keep/revert result: {value!r}")
    raise AssertionError("unreachable")


def check_decision(baseline: Mapping[str, Any]) -> tuple[dict[str, Any], bool]:
    decision = obj(load(ROOT / "decision.json", "decision.json"), "decision.json")
    schema = decision.get("schema")
    if schema is not None and schema != "litchi-0463-decision-v1":
        fail("decision schema differs")
    if type(decision.get("production_change_retained")) is not bool:
        fail("decision.production_change_retained must be boolean")
    kept = bool(decision["production_change_retained"])
    if "candidate_kept" in decision and decision["candidate_kept"] != kept:
        fail("decision candidate_kept disagrees with production_change_retained")
    if "keep_candidate" in decision and decision["keep_candidate"] != kept:
        fail("decision keep_candidate disagrees with production_change_retained")
    status = decision.get("status")
    if isinstance(status, str) and status.lower() == "rejected" and kept:
        fail("rejected decision cannot retain the candidate")
    if isinstance(status, str) and status.lower() in {"accepted", "kept", "retained"} and not kept:
        fail("accepted decision must retain the candidate")
    summary_path = ROOT / "summary.json"
    summary = obj(load(summary_path, "summary.json"), "summary.json")
    if digest(decision.get("summary_sha256"), "decision.summary_sha256") != sha_file(summary_path):
        fail("decision.summary_sha256 does not bind summary.json")
    comparisons = summary.get("comparisons")
    if not isinstance(comparisons, list):
        fail("summary.comparisons must be an array")
    targets = []
    for row_value in comparisons:
        row = obj(row_value, "summary.comparisons[]")
        lane = obj(row.get("lane"), "summary.comparisons[].lane")
        if lane.get("instrumentation") == "normal" and lane.get("shape") in {"medium", "large"}:
            targets.append(row)
    if len(targets) != 4:
        fail("summary must contain four normal medium/large gate comparisons")
    gate = True
    for row in targets:
        delta = obj(row.get("elapsed_delta_pct"), "summary comparison elapsed_delta_pct")
        interval = obj(row.get("p50_delta_95pct_interval"), "summary comparison interval")
        gate = gate and isinstance(delta.get("p50"), (int, float)) and delta["p50"] <= -3 and isinstance(interval.get("high"), (int, float)) and interval["high"] < 0
    if summary.get("predeclared_latency_gate_pass") is not gate:
        fail("summary predeclared gate does not match independent verification")
    if kept and not gate:
        fail("decision retains candidate without passing the 3% and CI gate")
    if "allocation_increase_review_flags" not in summary:
        fail("summary must disclose allocation_increase_review_flags")
    if decision.get("allocation_increase_review_flags") != summary["allocation_increase_review_flags"]:
        fail("decision allocation_increase_review_flags do not match summary")
    if "adverse_flags" in decision and decision["adverse_flags"] != summary.get("adverse_5pct_flags"):
        fail("decision adverse_flags do not match summary")
    elif "adverse_5pct_flags" in decision and decision["adverse_5pct_flags"] != summary.get("adverse_5pct_flags"):
        fail("decision adverse_5pct_flags do not match summary")
    elif "adverse_flags" not in decision and "adverse_5pct_flags" not in decision:
        fail("decision must retain the adverse timing/RSS flag review")
    if not kept:
        _, baseline_manifest, _ = source_record(baseline["source_manifest"], "baseline source manifest")
        restorations: list[tuple[str, Any]] = []
        if "restored_file" in decision:
            restorations.append((text(decision.get("restored_file"), "decision.restored_file"), decision.get("restored_sha256")))
        if "restored_files" in decision:
            values = decision["restored_files"]
            if not isinstance(values, list) or not values:
                fail("decision.restored_files must be a non-empty array")
            for index, value in enumerate(values):
                row = obj(value, f"decision.restored_files[{index}]")
                restorations.append((text(row.get("path"), f"decision.restored_files[{index}].path"), row.get("sha256")))
        for restored_file, restored_sha in restorations:
            relative = Path(restored_file)
            if relative.is_absolute() or ".." in relative.parts or relative.as_posix() != restored_file:
                fail("decision restored source path is unsafe")
            expected_restored = baseline_manifest.get(restored_file)
            if expected_restored is None:
                fail("decision restored source file is absent from the baseline source epoch")
            if digest(restored_sha, f"decision restored hash {restored_file}") != expected_restored:
                fail(f"decision restored hash does not bind baseline source epoch: {restored_file}")
    return decision, kept


def check_assembly(bindings: Mapping[str, Mapping[str, Any]], *, compare_live: bool = False) -> None:
    """Authenticate every retained disassembly, including absent/inlined targets."""
    expected_all_asm = {
        f"{variant}-{target_id}.asm.txt"
        for variant in bindings
        for target_id, _ in ASSEMBLY_TARGETS
    }
    for variant, binding in bindings.items():
        label = variant + "-assembly"
        row = obj(load(ROOT / (label + ".json"), label), label)
        if row.get("change") != CHANGE or row.get("variant") != variant:
            fail(f"{label}: batch or variant differs")
        if row.get("driver_sha256") != sha_file(ROOT / "assembly.py"):
            fail(f"{label}: assembly driver differs")
        if row.get("binding_sha256") != sha_file(ROOT / (variant + "-binding.json")):
            fail(f"{label}: executable binding differs")
        binary = binding["binaries"]["normal"]
        if row.get("binary_sha256") != binary["sha256"]:
            fail(f"{label}: executable identity differs")
        targets = row.get("targets")
        if not isinstance(targets, list) or len(targets) != len(ASSEMBLY_TARGETS):
            fail(f"{label}: target count differs")
        artifacts = row.get("artifacts")
        if not isinstance(artifacts, list) or len(artifacts) != len(ASSEMBLY_TARGETS):
            fail(f"{label}: artifact count differs")
        expected_commands: list[list[str]] = []
        expected_paths = {f"{variant}-{target_id}.asm.txt" for target_id, _ in ASSEMBLY_TARGETS}
        normalized_artifacts: dict[str, Mapping[str, Any]] = {}
        for index, (target_id, symbol) in enumerate(ASSEMBLY_TARGETS):
            target = obj(targets[index], f"{label}.targets[{index}]")
            if target.get("id") != target_id or target.get("symbol") != symbol:
                fail(f"{label}.targets[{index}]: target identity differs")
            command = ["objdump", "-d", "--demangle", "--no-show-raw-insn", "--disassemble=" + symbol, binary["path"]]
            if target.get("command") != command:
                fail(f"{label}.targets[{index}]: command differs")
            expected_commands.append(command)
            expected_path = f"{variant}-{target_id}.asm.txt"
            target_artifact = obj(target.get("artifact"), f"{label}.targets[{index}].artifact")
            if target_artifact.get("path") != expected_path:
                fail(f"{label}.targets[{index}]: artifact path differs")
            path = artifact(ROOT, target_artifact, f"{label}.targets[{index}].artifact")
            observations = assembly_observations(path.read_bytes(), symbol)
            for key in ("status", "instruction_count", "stack_frame_bytes", "calls"):
                if target.get(key) != observations[key]:
                    fail(f"{label}.targets[{index}]: {key} differs from retained objdump")
            normalized_artifacts[expected_path] = target_artifact
            if compare_live:
                try:
                    live = subprocess.check_output(command)
                except (OSError, subprocess.CalledProcessError) as error:
                    fail(f"{label}.targets[{index}]: live objdump failed ({error})")
                if live != path.read_bytes():
                    fail(f"{label}.targets[{index}]: live disassembly differs")
        if row.get("commands") != expected_commands:
            fail(f"{label}: command inventory differs")
        actual_paths = set()
        for index, item in enumerate(artifacts):
            item = obj(item, f"{label}.artifacts[{index}]")
            path = artifact(ROOT, item, f"{label}.artifacts[{index}]")
            relative = path.relative_to(ROOT).as_posix()
            if relative not in expected_paths or normalized_artifacts.get(relative) != item:
                fail(f"{label}: artifact inventory differs")
            actual_paths.add(relative)
        if actual_paths != expected_paths:
            fail(f"{label}: artifact coverage differs")
    actual_asm = {
        path.relative_to(ROOT).as_posix()
        for path in ROOT.rglob("*.asm.txt")
        if path.is_file() and not path.is_symlink()
    }
    if actual_asm != expected_all_asm:
        fail(f"assembly: exact .asm.txt whitelist differs")


def check_live_bindings(baseline: Mapping[str, Any], candidate: Mapping[str, Any], final_source: Mapping[str, Any]) -> None:
    for variant, binding in (("baseline", baseline), ("candidate", candidate)):
        for mode, raw in obj(binding.get("binaries"), f"{variant}.binaries").items():
            row = obj(raw, f"{variant}.binaries.{mode}")
            check_live_file(row.get("path"), row.get("sha256"), row.get("bytes"), f"{variant}.{mode}")
    check_current_sources(final_source)


def verify(precleanup: bool = False, portable: bool = False) -> dict[str, Any]:
    if precleanup and portable:
        fail("--precleanup and --portable are mutually exclusive")
    sealed_files = verify_sum_file(ROOT / "SHA256SUMS", ROOT)
    protocol, baseline, candidate = check_protocol()
    bindings = (baseline, candidate)
    receipts = check_check_receipts(protocol, bindings)
    selected = check_required_gates(protocol, receipts)
    check_prior_control(protocol)
    check_source_delta(baseline, candidate)
    check_assembly({"baseline": baseline, "candidate": candidate})
    oracle = load_oracle()
    check_runs(protocol, {"baseline": baseline, "candidate": candidate}, oracle)
    check_counters(protocol, {"baseline": baseline, "candidate": candidate}, oracle)
    supplementary = check_supplementary_protocol(protocol)
    check_phase_diagnostic(protocol, {"baseline": baseline, "candidate": candidate}, oracle, supplementary)
    recompute_json(ROOT / "summary.json", ROOT / "compare.py", "summary.json")
    recompute_json(ROOT / "phase-summary.json", ROOT / "derive-phases.py", "phase-summary.json")
    decision, kept = check_decision(baseline)
    final_source = candidate["source_manifest"] if kept else baseline["source_manifest"]
    if precleanup:
        check_live_bindings(baseline, candidate, final_source)
        check_assembly({"baseline": baseline, "candidate": candidate}, compare_live=True)
    return {
        "schema": SCHEMA,
        "change": CHANGE,
        "status": "pass",
        "precleanup": precleanup,
        "portable": portable,
        "sealed_files": sealed_files,
        "selected_gates": selected,
        "lanes": 24,
        "samples": 720,
        "phase_reports": 4,
        "phase_samples": 120,
        "counters": 2,
        "candidate_kept": kept,
        "decision_keys": sorted(decision),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true", help="also authenticate live binaries and final current source")
    parser.add_argument("--portable", action="store_true", help="mark this as a portable bundle replay")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify(args.precleanup, args.portable)
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        result = {
            "schema": SCHEMA,
            "change": CHANGE,
            "status": "fail",
            "precleanup": args.precleanup,
            "portable": args.portable,
            "error": str(error),
        }
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result.get("status") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
