#!/usr/bin/env python3
"""0943 DOC data-COW comparison runner.
This external driver owns the frozen schedule, source/build custody, synthetic
qualification preflight, one-shot child execution, and numerical reduction for
0943.  It never builds Rust code.  ``--check-schedule`` is read-only;
``--preflight`` validates synthetic reports and custody controls only;
``--capture`` requires a final authorized manifest and the explicit
``LITCHI_PERF_0943_CAPTURE_APPROVED=1`` environment variable.  Every formal
child is run once, its terminal receipt is durably written before report or
RSS-sidecar parsing, and all failed artifacts remain available for review.
"""
from __future__ import annotations
import argparse
import copy
import datetime as dt
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import random
import re
import signal
import shutil
import statistics
import subprocess
import sys
import time
from typing import Any, Mapping, Sequence
sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
PERF_ROOT = HERE.parent
REPO_ROOT = Path("/home/zhuhe/code/litchi").resolve()
PROTOCOL_PATH = HERE / "comparison-protocol.json"
TEMPLATE_NATIVE = HERE / "controls/synthetic-template-native.json"
TEMPLATE_ALLOCATOR = HERE / "controls/synthetic-template-allocator.json"
VALIDATOR_PATH = REPO_ROOT / "docs/performance/results/change-0917/capture.py"
RUNNER_PATH = HERE / "comparison-runner.py"
DEFAULT_FORMAL_ROOT = PERF_ROOT / "formal-authorized-v1"
DEFAULT_PREFLIGHT_ROOT = PERF_ROOT / "preflight-v1"
APPROVAL_ENV = "LITCHI_PERF_0943_CAPTURE_APPROVED"
APPROVAL_VALUE = "1"
QUALIFICATION_APPROVAL_ENV = "LITCHI_PERF_0943_QUALIFY_APPROVED"
MAX_ARTIFACT_BYTES = 1 << 20
SOURCE_REVISION = "13562481317d3d219c6fcd93babd9377e84ea834"
EDITS = (1, 2, 8, 32)
FIXTURES = ("docnohf", "docfloat")
ROUTES = ("zero", "default", "release")
VERSIONS = ("baseline", "candidate")
LANES = ("native", "allocator", "all")
NATIVE_CPU = 14
ALLOCATOR_CPU = 14
DEADLINE_SECONDS = 900.0
ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
}
HEX = set("0123456789abcdef")
class RunnerError(RuntimeError):
    """A schedule, custody, child, or reduction contract failed closed."""
def require(condition: bool, message: str) -> None:
    if not condition:
        raise RunnerError(message)
def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")
def parse_utc(value: object, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label}: timestamp missing")
    try:
        parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise RunnerError(f"{label}: malformed timestamp: {error}") from error
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")
    return parsed.astimezone(dt.timezone.utc)
def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()
def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()
def descriptor(path: Path, *, allow_large: bool = True) -> dict[str, Any]:
    path = Path(path).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    size = path.stat().st_size
    if not allow_large:
        require(size < MAX_ARTIFACT_BYTES, f"artifact is at least 1 MiB: {path}")
    return {"path": str(path), "bytes": size, "sha256": sha256_path(path)}
def packet_descriptor(path: Path, packet_root: Path, *, allow_large: bool = True) -> dict[str, Any]:
    """Describe an artifact using a path relative to its retained packet.
    Formal packets may be copied before an independent reduction.  Absolute
    paths in child receipts would make that replay silently consult the
    original capture tree, so every child-owned artifact is bound to the
    selected packet root by its relative path and digest.
    """
    path = Path(path).resolve()
    packet_root = Path(packet_root).resolve()
    try:
        relative = path.relative_to(packet_root)
    except ValueError as error:
        raise RunnerError(f"packet artifact is outside packet root: {path}") from error
    require(relative != Path(".") and ".." not in relative.parts,
            f"packet artifact path is unsafe: {relative}")
    identity = descriptor(path, allow_large=allow_large)
    identity["path"] = relative.as_posix()
    return identity
def _packet_descriptor_from_value(value: object, label: str, packet_root: Path,
                                  *, allow_large: bool = True) -> tuple[Path, dict[str, Any]]:
    require(isinstance(value, Mapping), f"{label}: packet descriptor missing")
    raw_path = value.get("path")
    relative = _relative_path(raw_path, f"{label} packet path")
    path = (Path(packet_root).resolve() / relative).resolve()
    require(path.is_relative_to(Path(packet_root).resolve()),
            f"{label}: packet path escapes packet root")
    actual = packet_descriptor(path, Path(packet_root), allow_large=allow_large)
    require(dict(value) == actual, f"{label}: packet artifact changed")
    return path, actual
def write_exclusive(path: Path, value: bytes) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    require(not path.exists(), f"refusing to replace retained artifact: {path}")
    temporary = path.with_name(path.name + f".tmp-{os.getpid()}")
    require(not temporary.exists(), f"temporary artifact already exists: {temporary}")
    try:
        with temporary.open("xb") as stream:
            stream.write(value)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
def write_json(path: Path, value: Any, *, allow_large: bool = False) -> None:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if not allow_large:
        require(len(data) < MAX_ARTIFACT_BYTES, f"JSON artifact is at least 1 MiB: {path}")
    write_exclusive(path, data)
def json_load(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return strict_json_load_bytes(path.read_bytes(), str(path))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"invalid JSON {path}: {error}") from error
def strict_json_load_bytes(data: bytes, label: str) -> Any:
    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"{label}: duplicate JSON key {key!r}")
            result[key] = value
        return result
    try:
        return json.loads(data.decode("utf-8"), object_pairs_hook=reject_duplicates)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"{label}: strict JSON parse failed: {error}") from error
def digest(value: object, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label}: invalid SHA-256")
    return value
def revision(value: object, label: str) -> str:
    require(isinstance(value, str) and len(value) == 40 and set(value) <= HEX,
            f"{label}: invalid source revision")
    return value
def integer(value: object, label: str, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: invalid integer")
    return value
def protocol() -> dict[str, Any]:
    value = json_load(PROTOCOL_PATH)
    require(isinstance(value, dict), "protocol is not an object")
    require(value.get("schema") == "litchi.performance.0943.doc-comparison.protocol.v1",
            "protocol schema changed")
    require(value.get("status") == "prospective", "protocol is no longer prospective")
    require(value.get("baseline_source_revision") == SOURCE_REVISION,
            "baseline source revision changed")
    require(value.get("candidate_source_revision") is None
            and value.get("candidate_source_root") is None,
            "prospective protocol binds a candidate")
    require(value.get("edits") == list(EDITS), "edit matrix changed")
    require(value.get("retention_routes") == list(ROUTES), "route matrix changed")
    fixtures = value.get("fixtures")
    require(isinstance(fixtures, list) and [row.get("case") for row in fixtures] == list(FIXTURES),
            "fixture order changed")
    expected_fixtures = {
        "docnohf": ("test-data/ole/doc/NoHeadFoot.doc", 26112,
                     "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48"),
        "docfloat": ("test-data/ole/doc/FloatingPictures.doc", 335360,
                      "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"),
    }
    for row in fixtures:
        require(isinstance(row, dict), "fixture row is malformed")
        case = row.get("case")
        require(case in expected_fixtures, f"unknown fixture: {case!r}")
        expected = expected_fixtures[case]
        require((row.get("relative_path"), row.get("bytes"), row.get("sha256")) == expected,
                f"fixture identity changed: {case}")
    native = value.get("native")
    require(isinstance(native, dict)
            and native.get("cpu") == NATIVE_CPU
            and native.get("rounds") == 12
            and native.get("retention") == "default"
            and native.get("warmups_per_child") == 5
            and native.get("samples_per_child") == 50
            and native.get("children") == 192
            and native.get("paired_cells") == 96,
            "native schedule changed")
    allocator = value.get("allocator")
    require(isinstance(allocator, dict)
            and allocator.get("cpu") == ALLOCATOR_CPU
            and allocator.get("repetitions") == 2
            and allocator.get("warmups_per_child") == 0
            and allocator.get("samples_per_child") == 1
            and allocator.get("children") == 96
            and allocator.get("paired_cells") == 48,
            "allocator schedule changed")
    combined = value.get("combined_capture")
    require(isinstance(combined, dict)
            and combined.get("children") == 288
            and combined.get("order") ==
            "native schedule followed by allocator schedule in one serial evidence root"
            and combined.get("required_cli_lane") == "all",
            "combined capture schedule changed")
    qualification = value.get("qualification")
    require(isinstance(qualification, dict)
            and qualification.get("children") == 64
            and qualification.get("native_children") == 16
            and qualification.get("allocator_children") == 48
            and qualification.get("native_warmups_per_child") == 0
            and qualification.get("native_samples_per_child") == 1
            and qualification.get("allocator_warmups_per_child") == 0
            and qualification.get("allocator_samples_per_child") == 1
            and qualification.get("claim_scope") ==
            "actual binary custody and correctness preflight only; no timing, allocation, RSS, or adoption claim"
            and qualification.get("separate_root_before_formal") is True,
            "qualification schedule changed")
    bootstrap = value.get("bootstrap")
    require(isinstance(bootstrap, dict)
            and bootstrap.get("draws") == 10000
            and bootstrap.get("seed") == 943
            and bootstrap.get("minimum_lower_ratio") == 1 / 1.05,
            "bootstrap contract changed")
    flags = value.get("review_flags")
    require(isinstance(flags, dict)
            and flags.get("p50_increase_fraction") == 0.05
            and flags.get("p95_increase_fraction") == 0.05
            and flags.get("p99_increase_fraction") == 0.05
            and flags.get("rss_increase_fraction") == 0.05
            and flags.get("allocator_gross_requested_increase_fraction") == 0.01
            and flags.get("allocator_peak_live_increase_fraction") == 0.05
            and flags.get("automatic_waiver") is False,
            "review flag contract changed")
    custody = value.get("custody")
    require(isinstance(custody, dict)
            and custody.get("child_deadline_seconds") == 900
            and custody.get("retry") is False
            and custody.get("serial") is True
            and custody.get("terminal_written_before_report_parse") is True
            and custody.get("max_report_bytes") == MAX_ARTIFACT_BYTES,
            "custody contract changed")
    max_rss = custody.get("max_rss")
    require(isinstance(max_rss, dict)
            and max_rss.get("argv") == [
                "/usr/bin/time", "-f", "%M", "-o",
                "<capture-root>/children/<label>/max-rss-kib.txt",
            ]
            and max_rss.get("unit") == "KiB"
            and max_rss.get("parse") == "only after the terminal receipt is durably written"
            and max_rss.get("raw_sidecar_required") is True,
            "RSS custody contract changed")
    require(custody.get("child_cwd") ==
            "bound immutable source root for the selected version; packet artifacts are outside cwd",
            "child cwd custody contract changed")
    require(custody.get("packet_paths") ==
            "child artifact descriptors are relative to the selected capture root; the executed RSS argv is absolute and normalized through recorded capture_origin; binary, source, and fixture identities remain absolute bound descriptors",
            "packet path custody contract changed")
    return value
def protocol_descriptor() -> dict[str, Any]:
    return descriptor(PROTOCOL_PATH, allow_large=False)
def script_descriptor(path: Path) -> dict[str, Any]:
    return descriptor(path, allow_large=False)
def _relative_path(value: object, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path missing")
    relative = Path(value)
    require(not relative.is_absolute() and ".." not in relative.parts and relative != Path("."),
            f"{label}: unsafe relative path")
    return relative
def source_listing(root: Path, listing_path: Path) -> tuple[int, int]:
    listed: dict[str, tuple[int, str]] = {}
    with listing_path.open("r", encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, start=1):
            try:
                row = strict_json_load_bytes(line.encode("utf-8"), f"source listing line {line_number}")
            except json.JSONDecodeError as error:
                raise RunnerError(f"source listing line {line_number} is invalid JSON: {error}") from error
            require(isinstance(row, dict), f"source listing line {line_number} is not an object")
            relative = _relative_path(row.get("path"), f"source listing line {line_number}")
            key = relative.as_posix()
            require(key not in listed, f"source listing duplicate path: {key}")
            listed[key] = (integer(row.get("bytes"), f"source listing {key} bytes"),
                           digest(row.get("sha256"), f"source listing {key} hash"))
    actual: dict[str, tuple[int, str]] = {}
    for directory, directories, files in os.walk(root, followlinks=False):
        directory_path = Path(directory)
        for name in directories + files:
            path = directory_path / name
            require(not path.is_symlink(), f"source tree contains symlink: {path}")
        for name in files:
            path = directory_path / name
            require(path.is_file(), f"source tree entry is not regular: {path}")
            relative = path.relative_to(root).as_posix()
            require(relative not in actual, f"source tree duplicate path: {relative}")
            actual[relative] = (path.stat().st_size, sha256_path(path))
    require(actual == listed,
            f"source tree/listing mismatch: actual={len(actual)} listed={len(listed)}")
    return len(actual), sum(size for size, _ in actual.values())
def _descriptor_from_value(value: object, label: str, *, allow_large: bool = True) -> tuple[Path, dict[str, Any]]:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}: descriptor path missing")
    path = Path(raw_path).resolve()
    actual = descriptor(path, allow_large=allow_large)
    require(value.get("bytes") == actual["bytes"] and value.get("sha256") == actual["sha256"],
            f"{label}: descriptor changed")
    return path, actual
def _source_receipt_listing(receipt: Mapping[str, Any], label: str) -> Mapping[str, Any]:
    candidate = receipt.get("listing", receipt.get("full_source_listing"))
    require(isinstance(candidate, Mapping), f"{label}: listing binding missing")
    return candidate
def validate_source_binding(binding: Mapping[str, Any], *, label: str,
                            expected_revision: str | None = None) -> dict[str, Any]:
    root_value = binding.get("source_root", binding.get("root"))
    require(isinstance(root_value, str) and root_value, f"{label}: source root missing")
    root = Path(root_value).resolve()
    require(root.is_dir() and not root.is_symlink(), f"{label}: source root missing: {root}")
    receipt_path, receipt_desc = _descriptor_from_value(binding.get("source_receipt"),
                                                         f"{label} source receipt",
                                                         allow_large=False)
    receipt = json_load(receipt_path)
    require(isinstance(receipt, dict), f"{label}: source receipt is not an object")
    source_revision = receipt.get("source_revision", receipt.get("revision", receipt.get("head")))
    source_revision = revision(source_revision, f"{label}: source revision")
    if expected_revision is not None:
        require(source_revision == expected_revision, f"{label}: source revision changed")
    receipt_root = receipt.get("source_root", receipt.get("root"))
    require(isinstance(receipt_root, str) and Path(receipt_root).resolve() == root,
            f"{label}: receipt root does not match binding")
    listing_value = _source_receipt_listing(receipt, label)
    listing_path, listing_desc = _descriptor_from_value(
        binding.get("full_source_listing", listing_value), f"{label} source listing")
    receipt_listing_path, receipt_listing_desc = _descriptor_from_value(
        listing_value, f"{label} receipt listing")
    require(receipt_listing_path == listing_path and receipt_listing_desc == listing_desc,
            f"{label}: receipt listing does not match binding")
    count, total = source_listing(root, listing_path)
    for field, observed in (("files", count), ("source_file_count", count),
                            ("bytes", total), ("source_bytes", total)):
        if field in receipt:
            require(receipt[field] == observed, f"{label}: receipt {field} changed")
    return {
        "source_root": root,
        "source_revision": source_revision,
        "source_receipt": receipt_desc,
        "source_receipt_path": receipt_path,
        "full_source_listing": listing_desc,
        "full_source_listing_path": listing_path,
        "file_count": count,
        "bytes": total,
    }
def _find_build_field(record: Mapping[str, Any], key: str, label: str) -> object:
    value = record.get(key)
    require(value is not None, f"{label}: {key} missing")
    return value
def validate_build_binding(binding: Mapping[str, Any], *, label: str,
                           source: Mapping[str, Any], expected_binary_name: str) -> dict[str, Any]:
    for field in ("source_root", "source_revision"):
        if field in binding:
            require(binding[field] == (str(source["source_root"])
                                       if field == "source_root" else source["source_revision"]),
                    f"{label}: {field} is not source-bound")
    start_path, start_desc = _descriptor_from_value(_find_build_field(binding, "start", label),
                                                     f"{label} start", allow_large=False)
    terminal_path, terminal_desc = _descriptor_from_value(_find_build_field(binding, "terminal", label),
                                                           f"{label} terminal", allow_large=False)
    binary_receipt_path, binary_receipt_desc = _descriptor_from_value(
        _find_build_field(binding, "binary_receipt", label), f"{label} binary receipt", allow_large=False)
    binary_path, binary_desc = _descriptor_from_value(_find_build_field(binding, "binary", label),
                                                      f"{label} binary")
    start = json_load(start_path)
    terminal = json_load(terminal_path)
    binary_receipt = json_load(binary_receipt_path)
    require(isinstance(start, dict) and isinstance(terminal, dict) and isinstance(binary_receipt, dict),
            f"{label}: build receipts must be objects")
    require(start.get("cwd") == str(source["source_root"])
            and terminal.get("cwd") == str(source["source_root"]),
            f"{label}: build cwd is not source root")
    require(start.get("argv") == terminal.get("argv"), f"{label}: build argv changed")
    require(terminal.get("exit_code") == 0, f"{label}: build failed")
    require(terminal.get("terminal_written_before_binary_descriptor") is True,
            f"{label}: build terminal ordering missing")
    argv = start.get("argv")
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
            f"{label}: build argv missing")
    require("--release" in argv and "--offline" in argv and "--locked" in argv and "-j2" in argv,
            f"{label}: build command is not the reviewed locked release build")
    require("--bin" in argv and argv[argv.index("--bin") + 1] == expected_binary_name,
            f"{label}: binary identity is not bound")
    source_receipt = start.get("source_receipt", terminal.get("source_receipt"))
    if isinstance(source_receipt, Mapping):
        require(Path(str(source_receipt.get("path", ""))).resolve()
                == Path(str(source["source_receipt_path"])).resolve(),
                f"{label}: build source receipt path changed")
    require(binary_receipt.get("path") in (None, str(binary_path)),
            f"{label}: binary receipt path changed")
    if binary_receipt.get("bytes") is not None:
        require(binary_receipt["bytes"] == binary_desc["bytes"], f"{label}: binary receipt size changed")
    if binary_receipt.get("sha256") is not None:
        require(binary_receipt["sha256"] == binary_desc["sha256"], f"{label}: binary receipt hash changed")
    require(binding.get("binary_identity", expected_binary_name) == expected_binary_name,
            f"{label}: binary identity changed")
    return {
        "start": start_desc,
        "start_path": start_path,
        "terminal": terminal_desc,
        "terminal_path": terminal_path,
        "binary_receipt": binary_receipt_desc,
        "binary_receipt_path": binary_receipt_path,
        "binary": binary_desc,
        "binary_path": binary_path,
        "binary_identity": expected_binary_name,
        "source_root": source["source_root"],
        "source_revision": source["source_revision"],
    }
def fixture_rows(p: Mapping[str, Any], source_root: Path) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for row in p["fixtures"]:
        relative = _relative_path(row["relative_path"], f"fixture {row.get('case')}")
        path = (source_root / relative).resolve()
        require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
        require(path.stat().st_size == row["bytes"] and sha256_path(path) == row["sha256"],
                f"fixture changed: {row.get('case')}")
        rows[str(row["case"])] = {**row, "path": relative.as_posix(), "path_abs": str(path)}
    return rows
def safe_label(value: str) -> str:
    require(re.fullmatch(r"[a-z0-9-]+", value) is not None, f"unsafe child label: {value}")
    return value
def native_schedule() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for round_no in range(1, 13):
        cases = list(FIXTURES) if round_no % 2 else list(reversed(FIXTURES))
        counts = list(EDITS) if round_no % 2 else list(reversed(EDITS))
        versions = list(VERSIONS) if round_no % 2 else list(reversed(VERSIONS))
        for case in cases:
            for edits in counts:
                pair_id = f"native-r{round_no:02d}-{case}-e{edits:02d}"
                for version in versions:
                    rows.append({
                        "lane": "native", "round": round_no, "case": case, "edits": edits,
                        "retention": "default", "version": version, "pair_id": pair_id,
                        "label": f"{pair_id}-{version}", "warmups": 5, "samples": 50,
                        "cpu": NATIVE_CPU, "deadline_seconds": int(DEADLINE_SECONDS),
                    })
    require(len(rows) == 192 and len({row["label"] for row in rows}) == 192,
            "native schedule count or uniqueness changed")
    return rows
def allocator_schedule() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    cells = [(case, edits, route) for case in FIXTURES for edits in EDITS for route in ROUTES]
    for repetition in (1, 2):
        ordered = cells if repetition == 1 else list(reversed(cells))
        versions = list(VERSIONS) if repetition == 1 else list(reversed(VERSIONS))
        for case, edits, route in ordered:
            pair_id = f"allocator-r{repetition:02d}-{case}-e{edits:02d}-{route}"
            for version in versions:
                rows.append({
                    "lane": "allocator", "repetition": repetition, "case": case,
                    "edits": edits, "retention": route, "version": version,
                    "pair_id": pair_id, "label": f"{pair_id}-{version}",
                    "warmups": 0, "samples": 1, "cpu": ALLOCATOR_CPU,
                    "deadline_seconds": int(DEADLINE_SECONDS),
                })
    require(len(rows) == 96 and len({row["label"] for row in rows}) == 96,
            "allocator schedule count or uniqueness changed")
    return rows
def schedule(lane: str) -> list[dict[str, Any]]:
    require(lane in LANES, f"unknown lane: {lane}")
    if lane == "native":
        return native_schedule()
    if lane == "allocator":
        return allocator_schedule()
    return all_schedule()
def all_schedule() -> list[dict[str, Any]]:
    return native_schedule() + allocator_schedule()


def qualification_schedule() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for case in FIXTURES:
        for edits in EDITS:
            pair_id = f"qualification-native-{case}-e{edits:02d}"
            for version in VERSIONS:
                rows.append({
                    "lane": "native", "round": 1, "case": case, "edits": edits,
                    "retention": "default", "version": version, "pair_id": pair_id,
                    "label": f"{pair_id}-{version}", "warmups": 0, "samples": 1,
                    "cpu": NATIVE_CPU, "deadline_seconds": int(DEADLINE_SECONDS),
                })
    for case in FIXTURES:
        for edits in EDITS:
            for route in ROUTES:
                pair_id = f"qualification-allocator-{case}-e{edits:02d}-{route}"
                for version in VERSIONS:
                    rows.append({
                        "lane": "allocator", "repetition": 1, "case": case,
                        "edits": edits, "retention": route, "version": version,
                        "pair_id": pair_id, "label": f"{pair_id}-{version}",
                        "warmups": 0, "samples": 1, "cpu": ALLOCATOR_CPU,
                        "deadline_seconds": int(DEADLINE_SECONDS),
                    })
    require(len(rows) == 64 and len({row["label"] for row in rows}) == 64,
            "qualification schedule count or uniqueness changed")
    return rows
def schedule_sha256(rows: Sequence[Mapping[str, Any]]) -> str:
    data = json.dumps(list(rows), sort_keys=True, separators=(",", ":")).encode("utf-8")
    return sha256_bytes(data)
def _load_validator(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"validator missing: {path}")
    spec = importlib.util.spec_from_file_location("litchi_0943_validator", path)
    require(spec is not None and spec.loader is not None, f"cannot load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    require(callable(getattr(module, "validate_report", None)),
            "validator has no validate_report function")
    return module
def _manifest_descriptor(value: object, label: str, *, allow_missing: bool = False) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label}: descriptor missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: descriptor path missing")
    path = Path(path_value).resolve()
    if allow_missing and not path.exists():
        integer(value.get("bytes"), f"{label}: bytes", 1)
        digest(value.get("sha256"), f"{label}: hash")
        return dict(value)
    actual = descriptor(path)
    require(dict(value) == actual, f"{label}: descriptor changed")
    return actual

def _manifest_source(manifest: Mapping[str, Any], version: str) -> dict[str, Any]:
    source = manifest.get("sources", {}).get(version) if isinstance(manifest.get("sources"), Mapping) else None
    require(isinstance(source, Mapping), f"manifest source missing: {version}")
    expected = SOURCE_REVISION if version == "baseline" else None
    binding = validate_source_binding(source, label=f"{version} source", expected_revision=expected)
    if version == "candidate":
        require(binding["source_revision"] != "", "candidate source revision missing")
    return binding


def validate_manifest(manifest_path: Path, *, authorization_kind: str = "capture") -> dict[str, Any]:
    manifest = json_load(manifest_path)
    require(isinstance(manifest, dict), "manifest is not an object")
    require(manifest.get("schema") == "litchi.performance.0943.doc-comparison.manifest.v1",
            "manifest schema changed")
    p = protocol()
    require(manifest.get("protocol_sha256") == sha256_path(PROTOCOL_PATH),
            "manifest protocol digest changed")
    _manifest_descriptor(manifest.get("protocol"), "manifest protocol")
    _manifest_descriptor(manifest.get("runner"), "manifest runner")
    _manifest_descriptor(manifest.get("validator"), "manifest validator")
    require(Path(str(manifest["runner"]["path"])).resolve() == RUNNER_PATH,
            "manifest runner path changed")
    require(Path(str(manifest["validator"]["path"])).resolve() == VALIDATOR_PATH,
            "manifest validator path changed")
    sources = manifest.get("sources")
    builds = manifest.get("builds")
    require(isinstance(sources, Mapping) and isinstance(builds, Mapping),
            "manifest source/build mappings missing")
    source_bindings = {version: _manifest_source(manifest, version) for version in VERSIONS}
    build_bindings: dict[str, dict[str, Any]] = {}
    expected_names = {"native": "doc_retention_probe", "allocator": "doc_retention_probe_alloc"}
    for version in VERSIONS:
        source = source_bindings[version]
        build_rows = builds.get(version)
        require(isinstance(build_rows, Mapping), f"manifest builds missing: {version}")
        build_bindings[version] = {}
        for lane, binary_name in expected_names.items():
            entry = build_rows.get(lane)
            require(isinstance(entry, Mapping), f"manifest build missing: {version}/{lane}")
            build_bindings[version][lane] = validate_build_binding(
                entry, label=f"{version}/{lane} build", source=source,
                expected_binary_name=binary_name)
    environment = manifest.get("child_environment")
    require(environment == ENVIRONMENT, "manifest child environment changed")
    require(manifest.get("schedule_sha256", {}).get("native") == schedule_sha256(native_schedule()),
            "manifest native schedule digest changed")
    require(manifest.get("schedule_sha256", {}).get("allocator") == schedule_sha256(allocator_schedule()),
            "manifest allocator schedule digest changed")
    require(manifest.get("schedule_sha256", {}).get("combined") == schedule_sha256(all_schedule()),
            "manifest combined schedule digest changed")
    authorization = manifest.get("authorization")
    authorization_key = ("capture_authorized" if authorization_kind == "capture"
                         else "qualification_authorized")
    require(isinstance(authorization, Mapping) and authorization.get(authorization_key) is True,
            f"manifest is not authorized for {authorization_kind}")
    external = Path(str(manifest.get("external_root", ""))).resolve()
    require(external != REPO_ROOT and REPO_ROOT not in external.parents,
            "formal evidence root must be outside checkout")
    return {
        "manifest": manifest,
        "protocol": p,
        "source_bindings": source_bindings,
        "build_bindings": build_bindings,
        "external_root": external,
        "validator": VALIDATOR_PATH,
    }


def binary_for(binding: Mapping[str, Any], version: str, lane: str) -> Path:
    return Path(str(binding["build_bindings"][version][lane]["binary_path"])).resolve()


def fixture_for(binding: Mapping[str, Any], version: str, case: str) -> dict[str, Any]:
    return fixture_rows(binding["protocol"], binding["source_bindings"][version]["source_root"])[case]


def child_command(spec: Mapping[str, Any], binary: Path, fixture: Mapping[str, Any],
                  rss_path: Path) -> tuple[str, ...]:
    cpu = NATIVE_CPU if spec["lane"] == "native" else ALLOCATOR_CPU
    return (
        "/usr/bin/time", "-f", "%M", "-o", str(Path(rss_path).resolve()),
        "/usr/bin/taskset", "--cpu-list", str(cpu), str(binary),
        "--case", str(spec["case"]), "--input", str(fixture["path_abs"]),
        "--edits", str(spec["edits"]), "--retention", str(spec["retention"]),
        "--warmup", str(spec["warmups"]), "--samples", str(spec["samples"]),
    )


def packet_command_view(command: Sequence[str], capture_origin: Path) -> list[str]:
    """Normalize packet-local absolute argv tokens for moved-packet replay."""
    origin = str(Path(capture_origin).resolve())
    result: list[str] = []
    for token in command:
        require(isinstance(token, str), "child command contains a non-string token")
        if token == origin:
            result.append("<packet>")
        elif token.startswith(origin + "/"):
            result.append("<packet>/" + token[len(origin) + 1:])
        else:
            result.append(token)
    return result


def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON key: {key}")
        result[key] = value
    return result


def read_report(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing child report: {path}")
    require(path.stat().st_size < MAX_ARTIFACT_BYTES, f"child report is at least 1 MiB: {path}")
    raw = path.read_bytes()
    try:
        value = json.loads(raw.decode("utf-8"), object_pairs_hook=reject_duplicates)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"child report is not strict JSON: {error}") from error
    require(isinstance(value, dict), "child report must be an object")
    return value


def parse_rss(path: Path) -> int:
    require(path.is_file() and not path.is_symlink(), f"missing RSS sidecar: {path}")
    raw = path.read_text(encoding="utf-8")
    require(len(raw.encode("utf-8")) < MAX_ARTIFACT_BYTES, f"RSS sidecar is too large: {path}")
    match = re.fullmatch(r"\s*([0-9]+)\s*\n?", raw)
    require(match is not None, f"RSS sidecar is not one KiB integer: {path}")
    return int(match.group(1))


def process_group_exists(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def close_process_group(pgid: int) -> tuple[bool, bool, int]:
    forced = process_group_exists(pgid)
    if forced:
        try:
            os.killpg(pgid, signal.SIGTERM)
        except ProcessLookupError:
            pass
    for probe in range(1, 101):
        if not process_group_exists(pgid):
            return True, forced, probe
        if probe == 50 and forced:
            try:
                os.killpg(pgid, signal.SIGKILL)
            except ProcessLookupError:
                pass
        time.sleep(0.01)
    return not process_group_exists(pgid), forced, 100


def prior_rows(rows: Sequence[Mapping[str, Any]], spec: Mapping[str, Any]) -> Sequence[Mapping[str, Any]]:
    index = next(index for index, row in enumerate(rows) if row["label"] == spec["label"])
    return rows[:index]


def receipt_path(external: Path, spec: Mapping[str, Any]) -> Path:
    return external / "receipts" / f"{safe_label(str(spec['label']))}.json"


def run_one(binding: Mapping[str, Any], spec: Mapping[str, Any], rows: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    external = binding["external_root"]
    child = external / "children" / safe_label(str(spec["label"]))
    receipt = receipt_path(external, spec)
    require(not child.exists() and not receipt.exists(), f"refusing to replace child: {spec['label']}")
    for previous in prior_rows(rows, spec):
        previous_receipt = receipt_path(external, previous)
        require(previous_receipt.is_file(), f"schedule gap before {spec['label']}: {previous['label']}")
        previous_value = json_load(previous_receipt)
        require(previous_value.get("status") == "passed", f"prior child failed: {previous['label']}")
        require(parse_utc(previous_value.get("started_utc"), f"{previous['label']} start")
                <= parse_utc(previous_value.get("finished_utc"), f"{previous['label']} finish"),
                f"prior child timestamps invalid: {previous['label']}")
    child.mkdir(parents=True, exist_ok=False)
    report_path = child / "report.jsonl"
    stderr_path = child / "stderr.log"
    rss_path = child / "max-rss-kib.txt"
    terminal_path = child / "terminal.json"
    validation_path = child / "validation.json"
    binary = binary_for(binding, str(spec["version"]), str(spec["lane"]))
    fixture = fixture_for(binding, str(spec["version"]), str(spec["case"]))
    command = child_command(spec, binary, fixture, rss_path)
    capture_origin = external.resolve()
    command_packet_view = packet_command_view(command, capture_origin)
    expected_packet_rss = f"<packet>/children/{safe_label(str(spec['label']))}/max-rss-kib.txt"
    require(command_packet_view.count(expected_packet_rss) == 1,
            f"{spec['label']}: RSS command path is not a strict packet-relative child path")
    started = now()
    monotonic_start = time.monotonic()
    process: subprocess.Popen[bytes] | None = None
    exit_code: int | None = None
    deadline_exceeded = False
    runner_error: str | None = None
    timeout_count = 0
    with report_path.open("xb") as report_stream, stderr_path.open("xb") as stderr_stream:
        try:
            process = subprocess.Popen(
                list(command),
                cwd=str(binding["source_bindings"][str(spec["version"])] ["source_root"]),
                env=dict(ENVIRONMENT), stdin=subprocess.DEVNULL,
                stdout=report_stream, stderr=stderr_stream, start_new_session=True, close_fds=True)
            while True:
                try:
                    exit_code = process.wait(timeout=0.25)
                    break
                except subprocess.TimeoutExpired:
                    timeout_count += 1
                    # Keep an explicit monotonic deadline without trusting child output.
                    if time.monotonic() - monotonic_start <= DEADLINE_SECONDS:
                        continue
                    deadline_exceeded = True
                    try:
                        os.killpg(process.pid, signal.SIGTERM)
                    except ProcessLookupError:
                        pass
                    try:
                        exit_code = process.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        try:
                            os.killpg(process.pid, signal.SIGKILL)
                        except ProcessLookupError:
                            pass
                        exit_code = process.wait(timeout=5)
                    break
        except Exception as error:
            runner_error = f"{type(error).__name__}: {error}"
    finished = now()
    group_closed, forced_cleanup, group_probes = (True, False, 0)
    if process is not None:
        group_closed, forced_cleanup, group_probes = close_process_group(process.pid)
    report_desc = packet_descriptor(report_path, external, allow_large=True) if report_path.is_file() else None
    stderr_desc = packet_descriptor(stderr_path, external, allow_large=True) if stderr_path.is_file() else None
    rss_desc = (packet_descriptor(rss_path, external, allow_large=False)
                if rss_path.is_file() else None)
    terminal = {
        "schema": "litchi.performance.0943.child-terminal.v1",
        "status": "terminal",
        "label": spec["label"],
        "spec": dict(spec),
        "argv": list(command),
        "packet_command": command_packet_view,
        "capture_origin": str(capture_origin),
        "cwd": str(binding["source_bindings"][str(spec["version"])] ["source_root"]),
        "environment": dict(ENVIRONMENT),
        "started_utc": started,
        "ended_utc": finished,
        "exit_code": exit_code,
        "deadline_seconds": int(DEADLINE_SECONDS),
        "deadline_exceeded": deadline_exceeded,
        "timeout_poll_count": timeout_count,
        "runner_error": runner_error,
        "process_group_id": process.pid if process is not None else None,
        "process_group_closed": group_closed,
        "process_group_forced_cleanup": forced_cleanup,
        "process_group_probe_count": group_probes,
        "report": report_desc,
        "stderr": stderr_desc,
        "max_rss_kib": rss_desc,
        "terminal_written_before_report_parse": True,
        "source_revision": binding["source_bindings"][str(spec["version"])]["source_revision"],
        "binary": binding["build_bindings"][str(spec["version"])][str(spec["lane"])]["binary"],
    }
    # The terminal receipt is the custody boundary.  No report or RSS value is
    # parsed before this create-only, fsynced write completes.
    write_json(terminal_path, terminal)

    validation: dict[str, Any]
    status = "passed"
    report: dict[str, Any] | None = None
    try:
        require(exit_code == 0, f"child exited with code {exit_code!r}")
        require(not deadline_exceeded and runner_error is None, "child did not reach a clean terminal")
        require(group_closed and not forced_cleanup, "child process group required cleanup")
        require(report_desc is not None, "child report missing")
        require(rss_desc is not None, "RSS sidecar missing")
        # These two reads intentionally occur after terminal.json exists.
        rss_kib = parse_rss(rss_path)
        report = read_report(report_path)
        validator = binding["validator_module"]
        report_summary = validator.validate_report(
            report,
            dict(spec),
            fixture,
            str(spec["lane"]),
            expected_warmups=int(spec["warmups"]),
            expected_samples=int(spec["samples"]),
        )
        require(isinstance(report_summary, Mapping), "validator did not return an object")
        validation = {
            "schema": "litchi.performance.0943.child-validation.v1",
            "status": "passed",
            "label": spec["label"],
            "report": report_desc,
            "rss_kib": rss_kib,
            "report_summary": dict(report_summary),
        }
    except Exception as error:
        status = "failed"
        validation = {
            "schema": "litchi.performance.0943.child-validation.v1",
            "status": "failed",
            "label": spec["label"],
            "report": report_desc,
            "rss": rss_desc,
            "error": f"{type(error).__name__}: {error}",
        }
    write_json(validation_path, validation)
    receipt_value = {
        "schema": "litchi.performance.0943.child-receipt.v1",
        "status": status,
        "spec": dict(spec),
        "schedule_sha256": schedule_sha256(rows),
        "protocol": protocol_descriptor(),
        "manifest": binding["manifest_descriptor"],
        "source": {
            "revision": binding["source_bindings"][str(spec["version"])]["source_revision"],
            "root": str(binding["source_bindings"][str(spec["version"])]["source_root"]),
        },
        "binary": binding["build_bindings"][str(spec["version"])][str(spec["lane"])]["binary"],
        "command": list(command),
        "packet_command": command_packet_view,
        "capture_origin": str(capture_origin),
        "cwd": str(binding["source_bindings"][str(spec["version"])] ["source_root"]),
        "environment": dict(ENVIRONMENT),
        "started_utc": started,
        "finished_utc": finished,
        "exit_code": exit_code,
        "terminal": packet_descriptor(terminal_path, external, allow_large=False),
        "validation": packet_descriptor(validation_path, external, allow_large=False),
        "report": report_desc,
        "stderr": stderr_desc,
        "max_rss_kib": rss_desc,
        "rss_kib": validation.get("rss_kib"),
    }
    write_json(receipt, receipt_value)
    if status != "passed":
        raise RunnerError(f"{spec['label']} failed; all artifacts retained at {child}")
    return {"spec": dict(spec), "receipt": receipt_value, "report": report,
            "validation": validation}


def load_child_receipt(binding: Mapping[str, Any], spec: Mapping[str, Any],
                       rows: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    packet_root = Path(binding["external_root"]).resolve()
    receipt = json_load(receipt_path(binding["external_root"], spec))
    require(isinstance(receipt, dict)
            and receipt.get("schema") == "litchi.performance.0943.child-receipt.v1"
            and receipt.get("status") == "passed"
            and receipt.get("spec") == dict(spec),
            f"{spec['label']}: receipt identity/status changed")
    require(receipt.get("schedule_sha256") == schedule_sha256(rows),
            f"{spec['label']}: schedule binding changed")
    require(receipt.get("protocol") == binding["protocol_descriptor"],
            f"{spec['label']}: protocol binding changed")
    require(receipt.get("manifest") == binding["manifest_descriptor"],
            f"{spec['label']}: manifest binding changed")
    started = parse_utc(receipt.get("started_utc"), f"{spec['label']} start")
    finished = parse_utc(receipt.get("finished_utc"), f"{spec['label']} finish")
    require(started <= finished, f"{spec['label']}: timestamps inverted")
    for previous in prior_rows(rows, spec):
        prior = json_load(receipt_path(binding["external_root"], previous))
        require(parse_utc(prior.get("finished_utc"), f"{previous['label']} finish") <= started,
                f"{spec['label']}: schedule overlap")
    for field in ("terminal", "validation", "report", "stderr", "max_rss_kib"):
        require(receipt.get(field) is not None, f"{spec['label']}: {field} artifact missing")
        _packet_descriptor_from_value(receipt[field], f"{spec['label']} {field}", packet_root)
    terminal_path, _ = _packet_descriptor_from_value(
        receipt["terminal"], f"{spec['label']} terminal", packet_root)
    terminal = json_load(terminal_path)
    require(terminal.get("terminal_written_before_report_parse") is True,
            f"{spec['label']}: terminal ordering missing")
    require(terminal.get("schema") == "litchi.performance.0943.child-terminal.v1"
            and terminal.get("status") == "terminal"
            and terminal.get("label") == spec["label"] and terminal.get("spec") == dict(spec),
            f"{spec['label']}: terminal identity changed")
    require(terminal.get("exit_code") == receipt.get("exit_code") == 0
            and terminal.get("deadline_seconds") == int(DEADLINE_SECONDS)
            and terminal.get("deadline_exceeded") is False
            and "runner_error" in terminal and terminal["runner_error"] is None
            and terminal.get("process_group_closed") is True
            and terminal.get("process_group_forced_cleanup") is False,
            f"{spec['label']}: unclean terminal state")
    require(terminal.get("started_utc") == receipt.get("started_utc")
            and terminal.get("ended_utc") == receipt.get("finished_utc"),
            f"{spec['label']}: terminal timestamps changed")
    require(terminal.get("environment") == receipt.get("environment") == ENVIRONMENT,
            f"{spec['label']}: child environment changed")
    source_binding = binding["source_bindings"][str(spec["version"])]
    expected_source = {"root": str(source_binding["source_root"]),
                       "revision": source_binding["source_revision"]}
    require(receipt.get("source") == expected_source
            and terminal.get("source_revision") == source_binding["source_revision"],
            f"{spec['label']}: child source identity changed")
    expected_binary = binding["build_bindings"][str(spec["version"])][str(spec["lane"])]["binary"]
    require(receipt.get("binary") == terminal.get("binary") == expected_binary,
            f"{spec['label']}: child binary identity changed")
    origin_value = receipt.get("capture_origin")
    require(isinstance(origin_value, str) and Path(origin_value).is_absolute(),
            f"{spec['label']}: capture origin missing")
    command = receipt.get("command")
    command_packet = receipt.get("packet_command")
    expected_command = list(child_command(
        spec, binary_for(binding, str(spec["version"]), str(spec["lane"])),
        fixture_for(binding, str(spec["version"]), str(spec["case"])),
        Path(origin_value) / "children" / safe_label(str(spec["label"])) / "max-rss-kib.txt"))
    require(command == expected_command, f"{spec['label']}: executed argv differs from schedule")
    require(isinstance(command, list) and isinstance(command_packet, list),
            f"{spec['label']}: command custody missing")
    require(command_packet == packet_command_view(command, Path(origin_value)),
            f"{spec['label']}: packet command normalization changed")
    expected_packet_rss = f"<packet>/children/{safe_label(str(spec['label']))}/max-rss-kib.txt"
    require(command_packet.count(expected_packet_rss) == 1,
            f"{spec['label']}: command packet RSS path changed")
    expected_cwd = str(binding["source_bindings"][str(spec["version"])] ["source_root"])
    require(terminal.get("argv") == command
            and terminal.get("packet_command") == command_packet
            and terminal.get("capture_origin") == origin_value
            and terminal.get("cwd") == receipt.get("cwd") == expected_cwd,
            f"{spec['label']}: terminal command/cwd changed")
    for field in ("report", "stderr", "max_rss_kib"):
        require(field in terminal and terminal[field] is not None,
                f"{spec['label']}: terminal {field} descriptor missing")
        _, terminal_descriptor = _packet_descriptor_from_value(
            terminal[field], f"{spec['label']} terminal {field}", packet_root)
        require(terminal_descriptor == receipt[field],
                f"{spec['label']}: terminal {field} descriptor differs from receipt")
    rss_path, _ = _packet_descriptor_from_value(
        receipt["max_rss_kib"], f"{spec['label']} RSS", packet_root)
    rss = parse_rss(rss_path)
    require(receipt.get("rss_kib") == rss, f"{spec['label']}: RSS value changed")
    report_path, report_descriptor_value = _packet_descriptor_from_value(
        receipt["report"], f"{spec['label']} report", packet_root)
    report = read_report(report_path)
    fixture = fixture_for(binding, str(spec["version"]), str(spec["case"]))
    summary = binding["validator_module"].validate_report(
        report, dict(spec), fixture, str(spec["lane"]),
        expected_warmups=int(spec["warmups"]), expected_samples=int(spec["samples"]),
    )
    validation_path, _ = _packet_descriptor_from_value(
        receipt["validation"], f"{spec['label']} validation", packet_root)
    validation = json_load(validation_path)
    require(validation.get("status") == "passed" and validation.get("report_summary") == summary,
            f"{spec['label']}: validation receipt differs from report")
    require(validation.get("report") == report_descriptor_value,
            f"{spec['label']}: validation report descriptor changed")
    return {"spec": dict(spec), "receipt": receipt, "report": report, "rss_kib": rss,
            "summary": summary}


def nearest_rank(values: Sequence[float], fraction: float) -> float:
    require(values, "empty statistic vector")
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def stats(values: Sequence[float]) -> dict[str, float]:
    require(values, "empty statistic vector")
    return {
        "p50": nearest_rank(values, 0.50),
        "p95": nearest_rank(values, 0.95),
        "p99": nearest_rank(values, 0.99),
        "mean": statistics.fmean(values),
    }


def fraction_flag(before: float, after: float, limit: float) -> dict[str, Any]:
    increase = float("inf") if before == 0 and after > 0 else (after / before - 1.0 if before else 0.0)
    return {"before": before, "after": after, "increase_fraction": increase,
            "flag": increase > limit}


def bootstrap(ratios: Sequence[float]) -> dict[str, Any]:
    require(ratios, "empty paired ratio vector")
    draws = 10000
    rng = random.Random(943)
    medians = [statistics.median(ratios[rng.randrange(len(ratios))] for _ in ratios)
               for _ in range(draws)]
    return {
        "draws": draws,
        "seed": 943,
        "ratios": list(ratios),
        "median_ratio": statistics.median(ratios),
        "confidence_interval_95": {
            "method": "nearest-rank 2.5th/97.5th percentiles",
            "lower": nearest_rank(medians, 0.025),
            "upper": nearest_rank(medians, 0.975),
        },
    }


def report_elapsed(report: Mapping[str, Any]) -> list[float]:
    values = report.get("elapsed_ns")
    require(isinstance(values, list) and values, "native elapsed vector missing")
    return [float(integer(value, "elapsed sample", 1)) for value in values]


def report_allocation(report: Mapping[str, Any]) -> dict[str, int]:
    value = report.get("allocation")
    require(isinstance(value, Mapping), "allocator allocation metrics missing")
    return {field: integer(value.get(field), f"allocation.{field}")
            for field in ("allocated_bytes", "deallocated_bytes", "allocation_calls",
                          "peak_live_bytes", "retained_bytes")}


def capacity_signature(report: Mapping[str, Any]) -> dict[str, Any]:
    routes: dict[str, Any] = {}
    for route in report.get("route_outputs", []):
        require(isinstance(route, Mapping), "route output malformed")
        name = str(route["retention"])
        require(name not in routes, f"duplicate route in capacity signature: {name}")
        routes[name] = {
            "output_capacity": route["output_capacity"],
            "final_retained_before_commit": route["final_retained_before_commit"],
            "stages": [{key: stage[key] for key in ("before", "after", "after_release")}
                       for stage in route["stages"]],
        }
    samples = []
    for sample in report.get("samples", []):
        allocation = sample.get("allocation")
        samples.append({
            "output_capacity": sample["output_capacity"],
            "retained_bytes": allocation.get("retained_bytes") if isinstance(allocation, Mapping) else None,
            "stages": [{key: stage[key] for key in ("before", "after", "after_release")}
                       for stage in sample.get("stages", sample.get("stage_observations", []))],
        })
    allocation = report.get("allocation")
    return {
        "output_capacity": report["output_capacity"],
        "retained_bytes": allocation.get("retained_bytes") if isinstance(allocation, Mapping) else None,
        "routes": routes,
        "samples": samples,
    }


def _without_metrics(value: Any) -> None:
    if not isinstance(value, dict):
        return
    for key in ("allocated_bytes", "deallocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes"):
        value.pop(key, None)
    for child in value.values():
        if isinstance(child, dict):
            _without_metrics(child)
        elif isinstance(child, list):
            for item in child:
                _without_metrics(item)


def semantic_signature(report: Mapping[str, Any]) -> Any:
    value = copy.deepcopy(dict(report))
    # Each input location is independently source-bound by validate_report.
    # Isolated exports have different paths but identical fixture digest/size.
    value.pop("input", None)
    value.pop("elapsed_ns", None)
    for sample in value.get("samples", []):
        if isinstance(sample, dict):
            sample.pop("elapsed_ns", None)
    _without_metrics(value)
    capacities = capacity_signature(report)
    def remove_capacity(node: Any) -> None:
        if not isinstance(node, dict):
            return
        for key in ("output_capacity", "final_retained_before_commit"):
            node.pop(key, None)
        for child in node.values():
            if isinstance(child, dict):
                remove_capacity(child)
            elif isinstance(child, list):
                for item in child:
                    remove_capacity(item)
    remove_capacity(value)
    # Keep the existence and shape of capacity data in a separate exact check.
    value["__capacity_shape"] = {
        "route_count": len(capacities["routes"]),
        "sample_count": len(capacities["samples"]),
    }
    return value


def compare_reports(baseline: Mapping[str, Any], candidate: Mapping[str, Any], label: str) -> dict[str, Any]:
    require(semantic_signature(baseline) == semantic_signature(candidate),
            f"{label}: non-metric semantic identity differs")
    base = capacity_signature(baseline)
    cand = capacity_signature(candidate)
    require(base == cand, f"{label}: explicit capacity/retention fields differ")
    return {"all_explicit_capacity_fields_equal": True,
            "capacity_sha256": sha256_bytes(json.dumps(base, sort_keys=True).encode())}


def pair_rows(items: Sequence[Mapping[str, Any]]) -> dict[tuple[Any, ...], dict[str, Mapping[str, Any]]]:
    result: dict[tuple[Any, ...], dict[str, Mapping[str, Any]]] = {}
    for item in items:
        spec = item["spec"]
        key = ((spec["round"], spec["case"], spec["edits"])
               if spec["lane"] == "native"
               else (spec["repetition"], spec["case"], spec["edits"], spec["retention"]))
        require(spec["version"] not in result.setdefault(key, {}), f"duplicate result pair: {key}")
        result[key][str(spec["version"])] = item
    return result


def reduce_native(binding: Mapping[str, Any]) -> dict[str, Any]:
    rows = all_schedule()
    lane_rows = native_schedule()
    loaded = [load_child_receipt(binding, spec, rows) for spec in lane_rows]
    pairs = pair_rows(loaded)
    cells: list[dict[str, Any]] = []
    for case in FIXTURES:
        for edits in EDITS:
            rounds: list[dict[str, Any]] = []
            ratios: list[float] = []
            for round_no in range(1, 13):
                pair = pairs[(round_no, case, edits)]
                require(set(pair) == set(VERSIONS), f"native pair incomplete: {case}/e{edits}/r{round_no}")
                base = pair["baseline"]
                cand = pair["candidate"]
                compare = compare_reports(base["report"], cand["report"],
                                          f"native {case}/e{edits}/r{round_no}")
                base_stats = stats(report_elapsed(base["report"]))
                cand_stats = stats(report_elapsed(cand["report"]))
                ratio = base_stats["p50"] / cand_stats["p50"]
                ratios.append(ratio)
                rounds.append({
                    "round": round_no,
                    "baseline": {"stats": base_stats, "rss_kib": base["rss_kib"]},
                    "candidate": {"stats": cand_stats, "rss_kib": cand["rss_kib"]},
                    "p50_ratio_baseline_over_candidate": ratio,
                    "latency": {field: fraction_flag(base_stats[field], cand_stats[field], 0.05)
                                 for field in ("p50", "p95", "p99")},
                    "rss": fraction_flag(float(base["rss_kib"]), float(cand["rss_kib"]), 0.05),
                    "exact": compare,
                })
            boot = bootstrap(ratios)
            lower = float(boot["confidence_interval_95"]["lower"])
            cells.append({
                "case": case, "edits": edits, "retention": "default", "rounds": rounds,
                "baseline_pooled": stats([x for round_ in rounds for x in
                                            report_elapsed(pairs[(round_["round"], case, edits)]["baseline"]["report"])]),
                "candidate_pooled": stats([x for round_ in rounds for x in
                                             report_elapsed(pairs[(round_["round"], case, edits)]["candidate"]["report"])]),
                "paired_bootstrap": boot,
                "latency_gate": lower >= 1 / 1.05,
            })
    return {"schema": "litchi.performance.0943.native-reduction.v1",
            "children": len(loaded), "paired_cells": 96, "cells": cells,
            "claim": "paired native timing comparison; every p50/p95/p99 and RSS review flag retained"}


def reduce_allocator(binding: Mapping[str, Any]) -> dict[str, Any]:
    rows = all_schedule()
    lane_rows = allocator_schedule()
    loaded = [load_child_receipt(binding, spec, rows) for spec in lane_rows]
    pairs = pair_rows(loaded)
    cells: list[dict[str, Any]] = []
    for key in sorted(pairs):
        pair = pairs[key]
        require(set(pair) == set(VERSIONS), f"allocator pair incomplete: {key}")
        base = pair["baseline"]
        cand = pair["candidate"]
        compare = compare_reports(base["report"], cand["report"], f"allocator {key}")
        base_alloc = report_allocation(base["report"])
        cand_alloc = report_allocation(cand["report"])
        cells.append({
            "repetition": key[0], "case": key[1], "edits": key[2], "retention": key[3],
            "baseline": base_alloc, "candidate": cand_alloc, "exact": compare,
            "allocation": fraction_flag(float(base_alloc["allocated_bytes"]),
                                         float(cand_alloc["allocated_bytes"]), 0.01),
            "peak_live": fraction_flag(float(base_alloc["peak_live_bytes"]),
                                        float(cand_alloc["peak_live_bytes"]), 0.05),
            "rss": fraction_flag(float(base["rss_kib"]), float(cand["rss_kib"]), 0.05),
        })
    return {"schema": "litchi.performance.0943.allocator-reduction.v1",
            "children": len(loaded), "paired_cells": 48, "cells": cells,
            "claim": "paired allocator metrics only; no allocator timing claim"}


def reduce_data(binding: Mapping[str, Any]) -> dict[str, Any]:
    native = reduce_native(binding)
    allocator = reduce_allocator(binding)
    allocation_cells = allocator["cells"]
    target_rows = []
    for case in FIXTURES:
        values = [cell for cell in allocation_cells if cell["case"] == case
                  and cell["edits"] == 32 and cell["retention"] == "default"]
        base = statistics.median(float(cell["baseline"]["allocated_bytes"]) for cell in values)
        cand = statistics.median(float(cell["candidate"]["allocated_bytes"]) for cell in values)
        target_rows.append({"case": case, "baseline_median_allocated_bytes": base,
                            "candidate_median_allocated_bytes": cand,
                            "candidate_over_baseline": cand / base if base else float("inf"),
                            "saves_at_least_5_percent": cand <= base * 0.95})
    gates = {
        "all_8_native_bootstrap_lower_bounds": all(cell["latency_gate"] for cell in native["cells"]),
        "candidate_saves_at_least_5_percent_in_one_32_edit_default_fixture": any(
            row["saves_at_least_5_percent"] for row in target_rows),
        "every_allocator_pair_gross_requested_within_1_percent": all(
            not cell["allocation"]["flag"] for cell in allocation_cells),
        "every_allocator_pair_peak_live_within_5_percent": all(
            not cell["peak_live"]["flag"] for cell in allocation_cells),
        "all_explicit_capacity_fields_equal": all(
            cell["exact"]["all_explicit_capacity_fields_equal"]
            for cell in allocation_cells
        ) and all(
            round_["exact"]["all_explicit_capacity_fields_equal"]
            for cell in native["cells"] for round_ in cell["rounds"]
        ),
    }
    gates["adoption"] = all(gates.values())
    return {
        "schema": "litchi.performance.0943.doc-comparison.reduction.v1",
        "protocol": protocol_descriptor(),
        "manifest": binding["manifest_descriptor"],
        "native": native,
        "allocator": allocator,
        "admission_targets": target_rows,
        "gates": gates,
        "review_flags": {
            "native": [round_ for cell in native["cells"] for round_ in cell["rounds"]],
            "allocator": allocation_cells,
            "automatic_waiver": False,
        },
        "claim_scope": {
            "timing": True,
            "allocator_timing": False,
            "adoption": bool(gates["adoption"]),
        },
    }


def _fixture_for_synthetic(p: Mapping[str, Any], case: str) -> dict[str, Any]:
    for row in p["fixtures"]:
        if row["case"] == case:
            root = Path("/home/zhuhe/litchi-performance-0943/baseline-source-v1")
            return {**row, "path": row["relative_path"],
                    "path_abs": str((root / row["relative_path"]).resolve())}
    raise RunnerError(f"unknown synthetic fixture: {case}")


def _zero_stage(stage: Mapping[str, Any]) -> dict[str, Any]:
    return {**stage, "stage": 0, "requested_text": "", "before": None,
            "after": None, "after_release": None}


def synthetic_report(template: Mapping[str, Any], *, case: str, edits: int,
                     retention: str, lane: str, warmups: int, samples: int) -> dict[str, Any]:
    report = copy.deepcopy(dict(template))
    fixture = _fixture_for_synthetic(protocol(), case)
    report["case"] = case
    report["input"] = fixture["path"]
    report["source_sha256"] = fixture["sha256"]
    report["source_bytes"] = fixture["bytes"]
    report["edits"] = edits
    report["retention"] = {"zero": "zero", "default": "default_8mib", "release": "release_8mib"}[retention]
    report["retention_ceiling_bytes"] = 0 if retention == "zero" else 8 * 1024 * 1024
    report["warmups"] = warmups
    report["samples_requested"] = samples
    report["timing_claim"] = lane == "native"
    report["allocator_instrumented"] = lane == "allocator"
    report["elapsed_ns"] = [1000 + index for index in range(samples)] if lane == "native" else []
    template_samples = report.get("samples")
    require(isinstance(template_samples, list) and template_samples,
            "synthetic template has no sample")
    report["samples"] = [copy.deepcopy(template_samples[index % len(template_samples)])
                         for index in range(samples)]
    route_name = report["retention"]
    selected_route_stages: list[dict[str, Any]] | None = None
    for route in report["route_outputs"]:
        route["semantic_change_count"] = edits
        stages = route["stages"]
        for index, stage in enumerate(stages):
            if index >= edits:
                stages[index] = _zero_stage(stage)
        if route["retention"] == route_name:
            report["output_capacity"] = route["output_capacity"]
            report["final_retained_before_commit"] = route.get("final_retained_before_commit")
            selected_route_stages = copy.deepcopy(route["stages"])
    require(selected_route_stages is not None, "synthetic selected route is missing")
    if lane == "allocator":
        selected_capacity = int(report["output_capacity"])
        allocation = report["allocation"]
        allocation["retained_bytes"] = selected_capacity
    for sample_index, sample in enumerate(report["samples"]):
        sample["index"] = sample_index
        if lane == "native":
            sample["elapsed_ns"] = 1000 + sample_index
        sample["semantic_change_count"] = edits
        sample["stages"] = copy.deepcopy(selected_route_stages)
        stages = sample.get("stages", sample.get("stage_observations"))
        for index, stage in enumerate(stages):
            if index >= edits:
                stages[index] = _zero_stage(stage)
        if lane == "allocator":
            sample["allocation"]["retained_bytes"] = int(report["output_capacity"])
    report["correctness"]["semantic_change_count_matches_edit_count"] = True
    return report


def synthetic_child(root: Path, spec: Mapping[str, Any], report: Mapping[str, Any],
                    fixture: Mapping[str, Any], outcome: list[dict[str, Any]],
                    *, formal_binding: Mapping[str, Any] | None = None,
                    schedule_rows: Sequence[Mapping[str, Any]] = (),
                    synthetic_index: int = 0) -> None:
    child = root / "children" / safe_label(str(spec["label"]))
    child.mkdir(parents=True, exist_ok=False)
    report_path = child / "report.jsonl"
    stderr_path = child / "stderr.log"
    rss_path = child / "max-rss-kib.txt"
    terminal_path = child / "terminal.json"
    validation_path = child / "validation.json"
    write_exclusive(report_path, (json.dumps(report, sort_keys=True) + "\n").encode())
    write_exclusive(stderr_path, b"")
    write_exclusive(rss_path, b"0\n")
    source_root = Path(fixture["path_abs"]).resolve().parents[3]
    command = (
        "/usr/bin/time", "-f", "%M", "-o", str(rss_path.resolve()),
        "/usr/bin/taskset", "--cpu-list", str(NATIVE_CPU if spec["lane"] == "native" else ALLOCATOR_CPU),
        "/synthetic/doc_retention_probe", "--case", str(spec["case"]),
        "--input", str(fixture["path_abs"]), "--edits", str(spec["edits"]),
        "--retention", str(spec["retention"]), "--warmup", str(spec["warmups"]),
        "--samples", str(spec["samples"]),
    )
    if formal_binding is not None:
        command = child_command(spec, binary_for(formal_binding, str(spec["version"]), str(spec["lane"])),
                                fixture_for(formal_binding, str(spec["version"]), str(spec["case"])), rss_path)
    command_packet_view = packet_command_view(command, root)
    report_desc = packet_descriptor(report_path, root)
    stderr_desc = packet_descriptor(stderr_path, root)
    rss_desc = packet_descriptor(rss_path, root, allow_large=False)
    started_utc = (dt.datetime(2026, 10, 6, tzinfo=dt.timezone.utc)
                   + dt.timedelta(microseconds=synthetic_index * 2)).isoformat().replace("+00:00", "Z")
    finished_utc = (dt.datetime(2026, 10, 6, tzinfo=dt.timezone.utc)
                    + dt.timedelta(microseconds=synthetic_index * 2 + 1)).isoformat().replace("+00:00", "Z")
    binary_identity = (formal_binding["build_bindings"][str(spec["version"])]
                       [str(spec["lane"])] ["binary"] if formal_binding is not None else None)
    terminal = {
        "schema": "litchi.performance.0943.child-terminal.v1"
        if formal_binding is not None else "litchi.performance.0943.synthetic-terminal.v1",
        "status": "terminal", "label": spec["label"], "spec": dict(spec),
        "argv": list(command), "packet_command": command_packet_view,
        "capture_origin": str(root.resolve()), "cwd": str(source_root),
        "environment": dict(ENVIRONMENT), "exit_code": 0,
        "terminal_written_before_report_parse": True,
        "started_utc": started_utc, "ended_utc": finished_utc,
        "deadline_seconds": int(DEADLINE_SECONDS), "deadline_exceeded": False, "runner_error": None,
        "process_group_closed": True, "process_group_forced_cleanup": False,
        "report": report_desc, "stderr": stderr_desc, "max_rss_kib": rss_desc,
        "source_revision": SOURCE_REVISION,
        "binary": binary_identity,
    }
    write_json(terminal_path, terminal)
    try:
        parsed = read_report(report_path)
        validator = (outcome[0]["validator"] if isinstance(outcome[0], Mapping)
                     else outcome[0])
        summary = validator.validate_report(
            parsed, dict(spec), fixture, str(spec["lane"]),
            expected_warmups=int(spec["warmups"]), expected_samples=int(spec["samples"]),
        )
        validation = {"schema": "litchi.performance.0943.child-validation.v1"
                      if formal_binding is not None else "litchi.performance.0943.synthetic-validation.v1",
                      "status": "passed", "report_summary": dict(summary),
                      "timing_claim": False, "report": report_desc, "rss_kib": 0}
        status = "passed"
    except Exception as error:
        validation = {"schema": "litchi.performance.0943.child-validation.v1"
                      if formal_binding is not None else "litchi.performance.0943.synthetic-validation.v1",
                      "status": "failed", "error": f"{type(error).__name__}: {error}"}
        status = "failed"
    write_json(validation_path, validation)
    if formal_binding is not None:
        receipt_value = {
            "schema": "litchi.performance.0943.child-receipt.v1", "status": status,
            "spec": dict(spec), "schedule_sha256": schedule_sha256(schedule_rows),
            "protocol": formal_binding["protocol_descriptor"],
            "manifest": formal_binding["manifest_descriptor"],
            "source": {"revision": SOURCE_REVISION, "root": str(source_root)},
            "binary": binary_identity, "command": list(command),
            "packet_command": command_packet_view, "capture_origin": str(root.resolve()),
            "cwd": str(source_root), "environment": dict(ENVIRONMENT),
            "started_utc": started_utc, "finished_utc": finished_utc, "exit_code": 0,
            "terminal": packet_descriptor(terminal_path, root, allow_large=False),
            "validation": packet_descriptor(validation_path, root, allow_large=False),
            "report": report_desc, "stderr": stderr_desc, "max_rss_kib": rss_desc,
            "rss_kib": 0,
        }
    else:
        receipt_value = {
        "schema": "litchi.performance.0943.synthetic-receipt.v1", "status": status,
        "spec": dict(spec), "terminal": packet_descriptor(terminal_path, root, allow_large=False),
        "validation": packet_descriptor(validation_path, root, allow_large=False),
        "report": packet_descriptor(report_path, root),
        "stderr": packet_descriptor(stderr_path, root),
        "max_rss_kib": packet_descriptor(rss_path, root, allow_large=False),
        "command": list(command), "packet_command": command_packet_view,
        "capture_origin": str(root.resolve()), "cwd": str(source_root),
        "timing_claim": False,
        }
    write_json(receipt_path(root, spec) if formal_binding is not None else child / "receipt.json",
               receipt_value)
    if status != "passed":
        raise RunnerError(f"synthetic positive failed: {spec['label']}")


def synthetic_packet_replay(root: Path, spec: Mapping[str, Any]) -> dict[str, Any]:
    """Replay one real-shaped synthetic packet after relocating its root."""
    source = root / "children" / safe_label(str(spec["label"]))
    moved_root = root / "relocation-control"
    moved = moved_root / "children" / safe_label(str(spec["label"]))
    require(not moved_root.exists(), "synthetic relocation root already exists")
    moved.parent.mkdir(parents=True, exist_ok=False)
    shutil.copytree(source, moved)
    receipt = json_load(moved / "receipt.json")
    require(receipt.get("status") == "passed" and receipt.get("spec") == dict(spec),
            "synthetic moved receipt identity changed")
    for field in ("terminal", "validation", "report", "stderr", "max_rss_kib"):
        _packet_descriptor_from_value(receipt[field], f"moved synthetic {field}", moved_root)
    origin = receipt.get("capture_origin")
    command = receipt.get("command")
    packet_command = receipt.get("packet_command")
    require(isinstance(origin, str) and Path(origin).is_absolute()
            and isinstance(command, list) and isinstance(packet_command, list),
            "synthetic moved command custody missing")
    require(packet_command == packet_command_view(command, Path(origin)),
            "synthetic moved command normalization changed")
    expected_rss = f"<packet>/children/{safe_label(str(spec['label']))}/max-rss-kib.txt"
    require(packet_command.count(expected_rss) == 1,
            "synthetic moved RSS command path changed")
    terminal_path, _ = _packet_descriptor_from_value(
        receipt["terminal"], "moved synthetic terminal", moved_root)
    terminal = json_load(terminal_path)
    require(terminal.get("argv") == command
            and terminal.get("packet_command") == packet_command
            and terminal.get("capture_origin") == origin
            and terminal.get("cwd") == receipt.get("cwd"),
            "synthetic moved terminal command/cwd changed")
    for field in ("report", "stderr", "max_rss_kib"):
        _, actual = _packet_descriptor_from_value(
            terminal[field], f"moved synthetic terminal {field}", moved_root)
        require(actual == receipt[field], f"synthetic moved {field} descriptor changed")
    return {"status": "passed", "source_root": str(root.resolve()),
            "moved_root": str(moved_root.resolve()), "label": spec["label"]}


def run_preflight(root: Path) -> dict[str, Any]:
    p = protocol()
    require(not root.exists(), f"refusing to replace preflight root: {root}")
    root.mkdir(parents=True, exist_ok=False)
    native_template = json_load(TEMPLATE_NATIVE)
    allocator_template = json_load(TEMPLATE_ALLOCATOR)
    validator = _load_validator(VALIDATOR_PATH)
    context = [{"validator": validator}]
    positive: list[str] = []
    native_rows = []
    for case in FIXTURES:
        for edits in EDITS:
            for version in VERSIONS:
                native_rows.append({
                    "lane": "native", "case": case, "edits": edits, "retention": "default",
                    "version": version, "warmups": 5, "samples": 50,
                    "label": f"native-{version}-{case}-e{edits:02d}",
                })
    allocator_rows = []
    for case in FIXTURES:
        for edits in EDITS:
            for route in ROUTES:
                for version in VERSIONS:
                    allocator_rows.append({
                        "lane": "allocator", "case": case, "edits": edits, "retention": route,
                        "version": version, "warmups": 0, "samples": 1,
                        "label": f"allocator-{version}-{case}-e{edits:02d}-{route}",
                    })
    for spec in native_rows + allocator_rows:
        template = native_template if spec["lane"] == "native" else allocator_template
        report = synthetic_report(template, case=spec["case"], edits=spec["edits"],
                                  retention=spec["retention"], lane=spec["lane"],
                                  warmups=spec["warmups"], samples=spec["samples"])
        synthetic_child(root, spec, report, _fixture_for_synthetic(p, spec["case"]), context)
        positive.append(spec["label"])
    relocation = synthetic_packet_replay(root, native_rows[0])
    controls: list[dict[str, Any]] = []
    base_spec = native_rows[0]
    base_report = synthetic_report(native_template, case="docnohf", edits=1, retention="default",
                                   lane="native", warmups=5, samples=50)
    negative_cases: list[tuple[str, Any]] = []
    wrong_fixture = copy.deepcopy(base_report); wrong_fixture["source_sha256"] = "0" * 64
    negative_cases.append(("wrong-fixture-digest", wrong_fixture))
    wrong_edits = copy.deepcopy(base_report); wrong_edits["edits"] = 2
    negative_cases.append(("wrong-edit-count", wrong_edits))
    wrong_route = copy.deepcopy(base_report); wrong_route["retention"] = "zero"
    negative_cases.append(("wrong-retention-route", wrong_route))
    wrong_output = copy.deepcopy(base_report); wrong_output["output_sha256"] = "0" * 64
    negative_cases.append(("wrong-output-digest", wrong_output))
    wrong_warmups = copy.deepcopy(base_report); wrong_warmups["warmups"] = 4
    negative_cases.append(("wrong-warmup-count", wrong_warmups))
    for name, candidate in negative_cases:
        try:
            validator.validate_report(candidate, dict(base_spec), _fixture_for_synthetic(p, "docnohf"),
                                      "native", expected_warmups=5, expected_samples=50)
        except Exception as error:
            controls.append({"name": name, "status": "rejected", "error": f"{type(error).__name__}: {error}"})
        else:
            raise RunnerError(f"negative control accepted: {name}")
    for name, data in (
        ("duplicate-json-key", b'{"x":1,"x":2}\n'),
        ("malformed-json", b'{not-json}\n'),
    ):
        try:
            strict_json_load_bytes(data, name)
        except Exception as error:
            controls.append({"name": name, "status": "rejected", "error": f"{type(error).__name__}: {error}"})
        else:
            raise RunnerError(f"negative control accepted: {name}")
    for name, data in (("wrong-binary-descriptor", {"path": "/no/such/binary", "bytes": 1, "sha256": "0" * 64}),
                       ("wrong-source-descriptor", {"path": "/no/such/source", "bytes": 1, "sha256": "0" * 64})):
        try:
            _descriptor_from_value(data, name)
        except Exception as error:
            controls.append({"name": name, "status": "rejected", "error": f"{type(error).__name__}: {error}"})
        else:
            raise RunnerError(f"negative control accepted: {name}")
    bad_rss = root / "controls-bad-rss.txt"
    write_exclusive(bad_rss, b"4096 KiB\n")
    try:
        parse_rss(bad_rss)
    except Exception as error:
        controls.append({"name": "malformed-rss-sidecar", "status": "rejected",
                         "error": f"{type(error).__name__}: {error}"})
    else:
        raise RunnerError("negative control accepted: malformed-rss-sidecar")
    summary = {
        "schema": "litchi.performance.0943.synthetic-preflight.v1",
        "status": "passed",
        "claim_scope": p["preflight"]["claim_scope"],
        "protocol": protocol_descriptor(),
        "runner": script_descriptor(RUNNER_PATH),
        "positive_children": len(positive),
        "expected_positive_children": 64,
        "native_positive_children": len(native_rows),
        "allocator_positive_children": len(allocator_rows),
        "positive_labels": positive,
        "negative_controls": controls,
        "terminal_written_before_report_parse": True,
        "packet_relocation": relocation,
        "templates": {
            "native": descriptor(TEMPLATE_NATIVE, allow_large=False),
            "allocator": descriptor(TEMPLATE_ALLOCATOR, allow_large=False),
        },
        "validator": script_descriptor(VALIDATOR_PATH),
        "artifact_root": str(root.resolve()),
    }
    write_json(root / "preflight-summary.json", summary)
    return summary


def run_synthetic_full(root: Path) -> dict[str, Any]:
    """Exercise the full shared schedule through child load and reduction."""
    p = protocol()
    require(not root.exists(), f"refusing to replace synthetic root: {root}")
    root.mkdir(parents=True, exist_ok=False)
    manifest_path = root / "synthetic-manifest.json"
    write_json(manifest_path, {"schema": "litchi.performance.0943.synthetic-manifest.v1",
                               "status": "synthetic_only", "source_revision": SOURCE_REVISION})
    source_root = Path("/home/zhuhe/litchi-performance-0943/baseline-source-v1").resolve()
    native_binary = Path(p["baseline_builds"]["native"]["binary"]["path"]).resolve()
    allocator_binary = Path(p["baseline_builds"]["allocator"]["binary"]["path"]).resolve()
    source_bindings = {version: {"source_root": source_root, "source_revision": SOURCE_REVISION}
                       for version in VERSIONS}
    build_bindings: dict[str, dict[str, Any]] = {}
    for version in VERSIONS:
        build_bindings[version] = {
            "native": {"binary": descriptor(native_binary), "binary_path": native_binary},
            "allocator": {"binary": descriptor(allocator_binary), "binary_path": allocator_binary},
        }
    validator = _load_validator(VALIDATOR_PATH)
    binding: dict[str, Any] = {
        "external_root": root, "protocol": p,
        "protocol_descriptor": protocol_descriptor(),
        "manifest_descriptor": descriptor(manifest_path, allow_large=False),
        "source_bindings": source_bindings, "build_bindings": build_bindings,
        "validator_module": validator,
    }
    rows = all_schedule()
    for index, spec in enumerate(rows):
        template = (json_load(TEMPLATE_NATIVE) if spec["lane"] == "native"
                    else json_load(TEMPLATE_ALLOCATOR))
        report = synthetic_report(template, case=str(spec["case"]), edits=int(spec["edits"]),
                                  retention=str(spec["retention"]), lane=str(spec["lane"]),
                                  warmups=int(spec["warmups"]), samples=int(spec["samples"]))
        synthetic_child(root, spec, report,
                        _fixture_for_synthetic(p, str(spec["case"])), [validator],
                        formal_binding=binding, schedule_rows=rows, synthetic_index=index)
    reduction = reduce_data(binding)
    reduction["claim_scope"] = {"timing": False, "allocation": False,
                                "adoption": False, "synthetic_only": True}
    reduction_path = root / "synthetic-reduction.json"
    write_json(reduction_path, reduction, allow_large=True)
    summary = {
        "schema": "litchi.performance.0943.synthetic-full.v1", "status": "passed",
        "runner": descriptor(RUNNER_PATH), "protocol": protocol_descriptor(),
        "claim_scope": "synthetic custody/load/reduction only; no workload or adoption claim",
        "children": len(rows), "native_children": 192, "allocator_children": 96,
        "schedule_sha256": schedule_sha256(rows), "reduction": descriptor(reduction_path),
        "artifact_root": str(root.resolve()), "terminal_written_before_report_parse": True,
    }
    write_json(root / "synthetic-full-summary.json", summary)
    return summary


def check_schedule() -> dict[str, Any]:
    p = protocol()
    native = native_schedule(); allocator = allocator_schedule()
    combined = all_schedule()
    require(len(combined) == 288 and combined == native + allocator,
            "combined schedule is not native followed by allocator")
    return {
        "schema": "litchi.performance.0943.schedule.v1",
        "status": "prospective_only_no_workload",
        "protocol": protocol_descriptor(),
        "native": {"children": len(native), "paired_cells": 96,
                    "schedule_sha256": schedule_sha256(native)},
        "allocator": {"children": len(allocator), "paired_cells": 48,
                       "schedule_sha256": schedule_sha256(allocator)},
        "combined": {"children": len(combined), "schedule_sha256": schedule_sha256(combined)},
        "total_children": len(native) + len(allocator),
        "source_revision": p["baseline_source_revision"],
        "candidate_source_revision": None,
        "timing_claim": False,
    }


def capture(manifest_path: Path, lane: str) -> dict[str, Any]:
    require(os.environ.get(APPROVAL_ENV) == APPROVAL_VALUE,
            f"set {APPROVAL_ENV}=1 only for an authorized formal capture")
    require(lane == "all", "formal capture must use the shared --lane all schedule")
    binding = validate_manifest(manifest_path)
    binding = dict(binding)
    binding["manifest_descriptor"] = descriptor(manifest_path, allow_large=False)
    binding["protocol_descriptor"] = protocol_descriptor()
    binding["validator_module"] = _load_validator(VALIDATOR_PATH)
    rows = schedule(lane)
    external = binding["external_root"]
    require(not external.exists(), f"formal root already exists: {external}")
    external.mkdir(parents=True, exist_ok=False)
    write_json(external / "capture-start.json", {
        "schema": "litchi.performance.0943.capture-start.v1",
        "started_utc": now(), "lane": lane, "schedule": rows,
        "schedule_sha256": schedule_sha256(rows),
        "manifest": binding["manifest_descriptor"], "protocol": binding["protocol_descriptor"],
        "source_revision": SOURCE_REVISION, "candidate_bound": True,
    })
    outcomes = []
    for spec in rows:
        result = run_one(binding, spec, rows)
        outcomes.append({"label": spec["label"], "status": result["validation"]["status"]})
    summary = {
        "schema": "litchi.performance.0943.capture-state.v1",
        "status": "completed" if all(item["status"] == "passed" for item in outcomes)
        else "completed_with_failures",
        "lane": lane, "children": len(outcomes),
        "passed": sum(item["status"] == "passed" for item in outcomes),
        "failed": sum(item["status"] != "passed" for item in outcomes),
        "schedule_sha256": schedule_sha256(rows), "results": outcomes,
        "finished_utc": now(), "manifest": binding["manifest_descriptor"],
    }
    write_json(external / "capture-state.json", summary)
    return summary


def qualify(manifest_path: Path, root: Path) -> dict[str, Any]:
    require(os.environ.get(QUALIFICATION_APPROVAL_ENV) == APPROVAL_VALUE,
            f"set {QUALIFICATION_APPROVAL_ENV}=1 only for an authorized actual qualification")
    binding = validate_manifest(manifest_path, authorization_kind="qualification")
    binding = dict(binding)
    binding["manifest_descriptor"] = descriptor(manifest_path, allow_large=False)
    binding["protocol_descriptor"] = protocol_descriptor()
    binding["validator_module"] = _load_validator(VALIDATOR_PATH)
    rows = qualification_schedule()
    require(binding["manifest"].get("schedule_sha256", {}).get("qualification")
            == schedule_sha256(rows), "manifest qualification schedule digest changed")
    external = root.resolve()
    binding["external_root"] = external
    require(not external.exists(), f"qualification root already exists: {external}")
    external.mkdir(parents=True, exist_ok=False)
    write_json(external / "qualification-start.json", {
        "schema": "litchi.performance.0943.qualification-start.v1",
        "started_utc": now(), "schedule": rows, "schedule_sha256": schedule_sha256(rows),
        "manifest": binding["manifest_descriptor"], "protocol": binding["protocol_descriptor"],
        "claim_scope": binding["protocol"]["qualification"]["claim_scope"],
    })
    outcomes = []
    for spec in rows:
        result = run_one(binding, spec, rows)
        outcomes.append({"label": spec["label"], "status": result["validation"]["status"]})
    audited = [load_child_receipt(binding, spec, rows) for spec in rows]
    pairs = {}
    for item in audited:
        spec = item["spec"]
        key = (spec["lane"], spec["case"], spec["edits"], spec["retention"])
        pairs.setdefault(key, {})[spec["version"]] = item["report"]
    for key, pair in pairs.items():
        require(set(pair) == set(VERSIONS), f"incomplete qualification pair: {key}")
        compare_reports(pair["baseline"], pair["candidate"], f"qualification {key}")
    summary = {
        "schema": "litchi.performance.0943.qualification-state.v1", "status": "passed",
        "claim_scope": binding["protocol"]["qualification"]["claim_scope"],
        "children": len(outcomes), "audited_children": len(audited), "exact_pairs": len(pairs),
        "schedule_sha256": schedule_sha256(rows), "results": outcomes, "finished_utc": now(),
        "manifest": binding["manifest_descriptor"], "formal_capture_started": False,
    }
    write_json(external / "qualification-state.json", summary)
    return summary


def reduce_capture(manifest_path: Path, lane: str, output: Path | None) -> dict[str, Any]:
    binding = validate_manifest(manifest_path)
    binding = dict(binding)
    binding["manifest_descriptor"] = descriptor(manifest_path, allow_large=False)
    binding["protocol_descriptor"] = protocol_descriptor()
    binding["validator_module"] = _load_validator(VALIDATOR_PATH)
    result = reduce_data(binding) if lane == "all" else (
        reduce_native(binding) if lane == "native" else reduce_allocator(binding))
    if output is not None:
        write_json(output, result)
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    actions = parser.add_mutually_exclusive_group(required=True)
    actions.add_argument("--check-schedule", action="store_true")
    actions.add_argument("--preflight", action="store_true")
    actions.add_argument("--synthetic-full", action="store_true")
    actions.add_argument("--qualify", action="store_true")
    actions.add_argument("--capture", action="store_true")
    actions.add_argument("--reduce", action="store_true")
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--lane", choices=["native", "allocator", "all"], default="native")
    parser.add_argument("--preflight-root", type=Path, default=DEFAULT_PREFLIGHT_ROOT)
    parser.add_argument("--synthetic-root", type=Path, default=PERF_ROOT / "synthetic-full-v1")
    parser.add_argument("--qualification-root", type=Path, default=PERF_ROOT / "qualification-v1")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    try:
        if args.check_schedule:
            result = check_schedule()
        elif args.preflight:
            result = run_preflight(args.preflight_root.resolve())
        elif args.synthetic_full:
            result = run_synthetic_full(args.synthetic_root.resolve())
        elif args.qualify:
            require(args.manifest is not None, "qualification requires --manifest")
            result = qualify(args.manifest.resolve(), args.qualification_root)
        elif args.capture:
            require(args.manifest is not None, "capture requires --manifest")
            result = capture(args.manifest.resolve(), args.lane)
        else:
            require(args.manifest is not None, "reduce requires --manifest")
            result = reduce_capture(args.manifest.resolve(), args.lane, args.output.resolve() if args.output else None)
        print(json.dumps(result, sort_keys=True))
        return 0
    except Exception as error:
        print(json.dumps({"schema": "litchi.performance.0943.runner-error.v1",
                          "status": "failed", "error": f"{type(error).__name__}: {error}"},
                         sort_keys=True))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
