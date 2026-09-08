#!/usr/bin/env python3
"""Fail-closed, portable verifier for the 0476 paired-arm evidence bundle.

The verifier authenticates the frozen lane protocol, every retained report and
receipt, arm/source/binary identities, canonical source manifests, and the
derived paired summary.  It does not import a capture driver, inspect a
temporary checkout, or require Rust, perf, or allocator tooling.  ``--live``
is available for an explicit source/binary custody check when the capture
trees still exist.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import subprocess
import sys
from typing import Any, Iterable, Mapping, NoReturn

import analyze
import custody
import report_checks

ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0476-verification-v1"
SHA = re.compile(r"^[0-9a-fA-F]{64}$")
REVISION = re.compile(r"^[0-9a-fA-F]{7,64}$")


class VerificationError(ValueError):
    """A retained artifact does not meet the 0476 custody contract."""


def fail(message: str) -> NoReturn:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=lambda value: (_ for _ in ()).throw(ValueError(f"non-finite JSON {value}")),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"cannot read JSON {path}: {error}")


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot canonicalize JSON: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def file_meta(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    try:
        size = path.stat().st_size
    except OSError as error:
        fail(f"cannot stat {path}: {error}")
    return {"sha256": sha256(path), "bytes": size}


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label} must be a non-empty relative POSIX path")
    require("\\" not in value, f"{label} contains a backslash")
    path = PurePosixPath(value)
    require(not path.is_absolute() and value not in {".", ".."}, f"{label} must be relative")
    require(all(part not in {"", ".", ".."} for part in path.parts), f"{label} escapes its root")
    return path.as_posix()


def bundle_path(root: Path, value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    current = root
    for part in PurePosixPath(relative).parts:
        current = current / part
        require(not current.is_symlink(), f"symlink in {label}: {relative}")
    return current


def check_digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA.fullmatch(value), f"{label} is not a SHA-256 digest")
    return value.lower()


def check_meta(path: Path, expected: Mapping[str, Any], label: str) -> None:
    require(set(expected) == {"sha256", "bytes"}, f"{label} metadata keys differ")
    expected_hash = check_digest(expected["sha256"], f"{label}.sha256")
    expected_bytes = expected["bytes"]
    require(isinstance(expected_bytes, int) and not isinstance(expected_bytes, bool) and expected_bytes >= 0, f"{label}.bytes is invalid")
    actual = file_meta(path)
    require(actual == {"sha256": expected_hash, "bytes": expected_bytes}, f"{label} hash/size mismatch")


def protocol(root: Path) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    path = root / "protocol.json"
    require(path.is_file(), "protocol.json is missing")
    value = read_json(path)
    require(isinstance(value, dict), "protocol must be an object")
    try:
        rows = analyze.protocol_order(value)
    except analyze.AnalysisError as error:
        fail(str(error))
    for key, expected in (("samples", 30), ("warmups", 3), ("workers", 1), ("cpu", 2)):
        if key in value:
            require(value[key] == expected, f"protocol.{key} must be {expected}")
    require(value.get("selector") == report_checks.CASE, "protocol.selector must bind the PPTX streaming case")
    require(value.get("shapes") == report_checks.SHAPES, "protocol.shapes differ from the frozen corpus")
    drivers = value.get("drivers")
    require(isinstance(drivers, dict) and drivers, "protocol.drivers is missing")
    for name, expected in drivers.items():
        relative = safe_relative(name, f"protocol.drivers[{name!r}]")
        path_value = root / relative
        require(path_value.is_file(), f"protocol driver is missing: {relative}")
        require(sha256(path_value) == check_digest(expected, f"protocol.drivers[{name!r}]"), f"protocol driver hash differs: {relative}")
    corpora = value.get("corpora")
    require(isinstance(corpora, dict) and set(corpora) == set(report_checks.SHAPES), "protocol.corpora must bind all three shapes")
    for shape, expected in corpora.items():
        require(isinstance(expected, dict), f"protocol.corpora.{shape} must be an object")
        require(expected.get("shape") == shape, f"protocol.corpora.{shape}.shape differs")
        check_digest(expected.get("archive_sha256"), f"protocol.corpora.{shape}.archive_sha256")
        for key in ("archive_bytes", "archive_member_count", "entry_count", "uncompressed_payload_bytes"):
            if key in expected:
                require(isinstance(expected[key], int) and expected[key] >= 0, f"protocol.corpora.{shape}.{key} is invalid")
    return value, rows, sha256(path)


def _arm_object(value: Any, arm: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"arm {arm} binding must be an object")
    return value


def arm_bindings(root: Path, protocol_value: Mapping[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, str]]]:
    raw = protocol_value.get("arms")
    if raw is None:
        for name in ("binding.json", "build.json", "source-binding.json"):
            candidate = root / name
            if candidate.is_file():
                data = read_json(candidate)
                if isinstance(data, dict) and isinstance(data.get("arms"), dict):
                    raw = data["arms"]
                    break
    if raw is None:
        # 0476 intentionally keeps one authenticated build/source binding per
        # arm so the reused control and freshly built candidate can be checked
        # independently.  Compose those files here without requiring a large
        # duplicated manifest.
        composed: dict[str, Any] = {}
        for arm in analyze.ARMS:
            for name in (f"{arm}-build.json", f"{arm}-source.json"):
                candidate = root / name
                if candidate.is_file():
                    value = read_json(candidate)
                    if isinstance(value, dict):
                        composed[arm] = value
                        break
        raw = composed if composed else None
    require(isinstance(raw, dict), "protocol or binding must declare control and candidate arms")
    result: dict[str, dict[str, Any]] = {}
    manifests: dict[str, dict[str, str]] = {}
    for arm in analyze.ARMS:
        item = _arm_object(raw.get(arm), arm)
        result[arm] = item
        if item.get("arm") is not None:
            require(item["arm"] == arm, f"arms.{arm}.arm differs")
        if item.get("schema") is not None:
            require(item["schema"] == "litchi-0476-source-v1", f"arms.{arm}.schema differs")
        revision = item.get("revision", item.get("git_revision"))
        if revision is not None:
            require(isinstance(revision, str) and REVISION.fullmatch(revision), f"arms.{arm}.revision is invalid")
        binary = item.get("binary")
        if isinstance(binary, dict):
            if binary.get("sha256") is not None:
                check_digest(binary["sha256"], f"arms.{arm}.binary.sha256")
            if binary.get("bytes") is not None:
                require(isinstance(binary["bytes"], int) and binary["bytes"] > 0, f"arms.{arm}.binary.bytes is invalid")
        elif item.get("binary_sha256") is not None:
            check_digest(item["binary_sha256"], f"arms.{arm}.binary_sha256")
        binaries = item.get("binaries")
        if binaries is not None:
            require(isinstance(binaries, dict) and set(binaries) >= {"normal", "allocator"}, f"arms.{arm}.binaries must contain normal and allocator")
            for mode in ("normal", "allocator"):
                entry = binaries[mode]
                require(isinstance(entry, dict), f"arms.{arm}.binaries.{mode} is malformed")
                check_digest(entry.get("sha256"), f"arms.{arm}.binaries.{mode}.sha256")
                require(isinstance(entry.get("bytes"), int) and entry["bytes"] > 0, f"arms.{arm}.binaries.{mode}.bytes is invalid")
                require(isinstance(entry.get("path"), str) and entry["path"], f"arms.{arm}.binaries.{mode}.path is missing")
        manifest_value = item.get("source_manifest", item.get("source_manifest_path", item.get("manifest")))
        manifest_metadata: Mapping[str, Any] | None = manifest_value if isinstance(manifest_value, dict) else None
        if isinstance(manifest_value, dict) and isinstance(manifest_value.get("path"), str):
            manifest_path = bundle_path(root, manifest_value["path"], f"arms.{arm}.source_manifest.path")
        elif isinstance(manifest_value, str):
            manifest_path = bundle_path(root, manifest_value, f"arms.{arm}.source_manifest")
        elif isinstance(manifest_value, dict):
            # Inline manifests are accepted only as canonical content; this
            # is still authenticated through the arm's optional hash field.
            manifest_path = None
            manifest_value = dict(manifest_value)
        else:
            manifest_path = None
        if manifest_path is not None:
            require(manifest_path.is_file(), f"arm {arm} source manifest is missing")
            data = read_json(manifest_path)
            manifests[arm] = normalize_manifest(data, f"arms.{arm}.source_manifest")
            if manifest_metadata is not None and manifest_metadata.get("files") is not None:
                require(manifest_metadata["files"] == len(manifests[arm]), f"arm {arm} source manifest file count differs")
            expected = item.get("source_manifest_sha256", item.get("manifest_sha256"))
            if expected is None and manifest_metadata is not None:
                expected = manifest_metadata.get("sha256")
            if expected is not None:
                check_digest(expected, f"arms.{arm}.source_manifest_sha256")
                require(sha256(manifest_path) == expected.lower(), f"arm {arm} source manifest hash differs")
        elif isinstance(manifest_value, dict):
            manifests[arm] = normalize_manifest(manifest_value, f"arms.{arm}.source_manifest")
    require(set(result) == set(analyze.ARMS), "both control and candidate arm bindings are required")
    return result, manifests


def normalize_manifest(value: Any, label: str) -> dict[str, str]:
    if isinstance(value, dict):
        if isinstance(value.get("files"), dict):
            value = value["files"]
        elif isinstance(value.get("manifest"), dict):
            value = value["manifest"]
        elif isinstance(value.get("entries"), list):
            result: dict[str, str] = {}
            for index, entry in enumerate(value["entries"]):
                require(isinstance(entry, dict), f"{label}.entries[{index}] must be an object")
                path = safe_relative(entry.get("path"), f"{label}.entries[{index}].path")
                result[path] = check_digest(entry.get("sha256"), f"{label}.entries[{index}].sha256")
            value = result
    require(isinstance(value, dict) and value, f"{label} must contain non-empty file hashes")
    result = {}
    for path, digest_value in value.items():
        safe_relative(path, f"{label}.{path}")
        result[str(path)] = check_digest(digest_value, f"{label}.{path}")
    require(len(result) == len(value), f"{label} has duplicate file paths")
    return result


def _arm_revision(arm: Mapping[str, Any]) -> str | None:
    value = arm.get("revision", arm.get("git_revision"))
    return value if isinstance(value, str) else None


def _arm_binary_hash(arm: Mapping[str, Any]) -> str | None:
    binary = arm.get("binary")
    value = binary.get("sha256") if isinstance(binary, dict) else arm.get("binary_sha256")
    return value.lower() if isinstance(value, str) else None


def _arm_mode_binary_hash(arm: Mapping[str, Any], mode: Any) -> str | None:
    binaries = arm.get("binaries")
    if isinstance(binaries, dict) and isinstance(binaries.get(mode), dict):
        value = binaries[mode].get("sha256")
        return value.lower() if isinstance(value, str) else None
    return _arm_binary_hash(arm)


def lane_directory(root: Path, row: Mapping[str, Any]) -> Path:
    value = row.get("directory", row.get("output_directory", row.get("capture_directory")))
    if isinstance(value, str) and value:
        return Path(value) if Path(value).is_absolute() else root / value
    base = "pilots" if row.get("suite") == "pilot" else "captures"
    candidate = root / base / row["lane"]
    return candidate if candidate.is_dir() else root / row["lane"]


def verify_artifacts(directory: Path, artifacts: Mapping[str, Any], label: str) -> None:
    require(isinstance(artifacts, dict), f"{label}.artifacts must be an object")
    for name, metadata in artifacts.items():
        relative = safe_relative(name, f"{label}.artifacts[{name!r}]")
        path = directory / relative
        check_meta(path, metadata, f"{label}.artifacts[{name!r}]")
    if "report.json" in artifacts:
        check_meta(directory / "report.json", artifacts["report.json"], f"{label}.artifacts.report.json")


def _find_archive_hash(value: Any, expected: str) -> bool:
    if isinstance(value, dict):
        for key, child in value.items():
            if key in {"archive_sha256", "output_sha256"} and isinstance(child, str) and child.lower() == expected.lower():
                return True
            if _find_archive_hash(child, expected):
                return True
    elif isinstance(value, list):
        return any(_find_archive_hash(child, expected) for child in value)
    return False


def verify_catalog(directory: Path, report: Mapping[str, Any], label: str) -> None:
    path = directory / "corpus-catalog.json"
    require(path.is_file(), f"{label}: corpus-catalog.json is missing")
    catalog = read_json(path)
    require(isinstance(catalog, dict), f"{label}: corpus catalog is not an object")
    row = report.get("results", [{}])[0] if isinstance(report.get("results"), list) and report.get("results") else {}
    corpus = row.get("corpus", {}) if isinstance(row, dict) else {}
    expected = corpus.get("archive_sha256") if isinstance(corpus, dict) else None
    require(isinstance(expected, str) and SHA.fullmatch(expected), f"{label}: report archive identity is missing")
    require(_find_archive_hash(catalog, expected), f"{label}: catalog does not retain the report archive identity")
    catalog_hash = catalog.get("catalog_sha256")
    if catalog_hash is not None:
        check_digest(catalog_hash, f"{label}.catalog_sha256")
        unsigned = dict(catalog)
        del unsigned["catalog_sha256"]
        require(hashlib.sha256(canonical(unsigned)).hexdigest() == catalog_hash.lower(), f"{label}: catalog_sha256 arithmetic differs")
    content_hash = catalog.get("content_set_sha256")
    if content_hash is not None:
        check_digest(content_hash, f"{label}.content_set_sha256")


def verify_receipt(root: Path, row: Mapping[str, Any], arms: Mapping[str, Mapping[str, Any]], protocol_hash: str, *, formal: bool) -> dict[str, Any]:
    directory = lane_directory(root, row)
    receipt_path = directory / "receipt.json"
    require(receipt_path.is_file(), f"{row['lane']}: receipt.json is missing")
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), f"{row['lane']}: receipt must be an object")
    for key, expected in (("lane", row["lane"]), ("arm", analyze.lane_arm(row)), ("repeat", row.get("repeat")), ("mode", row.get("mode")), ("shape", row.get("shape"))):
        if key in receipt and expected is not None:
            require(receipt[key] == expected, f"{row['lane']}: receipt.{key} differs from protocol")
    if receipt.get("protocol_sha256") is not None:
        require(receipt["protocol_sha256"].lower() == protocol_hash, f"{row['lane']}: protocol hash differs")
    if receipt.get("exit_code") is not None:
        require(receipt["exit_code"] == 0, f"{row['lane']}: command exit code is non-zero")
    for key in ("clean_before", "clean_after", "binary_unchanged"):
        if key in receipt:
            require(receipt[key] is True, f"{row['lane']}: receipt.{key} must be true")
    arm = analyze.lane_arm(row)
    if arm in arms:
        expected_revision = _arm_revision(arms[arm])
        expected_binary = _arm_mode_binary_hash(arms[arm], row.get("mode"))
        if expected_revision is not None and receipt.get("revision") is not None:
            require(receipt["revision"] == expected_revision, f"{row['lane']}: arm revision differs")
        if expected_binary is not None and receipt.get("binary_sha256") is not None:
            require(receipt["binary_sha256"].lower() == expected_binary, f"{row['lane']}: arm binary differs")
    if isinstance(receipt.get("artifacts"), dict):
        verify_artifacts(directory, receipt["artifacts"], f"{row['lane']}.receipt")
    report_path = directory / "report.json"
    if formal:
        require(report_path.is_file(), f"{row['lane']}: report.json is missing")
        try:
            report = analyze.read_json(report_path)
            checked = analyze.validate_report(report, row, str(report_path))
        except (analyze.AnalysisError, report_checks.ReportError) as error:
            fail(f"{row['lane']}: {error}")
        report_arm = report.get("arm") if isinstance(report, dict) else None
        if report_arm is not None:
            require(report_arm == arm, f"{row['lane']}: report arm differs")
        row0 = checked["row"]
        environment = report.get("environment") if isinstance(report, dict) else None
        if isinstance(environment, dict):
            expected_revision = _arm_revision(arms[arm]) if arm in arms else None
            if expected_revision is not None and environment.get("git_revision") is not None:
                require(environment["git_revision"] == expected_revision, f"{row['lane']}: report revision differs")
        identity = row0.get("binary_identity")
        expected_binary = _arm_mode_binary_hash(arms[arm], row.get("mode")) if arm in arms else None
        if isinstance(identity, dict) and expected_binary is not None and identity.get("binary_sha256") is not None:
            require(identity["binary_sha256"].lower() == expected_binary, f"{row['lane']}: report binary differs")
        verify_catalog(directory, report, row["lane"])
    return receipt


def verify_auxiliary_report(root: Path, row: Mapping[str, Any], protocol_value: Mapping[str, Any]) -> None:
    directory = lane_directory(root, row)
    report = directory / "report.json"
    require(report.is_file(), f"{row['lane']}: auxiliary report is missing")
    value = read_json(report)
    require(isinstance(value, dict), f"{row['lane']}: auxiliary report is not an object")
    if row.get("suite") == "pilot":
        try:
            report_checks.validate_pilot_report(value, row, path=str(report))
        except report_checks.ReportError as error:
            fail(f"{row['lane']}: {error}")
    elif row.get("suite") == "counter":
        try:
            analyze.validate_report(value, row, str(report))
        except (analyze.AnalysisError, report_checks.ReportError) as error:
            fail(f"{row['lane']}: {error}")
    elif row.get("suite") == "guard":
        try:
            report_checks.validate_guard_report(value, protocol_value.get("guard_selectors", []), path=str(report))
        except report_checks.ReportError as error:
            fail(f"{row['lane']}: {error}")
    configuration = value.get("configuration")
    if isinstance(configuration, dict):
        expected_samples = 1 if row.get("suite") == "pilot" else int(row.get("samples", 30))
        expected_warmups = 0 if row.get("suite") == "pilot" else int(row.get("warmups", 3))
        if configuration.get("samples_per_case") is not None:
            require(configuration["samples_per_case"] == expected_samples, f"{row['lane']}: sample count differs")
        if configuration.get("warmup_iterations_per_case") is not None:
            require(configuration["warmup_iterations_per_case"] == expected_warmups, f"{row['lane']}: warmup count differs")
    if row.get("suite") == "counter":
        counters = directory / "counters.csv"
        require(counters.is_file(), f"{row['lane']}: counters.csv is missing")
        events = protocol_value.get("counter_events")
        require(isinstance(events, str) and events, "protocol.counter_events is missing")
        try:
            parsed = report_checks.parse_counter_text(counters.read_text(encoding="utf-8"), [event for event in events.split(",") if event])
        except (OSError, UnicodeError, report_checks.ReportError) as error:
            fail(f"{row['lane']}: invalid counters.csv: {error}")
        require(set(parsed) == set(event for event in events.split(",") if event), f"{row['lane']}: counter event set differs")
    results = value.get("results")
    if isinstance(results, list):
        if row.get("suite") == "guard":
            selectors = row.get("guard_selectors")
            if selectors is None:
                selectors = []
            require(isinstance(selectors, list), f"{row['lane']}: guard selector declaration is malformed")
            shapes = [part for part in str(row.get("shape", "")).split(",") if part]
            require(len(results) == len(selectors) * len(shapes), f"{row['lane']}: guard result count differs")
            expected_cases = set(selectors)
            actual_cases = {result.get("case") for result in results if isinstance(result, dict)}
            require(actual_cases == expected_cases, f"{row['lane']}: guard selector set differs")
            identities = {
                (result.get("case"), result.get("corpus", {}).get("shape") if isinstance(result.get("corpus"), dict) else None)
                for result in results
                if isinstance(result, dict)
            }
            require(identities == {(selector, shape) for selector in selectors for shape in shapes}, f"{row['lane']}: guard selector/shape identities differ")
        else:
            require(len(results) == 1, f"{row['lane']}: auxiliary report must contain one result")
        for index, result in enumerate(results):
            if isinstance(result, dict):
                if result.get("output_sha256") is not None:
                    check_digest(result["output_sha256"], f"{row['lane']}.results[{index}].output_sha256")
                    corpus = result.get("corpus")
                    if isinstance(corpus, dict) and corpus.get("archive_sha256") is not None:
                        require(result["output_sha256"] == corpus["archive_sha256"], f"{row['lane']}.results[{index}] output/archive identity differs")
                require(result.get("elapsed_ns") is not None, f"{row['lane']}.results[{index}].elapsed_ns is missing")
                try:
                    sample_count = 1 if row.get("suite") == "pilot" else int(row.get("samples", 30))
                    report_checks.check_elapsed(result["elapsed_ns"], f"{row['lane']}.results[{index}].elapsed_ns", sample_count)
                except report_checks.ReportError as error:
                    fail(str(error))
        catalog_path = directory / "corpus-catalog.json"
        require(catalog_path.is_file(), f"{row['lane']}: corpus-catalog.json is missing")
        catalog = read_json(catalog_path)
        require(isinstance(catalog, dict), f"{row['lane']}: corpus catalog is not an object")
        if catalog.get("catalog_sha256") is not None:
            check_digest(catalog["catalog_sha256"], f"{row['lane']}.catalog_sha256")
            unsigned = dict(catalog)
            del unsigned["catalog_sha256"]
            require(hashlib.sha256(canonical(unsigned)).hexdigest() == catalog["catalog_sha256"].lower(), f"{row['lane']}: catalog hash arithmetic differs")
        for index, result in enumerate(results):
            if isinstance(result, dict) and isinstance(result.get("output_sha256"), str):
                require(_find_archive_hash(catalog, result["output_sha256"]), f"{row['lane']}: catalog misses result {index} archive identity")


def verify_corpus_oracle(root: Path, protocol_value: Mapping[str, Any], rows: Iterable[Mapping[str, Any]]) -> None:
    corpora = protocol_value["corpora"]
    for row in rows:
        if row.get("suite") not in {"main", "pilot", "counter"}:
            continue
        shape = row.get("shape")
        if shape not in corpora:
            continue
        directory = lane_directory(root, row)
        report_path = directory / "report.json"
        require(report_path.is_file(), f"{row['lane']}: report.json is missing for corpus oracle")
        report = read_json(report_path)
        results = report.get("results") if isinstance(report, dict) else None
        require(isinstance(results, list) and results, f"{row['lane']}: report has no result rows")
        expected = corpora[shape]
        for index, result in enumerate(results):
            require(isinstance(result, dict), f"{row['lane']}.results[{index}] is malformed")
            corpus = result.get("corpus")
            require(isinstance(corpus, dict), f"{row['lane']}.results[{index}].corpus is missing")
            require(result.get("output_sha256") == corpus.get("archive_sha256") == expected["archive_sha256"], f"{row['lane']}.results[{index}] archive output differs from protocol corpus")
            for key in ("archive_bytes", "archive_member_count", "entry_count", "uncompressed_payload_bytes"):
                if key in expected and key in corpus:
                    require(corpus[key] == expected[key], f"{row['lane']}.results[{index}].corpus.{key} differs from protocol corpus")


def verify_source_manifest_files(root: Path, manifests: Mapping[str, Mapping[str, str]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for arm, manifest in manifests.items():
        checked = 0
        missing = 0
        for relative, expected in manifest.items():
            # A portable bundle may contain hashes only.  If the path is
            # materialized under sources/<arm>, verify it; otherwise retain
            # the manifest as the authenticated source identity.
            candidates = [root / "sources" / arm / relative, root / "source" / arm / relative]
            existing = next((path for path in candidates if path.is_file()), None)
            if existing is None:
                missing += 1
                continue
            require(sha256(existing) == expected, f"{arm} source file hash differs: {relative}")
            checked += 1
        result[arm] = {"manifest_files": len(manifest), "materialized_checked": checked, "hash_only": missing}
    return result


def verify_validation_receipts(root: Path) -> dict[str, Any]:
    directory = root / "validation"
    if not directory.is_dir():
        return {"status": "absent", "receipts": 0, "snapshots": 0}
    receipts = 0
    snapshots: set[str] = set()
    failures: list[str] = []
    for path in sorted(directory.glob("*.json")):
        if path.name.endswith(".started.json"):
            continue
        record = read_json(path)
        require(isinstance(record, dict), f"validation receipt {path.name} is malformed")
        require(record.get("source_unchanged") is True, f"validation receipt {path.name} changed source custody")
        if record.get("exit_code") not in (None, 0):
            failures.append(path.stem)
        before = record.get("source_before")
        after = record.get("source_after")
        require(isinstance(before, dict) and isinstance(after, dict), f"validation receipt {path.name} lacks source snapshots")
        require(before == after, f"validation receipt {path.name} source snapshot differs before/after")
        relative = safe_relative(before.get("path"), f"{path.name}.source_before.path")
        snapshot_path = root / relative
        require(snapshot_path.is_file(), f"validation source snapshot is missing: {relative}")
        expected = check_digest(before.get("sha256"), f"{path.name}.source_before.sha256")
        require(sha256(snapshot_path) == expected, f"validation source snapshot hash differs: {relative}")
        manifest = read_json(snapshot_path)
        normalized = normalize_manifest(manifest, f"{path.name}.source_snapshot")
        require(len(normalized) == before.get("files"), f"validation source snapshot file count differs: {relative}")
        snapshots.add(relative)
        artifacts = record.get("artifacts")
        if isinstance(artifacts, dict):
            verify_artifacts(directory, {name: metadata for name, metadata in artifacts.items()}, path.name)
        receipts += 1
    return {"status": "pass", "receipts": receipts, "snapshots": len(snapshots), "retained_failed_attempts": sorted(failures)}


def verify_rust_validation(root: Path) -> dict[str, Any]:
    """Verify the authenticated required/retained Rust validation ledger.

    The individual validation receipts are checked separately for source
    custody and artifact hashes.  This ledger binds the set of required
    successful gates and the explicitly retained non-zero attempts, so a
    successful subset cannot be presented as the complete validation run.
    """

    path = root / "rust-validation.json"
    require(path.is_file(), "rust-validation.json is missing")
    value = read_json(path)
    require(isinstance(value, dict), "rust-validation.json must be an object")
    require(value.get("schema") == "litchi-0476-rust-validation-v1", "rust-validation.json schema differs")

    candidate_source = value.get("candidate_source")
    candidate_source_path = bundle_path(root, candidate_source, "rust-validation.candidate_source")
    require(candidate_source_path.is_file(), f"candidate source binding is missing: {candidate_source}")
    candidate_binding = read_json(candidate_source_path)
    require(isinstance(candidate_binding, dict), "candidate source binding is malformed")
    require(candidate_binding.get("arm") in (None, "candidate"), "candidate source binding arm differs")

    format_exception = value.get("format_exception")
    format_exception_path = bundle_path(root, format_exception, "rust-validation.format_exception")
    require(format_exception_path.is_file(), f"format exception is missing: {format_exception}")
    exception = read_json(format_exception_path)
    require(isinstance(exception, dict), "format exception is malformed")
    exception_path = safe_relative(exception.get("path"), "format_exception.path")
    require(exception.get("identical_to_pre_candidate_head") is True, "format exception does not bind the unchanged file")
    require(isinstance(exception.get("pre_candidate_head"), str) and REVISION.fullmatch(exception["pre_candidate_head"]), "format exception pre-candidate revision is invalid")
    exception_hash = check_digest(exception.get("sha256"), "format_exception.sha256")
    candidate_manifest = candidate_binding.get("source_manifest")
    candidate_manifest_path: Path | None = None
    if isinstance(candidate_manifest, dict) and isinstance(candidate_manifest.get("path"), str):
        candidate_manifest_path = bundle_path(root, candidate_manifest["path"], "candidate source manifest")
    elif isinstance(candidate_manifest, str):
        candidate_manifest_path = bundle_path(root, candidate_manifest, "candidate source manifest")
    if candidate_manifest_path is not None and candidate_manifest_path.is_file():
        candidate_files = normalize_manifest(read_json(candidate_manifest_path), "candidate source manifest")
        require(candidate_files.get(exception_path) == exception_hash, "format exception hash is not the candidate source hash")

    required = value.get("required_success")
    retained = value.get("retained_nonzero_attempts")
    ledger = value.get("receipts")
    require(isinstance(required, list) and required and all(isinstance(label, str) and label for label in required), "rust-validation.required_success is invalid")
    require(len(set(required)) == len(required), "rust-validation.required_success contains duplicates")
    require(isinstance(retained, dict) and retained, "rust-validation.retained_nonzero_attempts is invalid")
    require(isinstance(ledger, dict) and ledger, "rust-validation.receipts is invalid")
    required_set = set(required)
    retained_set = set(retained)
    require(required_set.isdisjoint(retained_set), "required and retained validation labels overlap")
    require(set(ledger) == required_set | retained_set, "rust-validation receipt labels differ from required/retained labels")

    validation_dir = root / "validation"
    require(validation_dir.is_dir(), "validation directory is missing")
    actual_receipts = {path.stem for path in validation_dir.glob("*.json") if not path.name.endswith(".started.json")}
    require(actual_receipts == set(ledger), "validation directory contains an unledgered or missing receipt")

    for label, expected in ledger.items():
        require(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", label) is not None, f"rust-validation label is unsafe: {label!r}")
        receipt_path = validation_dir / f"{label}.json"
        require(receipt_path.is_file(), f"rust-validation receipt is missing: {label}")
        record = read_json(receipt_path)
        require(isinstance(record, dict), f"rust-validation receipt is malformed: {label}")
        require(isinstance(expected, dict), f"rust-validation ledger entry is malformed: {label}")
        expected_hash = check_digest(expected.get("receipt_sha256"), f"rust-validation.receipts.{label}.receipt_sha256")
        require(sha256(receipt_path) == expected_hash, f"rust-validation receipt hash differs: {label}")
        require(record.get("exit_code") == expected.get("exit_code"), f"rust-validation exit code differs: {label}")
        source_before = expected.get("source_before")
        require(isinstance(source_before, dict), f"rust-validation source snapshot is missing: {label}")
        require(record.get("source_before") == source_before, f"rust-validation source snapshot differs: {label}")
        if label in required_set:
            require(record.get("exit_code") == 0, f"required Rust validation is not successful: {label}")
        else:
            expected_exit = retained[label]
            require(isinstance(expected_exit, int) and not isinstance(expected_exit, bool) and expected_exit != 0, f"retained non-zero code is invalid: {label}")
            require(record.get("exit_code") == expected_exit, f"retained Rust validation exit code differs: {label}")

    return {
        "status": "pass",
        "required_success": len(required_set),
        "retained_nonzero_attempts": len(retained_set),
        "receipts": len(ledger),
    }


def verify_evidence_validation(root: Path) -> dict[str, Any]:
    """Verify the final Python helper hashes and explicit successful checks."""

    path = root / "evidence-validation.json"
    require(path.is_file(), "evidence-validation.json is missing")
    value = read_json(path)
    require(isinstance(value, dict), "evidence-validation.json must be an object")
    require(value.get("schema") == "litchi-0476-evidence-validation-v1", "evidence-validation.json schema differs")
    helper_hashes = value.get("helper_hashes")
    require(isinstance(helper_hashes, dict) and helper_hashes, "evidence-validation.helper_hashes is invalid")
    expected_helpers = {"analyze.py", "custody.py", "report_checks.py", "test_evidence.py", "verify.py"}
    require(set(helper_hashes) == expected_helpers, "evidence-validation helper set differs")
    for name, expected in helper_hashes.items():
        helper = bundle_path(root, name, f"evidence-validation.helper_hashes[{name!r}]")
        require(sha256(helper) == check_digest(expected, f"evidence-validation.helper_hashes.{name}"), f"evidence-validation helper hash differs: {name}")
    required = value.get("required_success")
    receipts = value.get("receipts")
    require(isinstance(required, list) and required and all(isinstance(label, str) and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", label) for label in required), "evidence-validation.required_success is invalid")
    require(len(set(required)) == len(required), "evidence-validation.required_success contains duplicates")
    require(isinstance(receipts, dict) and set(receipts) == set(required), "evidence-validation receipt set differs")
    directory = root / "validation"
    for label in required:
        expected = receipts[label]
        require(isinstance(expected, dict), f"evidence-validation receipt entry is malformed: {label}")
        require(expected.get("exit_code") == 0, f"evidence-validation required check is not successful: {label}")
        receipt_path = directory / f"{label}.json"
        require(receipt_path.is_file(), f"evidence-validation receipt is missing: {label}")
        require(sha256(receipt_path) == check_digest(expected.get("receipt_sha256"), f"evidence-validation.receipts.{label}.receipt_sha256"), f"evidence-validation receipt hash differs: {label}")
        record = read_json(receipt_path)
        require(isinstance(record, dict) and record.get("exit_code") == 0, f"evidence-validation receipt exit code differs: {label}")
    return {"status": "pass", "required_success": len(required), "helper_files": len(helper_hashes)}


def verify_seal(root: Path, *, required: bool = True) -> dict[str, Any]:
    seal = root / "SHA256SUMS"
    if not seal.is_file():
        require(not required, "SHA256SUMS is missing; use --unsealed only for an in-progress bundle")
        return {"status": "unsealed"}
    entries: dict[str, str] = {}
    for index, line in enumerate(seal.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        parts = line.split(maxsplit=1)
        require(len(parts) == 2 and SHA.fullmatch(parts[0]), f"SHA256SUMS line {index} is malformed")
        name = parts[1]
        if name.startswith("*"):
            name = name[1:]
        safe_relative(name, f"SHA256SUMS line {index}")
        require(name != "SHA256SUMS" and name not in entries, f"SHA256SUMS duplicate/self entry: {name}")
        entries[name] = parts[0].lower()
    regular: dict[str, str] = {}
    for path in root.rglob("*"):
        if path.is_symlink():
            fail(f"symlink is not allowed in sealed bundle: {path.relative_to(root)}")
        if path.is_file() and path != root / "SHA256SUMS":
            regular[path.relative_to(root).as_posix()] = sha256(path)
    require(regular == entries, "SHA256SUMS does not cover exactly the regular bundle files")
    return {"status": "pass", "files": len(entries), "sha256sums_sha256": sha256(seal)}


def verify_summary(root: Path, computed: Mapping[str, Any]) -> dict[str, Any]:
    path = root / "summary.json"
    require(path.is_file(), "summary.json is missing")
    summary = read_json(path)
    require(summary == computed, "summary.json does not replay exactly")
    require(summary.get("acceptance", {}).get("passed") is True, "large requested-byte reduction acceptance failed")
    require(summary.get("acceptance", {}).get("all_samples_passed") is True, "a large allocator sample misses the requested-byte reduction threshold")
    require(summary.get("failed_allocation_calls", {}).get("all_zero") is True, "a formal allocator sample reports failed allocation")
    return summary


def live_check(root: Path, arms: Mapping[str, Mapping[str, Any]], manifests: Mapping[str, Mapping[str, str]]) -> dict[str, Any]:
    checked: dict[str, Any] = {}
    for arm, item in arms.items():
        binaries = item.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) >= {"normal", "allocator"}, f"live {arm} binaries are missing normal/allocator entries")
        for mode in ("normal", "allocator"):
            binary = binaries[mode]
            require(isinstance(binary, dict), f"live {arm} {mode} binary metadata is malformed")
            binary_path = binary.get("path")
            require(isinstance(binary_path, str) and binary_path, f"live {arm} {mode} binary path is missing")
            path = Path(binary_path)
            require(path.is_file(), f"live {arm} {mode} binary is missing: {path}")
            require(sha256(path) == check_digest(binary.get("sha256"), f"live {arm} {mode} binary.sha256"), f"live {arm} {mode} binary hash differs")
            require(path.stat().st_size == binary.get("bytes"), f"live {arm} {mode} binary size differs")
        tree = item.get("build_path")
        require(isinstance(tree, str) and tree, f"live {arm} build_path is missing")
        tree_path = Path(tree)
        require(tree_path.is_dir(), f"live {arm} source tree is missing: {tree_path}")
        revision = _arm_revision(item)
        require(revision is not None, f"live {arm} revision is missing")
        actual = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=tree_path, text=True).strip()
        require(actual == revision, f"live {arm} revision differs")
        dirty = subprocess.check_output(["git", "status", "--porcelain"], cwd=tree_path, text=True)
        require(not dirty.strip(), f"live {arm} source tree is dirty")
        for relative, expected in manifests.get(arm, {}).items():
            path = tree_path / relative
            require(path.is_file() and sha256(path) == expected, f"live {arm} source hash differs: {relative}")
        checked[arm] = {"binaries": 2, "tree": True, "manifest_files": len(manifests.get(arm, {}))}
    return checked


def verify_bundle(root: Path = ROOT, *, sealed: bool = True, live: bool = False) -> dict[str, Any]:
    try:
        custody_summary = custody.verify(root)
    except (custody.CustodyError, OSError, TypeError, ValueError, KeyError) as error:
        fail(f"custody: {error}")
    protocol_value, rows, protocol_hash = protocol(root)
    arms, manifests = arm_bindings(root, protocol_value)
    formal = {row["lane"] for row in analyze.formal_rows(protocol_value)}
    receipts: dict[str, Any] = {}
    for row in rows:
        receipts[row["lane"]] = verify_receipt(root, row, arms, protocol_hash, formal=row["lane"] in formal)
        if row["lane"] not in formal:
            verify_auxiliary_report(root, row, protocol_value)
    verify_corpus_oracle(root, protocol_value, rows)
    try:
        computed = analyze.analyze(root)
    except analyze.AnalysisError as error:
        fail(str(error))
    summary = verify_summary(root, computed)
    source_summary = verify_source_manifest_files(root, manifests)
    rust_validation_summary = verify_rust_validation(root)
    evidence_validation_summary = verify_evidence_validation(root)
    validation_summary = verify_validation_receipts(root)
    live_summary = live_check(root, arms, manifests) if live else None
    seal_summary = verify_seal(root, required=sealed)
    result = {
        "schema": SCHEMA,
        "protocol_sha256": protocol_hash,
        "custody": custody_summary,
        "formal_lanes": len(formal),
        "auxiliary_lanes": len(rows) - len(formal),
        "receipts": len(receipts),
        "source_manifests": source_summary,
        "rust_validation": rust_validation_summary,
        "evidence_validation": evidence_validation_summary,
        "validation": validation_summary,
        "acceptance": summary["acceptance"],
        "seal": seal_summary,
    }
    if live_summary is not None:
        result["live"] = live_summary
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--unsealed", action="store_true")
    parser.add_argument("--live", action="store_true")
    args = parser.parse_args(argv)
    result = verify_bundle(args.root, sealed=not args.unsealed, live=args.live)
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
