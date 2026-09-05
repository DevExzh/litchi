#!/usr/bin/env python3
"""Replay and summarize the portable 0421 allocator evidence bundle.

The summary is a correctness/custody report.  It intentionally never compares
elapsed samples and never needs the original source worktree or copied binary.
It reruns the retained single-report verifier through the shared repository
validators.  ``--replay`` checks the journals, reports, catalogs, logs,
protocol, and build record against the already rendered summary and table.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


CHANGE = 421
CPU = 2
REPEATS = ("R1", "R2")
SELECTORS = (
    "pptx_cross_copy_media_rich_lifecycle",
    "pptx_cross_copy_plain_lifecycle",
)
SAMPLES = 30
WARMUPS = 3
SHA256 = re.compile(r"^[0-9a-f]{64}$")
RSS = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE)
VECTOR_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
)


class SummaryError(RuntimeError):
    """A custody, schema, or correctness failure."""


def fail(message: str) -> None:
    raise SummaryError(message)


def strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON number {value!r}")


def load_json(path: Path, label: str) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=strict_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"cannot load {label} {path}: {error}")


def write_json(path: Path, value: Any) -> None:
    try:
        path.write_text(
            json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False)
            + "\n",
            encoding="utf-8",
        )
    except (OSError, TypeError, ValueError, OverflowError) as error:
        fail(f"cannot write {path}: {error}")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def digest_json(value: Any) -> str:
    try:
        encoded = json.dumps(
            value, sort_keys=True, separators=(",", ":"),
            ensure_ascii=False, allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot hash JSON identity: {error}")
    return hashlib.sha256(encoded).hexdigest()


def object_value(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def list_value(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        fail(f"{label} must be a list")
    return value


def string_value(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a non-empty string")
    return value


def sha_value(value: Any, label: str) -> str:
    value = string_value(value, label).lower()
    if SHA256.fullmatch(value) is None:
        fail(f"{label} must be a lowercase SHA-256")
    return value


def chronological_values(sample_order: list[int], values: list[int]) -> list[int]:
    """Restore original sample order from vectors aligned to elapsed order."""

    if len(sample_order) != len(values) or sorted(sample_order) != list(range(len(values))):
        fail("allocator sample order is not a complete permutation")
    return [value for _, value in sorted(zip(sample_order, values), key=lambda pair: pair[0])]


def integer_value(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label} must be an integer >= {minimum}")
    return value


def relative(root: Path, value: Any, label: str) -> Path:
    raw = string_value(value, label)
    path = Path(raw)
    if path.is_absolute() or ".." in path.parts:
        fail(f"{label} must be a relative path without '..'")
    resolved = (root / path).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError:
        fail(f"{label} escapes the evidence root")
    return resolved


def artifact(root: Path, value: Any, label: str, *, allow_empty: bool) -> tuple[Path, dict[str, Any]]:
    entry = object_value(value, label)
    path = relative(root, entry.get("path"), f"{label}.path")
    if not path.is_file():
        fail(f"{label} is missing: {path}")
    size = path.stat().st_size
    if size == 0 and not allow_empty:
        fail(f"{label} is empty")
    expected_bytes = integer_value(entry.get("bytes"), f"{label}.bytes")
    expected_sha = sha_value(entry.get("sha256"), f"{label}.sha256")
    if size != expected_bytes or sha256_file(path) != expected_sha:
        fail(f"{label} custody hash or byte count does not match")
    if entry.get("allow_empty") is not allow_empty:
        fail(f"{label}.allow_empty does not match the fixed log contract")
    return path, {"path": str(path.relative_to(root)), "bytes": size, "sha256": expected_sha}


def load_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = root / "protocol.json"
    protocol = object_value(load_json(path, "protocol"), str(path))
    protocol_sha = sha256_file(path)
    if protocol.get("change") != CHANGE or protocol.get("mode") != "allocator":
        fail("protocol is not the frozen 0421 allocator protocol")
    if protocol.get("cpu") != CPU or protocol.get("workers") != 1:
        fail("protocol CPU/worker identity changed")
    if protocol.get("repeats") != 2 or protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        fail("protocol sample contract changed")
    if protocol.get("selectors") != list(SELECTORS):
        fail("protocol selector set or order changed")
    flags = protocol.get("common_flags")
    if not isinstance(flags, list) or any(not isinstance(flag, str) or not flag for flag in flags):
        fail("protocol.common_flags is invalid")
    return protocol, protocol_sha


def load_build(root: Path, protocol_sha: str) -> tuple[dict[str, Any], str, dict[str, Any]]:
    path = root / "build-candidate.json"
    build = object_value(load_json(path, "candidate build"), str(path))
    build_sha = sha256_file(path)
    if build.get("change") != CHANGE or build.get("role") != "candidate":
        fail("candidate build has the wrong identity")
    if build.get("status") != "pass" or build.get("exit_code") != 0:
        fail("candidate build is not a passing record")
    if str(build.get("protocol_sha256", "")).lower() != protocol_sha:
        fail("candidate build is not bound to protocol.json")
    before = object_value(build.get("source_before"), "build.source_before")
    after = object_value(build.get("source_after"), "build.source_after")
    if before != after or before.get("clean") is not True or before.get("git_status_porcelain") != "":
        fail("candidate build source identities are not an unchanged clean pair")
    binary = object_value(object_value(build.get("binaries"), "build.binaries").get("allocator"), "build allocator")
    binary_sha = sha_value(binary.get("sha256", binary.get("binary_sha256")), "build allocator sha256")
    binary_bytes = integer_value(binary.get("bytes", binary.get("binary_bytes")), "build allocator bytes", 1)
    return build, build_sha, {
        "sha256": binary_sha,
        "bytes": binary_bytes,
        "path": string_value(binary.get("path"), "build allocator path"),
    }


def find_repo_root(explicit: Path | None) -> Path:
    if explicit is not None:
        candidate = explicit.expanduser().resolve()
        if not (candidate / "tools" / "perf_abba_summary.py").is_file():
            fail(f"--repo-root lacks the retained shared validators: {candidate}")
        return candidate
    here = Path(__file__).resolve()
    for candidate in (here, *here.parents):
        if (candidate / "tools" / "perf_abba_summary.py").is_file():
            return candidate
    fail("cannot locate the shared validators; pass --repo-root")
    raise AssertionError("unreachable")


def rerun_verifier(
    repo_root: Path, report: Path, catalog: Path, selector: str,
    expected_summary_sha: str, label: str,
) -> dict[str, Any]:
    verifier = Path(__file__).resolve().with_name("verify.py")
    if not verifier.is_file():
        fail(f"{label} cannot locate verify.py for replay")
    command = [
        sys.executable, str(verifier), "--repo-root", str(repo_root),
        "--report", str(report), "--catalog", str(catalog),
        "--selector", selector, "--lane", "allocator",
        "--samples", str(SAMPLES), "--warmups", str(WARMUPS),
    ]
    try:
        process = subprocess.run(
            command, cwd=repo_root, capture_output=True, text=True, check=False,
        )
    except OSError as error:
        fail(f"{label} verifier could not start: {error}")
    if process.returncode != 0:
        detail = (process.stderr or process.stdout).strip()
        fail(f"{label} replay verifier rejected the report: {detail}")
    try:
        value = json.loads(
            process.stdout, object_pairs_hook=strict_pairs,
            parse_constant=reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label} replay verifier emitted invalid JSON: {error}")
    value = object_value(value, f"{label}.replay_verifier")
    if digest_json(value) != expected_summary_sha:
        fail(f"{label} replay verifier result differs from retained verifier custody")
    if value.get("claim_authorized") is not False or value.get("performance_claim") is not None:
        fail(f"{label} replay verifier unexpectedly authorizes a claim")
    return value


def verify_journal_artifacts(root: Path, journal: dict[str, Any], label: str) -> dict[str, dict[str, Any]]:
    raw = object_value(journal.get("artifacts"), f"{label}.artifacts")
    expected = {
        "report": False,
        "catalog": False,
        "time_v": False,
        "stdout": True,
        "stderr": True,
        "verify_stdout": False,
        "verify_stderr": True,
    }
    if set(raw) != set(expected):
        fail(f"{label}.artifacts has an unexpected set of files")
    result: dict[str, dict[str, Any]] = {}
    parent: Path | None = None
    for name, allow_empty in expected.items():
        path, record = artifact(root, raw[name], f"{label}.artifacts.{name}", allow_empty=allow_empty)
        if parent is None:
            parent = path.parent
        elif path.parent != parent:
            fail(f"{label}.artifacts are not confined to one run directory")
        result[name] = record
    return result


def parse_time(path: Path, label: str) -> int:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot read {label}: {error}")
    exits = re.findall(r"^\s*Exit status:\s*(-?\d+)\s*$", text, re.MULTILINE)
    if exits != ["0"]:
        fail(f"{label} must contain exactly one successful Exit status")
    matches = RSS.findall(text)
    if len(matches) != 1:
        fail(f"{label} must contain exactly one Maximum resident set size")
    return int(matches[0])


def validate_report(
    report: dict[str, Any], catalog: dict[str, Any], *, label: str,
    selector: str, journal: dict[str, Any], report_sha: str,
) -> dict[str, Any]:
    expected_tool = {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": "litchi-perf-baseline-alloc",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": "system_allocator_operation_scoped",
        "allocator_counter_revision": "post_update_peak_v2",
    }
    if report.get("tool") != expected_tool:
        fail(f"{label}.tool is not the post_update_peak_v2 allocator identity")
    if report.get("schema_version") != 1 or not isinstance(report.get("results"), list) or len(report["results"]) != 1:
        fail(f"{label} has an invalid report envelope")
    binary_identity = object_value(report.get("binary_identity"), f"{label}.binary_identity")
    if binary_identity.get("binary_sha256") != journal["binary_sha256"] or binary_identity.get("binary_bytes") != journal["binary_bytes"]:
        fail(f"{label}.binary_identity is not bound to the journal")
    environment = object_value(report.get("environment"), f"{label}.environment")
    if environment.get("git_revision") != journal["source_revision"] or environment.get("git_worktree_dirty") is not False:
        fail(f"{label}.environment source identity differs from journal")
    configuration = object_value(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != SAMPLES or configuration.get("warmup_iterations_per_case") != WARMUPS or configuration.get("cases") != [selector]:
        fail(f"{label}.configuration does not match the allocator contract")
    result = object_value(report["results"][0], f"{label}.results[0]")
    if result.get("case") != selector:
        fail(f"{label}.results[0].case differs from journal")
    operation = object_value(result.get("operation_metrics"), f"{label}.operation_metrics")
    elapsed = object_value(result.get("elapsed_ns"), f"{label}.elapsed_ns")
    sample_order = list_value(elapsed.get("sample_order"), f"{label}.elapsed_ns.sample_order")
    if len(sample_order) != SAMPLES or sorted(sample_order) != list(range(SAMPLES)):
        fail(f"{label}.elapsed_ns.sample_order is not a complete sample permutation")
    if operation.get("sample_indices") != sample_order:
        fail(f"{label}.operation_metrics.sample_indices is not aligned to elapsed samples")
    allocation = object_value(operation.get("allocation"), f"{label}.operation_metrics.allocation")
    if allocation.get("status") != "measured" or allocation.get("scope") != "operation_global_system_allocator":
        fail(f"{label}.allocation is not measured operation-scoped evidence")
    vectors: dict[str, list[int]] = {}
    for field in VECTOR_FIELDS:
        metric = object_value(allocation.get(field), f"{label}.allocation.{field}")
        if metric.get("status") != "measured" or metric.get("scope") != "operation_global_system_allocator":
            fail(f"{label}.allocation.{field} status/scope changed")
        values = list_value(metric.get("values"), f"{label}.allocation.{field}.values")
        if len(values) != SAMPLES:
            fail(f"{label}.allocation.{field}.values must contain {SAMPLES} values")
        if any(isinstance(value, bool) or not isinstance(value, int) or value < 0 for value in values):
            fail(f"{label}.allocation.{field}.values contains an invalid counter")
        vectors[field] = list(values)
    before_ok = all(peak >= live for peak, live in zip(vectors["peak_live_bytes_before"], vectors["live_bytes_before"]))
    after_ok = all(peak >= live for peak, live in zip(vectors["peak_live_bytes_after"], vectors["live_bytes_after"]))
    chronological_before = chronological_values(sample_order, vectors["peak_live_bytes_before"])
    chronological_after = chronological_values(sample_order, vectors["peak_live_bytes_after"])
    before_monotonic = all(left <= right for left, right in zip(chronological_before, chronological_before[1:]))
    after_monotonic = all(left <= right for left, right in zip(chronological_after, chronological_after[1:]))
    after_ge_before = all(
        after >= before
        for before, after in zip(vectors["peak_live_bytes_before"], vectors["peak_live_bytes_after"])
    )
    boundary_monotonic = all(
        next_before >= previous_after
        for previous_after, next_before in zip(chronological_after, chronological_before[1:])
    )
    if not (before_ok and after_ok and before_monotonic and after_monotonic and after_ge_before and boundary_monotonic):
        fail(f"{label} violates peak/live or monotonic-peak invariants")
    output_sha = sha_value(result.get("output_sha256"), f"{label}.output_sha256")
    source = object_value(result.get("source"), f"{label}.source")
    special = object_value(source.get("pptx_cross_copy"), f"{label}.source.pptx_cross_copy")
    if special.get("expected_output_sha256") != output_sha:
        fail(f"{label} output does not match expected output")
    output_vector = special.get("output_sha256")
    if not isinstance(output_vector, list) or len(output_vector) != SAMPLES or any(value != output_sha for value in output_vector):
        fail(f"{label}.source.pptx_cross_copy.output_sha256 is not deterministic")
    catalog_ref = object_value(report.get("corpus_catalog"), f"{label}.corpus_catalog")
    catalog_identity = {
        key: catalog.get(key)
        for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256")
    }
    if catalog_ref != catalog_identity:
        fail(f"{label}.corpus_catalog differs from catalog sidecar")
    return {
        "report_sha256": report_sha,
        "catalog_sha256": digest_json(catalog),
        "catalog_identity_sha256": digest_json(catalog_identity),
        "corpus_sha256": digest_json(result.get("corpus")),
        "output_sha256": output_sha,
        "output_vector_sha256": digest_json(output_vector),
        "sample_order": list(sample_order),
        "sample_order_sha256": digest_json(sample_order),
        "allocation_vectors": vectors,
        "invariants": {
            "peak_before_ge_live_before": before_ok,
            "peak_after_ge_live_after": after_ok,
            "peak_after_ge_peak_before": after_ge_before,
            "peak_before_monotonic": before_monotonic,
            "peak_after_monotonic": after_monotonic,
            "peak_boundary_monotonic": boundary_monotonic,
        },
    }


def allocation_stats(vectors: dict[str, list[int]]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for field in VECTOR_FIELDS:
        values = vectors[field]
        result[field] = {
            "count": len(values),
            "mean": sum(values) / len(values),
            "min": min(values),
            "max": max(values),
        }
    return result


def verify_single_run(
    root: Path, repo_root: Path, manifest_run: dict[str, Any], build_sha: str, protocol_sha: str,
    binary: dict[str, Any], source_identity_sha: str, common_flags: list[str],
) -> dict[str, Any]:
    repeat = manifest_run.get("repeat")
    selector = manifest_run.get("selector")
    if repeat not in REPEATS or selector not in SELECTORS:
        fail("capture run has an unknown repeat or selector")
    journal_path = relative(root, manifest_run.get("journal"), "capture.run.journal")
    if not journal_path.is_file():
        fail(f"missing journal: {journal_path}")
    observed_journal_sha = sha256_file(journal_path)
    if observed_journal_sha != sha_value(manifest_run.get("journal_sha256"), "capture.run.journal_sha256"):
        fail(f"journal custody hash differs: {journal_path}")
    journal = object_value(load_json(journal_path, str(journal_path)), str(journal_path))
    label = f"{repeat}/{selector}"
    if journal.get("status") != "pass" or journal.get("change") != CHANGE or journal.get("repeat") != repeat or journal.get("selector") != selector:
        fail(f"{label} journal is not a passing identity record")
    if journal.get("fresh_process") is not True or journal.get("cpu") != CPU or journal.get("exit_code") != 0:
        fail(f"{label} is not a successful fresh CPU-2 process")
    if journal.get("build_sha256") != build_sha or journal.get("protocol_sha256") != protocol_sha:
        fail(f"{label} build/protocol custody differs")
    if journal.get("binary_sha256") != binary["sha256"] or journal.get("binary_bytes") != binary["bytes"]:
        fail(f"{label} binary custody differs")
    if journal.get("binary_before") != {"sha256": binary["sha256"], "bytes": binary["bytes"]} or journal.get("binary_after") != journal.get("binary_before"):
        fail(f"{label} binary before/after identity changed")
    if journal.get("source_identity_sha256") != source_identity_sha:
        fail(f"{label} source identity custody differs")
    source_after = object_value(journal.get("source_after"), f"{label}.source_after")
    if source_after.get("revision") != journal.get("source_revision") or source_after.get("clean") is not True or source_after.get("git_status_porcelain") != "":
        fail(f"{label} source_after is not the unchanged clean identity")
    if journal.get("common_flags") != common_flags:
        fail(f"{label} common flags differ from protocol")
    contract = journal.get("contract")
    if contract != {"lane": "allocator", "samples": SAMPLES, "warmups": WARMUPS}:
        fail(f"{label} contract differs")
    argv = list_value(journal.get("argv"), f"{label}.argv")
    if argv[:6] != ["taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o"]:
        fail(f"{label}.argv does not pin taskset/time -v")
    if len(argv) < 10 or argv[7:9] != [binary.get("path"), "--case"] or argv[9] != selector:
        fail(f"{label}.argv binary/case binding differs")
    try:
        flag_start = 10
        flag_end = flag_start + len(common_flags)
        if argv[flag_start:flag_end] != common_flags:
            fail(f"{label}.argv common flags differ")
        tail = argv[flag_end:]
        if len(tail) != 8 or tail[:5] != ["--samples", str(SAMPLES), "--warmup", str(WARMUPS), "--json"]:
            fail(f"{label}.argv capture-owned flags differ")
        report_relative = Path(journal["artifacts"]["report"]["path"])
        catalog_relative = Path(journal["artifacts"]["catalog"]["path"])
        relative(root, str(report_relative), f"{label}.report artifact")
        relative(root, str(catalog_relative), f"{label}.catalog artifact")
        recorded_report = Path(tail[5])
        if not recorded_report.is_absolute() or ".." in recorded_report.parts:
            fail(f"{label}.argv report path must be absolute and normalized")
        # Commands retain the original capture location. Resolve artifact
        # contents below the current export, but bind all command paths to
        # their original common root instead of requiring that root to exist.
        recorded_root = recorded_report.parents[len(report_relative.parts) - 1]
        if recorded_report != recorded_root / report_relative or Path(tail[7]) != recorded_root / catalog_relative or tail[6] != "--corpus-manifest":
            fail(f"{label}.argv report/catalog paths differ")
        if Path(argv[6]) != recorded_root / journal["artifacts"]["time_v"]["path"]:
            fail(f"{label}.argv time-v path differs")
    except (IndexError, KeyError, TypeError):
        fail(f"{label}.argv is incomplete")
    verify_argv = list_value(journal.get("verify_argv"), f"{label}.verify_argv")
    try:
        lane_position = verify_argv.index("--lane")
        warmups_present = "--warmups" in verify_argv
        lane_valid = lane_position + 1 < len(verify_argv) and verify_argv[lane_position + 1] == "allocator"
    except ValueError:
        lane_valid = False
        warmups_present = False
    if not lane_valid or not warmups_present:
        fail(f"{label}.verify_argv does not use the exact allocator verifier flags")
    artifacts = verify_journal_artifacts(root, journal, label)
    report_path = root / artifacts["report"]["path"]
    catalog_path = root / artifacts["catalog"]["path"]
    report = object_value(load_json(report_path, f"{label}.report"), f"{label}.report")
    catalog = object_value(load_json(catalog_path, f"{label}.catalog"), f"{label}.catalog")
    report_proof = validate_report(
        report, catalog, label=label, selector=selector, journal=journal,
        report_sha=artifacts["report"]["sha256"],
    )
    verification = object_value(load_json(root / artifacts["verify_stdout"]["path"], f"{label}.verifier"), f"{label}.verifier")
    if verification.get("claim_authorized") is not False or verification.get("performance_claim") is not None or verification.get("selector") != selector or verification.get("lane") != "allocator" or verification.get("samples") != SAMPLES or verification.get("warmups") != WARMUPS:
        fail(f"{label} retained verifier result does not withhold claims")
    verified_reports = list_value(verification.get("reports"), f"{label}.verifier.reports")
    if len(verified_reports) != 1 or object_value(verified_reports[0], f"{label}.verifier.reports[0]").get("report_sha256") != report_proof["report_sha256"]:
        fail(f"{label} retained verifier result is not bound to the report")
    if journal.get("verification", {}).get("status") != "pass" or journal["verification"].get("summary_sha256") != digest_json(verification):
        fail(f"{label} verifier custody differs")
    rerun_verifier(
        repo_root, report_path, catalog_path, selector,
        journal["verification"]["summary_sha256"], label,
    )
    rss = parse_time(root / artifacts["time_v"]["path"], f"{label}.time-v")
    report_proof["allocation_stats"] = allocation_stats(report_proof["allocation_vectors"])
    report_proof["whole_process_rss_kib"] = rss
    report_proof.update({
        "repeat": repeat,
        "selector": selector,
        "journal_sha256": observed_journal_sha,
        "build_sha256": build_sha,
        "protocol_sha256": protocol_sha,
        "binary_sha256": journal["binary_sha256"],
        "source_revision": journal["source_revision"],
        "source_identity_sha256": journal["source_identity_sha256"],
        "artifact_hashes": {name: item["sha256"] for name, item in artifacts.items()},
        "verification_summary_sha256": journal["verification"]["summary_sha256"],
    })
    return report_proof


def render_table(rows: list[dict[str, Any]]) -> str:
    lines = [
        "# 0421 allocator correctness diagnostic",
        "",
        "Four fresh allocator processes (two selectors, two repeats), 30 samples and 3 warmups each. Elapsed samples are deliberately omitted; this evidence makes no performance claim.",
        "",
        "| Repeat | Selector | Allocation calls mean | Allocated bytes mean | Live-before mean | Live-after mean | Peak-before mean | Peak-after mean | Whole-process RSS KiB |",
        "|---|---|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in rows:
        stats = row["allocation_stats"]
        lines.append(
            "| {repeat} | {selector} | {calls:.3f} | {allocated:.3f} | {before:.3f} | {after:.3f} | {peak_before:.3f} | {peak_after:.3f} | {rss:,} |".format(
                repeat=row["repeat"], selector=row["selector"],
                calls=stats["allocation_calls"]["mean"],
                allocated=stats["allocated_bytes"]["mean"],
                before=stats["live_bytes_before"]["mean"],
                after=stats["live_bytes_after"]["mean"],
                peak_before=stats["peak_live_bytes_before"]["mean"],
                peak_after=stats["peak_live_bytes_after"]["mean"],
                rss=row["whole_process_rss_kib"],
            )
        )
    lines.extend([
        "",
        "The JSON retains every allocation vector and each vector's count, mean, minimum, and maximum. Peak-before/live-before, peak-after/live-after, and peak-after/peak-before hold for every sample; peak high-water and inter-sample boundaries are checked after restoring chronological sample order. RSS is whole-process GNU `time -v` evidence; no elapsed-time or allocator-speed comparison is made.",
        "",
    ])
    return "\n".join(lines)


def summarize(root: Path, replay: bool, repo_root: Path | None) -> None:
    root = root.expanduser().resolve()
    manifest = object_value(load_json(root / "capture.json", "capture"), "capture")
    if manifest.get("status") != "pass" or manifest.get("change") != CHANGE or manifest.get("claim_authorized") is not False or manifest.get("performance_claim") is not None:
        fail("capture manifest is not a passing no-claim 0421 capture")
    protocol, protocol_sha = load_protocol(root)
    build, build_sha, binary = load_build(root, protocol_sha)
    if manifest.get("protocol", {}).get("sha256") != protocol_sha or manifest.get("build", {}).get("sha256") != build_sha:
        fail("capture manifest protocol/build custody differs")
    if manifest.get("binary", {}).get("sha256") != binary["sha256"] or manifest.get("binary", {}).get("bytes") != binary["bytes"]:
        fail("capture manifest binary custody differs")
    source_identity_sha = manifest.get("source", {}).get("identity_sha256")
    if not isinstance(source_identity_sha, str) or source_identity_sha != digest_json(build["source_before"]):
        fail("capture manifest source custody differs")
    if manifest.get("repeats") != list(REPEATS) or manifest.get("selectors") != list(SELECTORS) or manifest.get("samples") != SAMPLES or manifest.get("warmups") != WARMUPS or manifest.get("cpu") != CPU:
        fail("capture manifest contract differs")
    raw_runs = list_value(manifest.get("runs"), "capture.runs")
    if len(raw_runs) != 4:
        fail("capture must contain exactly four runs")
    expected = [(repeat, selector) for repeat in REPEATS for selector in SELECTORS]
    observed = [(item.get("repeat"), item.get("selector")) for item in raw_runs]
    if observed != expected:
        fail(f"capture run order must be {expected!r}")
    repo_root = find_repo_root(repo_root)
    rows = [
        verify_single_run(
            root, repo_root, item, build_sha, protocol_sha, binary, source_identity_sha,
            protocol["common_flags"],
        )
        for item in raw_runs
    ]
    by_selector: dict[str, list[dict[str, Any]]] = {selector: [] for selector in SELECTORS}
    for row in rows:
        by_selector[row["selector"]].append(row)
    repeat_gates: dict[str, Any] = {}
    for selector, selector_rows in by_selector.items():
        if len(selector_rows) != 2:
            fail(f"{selector} does not have both repeats")
        first, second = selector_rows
        fields = ("corpus_sha256", "catalog_identity_sha256", "output_sha256", "output_vector_sha256")
        for field in fields:
            if first[field] != second[field]:
                fail(f"{selector} differs across repeats in {field}")
        repeat_gates[selector] = {
            "corpus_same_across_repeats": True,
            "catalog_identity_same_across_repeats": True,
            "output_same_across_repeats": True,
            "corpus_sha256": first["corpus_sha256"],
            "catalog_identity_sha256": first["catalog_identity_sha256"],
            "output_sha256": first["output_sha256"],
        }
    summary = {
        "change": CHANGE,
        "classification": "allocator high-water correctness diagnostic; no optimization claim",
        "claim_authorized": False,
        "performance_claim": None,
        "claim_withheld_reason": "0421 registers allocator correctness/resource evidence and excludes elapsed comparison",
        "protocol": {"path": "protocol.json", "sha256": protocol_sha},
        "build": {"path": "build-candidate.json", "sha256": build_sha},
        "binary": binary,
        "source": {"revision": build["source_before"]["revision"], "identity_sha256": source_identity_sha},
        "configuration": {"cpu": CPU, "repeats": 2, "selectors": list(SELECTORS), "samples": SAMPLES, "warmups": WARMUPS, "lane": "allocator"},
        "repeat_gates": repeat_gates,
        "invariants": {
            "every_peak_before_ge_live_before": all(row["invariants"]["peak_before_ge_live_before"] for row in rows),
            "every_peak_after_ge_live_after": all(row["invariants"]["peak_after_ge_live_after"] for row in rows),
            "every_peak_after_ge_peak_before": all(row["invariants"]["peak_after_ge_peak_before"] for row in rows),
            "peak_before_monotonic_for_every_run": all(row["invariants"]["peak_before_monotonic"] for row in rows),
            "peak_after_monotonic_for_every_run": all(row["invariants"]["peak_after_monotonic"] for row in rows),
            "peak_boundary_monotonic_for_every_run": all(row["invariants"]["peak_boundary_monotonic"] for row in rows),
        },
        "runs": rows,
        "limitations": [
            "No allocator elapsed-time comparison or performance improvement claim is made.",
            "Allocation counters describe requested operations and process snapshots; they do not measure physical realloc overlap or operation-local peaks.",
            "RSS is whole-process GNU time -v maximum resident set size, including setup and warmups.",
            "Replay validates retained verifier output and custody hashes without requiring the original worktree or binary.",
        ],
    }
    summary_path = root / "summary.json"
    table_path = root / "result-table.md"
    table = render_table(rows)
    rendered = json.dumps(summary, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False) + "\n"
    if replay:
        try:
            if summary_path.read_text(encoding="utf-8") != rendered:
                fail("summary.json differs from deterministic replay")
            if table_path.read_text(encoding="utf-8") != table:
                fail("result-table.md differs from deterministic replay")
        except OSError as error:
            fail(f"replay output is missing: {error}")
    else:
        if summary_path.exists() or table_path.exists():
            fail("refusing to overwrite existing summary outputs; use --replay")
        write_json(summary_path, summary)
        try:
            table_path.write_text(table, encoding="utf-8")
        except OSError as error:
            fail(f"cannot write result-table.md: {error}")
    print(json.dumps({"status": "pass", "change": CHANGE, "runs": 4, "replay": replay}, sort_keys=True))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--replay", action="store_true")
    args = parser.parse_args()
    try:
        summarize(args.root, args.replay, args.repo_root)
    except SummaryError as error:
        print(f"0421 summary failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
