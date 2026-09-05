#!/usr/bin/env python3
"""Fail-closed verifier for the change-0419 allocation diagnostic reports.

The 0419 capture is intentionally smaller than the registered latency batch:
it first validates one heaptrack report (one sample, no warmups), then may
validate two four-leg ABBA groups.  This file is a path-oriented wrapper
around the stable report, corpus, operation-metric, and PPTX-row validators in
``tools``.  It does not turn the diagnostic samples into a latency claim.

Examples::

    python3 verify.py --report heaptrack/control/report.json \
        --catalog heaptrack/control/catalog.json --selector \
        pptx_cross_copy_media_rich_lifecycle --samples 1 --warmups 0 \
        --capture heaptrack/control/capture.json

    python3 verify.py --root . --mode normal \
        --selector pptx_cross_copy_plain_lifecycle --samples 100 --warmups 10

The second form reads ``runs/{mode}/{leg}/{selector}/report.json`` and its
neighboring catalog for A1, B1, B2, and A2.  ``--report``/``--reports`` and
``--catalog``/``--catalogs`` may be used for exported bundles whose paths do
not follow that layout.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import math
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Iterable, Mapping


EXPECTED_CHANGE = 419
ABBA_LEGS = ("A1", "B1", "B2", "A2")
ABBA_ROLES = ("control", "candidate", "candidate", "control")
NORMAL_BINARY = "litchi-perf-baseline"
ALLOCATOR_BINARY = "litchi-perf-baseline-alloc"
NORMAL_INSTRUMENTATION = "none"
ALLOCATOR_INSTRUMENTATION = "system_allocator_operation_scoped"
NORMAL_CLAIM_MIN_SAMPLES = 500
MAX_JSON_BYTES = 512 * 1024 * 1024
ZSTD_TIMEOUT_SECONDS = 120
SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")

_PPTX_DYNAMIC_FIELDS = frozenset(
    {"plan_ns", "commit_ns", "publication_ns", "reopen_ns", "lifecycle_ns", "output_sha256"}
)
_PPTX_DESCRIPTIVE_FIELDS = frozenset(
    {"implementation", "timing_scope", "performance_claim"}
)


class VerificationError(ValueError):
    """A report or its bound evidence violates the diagnostic contract."""


def fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")


def object_value(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "must be an object")
    return value


def list_value(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "must be a list")
    return value


def string_value(value: Any, path: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        fail(path, "must be a non-empty string" if nonempty else "must be a string")
    return value


def integer_value(value: Any, path: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(path, f"must be an integer at least {minimum}")
    return value


def sha256_value(value: Any, path: str) -> str:
    value = string_value(value, path)
    if SHA256_RE.fullmatch(value) is None:
        fail(path, "must be a 64-character hexadecimal SHA-256")
    return value.lower()


def canonical(value: Any, path: str = "value") -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError, OverflowError) as error:
        fail(path, f"is not canonical JSON: {error}")
    raise AssertionError("unreachable")


def digest_json(value: Any, path: str = "value") -> str:
    return hashlib.sha256(canonical(value, path).encode("utf-8")).hexdigest()


def _strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def _reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON constant {value!r}")


def load_json(path: Path, label: str) -> tuple[dict[str, Any], str]:
    """Load a bounded plain or zstd-compressed JSON artifact."""

    try:
        if not path.is_file():
            fail(label, "file is missing")
        if path.stat().st_size > MAX_JSON_BYTES:
            fail(label, f"compressed/raw size exceeds {MAX_JSON_BYTES} bytes")
    except OSError as error:
        fail(label, f"cannot inspect file: {error}")

    try:
        if path.name.endswith(".zst"):
            process = subprocess.Popen(
                ["zstd", "-q", "-dc", str(path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            try:
                raw, stderr = process.communicate(timeout=ZSTD_TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                process.kill()
                process.communicate()
                fail(label, "zstd decompression timed out")
            if process.returncode != 0:
                detail = stderr.decode("utf-8", "replace").strip()
                fail(label, f"zstd decompression failed{': ' + detail if detail else ''}")
        else:
            raw = path.read_bytes()
    except FileNotFoundError:
        fail(label, "zstd executable is unavailable")
    except OSError as error:
        fail(label, f"cannot read file: {error}")
    if len(raw) > MAX_JSON_BYTES:
        fail(label, f"decoded JSON exceeds {MAX_JSON_BYTES} bytes")

    try:
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=_strict_pairs,
            parse_constant=_reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        fail(label, f"invalid JSON: {error}")
    return object_value(value, label), hashlib.sha256(raw).hexdigest()


def find_repo_root(explicit: Path | None) -> Path:
    if explicit is not None:
        root = explicit.expanduser().resolve()
        if not (root / "tools" / "perf_abba_summary.py").is_file():
            fail("--repo-root", "does not contain tools/perf_abba_summary.py")
        return root
    here = Path(__file__).resolve()
    for candidate in (here, *here.parents):
        if (candidate / "tools" / "perf_abba_summary.py").is_file():
            return candidate
    candidate = Path.cwd().resolve()
    if (candidate / "tools" / "perf_abba_summary.py").is_file():
        return candidate
    fail("repo", "cannot locate tools/perf_abba_summary.py; pass --repo-root")
    raise AssertionError("unreachable")


def import_validators(repo_root: Path) -> tuple[Any, Any, Any]:
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    try:
        from tools import perf_abba_summary, perf_compare, validate_perf_corpus_binding
    except ImportError as error:
        fail("repo", f"cannot import stable validators: {error}")
    return perf_abba_summary, perf_compare, validate_perf_corpus_binding


def validate_elapsed(
    value: Any,
    path: str,
    expected_samples: int,
    perf_abba_summary: Any,
) -> tuple[list[int], list[int], dict[str, Any]]:
    elapsed = object_value(value, path)
    if elapsed.get("unit") != "ns":
        fail(f"{path}.unit", "must be 'ns'")
    raw_samples = list_value(elapsed.get("samples"), f"{path}.samples")
    if len(raw_samples) != expected_samples:
        fail(f"{path}.samples", f"must contain exactly {expected_samples} samples")
    samples = [
        integer_value(item, f"{path}.samples[{index}]", minimum=1)
        for index, item in enumerate(raw_samples)
    ]
    if samples != sorted(samples):
        fail(f"{path}.samples", "must be sorted ascending")

    raw_order = elapsed.get("sample_order")
    if raw_order is None:
        if expected_samples != 1:
            fail(f"{path}.sample_order", "is required when more than one sample is retained")
        order = [0]
    else:
        order_values = list_value(raw_order, f"{path}.sample_order")
        if len(order_values) != expected_samples:
            fail(f"{path}.sample_order", f"must contain exactly {expected_samples} entries")
        order = [
            integer_value(item, f"{path}.sample_order[{index}]")
            for index, item in enumerate(order_values)
        ]
        if sorted(order) != list(range(expected_samples)):
            fail(f"{path}.sample_order", "must be a complete sample permutation")

    if expected_samples >= 15:
        try:
            stats = perf_abba_summary.recompute_statistics(elapsed, path)
        except Exception as error:
            fail(path, f"elapsed statistics rejected: {error}")
    else:
        # The shared recomputer deliberately has the registered 15-sample
        # floor.  Diagnostic one/ten-sample rows still need finite, coherent
        # metadata, but are never passed to a claim-producing summary.
        for name in ("min", "p50", "p95", "p99", "max"):
            integer_value(elapsed.get(name), f"{path}.{name}")
        if any(elapsed[name] != expected for name, expected in {
            "min": samples[0],
            "max": samples[-1],
            "p50": (
                samples[(len(samples) - 1) // 2] // 2
                + samples[len(samples) // 2] // 2
                + (
                    samples[(len(samples) - 1) // 2] % 2
                    + samples[len(samples) // 2] % 2
                )
                // 2
            ),
            "p95": samples[min(((95 * len(samples) + 99) // 100) - 1, len(samples) - 1)],
            "p99": samples[min(((99 * len(samples) + 99) // 100) - 1, len(samples) - 1)],
        }.items()):
            fail(path, "reported integer statistics disagree with retained samples")
        mean = elapsed.get("mean")
        standard_deviation = elapsed.get("standard_deviation")
        if (
            isinstance(mean, bool)
            or not isinstance(mean, (int, float))
            or not math.isfinite(float(mean))
            or float(mean) < 0
            or isinstance(standard_deviation, bool)
            or not isinstance(standard_deviation, (int, float))
            or not math.isfinite(float(standard_deviation))
            or float(standard_deviation) < 0
        ):
            fail(path, "mean and standard_deviation must be finite non-negative numbers")
        confidence = object_value(elapsed.get("confidence_interval_95"), f"{path}.confidence_interval_95")
        if confidence.get("method") != "two-sided Student's t interval for the mean":
            fail(f"{path}.confidence_interval_95.method", "does not match the harness")
        for name in ("lower", "upper"):
            number = confidence.get(name)
            if (
                isinstance(number, bool)
                or not isinstance(number, (int, float))
                or not math.isfinite(float(number))
                or float(number) < 0
            ):
                fail(f"{path}.confidence_interval_95.{name}", "must be finite and non-negative")
        if float(confidence["lower"]) > float(confidence["upper"]):
            fail(f"{path}.confidence_interval_95", "lower must not exceed upper")
        stats = {
            "sample_count": expected_samples,
            "min": samples[0],
            "p50": elapsed["p50"],
            "p95": elapsed["p95"],
            "p99": elapsed["p99"],
            "max": samples[-1],
            "mean": float(mean),
            "standard_deviation": float(standard_deviation),
        }
    return samples, order, stats


def expected_tool(lane: str) -> dict[str, str]:
    if lane == "normal":
        return {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": NORMAL_BINARY,
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": NORMAL_INSTRUMENTATION,
        }
    return {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": ALLOCATOR_BINARY,
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": ALLOCATOR_INSTRUMENTATION,
    }


def validate_operation(
    result: dict[str, Any],
    path: str,
    selector: str,
    lane: str,
    perf_compare: Any,
    sample_order: list[int],
) -> tuple[dict[str, Any], str | None]:
    operation = object_value(result.get("operation_metrics"), f"{path}.operation_metrics")
    elapsed = object_value(result.get("elapsed_ns"), f"{path}.elapsed_ns")
    try:
        perf_compare._validate_operation_metrics(
            operation,
            f"{path}.operation_metrics",
            elapsed["samples"],
            1,
            elapsed_sample_order=sample_order,
        )
    except Exception as error:
        fail(path, f"operation metrics rejected: {error}")
    allocation = operation.get("allocation")
    if selector.endswith("_lifecycle"):
        if allocation is None:
            fail(f"{path}.operation_metrics", "lifecycle rows must carry allocation status")
        allocation_object = object_value(allocation, f"{path}.operation_metrics.allocation")
        expected_status = "measured" if lane == "allocator" else "unavailable"
        if allocation_object.get("status") != expected_status:
            fail(
                f"{path}.operation_metrics.allocation.status",
                f"must be {expected_status!r} for the {lane} lane",
            )
        if lane == "allocator":
            try:
                perf_compare._validate_allocator_operation_evidence(result, path, 1)
            except Exception as error:
                fail(path, f"allocator operation evidence rejected: {error}")
        return operation, expected_status
    if allocation is not None:
        fail(f"{path}.operation_metrics.allocation", "phase rows must omit allocation")
    return operation, None


def validate_report(
    report: dict[str, Any],
    catalog: dict[str, Any],
    path: str,
    *,
    selector: str,
    samples: int,
    warmups: int,
    lane: str,
    perf_abba_summary: Any,
    perf_compare: Any,
    corpus_binding: Any,
    raw_sha256: str,
) -> dict[str, Any]:
    expected_top = {
        "schema_version",
        "tool",
        "binary_identity",
        "environment",
        "configuration",
        "parallel_metrics",
        "results",
        "corpus_catalog",
    }
    if set(report) != expected_top:
        fail(path, "report has an unexpected top-level schema")
    if report.get("schema_version") != 1:
        fail(f"{path}.schema_version", "must be 1")

    tool = object_value(report.get("tool"), f"{path}.tool")
    expected = expected_tool(lane)
    if tool != expected:
        fail(f"{path}.tool", f"does not match the {lane} instrumentation contract")
    try:
        if perf_abba_summary.detect_report_profile(report, path) != "current-v1":
            fail(f"{path}.tool", "legacy reports are not accepted by this diagnostic")
        # The shared summary validator intentionally accepts only the
        # uninstrumented binary because it is a latency-summary boundary.
        # 0419's allocator rows use the same report schema but a distinct
        # executable/instrumentation identity, already checked above.
        if lane == "normal":
            perf_abba_summary._validate_tool(tool, path, "current-v1")
    except VerificationError:
        raise
    except Exception as error:
        fail(f"{path}.tool", f"shared validator rejected tool identity: {error}")

    try:
        binary = perf_abba_summary._validate_binary_identity(
            report.get("binary_identity"), path, tool
        )
        perf_abba_summary._validate_environment(report.get("environment"), path)
    except Exception as error:
        fail(path, f"shared identity validator rejected report: {error}")
    try:
        perf_compare.validate_parallel_metrics(report, path)
    except Exception as error:
        fail(path, f"parallel metrics rejected: {error}")

    configuration = object_value(report.get("configuration"), f"{path}.configuration")
    configured_samples = integer_value(
        configuration.get("samples_per_case"),
        f"{path}.configuration.samples_per_case",
        minimum=1,
    )
    if configured_samples != samples:
        fail(f"{path}.configuration.samples_per_case", f"must be {samples}")
    configured_warmups = integer_value(
        configuration.get("warmup_iterations_per_case"),
        f"{path}.configuration.warmup_iterations_per_case",
    )
    if configured_warmups != warmups:
        fail(f"{path}.configuration.warmup_iterations_per_case", f"must be {warmups}")
    if configuration.get("cases") != [selector]:
        fail(f"{path}.configuration.cases", "must contain exactly the selected case")
    corpus_shapes = list_value(configuration.get("corpus_shapes"), f"{path}.configuration.corpus_shapes")
    if not corpus_shapes or any(not isinstance(shape, str) or not shape for shape in corpus_shapes):
        fail(f"{path}.configuration.corpus_shapes", "must be a non-empty list of names")
    for field, expected_value in (
        ("filesystem_cache_states", ["warm"]),
        ("filesystem_fresh_child_per_sample", True),
        ("filesystem_process_isolated", True),
        ("filesystem_root_selected", False),
        ("execution_workers", [1]),
    ):
        if field in configuration and configuration[field] != expected_value:
            fail(f"{path}.configuration.{field}", f"does not match {expected_value!r}")

    catalog_ref = object_value(report.get("corpus_catalog"), f"{path}.corpus_catalog")
    expected_catalog_ref = {
        key: catalog.get(key)
        for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256")
    }
    if catalog_ref != expected_catalog_ref:
        fail(f"{path}.corpus_catalog", "does not match the catalog sidecar")
    try:
        corpus_binding.validate_binding(report, catalog)
    except Exception as error:
        fail(path, f"report/catalog binding rejected: {error}")
    catalog_build = object_value(catalog.get("build"), f"{path}.catalog.build")
    if catalog_build.get("git_revision") != report["environment"].get("git_revision"):
        fail(f"{path}.catalog.build.git_revision", "does not match report environment revision")

    try:
        indexed = perf_abba_summary._index_results(report, path)
        perf_abba_summary._validate_pptx_cross_copy_result_rows(indexed, configuration, path)
        perf_abba_summary._validate_configuration_rows(configuration, indexed, path)
    except Exception as error:
        fail(path, f"shared PPTX row validator rejected report: {error}")
    results = list_value(report.get("results"), f"{path}.results")
    if len(results) != 1:
        fail(f"{path}.results", "must contain exactly one result")
    result = object_value(results[0], f"{path}.results[0]")
    if result.get("case") != selector:
        fail(f"{path}.results[0].case", "does not match the selected case")
    elapsed_samples, sample_order, elapsed_stats = validate_elapsed(
        result.get("elapsed_ns"),
        f"{path}.results[0].elapsed_ns",
        samples,
        perf_abba_summary,
    )
    operation, allocation_status = validate_operation(
        result,
        f"{path}.results[0]",
        selector,
        lane,
        perf_compare,
        sample_order,
    )

    source = object_value(result.get("source"), f"{path}.results[0].source")
    for field in (
        "read_calls",
        "read_bytes",
        "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes",
        "max_in_flight_reads",
    ):
        values = list_value(source.get(field), f"{path}.results[0].source.{field}")
        for index, value in enumerate(values):
            integer_value(value, f"{path}.results[0].source.{field}[{index}]")
    special = object_value(source.get("pptx_cross_copy"), f"{path}.results[0].source.pptx_cross_copy")
    expected_output = sha256_value(
        special.get("expected_output_sha256"),
        f"{path}.results[0].source.pptx_cross_copy.expected_output_sha256",
    )
    output = sha256_value(result.get("output_sha256"), f"{path}.results[0].output_sha256")
    if output != expected_output:
        fail(f"{path}.results[0].output_sha256", "does not match expected_output_sha256")
    output_vector = list_value(
        special.get("output_sha256"),
        f"{path}.results[0].source.pptx_cross_copy.output_sha256",
    )
    if len(output_vector) != samples:
        fail(f"{path}.results[0].source.pptx_cross_copy.output_sha256", "does not match sample count")
    output_vector = [
        sha256_value(value, f"{path}.results[0].source.pptx_cross_copy.output_sha256[{index}]")
        for index, value in enumerate(output_vector)
    ]
    if output_vector != [expected_output] * samples:
        fail(f"{path}.results[0].source.pptx_cross_copy.output_sha256", "contains an unexpected digest")
    gates = object_value(special.get("gates"), f"{path}.results[0].source.pptx_cross_copy.gates")
    if not gates or any(value is not True for value in gates.values()):
        fail(f"{path}.results[0].source.pptx_cross_copy.gates", "every correctness gate must be true")

    return {
        "path": path,
        "raw_sha256": raw_sha256,
        "report": report,
        "catalog": catalog,
        "result": result,
        "binary": binary,
        "revision": string_value(report["environment"].get("git_revision"), f"{path}.environment.git_revision"),
        "corpus": result["corpus"],
        "source_special": special,
        "output_sha256": output,
        "output_vector_sha256": digest_json(output_vector, f"{path}.output_vector_sha256"),
        "elapsed_samples": elapsed_samples,
        "sample_order": sample_order,
        "elapsed_stats": elapsed_stats,
        "operation": operation,
        "allocation_status": allocation_status,
    }


def semantic_special(value: Mapping[str, Any]) -> dict[str, Any]:
    """Retain correctness/source identity while dropping run/implementation detail."""

    return {
        key: child
        for key, child in value.items()
        if key not in _PPTX_DYNAMIC_FIELDS and key not in _PPTX_DESCRIPTIVE_FIELDS
    }


def stable_environment(value: Mapping[str, Any]) -> dict[str, Any]:
    return {key: child for key, child in value.items() if key != "git_revision"}


def validate_abba(rows: list[dict[str, Any]], selector: str) -> None:
    if len(rows) != 4:
        return
    if [row["report"]["results"][0]["case"] for row in rows] != [selector] * 4:
        fail("abba", "all four rows must select the requested case")
    for left, right in ((0, 3), (1, 2)):
        left_binary = rows[left]["binary"]
        right_binary = rows[right]["binary"]
        for field in ("binary_sha256", "binary_bytes", "mode_bits", "executable", "profile"):
            if left_binary.get(field) != right_binary.get(field):
                fail(f"abba.{ABBA_LEGS[left]}->{ABBA_LEGS[right]}", f"binary {field} differs")
    if rows[0]["binary"]["binary_sha256"] == rows[1]["binary"]["binary_sha256"]:
        fail("abba", "control and candidate binaries must have distinct identities")
    if rows[0]["revision"] != rows[3]["revision"] or rows[1]["revision"] != rows[2]["revision"]:
        fail("abba", "same-role legs must retain the same source revision")
    if rows[0]["revision"] == rows[1]["revision"]:
        fail("abba", "control and candidate source revisions must differ")

    reference = rows[0]
    stable_source = canonical(semantic_special(reference["source_special"]), "abba.source")
    stable_corpus = canonical(reference["corpus"], "abba.corpus")
    stable_environment_value = canonical(stable_environment(reference["report"]["environment"]), "abba.environment")
    for index, row in enumerate(rows[1:], 1):
        label = f"abba.{ABBA_LEGS[index]}"
        if canonical(semantic_special(row["source_special"]), label) != stable_source:
            fail(label, "source/output correctness identity differs across ABBA legs")
        if canonical(row["corpus"], label) != stable_corpus:
            fail(label, "corpus identity differs across ABBA legs")
        if row["output_sha256"] != reference["output_sha256"]:
            fail(label, "output digest differs across ABBA legs")
        if canonical(stable_environment(row["report"]["environment"]), label) != stable_environment_value:
            fail(label, "non-revision environment identity differs across ABBA legs")
        if row["report"]["corpus_catalog"].get("content_set_sha256") != reference["report"]["corpus_catalog"].get("content_set_sha256"):
            fail(label, "catalog content-set identity differs across ABBA legs")


def validate_capture_manifest(
    capture_path: Path,
    report_path: Path,
    catalog_path: Path,
) -> dict[str, Any]:
    capture, raw_sha256 = load_json(capture_path, "capture")
    if capture.get("status") != "pass":
        fail("capture.status", "must be 'pass'")
    if capture.get("exit_code") != 0:
        fail("capture.exit_code", "must be zero")
    if capture.get("source_unchanged") is not True or capture.get("binary_unchanged") is not True:
        fail("capture", "source and binary identities must remain unchanged")
    files = list_value(capture.get("files"), "capture.files")
    names: set[str] = set()
    capture_dir = capture_path.parent.resolve()
    for index, item in enumerate(files):
        entry = object_value(item, f"capture.files[{index}]")
        name = string_value(entry.get("name"), f"capture.files[{index}].name")
        if name in names or Path(name).name != name or Path(name).is_absolute():
            fail(f"capture.files[{index}].name", "must be a unique basename")
        names.add(name)
        artifact = capture_dir / name
        expected_bytes = integer_value(entry.get("bytes"), f"capture.files[{index}].bytes")
        if expected_bytes > MAX_JSON_BYTES:
            fail(f"capture.files[{index}].bytes", "exceeds the bounded artifact limit")
        compressed = False
        if not artifact.is_file():
            # Measurement cleanup may gzip only retained text logs.  Keep this
            # fallback at the capture-manifest boundary: report and catalog
            # paths remain plain JSON inputs to load_json().
            if not name.endswith((".log", ".txt")):
                fail(f"capture.files[{index}]", "listed artifact is missing")
            artifact = capture_dir / f"{name}.gz"
            if not artifact.is_file():
                fail(f"capture.files[{index}]", "listed artifact is missing")
            compressed = True
        if not compressed and artifact.stat().st_size != expected_bytes:
            fail(f"capture.files[{index}].bytes", "does not match artifact")
        digest_builder = hashlib.sha256()
        observed_bytes = 0
        try:
            stream = gzip.open(artifact, "rb") if compressed else artifact.open("rb")
            with stream:
                while True:
                    # For a compressed sidecar, consume at most expected bytes
                    # plus one sentinel byte.  This detects a changed raw
                    # length without allowing a decompression bomb to expand
                    # beyond the verifier's artifact budget.
                    remaining = expected_bytes + 1 - observed_bytes
                    if remaining <= 0:
                        break
                    chunk = stream.read(min(1024 * 1024, remaining))
                    if not chunk:
                        break
                    observed_bytes += len(chunk)
                    if observed_bytes > expected_bytes:
                        fail(f"capture.files[{index}]", "compressed raw bytes exceed recorded length")
                    digest_builder.update(chunk)
        except (gzip.BadGzipFile, EOFError, OSError) as error:
            fail(f"capture.files[{index}]", f"cannot read compressed sidecar: {error}")
        if observed_bytes != expected_bytes:
            fail(f"capture.files[{index}].bytes", "does not match raw artifact length")
        digest = digest_builder.hexdigest()
        if sha256_value(entry.get("sha256"), f"capture.files[{index}].sha256") != digest:
            fail(f"capture.files[{index}].sha256", "does not match raw artifact bytes")
    for artifact in (report_path.resolve(), catalog_path.resolve()):
        if artifact.parent != capture_dir or artifact.name not in names:
            fail("capture.files", f"does not retain {artifact.name!r} beside capture manifest")
    return {"path": str(capture_path), "raw_sha256": raw_sha256, "artifact_count": len(files)}


def allocation_projection(operation: Mapping[str, Any]) -> dict[str, Any] | None:
    allocation = operation.get("allocation")
    if not isinstance(allocation, dict):
        return None
    result: dict[str, Any] = {"status": allocation.get("status")}
    for field in (
        "allocation_calls",
        "deallocation_calls",
        "reallocation_calls",
        "failed_allocation_calls",
        "allocated_bytes",
        "deallocated_bytes",
    ):
        metric = allocation.get(field)
        if not isinstance(metric, dict) or not isinstance(metric.get("values"), list):
            continue
        values = metric["values"]
        if all(isinstance(value, int) and not isinstance(value, bool) and value >= 0 for value in values):
            result[field] = {
                "count": len(values),
                "sum": sum(values),
                "min": min(values) if values else None,
                "max": max(values) if values else None,
            }
    return result


def make_summary(
    rows: list[dict[str, Any]],
    *,
    selector: str,
    lane: str,
    samples: int,
    warmups: int,
    capture: dict[str, Any] | None,
) -> dict[str, Any]:
    report_rows = []
    for index, row in enumerate(rows):
        report_rows.append(
            {
                "leg": ABBA_LEGS[index] if len(rows) == 4 else "single",
                "revision": row["revision"],
                "binary_sha256": row["binary"]["binary_sha256"],
                "report_sha256": row["raw_sha256"],
                "corpus_archive_sha256": row["corpus"].get("archive_sha256"),
                "output_sha256": row["output_sha256"],
                "output_vector_sha256": row["output_vector_sha256"],
                "elapsed": row["elapsed_stats"],
                "sample_order_sha256": digest_json(row["sample_order"], "sample_order"),
                "allocation": allocation_projection(row["operation"]),
            }
        )
    if lane != "normal":
        claim_reason = "allocator instrumentation is excluded from latency claims"
    elif len(rows) != 4:
        claim_reason = "a four-leg ABBA comparison is required"
    elif samples < NORMAL_CLAIM_MIN_SAMPLES:
        claim_reason = f"normal ABBA retains {samples} samples; at least {NORMAL_CLAIM_MIN_SAMPLES} are required"
    else:
        claim_reason = "0419 is an allocation diagnostic and registers no latency claim"
    return {
        "change": EXPECTED_CHANGE,
        "classification": "matched allocation/resource diagnostic",
        "selector": selector,
        "lane": lane,
        "samples": samples,
        "warmups": warmups,
        "report_count": len(rows),
        "claim_authorized": False,
        "performance_claim": None,
        "claim_withheld_reason": claim_reason,
        "capture": capture,
        "reports": report_rows,
    }


def flatten_paths(values: Iterable[list[str] | None]) -> list[Path]:
    result: list[Path] = []
    for group in values:
        if group:
            result.extend(Path(value).expanduser() for value in group)
    return result


def resolve_inputs(args: argparse.Namespace) -> tuple[list[Path], list[Path], Path | None]:
    reports = flatten_paths((args.report_values, *args.reports_values))
    catalogs = flatten_paths((args.catalog_values, *args.catalogs_values))
    root = args.root.expanduser().resolve() if args.root is not None else None
    if not reports:
        if root is None or not args.selector:
            fail("inputs", "pass reports or --root with --selector")
        reports = [
            root / "runs" / args.lane / leg / args.selector / "report.json"
            for leg in ABBA_LEGS
        ]
    if not catalogs:
        catalogs = [report.parent / "catalog.json" for report in reports]
    if len(catalogs) == 1 and len(reports) > 1:
        catalogs = catalogs * len(reports)
    if len(catalogs) != len(reports):
        fail("inputs", "the number of catalogs must match reports (or be one shared catalog)")
    if len(reports) not in (1, 4):
        fail("inputs", "pass exactly one report or four ABBA reports")
    return reports, catalogs, root


def run(args: argparse.Namespace) -> dict[str, Any]:
    repo_root = find_repo_root(args.repo_root)
    perf_abba_summary, perf_compare, corpus_binding = import_validators(repo_root)
    reports, catalogs, root = resolve_inputs(args)
    if args.samples is None:
        samples = 100 if args.lane == "normal" else 30
    else:
        samples = args.samples
    if args.warmups is None:
        warmups = 10 if args.lane == "normal" else 3
    else:
        warmups = args.warmups
    if samples < 1 or warmups < 0:
        fail("spec", "samples must be positive and warmups non-negative")

    loaded_rows: list[dict[str, Any]] = []
    selected_selector = args.selector
    for index, (report_path, catalog_path) in enumerate(zip(reports, catalogs)):
        report, report_sha256 = load_json(report_path, f"report[{index}]")
        catalog, _catalog_sha256 = load_json(catalog_path, f"catalog[{index}]")
        if selected_selector is None:
            result_rows = list_value(report.get("results"), f"report[{index}].results")
            if len(result_rows) != 1:
                fail("selector", "must be supplied when a report has multiple results")
            selected_selector = string_value(
                object_value(result_rows[0], f"report[{index}].results[0]").get("case"),
                f"report[{index}].results[0].case",
            )
        loaded_rows.append(
            validate_report(
                report,
                catalog,
                f"report[{index}]",
                selector=selected_selector,
                samples=samples,
                warmups=warmups,
                lane=args.lane,
                perf_abba_summary=perf_abba_summary,
                perf_compare=perf_compare,
                corpus_binding=corpus_binding,
                raw_sha256=report_sha256,
            )
        )
    assert selected_selector is not None
    validate_abba(loaded_rows, selected_selector)

    capture = None
    if args.capture is not None:
        if len(reports) != 1:
            fail("--capture", "is supported for a single report invocation")
        capture = validate_capture_manifest(args.capture, reports[0], catalogs[0])
    if args.trace is not None:
        if not args.trace.is_file() or args.trace.stat().st_size <= 0:
            fail("--trace", "must be a non-empty retained artifact")
    summary = make_summary(
        loaded_rows,
        selector=selected_selector,
        lane=args.lane,
        samples=samples,
        warmups=warmups,
        capture=capture,
    )
    if args.output is not None:
        output = args.output.expanduser()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(summary, sort_keys=True, indent=2) + "\n", encoding="utf-8")
    return summary


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--root", type=Path, help="capture root containing runs/{mode}/{leg}/{selector}")
    parser.add_argument("--mode", "--lane", dest="lane", choices=("normal", "allocator"), default="normal")
    parser.add_argument("--selector")
    parser.add_argument("--samples", type=int)
    parser.add_argument("--warmups", type=int)
    parser.add_argument("--report", dest="report_values", action="append")
    parser.add_argument("--reports", dest="reports_values", action="append", nargs="+")
    parser.add_argument("--catalog", dest="catalog_values", action="append")
    parser.add_argument("--catalogs", dest="catalogs_values", action="append", nargs="+")
    parser.add_argument("--capture", type=Path)
    parser.add_argument("--trace", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    args.report_values = args.report_values or []
    args.reports_values = args.reports_values or []
    args.catalog_values = args.catalog_values or []
    args.catalogs_values = args.catalogs_values or []
    try:
        summary = run(args)
    except (VerificationError, OSError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 2
    print(json.dumps(summary, sort_keys=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
