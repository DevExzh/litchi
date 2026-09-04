from __future__ import annotations

import copy
import contextlib
import io
import json
import math
import tempfile
import unittest
from pathlib import Path

from tools.validate_opc_serial_eager_allocator_abba import (
    ALLOCATOR,
    ALLOCATOR_BINARY,
    ALLOCATOR_SCOPE,
    ABBA_ORDER,
    CASE,
    CORPUS_ORACLE,
    COUNTER_METRICS,
    EXPECTED_DELTA_BY_SHAPE,
    HWM_SCOPE,
    NONCLAIMABLE_ALLOCATOR_METRICS,
    PROCESS_SCOPE,
    RSS_SCOPE,
    SAMPLE_COUNT,
    SHAPES,
    SOURCE_COMPRESSED_SCOPE,
    SOURCE_DECOMPRESSED_SCOPE,
    SOURCE_PATTERN_SCOPE,
    SOURCE_RECOMPRESSED_SCOPE,
    SOURCE_SCOPE,
    ValidationError,
    _expected_configuration,
    _validate_elapsed,
    main,
    validate_paths,
)


ROLES = ("A1", "B1", "B2", "A2")
ROLE_OFFSETS = {"A1": 100, "B1": 200, "B2": 300, "A2": 400}
CONTROL_REVISION = "cca80d89bac0aa4e2740a7879cf39cdcd8cbbb44"
CANDIDATE_REVISION = "fadf43722289fc78f565b8265a03d4763d2660b5"

# Independent producer-manifest cross-check.  Do not derive this table from
# CORPUS_ORACLE: this test must fail if the validator and its fixture drift
# together, as happened when the incompressible payload hashes were first
# transcribed.
INDEPENDENT_CORPUS_IDENTITIES = {
    "tiny": {
        "archive_bytes": 1_310,
        "archive_member_count": 5,
        "entry_count": 3,
        "entry_bytes": 512,
        "target_payload_sha256": "630b1da45fe604eda3b5468b7c9ca7facfbd404941779786276a69ff870e4bdd",
        "part_names_sha256": "5458f5d1eb9283e10cd7057abf8f63cce9d1e0b6c57c5f9f945a9bad3b99cda4",
        "part_payload_sha256": "d1baa4a40fc63856136504f95933bcb2bb3da28f2000cabe1153eaee88c723c0",
    },
    "many-small": {
        "archive_bytes": 303_003,
        "archive_member_count": 258,
        "entry_count": 256,
        "entry_bytes": 1_024,
        "target_payload_sha256": "05fd26cad1f538b7ed415a0f525a13896823b02abcf22ad1746172f035a2149d",
        "part_names_sha256": "82415ca7ad25155c41df5d93707c95e5fcc31e66cde226ff046fc84906f56bc2",
        "part_payload_sha256": "7bdf372948a4f914aea31187d1f2813254957cd907279690b022ef00737caaa7",
    },
    "few-large": {
        "archive_bytes": 16_783_632,
        "archive_member_count": 6,
        "entry_count": 4,
        "entry_bytes": 4 * 1024 * 1024,
        "target_payload_sha256": "3dbf6225021a99c1da8750a738bde21f57591c0be1a60aa510966c47ee25b098",
        "part_names_sha256": "d48e27d95e97a4de43e476096910540416f6e19eb54a3759d5ca081b4136166c",
        "part_payload_sha256": "ac1e942c87db2e622c1e1c2efd1046e5d791a44db73bd6255078f8816d922db3",
    },
}


def _digest(digit: str) -> str:
    return (digit * 64)[:64]


def _revision(digit: str) -> str:
    return (digit * 40)[:40]


def _metric(status: str, scope: str, value: object = None, *, reason: str | None = None) -> dict[str, object]:
    result: dict[str, object] = {"status": status}
    if value is not None:
        result["value"] = value
    result["scope"] = scope
    if reason is not None:
        result["reason"] = reason
    return result


def _absent(status: str, scope: str) -> dict[str, object]:
    return {"status": status, "scope": scope}


def _elapsed(offset: int) -> dict[str, object]:
    samples = [offset + index for index in range(1, SAMPLE_COUNT + 1)]
    mean = 0.0
    squared = 0.0
    for index, sample in enumerate(samples):
        current = float(sample)
        count = float(index + 1)
        delta = current - mean
        next_mean = mean + delta / count
        squared += delta * (current - next_mean)
        mean = next_mean
    standard_deviation = math.sqrt(squared / (SAMPLE_COUNT - 1))
    margin = 2.145 * standard_deviation / math.sqrt(SAMPLE_COUNT)
    return {
        "unit": "ns",
        "samples": samples,
        "sample_order": list(reversed(range(SAMPLE_COUNT))),
        "min": samples[0],
        "p50": samples[7],
        "p95": samples[-1],
        "p99": samples[-1],
        "max": samples[-1],
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(mean - margin, 0.0),
            "upper": mean + margin,
        },
    }


def _contract() -> dict[str, object]:
    control = {
        "implementation": "control",
        "revision": CONTROL_REVISION,
        "binary_sha256": _digest("1"),
        "binary_bytes": 1001,
        "mode_bits": 509,
        "profile": "release",
    }
    candidate = {
        "implementation": "candidate",
        "revision": CANDIDATE_REVISION,
        "binary_sha256": _digest("2"),
        "binary_bytes": 1002,
        "mode_bits": 509,
        "profile": "release",
    }
    return {
        "schema_version": 1,
        "case": CASE,
        "cache_state": "warm",
        "samples_per_case": SAMPLE_COUNT,
        "warmup_iterations_per_case": 3,
        "execution_workers": [1],
        "abba_order": ABBA_ORDER,
        "tool": {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": ALLOCATOR_BINARY,
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": "system_allocator_operation_scoped",
        },
        "environment": {
            "rustc_version": "rustc 1.98.1 (test)",
            "allocator": ALLOCATOR,
            "target_os": "linux",
            "target_arch": "x86_64",
            "logical_cpus_available": 32,
            "cpu_model": "test-cpu",
            "cpu_affinity": "2",
            "rustflags": None,
            "cargo_build_target": None,
        },
        "legs": {"A1": control, "A2": copy.deepcopy(control), "B1": candidate, "B2": copy.deepcopy(candidate)},
        "corpora": [copy.deepcopy(CORPUS_ORACLE[shape]) for shape in SHAPES],
        "expected_deltas": copy.deepcopy(EXPECTED_DELTA_BY_SHAPE),
    }


def _operation(role: str, shape: str) -> dict[str, object]:
    source: dict[str, object] = {
        "status": "not_applicable",
        "counter_scope": "not_applicable_in_process_sink",
    }
    for key in (
        "logical_read_calls",
        "logical_read_requested_bytes",
        "logical_read_returned_bytes",
        "logical_read_largest_requested_bytes",
        "logical_read_largest_returned_bytes",
        "max_concurrent_reads",
    ):
        source[key] = _absent("not_applicable", SOURCE_SCOPE)
    source["logical_read_pattern"] = _absent("not_applicable", SOURCE_PATTERN_SCOPE)
    source["compressed_bytes"] = _absent("not_applicable", SOURCE_COMPRESSED_SCOPE)
    source["decompressed_bytes"] = _absent("not_applicable", SOURCE_DECOMPRESSED_SCOPE)
    source["recompressed_bytes"] = _absent("not_applicable", SOURCE_RECOMPRESSED_SCOPE)

    process: dict[str, object] = {"status": "unavailable"}
    for key in (
        "user_cpu_ticks",
        "system_cpu_ticks",
        "clock_ticks_per_second",
        "minor_faults",
        "major_faults",
        "voluntary_context_switches",
        "nonvoluntary_context_switches",
        "rchar",
        "wchar",
        "read_bytes",
        "write_bytes",
        "cancelled_write_bytes",
        "syscr",
        "syscw",
    ):
        process[key] = _absent("unavailable", PROCESS_SCOPE)
    process["rss_delta_bytes"] = _absent("unavailable", RSS_SCOPE)
    process["peak_rss_bytes"] = _absent("unavailable", HWM_SCOPE)

    sink = {
        "status": "not_applicable",
        "output_bytes": _absent("not_applicable", "post_operation_output_length_not_sink_write_volume"),
        "write_status": "not_applicable",
        "accepted_bytes": _absent("not_applicable", "logical_sink_accepted_write_bytes"),
        "write_calls": _absent("not_applicable", "logical_sink_accepted_write_calls"),
        "largest_write": _absent("not_applicable", "logical_sink_largest_accepted_write"),
        "write_size_buckets": {"status": "not_applicable"},
    }
    for key in (
        "bytes_0",
        "bytes_1_to_512",
        "bytes_513_to_4096",
        "bytes_4097_to_16384",
        "bytes_16385_to_65536",
        "bytes_over_65536",
    ):
        sink["write_size_buckets"][key] = _absent("not_applicable", "logical_sink_accepted_write_size_bucket_counts")

    publication = {
        "status": "not_applicable",
        "changed_spans": _absent("not_applicable", "logical_publication_counter"),
        "published_bytes": _absent("not_applicable", "logical_publication_counter"),
    }
    materialization = {"status": "not_applicable", "opc_parts": _absent("not_applicable", "logical_materialization_counter")}
    cfb = {"status": "not_applicable"}
    for phase in ("open", "plan", "atomic_publication"):
        cfb[phase] = {
            "elapsed_ns": _absent("not_applicable", "timed_cfb_phase_elapsed_ns"),
            "logical_read_calls": _absent("not_applicable", "timed_cfb_phase_logical_read_at"),
            "logical_read_requested_bytes": _absent("not_applicable", "timed_cfb_phase_logical_read_at"),
            "logical_read_returned_bytes": _absent("not_applicable", "timed_cfb_phase_logical_read_at"),
        }

    shape_index = SHAPES.index(shape)
    values = {
        "allocation_calls": 1000 + shape_index * 100,
        "deallocation_calls": 1100 + shape_index * 100,
        "reallocation_calls": 7 + shape_index,
        "failed_allocation_calls": 0,
        "allocated_bytes": 100_000_000 + shape_index * 100_000,
        "deallocated_bytes": 90_000_000 + shape_index * 100_000,
    }
    if role in ("B1", "B2"):
        for metric, delta in EXPECTED_DELTA_BY_SHAPE[shape].items():
            values[metric] += delta
    allocation = {"status": "measured", "scope": ALLOCATOR_SCOPE}
    for metric in COUNTER_METRICS:
        allocation[metric] = {"values": [values[metric]] * SAMPLE_COUNT, "status": "measured", "scope": ALLOCATOR_SCOPE}
    for index, metric in enumerate(NONCLAIMABLE_ALLOCATOR_METRICS):
        allocation[metric] = {"values": [10_000 + index + (role in ("B1", "B2"))] * SAMPLE_COUNT, "status": "measured", "scope": ALLOCATOR_SCOPE}
    return {
        "sample_count": SAMPLE_COUNT,
        "sample_indices": list(reversed(range(SAMPLE_COUNT))),
        "alignment": "elapsed_ns.samples_by_elapsed_then_sample_index",
        "latency_claim": "comparable_timed_operation",
        "source": source,
        "process": process,
        "sink": sink,
        "publication": publication,
        "materialization": materialization,
        "cfb_phases": cfb,
        "allocation": allocation,
    }


def _summary(shape: str) -> dict[str, object]:
    oracle = CORPUS_ORACLE[shape]
    model = {"comparison": "candidate-control", "status": "expected_not_observed"}
    model.update(EXPECTED_DELTA_BY_SHAPE[shape])
    return {
        "implementation": "OpcPackage::from_bytes",
        "timing_scope": "OpcPackage::from_bytes constructor only; ZIP preflight and all package semantic oracles excluded",
        "performance_claim": "none",
        "predeclared_allocator_model": model,
        "worker_count": 1,
        "source_archive_bytes": oracle["archive_bytes"],
        "source_archive_sha256": oracle["archive_sha256"],
        "archive_member_count": oracle["archive_member_count"],
        "part_count": oracle["part_count"],
        "part_names_sha256": oracle["part_names_sha256"],
        "part_payload_sha256": oracle["part_payload_sha256"],
        "target_name": oracle["target_entry"],
        "target_payload_sha256": oracle["target_payload_sha256"],
        "all_ordinary_parts_deflated_verified": True,
        "observed_part_counts": [oracle["part_count"]] * SAMPLE_COUNT,
        "observed_part_names_sha256": [oracle["part_names_sha256"]] * SAMPLE_COUNT,
        "observed_part_payload_sha256": [oracle["part_payload_sha256"]] * SAMPLE_COUNT,
        "observed_content_types_verified": [True] * SAMPLE_COUNT,
        "observed_root_relationship_verified": [True] * SAMPLE_COUNT,
        "observed_main_target_verified": [True] * SAMPLE_COUNT,
        "observed_deterministic_payload_hashes_verified": [True] * SAMPLE_COUNT,
    }


def _report(role: str) -> dict[str, object]:
    contract = _contract()
    leg = contract["legs"][role]
    environment = {
        "rustc_version": contract["environment"]["rustc_version"],
        "git_revision": leg["revision"],
        "git_worktree_dirty": False,
        "logical_cpus_available": 32,
        "allocator": ALLOCATOR,
        "rustflags": None,
        "cargo_build_target": None,
        "perf_event_paranoid": "1",
        "os": "linux",
        "kernel": "Linux test",
        "cpu_model": "test-cpu",
        "total_memory_bytes": 1_073_741_824,
        "page_size_bytes": 4096,
        "filesystem_type": "tmpfs",
        "source_destination_same_device": True,
        "cpu_affinity": "2",
        "storage_identifier": None,
    }
    results = []
    parallel_cases = []
    for shape in SHAPES:
        oracle = CORPUS_ORACLE[shape]
        corpus = {key: value for key, value in oracle.items() if key not in {"part_count", "part_names_sha256", "part_payload_sha256"}}
        corpus.update({"xlsx": None})
        source = {
            "read_calls": [],
            "read_bytes": [],
            "ordinary_payload_read_calls": [],
            "ordinary_payload_read_bytes": [],
            "max_in_flight_reads": [],
            "opc_serial_eager_open": _summary(shape),
        }
        results.append({
            "case": CASE,
            "corpus": corpus,
            "elapsed_ns": _elapsed(ROLE_OFFSETS[role]),
            "sink": None,
            "source": source,
            "execution": {"worker_count": 1, "logical_tasks": 1, "logical_bytes": oracle["archive_bytes"]},
            "output_sha256": oracle["archive_sha256"],
            "operation_metrics": _operation(role, shape),
        })
        parallel_cases.append({
            "case": CASE,
            "corpus_sha256": oracle["archive_sha256"],
            "configured_worker_count": _metric("measured", "result.execution.worker_count", 1),
            "observed_local_worker_count": _metric("not_applicable", "result.source.opc_cache.worker_count_with_one_created_local_worker_team", reason="serial result created no local worker team"),
            "deterministic_task_count": _metric("measured", "result.execution.logical_tasks", 1),
            "deterministic_chunk_count": _metric("unavailable", "result.execution.deterministic_chunk_count", reason="no deterministic chunk counter is exposed; byte totals are not used as a proxy"),
            "lock_wait_ns": _metric("unavailable", "lock_wait_ns", reason="no exact instrumented lock boundary is present; waiter counts are not timed"),
        })
    return {
        "schema_version": 1,
        "tool": {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": ALLOCATOR_BINARY,
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": "system_allocator_operation_scoped",
        },
        "binary_identity": {
            "path": f"/tmp/{role.lower()}/{ALLOCATOR_BINARY}",
            "binary_sha256": leg["binary_sha256"],
            "binary_bytes": leg["binary_bytes"],
            "mode_bits": leg["mode_bits"],
            "executable": True,
            "profile": "release",
        },
        "environment": environment,
        "configuration": copy.deepcopy(_expected_configuration()),
        "parallel_metrics": {
            "schema_version": 1,
            "scope": "explicit_local_execution_only",
            "claim": "descriptive",
            "configured_worker_budget": _metric("measured", "configuration.execution_workers", [1]),
            "observed_process_thread_count": _metric("unavailable", "process_thread_count", reason="no process-global thread counter is collected; local worker boundaries only"),
            "cases": parallel_cases,
        },
        "results": results,
    }


class SerialEagerAllocatorValidatorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.contract = _contract()
        self.reports = {role: _report(role) for role in ROLES}

    def _validate(self, *, reports: dict[str, object] | None = None, contract: dict[str, object] | None = None) -> dict[str, object]:
        reports = self.reports if reports is None else reports
        contract = self.contract if contract is None else contract
        with tempfile.TemporaryDirectory(prefix="litchi-0403-serial-eager-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(contract, indent=2) + chr(10), encoding="utf-8")
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(reports[role], indent=2) + chr(10), encoding="utf-8")
                paths[role] = path
            return validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)

    def test_fixed_hashes_match_independent_producer_manifest(self) -> None:
        for shape, expected in INDEPENDENT_CORPUS_IDENTITIES.items():
            with self.subTest(shape=shape):
                for key, value in expected.items():
                    self.assertEqual(CORPUS_ORACLE[shape][key], value)

    def test_rejects_unstable_elapsed_ties(self) -> None:
        elapsed = _elapsed(100)
        elapsed["samples"] = [101] * SAMPLE_COUNT
        elapsed["sample_order"] = list(reversed(range(SAMPLE_COUNT)))
        elapsed["min"] = 101
        elapsed["p50"] = 101
        elapsed["p95"] = 101
        elapsed["p99"] = 101
        elapsed["max"] = 101
        elapsed["mean"] = 101.0
        elapsed["standard_deviation"] = 0.0
        elapsed["confidence_interval_95"]["lower"] = 101.0
        elapsed["confidence_interval_95"]["upper"] = 101.0
        with self.assertRaisesRegex(ValidationError, "stable by"):
            _validate_elapsed(elapsed, "tie-test")

    def test_accepts_exact_in_process_allocator_abba_and_separates_model(self) -> None:
        projection = self._validate()
        self.assertEqual(projection["validation"]["report_count"], 4)
        self.assertEqual(projection["validation"]["matrix_rows"], 3)
        self.assertTrue(projection["validation"]["exact_candidate_control_deltas_verified"])
        self.assertTrue(projection["validation"]["model_is_expected_not_observed"])
        self.assertFalse(projection["claimability"]["allocator_elapsed_ns"]["claimable"])
        self.assertFalse(projection["claimability"]["fresh_child_per_sample"]["claimable"])
        self.assertFalse(projection["claimability"]["logical_io"]["claimable"])
        self.assertFalse(projection["claimability"]["physical_io"]["claimable"])
        self.assertFalse(projection["provenance"]["binary_identity"]["file_rehashed"])
        self.assertEqual(projection["rows"][0]["predeclared_allocator_model"]["status"], "expected_not_observed")
        self.assertEqual(projection["model_vs_observation"]["observed_allocator_vectors"]["status"], "measured")

    def test_rejects_duplicate_keys_and_nonfinite_numbers(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0403-serial-eager-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                encoded = json.dumps(self.reports[role])
                if role == "A1":
                    encoded = encoded.replace('{"schema_version": 1,', '{"schema_version": 1, "schema_version": 1,', 1)
                path.write_text(encoded, encoding="utf-8")
                paths[role] = path
            with self.assertRaisesRegex(ValidationError, "duplicate JSON object key"):
                validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)
            paths["A1"].write_text(json.dumps(self.reports["A1"]).replace('"mean": 108.0', '"mean": NaN'), encoding="utf-8")
            with self.assertRaisesRegex(ValidationError, "non-finite JSON number"):
                validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)

    def test_rejects_swapped_roles_dirty_or_mutated_binary(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0403-serial-eager-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(self.reports[role]), encoding="utf-8")
                paths[role] = path
            with self.assertRaisesRegex(ValidationError, "binary_sha256"):
                validate_paths(paths["B1"], paths["A1"], paths["B2"], paths["A2"], contract_path)
        reports = copy.deepcopy(self.reports)
        reports["A2"]["environment"]["git_worktree_dirty"] = True
        with self.assertRaisesRegex(ValidationError, "git_worktree_dirty"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["B2"]["binary_identity"]["binary_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValidationError, "binary_sha256"):
            self._validate(reports=reports)

    def test_rejects_toolchain_cpu_workers_and_cold_or_fresh_child_configuration(self) -> None:
        for path, key, value, message in (
            ("environment", "rustc_version", "rustc 1.95.0", "rustc_version"),
            ("environment", "cpu_affinity", "3", "cpu_affinity"),
            ("configuration", "execution_workers", [1, 2], "execution_workers"),
            ("configuration", "filesystem_cache_states", ["warm", "cold-requested"], "filesystem_cache_states"),
            ("configuration", "filesystem_fresh_child_per_sample", False, "filesystem_fresh_child_per_sample"),
        ):
            reports = copy.deepcopy(self.reports)
            reports["B1"][path][key] = value
            with self.assertRaisesRegex(ValidationError, message):
                self._validate(reports=reports)

    def test_rejects_matrix_corpus_member_part_and_semantic_mutations(self) -> None:
        mutations = (
            (lambda report: report["results"].pop(), "fixed rows"),
            (lambda report: report["results"][0]["corpus"].__setitem__("archive_sha256", "0" * 64), "archive_sha256"),
            (lambda report: report["results"][1]["corpus"].__setitem__("archive_member_count", 257), "archive_member_count"),
            (lambda report: report["results"][2]["source"]["opc_serial_eager_open"].__setitem__("part_count", 5), "part_count"),
            (lambda report: report["results"][0]["source"]["opc_serial_eager_open"].__setitem__("all_ordinary_parts_deflated_verified", False), "all_ordinary_parts_deflated_verified"),
            (lambda report: report["results"][0]["source"]["opc_serial_eager_open"]["observed_content_types_verified"].__setitem__(0, False), "observed_content_types_verified"),
        )
        for mutation, message in mutations:
            reports = copy.deepcopy(self.reports)
            mutation(reports["A1"])
            with self.assertRaisesRegex(ValidationError, message):
                self._validate(reports=reports)

    def test_rejects_sample_order_vector_cardinality_and_overflow(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["operation_metrics"]["sample_indices"][0] = 0
        with self.assertRaisesRegex(ValidationError, "sample_indices"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["B1"]["results"][0]["operation_metrics"]["allocation"]["allocated_bytes"]["values"].pop()
        with self.assertRaisesRegex(ValidationError, "allocated_bytes"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["B2"]["results"][1]["operation_metrics"]["allocation"]["allocation_calls"]["values"][0] = 1 << 64
        with self.assertRaisesRegex(ValidationError, "allocation_calls"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["elapsed_ns"]["mean"] = 10**1000
        with self.assertRaisesRegex(ValidationError, "finite number"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["source"]["opc_serial_eager_open"]["observed_part_counts"][0] = 3.0
        with self.assertRaisesRegex(ValidationError, r"observed_part_counts\[0\]"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["operation_metrics"]["allocation"]["allocated_bytes"]["values"][0] = True
        with self.assertRaisesRegex(ValidationError, r"allocated_bytes\.values\[0\]"):
            self._validate(reports=reports)

    def test_rejects_claimable_abba_equality_or_exact_delta_mutations(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A2"]["results"][0]["operation_metrics"]["allocation"]["allocation_calls"]["values"][0] += 1
        with self.assertRaisesRegex(ValidationError, "control allocation_calls"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["B2"]["results"][1]["operation_metrics"]["allocation"]["allocated_bytes"]["values"][4] += 1
        with self.assertRaisesRegex(ValidationError, "candidate allocated_bytes"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["B1"]["results"][2]["operation_metrics"]["allocation"]["deallocated_bytes"]["values"][0] += 1
        reports["B2"]["results"][2]["operation_metrics"]["allocation"]["deallocated_bytes"]["values"][0] += 1
        with self.assertRaisesRegex(ValidationError, "A1_to_B1"):
            self._validate(reports=reports)

    def test_rejects_forbidden_process_elapsed_and_model_claims(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["operation_metrics"]["process"]["status"] = "measured"
        with self.assertRaisesRegex(ValidationError, "operation_metrics.process.status"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["operation_metrics"]["latency_claim"] = "allocator_speedup"
        with self.assertRaisesRegex(ValidationError, "latency_claim"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["source"]["opc_serial_eager_open"]["predeclared_allocator_model"]["status"] = "measured"
        with self.assertRaisesRegex(ValidationError, "predeclared_allocator_model.status"):
            self._validate(reports=reports)

    def test_rejects_contract_oracle_redefinition_and_duplicate_paths(self) -> None:
        contract = copy.deepcopy(self.contract)
        contract["corpora"][0]["archive_bytes"] += 1
        with self.assertRaisesRegex(ValidationError, "retained fixed identity"):
            self._validate(contract=contract)
        contract = copy.deepcopy(self.contract)
        contract["corpora"] = list(reversed(contract["corpora"]))
        with self.assertRaisesRegex(ValidationError, "fixed shape order"):
            self._validate(contract=contract)
        with tempfile.TemporaryDirectory(prefix="litchi-0403-serial-eager-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(self.reports[role]), encoding="utf-8")
                paths[role] = path
            with self.assertRaisesRegex(ValidationError, "distinct report paths"):
                validate_paths(paths["A1"], paths["A1"], paths["B2"], paths["A2"], contract_path)

    def test_cli_accepts_explicit_identity_contract(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0403-serial-eager-validator-") as directory:
            root = Path(directory)
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(self.reports[role]), encoding="utf-8")
                paths[role] = path
            args = [
                "--a1", str(paths["A1"]), "--b1", str(paths["B1"]),
                "--b2", str(paths["B2"]), "--a2", str(paths["A2"]),
                "--control-revision", CONTROL_REVISION, "--candidate-revision", CANDIDATE_REVISION,
                "--control-binary-sha256", _digest("1"), "--candidate-binary-sha256", _digest("2"),
                "--control-binary-bytes", "1001", "--candidate-binary-bytes", "1002",
                "--control-mode-bits", "509", "--candidate-mode-bits", "509",
                "--rustc-version", "rustc 1.98.1 (test)", "--logical-cpus", "32",
                "--cpu-model", "test-cpu",
            ]
            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                status = main(args)
            self.assertEqual(status, 0)
            self.assertIn('"validator":"litchi-opc-serial-eager-allocator-abba"', stdout.getvalue())

            same_identity = list(args)
            candidate_revision_index = same_identity.index(CANDIDATE_REVISION)
            same_identity[candidate_revision_index] = CONTROL_REVISION
            stderr = io.StringIO()
            with contextlib.redirect_stderr(stderr):
                status = main(same_identity)
            self.assertEqual(status, 1)
            self.assertIn("control and candidate revisions must differ", stderr.getvalue())

    def test_output_collision_is_rejected_without_overwrite(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0403-serial-eager-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(self.reports[role]), encoding="utf-8")
                paths[role] = path
            before = contract_path.read_bytes()
            stderr = io.StringIO()
            with contextlib.redirect_stderr(stderr):
                status = main([
                    "--contract", str(contract_path),
                    "--a1", str(paths["A1"]), "--b1", str(paths["B1"]),
                    "--b2", str(paths["B2"]), "--a2", str(paths["A2"]),
                    "--output", str(contract_path),
                ])
            self.assertEqual(status, 1)
            self.assertIn("--output must differ", stderr.getvalue())
            self.assertEqual(contract_path.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
