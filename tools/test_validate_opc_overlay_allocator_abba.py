from __future__ import annotations

import copy
import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path

from tools.validate_opc_overlay_allocator_abba import (
    ALLOCATION_METRICS,
    ALLOCATION_SCOPE,
    EXPECTED_DELTA_BY_COUNT,
    ValidationError,
    main,
    validate_paths,
)


def _sha(seed: str) -> str:
    return (seed.encode("ascii") * 64)[:64].decode("ascii")


def _revision(seed: str) -> str:
    return (seed.encode("ascii") * 40)[:40].decode("ascii")


def _measured(values: list[int], scope: str = ALLOCATION_SCOPE) -> dict[str, object]:
    return {"values": values, "status": "measured", "scope": scope}


def _absent(status: str, scope: str) -> dict[str, object]:
    return {"status": status, "scope": scope}


def _contract() -> dict[str, object]:
    return {
        "schema_version": 1,
        "case": "opc_source_overlay_multi_part_noop",
        "cache_state": "warm",
        "samples_per_case": 30,
        "warmup_iterations_per_case": 3,
        "execution_workers": [1],
        "abba_order": ["A1_control", "B1_candidate", "B2_candidate", "A2_control"],
        "tool": {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": "litchi-perf-baseline-alloc",
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": "system_allocator_operation_scoped",
        },
        "environment": {
            "rustc_version": "rustc 1.98.1 (test)",
            "allocator": "CountingSystemAllocator(std::alloc::System)",
            "target_os": "linux",
            "target_arch": "x86_64",
            "logical_cpus_available": 1,
            "rustflags": None,
            "cargo_build_target": None,
        },
        "legs": {
            "A1": {
                "implementation": "control",
                "revision": _revision("a"),
                "binary_sha256": _sha("1"),
                "binary_bytes": 1001,
                "mode_bits": 509,
                "profile": "release",
            },
            "B1": {
                "implementation": "candidate",
                "revision": _revision("b"),
                "binary_sha256": _sha("2"),
                "binary_bytes": 1002,
                "mode_bits": 509,
                "profile": "release",
            },
            "B2": {
                "implementation": "candidate",
                "revision": _revision("b"),
                "binary_sha256": _sha("2"),
                "binary_bytes": 1002,
                "mode_bits": 509,
                "profile": "release",
            },
            "A2": {
                "implementation": "control",
                "revision": _revision("a"),
                "binary_sha256": _sha("1"),
                "binary_bytes": 1001,
                "mode_bits": 509,
                "profile": "release",
            },
        },
        "corpora": [
            {
                "shape": "overlay-small",
                "payload_kind": "compressible",
                "base_name": "overlay-small-compressible",
                "generator": "litchi-opc-source-overlay-multi-part-v1",
                "package_format": "OPC/ZIP",
                "compression": "deflate",
                "entry_count": 32,
                "archive_member_count": 34,
                "entry_bytes": 1024,
                "uncompressed_payload_bytes": 32768,
                "archive_bytes": 7451,
                "archive_sha256": "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509",
                "target_entry": "benchmark/parts/00016.bin",
                "target_payload_bytes": 1024,
                "target_payload_sha256": "5b7b9793a43d08ca2c0670289d932541377407ed352ab9f6c145f63d19de9f98",
            },
            {
                "shape": "overlay-large",
                "payload_kind": "incompressible",
                "base_name": "overlay-large-incompressible",
                "generator": "litchi-opc-source-overlay-multi-part-v1",
                "package_format": "OPC/ZIP",
                "compression": "deflate",
                "entry_count": 32,
                "archive_member_count": 34,
                "entry_bytes": 65536,
                "uncompressed_payload_bytes": 2097152,
                "archive_bytes": 2103195,
                "archive_sha256": "8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73",
                "target_entry": "benchmark/parts/00016.bin",
                "target_payload_bytes": 65536,
                "target_payload_sha256": "e17b543eec6b4d3534978d7d59e7240dcbf0f2a2050fd80f32ea3daec266aa73",
            },
            {
                "shape": "overlay-media-incompressible",
                "payload_kind": "incompressible",
                "base_name": "overlay-media-incompressible-incompressible",
                "generator": "litchi-opc-source-overlay-multi-part-v1",
                "package_format": "OPC/ZIP",
                "compression": "deflate",
                "entry_count": 32,
                "archive_member_count": 34,
                "entry_bytes": 262144,
                "uncompressed_payload_bytes": 8388608,
                "archive_bytes": 8396580,
                "archive_sha256": "bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883",
                "target_entry": "benchmark/parts/00016.bin",
                "target_payload_bytes": 262144,
                "target_payload_sha256": "3ad07c7e34d3dd6d9ff75b696ccbdd702777b6e4dea04b19bbe3d0aa6d21cdeb",
            },
        ],
        "expected_deltas": {
            str(count): values for count, values in EXPECTED_DELTA_BY_COUNT.items()
        },
    }


def _configuration() -> dict[str, object]:
    return {
        "samples_per_case": 30,
        "warmup_iterations_per_case": 3,
        "filesystem_cache_states": ["warm", "cold-requested"],
        "filesystem_fresh_child_per_sample": True,
        "filesystem_process_isolated": True,
        "filesystem_root_selected": False,
        "cases": ["opc_source_overlay_multi_part_noop"],
        "corpus_shapes": ["tiny", "many-small", "few-large", "wide-root"],
        "payload_kinds": ["compressible", "incompressible"],
        "writer_shapes": ["tiny", "large", "payload-heavy"],
        "xlsx_shapes": ["tiny", "medium", "dense-wide"],
        "xlsb_shapes": ["tiny", "medium", "large", "sparse"],
        "xlsx_cell_crud_shapes": ["medium", "dense-sparse"],
        "xlsx_row_visibility_shapes": ["medium", "large"],
        "semantic_shapes": ["tiny", "medium", "large"],
        "rtf_variants": ["plain"],
        "range_simulation": {
            "fixed_latency_us": 100,
            "request_overhead_us": 25,
            "bandwidth_bytes_per_second": 50 * 1024 * 1024,
            "max_physical_range_bytes": 4 * 1024,
        },
        "execution_workers": [1],
    }


def _environment(role: str, contract: dict[str, object]) -> dict[str, object]:
    revisions = contract["legs"]
    return {
        "rustc_version": contract["environment"]["rustc_version"],
        "git_revision": revisions[role]["revision"],
        "git_worktree_dirty": False,
        "logical_cpus_available": 1,
        "allocator": "CountingSystemAllocator(std::alloc::System)",
        "rustflags": None,
        "cargo_build_target": None,
        "perf_event_paranoid": "1",
        "os": "linux",
        "kernel": "Linux test",
        "cpu_model": "test-cpu",
        "total_memory_bytes": 1024 * 1024,
        "page_size_bytes": 4096,
        "filesystem_type": "tmpfs",
        "source_destination_same_device": True,
        "cpu_affinity": "2",
        "storage_identifier": None,
    }


def _metric_status(status: str, scope: str, reason: str) -> dict[str, object]:
    return {"status": status, "scope": scope, "reason": reason}


def _parallel(results: list[dict[str, object]]) -> dict[str, object]:
    cases = []
    for result in results:
        cases.append(
            {
                "case": result["case"],
                "corpus_sha256": result["corpus"]["archive_sha256"],
                "configured_worker_count": _metric_status("not_applicable", "configuration.execution_workers", "result has no explicit worker-budget field"),
                "observed_local_worker_count": _metric_status("not_applicable", "result.source.opc_cache.worker_count", "result does not create an explicit local worker team"),
                "deterministic_task_count": _metric_status("not_applicable", "result.execution.logical_tasks", "result has neither an explicit execution context nor range simulation"),
                "deterministic_chunk_count": _metric_status("not_applicable", "result.source.simulation.physical_request_count", "result has neither an explicit execution context nor range simulation"),
                "lock_wait_ns": _metric_status("unavailable", "lock_wait_ns", "no exact instrumented lock boundary is present; waiter counts are not timed"),
            }
        )
    return {
        "schema_version": 1,
        "scope": "explicit_local_execution_only",
        "claim": "descriptive",
        "configured_worker_budget": {"status": "measured", "value": [1], "scope": "configuration.execution_workers"},
        "observed_process_thread_count": {"status": "unavailable", "scope": "process_thread_count", "reason": "no process-global thread counter is collected"},
        "cases": cases,
    }


def _allocation(role: str, count: int) -> dict[str, object]:
    control = role in ("A1", "A2")
    # Keep all six vectors constant so the ABBA equality and exact-delta gates
    # are independent of elapsed-order permutations.
    values = {
        "allocation_calls": 1000,
        "deallocation_calls": 900,
        "reallocation_calls": 7,
        "failed_allocation_calls": 0,
        "allocated_bytes": 10_000_000,
        "deallocated_bytes": 9_000_000,
    }
    if not control:
        for metric, delta in EXPECTED_DELTA_BY_COUNT[count].items():
            values[metric] += delta
    values.update(
        {
            "live_bytes_before": 7000 + (1 if role == "B1" else 2 if role == "B2" else 0),
            "live_bytes_after": 7100 + (1 if role == "B1" else 2 if role == "B2" else 0),
            "peak_live_bytes_before": 8000,
            "peak_live_bytes_after": 8200,
        }
    )
    return {
        "status": "measured",
        "scope": ALLOCATION_SCOPE,
        **{metric: _measured([values[metric]] * 30) for metric in ALLOCATION_METRICS},
    }


def _operation(role: str, count: int, elapsed: list[int], order: list[int], sink: dict[str, object]) -> dict[str, object]:
    source = {
        "status": "not_applicable",
        "counter_scope": "not_applicable_in_process_sink",
        "logical_read_calls": _absent("not_applicable", "operation_logical_read_at"),
        "logical_read_requested_bytes": _absent("not_applicable", "operation_logical_read_at"),
        "logical_read_returned_bytes": _absent("not_applicable", "operation_logical_read_at"),
        "logical_read_largest_requested_bytes": _absent("not_applicable", "operation_logical_read_at"),
        "logical_read_largest_returned_bytes": _absent("not_applicable", "operation_logical_read_at"),
        "logical_read_pattern": _absent("not_applicable", "operation_logical_read_at_range_order_not_physical_io"),
        "compressed_bytes": _absent("not_applicable", "unavailable_read_at_has_no_compressed_member_boundary"),
        "decompressed_bytes": _absent("not_applicable", "unavailable_read_at_has_no_decompressed_byte_boundary"),
        "recompressed_bytes": _absent("not_applicable", "unavailable_atomic_save_has_no_recompressed_byte_boundary"),
        "max_concurrent_reads": _absent("not_applicable", "operation_logical_read_at"),
    }
    process_keys = (
        "user_cpu_ticks", "system_cpu_ticks", "clock_ticks_per_second", "minor_faults", "major_faults",
        "voluntary_context_switches", "nonvoluntary_context_switches", "rss_delta_bytes", "peak_rss_bytes",
        "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes", "syscr", "syscw",
    )
    process = {"status": "unavailable"}
    process.update({key: _absent("unavailable", "procfs_operation_delta") for key in process_keys})
    process["rchar"]["scope"] = "procfs_operation_delta"
    process["wchar"]["scope"] = "procfs_operation_delta"
    process["read_bytes"]["scope"] = "procfs_operation_delta"
    process["write_bytes"]["scope"] = "procfs_operation_delta"
    process["cancelled_write_bytes"]["scope"] = "procfs_operation_delta"
    process["syscr"]["scope"] = "procfs_operation_delta"
    process["syscw"]["scope"] = "procfs_operation_delta"
    op_sink = {
        "status": "not_applicable",
        "output_bytes": _absent("not_applicable", "post_operation_output_length_not_sink_write_volume"),
        "write_status": "measured",
        "accepted_bytes": _measured([sink["accepted_bytes"]] * 30, "logical_sink_accepted_write_bytes"),
        "write_calls": _measured([sink["write_calls"]] * 30, "logical_sink_accepted_write_calls"),
        "largest_write": _measured([sink["largest_write"]] * 30, "logical_sink_largest_accepted_write"),
        "write_size_buckets": {"status": "measured"},
    }
    for key, scope in (
        ("bytes_0", "logical_sink_accepted_write_size_bucket_counts"),
        ("bytes_1_to_512", "logical_sink_accepted_write_size_bucket_counts"),
        ("bytes_513_to_4096", "logical_sink_accepted_write_size_bucket_counts"),
        ("bytes_4097_to_16384", "logical_sink_accepted_write_size_bucket_counts"),
        ("bytes_16385_to_65536", "logical_sink_accepted_write_size_bucket_counts"),
        ("bytes_over_65536", "logical_sink_accepted_write_size_bucket_counts"),
    ):
        op_sink["write_size_buckets"][key] = _measured([sink["write_size_buckets"][key]] * 30, scope)
    absent_publication = {"status": "not_applicable", "changed_spans": _absent("not_applicable", "logical_publication_counter"), "published_bytes": _absent("not_applicable", "logical_publication_counter")}
    absent_materialization = {"status": "not_applicable", "opc_parts": _absent("not_applicable", "logical_materialization_counter")}
    absent_phase = {key: _absent("not_applicable", "timed_cfb_phase_elapsed_ns" if key == "elapsed_ns" else "timed_cfb_phase_logical_read_at") for key in ("elapsed_ns", "logical_read_calls", "logical_read_requested_bytes", "logical_read_returned_bytes")}
    allocation = _allocation(role, count)
    return {
        "sample_count": 30,
        "sample_indices": order,
        "alignment": "elapsed_ns.samples_by_elapsed_then_sample_index",
        "latency_claim": "comparable_timed_operation",
        "source": source,
        "process": process,
        "sink": op_sink,
        "publication": absent_publication,
        "materialization": absent_materialization,
        "cfb_phases": {"status": "not_applicable", "open": absent_phase, "plan": copy.deepcopy(absent_phase), "atomic_publication": copy.deepcopy(absent_phase)},
        "allocation": allocation,
    }


def _result(role: str, shape: str, count: int, contract: dict[str, object], role_index: int) -> dict[str, object]:
    corpus_spec = next(item for item in contract["corpora"] if item["shape"] == shape)
    corpus = {key: value for key, value in corpus_spec.items() if key != "base_name"}
    corpus["name"] = f"{corpus_spec['base_name']}-count-{count}"
    corpus["xlsx"] = None
    base = 1000 + role_index * 100
    elapsed = [base + index for index in range(30)]
    order = list(range(30))
    phases = {
        "preparation_ns": [100] * 30,
        "open_ns": [200] * 30,
        "planning_ns": [300] * 30,
        "publication_ns": [value - 600 for value in elapsed],
    }
    source_values = {
        "read_calls": [count] * 30,
        "read_bytes": [count * 123] * 30,
        "ordinary_payload_read_calls": [count] * 30,
        "ordinary_payload_read_bytes": [count * corpus_spec["entry_bytes"]] * 30,
        "max_in_flight_reads": [0] * 30,
    }
    observed_overlay = {
        "implementation": "SourceBackedPackage::write_part_overlays_to_stream",
        "timing_scope": "elapsed_ns is explicitly the sum of preparation_ns, open_ns, planning_ns, and publication_ns; operation_metrics.allocation covers only the write_part_overlays_to_stream publication call; structural setup belongs to those named phases",
        "performance_claim": "none",
        "overlay_mode": "noop",
        "replacement_semantics": "non-empty equal-payload replacement plan; semantic no-op",
        "overlay_count": count,
        "source_shape": shape,
        "payload_kind": corpus_spec["payload_kind"],
        "source_bytes": corpus_spec["archive_bytes"],
        "source_sha256": corpus_spec["archive_sha256"],
        "expected_eager_sha256": corpus_spec["archive_sha256"],
        "source_cache_max_bytes": corpus_spec["uncompressed_payload_bytes"],
        "source_cache_max_entries": 32,
        "sink_max_bytes": corpus_spec["archive_bytes"] * 2 + 65536,
        "sink_max_write": 65536,
        **phases,
        "cache_before_publication_hits": [0] * 30,
        "cache_before_publication_cold_loads": [0] * 30,
        "cache_before_publication_retained_entries": [0] * 30,
        "cache_before_publication_retained_bytes": [0] * 30,
        "source_cache_after_publication_probe_hits": [0] * 30,
        "source_cache_after_publication_probe_cold_loads": [count] * 30,
        "source_cache_after_publication_probe_retained_entries": [count] * 30,
        "source_cache_after_publication_probe_retained_bytes": [count * corpus_spec["entry_bytes"]] * 30,
        "reopened_output_cache_hits": [0] * 30,
        "reopened_output_cache_cold_loads": [0] * 30,
        "reopened_output_cache_retained_entries": [0] * 30,
        "reopened_output_cache_retained_bytes": [0] * 30,
        "observed_after_publication_source_read_calls": source_values["read_calls"],
        "observed_after_publication_source_read_bytes": source_values["read_bytes"],
        "observed_after_publication_ordinary_payload_read_calls": source_values["ordinary_payload_read_calls"],
        "observed_after_publication_ordinary_payload_read_bytes": source_values["ordinary_payload_read_bytes"],
        "expected_eager_semantic_verified": True,
        "raw_members_and_order_preservation_verified": True,
        "equal_payload_noop_source_verified": True,
        "observed_output_sha256": [corpus_spec["archive_sha256"]] * 30,
    }
    sink = {
        "accepted_bytes": 1000 + count,
        "write_calls": 10,
        "largest_write": 100,
        "write_size_buckets": {
            "bytes_0": 0,
            "bytes_1_to_512": 10,
            "bytes_513_to_4096": 0,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": 0,
            "bytes_over_65536": 0,
        },
    }
    source = {**source_values, "opc_source_overlay": observed_overlay}
    return {
        "case": "opc_source_overlay_multi_part_noop",
        "corpus": corpus,
        "elapsed_ns": {
            "unit": "ns",
            "samples": elapsed,
            "sample_order": order,
            "min": elapsed[0],
            "p50": elapsed[14],
            "p95": elapsed[28],
            "p99": elapsed[29],
            "max": elapsed[-1],
            "mean": float(sum(elapsed)) / len(elapsed),
            "standard_deviation": 1.0,
            "confidence_interval_95": {"method": "test", "lower": 1.0, "upper": 2.0},
        },
        "sink": sink,
        "source": source,
        "output_sha256": corpus_spec["archive_sha256"],
        "operation_metrics": _operation(role, count, elapsed, order, sink),
    }


def _documents() -> tuple[dict[str, object], dict[str, object]]:
    contract = _contract()
    reports = {}
    for role_index, role in enumerate(("A1", "B1", "B2", "A2")):
        results = []
        for shape in ("overlay-small", "overlay-large", "overlay-media-incompressible"):
            for count in (2, 8, 32):
                results.append(_result(role, shape, count, contract, role_index))
        reports[role] = {
            "schema_version": 1,
            "tool": copy.deepcopy(contract["tool"]),
            "binary_identity": {
                "path": f"/tmp/{role}/litchi-perf-baseline-alloc",
                "binary_sha256": contract["legs"][role]["binary_sha256"],
                "binary_bytes": contract["legs"][role]["binary_bytes"],
                "mode_bits": contract["legs"][role]["mode_bits"],
                "executable": True,
                "profile": "release",
            },
            "environment": _environment(role, contract),
            "configuration": _configuration(),
            "parallel_metrics": _parallel(results),
            "results": results,
        }
    return contract, reports


class OpcOverlayAllocatorAbbaValidatorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.contract, self.reports = _documents()

    def _validate(self, contract: object | None = None, reports: dict[str, object] | None = None) -> dict[str, object]:
        contract = copy.deepcopy(contract if contract is not None else self.contract)
        reports = copy.deepcopy(reports if reports is not None else self.reports)
        with tempfile.TemporaryDirectory(prefix="litchi-0402-overlay-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(contract, indent=2) + "\n", encoding="utf-8")
            paths = {}
            for role in ("A1", "B1", "B2", "A2"):
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(reports[role], indent=2) + "\n", encoding="utf-8")
                paths[role] = path
            return validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)

    def test_valid_matrix_produces_compact_projection(self) -> None:
        projection = self._validate()
        self.assertEqual(projection["validation"]["report_count"], 4)
        self.assertEqual(len(projection["rows"]), 9)
        self.assertEqual(projection["rows"][0]["allocation"]["allocation_calls"]["delta"], -2)
        self.assertFalse(projection["rows"][0]["claimability"]["allocator_elapsed_ns"])
        raw_hashes = projection["provenance"]["raw_report_sha256"]
        self.assertEqual(set(raw_hashes), {"A1", "B1", "B2", "A2"})
        self.assertEqual(len(set(raw_hashes.values())), 4)
        validator_source_sha256 = projection["provenance"]["validator_source_sha256"]
        self.assertEqual(len(validator_source_sha256), 64)
        self.assertEqual(validator_source_sha256, validator_source_sha256.lower())
        self.assertTrue(projection["validation"]["four_distinct_raw_report_byte_streams"])
        self.assertTrue(projection["validation"]["operation_vector_aligned_by_sample_indices_and_elapsed_order"])
        self.assertFalse(projection["claimability"]["independent_process_proof"]["claimable"])
        self.assertEqual(projection["protocol"]["configuration_cache_states"], ["warm", "cold-requested"])
        self.assertIn("does not imply a cold run", projection["protocol"]["cache_claim_scope"])

    def test_rejects_duplicate_json_keys(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0402-overlay-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ("A1", "B1", "B2", "A2"):
                path = root / f"{role.lower()}.json"
                text = json.dumps(self.reports[role])
                if role == "A1":
                    text = text.replace('{"schema_version": 1,', '{"schema_version": 1, "schema_version": 1,', 1)
                path.write_text(text, encoding="utf-8")
                paths[role] = path
            with self.assertRaisesRegex(ValidationError, "duplicate JSON object key"):
                validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)

    def test_rejects_nonfinite_json_number(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0402-overlay-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ("A1", "B1", "B2", "A2"):
                path = root / f"{role.lower()}.json"
                text = json.dumps(self.reports[role])
                if role == "A1":
                    text = text.replace('"mean": 1014.5', '"mean": NaN', 1)
                path.write_text(text, encoding="utf-8")
                paths[role] = path
            with self.assertRaisesRegex(ValidationError, "non-finite JSON number"):
                validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)

    def test_rejects_wrong_binary_identity(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["B1"]["binary_identity"]["binary_sha256"] = _sha("9")
        with self.assertRaisesRegex(ValidationError, "binary_sha256"):
            self._validate(reports=reports)

    def test_rejects_dirty_worktree(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A2"]["environment"]["git_worktree_dirty"] = True
        with self.assertRaisesRegex(ValidationError, "git_worktree_dirty"):
            self._validate(reports=reports)

    def test_rejects_nonprotocol_cpu_affinity(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["B2"]["environment"]["cpu_affinity"] = "3"
        with self.assertRaisesRegex(ValidationError, "cpu_affinity"):
            self._validate(reports=reports)

    def test_accepts_real_global_cache_selector_list_but_rejects_mutation(self) -> None:
        self._validate(reports=self.reports)
        reports = copy.deepcopy(self.reports)
        reports["B2"]["configuration"]["filesystem_cache_states"] = ["warm"]
        with self.assertRaisesRegex(ValidationError, "filesystem_cache_states"):
            self._validate(reports=reports)

    def test_rejects_missing_matrix_row(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["B2"]["results"].pop()
        with self.assertRaisesRegex(ValidationError, "matrix rows"):
            self._validate(reports=reports)

    def test_rejects_sample_order_or_allocation_cardinality_mutation(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["elapsed_ns"]["sample_order"][0] = 1
        with self.assertRaisesRegex(ValidationError, "sample_order"):
            self._validate(reports=reports)

        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["operation_metrics"]["allocation"]["allocated_bytes"]["values"].pop()
        with self.assertRaisesRegex(ValidationError, "allocated_bytes"):
            self._validate(reports=reports)

    def test_rejects_raw_alignment_and_phase_identity_mutations(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["B1"]["results"][1]["operation_metrics"]["allocation"]["allocation_calls"]["values"][0] += 1
        with self.assertRaisesRegex(ValidationError, "allocation_calls must be constant"):
            self._validate(reports=reports)

        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][2]["source"]["opc_source_overlay"]["publication_ns"][0] += 1
        with self.assertRaisesRegex(ValidationError, "phase sum"):
            self._validate(reports=reports)

    def test_rejects_source_sink_oracle_mutations(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A2"]["results"][3]["source"]["opc_source_overlay"]["equal_payload_noop_source_verified"] = False
        with self.assertRaisesRegex(ValidationError, "equal_payload_noop_source_verified"):
            self._validate(reports=reports)

        reports = copy.deepcopy(self.reports)
        reports["B2"]["results"][4]["sink"]["write_calls"] += 1
        with self.assertRaisesRegex(ValidationError, "write_size_buckets"):
            self._validate(reports=reports)

    def test_rejects_b2_source_identity_mutation(self) -> None:
        reports = copy.deepcopy(self.reports)
        result = reports["B2"]["results"][0]
        result["source"]["read_bytes"][0] += 1
        result["source"]["opc_source_overlay"]["observed_after_publication_source_read_bytes"][0] += 1
        with self.assertRaisesRegex(ValidationError, "A1/B2 source identity"):
            self._validate(reports=reports)

    def test_rejects_b2_sink_identity_mutation(self) -> None:
        reports = copy.deepcopy(self.reports)
        result = reports["B2"]["results"][0]
        result["sink"]["accepted_bytes"] += 1
        result["operation_metrics"]["sink"]["accepted_bytes"]["values"] = [result["sink"]["accepted_bytes"]] * 30
        with self.assertRaisesRegex(ValidationError, "A1/B2 sink identity"):
            self._validate(reports=reports)

    def test_rejects_insufficient_ordinary_payload_reads(self) -> None:
        reports = copy.deepcopy(self.reports)
        result = reports["B2"]["results"][0]
        result["source"]["ordinary_payload_read_calls"][0] = 1
        with self.assertRaisesRegex(ValidationError, "ordinary_payload_read_calls must be at least"):
            self._validate(reports=reports)

        reports = copy.deepcopy(self.reports)
        result = reports["B2"]["results"][0]
        result["source"]["ordinary_payload_read_bytes"][0] = 0
        with self.assertRaisesRegex(ValidationError, "ordinary_payload_read_bytes must be positive"):
            self._validate(reports=reports)

    def test_rejects_sink_bounds_mutations(self) -> None:
        reports = copy.deepcopy(self.reports)
        result = reports["A1"]["results"][0]
        max_bytes = self.contract["corpora"][0]["archive_bytes"] * 2 + 65536
        result["sink"]["accepted_bytes"] = max_bytes + 1
        result["operation_metrics"]["sink"]["accepted_bytes"]["values"] = [max_bytes + 1] * 30
        with self.assertRaisesRegex(ValidationError, "accepted_bytes must be <= sink_max_bytes"):
            self._validate(reports=reports)

        reports = copy.deepcopy(self.reports)
        result = reports["A1"]["results"][0]
        result["sink"]["largest_write"] = result["sink"]["accepted_bytes"] + 1
        result["operation_metrics"]["sink"]["largest_write"]["values"] = [result["sink"]["largest_write"]] * 30
        with self.assertRaisesRegex(ValidationError, "largest_write must be <= accepted_bytes"):
            self._validate(reports=reports)

    def test_rejects_nonconstant_counter_vector_even_when_delta_is_preserved(self) -> None:
        reports = copy.deepcopy(self.reports)
        for role in ("A1", "B1", "B2", "A2"):
            values = reports[role]["results"][0]["operation_metrics"]["allocation"]["allocated_bytes"]["values"]
            values[7] += 1
        with self.assertRaisesRegex(ValidationError, "allocated_bytes must be constant"):
            self._validate(reports=reports)

    def test_rejects_wrong_exact_delta_policy(self) -> None:
        contract = copy.deepcopy(self.contract)
        contract["expected_deltas"]["8"]["allocated_bytes"] += 1
        with self.assertRaisesRegex(ValidationError, "expected_deltas.8"):
            self._validate(contract=contract)

    def test_rejects_contract_corpus_oracle_redefinition(self) -> None:
        contract = copy.deepcopy(self.contract)
        contract["corpora"][0]["archive_bytes"] += 1
        with self.assertRaisesRegex(ValidationError, "retained corpus identity"):
            self._validate(contract=contract)

    def test_allocator_elapsed_statistics_are_not_claimed(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["elapsed_ns"]["mean"] = -12345.0
        projection = self._validate(reports=reports)
        self.assertFalse(projection["claimability"]["allocator_elapsed_ns"]["claimable"])

    def test_accepts_nonidentity_elapsed_sample_order(self) -> None:
        reports = copy.deepcopy(self.reports)
        result = reports["A1"]["results"][0]
        result["elapsed_ns"]["sample_order"] = list(reversed(range(30)))
        result["operation_metrics"]["sample_indices"] = list(reversed(range(30)))
        self._validate(reports=reports)

    def test_requires_distinct_report_paths(self) -> None:
        contract = copy.deepcopy(self.contract)
        reports = copy.deepcopy(self.reports)
        with tempfile.TemporaryDirectory(prefix="litchi-0402-overlay-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(contract), encoding="utf-8")
            paths = []
            for role in ("A1", "B1", "B2", "A2"):
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(reports[role]), encoding="utf-8")
                paths.append(path)
            with self.assertRaisesRegex(ValidationError, "distinct report paths"):
                validate_paths(paths[0], paths[1], paths[2], paths[2], contract_path)

    def test_rejects_output_collision_with_inputs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0402-overlay-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ("A1", "B1", "B2", "A2"):
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(self.reports[role]), encoding="utf-8")
                paths[role] = path
            base_args = [
                "--contract",
                str(contract_path),
                "--a1",
                str(paths["A1"]),
                "--b1",
                str(paths["B1"]),
                "--b2",
                str(paths["B2"]),
                "--a2",
                str(paths["A2"]),
            ]
            for colliding_input in (contract_path, paths["A1"]):
                before = colliding_input.read_bytes()
                stderr = io.StringIO()
                with contextlib.redirect_stderr(stderr):
                    status = main(base_args + ["--output", str(colliding_input)])
                self.assertEqual(status, 1)
                self.assertIn("--output must differ", stderr.getvalue())
                self.assertEqual(colliding_input.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
