from __future__ import annotations

import copy
import contextlib
import io
import json
import math
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from tools import validate_opc_serial_eager_latency_abba as validator


ROLES = ("A1", "B1", "B2", "A2")
CONTROL_REVISION = validator.CONTROL_REVISION
CANDIDATE_REVISION = validator.CANDIDATE_REVISION

INDEPENDENT_CORPUS_IDENTITIES = {
    "tiny": {
        "name": "tiny-compressible",
        "generator": "litchi-opc-synthetic-v2",
        "package_format": "OPC/ZIP",
        "shape": "tiny",
        "payload_kind": "compressible",
        "compression": "deflate",
        "entry_count": 3,
        "part_count": 3,
        "archive_member_count": 5,
        "entry_bytes": 512,
        "uncompressed_payload_bytes": 1_536,
        "archive_bytes": 1_310,
        "archive_sha256": "1e28b8a9049a82f07e8ea88b2d492ef522d2da793d22fa50e2fe7f354dca3e2a",
        "target_entry": "benchmark/parts/00001.bin",
        "target_payload_bytes": 512,
        "target_payload_sha256": "630b1da45fe604eda3b5468b7c9ca7facfbd404941779786276a69ff870e4bdd",
        "part_names_sha256": "5458f5d1eb9283e10cd7057abf8f63cce9d1e0b6c57c5f9f945a9bad3b99cda4",
        "part_payload_sha256": "d1baa4a40fc63856136504f95933bcb2bb3da28f2000cabe1153eaee88c723c0",
    },
    "many-small": {
        "name": "many-small-incompressible",
        "generator": "litchi-opc-synthetic-v2",
        "package_format": "OPC/ZIP",
        "shape": "many-small",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 256,
        "part_count": 256,
        "archive_member_count": 258,
        "entry_bytes": 1_024,
        "uncompressed_payload_bytes": 262_144,
        "archive_bytes": 303_003,
        "archive_sha256": "183178dec5b0fd578e5af04279032368598eec79da7caf0441fc979ce8fc14a0",
        "target_entry": "benchmark/parts/00128.bin",
        "target_payload_bytes": 1_024,
        "target_payload_sha256": "05fd26cad1f538b7ed415a0f525a13896823b02abcf22ad1746172f035a2149d",
        "part_names_sha256": "82415ca7ad25155c41df5d93707c95e5fcc31e66cde226ff046fc84906f56bc2",
        "part_payload_sha256": "7bdf372948a4f914aea31187d1f2813254957cd907279690b022ef00737caaa7",
    },
    "few-large": {
        "name": "few-large-incompressible",
        "generator": "litchi-opc-synthetic-v2",
        "package_format": "OPC/ZIP",
        "shape": "few-large",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 4,
        "part_count": 4,
        "archive_member_count": 6,
        "entry_bytes": 4 * 1024 * 1024,
        "uncompressed_payload_bytes": 16 * 1024 * 1024,
        "archive_bytes": 16_783_632,
        "archive_sha256": "a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6",
        "target_entry": "benchmark/parts/00002.bin",
        "target_payload_bytes": 4 * 1024 * 1024,
        "target_payload_sha256": "3dbf6225021a99c1da8750a738bde21f57591c0be1a60aa510966c47ee25b098",
        "part_names_sha256": "d48e27d95e97a4de43e476096910540416f6e19eb54a3759d5ca081b4136166c",
        "part_payload_sha256": "ac1e942c87db2e622c1e1c2efd1046e5d791a44db73bd6255078f8816d922db3",
    },
}


def _digest(digit: str) -> str:
    return (digit * 64)[:64]


def _absent(status: str, scope: str) -> dict[str, str]:
    return {"status": status, "scope": scope}


def _elapsed(start: int) -> dict[str, object]:
    samples = [start + index for index in range(validator.SAMPLE_COUNT)]
    mean = 0.0
    squared = 0.0
    for index, sample in enumerate(samples):
        current = float(sample)
        count = float(index + 1)
        delta = current - mean
        next_mean = mean + delta / count
        squared += delta * (current - next_mean)
        mean = next_mean
    deviation = math.sqrt(squared / (len(samples) - 1))
    margin = validator._student_t_critical_95(len(samples) - 1) * deviation / math.sqrt(len(samples))
    return {
        "unit": "ns",
        "samples": samples,
        "sample_order": list(range(validator.SAMPLE_COUNT)),
        "min": samples[0],
        "p50": (samples[249] + samples[250]) // 2,
        "p95": samples[474],
        "p99": samples[494],
        "max": samples[-1],
        "mean": mean,
        "standard_deviation": deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(mean - margin, 0.0),
            "upper": mean + margin,
        },
    }


def _operation() -> dict[str, object]:
    source = {
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
        source[key] = _absent("not_applicable", validator.SOURCE_SCOPE)
    source["logical_read_pattern"] = _absent("not_applicable", validator.SOURCE_PATTERN_SCOPE)
    source["compressed_bytes"] = _absent("not_applicable", validator.SOURCE_COMPRESSED_SCOPE)
    source["decompressed_bytes"] = _absent("not_applicable", validator.SOURCE_DECOMPRESSED_SCOPE)
    source["recompressed_bytes"] = _absent("not_applicable", validator.SOURCE_RECOMPRESSED_SCOPE)
    process = {"status": "unavailable"}
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
        process[key] = _absent("unavailable", validator.PROCESS_SCOPE)
    process["rss_delta_bytes"] = _absent("unavailable", validator.RSS_SCOPE)
    process["peak_rss_bytes"] = _absent("unavailable", validator.HWM_SCOPE)
    sink = {
        "status": "not_applicable",
        "output_bytes": _absent("not_applicable", validator.OUTPUT_SCOPE),
        "write_status": "not_applicable",
        "accepted_bytes": _absent("not_applicable", validator.SINK_SCOPE),
        "write_calls": _absent("not_applicable", validator.SINK_WRITE_CALLS_SCOPE),
        "largest_write": _absent("not_applicable", validator.SINK_LARGEST_WRITE_SCOPE),
        "write_size_buckets": {"status": "not_applicable"},
    }
    for key in validator.SINK_BUCKET_KEYS - {"status"}:
        sink["write_size_buckets"][key] = _absent("not_applicable", validator.SINK_BUCKET_SCOPE)
    publication = {
        "status": "not_applicable",
        "changed_spans": _absent("not_applicable", validator.PUBLICATION_SCOPE),
        "published_bytes": _absent("not_applicable", validator.PUBLICATION_SCOPE),
    }
    materialization = {
        "status": "not_applicable",
        "opc_parts": _absent("not_applicable", validator.MATERIALIZATION_SCOPE),
    }
    cfb = {"status": "not_applicable"}
    for phase in ("open", "plan", "atomic_publication"):
        cfb[phase] = {
            "elapsed_ns": _absent("not_applicable", validator.CFB_ELAPSED_SCOPE),
            "logical_read_calls": _absent("not_applicable", validator.CFB_SOURCE_SCOPE),
            "logical_read_requested_bytes": _absent("not_applicable", validator.CFB_SOURCE_SCOPE),
            "logical_read_returned_bytes": _absent("not_applicable", validator.CFB_SOURCE_SCOPE),
        }
    allocation = {"status": "unavailable", "scope": validator.ALLOCATOR_SCOPE}
    for metric in (
        *validator.COUNTER_METRICS,
        "live_bytes_before",
        "live_bytes_after",
        "peak_live_bytes_before",
        "peak_live_bytes_after",
    ):
        allocation[metric] = _absent("unavailable", validator.ALLOCATOR_SCOPE)
    return {
        "sample_count": validator.SAMPLE_COUNT,
        "sample_indices": list(range(validator.SAMPLE_COUNT)),
        "alignment": validator.ALIGNMENT,
        "latency_claim": validator.LATENCY_CLAIM,
        "source": source,
        "process": process,
        "sink": sink,
        "publication": publication,
        "materialization": materialization,
        "cfb_phases": cfb,
        "allocation": allocation,
    }


def _summary(shape: str) -> dict[str, object]:
    oracle = validator.CORPUS_ORACLE[shape]
    model = {
        "comparison": "candidate-control",
        "status": "expected_not_observed",
        **validator.EXPECTED_ALLOCATOR_MODEL[shape],
    }
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
        "observed_part_counts": [oracle["part_count"]] * validator.SAMPLE_COUNT,
        "observed_part_names_sha256": [oracle["part_names_sha256"]] * validator.SAMPLE_COUNT,
        "observed_part_payload_sha256": [oracle["part_payload_sha256"]] * validator.SAMPLE_COUNT,
        "observed_content_types_verified": [True] * validator.SAMPLE_COUNT,
        "observed_root_relationship_verified": [True] * validator.SAMPLE_COUNT,
        "observed_main_target_verified": [True] * validator.SAMPLE_COUNT,
        "observed_deterministic_payload_hashes_verified": [True] * validator.SAMPLE_COUNT,
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
        "case": validator.CASE,
        "cache_state": "warm",
        "samples_per_case": validator.SAMPLE_COUNT,
        "warmup_iterations_per_case": validator.WARMUP_COUNT,
        "execution_workers": [1],
        "abba_order": validator.ABBA_ORDER,
        "tool": {
            "name": validator.TOOL_NAME,
            "version": validator.TOOL_VERSION,
            "binary": validator.NORMAL_BINARY,
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": "none",
        },
        "environment": {
            "rustc_version": "rustc 1.98.1 (test)",
            "allocator": validator.ALLOCATOR,
            "target_os": "linux",
            "target_arch": "x86_64",
            "logical_cpus_available": 32,
            "cpu_model": "test-cpu",
            "cpu_affinity": "2",
            "rustflags": None,
            "cargo_build_target": None,
        },
        "legs": {"A1": control, "A2": copy.deepcopy(control), "B1": candidate, "B2": copy.deepcopy(candidate)},
        "corpora": [copy.deepcopy(validator.CORPUS_ORACLE[shape]) for shape in validator.SHAPES],
        "expected_deltas": copy.deepcopy(validator.EXPECTED_ALLOCATOR_MODEL),
    }


def _report(role: str, *, starts: dict[str, int] | None = None) -> dict[str, object]:
    contract = _contract()
    leg = contract["legs"][role]
    environment = {
        "rustc_version": contract["environment"]["rustc_version"],
        "git_revision": leg["revision"],
        "git_worktree_dirty": False,
        "logical_cpus_available": 32,
        "allocator": validator.ALLOCATOR,
        "rustflags": None,
        "cargo_build_target": None,
        "perf_event_paranoid": "1",
        "os": "linux",
        "kernel": "Linux test",
        "cpu_model": "test-cpu",
        "total_memory_bytes": 1_073_741_824,
        "page_size_bytes": 4096,
        "filesystem_type": "tmpfs",
        "source_destination_same_device": None,
        "cpu_affinity": "2",
        "storage_identifier": None,
    }
    role_starts = starts or {"A1": 1000, "B1": 900, "B2": 905, "A2": 1005}
    results = []
    parallel_cases = []
    for shape_index, shape in enumerate(validator.SHAPES):
        oracle = validator.CORPUS_ORACLE[shape]
        corpus = {key: value for key, value in oracle.items() if key not in {"part_count", "part_names_sha256", "part_payload_sha256"}}
        corpus["xlsx"] = None
        results.append(
            {
                "case": validator.CASE,
                "corpus": corpus,
                "elapsed_ns": _elapsed(role_starts[role] + shape_index * 100_000),
                "sink": None,
                "source": {
                    "read_calls": [],
                    "read_bytes": [],
                    "ordinary_payload_read_calls": [],
                    "ordinary_payload_read_bytes": [],
                    "max_in_flight_reads": [],
                    "opc_serial_eager_open": _summary(shape),
                },
                "execution": {
                    "worker_count": 1,
                    "logical_tasks": 1,
                    "logical_bytes": oracle["archive_bytes"],
                },
                "output_sha256": oracle["archive_sha256"],
                "operation_metrics": _operation(),
            }
        )
        parallel_cases.append(
            {
                "case": validator.CASE,
                "corpus_sha256": oracle["archive_sha256"],
                "configured_worker_count": {"status": "measured", "value": 1, "scope": "result.execution.worker_count"},
                "observed_local_worker_count": {"status": "not_applicable", "scope": "result.source.opc_cache.worker_count_with_one_created_local_worker_team", "reason": "result does not create an explicit local worker team"},
                "deterministic_task_count": {"status": "measured", "value": 1, "scope": "result.execution.logical_tasks"},
                "deterministic_chunk_count": {"status": "unavailable", "scope": "result.execution.deterministic_chunk_count", "reason": "no deterministic chunk counter is exposed"},
                "lock_wait_ns": {"status": "unavailable", "scope": "lock_wait_ns", "reason": "no exact instrumented lock boundary is present"},
            }
        )
    return {
        "schema_version": 1,
        "tool": copy.deepcopy(contract["tool"]),
        "binary_identity": {
            "path": f"/tmp/{role.lower()}/{validator.NORMAL_BINARY}",
            "binary_sha256": leg["binary_sha256"],
            "binary_bytes": leg["binary_bytes"],
            "mode_bits": leg["mode_bits"],
            "executable": True,
            "profile": "release",
        },
        "environment": environment,
        "configuration": validator._expected_configuration(),
        "parallel_metrics": {
            "schema_version": 1,
            "scope": "explicit_local_execution_only",
            "claim": "descriptive",
            "configured_worker_budget": {"status": "measured", "value": [1], "scope": "configuration.execution_workers"},
            "observed_process_thread_count": {"status": "unavailable", "scope": "process_thread_count", "reason": "no process-global thread counter is collected"},
            "cases": parallel_cases,
        },
        "results": results,
    }


class NormalLatencyValidatorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.contract = _contract()
        self.reports = {role: _report(role) for role in ROLES}

    def _validate(self, *, reports: dict[str, object] | None = None, contract: dict[str, object] | None = None) -> dict[str, object]:
        reports = self.reports if reports is None else reports
        contract = self.contract if contract is None else contract
        with tempfile.TemporaryDirectory(prefix="litchi-0403-normal-latency-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(contract, indent=2) + "\n", encoding="utf-8")
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(reports[role], indent=2) + "\n", encoding="utf-8")
                paths[role] = path
            return validator.validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)

    def _direct_contract(self) -> validator.Contract:
        return validator.Contract(
            raw_sha256="",
            tool=copy.deepcopy(self.contract["tool"]),
            environment=copy.deepcopy(self.contract["environment"]),
            legs=copy.deepcopy(self.contract["legs"]),
            corpora={
                shape: copy.deepcopy(self.contract["corpora"][index])
                for index, shape in enumerate(validator.SHAPES)
            },
        )

    def _validate_direct(
        self,
        contract: validator.Contract,
        *,
        reports: dict[str, object] | None = None,
    ) -> dict[str, object]:
        reports = self.reports if reports is None else reports
        with tempfile.TemporaryDirectory(prefix="litchi-0403-normal-latency-validator-") as directory:
            root = Path(directory)
            paths = {}
            for role in ROLES:
                path = root / f"{role.lower()}.json"
                path.write_text(json.dumps(reports[role], indent=2) + "\n", encoding="utf-8")
                paths[role] = path
            return validator.validate_paths(
                paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract=contract
            )

    def test_fixed_hashes_match_independent_producer_manifest(self) -> None:
        for shape, expected in INDEPENDENT_CORPUS_IDENTITIES.items():
            with self.subTest(shape=shape):
                self.assertEqual(set(expected), set(validator.CORPUS_ORACLE[shape]))
                for key, value in expected.items():
                    self.assertEqual(validator.CORPUS_ORACLE[shape][key], value)

    def test_accepts_rows_and_preserves_per_row_decisions_without_pooling(self) -> None:
        projection = self._validate()
        self.assertEqual(projection["validation"]["report_count"], 4)
        self.assertTrue(projection["validation"]["elapsed_statistics_recomputed"])
        self.assertFalse(projection["validation"]["pooled_claim"])
        self.assertEqual(projection["rows"][0]["elapsed_ns"]["accepted_statistics"], list(validator.STATISTICS))
        self.assertEqual(projection["rows"][0]["elapsed_ns"]["rejected_statistics"], {})
        self.assertFalse(projection["claimability"]["rss"]["claimable"])
        self.assertFalse(projection["claimability"]["cold_cache"]["claimable"])
        self.assertFalse(projection["claimability"]["fresh_child_per_sample"]["claimable"])
        self.assertFalse(projection["claimability"]["logical_io"]["claimable"])
        self.assertFalse(projection["claimability"]["physical_io"]["claimable"])
        self.assertFalse(projection["claimability"]["allocator"]["claimable"])

    def test_classifies_rejected_and_adverse_rows(self) -> None:
        reports = {role: _report(role, starts={"A1": 1000, "B1": 1200, "B2": 1300, "A2": 1100}) for role in ROLES}
        projection = self._validate(reports=reports)
        decision = projection["rows"][0]["elapsed_ns"]
        self.assertEqual(decision["accepted_statistics"], [])
        self.assertEqual(decision["adverse_both_statistics"], list(validator.STATISTICS))
        self.assertEqual(set(decision["rejected_statistics"]), set(validator.STATISTICS))

        reports = {role: _report(role, starts={"A1": 1000, "B1": 900, "B2": 905, "A2": 1400}) for role in ROLES}
        projection = self._validate(reports=reports)
        decision = projection["rows"][0]["elapsed_ns"]
        self.assertEqual(decision["accepted_statistics"], [])
        self.assertEqual(decision["adverse_both_statistics"], [])
        self.assertTrue(all("control drift" in reason for reason in decision["rejected_statistics"].values()))

    def test_rejects_duplicate_keys_nonfinite_and_unstable_ties(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0403-normal-latency-validator-") as directory:
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
            with self.assertRaisesRegex(validator.ValidationError, "duplicate JSON object key"):
                validator.validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)
            tie = copy.deepcopy(self.reports)
            elapsed = tie["A1"]["results"][0]["elapsed_ns"]
            elapsed["samples"] = [101] * validator.SAMPLE_COUNT
            elapsed["sample_order"] = list(reversed(range(validator.SAMPLE_COUNT)))
            elapsed["min"] = elapsed["p50"] = elapsed["p95"] = elapsed["p99"] = elapsed["max"] = 101
            elapsed["mean"] = 101.0
            elapsed["standard_deviation"] = 0.0
            elapsed["confidence_interval_95"]["lower"] = elapsed["confidence_interval_95"]["upper"] = 101.0
            with self.assertRaisesRegex(validator.ValidationError, "stable by"):
                validator._validate_elapsed(elapsed, "tie-test")
            paths["A1"].write_text(json.dumps(self.reports["A1"]).replace('"mean": 1249.5', '"mean": NaN'), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "non-finite JSON number"):
                validator.validate_paths(paths["A1"], paths["B1"], paths["B2"], paths["A2"], contract_path)

    def test_rejects_contract_corpus_semantic_binary_and_forbidden_claim_mutations(self) -> None:
        contract = copy.deepcopy(self.contract)
        contract["corpora"][1]["target_payload_sha256"] = "0" * 64
        with self.assertRaisesRegex(validator.ValidationError, "retained fixed identity"):
            self._validate(contract=contract)
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][2]["source"]["opc_serial_eager_open"]["observed_part_counts"][0] = True
        with self.assertRaisesRegex(validator.ValidationError, "observed_part_counts"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["B1"]["tool"]["instrumentation"] = "system_allocator_operation_scoped"
        with self.assertRaisesRegex(validator.ValidationError, "instrumentation"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["A2"]["results"][0]["operation_metrics"]["allocation"]["status"] = "measured"
        with self.assertRaisesRegex(validator.ValidationError, "allocation.status"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["B2"]["configuration"]["filesystem_cache_states"] = ["warm", "cold-requested"]
        with self.assertRaisesRegex(validator.ValidationError, "filesystem_cache_states"):
            self._validate(reports=reports)

    def test_direct_contract_revalidates_legs_binary_metadata_and_environment(self) -> None:
        direct = self._direct_contract()
        projection = self._validate_direct(direct)
        self.assertEqual(projection["validation"]["report_count"], 4)

        legs = copy.deepcopy(direct.legs)
        legs["A1"]["implementation"] = "candidate"
        with self.assertRaisesRegex(validator.ValidationError, "implementation"):
            self._validate_direct(replace(direct, legs=legs))

        legs = copy.deepcopy(direct.legs)
        legs["A1"]["revision"] = CANDIDATE_REVISION
        with self.assertRaisesRegex(validator.ValidationError, "revision"):
            self._validate_direct(replace(direct, legs=legs))

        legs = copy.deepcopy(direct.legs)
        legs["B1"]["binary_sha256"] = _digest("3")
        with self.assertRaisesRegex(validator.ValidationError, "binary_sha256"):
            self._validate_direct(replace(direct, legs=legs))

        legs = copy.deepcopy(direct.legs)
        legs["B1"]["binary_bytes"] = 0
        with self.assertRaisesRegex(validator.ValidationError, "binary_bytes"):
            self._validate_direct(replace(direct, legs=legs))

        legs = copy.deepcopy(direct.legs)
        legs["B1"]["mode_bits"] = 0o10000
        with self.assertRaisesRegex(validator.ValidationError, "mode_bits"):
            self._validate_direct(replace(direct, legs=legs))

        environment = copy.deepcopy(direct.environment)
        environment["cpu_affinity"] = "3"
        with self.assertRaisesRegex(validator.ValidationError, "cpu_affinity"):
            self._validate_direct(replace(direct, environment=environment))

        environment = copy.deepcopy(direct.environment)
        environment["allocator"] = "jemalloc"
        with self.assertRaisesRegex(validator.ValidationError, "allocator"):
            self._validate_direct(replace(direct, environment=environment))

    def test_direct_contract_rejects_malformed_recursive_types(self) -> None:
        direct = self._direct_contract()
        mutations = (
            ("raw_sha256", replace(direct, raw_sha256=True), "raw_sha256"),
            ("tool", replace(direct, tool=[]), "contract.tool"),
            ("environment", replace(direct, environment=[]), "contract.environment"),
            ("legs", replace(direct, legs=[]), "contract.legs"),
            ("corpora", replace(direct, corpora=[]), "contract.corpora"),
        )
        for name, forged, message in mutations:
            with self.subTest(name=name):
                with self.assertRaisesRegex(validator.ValidationError, message):
                    self._validate_direct(forged)

        legs = copy.deepcopy(direct.legs)
        legs["A1"]["binary_bytes"] = True
        with self.assertRaisesRegex(validator.ValidationError, "binary_bytes"):
            self._validate_direct(replace(direct, legs=legs))

    def test_rejects_roles_paths_and_recursive_type_drift(self) -> None:
        reports = copy.deepcopy(self.reports)
        reports["A2"]["binary_identity"]["binary_sha256"] = "0" * 64
        with self.assertRaisesRegex(validator.ValidationError, "binary_sha256"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        reports["A1"]["results"][0]["execution"]["worker_count"] = True
        with self.assertRaisesRegex(validator.ValidationError, "worker_count"):
            self._validate(reports=reports)
        reports = copy.deepcopy(self.reports)
        with tempfile.TemporaryDirectory(prefix="litchi-0403-normal-latency-validator-") as directory:
            root = Path(directory)
            contract_path = root / "contract.json"
            contract_path.write_text(json.dumps(self.contract), encoding="utf-8")
            paths = {}
            for role in ROLES:
                paths[role] = root / f"{role.lower()}.json"
                paths[role].write_text(json.dumps(reports[role]), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "distinct report paths"):
                validator.validate_paths(paths["A1"], paths["A1"], paths["B2"], paths["A2"], contract_path)

    def test_cli_accepts_explicit_identity_contract(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-0403-normal-latency-validator-") as directory:
            root = Path(directory)
            paths = {}
            for role in ROLES:
                paths[role] = root / f"{role.lower()}.json"
                paths[role].write_text(json.dumps(self.reports[role]), encoding="utf-8")
            args = [
                "--a1", str(paths["A1"]), "--b1", str(paths["B1"]),
                "--b2", str(paths["B2"]), "--a2", str(paths["A2"]),
                "--control-revision", CONTROL_REVISION, "--candidate-revision", CANDIDATE_REVISION,
                "--control-binary-sha256", _digest("1"), "--candidate-binary-sha256", _digest("2"),
                "--control-binary-bytes", "1001", "--candidate-binary-bytes", "1002",
                "--control-mode-bits", "509", "--candidate-mode-bits", "509",
                "--rustc-version", "rustc 1.98.1 (test)", "--logical-cpus", "32", "--cpu-model", "test-cpu",
            ]
            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                status = validator.main(args)
            self.assertEqual(status, 0)
            self.assertIn('"validator":"litchi-opc-serial-eager-latency-abba"', stdout.getvalue())


if __name__ == "__main__":
    unittest.main()
