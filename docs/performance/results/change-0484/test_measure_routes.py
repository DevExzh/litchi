"""Focused tests for the opt-in 0484 route protocol and storage receipts."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import sys
import unittest


HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import corpus_oracle  # noqa: E402
import measure_routes  # noqa: E402


def _sha(seed: str) -> str:
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()


def _proof(encoded: int = 100) -> dict:
    return {"authored": {"encoded_xml_bytes": encoded, "expected_encoded_sha256": _sha("encoded")}}


def _replay(route: str, *, encoded: int = 100) -> dict:
    histogram = {
        "bytes_0": 0,
        "bytes_1_to_512": 4,
        "bytes_513_to_4096": 0,
        "bytes_4097_to_16384": 0,
        "bytes_16385_to_65536": 0,
        "bytes_over_65536": 0,
    }
    value = {
        "route": route,
        "producer_invocations": 1,
        "store_prepare_calls": 1,
        "store_append_calls": 2,
        "store_appended_bytes": encoded,
        "store_finish_calls": 1,
        "replay_opens": 4,
        "replay_read_calls": 4,
        "replay_requested_bytes": encoded * 4,
        "replay_returned_bytes": encoded * 4,
        "replay_sha256_checks": 4,
        "replay_finish_calls": 4,
        "request_histogram": copy.deepcopy(histogram),
        "returned_histogram": copy.deepcopy(histogram),
        "retained_logical_bytes": encoded,
        "retained_capacity_bytes": None,
        "retained_capacity_provenance": None,
        "file_logical_bytes": None,
        "file_allocated_bytes": None,
        "file_write_calls": None,
        "file_sync_calls": None,
        "file_cleanup_verified": None,
        "seal_sha256_checks": 0,
        "cleanup_sha256_checks": 0,
        "durable_reference_bytes": encoded,
        "durable_reference_sha256": _sha("encoded"),
    }
    if route == "memory_store":
        value.update(
            retained_capacity_bytes=measure_routes.STORE_REPLAY_MAX_BYTES,
            retained_capacity_provenance=measure_routes.MEMORY_CAPACITY_PROVENANCE,
            durable_reference_kind="none",
            durable_reference_bytes=0,
            durable_reference_sha256=None,
        )
    else:
        value.update(
            retained_capacity_bytes=None,
            retained_logical_bytes=None,
            file_logical_bytes=encoded,
            file_allocated_bytes=4096,
            file_write_calls=1,
            file_sync_calls=1,
            file_cleanup_verified=True,
            seal_sha256_checks=1,
            cleanup_sha256_checks=1,
            durable_reference_kind="file",
            durable_reference_bytes=128,
        )
    return value


class RouteObservationTests(unittest.TestCase):
    def test_memory_store_requires_one_producer_and_four_authenticated_replays(self) -> None:
        spec = measure_routes.ROUTE_BY_NAME["memory_store"]
        measure_routes._check_replay_observation(
            {"replay": _replay("memory_store")}, spec, _proof(), "sample"
        )

        tampered = _replay("memory_store")
        tampered["retained_logical_bytes"] = 101
        tampered["durable_reference_bytes"] = 101
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation(
                {"replay": tampered}, spec, _proof(), "sample"
            )

        tampered = _replay("memory_store")
        tampered["retained_capacity_provenance"] = "configured_ceiling"
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation(
                {"replay": tampered}, spec, _proof(), "sample"
            )

        tampered = _replay("memory_store")
        tampered["file"] = {}
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation(
                {"replay": tampered}, spec, _proof(), "sample"
            )

    def test_file_store_requires_actual_length_sync_and_cleanup(self) -> None:
        spec = measure_routes.ROUTE_BY_NAME["file_store"]
        measure_routes._check_replay_observation(
            {"replay": _replay("file_store")}, spec, _proof(), "sample"
        )

        tampered = _replay("file_store")
        tampered["file_logical_bytes"] = spec.replay_max_bytes
        tampered["file_cleanup_verified"] = True
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation(
                {"replay": tampered}, spec, _proof(), "sample"
            )

        tampered = _replay("file_store")
        tampered["file_sync_calls"] = 2
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation(
                {"replay": tampered}, spec, _proof(), "sample"
            )

        tampered = _replay("file_store")
        tampered["file_cleanup_verified"] = False
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation(
                {"replay": tampered}, spec, _proof(), "sample"
            )

    def test_deterministic_route_has_no_store_observation(self) -> None:
        spec = measure_routes.ROUTE_BY_NAME["deterministic"]
        measure_routes._check_replay_observation({"replay": None}, spec, _proof(), "sample")
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation(
                {"replay": _replay("memory_store")}, spec, _proof(), "sample"
            )

    def test_route_histogram_rejects_self_consistent_counter_fabrication(self) -> None:
        spec = measure_routes.ROUTE_BY_NAME["memory_store"]
        replay = _replay("memory_store")
        replay["request_histogram"] = {
            "bytes_0": 0,
            "bytes_1_to_512": 4,
            "bytes_513_to_4096": 0,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": 0,
            "bytes_over_65536": 0,
        }
        replay["returned_histogram"] = copy.deepcopy(replay["request_histogram"])
        replay["replay_requested_bytes"] = 400
        measure_routes._check_replay_observation({"replay": replay}, spec, _proof(), "sample")
        replay["returned_histogram"] = {
            "bytes_0": 0,
            "bytes_1_to_512": 0,
            "bytes_513_to_4096": 4,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": 0,
            "bytes_over_65536": 0,
        }
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_replay_observation({"replay": replay}, spec, _proof(), "sample")


class RouteProtocolTests(unittest.TestCase):
    def test_machine_record_requires_current_recorder_and_real_observations(self):
        value = json.loads((HERE / "machine.json").read_text())
        measure_routes._validate_machine_record(value)
        edits = (
            (("schema",), "arbitrary"),
            (("driver_sha256",), "0" * 64),
            (("selected_cpu",), False),
            (("environment", "RUSTFLAGS"), "different flags"),
            (("commands", "rustc", "status"), "unavailable"),
            (("commands", "rustc", "exit_code"), False),
            (("commands", "cpu", "stdout"), "fabricated CPU inventory"),
            (("cache_policy", "cold_cache_claim"), True),
        )
        for keys, replacement in edits:
            changed = copy.deepcopy(value)
            record = changed
            for key in keys[:-1]:
                record = record[key]
            record[keys[-1]] = replacement
            with self.subTest(keys=keys), self.assertRaises(measure_routes.RouteMeasureError):
                measure_routes._validate_machine_record(changed)

    def test_inventory_and_axes_are_separate_from_default_driver(self) -> None:
        value = measure_routes._protocol_value()
        self.assertEqual(value["expected_formal_processes"], 120)
        self.assertEqual(value["expected_pilot_processes"], 60)
        self.assertEqual(len(value["formal_runs"]), 120)
        self.assertEqual(len(value["pilot_runs"]), 60)
        self.assertEqual(value["initial_axes"], {
            "input_mode": ["owned"],
            "sink_write_bytes": [4096],
            "compression": ["current"],
        })
        self.assertEqual(
            value["source_contract"],
            measure_routes.SOURCE_CONTRACTS["owned"],
        )
        self.assertEqual(value["old_deterministic_max_replay_bytes"], 1)
        self.assertEqual(value["store_replay_max_bytes"], 8 * 1024 * 1024)
        self.assertEqual(len(measure_routes.ROUTE_CASES), len(measure_routes.CASES))
        self.assertTrue(all("input_mode" in case for case in value["cases"]))
        self.assertEqual(value["separate_axis_arms"]["input"]["modes"], ["file", "short-read", "latency"])
        self.assertEqual(value["separate_axis_arms"]["sink"]["write_bytes"], [512, 4096, 65536])
        self.assertEqual(value["separate_axis_arms"]["compression"]["modes"], ["current", "store", "deflate"])
        self.assertEqual(value["expected_axis_formal_processes"], 108)
        self.assertEqual(value["expected_axis_pilot_processes"], 54)
        self.assertEqual(len(value["axis_arms"]), 27)
        self.assertEqual(len(value["axis_formal_runs"]), 108)
        self.assertEqual(len(value["axis_pilot_runs"]), 54)
        self.assertEqual(len({run["label"] for run in value["axis_formal_runs"]}), 108)
        self.assertEqual(len({run["label"] for run in value["axis_pilot_runs"]}), 54)
        self.assertEqual(len(value["missing_intersections"]), 7)
        self.assertEqual(
            {arm["axis"] for arm in value["axis_arms"]},
            {"input", "sink", "compression"},
        )
        self.assertEqual(
            {arm["value"] for arm in value["axis_arms"] if arm["axis"] == "compression"},
            {"current", "store", "deflate"},
        )
        self.assertTrue(all(run["storage_profile"]["replay_max_bytes"] == 1 for run in value["axis_pilot_runs"]))

    def test_one_factor_argv_binds_input_sink_compression_and_profiles(self) -> None:
        arm = measure_routes.AXIS_ARM_BY_LABEL["axis-input-latency-s64-a64-short-c64"]
        binary = {"path": "/bin/true", "sha256": _sha("binary"), "bytes": 1}
        argv = measure_routes._axis_argv(
            binary,
            measure_routes._axis_case(arm),
            arm,
            samples=3,
            warmups=1,
            report=Path("/var/tmp/report.json"),
            resource=Path("/var/tmp/resource.txt"),
        )
        self.assertEqual(argv[argv.index("--compression") + 1], "current")
        self.assertEqual(argv[argv.index("--sink-write") + 1], "4096")
        self.assertEqual(argv[argv.index("--input-mode") + 1], "latency")
        self.assertEqual(argv[argv.index("--input-max-range") + 1], str(measure_routes.INPUT_RANGE_BYTES))
        self.assertEqual(argv[argv.index("--input-delay-us") + 1], str(measure_routes.LATENCY_DELAY_US))
        self.assertEqual(argv[argv.index("--input-overhead-us") + 1], str(measure_routes.LATENCY_OVERHEAD_US))
        self.assertEqual(argv[argv.index("--input-bytes-per-second") + 1], str(measure_routes.LATENCY_BYTES_PER_SECOND))
        file_arm = measure_routes.AXIS_ARM_BY_LABEL["axis-input-file-s64-a64-short-c64"]
        file_argv = measure_routes._axis_argv(
            binary,
            measure_routes._axis_case(file_arm),
            file_arm,
            samples=3,
            warmups=1,
            report=Path("/var/tmp/file-report.json"),
            resource=Path("/var/tmp/file-resource.txt"),
        )
        self.assertEqual(
            file_argv[file_argv.index("--input-file") + 1],
            str((measure_routes.ROOT / file_arm["input_file"]).resolve()),
        )
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._exact_int(False, 0, "bool-zero")

    def test_independent_corpus_oracle_rejects_self_consistent_hash_tampering(self) -> None:
        expected = corpus_oracle.expected_case(64, 64, "fixed64", "short")
        observed = {
            "source_count": 64,
            "authored_count": 64,
            "chunk_mode": "fixed64",
            "text_mode": "short",
            **copy.deepcopy(expected),
        }
        observed["proof"]["authored"] = copy.deepcopy(observed["authored"])
        measure_routes._check_corpus_case(observed, "case")
        observed["authored"]["encoded_xml_bytes"] += 1
        observed["proof"]["authored"]["encoded_xml_bytes"] += 1
        with self.assertRaises(measure_routes.RouteMeasureError):
            measure_routes._check_corpus_case(observed, "case")


if __name__ == "__main__":
    unittest.main()
