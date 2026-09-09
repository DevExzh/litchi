"""Focused tests for the independent 0484 route analyzer."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import shutil
import sys
import tempfile
import unittest


HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import analyze_routes  # noqa: E402
import measure_routes  # noqa: E402


def _report(path: Path, samples: int = 30) -> dict:
    value = json.loads(path.read_text(encoding="utf-8"))
    value["config"]["samples"] = samples
    value["config"]["warmups"] = 3
    original = value["cases"][0]["samples"][0]
    value["cases"][0]["samples"] = []
    for index in range(samples):
        sample = copy.deepcopy(original)
        sample["sample"] = index
        sample["elapsed_ns"] += index
        value["cases"][0]["samples"].append(sample)
    return value


class AnalyzerTests(unittest.TestCase):
    def test_stats_retains_percentiles_without_fabricating_confidence(self) -> None:
        result = analyze_routes.stats([1, 2, 3, 4, 5])
        self.assertEqual(result["n"], 5)
        self.assertEqual(result["p50"], 3.0)
        self.assertIn("p95", result)
        self.assertIn("p99", result)
        self.assertNotIn("ci95_low", result)
        self.assertNotIn("ci95_high", result)

    def test_process_summary_keeps_source_authored_and_rss_metrics_separate(self) -> None:
        report_path = HERE / "route-diagnostics/dev100/deterministic-s64-a64-short-c64/report.json"
        resource_path = HERE / "route-diagnostics/dev100/deterministic-s64-a64-short-c64/resource.txt"
        expected = next(
            row for row in measure_routes._run_inventory(pilot=False)
            if row["label"] == "r1-deterministic-normal-s64-a64-short-c64"
        )
        row = analyze_routes._summarize_report(_report(report_path), resource_path, expected, "route")
        self.assertEqual(row["samples"], 30)
        for name in (
            "elapsed_ns",
            "source_throughput_bytes_per_second",
            "authored_throughput_bytes_per_second",
            "candidate_throughput_bytes_per_second",
            "time_max_rss_bytes",
        ):
            self.assertEqual(set(("n", "min", "max", "mean", "p50", "p95", "p99")), set(row["metrics"][name]))
        self.assertIsNone(row["allocator"])
        self.assertGreater(row["metrics"]["source_throughput_bytes_per_second"]["p50"], 0)
        self.assertGreater(row["metrics"]["authored_throughput_bytes_per_second"]["p50"], 0)
        self.assertEqual(row["metrics"]["time_max_rss_bytes"]["n"], 1)
        self.assertEqual(row["io"]["source"]["calls"]["n"], 30)
        self.assertEqual(row["io"]["sink"]["write_calls"]["n"], 30)
        self.assertEqual(row["io"]["authored"]["opens"]["n"], 30)
        self.assertEqual(row["io"]["source"]["request_histogram"]["kind"], "common")
        self.assertNotIn("replay", row["io"])

    def test_allocator_summary_is_separate_from_rss(self) -> None:
        report_path = HERE / "route-axis-diagnostics/dev109/axis-sink-4096-s64-a64-short-c64/report.json"
        resource_path = HERE / "route-axis-diagnostics/dev109/axis-sink-4096-s64-a64-short-c64/resource.txt"
        expected = next(
            row for row in measure_routes._axis_inventory(pilot=False)
            if row["label"] == "r1-axis-sink-4096-s64-a64-short-c64-allocator"
        )
        row = analyze_routes._summarize_report(_report(report_path), resource_path, expected, "axis")
        self.assertIsInstance(row["allocator"], dict)
        self.assertIn("region_peak_live_bytes", row["allocator"])
        self.assertIn("operation_peak_increment_bytes", row["allocator"])
        self.assertIn("time_max_rss_bytes", row["metrics"])
        self.assertNotEqual(row["allocator"]["region_peak_live_bytes"]["p50"], row["metrics"]["time_max_rss_bytes"]["p50"])

    def test_replay_io_summary_preserves_null_file_writes(self) -> None:
        report_path = HERE / "route-captures/formal1/r1-memory_store-normal-s64-a64-empty-c64/report.json"
        resource_path = HERE / "route-captures/formal1/r1-memory_store-normal-s64-a64-empty-c64/resource.txt"
        expected = next(
            row for row in measure_routes._run_inventory(pilot=False)
            if row["label"] == "r1-memory_store-normal-s64-a64-empty-c64"
        )
        row = analyze_routes._summarize_report(_report(report_path), resource_path, expected, "route")
        replay = row["io"]["replay"]
        self.assertEqual(replay["replay_opens"]["p50"], 4.0)
        self.assertEqual(replay["replay_read_calls"]["p50"], 8.0)
        self.assertEqual(replay["replay_returned_bytes"]["p50"], 32000.0)
        self.assertNotIn("file_write_calls", replay)
        self.assertEqual(replay["request_histogram"]["kind"], "common")

    def test_artifact_hash_mismatch_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            artifacts = {}
            for name, content in (("report.json", "{}"), ("resource.txt", "rss"), ("stdout.txt", ""), ("stderr.txt", "")):
                artifact = root / name
                artifact.write_text(content, encoding="utf-8")
                artifacts[name] = analyze_routes.metadata(artifact)
            artifact = root / "report.json"
            receipt = {"artifacts": artifacts}
            self.assertEqual(analyze_routes._artifact(root, receipt, "report.json"), artifact)
            artifact.write_text('{"tampered": true}', encoding="utf-8")
            with self.assertRaises(analyze_routes.AnalysisError):
                analyze_routes._artifact(root, receipt, "report.json")

    def test_canonical_route_validator_rejects_tampered_report_shape(self) -> None:
        source_dir = HERE / "route-captures/formal1/r1-deterministic-normal-s64-a64-empty-c64"
        expected = next(
            row for row in measure_routes._run_inventory(pilot=False)
            if row["label"] == "r1-deterministic-normal-s64-a64-empty-c64"
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            report = root / "report.json"
            resource = root / "resource.txt"
            shutil.copy2(source_dir / "report.json", report)
            shutil.copy2(source_dir / "resource.txt", resource)
            receipt = json.loads((source_dir / "receipt.json").read_text(encoding="utf-8"))
            binary = receipt["binary"]
            argv = measure_routes._route_argv(
                binary,
                measure_routes.ROUTE_CASE_BY_LABEL[expected["case"]],
                measure_routes.ROUTE_BY_NAME[expected["route"]],
                samples=30,
                warmups=3,
                report=report,
                resource=resource,
                replay_dir=None,
            )
            analyze_routes._canonical_report_check(report, resource, root / "receipt.json", expected, "route", binary, argv, {})
            value = json.loads(report.read_text(encoding="utf-8"))
            value["cases"][0]["samples"][0]["source_reads"]["calls"] += 1
            report.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(analyze_routes.AnalysisError):
                analyze_routes._canonical_report_check(report, resource, root / "receipt.json", expected, "route", binary, argv, {})

    def test_inventory_rejects_missing_and_extra_formal_runs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            bundle = Path(directory)
            root = bundle / "route-captures" / "formal1"
            (root / "r1-case").mkdir(parents=True)
            (root / "r1-case" / "receipt.json").write_text("{}", encoding="utf-8")
            expected = [{"label": "r1-case"}]
            analyze_routes._inventory_paths(bundle, "route", "formal1", expected)
            (root / "r2-extra").mkdir()
            with self.assertRaises(analyze_routes.AnalysisError):
                analyze_routes._inventory_paths(bundle, "route", "formal1", expected)

    def test_removed_replay_directory_requires_cleanup_manifest_and_rejects_symlink(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            receipt = root / "route-captures/formal1/r1-file_store-normal-case/receipt.json"
            receipt.parent.mkdir(parents=True)
            receipt.write_text("{}", encoding="utf-8")
            replay = receipt.parent / "replay"
            replay.mkdir()
            analyze_routes._validate_replay_directory(root, receipt, "file_store", None)
            replay.rmdir()
            relative = replay.relative_to(root).as_posix()
            manifest = {"entries": {relative: {"entry": {"path": relative}}}}
            analyze_routes._validate_replay_directory(root, receipt, "file_store", manifest)
            target = root / "target"
            target.mkdir()
            replay.symlink_to(target, target_is_directory=True)
            with self.assertRaises(analyze_routes.AnalysisError):
                analyze_routes._validate_replay_directory(root, receipt, "file_store", manifest)

    def test_protocol_script_binding_is_exact_current_helper_set(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copy2(HERE / "machine.json", root / "machine.json")
            protocol = json.loads((HERE / "route-protocol.json").read_text(encoding="utf-8"))
            protocol["scripts"]["measure_routes.py"] = "0" * 64
            path = root / "route-protocol.json"
            path.write_text(json.dumps(protocol), encoding="utf-8")
            with self.assertRaises(analyze_routes.AnalysisError):
                analyze_routes.validate_protocol(path)

    def test_role_source_mismatch_is_rejected(self) -> None:
        report_path = HERE / "route-diagnostics/dev100/deterministic-s64-a64-short-c64/report.json"
        resource_path = HERE / "route-diagnostics/dev100/deterministic-s64-a64-short-c64/resource.txt"
        expected_rows = [
            row for row in measure_routes._run_inventory(pilot=False)
            if row["label"] in {
                "r1-deterministic-normal-s64-a64-short-c64",
                "r1-deterministic-allocator-s64-a64-short-c64",
            }
        ]
        normal = analyze_routes._summarize_report(_report(report_path), resource_path, expected_rows[0], "route")
        allocator = copy.deepcopy(normal)
        allocator["role"] = "allocator"
        allocator["source_identity"]["archive_sha256"] = "0" * 64
        with self.assertRaises(analyze_routes.AnalysisError):
            analyze_routes.validate_identity_consistency([normal, allocator])

    def test_review_flag_is_descriptive_and_never_a_causal_speedup(self) -> None:
        before = {"metrics": {"elapsed_ns": {"p50": 100.0}}}
        after = {"metrics": {"elapsed_ns": {"p50": 106.0}}}
        result = analyze_routes._comparison_record(
            before,
            after,
            "elapsed_ns",
            kind="contemporaneous_route",
            control="deterministic",
            candidate="memory_store",
        )
        self.assertTrue(result["review_flag"])
        self.assertFalse(result["causal_speedup"])
        self.assertEqual(result["percent_change"], 6.0)

    def _route_identity_rows(self) -> list[dict]:
        report_path = HERE / "route-diagnostics/dev100/deterministic-s64-a64-short-c64/report.json"
        resource_path = HERE / "route-diagnostics/dev100/deterministic-s64-a64-short-c64/resource.txt"
        expected = next(
            row for row in measure_routes._run_inventory(pilot=False)
            if row["label"] == "r1-deterministic-normal-s64-a64-short-c64"
        )
        template = analyze_routes._summarize_report(_report(report_path), resource_path, expected, "route")
        rows = []
        for route in measure_routes.ROUTE_NAMES:
            for repeat in measure_routes.REPEATS:
                for role in measure_routes.ROLES:
                    row = copy.deepcopy(template)
                    row.update({
                        "identity": route,
                        "route": route,
                        "role": role,
                        "repeat": repeat,
                        "label": f"r{repeat}-{route}-{role}",
                    })
                    rows.append(row)
        return rows

    def test_identity_requires_full_candidate_identity_across_repeats_and_routes(self) -> None:
        rows = self._route_identity_rows()
        for row in rows:
            if row["route"] == "memory_store" and row["repeat"] == 2:
                row["candidate_identity"]["archive_sha256"] = "0" * 64
        with self.assertRaises(analyze_routes.AnalysisError):
            analyze_routes.validate_identity_consistency(rows)

    def test_compression_axis_allows_physical_bytes_but_requires_main_xml(self) -> None:
        rows = self._route_identity_rows()
        for role in measure_routes.ROLES:
            axis = copy.deepcopy(next(row for row in rows if row["route"] == "deterministic" and row["role"] == role and row["repeat"] == 1))
            axis.update({
                "family": "axis",
                "identity": "axis-compression-store-s64-a64-short-c64",
                "route": None,
                "axis": "compression",
                "value": "store",
                "role": role,
                "repeat": 1,
                "label": f"axis-{role}",
            })
            axis["source_identity"]["archive_bytes"] += 1
            axis["source_identity"]["archive_sha256"] = "1" * 64
            axis["candidate_identity"]["archive_bytes"] += 1
            axis["candidate_identity"]["archive_sha256"] = "2" * 64
            rows.append(axis)
        analyze_routes.validate_identity_consistency(rows)
        rows[-1]["candidate_identity"]["main_xml_sha256"] = "3" * 64
        with self.assertRaises(analyze_routes.AnalysisError):
            analyze_routes.validate_identity_consistency(rows)


if __name__ == "__main__":
    unittest.main()
