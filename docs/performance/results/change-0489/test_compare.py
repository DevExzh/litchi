#!/usr/bin/env python3
"""Read-only unit checks for the 0489 comparison helper."""

from __future__ import annotations

import copy
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock


ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import compare  # noqa: E402


class CompareHelperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.protocol = compare._protocol_value(include_freeze_source=False)

    @staticmethod
    def _retained_report(arm: dict, role: str, repeat: int = 1) -> tuple[Path, dict, dict, Path | None]:
        directory = compare.BASELINE_ROOT / "captures" / "after" / "formal1" / (
            f"after-r{repeat}-{role}-{arm['id']}"
        )
        report = directory / "report.json"
        build = compare._read(
            compare.BASELINE_ROOT / f"build-{role}.json"
        )["binary"]
        replay_dir = directory / "replay" if arm["route"] == "file_store" else None
        argv = compare._argv_for(
            arm,
            build,
            report=report,
            resource=directory / "resource.txt",
            samples=compare.FORMAL_SAMPLES,
            warmups=compare.FORMAL_WARMUPS,
            replay_dir=replay_dir,
        )
        return report, build, {"argv": argv, "directory": directory}, replay_dir

    def test_frozen_matrix_cardinality_and_order(self) -> None:
        self.assertEqual(len(compare.ARMS), 18)
        self.assertEqual(len(compare._runs(pilot=False)), 144)
        self.assertEqual(len(compare._runs(pilot=True)), 72)
        formal = compare._runs(pilot=False)
        before = [run for run in formal if run["phase"] == "before"]
        after = [run for run in formal if run["phase"] == "after"]
        self.assertEqual(len(before), 72)
        self.assertEqual(len(after), 72)
        self.assertEqual(
            [(run["arm"], run["role"]) for run in before if run["repeat"] == 2],
            list(reversed([(run["arm"], run["role"]) for run in before if run["repeat"] == 1])),
        )

    def test_all_retained_0487_report_validators_and_resource_parsers(self) -> None:
        """Exercise every selected arm with both retained 0487 roles."""

        checked = 0
        for arm in compare.ARMS:
            for role in compare.ROLE_NAMES:
                with self.subTest(arm=arm["id"], role=role):
                    report_path, binary, paths, replay_dir = self._retained_report(arm, role)
                    value = compare._validate_report(
                        report_path,
                        arm,
                        role,
                        samples=compare.FORMAL_SAMPLES,
                        warmups=compare.FORMAL_WARMUPS,
                        binary=binary,
                        argv=paths["argv"],
                        input_metadata=compare._input_metadata(self.protocol, arm),
                        replay_dir=replay_dir,
                    )
                    self.assertEqual(value["schema"], compare.old_routes.REPORT_SCHEMA)
                    self.assertEqual(len(value["cases"]), 1)
                    rss = compare._parse_resource(paths["directory"] / "resource.txt")
                    self.assertGreater(rss, 0)
                    checked += 1
        self.assertEqual(checked, 36)

    def test_metric_extraction_retains_samples_heap_and_rss_scope(self) -> None:
        for arm in compare.ARMS:
            for role in compare.ROLE_NAMES:
                with self.subTest(arm=arm["id"], role=role):
                    report_path, binary, paths, replay_dir = self._retained_report(arm, role)
                    report = compare._validate_report(
                        report_path,
                        arm,
                        role,
                        samples=compare.FORMAL_SAMPLES,
                        warmups=compare.FORMAL_WARMUPS,
                        binary=binary,
                        argv=paths["argv"],
                        input_metadata=compare._input_metadata(self.protocol, arm),
                        replay_dir=replay_dir,
                    )
                    metrics = compare._row_metrics(
                        report,
                        role,
                        compare._parse_resource(paths["directory"] / "resource.txt"),
                        str(report_path),
                    )
                    self.assertEqual(metrics["elapsed_ns"]["n"], compare.FORMAL_SAMPLES)
                    self.assertEqual(len(metrics["sample_elapsed_ns"]), compare.FORMAL_SAMPLES)
                    rss = metrics["time_max_rss_bytes"]
                    self.assertEqual(rss["n"], 1)
                    self.assertEqual(rss["observation_scope"], "whole_child_gnu_time_single_observation")
                    self.assertEqual(rss["p50"], rss["p95"])
                    if role == "allocator":
                        self.assertEqual(
                            metrics["allocator_operation_peak_increment_bytes"]["n"],
                            compare.FORMAL_SAMPLES,
                        )
                        self.assertEqual(
                            len(metrics["sample_allocator_operation_peak_increment_bytes"]),
                            compare.FORMAL_SAMPLES,
                        )
                        self.assertIn("allocation_calls", metrics["allocator_counters"])
                    else:
                        self.assertIsNone(metrics["allocator_operation_peak_increment_bytes"])

    @staticmethod
    def _synthetic_rows() -> list[dict]:
        rows: list[dict] = []
        common_source = {"archive_bytes": 10, "archive_sha256": "a" * 64, "main_xml_bytes": 11, "main_xml_sha256": "b" * 64}
        common_authored = {"encoded_xml_bytes": 12, "expected_event_sha256": "c" * 64, "expected_encoded_sha256": "d" * 64}
        common_candidate = {
            "archive_bytes": 13,
            "archive_sha256": "e" * 64,
            "main_xml_bytes": 14,
            "main_xml_sha256": "f" * 64,
            "semantic": {"paragraph_count": 2, "text_bytes": 3, "order_sha256": "1" * 64, "text_sha256": "2" * 64},
            "oracle_flags": {field: True for field in (
                "candidate_xml_exact", "candidate_semantic_exact", "untouched_member_metadata_exact",
                "untouched_raw_members_preserved", "physical_order_exact", "opaque_member_exact",
                "source_unchanged", "inverse_exact",
            )},
        }
        for phase in compare.PHASES:
            for arm in compare.ARMS:
                for role in compare.ROLE_NAMES:
                    for repeat in compare.REPEATS:
                        elapsed = 100.0 if phase == "before" else 106.0
                        rss = 1_000.0 if phase == "before" else 1_060.0
                        metrics = {
                            "elapsed_ns": {"n": 30, "p50": elapsed, "p95": elapsed, "p99": elapsed},
                            "time_max_rss_bytes": {
                                "n": 1,
                                "value": rss,
                                "p50": rss,
                                "p95": rss,
                                "p99": rss,
                                "observation_scope": "whole_child_gnu_time_single_observation",
                            },
                            "allocator_operation_peak_increment_bytes": None,
                        }
                        if role == "allocator":
                            heap = 200.0 if phase == "before" else 188.0
                            metrics["allocator_operation_peak_increment_bytes"] = {
                                "n": 30, "p50": heap, "p95": heap, "p99": heap,
                            }
                        rows.append({
                            "phase": phase,
                            "arm": arm["id"],
                            "role": role,
                            "repeat": repeat,
                            "label": f"{phase}-{arm['id']}-{role}-{repeat}",
                            "source_identity": copy.deepcopy(common_source),
                            "authored_identity": copy.deepcopy(common_authored),
                            "candidate_identity": copy.deepcopy(common_candidate),
                            "metrics": metrics,
                        })
        return rows

    def test_before_after_flags_include_rss_and_zero_baseline_is_safe(self) -> None:
        rows = self._synthetic_rows()
        comparisons = compare._comparisons(rows)
        self.assertEqual(len(comparisons), 180)
        elapsed = next(item for item in comparisons if item["metric"] == "elapsed_ns")
        rss = next(item for item in comparisons if item["metric"] == "time_max_rss_bytes")
        heap = next(item for item in comparisons if item["metric"] == "allocator_operation_peak_increment_bytes")
        self.assertTrue(elapsed["percent_changes"]["p50"]["review_flag_5pct"])
        self.assertTrue(rss["percent_changes"]["p50"]["review_flag_5pct"])
        self.assertTrue(heap["percent_changes"]["p50"]["review_flag_5pct"])
        self.assertIsNone(compare._percent_change(0.0, 0.0))
        self.assertTrue(compare.math.isinf(compare._percent_change(0.0, 1.0)))

    def test_artifact_hash_tamper_is_rejected(self) -> None:
        configured_tmp = os.environ.get("TMPDIR")
        temp_root = configured_tmp if configured_tmp and Path(configured_tmp).is_dir() else None
        with tempfile.TemporaryDirectory(dir=temp_root) as name:
            directory = Path(name)
            for filename in ("started.json", "stdout.txt", "stderr.txt", "resource.txt", "report.json"):
                (directory / filename).write_text("{}\n", encoding="utf-8")
            receipt = {"artifacts": compare._artifact_map(directory / filename for filename in (
                directory / "started.json", directory / "stdout.txt", directory / "stderr.txt",
                directory / "resource.txt", directory / "report.json",
            )), "missing_artifacts": []}
            compare._validate_artifacts(receipt, directory, file_store=False)
            (directory / "report.json").write_text("tampered\n", encoding="utf-8")
            with self.assertRaises(compare.CompareError):
                compare._validate_artifacts(receipt, directory, file_store=False)

    def test_content_identity_tamper_is_rejected(self) -> None:
        rows = self._synthetic_rows()
        original = rows[0]
        tampered = copy.deepcopy(original)
        tampered["source_identity"]["main_xml_sha256"] = "9" * 64
        with self.assertRaises(compare.CompareError):
            compare._check_content_identity([original, tampered])

    def test_protocol_loader_returns_protocol_hash_after_fixture_checks(self) -> None:
        configured_tmp = os.environ.get("TMPDIR")
        temp_root = configured_tmp if configured_tmp and Path(configured_tmp).is_dir() else None
        with tempfile.TemporaryDirectory(dir=temp_root) as name:
            protocol_path = Path(name) / "comparison-protocol.json"
            protocol = compare._protocol_value(include_freeze_source=False)
            protocol_path.write_text(json.dumps(protocol, sort_keys=True) + "\n", encoding="utf-8")
            with mock.patch.object(compare, "_protocol_path", return_value=protocol_path):
                loaded, digest = compare._load_protocol()
            self.assertEqual(loaded, protocol)
            self.assertEqual(digest, compare.sha(protocol_path))

    def test_capture_all_loads_only_the_selected_phase_builds(self) -> None:
        # The phase loader is checked before a protocol is frozen so that
        # source edits can follow the completed before lane.
        protocol = compare._protocol_value(include_freeze_source=False)
        loaded: list[tuple[str, str]] = []
        captured: list[dict] = []

        def load_build(_protocol: dict, phase: str, role: str) -> dict:
            loaded.append((phase, role))
            return {"phase": phase, "role": role, "source": {"phase": phase}}

        def capture_one(_protocol: dict, _builds: dict, run: dict, _attempt: str, *, pilot: bool, timeout_seconds: int) -> None:
            self.assertFalse(pilot)
            self.assertEqual(timeout_seconds, 1)
            captured.append(run)

        with (
            mock.patch.object(compare, "_load_protocol", return_value=(protocol, "p" * 64)),
            mock.patch.object(compare, "_load_build", side_effect=load_build),
            mock.patch.object(compare, "_capture_one", side_effect=capture_one),
        ):
            compare.capture_all("before", "unit-before", pilot=False, timeout_seconds=1)

        self.assertEqual(loaded, [("before", "normal"), ("before", "allocator")])
        self.assertEqual(len(captured), 72)
        self.assertTrue(all(run["phase"] == "before" for run in captured))

    def test_before_binding_uses_retained_0487_and_after_uses_0489(self) -> None:
        bindings = compare._protocol_value(include_freeze_source=False)["build_bindings"]
        self.assertEqual(bindings["before"]["normal"]["root"], "0487")
        self.assertEqual(bindings["before"]["normal"]["record_path"], "build-normal.json")
        self.assertEqual(bindings["after"]["normal"]["root"], "0489")
        self.assertEqual(bindings["after"]["normal"]["record_path"], "build-normal.json")
        self.assertEqual(compare._build_environment("before"), compare.baseline_environment())
        self.assertNotIn("CARGO_TARGET_DIR", compare._build_environment("before"))
        self.assertEqual(compare._build_environment("after"), compare.environment())
        self.assertEqual(compare.support.TEMP, Path("/home/zhuhe/.cache/litchi-goal-0489"))
        self.assertEqual(compare.support.TARGET_DIR, Path("/home/zhuhe/.cache/litchi-build-0489"))

    def test_source_allowlist_is_only_opc_and_focused_tests(self) -> None:
        self.assertEqual(
            compare.ALLOWED_SOURCE_PREFIXES,
            (
                "crates/litchi-opc/src/source_backed/splice.rs",
                "crates/litchi-opc/tests/source_part_splice.rs",
                "crates/litchi-opc/tests/source_part_splice_replay.rs",
            ),
        )


if __name__ == "__main__":
    unittest.main()
