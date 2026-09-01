import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tools import xls_source_attribution_abba as abba


def identity(path: str, digest: str) -> dict[str, object]:
    return {"path": path, "bytes": 10, "sha256": digest}


def child_report(
    selector: abba.Selector,
    oracle: dict[str, object],
    *,
    revision: str = "control-revision",
    binary_path: str = "/control",
    observation: dict[str, object] | None = None,
) -> dict[str, object]:
    metrics = {name: 1 for name in abba.REQUIRED_METRICS}
    return {
        "schema_version": abba.PROFILER_SCHEMA_VERSION,
        "mode": selector.mode,
        "operation": selector.operation,
        "warmups": 20,
        "samples": 1,
        "worksheet_index": 1,
        "row": 1,
        "column": 0,
        "input": identity("/corpus.xls", "a" * 64),
        "binary": identity(binary_path, "b" * 64),
        "revision": revision,
        "tool": {"revision": revision},
        "semantic_oracle": oracle,
        "elapsed_samples_ns": [10],
        "records": [
            {
                "elapsed_ns": 10,
                "metrics": metrics,
                "observation": (
                    observation
                    if observation is not None
                    else abba.expected_observation(oracle, selector)
                ),
                "source_version_stable": True,
                "eager_phases": None,
            }
        ],
    }


class XlsSourceAttributionAbbaTests(unittest.TestCase):
    def test_percentile_stats_use_declared_no_interpolation_rules(self):
        result = abba.percentile_stats([5, 1, 3, 2])
        self.assertEqual(result["count"], 4)
        self.assertEqual(result["min"], 1)
        self.assertEqual(result["max"], 5)
        self.assertEqual(result["p50"], 2)
        self.assertEqual(result["p95"], 5)
        self.assertEqual(result["p99"], 5)
        self.assertEqual(result["mean"], 2.75)

    def test_gate_direction_accepts_lower_candidate_and_rejects_regression(self):
        lower = abba._value_comparison(1000, 980)
        self.assertEqual(lower["delta"], 20)
        self.assertGreaterEqual(lower["improvement_percent"], 1.0)
        higher = abba._value_comparison(1000, 1051)
        self.assertLess(higher["improvement_percent"], -5.0)

    def test_schema_integer_validation_rejects_bool_float_and_negative(self):
        for value in (True, 1.0, -1, abba.MAX_U64 + 1):
            with self.assertRaises(abba.DriverError):
                abba._require_nonnegative_int(value, "fixture")

    def test_observation_integer_fields_reject_bool_and_float(self):
        oracle = {
            "source_implementation_projection": {
                "worksheet_count": 2,
                "worksheet_names": ["Sheet1", "Sheet2"],
                "selected_cell": "string:4:Date",
            }
        }
        expected = abba.expected_observation(
            oracle, abba.SELECTOR_BY_KEY["file-source/open"]
        )
        for invalid in (True, 2.0):
            observation = dict(expected)
            observation["worksheet_count"] = invalid
            with self.assertRaises(abba.DriverError):
                abba.validate_observation(observation, expected, "fixture observation")

    def test_metric_allowlist_and_sample_row_cap_are_fail_closed(self):
        selector = abba.SELECTOR_BY_KEY["file-source/open"]
        group = abba.GroupAccumulator(selector, "A1", 1)
        metrics = {name: 1 for name in abba.REQUIRED_METRICS}
        metrics["unexpected"] = 1
        with self.assertRaisesRegex(abba.DriverError, "unexpected metric"):
            group.add(1, metrics)
        with self.assertRaisesRegex(abba.DriverError, "row exceeds"):
            abba.encode_sample_row(
                {"payload": "x" * (abba.MAX_NORMALIZED_ROW_BYTES + 1)}, 0
            )

    def test_oracle_mismatch_rejects_observation(self):
        oracle = {
            "source_implementation_projection": {
                "worksheet_count": 2,
                "worksheet_names": ["Sheet1", "Sheet2"],
                "selected_cell": "string:4:Date",
            }
        }
        selector = abba.SELECTOR_BY_KEY["file-source/one-cell"]
        expected = abba.expected_observation(oracle, selector)
        self.assertEqual(expected["cell"], "string:4:Date")
        metrics = {name: 1 for name in abba.REQUIRED_METRICS}
        report = {
            "schema_version": abba.PROFILER_SCHEMA_VERSION,
            "mode": selector.mode,
            "operation": selector.operation,
            "warmups": 20,
            "samples": 1,
            "worksheet_index": 1,
            "row": 1,
            "column": 0,
            "input": identity("/corpus.xls", "a" * 64),
            "binary": identity("/control", "b" * 64),
            "revision": "control-revision",
            "tool": {"revision": "control-revision"},
            "semantic_oracle": oracle,
            "elapsed_samples_ns": [10],
            "records": [
                {
                    "elapsed_ns": 10,
                    "metrics": metrics,
                    "observation": {
                        "kind": "one-cell",
                        "worksheet_count": None,
                        "worksheet_names": None,
                        "cell": "wrong",
                    },
                    "source_version_stable": True,
                    "eager_phases": None,
                }
            ],
        }
        with self.assertRaisesRegex(abba.DriverError, "observation"):
            abba.validate_child_report(
                report,
                selector=selector,
                leg="A1",
                input_identity=identity("/corpus.xls", "a" * 64),
                binary_identity=identity("/control", "b" * 64),
                warmups=20,
                semantic_oracles={selector.key: oracle},
                seen_revisions={},
            )

        report["records"][0]["observation"] = expected
        report["input"]["path"] = "/different-corpus.xls"
        with self.assertRaisesRegex(abba.DriverError, "path mismatch"):
            abba.validate_child_report(
                report,
                selector=selector,
                leg="A1",
                input_identity=identity("/corpus.xls", "a" * 64),
                binary_identity=identity("/control", "b" * 64),
                warmups=20,
                semantic_oracles={selector.key: oracle},
                seen_revisions={},
            )

        report["input"]["path"] = "/corpus.xls"
        with self.assertRaisesRegex(abba.DriverError, "revision mismatch"):
            abba.validate_child_report(
                report,
                selector=selector,
                leg="A1",
                input_identity=identity("/corpus.xls", "a" * 64),
                binary_identity=identity("/control", "b" * 64),
                warmups=20,
                semantic_oracles={selector.key: oracle},
                seen_revisions={},
                expected_revision="different-revision",
            )

    def test_distinct_a1_oracles_are_accepted_per_selector_cell(self):
        open_selector = abba.SELECTOR_BY_KEY["file-source/open"]
        list_selector = abba.SELECTOR_BY_KEY["file-source/list"]
        open_oracle = {
            "source_implementation_projection": {
                "worksheet_count": 1,
                "worksheet_names": ["Only"],
                "selected_cell": None,
            }
        }
        list_oracle = {
            "source_implementation_projection": {
                "worksheet_count": 2,
                "worksheet_names": ["First", "Second"],
                "selected_cell": None,
            }
        }
        semantic_oracles = {}
        for selector, oracle in (
            (open_selector, open_oracle),
            (list_selector, list_oracle),
        ):
            abba.validate_child_report(
                child_report(selector, oracle),
                selector=selector,
                leg="A1",
                input_identity=identity("/corpus.xls", "a" * 64),
                binary_identity=identity("/control", "b" * 64),
                warmups=20,
                semantic_oracles=semantic_oracles,
                seen_revisions={},
                expected_revision="control-revision",
            )
        self.assertEqual(
            set(semantic_oracles),
            {"file-source/open", "file-source/list"},
        )

    def test_within_cell_oracle_or_observation_mismatch_is_rejected(self):
        selector = abba.SELECTOR_BY_KEY["file-source/open"]
        oracle = {
            "source_implementation_projection": {
                "worksheet_count": 1,
                "worksheet_names": ["Only"],
                "selected_cell": None,
            }
        }
        semantic_oracles = {}
        abba.validate_child_report(
            child_report(selector, oracle),
            selector=selector,
            leg="A1",
            input_identity=identity("/corpus.xls", "a" * 64),
            binary_identity=identity("/control", "b" * 64),
            warmups=20,
            semantic_oracles=semantic_oracles,
            seen_revisions={},
            expected_revision="control-revision",
        )

        bad_observation = abba.expected_observation(oracle, selector)
        bad_observation["worksheet_count"] = 2
        with self.assertRaisesRegex(abba.DriverError, "observation"):
            abba.validate_child_report(
                child_report(selector, oracle, observation=bad_observation),
                selector=selector,
                leg="B1",
                input_identity=identity("/corpus.xls", "a" * 64),
                binary_identity=identity("/control", "b" * 64),
                warmups=20,
                semantic_oracles=semantic_oracles,
                seen_revisions={},
                expected_revision="control-revision",
            )

        changed_oracle = {
            "source_implementation_projection": {
                "worksheet_count": 2,
                "worksheet_names": ["Only"],
                "selected_cell": None,
            }
        }
        with self.assertRaisesRegex(abba.DriverError, "semantic oracle changed"):
            abba.validate_child_report(
                child_report(selector, changed_oracle),
                selector=selector,
                leg="B1",
                input_identity=identity("/corpus.xls", "a" * 64),
                binary_identity=identity("/control", "b" * 64),
                warmups=20,
                semantic_oracles=semantic_oracles,
                seen_revisions={},
                expected_revision="control-revision",
            )

    def test_missing_cell_oracle_rejects_every_non_a1_leg(self):
        selector = abba.SELECTOR_BY_KEY["file-source/list"]
        oracle = {
            "source_implementation_projection": {
                "worksheet_count": 1,
                "worksheet_names": ["Only"],
                "selected_cell": None,
            }
        }
        for leg in ("B1", "B2", "A2"):
            with self.assertRaisesRegex(abba.DriverError, "only A1 may establish"):
                abba.validate_child_report(
                    child_report(selector, oracle),
                    selector=selector,
                    leg=leg,
                    input_identity=identity("/corpus.xls", "a" * 64),
                    binary_identity=identity("/control", "b" * 64),
                    warmups=20,
                    semantic_oracles={},
                    seen_revisions={},
                    expected_revision="control-revision",
                )

    def test_test_mode_verdict_is_always_non_evidence(self):
        verdict, failures = abba.resolve_verdict([], test_mode=True)
        self.assertEqual(verdict, "rejected")
        self.assertTrue(any("non-evidence" in failure for failure in failures))

    def test_samples_publish_from_partial_name_atomically(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            partial = root / "samples.jsonl.partial"
            final = root / "samples.jsonl"
            partial.write_bytes(b"sample\n")
            self.assertFalse(final.exists())
            abba.publish_samples(partial, final)
            self.assertFalse(partial.exists())
            self.assertEqual(final.read_bytes(), b"sample\n")
            with self.assertRaisesRegex(abba.DriverError, "missing partial"):
                abba.publish_samples(root / "missing.partial", root / "other.jsonl")

    def test_partial_samples_writer_preserves_encoded_bytes(self):
        with tempfile.TemporaryDirectory() as temporary:
            partial = Path(temporary) / "samples.jsonl.partial"
            encoded = abba.encode_sample_row({"sample": 1}, 0)
            with abba.open_samples_partial(partial) as stream:
                self.assertEqual(stream.mode, "xb")
                self.assertEqual(stream.write(encoded), len(encoded))
            self.assertEqual(partial.read_bytes(), encoded)

    def test_launch_subprocess_exception_maps_to_driver_error(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            selector = abba.SELECTOR_BY_KEY["file-source/open"]
            with mock.patch.object(
                abba.subprocess,
                "Popen",
                side_effect=abba.subprocess.SubprocessError("preexec failed"),
            ) as popen:
                with self.assertRaisesRegex(abba.DriverError, "cannot launch"):
                    abba.invoke_child(
                        binary=Path("/control"),
                        input_path=Path("/corpus.xls"),
                        revision="control-revision",
                        selector=selector,
                        warmups=1,
                        tmpdir=root,
                        cwd=root,
                        cpu=None,
                        memory_limit_bytes=2 * 1024**3,
                        timeout_seconds=1.0,
                    )
                self.assertEqual(popen.call_args.kwargs["stdout"], abba.subprocess.PIPE)
                self.assertEqual(popen.call_args.kwargs["stderr"], abba.subprocess.PIPE)

    def test_child_launch_injects_leg_revision_over_ambient_environment(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            selector = abba.SELECTOR_BY_KEY["file-source/open"]
            with mock.patch.dict(
                abba.os.environ, {"LITCHI_REVISION": "ambient-revision"}, clear=False
            ):
                with mock.patch.object(
                    abba.subprocess,
                    "Popen",
                    side_effect=abba.subprocess.SubprocessError("probe launch"),
                ) as popen:
                    for binary, revision in (
                        (Path("/control"), "control-revision"),
                        (Path("/candidate"), "candidate-revision"),
                    ):
                        with self.assertRaisesRegex(abba.DriverError, "cannot launch"):
                            abba.invoke_child(
                                binary=binary,
                                input_path=Path("/corpus.xls"),
                                revision=revision,
                                selector=selector,
                                warmups=1,
                                tmpdir=root,
                                cwd=root,
                                cpu=None,
                                memory_limit_bytes=2 * 1024**3,
                                timeout_seconds=1.0,
                            )

            self.assertEqual(len(popen.call_args_list), 2)
            for call, binary, revision in zip(
                popen.call_args_list,
                (Path("/control"), Path("/candidate")),
                ("control-revision", "candidate-revision"),
            ):
                self.assertEqual(call.args[0][0], str(binary))
                self.assertEqual(call.kwargs["env"]["LITCHI_REVISION"], revision)

    def test_protocol_declares_six_selectors_and_serial_fresh_children(self):
        protocol = abba.build_protocol(
            corpus=identity("/corpus.xls", "a" * 64),
            control_binary=identity("/control", "b" * 64),
            candidate_binary=identity("/candidate", "c" * 64),
            tmpdir="/home/zhuhe/CodeProjects/.cargo-targets/change-0358/tmp",
            cpu=2,
            memory_limit_bytes=2 * 1024**3,
            timeout_seconds=120.0,
        )
        self.assertEqual(protocol["selector_order"], [selector.key for selector in abba.SELECTORS])
        self.assertEqual(protocol["collection"]["total_children"], 12000)
        self.assertTrue(protocol["collection"]["serial"])
        self.assertTrue(protocol["collection"]["fresh_child_per_sample"])
        self.assertFalse(protocol["collection"]["test_mode"])
        self.assertTrue(protocol["collection"]["evidence_eligible"])
        self.assertEqual(protocol["collection"]["warmups_per_child"], 20)
        self.assertEqual(protocol["collection"]["retained_samples_per_selector_per_leg"], 500)
        self.assertEqual(protocol["gates"]["claim_selector"], "file-source/one-cell")
        self.assertEqual(protocol["gates"]["claim_minimum_p50_and_mean_improvement_percent"], 1.0)

        with self.assertRaisesRegex(abba.DriverError, "exactly 20 warmups"):
            abba.build_protocol(
                corpus=identity("/corpus.xls", "a" * 64),
                control_binary=identity("/control", "b" * 64),
                candidate_binary=identity("/candidate", "c" * 64),
                tmpdir="/home/zhuhe/CodeProjects/.cargo-targets/change-0358/tmp",
                cpu=2,
                memory_limit_bytes=2 * 1024**3,
                timeout_seconds=120.0,
                warmups=19,
            )

        test_protocol = abba.build_protocol(
            corpus=identity("/corpus.xls", "a" * 64),
            control_binary=identity("/control", "b" * 64),
            candidate_binary=identity("/candidate", "c" * 64),
            tmpdir="/home/zhuhe/CodeProjects/.cargo-targets/change-0358/tmp",
            cpu=2,
            memory_limit_bytes=2 * 1024**3,
            timeout_seconds=120.0,
            warmups=1,
            samples=1,
            test_mode=True,
        )
        self.assertTrue(test_protocol["collection"]["test_mode"])
        self.assertFalse(test_protocol["collection"]["evidence_eligible"])


if __name__ == "__main__":
    unittest.main()
