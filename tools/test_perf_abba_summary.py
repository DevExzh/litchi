import copy
import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path

from tools import perf_abba_summary


TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
}


CONFIGURATION = {
    "cases": ["synthetic_case"],
    "corpus_shapes": ["medium", "tiny"],
    "samples_per_case": 5,
    "warmup_iterations_per_case": 1,
}


ENVIRONMENT = {
    "rustc_version": "rustc 1.95.0 (test)",
    "git_revision": "control-revision",
    "git_worktree_dirty": False,
    "logical_cpus_available": 1,
    "allocator": "Rust system allocator",
    "rustflags": None,
    "cargo_build_target": None,
}


def elapsed(samples):
    ordered = sorted(samples)
    return {
        "unit": "ns",
        "samples": samples,
        "p50": (
            ordered[2]
            if len(ordered) % 2
            else (ordered[len(ordered) // 2 - 1] + ordered[len(ordered) // 2]) / 2
        ),
        "mean": sum(samples) / len(samples),
        "p95": ordered[max(1, (95 * len(samples) + 99) // 100) - 1],
        "p99": ordered[max(1, (99 * len(samples) + 99) // 100) - 1],
    }


def report(rows):
    return {
        "schema_version": 1,
        "tool": copy.deepcopy(TOOL),
        "environment": copy.deepcopy(ENVIRONMENT),
        "configuration": copy.deepcopy(CONFIGURATION),
        "results": rows,
    }


def row(shape, samples, *, source=None, sink=None):
    result = {
        "case": "synthetic_case",
        "corpus": {
            "name": f"synthetic-{shape}",
            "shape": shape,
            "archive_sha256": shape * 64,
        },
        "elapsed_ns": elapsed(samples),
        "source": {"read_calls": [1] * len(samples)} if source is None else source,
        "sink": {"accepted_bytes": 10, "write_calls": 1} if sink is None else sink,
    }
    return result


def four_legs():
    shapes = {
        "tiny": {
            "a1": [100, 100, 100, 100, 100],
            "b1": [80, 80, 80, 80, 80],
            "b2": [82, 82, 82, 82, 82],
            "a2": [102, 102, 102, 102, 102],
        },
        "medium": {
            "a1": [10, 20, 30, 40, 50],
            "b1": [8, 16, 24, 32, 40],
            "b2": [9, 18, 27, 36, 45],
            "a2": [11, 22, 33, 44, 55],
        },
    }
    return [
        report([row(shape, values[label]) for shape, values in shapes.items()])
        for label in ("a1", "b1", "b2", "a2")
    ]


class PerfAbbaSummaryTests(unittest.TestCase):
    def test_recomputes_statistics_and_emits_every_multi_shape_row(self):
        summary = perf_abba_summary.summarize_reports(four_legs())
        self.assertEqual(summary["verification"]["result_count"], 2)
        self.assertEqual([item["shape"] for item in summary["results"]], ["medium", "tiny"])
        medium = summary["results"][0]
        elapsed_summary = medium["elapsed_ns"]
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["p50"], 30.0)
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["mean"], 30.0)
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["p95"], 50.0)
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["p99"], 50.0)
        self.assertAlmostEqual(
            elapsed_summary["candidate_reduction_percent"]["a1_to_b1"]["mean"], 20.0
        )
        self.assertAlmostEqual(
            elapsed_summary["same_implementation_drift_percent"]["control"]["p50"], 10.0
        )
        self.assertEqual(elapsed_summary["accepted_statistics"], ["p99"])

    def test_default_drift_ceilings_and_custom_ceilings_are_applied_per_statistic(self):
        legs = [
            report([row("tiny", values)])
            for values in (
                [100, 100, 100, 100, 100],
                [80, 80, 80, 80, 80],
                [80, 80, 80, 80, 80],
                [106, 106, 106, 106, 106],
            )
        ]
        default = perf_abba_summary.summarize_reports(legs)["results"][0]["elapsed_ns"]
        self.assertEqual(default["accepted_statistics"], ["p95", "p99"])
        self.assertIn("p50", default["rejected_statistics"])
        custom = perf_abba_summary.summarize_reports(
            legs,
            drift_ceilings={"p50": 10, "mean": 10, "p95": 10, "p99": 15},
        )["results"][0]["elapsed_ns"]
        self.assertEqual(custom["accepted_statistics"], ["p50", "mean", "p95", "p99"])

    def test_adverse_both_and_sign_disagreement_are_classified(self):
        adverse = [
            report([row("tiny", values)])
            for values in (
                [100, 100, 100, 100, 100],
                [120, 120, 120, 120, 120],
                [130, 130, 130, 130, 130],
                [110, 110, 110, 110, 110],
            )
        ]
        elapsed_summary = perf_abba_summary.summarize_reports(adverse)["results"][0][
            "elapsed_ns"
        ]
        self.assertEqual(elapsed_summary["adverse_both_statistics"], [
            "p50",
            "mean",
            "p95",
            "p99",
        ])
        self.assertTrue(
            all(
                reason.startswith("candidate is not lower in both paired directions")
                for reason in elapsed_summary["rejected_statistics"].values()
            )
        )

        mixed = [
            report([row("tiny", values)])
            for values in (
                [100, 100, 100, 100, 100],
                [80, 80, 80, 80, 80],
                [120, 120, 120, 120, 120],
                [100, 100, 100, 100, 100],
            )
        ]
        elapsed_summary = perf_abba_summary.summarize_reports(mixed)["results"][0][
            "elapsed_ns"
        ]
        self.assertEqual(elapsed_summary["adverse_both_statistics"], [])
        self.assertTrue(
            all(
                reason.startswith("paired directions disagree")
                for reason in elapsed_summary["rejected_statistics"].values()
            )
        )

    def test_environment_provenance_allows_expected_variants_and_rejects_stable_drift(self):
        legs = four_legs()
        for index, leg in enumerate(legs):
            leg["environment"]["git_revision"] = f"revision-{index}"
            leg["environment"]["git_worktree_dirty"] = bool(index % 2)
        summary = perf_abba_summary.summarize_reports(legs)
        self.assertEqual(
            [
                summary["environment"]["legs"][label]["git_revision"]
                for label in ("a1", "b1", "b2", "a2")
            ],
            ["revision-0", "revision-1", "revision-2", "revision-3"],
        )
        self.assertEqual(summary["verification"]["environment_legs_recorded"], True)

        legs = four_legs()
        legs[1]["environment"]["allocator"] = "different allocator"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "stable environment identity"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_source_and_sink_identity_mismatches_fail_closed(self):
        for field in ("source", "sink"):
            legs = four_legs()
            legs[2]["results"][0][field] = {"changed": True}
            with self.subTest(field=field), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, f"{field} identity"
            ):
                perf_abba_summary.summarize_reports(legs)

    def test_tool_configuration_and_result_identity_mismatches_fail_closed(self):
        for mutation, message in (
            (lambda legs: legs[1]["tool"].update(version="other"), "tool identity"),
            (lambda legs: legs[1]["configuration"].update(samples_per_case=6), "configuration"),
            (
                lambda legs: legs[1]["results"].__setitem__(
                    0,
                    {
                        **row("tiny", [1, 1, 1, 1, 1]),
                        "corpus": {
                            **row("tiny", [1, 1, 1, 1, 1])["corpus"],
                            "archive_sha256": "changed" * 10 + "changed"[:4],
                        },
                    },
                ),
                "case/corpus",
            ),
        ):
            legs = four_legs()
            mutation(legs)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(legs)

    def test_reported_statistics_and_sample_cardinality_are_verified(self):
        legs = four_legs()
        legs[0]["results"][0]["elapsed_ns"]["p95"] += 1
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "disagrees"):
            perf_abba_summary.summarize_reports(legs)

        legs = four_legs()
        legs[3]["results"][0]["elapsed_ns"] = elapsed([1, 2, 3, 4])
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "samples_per_case|sample counts"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_schema_mapping_and_ceiling_validation_fail_closed(self):
        legs = four_legs()
        legs[2]["schema_version"] = 2
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "schema_version"):
            perf_abba_summary.summarize_reports(legs)

        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "non-negative"):
            perf_abba_summary.summarize_reports(
                {
                    label.upper(): value
                    for label, value in zip(("a1", "b1", "b2", "a2"), four_legs())
                },
                drift_ceilings={"p50": -1, "mean": 5, "p95": 10, "p99": 15},
            )

    def test_selectors_filter_after_full_identity_validation(self):
        summary = perf_abba_summary.summarize_reports(four_legs(), shapes=["tiny"])
        self.assertEqual(len(summary["results"]), 1)
        self.assertEqual(summary["results"][0]["shape"], "tiny")
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "did not match"):
            perf_abba_summary.summarize_reports(four_legs(), cases=["missing"])

    def test_output_is_deterministic_and_cli_accepts_positional_reports(self):
        legs = four_legs()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = []
            for index, leg in enumerate(legs):
                path = root / f"leg-{index}.json"
                path.write_text(json.dumps(leg, sort_keys=False), encoding="utf-8")
                paths.append(path)
            output = root / "summary.json"
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(
                    perf_abba_summary.main(
                        [*(str(path) for path in paths), "--json-out", str(output)]
                    ),
                    0,
                )
            from_file = json.loads(output.read_text(encoding="utf-8"))
        direct = perf_abba_summary.summarize_reports(legs)
        self.assertEqual(
            json.dumps(from_file, sort_keys=True, separators=(",", ":")),
            json.dumps(direct, sort_keys=True, separators=(",", ":")),
        )


if __name__ == "__main__":
    unittest.main()
