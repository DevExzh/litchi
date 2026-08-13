import copy
import json
import math
import tempfile
import unittest
from pathlib import Path

from tools import perf_compare


TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
}


def policy():
    corpus = report()["results"][0]["corpus"]
    corpus_identity = json.dumps(
        corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
    )
    return {
        "schema_version": 1,
        "policy_id": "test-policy-v1",
        "minimum_samples": 5,
        "expected_result_count": 1,
        "expected_result_keys_sha256": perf_compare.result_key_manifest_sha256(
            [("opc_open", corpus_identity)]
        ),
        "required_cases": ["opc_open"],
        "require_clean_worktree": True,
        "require_distinct_revisions": True,
        "tool_identity": copy.deepcopy(TOOL),
        "build_identity_fields": [
            "rustc_version",
            "allocator",
            "rustflags",
            "cargo_build_target",
            "logical_cpus_available",
        ],
        "nullable_build_identity_fields": [],
        "expected_configuration": {"samples_per_case": 5},
        "latency_thresholds_percent": {"p50": 5.0, "p95": 8.0},
        "metric_classes": [
            {
                "name": "allocation",
                "max_regression_percent": 3.0,
                "path_globs": ["**/allocation_calls"],
            },
            {
                "name": "rss",
                "max_regression_percent": 4.0,
                "path_globs": ["**/*rss_bytes"],
            },
            {
                "name": "work",
                "max_regression_percent": 5.0,
                "path_globs": ["source/read_calls", "**/instructions"],
            },
        ],
    }


def report(value=100, revision="baseline", corpus_sha="abc"):
    samples = [value] * 5
    return {
        "schema_version": 1,
        "tool": copy.deepcopy(TOOL),
        "environment": {
            "rustc_version": "rustc 1.95.0",
            "git_revision": revision,
            "git_worktree_dirty": False,
            "logical_cpus_available": 2,
            "allocator": "system",
            "rustflags": "-C target-cpu=x86-64-v3",
            "cargo_build_target": "x86_64-unknown-linux-gnu",
            "os": "linux",
            "kernel": "Linux 6.12",
            "cpu_model": "Test CPU",
            "total_memory_bytes": 16_000_000_000,
            "page_size_bytes": 4096,
        },
        "configuration": {
            "samples_per_case": 5,
            "cases": ["opc_open"],
            "execution_workers": [1, 2],
        },
        "results": [
            {
                "case": "opc_open",
                "corpus": {
                    "name": "opc-fixed",
                    "generator": "opc-v1",
                    "shape": "tiny",
                    "payload_kind": "compressible",
                    "archive_sha256": corpus_sha,
                },
                "elapsed_ns": {
                    "unit": "ns",
                    "samples": samples,
                    "p50": value,
                    "p95": value,
                },
                "metrics": {
                    "allocation_calls": [1000] * 5,
                    "peak_rss_bytes": [10_000] * 5,
                    "instructions": [50_000] * 5,
                },
                "source": {"read_calls": [10] * 5},
            }
        ],
    }


class PerfCompareTests(unittest.TestCase):
    def test_checked_policy_pins_identity_only_default_manifest(self):
        repository = Path(__file__).resolve().parents[1]
        checked_policy = json.loads(
            (repository / "docs/performance/perf-regression-policy-v1.json").read_text(
                encoding="utf-8"
            )
        )
        manifest = json.loads(
            (
                repository
                / "docs/performance/results/perf-regression-default-manifest-v1.json"
            ).read_text(encoding="utf-8")
        )

        self.assertEqual(manifest["manifest_kind"], "case-corpus-key-identity")
        self.assertEqual(manifest["result_count"], 198)
        self.assertEqual(manifest["case_count"], 36)
        self.assertEqual(manifest["source_report_samples_per_case"], 1)
        self.assertEqual(manifest["source_report_warmup_iterations_per_case"], 0)
        self.assertEqual(manifest["default_cases"], checked_policy["required_cases"])
        for field, expected in manifest["identity_configuration"].items():
            self.assertEqual(
                checked_policy["expected_configuration"][field], expected, field
            )

        keys = []
        case_corpora = manifest["case_corpora"]
        corpora = manifest["corpora"]
        self.assertEqual(set(case_corpora), set(checked_policy["required_cases"]))
        for case in manifest["default_cases"]:
            names = case_corpora[case]
            self.assertEqual(len(names), len(set(names)), case)
            for name in names:
                corpus = corpora[name]
                corpus_identity = json.dumps(
                    corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
                )
                keys.append((case, corpus_identity))

        self.assertEqual(len(keys), manifest["result_count"])
        self.assertEqual(len(keys), len(set(keys)))
        self.assertEqual(
            set(corpora), {name for names in case_corpora.values() for name in names}
        )
        self.assertEqual(
            sorted(len(names) for names in case_corpora.values()),
            [3] * 18 + [8] * 18,
        )
        for forbidden in ("elapsed_ns", "metrics", "sink", "output_sha256"):
            self.assertNotIn(forbidden, manifest)
        digest = perf_compare.result_key_manifest_sha256(keys)
        self.assertEqual(digest, manifest["result_keys_sha256"])
        self.assertEqual(digest, checked_policy["expected_result_keys_sha256"])

    def test_pass_compares_latency_and_available_resource_counters(self):
        baseline = report()
        current = report(104, "current")
        current["results"][0]["metrics"]["allocation_calls"] = [1020] * 5
        current["results"][0]["metrics"]["peak_rss_bytes"] = [10_300] * 5
        current["results"][0]["metrics"]["instructions"] = [52_000] * 5
        result = perf_compare.compare_reports(baseline, current, policy())
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["matched_results"], 1)
        self.assertEqual(result["summary"]["compared_metrics"], 6)
        self.assertFalse(result["regressions"])

    def test_p50_and_p95_regressions_are_reported(self):
        result = perf_compare.compare_reports(report(), report(120, "current"), policy())
        self.assertEqual(result["status"], "regression")
        metrics = {item["metric"] for item in result["regressions"]}
        self.assertEqual(metrics, {"elapsed_ns.p50", "elapsed_ns.p95"})

    def test_optional_resource_regression_is_reported(self):
        current = report(revision="current")
        current["results"][0]["metrics"]["allocation_calls"] = [1100] * 5
        result = perf_compare.compare_reports(report(), current, policy())
        self.assertEqual(result["status"], "regression")
        regression = result["regressions"][0]
        self.assertEqual(regression["metric_class"], "allocation")
        self.assertEqual(regression["metric"], "metrics.allocation_calls")

    def test_zero_baseline_metric_fails_on_nonzero_current(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["source"]["read_calls"] = [0] * 5
        current["results"][0]["source"]["read_calls"] = [1] * 5
        result = perf_compare.compare_reports(baseline, current, policy())
        regression = next(
            item for item in result["regressions"] if item["metric"] == "source.read_calls"
        )
        self.assertTrue(regression["delta_is_infinite"])
        self.assertIsNone(regression["delta_percent"])

    def test_schema_tool_and_build_identity_mismatches_fail_closed(self):
        mutations = {
            "schema": lambda item: item.__setitem__("schema_version", 2),
            "tool": lambda item: item["tool"].__setitem__("version", "9.9.9"),
            "build": lambda item: item["environment"].__setitem__(
                "allocator", "jemalloc"
            ),
            "dirty": lambda item: item["environment"].__setitem__(
                "git_worktree_dirty", True
            ),
        }
        for name, mutate in mutations.items():
            with self.subTest(name=name):
                current = report(revision="current")
                mutate(current)
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(report(), current, policy())

    def test_case_and_corpus_key_mismatches_fail_closed(self):
        current = report(revision="current", corpus_sha="changed")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "case/corpus key mismatch"
        ):
            perf_compare.compare_reports(report(), current, policy())

    def test_same_sided_corpus_replacement_fails_policy_manifest(self):
        baseline = report(corpus_sha="changed")
        current = report(revision="current", corpus_sha="changed")
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "manifest digest"):
            perf_compare.compare_reports(baseline, current, policy())
        current = report(revision="current")
        current["results"][0]["case"] = "other_case"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "case/corpus key mismatch"
        ):
            perf_compare.compare_reports(report(), current, policy())

    def test_duplicate_and_missing_result_keys_fail_closed(self):
        duplicate_policy = policy()
        duplicate_policy["expected_result_count"] = 2
        baseline = report()
        current = report(revision="current")
        baseline["results"].append(copy.deepcopy(baseline["results"][0]))
        current["results"].append(copy.deepcopy(current["results"][0]))
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "duplicate"):
            perf_compare.compare_reports(baseline, current, duplicate_policy)

        current = report(revision="current")
        current["results"] = []
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "expected 1"):
            perf_compare.compare_reports(report(), current, policy())

    def test_extra_same_sided_case_fails_exact_policy_set(self):
        comparison_policy = policy()
        comparison_policy["expected_result_count"] = 2
        baseline = report()
        current = report(revision="current")
        for item in (baseline, current):
            rogue = copy.deepcopy(item["results"][0])
            rogue["case"] = "rogue_case"
            rogue["corpus"]["archive_sha256"] = "rogue"
            item["results"].append(rogue)
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "case set"):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_insufficient_samples_and_invalid_numbers_fail_closed(self):
        current = report(revision="current")
        current["results"][0]["elapsed_ns"]["samples"] = [100, 100]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "minimum is 5"):
            perf_compare.compare_reports(report(), current, policy())

        current = report(revision="current")
        current["results"][0]["metrics"]["allocation_calls"] = [1, 2]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "minimum is 5"):
            perf_compare.compare_reports(report(), current, policy())

        for invalid in (math.nan, math.inf, -1):
            with self.subTest(invalid=invalid):
                current = report(revision="current")
                current["results"][0]["elapsed_ns"]["samples"][0] = invalid
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(report(), current, policy())

        for invalid in ([], "bad", [1, "bad", 3, 4, 5]):
            with self.subTest(optional_metric=invalid):
                baseline = report()
                current = report(revision="current")
                baseline["results"][0]["source"]["read_calls"] = invalid
                current["results"][0]["source"]["read_calls"] = invalid
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(baseline, current, policy())

        current = report(revision="current")
        current["results"][0]["elapsed_ns"]["samples"][0] = 10**10000
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "finite range"):
            perf_compare.compare_reports(report(), current, policy())

    def test_reported_percentile_must_agree_with_samples(self):
        current = report(revision="current")
        current["results"][0]["elapsed_ns"]["p95"] = 999
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "disagrees"):
            perf_compare.compare_reports(report(), current, policy())

    def test_optional_metric_presence_must_match(self):
        current = report(revision="current")
        del current["results"][0]["metrics"]["peak_rss_bytes"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "optional metric mismatch"
        ):
            perf_compare.compare_reports(report(), current, policy())

    def test_configuration_must_match(self):
        current = report(revision="current")
        current["configuration"]["samples_per_case"] = 6
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "configuration"):
            perf_compare.compare_reports(report(), current, policy())

        current = report(revision="current")
        current["configuration"]["cases"] = ["rogue_case"]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "exact policy"):
            perf_compare.compare_reports(report(), current, policy())

        baseline = report()
        current = report(revision="current")
        baseline["configuration"]["execution_workers"] = [1]
        current["configuration"]["execution_workers"] = [1]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "derived default"):
            perf_compare.compare_reports(baseline, current, policy())

    def test_unapproved_corpus_manifest_fails_closed(self):
        comparison_policy = policy()
        comparison_policy["expected_result_keys_sha256"] = None
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "no approved"):
            perf_compare.compare_reports(
                report(), report(revision="current"), comparison_policy
            )

    def test_build_identity_values_are_typed_and_nullable_only_by_policy(self):
        current = report(revision="current")
        current["environment"]["allocator"] = None
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "must not be null"):
            perf_compare.compare_reports(report(), current, policy())

        nullable_policy = policy()
        nullable_policy["nullable_build_identity_fields"] = ["rustflags"]
        baseline = report()
        current = report(revision="current")
        baseline["environment"]["rustflags"] = None
        current["environment"]["rustflags"] = None
        result = perf_compare.compare_reports(baseline, current, nullable_policy)
        self.assertEqual(result["status"], "pass")

    def test_reference_and_current_revisions_must_differ(self):
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "must differ"):
            perf_compare.compare_reports(report(), report(), policy())

    def test_cli_emits_machine_and_human_reports_and_exit_codes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            policy_path = root / "policy.json"
            baseline_path = root / "baseline.json"
            current_path = root / "current.json"
            json_path = root / "comparison.json"
            text_path = root / "comparison.txt"
            policy_path.write_text(json.dumps(policy()), encoding="utf-8")
            baseline_path.write_text(json.dumps(report()), encoding="utf-8")
            current_path.write_text(
                json.dumps(report(104, "current")), encoding="utf-8"
            )
            args = [
                "--policy",
                str(policy_path),
                "--baseline",
                str(baseline_path),
                "--current",
                str(current_path),
                "--json-out",
                str(json_path),
                "--summary-out",
                str(text_path),
            ]
            self.assertEqual(perf_compare.main(args), 0)
            self.assertEqual(json.loads(json_path.read_text())["status"], "pass")
            self.assertIn("PASS:", text_path.read_text())

            current_path.write_text(
                json.dumps(report(120, "current")), encoding="utf-8"
            )
            self.assertEqual(perf_compare.main(args), 1)
            self.assertEqual(json.loads(json_path.read_text())["status"], "regression")

            current_path.write_text("{invalid", encoding="utf-8")
            self.assertEqual(perf_compare.main(args), 2)
            self.assertEqual(json.loads(json_path.read_text())["status"], "invalid")
            self.assertIn("INVALID:", text_path.read_text())


if __name__ == "__main__":
    unittest.main()
