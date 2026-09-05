"""Focused corruption tests for :mod:`summarize_crud_baseline`."""

from __future__ import annotations

import copy
import json
import shutil
import tempfile
import unittest
from pathlib import Path

from tools import summarize_crud_baseline as summarizer
from tools.generate_corpus_manifest_v2 import generate


ROOT = Path(__file__).resolve().parents[1]
V1 = ROOT / "docs/performance/results/perf-regression-default-manifest-v1.json"


class CrudBaselineSummaryTests(unittest.TestCase):
    def setUp(self) -> None:
        source = json.loads(V1.read_text(encoding="utf-8"))
        case = sorted(source["case_corpora"])[0]
        corpus_name = source["case_corpora"][case][0]
        self.case = case
        self.corpus = source["corpora"][corpus_name]
        self.catalog = generate(
            {**source, "case_corpora": {case: [corpus_name]}}, revision="revision-test"
        )

    def _stats(self, values: list[int]) -> dict:
        order = list(range(len(values)))
        stats = summarizer._rust_statistics(sorted(values), order)
        stats["samples"] = sorted(values)
        return stats

    def _report(self, *, phase: str, values: list[int], digest: str) -> dict:
        tool = {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": "litchi-perf-baseline",
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": "system_allocator_operation_scoped"
            if phase == "allocator"
            else "none",
        }
        binary_sha = "b" * 64 if phase == "allocator" else "a" * 64
        elapsed = self._stats(values)
        binary_path = (
            "/tmp/litchi-goal-0417-normal"
            if phase == "normal"
            else "/tmp/litchi-goal-0417-allocator"
        )
        return {
            "schema_version": 1,
            "tool": tool,
            "binary_identity": {
                "path": binary_path,
                "binary_sha256": binary_sha,
                "binary_bytes": 100,
                "mode_bits": 0o755,
                "executable": True,
                "profile": "release",
            },
            "environment": {
                "git_revision": "revision-test",
                "git_worktree_dirty": False,
            },
            "configuration": {
                "samples_per_case": len(values),
                "warmup_iterations_per_case": 2 if phase == "normal" else 1,
                "cases": [self.case],
            },
            "results": [
                {
                    "case": self.case,
                    "corpus": copy.deepcopy(self.corpus),
                    "elapsed_ns": elapsed,
                    "output_sha256": digest,
                }
            ],
            "corpus_catalog": {
                key: self.catalog[key]
                for key in (
                    "manifest_version",
                    "catalog_id",
                    "catalog_sha256",
                    "content_set_sha256",
                )
            },
        }

    def _bundle(self, *, measured_allocator: bool = False) -> Path:
        directory = Path(tempfile.mkdtemp(prefix="crud-baseline-test-"))
        self.addCleanup(shutil.rmtree, directory, ignore_errors=True)
        runs = []
        for phase in ("normal", "allocator"):
            count = 5
            values = [100, 101, 102, 103, 104]
            for repeat in (1, 2):
                report = self._report(
                    phase=phase,
                    values=[value + (repeat - 1) for value in values],
                    digest="c" * 64,
                )
                if measured_allocator and phase == "allocator":
                    from tools.test_perf_compare import allocator_operation_metrics

                    report["results"][0]["operation_metrics"] = allocator_operation_metrics(
                        value=10 + repeat, sample_count=count
                    )
                report_name = f"{phase}-{repeat}.json"
                catalog_name = f"{phase}-{repeat}.catalog.json"
                (directory / report_name).write_text(
                    json.dumps(report), encoding="utf-8"
                )
                (directory / catalog_name).write_text(
                    json.dumps(self.catalog), encoding="utf-8"
                )
                runs.append(
                    {
                        "phase": phase,
                        "repeat": repeat,
                        "selector": self.case,
                        "report": report_name,
                        "catalog": catalog_name,
                        "argv": [
                            "/tmp/litchi-goal-0417-normal"
                            if phase == "normal"
                            else "/tmp/litchi-goal-0417-allocator",
                            "--case",
                            self.case,
                            "--samples",
                            str(count),
                            "--warmup",
                            "2" if phase == "normal" else "1",
                            "--json",
                            str((directory / report_name).resolve()),
                            "--corpus-manifest",
                            str((directory / catalog_name).resolve()),
                        ],
                        "exit_code": 0,
                    }
                )
        (directory / "capture.json").write_text(
            json.dumps(
                {
                    "revision": "revision-test",
                    "binaries": {
                        "normal": {
                            "path": "/tmp/litchi-goal-0417-normal",
                            "sha256": "a" * 64,
                            "bytes": 100,
                        },
                        "allocator": {
                            "path": "/tmp/litchi-goal-0417-allocator",
                            "sha256": "b" * 64,
                            "bytes": 100,
                        },
                    },
                    "runs": runs,
                }
            ),
            encoding="utf-8",
        )
        return directory

    def test_valid_phases_and_unavailable_allocator_are_reported(self) -> None:
        directory = self._bundle()
        result = summarizer.summarize(
            directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
        )
        self.assertEqual(len(result["normal"]), 1)
        self.assertEqual(result["allocator"][0]["allocation_delta"]["status"], "unavailable")
        self.assertEqual(
            result["allocator"][0]["latency_comparison"],
            "excluded_for_allocator_instrumentation",
        )

    def test_measured_allocator_delta_is_summarized(self) -> None:
        directory = self._bundle(measured_allocator=True)
        result = summarizer.summarize(
            directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
        )
        allocation = result["allocator"][0]["allocation_delta"]
        self.assertEqual(allocation["status"], "measured")
        self.assertEqual(allocation["metrics"]["allocation_calls"]["delta"]["p50"], 1)

    def test_copied_bundle_replays_with_historical_absolute_argv(self) -> None:
        directory = self._bundle()
        replay = directory.parent / f"{directory.name}-replay"
        shutil.copytree(directory, replay)
        self.addCleanup(shutil.rmtree, replay, ignore_errors=True)
        result = summarizer.summarize(
            replay, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
        )
        self.assertEqual(len(result["normal"]), 1)

    def test_even_percentiles_use_documented_producer_convention(self) -> None:
        ordered = [100, 101, 104, 105]
        self.assertEqual(summarizer._midpoint(ordered[1], ordered[2]), 102)
        self.assertEqual(summarizer._nearest_rank(ordered, 95), 105)
        self.assertEqual(summarizer._nearest_rank(ordered, 99), 105)

    def test_elapsed_stat_tampering_is_refused(self) -> None:
        directory = self._bundle()
        report_path = directory / "normal-1.json"
        report = json.loads(report_path.read_text(encoding="utf-8"))
        report["results"][0]["elapsed_ns"]["p50"] += 1
        report_path.write_text(json.dumps(report), encoding="utf-8")
        with self.assertRaises(summarizer.SummaryError):
            summarizer.summarize(
                directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
            )

    def test_oracle_repeat_mismatch_is_refused(self) -> None:
        directory = self._bundle()
        report_path = directory / "normal-2.json"
        report = json.loads(report_path.read_text(encoding="utf-8"))
        report["results"][0]["output_sha256"] = "e" * 64
        report_path.write_text(json.dumps(report), encoding="utf-8")
        with self.assertRaises(summarizer.SummaryError):
            summarizer.summarize(
                directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
            )

    def test_phase_binary_argv_mismatch_is_refused(self) -> None:
        directory = self._bundle()
        capture_path = directory / "capture.json"
        capture = json.loads(capture_path.read_text(encoding="utf-8"))
        capture["runs"][0]["argv"][0] = "/tmp/other-binary"
        capture_path.write_text(json.dumps(capture), encoding="utf-8")
        with self.assertRaises(summarizer.SummaryError):
            summarizer.summarize(
                directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
            )

    def test_boolean_elapsed_scalar_is_refused(self) -> None:
        directory = self._bundle()
        report_path = directory / "normal-1.json"
        report = json.loads(report_path.read_text(encoding="utf-8"))
        report["results"][0]["elapsed_ns"]["p50"] = True
        report_path.write_text(json.dumps(report), encoding="utf-8")
        with self.assertRaises(summarizer.SummaryError):
            summarizer.summarize(
                directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
            )

    def test_missing_repeat_is_refused(self) -> None:
        directory = self._bundle()
        capture_path = directory / "capture.json"
        capture = json.loads(capture_path.read_text(encoding="utf-8"))
        capture["runs"] = [
            run
            for run in capture["runs"]
            if not (run["phase"] == "normal" and run["repeat"] == 2)
        ]
        capture_path.write_text(json.dumps(capture), encoding="utf-8")
        with self.assertRaises(summarizer.SummaryError):
            summarizer.summarize(
                directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
            )

    def test_catalog_tampering_is_refused(self) -> None:
        directory = self._bundle()
        catalog_path = directory / "normal-1.catalog.json"
        catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
        catalog["corpora"][0]["name"] = "tampered"
        catalog_path.write_text(json.dumps(catalog), encoding="utf-8")
        with self.assertRaises(summarizer.SummaryError):
            summarizer.summarize(
                directory, samples=5, warmups=2, allocation_samples=5, allocation_warmups=1
            )


if __name__ == "__main__":
    unittest.main()
