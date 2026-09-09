"""Focused invariants for the 0490 variance evidence driver.

These tests exercise protocol construction and fail-closed analysis helpers;
they never build binaries or launch a capture.
"""

from __future__ import annotations

import copy
import importlib.util
from pathlib import Path
import tempfile
import unittest


HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("variance_0490_under_test", HERE / "variance.py")
assert SPEC is not None and SPEC.loader is not None
variance = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(variance)


class VarianceProtocolTests(unittest.TestCase):
    def test_formal_inventory_has_72_unique_children_and_4320_samples(self) -> None:
        runs = variance.formal_runs()
        self.assertEqual(len(runs), 72)
        self.assertEqual(len({row["label"] for row in runs}), 72)
        self.assertEqual({row["block"] for row in runs}, set(range(1, 7)))
        self.assertEqual(len(runs) * variance.FORMAL_SAMPLES, 4320)
        for block in range(1, 7):
            children = [row for row in runs if row["block"] == block]
            self.assertEqual(len(children), 12)
            self.assertEqual({row["phase"] for row in children}, {"before", "after"})
            self.assertEqual(len({(row["phase"], row["role"], row["route"]) for row in children}), 12)

    def test_phase_and_role_route_order_is_alternated(self) -> None:
        runs = variance.formal_runs()
        expected_pairs = [(role, route) for role in variance.ROLES for route in variance.ROUTES]
        for block in range(1, 7):
            children = [row for row in runs if row["block"] == block]
            expected_phases = ("before", "after") if block % 2 else ("after", "before")
            self.assertEqual([row["phase"] for row in children[:6]], [expected_phases[0]] * 6)
            self.assertEqual([row["phase"] for row in children[6:]], [expected_phases[1]] * 6)
            expected = expected_pairs if block % 2 else list(reversed(expected_pairs))
            self.assertEqual([(row["role"], row["route"]) for row in children[:6]], expected)
            self.assertEqual([(row["role"], row["route"]) for row in children[6:]], expected)

    def test_stats_and_percentile_keep_all_sixty_samples(self) -> None:
        values = list(range(1, variance.FORMAL_SAMPLES + 1))
        summary = variance._stats(values)
        self.assertEqual(summary["n"], 60)
        self.assertEqual(summary["min"], 1.0)
        self.assertEqual(summary["max"], 60.0)
        self.assertEqual(summary["p50"], 30.5)
        self.assertEqual(summary["p95"], 57.05)
        self.assertEqual(summary["p99"], 59.41)

    def test_block_bootstrap_is_deterministic_and_uses_six_blocks(self) -> None:
        values = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]
        first = variance._bootstrap(values)
        second = variance._bootstrap(values)
        self.assertEqual(first, second)
        assert first is not None
        self.assertEqual(first["n_blocks"], 6)
        self.assertEqual(first["resamples"], 10_000)
        self.assertEqual(first["seed"], 490)

    def test_inventory_validator_rejects_missing_or_extra_child(self) -> None:
        expected = variance.formal_runs()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            captures = root / "captures" / "attempt"
            for run in expected:
                path = captures / f"block-{run['block']:02d}" / run["label"]
                path.mkdir(parents=True)
            original_root = variance.ROOT
            try:
                variance.ROOT = root
                variance._validate_inventory("attempt", expected)
                (captures / "block-01" / expected[0]["label"]).rmdir()
                with self.assertRaises(variance.VarianceError):
                    variance._validate_inventory("attempt", expected)
            finally:
                variance.ROOT = original_root

    def test_sample_count_validator_rejects_retained_0489_thirty_sample_report(self) -> None:
        report_path = (
            variance.CANDIDATE_ROOT
            / "captures/after/formal1/after-r1-normal-memory_store-owned-s64-a64-short-c64/report.json"
        )
        report = variance.read_json(report_path)
        with tempfile.TemporaryDirectory() as directory:
            resource = Path(directory) / "resource.txt"
            resource.write_text("Maximum resident set size (kbytes): 1\n", encoding="utf-8")
            with self.assertRaises(variance.VarianceError):
                variance._sample_values(report, resource, "normal", "retained-0489")

    def test_oracle_and_sink_identity_mismatch_is_rejected(self) -> None:
        report_path = (
            variance.CANDIDATE_ROOT
            / "captures/after/formal1/after-r1-normal-memory_store-owned-s64-a64-short-c64/report.json"
        )
        report = variance.read_json(report_path)
        observed = copy.deepcopy(report["cases"][0])
        original = variance._identity(observed, "identity")
        observed["samples"][0]["sink"]["sha256"] = "0" * 64
        with self.assertRaises(variance.VarianceError):
            variance._sample_values(
                {"cases": [observed]},
                HERE / "missing-resource.txt",
                "normal",
                "tampered",
            )
        self.assertEqual(original["candidate"]["archive_bytes"], variance._identity(report["cases"][0], "identity")["candidate"]["archive_bytes"])

    def test_binary_binding_helper_rejects_changed_executable(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "binary"
            path.write_bytes(b"binary")
            path.chmod(0o755)
            expected = {"path": str(path), "bytes": 6, "sha256": variance.sha(path), "executable": True}
            self.assertEqual(variance.metadata(path)["sha256"], expected["sha256"])
            path.write_bytes(b"tampered")
            with self.assertRaises(variance.VarianceError):
                variance._binary_for(
                    {"builds": {"before": {"normal": {"binary": expected}}}},
                    "before",
                    "normal",
                )


if __name__ == "__main__":
    unittest.main()
