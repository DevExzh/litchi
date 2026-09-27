"""Focused tests for the manifest-backed default performance matrix check."""

from __future__ import annotations

import copy
import json
import tempfile
import unittest
from pathlib import Path

from tools import validate_perf_default_matrix as matrix


ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = ROOT / "docs/performance/results/perf-regression-default-manifest-v1.json"


class DefaultPerformanceMatrixTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = matrix.load_manifest(MANIFEST_PATH)

    def _report(
        self,
        keys: set[tuple[str, str]],
        *,
        configuration: dict[str, object],
        samples: int,
    ) -> dict[str, object]:
        report_configuration = copy.deepcopy(configuration)
        report_configuration["samples_per_case"] = samples
        for field in matrix.IDENTITY_FIXED_FIELDS:
            report_configuration[field] = copy.deepcopy(
                self.manifest["identity_configuration"][field]
            )
        results = []
        for case, corpus_json in sorted(keys):
            results.append(
                {
                    "case": case,
                    "corpus": json.loads(corpus_json),
                    "elapsed_ns": {
                        "unit": "ns",
                        "samples": [1] * samples,
                        "sample_order": list(range(samples)),
                    },
                }
            )
        return {
            "schema_version": self.manifest["report_schema_version"],
            "configuration": report_configuration,
            "results": results,
        }

    def _write_report(self, report: dict[str, object]) -> Path:
        handle = tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", suffix=".json", delete=False
        )
        with handle:
            json.dump(report, handle)
        self.addCleanup(lambda: Path(handle.name).unlink(missing_ok=True))
        return Path(handle.name)

    def test_manifest_derives_current_full_and_tiny_matrices(self) -> None:
        full = matrix.expected_keys(self.manifest, mode="full")
        smoke = matrix.expected_keys(
            self.manifest,
            mode="smoke",
            shape="tiny",
            payload="compressible",
        )
        self.assertEqual(len(full), self.manifest["result_count"])
        self.assertEqual(
            {case for case, _ in full},
            set(self.manifest["default_cases"]),
        )
        self.assertEqual(matrix._result_key_digest(full), self.manifest["result_keys_sha256"])
        self.assertEqual(len(smoke), len(self.manifest["default_cases"]))
        self.assertTrue(smoke <= full)

    def test_full_report_rejects_a_truncated_legacy_matrix(self) -> None:
        full = matrix.expected_keys(self.manifest, mode="full")
        truncated = set(sorted(full)[:-1])
        report = self._report(
            truncated,
            configuration={
                "cases": self.manifest["default_cases"],
                **self.manifest["identity_configuration"],
            },
            samples=15,
        )
        with self.assertRaises(matrix.MatrixValidationError):
            matrix.validate_report(
                self._write_report(report),
                self.manifest,
                mode="full",
                samples=15,
            )

    def test_smoke_report_rejects_a_wrong_payload_key(self) -> None:
        smoke = matrix.expected_keys(
            self.manifest,
            mode="smoke",
            shape="tiny",
            payload="compressible",
        )
        case, corpus_json = sorted(smoke)[0]
        corpus = json.loads(corpus_json)
        corpus["payload_kind"] = "tampered"
        wrong_key = (case, matrix._canonical_json(corpus))
        report = self._report(
            (smoke - {(case, corpus_json)}) | {wrong_key},
            configuration={
                "cases": self.manifest["default_cases"],
                "corpus_shapes": ["tiny"],
                "payload_kinds": ["compressible"],
                "writer_shapes": ["tiny"],
                "xlsx_shapes": ["tiny"],
                "semantic_shapes": ["tiny"],
            },
            samples=2,
        )
        with self.assertRaises(matrix.MatrixValidationError):
            matrix.validate_report(
                self._write_report(report),
                self.manifest,
                mode="smoke",
                samples=2,
                shape="tiny",
                payload="compressible",
            )

    def test_report_rejects_non_integer_elapsed_samples(self) -> None:
        smoke = matrix.expected_keys(
            self.manifest,
            mode="smoke",
            shape="tiny",
            payload="compressible",
        )
        report = self._report(
            smoke,
            configuration={
                "cases": self.manifest["default_cases"],
                "corpus_shapes": ["tiny"],
                "payload_kinds": ["compressible"],
                "writer_shapes": ["tiny"],
                "xlsx_shapes": ["tiny"],
                "semantic_shapes": ["tiny"],
            },
            samples=2,
        )
        for malformed in (True, float("nan"), -1):
            with self.subTest(malformed=malformed):
                candidate = copy.deepcopy(report)
                candidate["results"][0]["elapsed_ns"]["samples"][0] = malformed
                with self.assertRaises(matrix.MatrixValidationError):
                    matrix.validate_report(
                        self._write_report(candidate),
                        self.manifest,
                        mode="smoke",
                        samples=2,
                        shape="tiny",
                        payload="compressible",
                    )

    def test_report_rejects_elapsed_and_configuration_contract_mismatches(self) -> None:
        smoke = matrix.expected_keys(
            self.manifest,
            mode="smoke",
            shape="tiny",
            payload="compressible",
        )
        report = self._report(
            smoke,
            configuration={
                "cases": self.manifest["default_cases"],
                "corpus_shapes": ["tiny"],
                "payload_kinds": ["compressible"],
                "writer_shapes": ["tiny"],
                "xlsx_shapes": ["tiny"],
                "semantic_shapes": ["tiny"],
            },
            samples=2,
        )
        mutations = {
            "elapsed unit": lambda candidate: candidate["results"][0]["elapsed_ns"].update(
                unit="us"
            ),
            "sample order": lambda candidate: candidate["results"][0]["elapsed_ns"].update(
                sample_order=[0, 0]
            ),
            "samples per case": lambda candidate: candidate["configuration"].update(
                samples_per_case=1
            ),
            "filesystem process flag": lambda candidate: candidate[
                "configuration"
            ].update(filesystem_process_isolated=False),
            "range simulation": lambda candidate: candidate[
                "configuration"
            ]["range_simulation"].update(fixed_latency_us=101),
        }
        for name, mutate in mutations.items():
            with self.subTest(mutation=name):
                candidate = copy.deepcopy(report)
                mutate(candidate)
                with self.assertRaises(matrix.MatrixValidationError):
                    matrix.validate_report(
                        self._write_report(candidate),
                        self.manifest,
                        mode="smoke",
                        samples=2,
                        shape="tiny",
                        payload="compressible",
                    )

    def test_manifest_rejects_tampered_result_count(self) -> None:
        tampered = copy.deepcopy(self.manifest)
        tampered["result_count"] -= 1
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", suffix=".json", delete=False
        ) as handle:
            json.dump(tampered, handle)
            path = Path(handle.name)
        self.addCleanup(lambda: path.unlink(missing_ok=True))
        with self.assertRaises(matrix.MatrixValidationError):
            matrix.load_manifest(path)


if __name__ == "__main__":
    unittest.main()
