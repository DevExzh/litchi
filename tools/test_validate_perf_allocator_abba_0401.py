from __future__ import annotations

import copy
import hashlib
import json
import os
import tempfile
import unittest
from pathlib import Path

from tools.validate_perf_allocator_abba_0401 import (
    ValidationError,
    build_projection,
    validate_paths,
)


ROOT = Path(__file__).resolve().parents[1]
EVIDENCE = Path(
    os.environ.get("LITCHI_0401_EVIDENCE", "/tmp/litchi-0401-evidence.M4SZcY")
)
LEG_NAMES = ("a1", "b1", "b2", "a2")
_MISSING = object()


def _evidence_paths(root: Path) -> dict[str, Path]:
    return {leg: root / f"{leg}-allocator.json" for leg in LEG_NAMES}


@unittest.skipUnless(
    all(path.is_file() for path in _evidence_paths(EVIDENCE).values()),
    "Change 0401 allocator evidence is required",
)
class AllocatorAbba0401ValidatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.real_paths = _evidence_paths(EVIDENCE)
        cls.raw_documents = {
            leg: path.read_bytes() for leg, path in cls.real_paths.items()
        }
        cls.documents = {
            leg: json.loads(raw)
            for leg, raw in cls.raw_documents.items()
        }
        cls.base_projection = build_projection(
            cls.real_paths["a1"],
            cls.real_paths["b1"],
            cls.real_paths["b2"],
            cls.real_paths["a2"],
        )

    def _fixture(
        self,
        documents: dict[str, object] | None = None,
        projection: object = _MISSING,
    ) -> tuple[tempfile.TemporaryDirectory[str], dict[str, Path], Path | None]:
        temporary = tempfile.TemporaryDirectory(prefix="litchi-0401-allocator-validator-")
        root = Path(temporary.name)
        paths: dict[str, Path] = {}
        for leg in LEG_NAMES:
            path = root / f"{leg}.json"
            if documents is None:
                path.write_bytes(self.raw_documents[leg])
            else:
                path.write_text(json.dumps(documents[leg], indent=2) + "\n", encoding="utf-8")
            paths[leg] = path
        if projection is _MISSING:
            projection = self.base_projection
        if projection is None:
            return temporary, paths, None
        projection_path = root / "allocation-metrics.json"
        projection_path.write_text(json.dumps(projection, indent=2) + "\n", encoding="utf-8")
        return temporary, paths, projection_path

    def _validate(
        self,
        documents: dict[str, object] | None = None,
        projection: object = _MISSING,
    ) -> None:
        temporary, paths, projection_path = self._fixture(documents, projection)
        try:
            validate_paths(
                paths["a1"],
                paths["b1"],
                paths["b2"],
                paths["a2"],
                projection_path,
            )
        finally:
            temporary.cleanup()

    def test_real_reports_validate_read_only(self) -> None:
        before = {leg: raw for leg, raw in self.raw_documents.items()}
        with tempfile.TemporaryDirectory(prefix="litchi-0401-allocator-projection-") as name:
            projection_path = Path(name) / "allocation-metrics.json"
            projection_path.write_text(
                json.dumps(self.base_projection, indent=2) + "\n", encoding="utf-8"
            )
            result = validate_paths(
                self.real_paths["a1"],
                self.real_paths["b1"],
                self.real_paths["b2"],
                self.real_paths["a2"],
                projection_path,
            )
        self.assertEqual(result, {
            "reports": 4,
            "samples": 120,
            "unique_child_process_ids": 120,
            "allocator_vectors": 40,
        })
        self.assertEqual(
            before,
            {leg: path.read_bytes() for leg, path in self.real_paths.items()},
        )

    def test_projection_builder_is_deterministic(self) -> None:
        first = build_projection(
            self.real_paths["a1"],
            self.real_paths["b1"],
            self.real_paths["b2"],
            self.real_paths["a2"],
        )
        second = build_projection(
            self.real_paths["a1"],
            self.real_paths["b1"],
            self.real_paths["b2"],
            self.real_paths["a2"],
        )
        first_bytes = (json.dumps(first, indent=2) + "\n").encode()
        second_bytes = (json.dumps(second, indent=2) + "\n").encode()
        self.assertEqual(first, second)
        self.assertEqual(hashlib.sha256(first_bytes).hexdigest(), hashlib.sha256(second_bytes).hexdigest())

    def test_rejects_duplicate_json_keys(self) -> None:
        temporary, paths, projection_path = self._fixture()
        try:
            raw = paths["a1"].read_text()
            paths["a1"].write_text(raw.replace("{", '{"schema_version": 1,', 1), encoding="utf-8")
            with self.assertRaisesRegex(ValidationError, "duplicate JSON object key"):
                validate_paths(paths["a1"], paths["b1"], paths["b2"], paths["a2"], projection_path)
        finally:
            temporary.cleanup()

    def test_rejects_non_finite_json_number(self) -> None:
        temporary, paths, projection_path = self._fixture()
        try:
            raw = paths["a1"].read_text()
            paths["a1"].write_text(raw.replace('"mean": ', '"mean": NaN, "discarded_mean": ', 1), encoding="utf-8")
            with self.assertRaisesRegex(ValidationError, "non-finite JSON number"):
                validate_paths(paths["a1"], paths["b1"], paths["b2"], paths["a2"], projection_path)
        finally:
            temporary.cleanup()

    def test_rejects_wrong_revision(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["b1"]["environment"]["git_revision"] = "0" * 40
        with self.assertRaisesRegex(ValidationError, "git_revision"):
            self._validate(documents)

    def test_rejects_binary_provenance_mutation(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["b2"]["binary_identity"]["binary_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValidationError, "binary_sha256"):
            self._validate(documents)

    def test_rejects_duplicate_report_path(self) -> None:
        temporary, paths, projection_path = self._fixture()
        try:
            with self.assertRaisesRegex(ValidationError, "four distinct report paths"):
                validate_paths(paths["a1"], paths["a1"], paths["b2"], paths["a2"], projection_path)
        finally:
            temporary.cleanup()

    def test_rejects_cross_leg_child_pid_reuse(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["b1"]["filesystem_evidence"][0]["samples"][0]["child_process_id"] = (
            documents["a1"]["filesystem_evidence"][0]["samples"][0]["child_process_id"]
        )
        with self.assertRaisesRegex(ValidationError, "process IDs"):
            self._validate(documents)

    def test_rejects_operation_vector_raw_sample_mismatch(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["b2"]["results"][0]["operation_metrics"]["allocation"][
            "allocated_bytes"
        ]["values"][0] += 1
        with self.assertRaisesRegex(ValidationError, "raw alignment"):
            self._validate(documents)

    def test_rejects_exact_control_vector_mutation(self) -> None:
        documents = copy.deepcopy(self.documents)
        sample = documents["a2"]["filesystem_evidence"][0]["samples"]
        vector = documents["a2"]["results"][0]["operation_metrics"]["allocation"][
            "allocation_calls"
        ]["values"]
        for item in sample:
            item["allocation_metrics"]["allocation_calls"] += 1
        for index in range(len(vector)):
            vector[index] += 1
        with self.assertRaisesRegex(ValidationError, "exact Change 0401 value"):
            self._validate(documents, projection=None)

    def test_rejects_same_implementation_nonclaimable_difference(self) -> None:
        documents = copy.deepcopy(self.documents)
        sample = documents["b2"]["filesystem_evidence"][0]["samples"]
        vector = documents["b2"]["results"][0]["operation_metrics"]["allocation"][
            "live_bytes_after"
        ]["values"]
        for item in sample:
            item["allocation_metrics"]["live_bytes_after"] += 1
        for index in range(len(vector)):
            vector[index] += 1
        with self.assertRaisesRegex(ValidationError, "candidate live_bytes_after vector equality"):
            self._validate(documents, projection=None)

    def test_rejects_selected_cell_oracle_mutation(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["a2"]["filesystem_evidence"][0]["samples"][7][
            "xlsx_selected_cell"
        ]["lexical_value"] = "1028013"
        with self.assertRaisesRegex(ValidationError, "xlsx_selected_cell"):
            self._validate(documents)

    def test_rejects_operation_sample_index_alignment_mutation(self) -> None:
        documents = copy.deepcopy(self.documents)
        indices = documents["b1"]["results"][0]["operation_metrics"]["sample_indices"]
        indices[0], indices[1] = indices[1], indices[0]
        with self.assertRaisesRegex(ValidationError, "sample_indices"):
            self._validate(documents, projection=None)

    def test_rejects_elapsed_raw_sample_alignment_mutation(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["a1"]["results"][0]["elapsed_ns"]["samples"][0] += 1
        with self.assertRaisesRegex(ValidationError, "elapsed/raw sample alignment"):
            self._validate(documents, projection=None)

    def test_rejects_projection_report_digest_mutation(self) -> None:
        projection = copy.deepcopy(self.base_projection)
        projection["per_leg"]["A1"]["raw_report_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValidationError, "raw_report_sha256"):
            self._validate(projection=projection)

    def test_rejects_projection_delta_mutation(self) -> None:
        projection = copy.deepcopy(self.base_projection)
        projection["matched_deltas"]["A1_to_B1"]["metrics"]["allocation_calls"]["delta"] += 1
        with self.assertRaisesRegex(ValidationError, "allocation_calls"):
            self._validate(projection=projection)

    def test_rejects_claimability_scope_mutation(self) -> None:
        projection = copy.deepcopy(self.base_projection)
        projection["claimability"]["allocated_bytes"]["scope"] = "all workloads"
        with self.assertRaisesRegex(ValidationError, "allocated_bytes"):
            self._validate(projection=projection)

    def test_rejects_nonclaimable_claim_mutation(self) -> None:
        projection = copy.deepcopy(self.base_projection)
        projection["claimability"]["peak_live_bytes_after"]["claimable"] = True
        with self.assertRaisesRegex(ValidationError, "peak_live_bytes_after"):
            self._validate(projection=projection)

    def test_rejects_allocator_latency_statistic_in_projection(self) -> None:
        projection = copy.deepcopy(self.base_projection)
        projection["claimability"]["allocator_elapsed_ns"]["mean"] = 1.0
        with self.assertRaisesRegex(ValidationError, "allocator_elapsed_ns"):
            self._validate(projection=projection)

    def test_allocator_elapsed_summary_is_not_claimed(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["a1"]["results"][0]["elapsed_ns"]["mean"] = -12345.0
        temporary, paths, projection_path = self._fixture(documents, projection=None)
        try:
            result = validate_paths(paths["a1"], paths["b1"], paths["b2"], paths["a2"])
            self.assertEqual(result["reports"], 4)
        finally:
            temporary.cleanup()


if __name__ == "__main__":
    unittest.main()
