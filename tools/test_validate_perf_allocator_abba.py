from __future__ import annotations

import copy
import hashlib
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from tools.validate_perf_allocator_abba import ValidationError, validate_paths


ROOT = Path(__file__).resolve().parents[1]
BUNDLE = ROOT / "docs" / "performance" / "results" / "change-0400"
FRAME_NAMES = {
    "a1": "a1-allocator.json.zst",
    "b1": "b1-allocator.json.zst",
    "b2": "b2-allocator.json.zst",
    "a2": "a2-allocator.json.zst",
}


@unittest.skipUnless(shutil.which("zstd"), "zstd is required for evidence-frame tests")
class AllocatorAbbaValidatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.documents = {}
        cls.raw_documents = {}
        for leg, name in FRAME_NAMES.items():
            raw = subprocess.run(
                ["zstd", "-q", "-d", "-c", str(BUNDLE / name)],
                check=True,
                stdout=subprocess.PIPE,
            ).stdout
            cls.raw_documents[leg] = raw
            cls.documents[leg] = json.loads(raw)
        cls.projection = json.loads((BUNDLE / "allocation-metrics.json").read_text())

    def _write_fixture(
        self,
        documents: dict[str, object] | None = None,
        projection: object | None = None,
    ) -> tuple[tempfile.TemporaryDirectory[str], dict[str, Path], Path]:
        temporary = tempfile.TemporaryDirectory(prefix="litchi-0400-allocator-validator-")
        root = Path(temporary.name)
        docs = copy.deepcopy(documents if documents is not None else self.documents)
        projection_document = copy.deepcopy(
            projection if projection is not None else self.projection
        )
        paths = {}
        for leg in ("a1", "b1", "b2", "a2"):
            path = root / f"{leg}.json"
            if documents is None:
                path.write_bytes(self.raw_documents[leg])
            else:
                path.write_text(json.dumps(docs[leg], indent=2) + "\n")
            paths[leg] = path
            if documents is not None:
                projection_document["per_leg"][leg.upper()]["raw_report_sha256"] = (
                    hashlib.sha256(path.read_bytes()).hexdigest()
                )
        projection_path = root / "allocation-metrics.json"
        projection_path.write_text(
            json.dumps(projection_document, indent=2) + "\n"
        )
        return temporary, paths, projection_path

    def _validate(self, documents=None, projection=None) -> None:
        temporary, paths, projection_path = self._write_fixture(documents, projection)
        try:
            validate_paths(
                paths["a1"], paths["b1"], paths["b2"], paths["a2"], projection_path
            )
        finally:
            temporary.cleanup()

    def test_retained_evidence_and_projection_validate(self) -> None:
        self._validate()

    def test_rejects_duplicate_json_keys(self) -> None:
        temporary, paths, projection_path = self._write_fixture()
        try:
            raw = paths["a1"].read_text()
            paths["a1"].write_text(raw.replace("{", '{"schema_version": 1,', 1))
            with self.assertRaisesRegex(ValidationError, "duplicate JSON object key"):
                validate_paths(
                    paths["a1"], paths["b1"], paths["b2"], paths["a2"], projection_path
                )
        finally:
            temporary.cleanup()

    def test_rejects_non_finite_json_number(self) -> None:
        temporary, paths, projection_path = self._write_fixture()
        try:
            raw = paths["a1"].read_text()
            paths["a1"].write_text(raw.replace('"mean": ', '"mean": NaN, "discarded_mean": ', 1))
            with self.assertRaisesRegex(ValidationError, "non-finite JSON number"):
                validate_paths(
                    paths["a1"], paths["b1"], paths["b2"], paths["a2"], projection_path
                )
        finally:
            temporary.cleanup()

    def test_rejects_overflowed_json_float(self) -> None:
        temporary, paths, projection_path = self._write_fixture()
        try:
            raw = paths["a1"].read_text()
            paths["a1"].write_text(
                raw.replace('"mean": ', '"mean": 1e9999, "discarded_mean": ', 1)
            )
            with self.assertRaisesRegex(ValidationError, "non-finite JSON number"):
                validate_paths(
                    paths["a1"],
                    paths["b1"],
                    paths["b2"],
                    paths["a2"],
                    projection_path,
                )
        finally:
            temporary.cleanup()

    def test_rejects_wrong_revision(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["b1"]["environment"]["git_revision"] = "0" * 40
        with self.assertRaisesRegex(ValidationError, "git_revision"):
            self._validate(documents)

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

    def test_rejects_bool_allocator_value(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["b2"]["filesystem_evidence"][0]["samples"][0][
            "allocation_metrics"
        ]["allocated_bytes"] = True
        with self.assertRaisesRegex(ValidationError, "must be an integer"):
            self._validate(documents)

    def test_rejects_same_implementation_vector_difference(self) -> None:
        documents = copy.deepcopy(self.documents)
        sample = documents["a2"]["filesystem_evidence"][0]["samples"][0]
        sample["allocation_metrics"]["allocation_calls"] += 1
        order = documents["a2"]["results"][0]["elapsed_ns"]["sample_order"]
        rank = order.index(0)
        documents["a2"]["results"][0]["operation_metrics"]["allocation"][
            "allocation_calls"
        ]["values"][rank] += 1
        with self.assertRaisesRegex(ValidationError, "not constant"):
            self._validate(documents)

    def test_rejects_selected_cell_oracle_mutation(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["a2"]["filesystem_evidence"][0]["samples"][7][
            "xlsx_selected_cell"
        ]["lexical_value"] = "1028013"
        with self.assertRaisesRegex(ValidationError, "xlsx_selected_cell"):
            self._validate(documents)

    def test_rejects_projection_report_digest_mutation(self) -> None:
        projection = copy.deepcopy(self.projection)
        projection["per_leg"]["A1"]["raw_report_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValidationError, "raw_report_sha256"):
            self._validate(projection=projection)

    def test_rejects_projection_delta_mutation(self) -> None:
        projection = copy.deepcopy(self.projection)
        projection["matched_deltas"]["A1_to_B1"]["metrics"]["allocation_calls"][
            "delta"
        ] += 1
        with self.assertRaisesRegex(ValidationError, "allocation_calls"):
            self._validate(projection=projection)

    def test_rejects_allocator_latency_statistic_in_projection(self) -> None:
        projection = copy.deepcopy(self.projection)
        projection["claimability"]["allocator_elapsed_ns"]["mean"] = 1.0
        with self.assertRaisesRegex(ValidationError, "allocator_elapsed_ns"):
            self._validate(projection=projection)

    def test_rejects_claimable_projection_scope_mutation(self) -> None:
        projection = copy.deepcopy(self.projection)
        projection["claimability"]["allocated_bytes"]["scope"] = "all workloads"
        with self.assertRaisesRegex(ValidationError, "allocated_bytes.scope"):
            self._validate(projection=projection)

    def test_rejects_nonclaimable_projection_reason_mutation(self) -> None:
        projection = copy.deepcopy(self.projection)
        projection["claimability"]["peak_live_bytes_after"]["reason"] = (
            "operation peak memory reduction"
        )
        with self.assertRaisesRegex(ValidationError, "peak_live_bytes_after.reason"):
            self._validate(projection=projection)

    def test_allocator_elapsed_summary_is_not_evaluated(self) -> None:
        documents = copy.deepcopy(self.documents)
        documents["a1"]["results"][0]["elapsed_ns"]["mean"] = -12345.0
        self._validate(documents)


if __name__ == "__main__":
    unittest.main()
