"""Portable evidence checks; no Cargo, capture, or external tools required."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location("analyzer", Path(__file__).with_name("analyze_profile.py"))
a = importlib.util.module_from_spec(spec)
spec.loader.exec_module(a)


class EvidenceTests(unittest.TestCase):
    def test_typed_outcomes_and_dimensions_are_not_interchangeable(self):
        self.assertFalse(a.same_expected_outcome({"number": 1}, {"number": True}))
        self.assertFalse(a.same_expected_outcome({"numbers": [[1, 2]]}, {"numbers": [[1], [2]]}))
        self.assertEqual(a.shape_geometry("2x1", "test"), {"kind": "array", "rows": 2, "columns": 1, "elements": 2})
        for bad in ("0x1", "1x0", "1x2x3", None):
            with self.subTest(bad=bad), self.assertRaises(a.CaptureError):
                a.shape_geometry(bad, "test")

    def test_bootstrap_is_reproducible_and_preserves_known_shift(self):
        left = a.bootstrap_delta_interval([10.0] * 15, [12.0] * 15, 0.5, 123)
        right = a.bootstrap_delta_interval([10.0] * 15, [12.0] * 15, 0.5, 123)
        self.assertEqual(left, right)
        self.assertEqual(left["difference_ci95"], {"low": 2.0, "high": 2.0})
        self.assertEqual(left["relative_ci95"], {"low": 0.2, "high": 0.2})
        self.assertEqual(left["replicates"], 10000)

    def test_freeze_requires_available_matching_bytes_and_manifest(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            path = root / "freeze.json"
            manifest = {"selected_files": {"source.rs": "a" * 64}, "base_commit": "base"}
            path.write_text(json.dumps(manifest))
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            summary = {"candidate_freeze": {"path": str(path), "sha256": digest, "manifest": manifest}, "captures": [{"label": "candidate-final", "freeze_sha256": digest, "freeze_path": str(path), "freeze_base_commit": "base"}]}
            a.validate_freeze_receipt(summary)
            changed = copy.deepcopy(summary)
            changed["candidate_freeze"]["manifest"]["base_commit"] = "different"
            with self.assertRaises(a.CaptureError):
                a.validate_freeze_receipt(changed)
            path.write_text("{}")
            with self.assertRaises(a.CaptureError):
                a.validate_freeze_receipt(summary)
            path.unlink()
            with patch.object(a, "HERE", root / "performance"), self.assertRaises(a.CaptureError):
                a.validate_freeze_receipt(summary)

    def test_timed_geometry_and_sticky_reads_are_bound_to_contract(self):
        sample = {key: 0 for key in a.SAMPLE_COUNTERS}
        sample["reference_reads"] = 1
        row = {key: 0 for key in a.ROW_MEDIANS}
        row.update(case="case", phase="evaluate", sample_index=1, supported=True,
                   warmups=1, iterations=1, repeat=4, input_bytes=0,
                   elapsed_ns_per_repeat=0, work_per_repeat=0,
                   reference_reads_p50=1, reference_reads_per_repeat=0.25,
                   samples=[sample], rss_kib=1, binary_sha256="binary",
                   source_git_head="source", raw_stdout="raw", raw_time="time",
                   shape="array", rows=2, columns=1, elements=2)
        contract = {"geometry": a.shape_geometry("2x1", "test"),
                    "reference_reads": 1, "row": {"failure": "Cancelled"}}
        with patch.object(a, "raw_child", return_value=(row, 1)):
            a.validate_row(row, Path("."), "test", 1, contract)
            normal = copy.deepcopy(contract)
            normal["row"] = {}
            with self.assertRaisesRegex(a.CaptureError, "timed reference reads"):
                a.validate_row(row, Path("."), "test", 1, normal)
            row["rows"], row["columns"] = 1, 2
            with self.assertRaisesRegex(a.CaptureError, "execution rows"):
                a.validate_row(row, Path("."), "test", 1, contract)

    def test_missing_preflight_is_refused(self):
        with TemporaryDirectory() as temporary, self.assertRaises(a.CaptureError):
            a.validate_preflight(Path(temporary), ["case"], {})


if __name__ == "__main__":
    unittest.main()
