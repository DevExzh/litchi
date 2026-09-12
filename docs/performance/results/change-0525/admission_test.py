#!/usr/bin/env python3
"""Small in-memory negative vectors for the 0525 admission evaluators."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import unittest


HERE = Path(__file__).resolve().parent


def load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


NUMERIC = load(HERE / "analyze.py", "xlsx_0525_admission_numeric")
PROFILES = load(HERE / "analyze_profiles.py", "xlsx_0525_admission_profiles")


def plan():
    return {
        "primary": {"repeats": 1, "shapes": ["medium"]},
        "admission": (
            "Require at least 5% p50 improvement in primary total and commit "
            "phase in both paired repeats for both medium and dense-sparse. "
            "Require at least 15% commit Ir reduction in both shapes/repeats."
        ),
    }


def row(*, elapsed=100, commit=100, repeat=1, shape="medium"):
    return {
        "kind": "primary",
        "guard": None,
        "repeat": repeat,
        "shape": shape,
        "timing": {
            "elapsed_ns": {"p50": elapsed},
            "commit_ns": {"p50": commit},
        },
    }


def evidence(rows):
    return {"native": {"rows": rows}}


class AdmissionNegativeVectors(unittest.TestCase):
    def test_missing_primary_row_is_rejected(self):
        with self.assertRaises(NUMERIC.BASE.EvidenceError):
            NUMERIC._primary_rows(evidence([]), plan(), "baseline")

    def test_duplicate_primary_row_is_rejected(self):
        duplicate = row()
        with self.assertRaises(NUMERIC.BASE.EvidenceError):
            NUMERIC._primary_rows(evidence([duplicate, dict(duplicate)]), plan(), "baseline")

    def test_below_threshold_total_fails_only_total_gate(self):
        result = {
            "baseline": evidence([row(elapsed=100, commit=100)]),
            "candidate": evidence([row(elapsed=96, commit=90)]),
        }
        admission = NUMERIC._native_admission(result, plan())
        measured = admission["rows"][0]
        self.assertFalse(measured["native_primary_total_p50"]["passed"])
        self.assertTrue(measured["native_primary_commit_p50"]["passed"])
        self.assertFalse(admission["passed"])

    def test_below_threshold_commit_fails_only_commit_gate(self):
        result = {
            "baseline": evidence([row(elapsed=100, commit=100)]),
            "candidate": evidence([row(elapsed=90, commit=96)]),
        }
        admission = NUMERIC._native_admission(result, plan())
        measured = admission["rows"][0]
        self.assertTrue(measured["native_primary_total_p50"]["passed"])
        self.assertFalse(measured["native_primary_commit_p50"]["passed"])
        self.assertFalse(admission["passed"])

    def test_invalid_numeric_value_is_rejected(self):
        invalid_baseline = {
            "baseline": evidence([row(elapsed=0)]),
            "candidate": evidence([row(elapsed=90)]),
        }
        invalid_candidate = {
            "baseline": evidence([row()]),
            "candidate": evidence([row(elapsed="90")]),
        }
        with self.assertRaises(NUMERIC.BASE.EvidenceError):
            NUMERIC._native_admission(invalid_baseline, plan())
        with self.assertRaises(NUMERIC.BASE.EvidenceError):
            NUMERIC._native_admission(invalid_candidate, plan())

    def test_profile_commit_ir_below_threshold_is_rejected(self):
        admission = PROFILES.profile_commit_ir_admission(100, 86, 15.0)
        self.assertAlmostEqual(admission["reduction_percent"], 14.0)
        self.assertFalse(admission["passed"])


if __name__ == "__main__":
    unittest.main()
