#!/usr/bin/env python3
"""Static source-pair and preparation-receipt tests for the XLSX raw lane."""

from __future__ import annotations

import subprocess
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from native_qualification import prepare_xlsx_raw_inputs as preparation  # noqa: E402


def _git(root: Path, *arguments: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.stdout.strip()


def _repo(root: Path, marker: str) -> str:
    root.mkdir()
    _git(root, "init", "--quiet")
    _git(root, "config", "user.email", "xlsx-raw-protocol@test.invalid")
    _git(root, "config", "user.name", "xlsx raw protocol test")
    for relative in preparation.ADAPTER_PATHS:
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(f"{marker}:{relative}\n", encoding="utf-8")
    _git(root, "add", ".")
    _git(root, "commit", "--quiet", "-m", "source")
    return _git(root, "rev-parse", "HEAD")


class PrepareXlsxRawInputsTests(unittest.TestCase):
    def test_prepares_real_source_pair_and_binds_receipt(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xlsx-raw-preparation-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline_revision = _repo(baseline, "baseline")
            candidate_revision = _repo(candidate, "candidate")
            receipt = root / "preparation.json"
            with mock.patch.object(preparation, "BASELINE_REVISION", baseline_revision):
                preparation.prepare(baseline, candidate, candidate_revision, receipt)
                preparation.validate_preparation_receipt(
                    receipt,
                    baseline_root=baseline,
                    candidate_root=candidate,
                    candidate_revision=candidate_revision,
                )

    def test_rejects_unapproved_baseline_change_after_receipt(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xlsx-raw-preparation-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline_revision = _repo(baseline, "baseline")
            candidate_revision = _repo(candidate, "candidate")
            receipt = root / "preparation.json"
            with mock.patch.object(preparation, "BASELINE_REVISION", baseline_revision):
                preparation.prepare(baseline, candidate, candidate_revision, receipt)
                (baseline / "unapproved.txt").write_text("changed\n", encoding="utf-8")
                with self.assertRaisesRegex(preparation.PreparationError, "baseline preparation identity changed"):
                    preparation.validate_preparation_receipt(
                        receipt,
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_revision=candidate_revision,
                    )

    def test_rejects_candidate_mutation_and_revision_echo_mismatch(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xlsx-raw-preparation-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline_revision = _repo(baseline, "baseline")
            candidate_revision = _repo(candidate, "candidate")
            receipt = root / "preparation.json"
            with mock.patch.object(preparation, "BASELINE_REVISION", baseline_revision):
                preparation.prepare(baseline, candidate, candidate_revision, receipt)
                path = candidate / preparation.PROBE_PATHS[0]
                path.write_text("candidate mutation\n", encoding="utf-8")
                with self.assertRaisesRegex(preparation.PreparationError, "candidate preparation identity changed"):
                    preparation.validate_preparation_receipt(
                        receipt,
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_revision=candidate_revision,
                    )


if __name__ == "__main__":
    unittest.main()
