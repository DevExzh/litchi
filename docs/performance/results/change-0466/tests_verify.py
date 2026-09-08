#!/usr/bin/env python3
"""Focused fail-closed tests for the portable 0466 bundle verifier."""

from __future__ import annotations

import gzip
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("litchi_0466_verify", ROOT / "verify.py")
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("cannot load verify.py")
VERIFY = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(VERIFY)


def file_sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def reseal_sums(bundle: Path) -> None:
    sums = bundle / "SHA256SUMS"
    if not sums.exists():
        return
    names = []
    for line in sums.read_text(encoding="utf-8").splitlines():
        fields = line.split("  ", 1)
        if len(fields) == 2 and fields[1] != "SHA256SUMS":
            names.append(fields[1])
    actual = {
        path.relative_to(bundle).as_posix()
        for path in bundle.rglob("*")
        if path.is_file() and path != sums
    }
    names = sorted(actual | set(names))
    sums.unlink()
    sums.write_text("".join(f"{file_sha(bundle / name)}  {name}\n" for name in names), encoding="utf-8")


class VerifyBundleTests(unittest.TestCase):
    def run_verifier(self, bundle: Path) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, "-B", str(bundle / "verify.py")],
            cwd=bundle,
            text=True,
            capture_output=True,
            check=False,
        )

    def copied_bundle(self) -> tuple[tempfile.TemporaryDirectory[str], Path]:
        # Keep the temporary tree on the repository filesystem so hard links
        # work even when the system /tmp is a different mount.
        temporary = tempfile.TemporaryDirectory(prefix=".litchi-0466-verify-", dir=ROOT.parents[3])
        parent = Path(temporary.name)
        bundle = parent / "change-0466"
        # Hard links keep the 254 MiB perf.data fixture cheap to copy.  Every
        # test mutates only after unlinking its target, so the source bundle is
        # never changed.
        shutil.copytree(ROOT, bundle, copy_function=os.link)
        shutil.copytree(ROOT.parent / "change-0465", parent / "change-0465", copy_function=os.link)
        return temporary, bundle

    @staticmethod
    def write_json(path: Path, value: object) -> None:
        path.unlink()
        path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    def test_current_bundle_passes(self) -> None:
        result = VERIFY.verify()
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["lanes"], list(VERIFY.LANES))
        self.assertEqual(result["case"], VERIFY.CASE)
        self.assertEqual(result["corpus_archive_sha256"], VERIFY.CORPUS_SHA256)

    def test_receipt_update_cannot_hide_changed_latency_sample(self) -> None:
        temporary, bundle = self.copied_bundle()
        try:
            report_path = bundle / "normal-r1" / "report.json"
            report = json.loads(report_path.read_text(encoding="utf-8"))
            samples = report["results"][0]["elapsed_ns"]["samples"]
            samples[-1] += 12345
            self.write_json(report_path, report)
            receipt_path = bundle / "normal-r1" / "receipt.json"
            receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
            row = receipt["artifacts"]["report.json"]
            row["bytes"] = report_path.stat().st_size
            row["sha256"] = file_sha(report_path)
            self.write_json(receipt_path, receipt)
            reseal_sums(bundle)
            result = self.run_verifier(bundle)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("statistics", result.stdout)
        finally:
            temporary.cleanup()

    def test_receipt_update_cannot_hide_changed_corpus_identity(self) -> None:
        temporary, bundle = self.copied_bundle()
        try:
            report_path = bundle / "normal-r2" / "report.json"
            report = json.loads(report_path.read_text(encoding="utf-8"))
            report["results"][0]["corpus"]["archive_sha256"] = "0" * 64
            self.write_json(report_path, report)
            receipt_path = bundle / "normal-r2" / "receipt.json"
            receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
            row = receipt["artifacts"]["report.json"]
            row["bytes"] = report_path.stat().st_size
            row["sha256"] = file_sha(report_path)
            self.write_json(receipt_path, receipt)
            reseal_sums(bundle)
            result = self.run_verifier(bundle)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("corpus archive identity", result.stdout)
        finally:
            temporary.cleanup()

    def test_gzipped_raw_log_preserves_receipt_identity(self) -> None:
        temporary, bundle = self.copied_bundle()
        try:
            raw = bundle / "normal-r1" / "resource.log"
            payload = raw.read_bytes()
            raw.unlink()
            with gzip.open(str(raw) + ".gz", "wb") as stream:
                stream.write(payload)
            reseal_sums(bundle)
            result = self.run_verifier(bundle)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        finally:
            temporary.cleanup()

    def test_fp_binding_cannot_be_repointed_to_normal_binary(self) -> None:
        temporary, bundle = self.copied_bundle()
        try:
            path = bundle / "profile-binding.json"
            binding = json.loads(path.read_text(encoding="utf-8"))
            binding["binary_sha256"] = "0" * 64
            # The altered value is intentionally a valid SHA-shaped value;
            # resealing must not make a false profile identity acceptable.
            self.write_json(path, binding)
            reseal_sums(bundle)
            result = self.run_verifier(bundle)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("binary identity differs", result.stdout)
        finally:
            temporary.cleanup()

    def test_fp_report_statistics_are_recomputed(self) -> None:
        temporary, bundle = self.copied_bundle()
        try:
            path = bundle / "samples-fp" / "report.json"
            report = json.loads(path.read_text(encoding="utf-8"))
            report["results"][0]["elapsed_ns"]["samples"][-1] += 12345
            self.write_json(path, report)
            reseal_sums(bundle)
            result = self.run_verifier(bundle)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("statistics", result.stdout)
        finally:
            temporary.cleanup()

    def test_derived_fp_summary_cannot_be_changed_after_resealing(self) -> None:
        temporary, bundle = self.copied_bundle()
        try:
            path = bundle / "summary-fp.json"
            summary = json.loads(path.read_text(encoding="utf-8"))
            summary["exact_contexts"]["rows"]["commit"]["weighted_event_period"] += 1
            self.write_json(path, summary)
            reseal_sums(bundle)
            result = self.run_verifier(bundle)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("derived summary", result.stdout)
        finally:
            temporary.cleanup()


if __name__ == "__main__":
    unittest.main()
