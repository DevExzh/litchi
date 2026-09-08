#!/usr/bin/env python3
"""Semantic and custody tests for the portable 0467 bundle verifier."""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import tempfile
import unittest


HERE = Path(__file__).resolve().parent


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _load(path: Path) -> object:
    return json.loads(path.read_text(encoding="utf-8"))


def _refresh_artifact(root: Path, lane: str, name: str) -> None:
    receipt_path = root / lane / "receipt.json"
    receipt = _load(receipt_path)
    row = receipt["artifacts"][name]
    path = root / lane / row.get("path", name)
    row["bytes"] = path.stat().st_size
    row["sha256"] = _sha(path)
    _write_json(receipt_path, receipt)


def _seal(root: Path) -> None:
    names = sorted(
        path.relative_to(root).as_posix()
        for path in root.rglob("*")
        if path.is_file() and path.name != "SHA256SUMS"
    )
    lines = [f"{_sha(root / name)}  {name}" for name in names]
    (root / "SHA256SUMS").write_text("\n".join(lines) + "\n", encoding="utf-8")


def _load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class BundleFixture:
    def __init__(self):
        self.temp = tempfile.TemporaryDirectory(prefix="litchi-0467-verify-")
        self.root = Path(self.temp.name)
        self._make()

    def close(self):
        self.temp.cleanup()

    def _make(self):
        for name in (
            "protocol-r1.json",
            "capture.py",
            "analyze.py",
            "source-bindings.json",
            "review.json",
        ):
            shutil.copy2(HERE / name, self.root / name)
        shutil.copytree(HERE / "sources", self.root / "sources")
        policy = HERE.parents[3] / "docs/performance/perf-regression-policy-v1.json"
        (self.root / "docs/performance").mkdir(parents=True)
        shutil.copy2(policy, self.root / "docs/performance/perf-regression-policy-v1.json")
        shutil.copy2(HERE / "control-binding.json", self.root / "control-binding.json")
        shutil.copy2(HERE / "control-build.json", self.root / "control-build.json")

        control = _load(self.root / "control-binding.json")
        candidate_revision = "b" * 40
        candidate_sha = "b" * 64
        candidate_build = _load(self.root / "control-build.json")
        candidate_build["role"] = "candidate"
        candidate_build["revision"] = candidate_revision
        _write_json(self.root / "candidate-build.json", candidate_build)
        candidate = {
            "revision": candidate_revision,
            "binary_sha256": candidate_sha,
            "bytes": control["bytes"],
            "build_receipt": "candidate-build.json",
            "build_receipt_sha256": _sha(self.root / "candidate-build.json"),
            "clean_build": True,
        }
        _write_json(self.root / "candidate-binding.json", candidate)
        source_bindings = _load(self.root / "source-bindings.json")
        source_bindings["roles"]["candidate"]["revision"] = candidate_revision
        source_bindings["roles"]["candidate"]["git_tree"] = "c" * 40
        _write_json(self.root / "source-bindings.json", source_bindings)

        self._copy_lane("A1-clean", "control", "A1-clean")
        self._copy_lane("A1-clean", "control", "A2-clean")
        self._copy_lane("A1-clean", "candidate", "B1-clean")
        self._copy_lane("A1-clean", "candidate", "B2-clean")
        self._copy_lane("A-full-clean", "control", "A-full-clean")
        self._copy_lane("A-full-clean", "candidate", "B-full-clean")
        self._copy_lane("A-heap-clean", "control", "A-heap-clean")
        self._copy_lane("A-heap-clean", "candidate", "B-heap-clean")
        self._write_analysis()

    def _copy_lane(self, source_name: str, role: str, lane: str) -> None:
        source = HERE / source_name
        target = self.root / lane
        shutil.copytree(source, target)
        for export_name in ("print.txt", "print.stderr"):
            export_path = target / export_name
            if export_path.exists():
                export_path.unlink()
        binding = _load(self.root / f"{role}-binding.json")
        receipt = _load(target / "receipt.json")
        started = _load(target / "started.json")
        for record in (receipt, started):
            record["lane"] = lane
            record["role"] = role
            record["revision"] = binding["revision"]
            record["binary_sha256"] = binding["binary_sha256"]
            record["binding_sha256"] = _sha(self.root / f"{role}-binding.json")
            record["driver_sha256"] = _sha(self.root / "capture.py")
            record["argv"] = [
                value.replace(source_name, lane) if isinstance(value, str) else value
                for value in record["argv"]
            ]
        _write_json(target / "started.json", started)

        report = _load(target / "report.json")
        catalog = _load(target / "corpus-catalog.json")
        if role == "candidate":
            report["environment"]["git_revision"] = binding["revision"]
            report["binary_identity"]["binary_sha256"] = binding["binary_sha256"]
            report["binary_identity"]["path"] = "/tmp/formal-candidate"
            catalog["build"]["git_revision"] = binding["revision"]
            catalog_without_hash = dict(catalog)
            catalog_without_hash.pop("catalog_sha256", None)
            catalog["catalog_sha256"] = hashlib.sha256(
                json.dumps(catalog_without_hash, sort_keys=True, separators=(",", ":")).encode()
            ).hexdigest()
            report["corpus_catalog"]["catalog_sha256"] = catalog["catalog_sha256"]
        _write_json(target / "corpus-catalog.json", catalog)
        _write_json(target / "report.json", report)

        receipt["artifacts"]["started.json"]["sha256"] = _sha(target / "started.json")
        receipt["artifacts"]["started.json"]["bytes"] = (target / "started.json").stat().st_size
        receipt["artifacts"]["corpus-catalog.json"]["sha256"] = _sha(target / "corpus-catalog.json")
        receipt["artifacts"]["corpus-catalog.json"]["bytes"] = (target / "corpus-catalog.json").stat().st_size
        receipt["artifacts"]["report.json"]["sha256"] = _sha(target / "report.json")
        receipt["artifacts"]["report.json"]["bytes"] = (target / "report.json").stat().st_size
        _write_json(target / "receipt.json", receipt)

    def _write_analysis(self):
        analyzer = _load_module(HERE / "analyze.py", "litchi_0467_test_analyze")
        result = analyzer.analyze(self.root)
        _write_json(self.root / "analysis.json", result)


class VerifierTests(unittest.TestCase):
    def setUp(self):
        self.bundle = BundleFixture()
        self.verify = _load_module(HERE / "verify.py", "litchi_0467_test_verify")
        self.verify.ROOT = self.bundle.root
        # The production --preseal path also checks /tmp/litchi-goal-0467
        # binaries.  Fixture tests are portable and deliberately have no live
        # binaries; the custody check itself is exercised by the coordinator's
        # preseal run on the real bundle.
        self.verify._preseal_live_bindings = lambda bindings: None

    def tearDown(self):
        self.bundle.close()

    def test_preseal_passes_without_live_binaries_or_checksum_file(self):
        result = self.verify.verify(preseal=True)
        self.assertEqual(result["status"], "pass")
        self.assertFalse(result["sealed"])
        self.assertEqual(result["normal"]["result_count"], 6)
        self.assertEqual(result["full_guard"]["result_count"], 201)
        self.assertEqual(result["heap_guard"]["instrumentation"], "external-heaptrack-whole-process")

    def test_analysis_is_recomputed_when_retained(self):
        result = self.verify.verify(preseal=True)
        self.assertEqual(result["status"], "pass")

    def test_semantic_report_tamper_is_rejected_after_receipt_hash_repair(self):
        path = self.bundle.root / "B1-clean" / "report.json"
        report = _load(path)
        report["results"][0]["corpus"]["archive_sha256"] = "c" * 64
        _write_json(path, report)
        _refresh_artifact(self.bundle.root, "B1-clean", "report.json")
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_clean_build_tamper_is_rejected_after_binding_hash_repair(self):
        path = self.bundle.root / "candidate-build.json"
        build = _load(path)
        build["clean_after"] = False
        _write_json(path, build)
        binding_path = self.bundle.root / "candidate-binding.json"
        binding = _load(binding_path)
        binding["build_receipt_sha256"] = _sha(path)
        _write_json(binding_path, binding)
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_heaptrack_wrapper_tamper_is_rejected(self):
        path = self.bundle.root / "B-heap-clean" / "receipt.json"
        receipt = _load(path)
        receipt["argv"] = [value for value in receipt["argv"] if value != "heaptrack"]
        _write_json(path, receipt)
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_postprocess_exports_are_optional_but_authenticated(self):
        commands = []
        for lane in ("A-heap-clean", "B-heap-clean"):
            lane_dir = self.bundle.root / lane
            (lane_dir / "print.txt").write_text("heaptrack report\n", encoding="utf-8")
            (lane_dir / "print.stderr").write_text("\n", encoding="utf-8")
            commands.append(
                {
                    "lane": lane,
                    "argv": [
                        "heaptrack_print",
                        "-f",
                        str(lane_dir / "heaptrack.zst"),
                        "-n",
                        "30",
                    ],
                    "started_utc": "2026-09-08T00:00:00+00:00",
                    "exit_code": 0,
                    "finished_utc": "2026-09-08T00:00:01+00:00",
                    "artifacts": {
                        "print.txt": _sha(lane_dir / "print.txt"),
                        "print.stderr": _sha(lane_dir / "print.stderr"),
                    },
                }
            )
        _write_json(
            self.bundle.root / "postprocess.json",
            {
                "environment": {"DEBUGINFOD_URLS": "", "LC_ALL": "C"},
                "commands": commands,
            },
        )
        result = self.verify.verify(preseal=True)
        self.assertEqual(result["postprocess"]["lanes"], ["A-heap-clean", "B-heap-clean"])
        (self.bundle.root / "B-heap-clean" / "print.txt").write_text(
            "tampered\n", encoding="utf-8"
        )
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_heap_summary_recomputes_totals_after_repaired_export_hash(self):
        commands = []
        for lane in ("A-heap-clean", "B-heap-clean"):
            lane_dir = self.bundle.root / lane
            shutil.copy2(HERE / lane / "print.txt", lane_dir / "print.txt")
            shutil.copy2(HERE / lane / "print.stderr", lane_dir / "print.stderr")
            commands.append(
                {
                    "lane": lane,
                    "argv": [
                        "heaptrack_print",
                        "-f",
                        str(lane_dir / "heaptrack.zst"),
                        "-n",
                        "30",
                    ],
                    "started_utc": "2026-09-08T00:00:00+00:00",
                    "exit_code": 0,
                    "finished_utc": "2026-09-08T00:00:01+00:00",
                    "artifacts": {
                        "print.txt": _sha(lane_dir / "print.txt"),
                        "print.stderr": _sha(lane_dir / "print.stderr"),
                    },
                }
            )
        _write_json(
            self.bundle.root / "postprocess.json",
            {"environment": {"DEBUGINFOD_URLS": "", "LC_ALL": "C"}, "commands": commands},
        )
        shutil.copy2(HERE / "heap_summary.py", self.bundle.root / "heap_summary.py")
        shutil.copy2(HERE / "heap-summary.json", self.bundle.root / "heap-summary.json")
        self.assertEqual(self.verify.verify(preseal=True)["status"], "pass")

        candidate_print = self.bundle.root / "B-heap-clean" / "print.txt"
        candidate_print.write_text(
            candidate_print.read_text(encoding="utf-8").replace(
                "calls to allocation functions: 44815468 (",
                "calls to allocation functions: 44815469 (",
                1,
            ),
            encoding="utf-8",
        )
        postprocess = _load(self.bundle.root / "postprocess.json")
        postprocess["commands"][1]["artifacts"]["print.txt"] = _sha(candidate_print)
        _write_json(self.bundle.root / "postprocess.json", postprocess)
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_guard_summary_is_recomputed_when_retained(self):
        # Keep this fixture-level test independent of the expensive four-lane
        # capture.  It proves the verifier treats the retained summary as
        # derived semantic evidence, rather than trusting its JSON or only its
        # eventual checksum entry.
        (self.bundle.root / "guard_summary.py").write_text(
            "def summarize(root):\n    return {'schema': 'fixture-guard', 'value': 1}\n",
            encoding="utf-8",
        )
        _write_json(
            self.bundle.root / "guard-summary.json",
            {"schema": "fixture-guard", "value": 1},
        )
        self.assertEqual(self.verify.verify(preseal=True)["status"], "pass")
        _write_json(
            self.bundle.root / "guard-summary.json",
            {"schema": "fixture-guard", "value": 2},
        )
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_fixed_qualification_is_recomputed_when_retained(self):
        # Exercise the optional hook without requiring the four 500-sample
        # fixed-checkout lanes in this small portable fixture.
        (self.bundle.root / "fixed_qualify.py").write_text(
            "def qualify(root):\n    return {'schema': 'fixture-fixed', 'value': 1}\n",
            encoding="utf-8",
        )
        _write_json(
            self.bundle.root / "fixed-qualification.json",
            {"schema": "fixture-fixed", "value": 1},
        )
        self.assertEqual(self.verify.verify(preseal=True)["status"], "pass")
        _write_json(
            self.bundle.root / "fixed-qualification.json",
            {"schema": "fixture-fixed", "value": 2},
        )
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_full_guard_row_deletion_is_rejected(self):
        path = self.bundle.root / "B-full-clean" / "report.json"
        report = _load(path)
        report["results"].pop()
        _write_json(path, report)
        _refresh_artifact(self.bundle.root, "B-full-clean", "report.json")
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=True)

    def test_checksum_seal_detects_unlisted_or_modified_artifact(self):
        _seal(self.bundle.root)
        self.assertTrue(self.verify.verify(preseal=False)["sealed"])
        extra = self.bundle.root / "unlisted.txt"
        extra.write_text("tamper\n", encoding="utf-8")
        with self.assertRaises(self.verify.VerificationError):
            self.verify.verify(preseal=False)


if __name__ == "__main__":
    unittest.main()
