import hashlib
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from tools import perf_abba_package


ZSTD = shutil.which("zstd")


def canonical(value):
    return perf_abba_package.canonical_json(value, "fixture")


def make_fixture():
    reports = {}
    for role, revision in (
        ("a1", "control-revision"),
        ("b1", "candidate-revision"),
        ("b2", "candidate-revision"),
        ("a2", "control-revision"),
    ):
        reports[role] = {
            "schema_version": 1,
            "tool": {
                "name": "litchi-perf-baseline",
                "version": "0.1.0",
                "profile": "release",
                "target_os": "linux",
                "target_arch": "x86_64",
            },
            "environment": {
                "rustc_version": "rustc 1.95.0 (test)",
                "git_revision": revision,
                "git_worktree_dirty": False,
            },
            "configuration": {"samples_per_case": 15, "cases": ["fixture"]},
            "results": [{"case": "fixture", "elapsed_ns": {"samples": [1]}}],
        }
    report_identity = {
        role: {"canonical_sha256": hashlib.sha256(canonical(report)).hexdigest()}
        for role, report in reports.items()
    }
    configuration = reports["a1"]["configuration"]
    harness_tool = reports["a1"]["tool"]
    summary = {
        "schema_version": 1,
        "tool": {"name": "litchi-perf-abba-summary", "version": "0.1.0"},
        "protocol": {
            "order": ["a1_control", "b1_candidate", "b2_candidate", "a2_control"]
        },
        "harness_identity": {
            "schema_version": 1,
            "tool": harness_tool,
            "configuration": configuration,
        },
        "environment": {
            "stable": {"rustc_version": "rustc 1.95.0 (test)"},
            "legs": {
                role: report["environment"] for role, report in reports.items()
            },
        },
        "implementation_identity": {
            "control": {"git_revision": "control-revision", "legs": ["a1", "a2"]},
            "candidate": {
                "git_revision": "candidate-revision",
                "legs": ["b1", "b2"],
            },
            "distinct": True,
        },
        "report_identity": report_identity,
        "results": [{"case": "fixture"}],
        "verification": {
            "result_count": 1,
            "tool_identity_verified": True,
            "configuration_identity_verified": True,
            "environment_stable_identity_verified": True,
            "environment_legs_recorded": True,
            "case_corpus_identity_verified": True,
            "statistics_recomputed_from_samples": True,
        },
    }
    return reports, summary


@unittest.skipUnless(ZSTD, "zstd executable is required")
class PerfAbbaPackageTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        self.source = self.root / "source"
        self.source.mkdir()
        self.reports, self.summary = make_fixture()
        self.report_paths = {}
        for role, report in self.reports.items():
            path = self.source / f"report-{role}.json"
            # Deliberately use different whitespace in each source.  The
            # compressed artifact must preserve these bytes exactly while
            # canonical identity remains bound to the strict summary.
            path.write_bytes(
                (json.dumps(report, sort_keys=True, indent=2) + f"\n  \n").encode()
            )
            self.report_paths[role] = path
        self.summary_path = self.source / "summary.json"
        self.summary_path.write_text(
            json.dumps(self.summary, sort_keys=True, indent=2) + "\n", encoding="utf-8"
        )

    def tearDown(self):
        self.tempdir.cleanup()

    def specs(self):
        return {
            role: self.report_paths[role] for role in ("a1", "b1", "b2", "a2")
        }

    def test_packages_raw_json_and_records_all_digests(self):
        output = self.root / "package"
        manifest = perf_abba_package.package_artifacts(
            change_id="0238-perf-package",
            output_dir=output,
            summary=self.summary_path,
            artifacts=self.specs(),
        )

        self.assertEqual(manifest["change_id"], "0238-perf-package")
        self.assertEqual(manifest["change"], "0238-perf-package")
        self.assertEqual(manifest["self_excluded"], True)
        self.assertEqual(
            [item["role"] for item in manifest["artifacts"]], ["a1", "b1", "b2", "a2"]
        )
        self.assertEqual(manifest["summary_identity"]["path"], "summary.json")
        self.assertEqual(
            manifest["summary_identity"]["sha256"],
            hashlib.sha256(self.summary_path.read_bytes()).hexdigest(),
        )
        self.assertEqual(manifest["summary"], manifest["summary_identity"])

        for item in manifest["artifacts"]:
            compressed_path = output / item["path"]
            compressed = compressed_path.read_bytes()
            raw = self.report_paths[item["role"]].read_bytes()
            self.assertEqual(item["bytes"], len(compressed))
            self.assertEqual(item["sha256"], hashlib.sha256(compressed).hexdigest())
            self.assertEqual(item["uncompressed_bytes"], len(raw))
            self.assertEqual(item["uncompressed_sha256"], hashlib.sha256(raw).hexdigest())
            decompressed = subprocess.run(
                [ZSTD, "--quiet", "--decompress", "--stdout", str(compressed_path)],
                check=True,
                stdout=subprocess.PIPE,
            ).stdout
            self.assertEqual(decompressed, raw)

        manifest_path = output / "0238-perf-package-manifest.json"
        self.assertEqual(json.loads(manifest_path.read_text()), manifest)
        self.assertEqual((output / "summary.json").read_bytes(), self.summary_path.read_bytes())

    def test_compression_and_manifest_are_deterministic(self):
        first = perf_abba_package.package_artifacts(
            change_id="0238",
            output_dir=self.root / "first",
            summary=self.summary_path,
            artifacts=self.specs(),
        )
        second = perf_abba_package.package_artifacts(
            change_id="0238",
            output_dir=self.root / "second",
            summary=self.summary_path,
            artifacts=self.specs(),
        )
        self.assertEqual(first, second)
        first_files = sorted((path.relative_to(self.root / "first"), path.read_bytes()) for path in (self.root / "first").rglob("*") if path.is_file())
        second_files = sorted((path.relative_to(self.root / "second"), path.read_bytes()) for path in (self.root / "second").rglob("*") if path.is_file())
        self.assertEqual(first_files, second_files)

    def test_fails_closed_before_writing_on_summary_binding_mismatch(self):
        malformed = dict(self.summary)
        malformed["report_identity"] = dict(self.summary["report_identity"])
        malformed["report_identity"]["a1"] = {"canonical_sha256": "0" * 64}
        summary_path = self.source / "mismatch.json"
        summary_path.write_text(json.dumps(malformed), encoding="utf-8")
        output = self.root / "package"
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError, "canonical SHA-256"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=output,
                summary=summary_path,
                artifacts=self.specs(),
            )
        self.assertFalse(output.exists())

    def test_refuses_overwrite_and_path_escape(self):
        output = self.root / "package"
        perf_abba_package.package_artifacts(
            change_id="0238",
            output_dir=output,
            summary=self.summary_path,
            artifacts=self.specs(),
        )
        with self.assertRaisesRegex(perf_abba_package.ArtifactPackagingError, "overwrite"):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=output,
                summary=self.summary_path,
                artifacts=self.specs(),
            )

        with self.assertRaisesRegex(perf_abba_package.ArtifactPackagingError, "escapes"):
            perf_abba_package.package_artifacts(
                change_id="0239",
                output_dir=self.root / "escape",
                summary=self.summary_path,
                artifacts=self.specs(),
                artifact_names={"a1": "../outside.json.zst"},
            )
        self.assertFalse((self.root / "outside.json.zst").exists())

    def test_rejects_duplicate_or_missing_roles(self):
        with self.assertRaisesRegex(perf_abba_package.ArtifactPackagingError, "duplicate"):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=self.root / "duplicate",
                summary=self.summary_path,
                artifacts=[
                    ("a1", self.report_paths["a1"]),
                    ("a1-control", self.report_paths["a1"]),
                    ("b1", self.report_paths["b1"]),
                    ("b2", self.report_paths["b2"]),
                    ("a2", self.report_paths["a2"]),
                ],
            )
        with self.assertRaisesRegex(perf_abba_package.ArtifactPackagingError, "missing"):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=self.root / "missing",
                summary=self.summary_path,
                artifacts={role: path for role, path in self.specs().items() if role != "a2"},
            )


if __name__ == "__main__":
    unittest.main()
