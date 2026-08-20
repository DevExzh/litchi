import copy
import hashlib
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

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
            "leg": role,
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
            "results": [
                {
                    "case": "fixture",
                    "corpus": {"name": "fixture", "shape": "tiny"},
                    "elapsed_ns": {"samples": [1]},
                }
            ],
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
            "stable": {
                "rustc_version": "rustc 1.95.0 (test)",
                "git_worktree_dirty": False,
            },
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
        "results": [
            {
                "case": "fixture",
                "corpus": {"name": "fixture", "shape": "tiny"},
                "shape": "tiny",
                "identity": {
                    "corpus": canonical({"name": "fixture", "shape": "tiny"}).decode()
                },
            }
        ],
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

    def write_summary(self):
        self.summary_path.write_text(
            json.dumps(self.summary, sort_keys=True, indent=2) + "\n", encoding="utf-8"
        )

    def replace_report(self, role, mutate):
        report = copy.deepcopy(self.reports[role])
        mutate(report)
        self.reports[role] = report
        path = self.source / f"report-{role}.json"
        path.write_bytes((json.dumps(report, sort_keys=True, indent=2) + "\n").encode())
        self.report_paths[role] = path
        self.summary["report_identity"][role] = {
            "canonical_sha256": hashlib.sha256(canonical(report)).hexdigest()
        }
        self.write_summary()

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
        zstd_identity = manifest["compression"]["executable"]
        self.assertEqual(Path(zstd_identity["path"]), Path(ZSTD).resolve())
        self.assertTrue(zstd_identity["version"])
        self.assertEqual(
            zstd_identity["sha256"], hashlib.sha256(Path(ZSTD).resolve().read_bytes()).hexdigest()
        )
        self.assertEqual(zstd_identity["bytes"], Path(ZSTD).resolve().stat().st_size)
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

    def test_cross_checks_report_schema_tool_configuration_environment_and_results(self):
        mutations = (
            ("schema_version", lambda report: report.update(schema_version=2), "schema_version"),
            ("tool", lambda report: report["tool"].update(version="other"), "tool identity"),
            (
                "configuration",
                lambda report: report["configuration"].update(cases=["different"]),
                "configuration",
            ),
            (
                "environment",
                lambda report: report["environment"].update(git_revision="wrong"),
                "environment",
            ),
            (
                "results",
                lambda report: report["results"][0]["corpus"].update(name="different"),
                "case/corpus/shape",
            ),
        )
        for label, mutate, message in mutations:
            with self.subTest(label=label):
                self.replace_report("a1", mutate)
                with self.assertRaisesRegex(
                    perf_abba_package.ArtifactPackagingError, message
                ):
                    perf_abba_package.package_artifacts(
                        change_id="0238",
                        output_dir=self.root / f"cross-check-{label}",
                        summary=self.summary_path,
                        artifacts=self.specs(),
                    )
                # Restore a clean fixture for the next mutation.
                self.tearDown()
                self.setUp()

    def test_rejects_summary_implementation_leg_mismatch(self):
        self.summary["implementation_identity"]["control"]["git_revision"] = "wrong"
        self.write_summary()
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError, "implementation leg"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=self.root / "implementation-mismatch",
                summary=self.summary_path,
                artifacts=self.specs(),
            )

    def test_rejects_reusing_one_raw_report_for_two_roles(self):
        # Re-encode the same JSON object with different bytes to ensure the
        # canonical identity check catches semantic report reuse as well as
        # an exact path/byte-for-byte copy.
        reused_path = self.source / "reencoded-a1.json"
        reused_path.write_bytes(
            json.dumps(self.reports["a1"], sort_keys=True, separators=(",", ":")).encode()
        )
        self.report_paths["a2"] = reused_path
        self.summary["report_identity"]["a2"] = dict(self.summary["report_identity"]["a1"])
        self.write_summary()
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError, "reused"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=self.root / "reused",
                summary=self.summary_path,
                artifacts=self.specs(),
            )

    def test_publication_failure_removes_partial_artifacts_and_staging(self):
        output = self.root / "atomic-publication"
        real_link = perf_abba_package.os.link
        call_count = 0

        def fail_on_second_link(source, destination, *args, **kwargs):
            nonlocal call_count
            call_count += 1
            if call_count == 2:
                raise OSError("injected link failure")
            return real_link(source, destination, *args, **kwargs)

        with mock.patch.object(perf_abba_package.os, "link", side_effect=fail_on_second_link):
            with self.assertRaisesRegex(
                perf_abba_package.ArtifactPackagingError, "cannot publish"
            ):
                perf_abba_package.package_artifacts(
                    change_id="0238",
                    output_dir=output,
                    summary=self.summary_path,
                    artifacts=self.specs(),
                )
        self.assertFalse(output.exists())

    def test_staging_write_failure_removes_partial_staging_and_directories(self):
        output = self.root / "atomic-staging"
        real_write = perf_abba_package._write_exclusive
        call_count = 0

        def fail_on_second_write(path, data, location):
            nonlocal call_count
            call_count += 1
            if call_count == 2:
                path.write_bytes(data[:1])
                raise OSError("injected write failure")
            return real_write(path, data, location)

        with mock.patch.object(
            perf_abba_package, "_write_exclusive", side_effect=fail_on_second_write
        ):
            with self.assertRaisesRegex(
                perf_abba_package.ArtifactPackagingError, "artifact publication failed"
            ):
                perf_abba_package.package_artifacts(
                    change_id="0238",
                    output_dir=output,
                    summary=self.summary_path,
                    artifacts=self.specs(),
                )
        self.assertFalse(output.exists())

    def test_mkdir_failure_removes_directories_created_before_the_failure(self):
        output = self.root / "mkdir-failure" / "nested"
        real_mkdir = Path.mkdir
        call_count = 0

        def fail_on_second_mkdir(path, *args, **kwargs):
            nonlocal call_count
            call_count += 1
            if call_count == 2:
                raise OSError("injected mkdir failure")
            return real_mkdir(path, *args, **kwargs)

        with mock.patch.object(Path, "mkdir", autospec=True, side_effect=fail_on_second_mkdir):
            with self.assertRaisesRegex(
                perf_abba_package.ArtifactPackagingError, "cannot create output directory"
            ):
                perf_abba_package.package_artifacts(
                    change_id="0238",
                    output_dir=output,
                    summary=self.summary_path,
                    artifacts=self.specs(),
                )
        self.assertFalse(output.exists())
        self.assertFalse(output.parent.exists())

    def test_zstd_version_probe_failure_leaves_no_output(self):
        fake_zstd = self.root / "fake-zstd"
        fake_zstd.write_text("#!/bin/sh\nexit 19\n", encoding="utf-8")
        fake_zstd.chmod(0o755)
        output = self.root / "zstd-probe-failure"
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError, "version probe failed"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=output,
                summary=self.summary_path,
                artifacts=self.specs(),
                zstd_executable=fake_zstd,
            )
        self.assertFalse(output.exists())

    def test_rejects_output_directory_symlink(self):
        real_output = self.root / "real-output"
        real_output.mkdir()
        symlink_output = self.root / "symlink-output"
        try:
            symlink_output.symlink_to(real_output, target_is_directory=True)
        except OSError as error:
            self.skipTest(f"symlink unavailable: {error}")
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError, "must not be a symlink"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=symlink_output,
                summary=self.summary_path,
                artifacts=self.specs(),
            )
        self.assertEqual(list(real_output.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
