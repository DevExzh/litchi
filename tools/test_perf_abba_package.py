import copy
import hashlib
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tools import perf_abba_package, perf_abba_summary
from tools.test_perf_abba_summary import (
    four_legs,
    with_filesystem_evidence,
    with_operation_metrics,
)


ZSTD = shutil.which("zstd")


def canonical(value):
    return perf_abba_package.canonical_json(value, "fixture")


def make_fixture():
    roles = ("a1", "b1", "b2", "a2")
    reports = dict(zip(roles, four_legs()))
    summary = perf_abba_summary.summarize_reports(
        [reports[role] for role in roles]
    )
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
            perf_abba_package.ArtifactPackagingError, "summary"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=output,
                summary=summary_path,
                artifacts=self.specs(),
            )
        self.assertFalse(output.exists())

    def test_rejects_malformed_parallel_metrics_before_publication(self):
        self.replace_report("a1", lambda report: report.update(parallel_metrics={}))
        output = self.root / "parallel-metrics-invalid"
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError,
            "parallel_metrics validation failed",
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=output,
                summary=self.summary_path,
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
            ("schema_version", lambda report: report.update(schema_version=2), "canonical"),
            ("tool", lambda report: report["tool"].update(version="other"), "canonical"),
            (
                "configuration",
                lambda report: report["configuration"].update(cases=["different"]),
                "canonical",
            ),
            (
                "environment",
                lambda report: report["environment"].update(git_revision="wrong"),
                "canonical",
            ),
            (
                "results",
                lambda report: report["results"][0]["corpus"].update(name="different"),
                "canonical",
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

    def test_rejects_nested_operation_and_filesystem_identity_tamper(self):
        mutations = (
            (
                "operation_metrics.sample_count",
                with_operation_metrics,
                lambda report: report["results"][0]["operation_metrics"].update(
                    sample_count=14
                ),
            ),
            (
                "operation_metrics.schema",
                with_operation_metrics,
                lambda report: report["results"][0]["operation_metrics"].update(
                    schema=2
                ),
            ),
            (
                "operation_metrics.scope",
                with_operation_metrics,
                lambda report: report["results"][0]["operation_metrics"]["source"].update(
                    counter_scope="untimed_source_replay_only"
                ),
            ),
            (
                "filesystem_evidence.corpus",
                with_filesystem_evidence,
                lambda report: report["filesystem_evidence"][0]["corpus"].update(
                    name="changed-corpus"
                ),
            ),
            (
                "filesystem_evidence.tool",
                with_filesystem_evidence,
                lambda report: report["filesystem_evidence"][0]["tool"].update(
                    version="changed-tool"
                ),
            ),
            (
                "filesystem_evidence.configuration",
                with_filesystem_evidence,
                lambda report: report["filesystem_evidence"][0]["configuration"].update(
                    warmup_iterations_per_case=2
                ),
            ),
        )
        for name, prepare, mutate in mutations:
            with self.subTest(name=name):
                prepared = prepare(four_legs())
                self.reports = dict(zip(("a1", "b1", "b2", "a2"), prepared))
                for role, report in self.reports.items():
                    self.report_paths[role].write_bytes(
                        (json.dumps(report, sort_keys=True, indent=2) + "\n").encode()
                    )
                self.summary = perf_abba_summary.summarize_reports(prepared)
                self.write_summary()
                self.replace_report("a1", mutate)
                with self.assertRaisesRegex(
                    perf_abba_package.ArtifactPackagingError,
                    "recompute canonical ABBA summary|canonical ABBA",
                ):
                    perf_abba_package.package_artifacts(
                        change_id="0238",
                        output_dir=self.root / f"nested-{name.replace('.', '-')}",
                        summary=self.summary_path,
                        artifacts=self.specs(),
                    )
                self.tearDown()
                self.setUp()

    def test_rejects_summary_implementation_leg_mismatch(self):
        self.summary["implementation_identity"]["control"]["git_revision"] = "wrong"
        self.write_summary()
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError, "canonical"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=self.root / "implementation-mismatch",
                summary=self.summary_path,
                artifacts=self.specs(),
            )

    def test_rejects_full_summary_tamper_not_covered_by_report_bindings(self):
        self.summary["results"][0]["elapsed_ns"]["candidate_reduction_percent"]["a1_to_b1"][
            "p50"
        ] = 999.0
        self.write_summary()
        output = self.root / "tampered-summary"
        with self.assertRaisesRegex(
            perf_abba_package.ArtifactPackagingError, "canonical ABBA"
        ):
            perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=output,
                summary=self.summary_path,
                artifacts=self.specs(),
            )
        self.assertFalse(output.exists())

    def test_rejects_reusing_one_raw_report_for_two_roles(self):
        # Re-encode the same JSON object with different bytes to ensure the
        # canonical identity check catches semantic report reuse as well as
        # an exact path/byte-for-byte copy.
        reused_path = self.source / "reencoded-a1.json"
        reused_path.write_bytes(
            json.dumps(self.reports["a1"], sort_keys=True, separators=(",", ":")).encode()
        )
        self.report_paths["a2"] = reused_path
        self.reports["a2"] = copy.deepcopy(self.reports["a1"])
        self.summary = perf_abba_summary.summarize_reports(
            [self.reports[role] for role in ("a1", "b1", "b2", "a2")]
        )
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

        def fail_on_second_write(directory_fd, name, data, location):
            nonlocal call_count
            call_count += 1
            if call_count == 2:
                file_descriptor = perf_abba_package.os.open(
                    name,
                    perf_abba_package.os.O_WRONLY
                    | perf_abba_package.os.O_CREAT
                    | perf_abba_package.os.O_EXCL,
                    mode=0o600,
                    dir_fd=directory_fd,
                )
                try:
                    perf_abba_package.os.write(file_descriptor, data[:1])
                finally:
                    perf_abba_package.os.close(file_descriptor)
                raise OSError("injected write failure")
            return real_write(directory_fd, name, data, location)

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
        real_mkdir = perf_abba_package.os.mkdir
        call_count = 0

        def fail_on_second_mkdir(path, *args, **kwargs):
            nonlocal call_count
            call_count += 1
            if call_count == 2:
                raise OSError("injected mkdir failure")
            return real_mkdir(path, *args, **kwargs)

        with mock.patch.object(
            perf_abba_package.os, "mkdir", side_effect=fail_on_second_mkdir
        ):
            with self.assertRaisesRegex(
                perf_abba_package.ArtifactPackagingError, "cannot open output directory"
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

    def test_publication_stays_on_held_directory_after_path_swap(self):
        output = self.root / "directory-race"
        moved = self.root / "directory-race-original"
        real_link = perf_abba_package.os.link
        swapped = False

        def swap_before_first_publish(source, destination, *args, **kwargs):
            nonlocal swapped
            if not swapped:
                swapped = True
                output.rename(moved)
                output.mkdir()
            return real_link(source, destination, *args, **kwargs)

        with mock.patch.object(
            perf_abba_package.os, "link", side_effect=swap_before_first_publish
        ):
            manifest = perf_abba_package.package_artifacts(
                change_id="0238",
                output_dir=output,
                summary=self.summary_path,
                artifacts=self.specs(),
            )
        self.assertEqual(list(output.iterdir()), [])
        self.assertEqual(
            json.loads((moved / manifest["manifest_path"]).read_text()), manifest
        )
        self.assertFalse(any(path.name.startswith(".0238.staging-") for path in moved.iterdir()))

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
