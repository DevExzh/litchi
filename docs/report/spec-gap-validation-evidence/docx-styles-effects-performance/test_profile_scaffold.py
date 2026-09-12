"""Executable guards for the bounded styles-effects profile scaffold."""

from pathlib import Path
import os
import subprocess
from tempfile import TemporaryDirectory
import unittest
import hashlib
import json
import shlex
import zipfile
import sys

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import verify_profile


class ProfileScaffoldTests(unittest.TestCase):
    def test_publication_attribution_split_and_subphase_allocator_gate_are_wired(self):
        adapter = (HERE / "profile-harness/profile_adapter.rs").read_text()
        support = (HERE / "harness/support.rs").read_text()
        main = (HERE / "profile-harness/main.rs").read_text()
        runner = (HERE / "run_profile.sh").read_text()
        self.assertIn("fn timed_publication", adapter)
        self.assertIn("support::begin_phase_window()", adapter)
        self.assertIn("apply_before.phase_delta", adapter)
        self.assertIn("serialize_before.phase_delta", adapter)
        self.assertIn("static PHASE_PEAK_BYTES", support)
        self.assertIn('"apply_allocation"', main)
        self.assertIn('"serialize_allocation"', main)
        self.assertIn(verify_profile.PROFILE_SOURCE_COMMIT, main)
        self.assertIn(f"SOURCE_COMMIT={verify_profile.PROFILE_SOURCE_COMMIT}", runner)

        valid = {field: 0 for field in verify_profile.ALLOCATION_FIELDS}
        valid.update(
            {
                "direct_allocated_bytes": 2,
                "requested_alloc_bytes": 2,
                "live_before": 3,
                "live_after": 5,
                "peak_live_delta": 2,
                "alloc_balance_ok": True,
                "alloc_invalid": False,
            }
        )
        verify_profile.verify_allocation_delta(valid, "apply_allocation")

        impossible_peak = dict(valid, peak_live_delta=1)
        with self.assertRaisesRegex(AssertionError, "peak live delta"):
            verify_profile.verify_allocation_delta(impossible_peak, "apply_allocation")

        unbalanced = dict(valid, live_after=6)
        with self.assertRaisesRegex(AssertionError, "live-byte equation"):
            verify_profile.verify_allocation_delta(unbalanced, "apply_allocation")

    def test_host_probe_executes_and_emits_parseable_observations(self):
        with TemporaryDirectory(prefix="litchi-docx-styles-host-") as directory:
            output = Path(directory) / "host-before.txt"
            run = subprocess.run(
                ["bash", str(HERE / "host_probe.sh"), str(output)],
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(run.returncode, 0, run.stderr)
            values = dict(
                line.split("=", 1)
                for line in output.read_text().splitlines()
                if "=" in line
            )
            required = {
                "schema",
                "utc",
                "pid",
                "kernel",
                "os",
                "cpu_model",
                "logical_cpus",
                "memory_total_kib",
                "memory_available_kib",
                "load_1m",
                "time_version",
                "affinity",
            }
            self.assertTrue(required <= values.keys())
            self.assertEqual(values["schema"], "docx-styles-effects-host-v1")
            if values["logical_cpus"] != "unavailable":
                self.assertGreater(int(values["logical_cpus"]), 0)
            if values["memory_total_kib"] != "unavailable":
                self.assertGreater(int(values["memory_total_kib"]), 0)
            if values["load_1m"] != "unavailable":
                self.assertGreaterEqual(float(values["load_1m"]), 0.0)
            self.assertTrue(values["time_version"])

    def test_profile_runner_freeze_gate_runs_before_creating_output(self):
        with TemporaryDirectory(prefix="litchi-docx-styles-profile-gate-") as directory:
            results = Path(directory) / "results"
            environment = {
                key: value
                for key, value in os.environ.items()
                if key not in {"PROFILE_FROZEN", "DOCX_STYLES_EFFECTS_PROFILE_API_WIRED"}
            }
            environment["DOCX_STYLES_EFFECTS_PROFILE_RESULTS"] = str(results)
            run = subprocess.run(
                ["bash", str(HERE / "run_profile.sh")],
                cwd=HERE,
                env=environment,
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(run.returncode, 2, run.stderr)
            self.assertFalse(results.exists())

    def test_profile_runner_api_gate_runs_before_creating_output(self):
        with TemporaryDirectory(prefix="litchi-docx-styles-profile-api-gate-") as directory:
            results = Path(directory) / "results"
            environment = dict(os.environ)
            environment["PROFILE_FROZEN"] = "1"
            environment.pop("DOCX_STYLES_EFFECTS_PROFILE_API_WIRED", None)
            environment["DOCX_STYLES_EFFECTS_PROFILE_RESULTS"] = str(results)
            run = subprocess.run(
                ["bash", str(HERE / "run_profile.sh")],
                cwd=HERE,
                env=environment,
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(run.returncode, 2, run.stderr)
            self.assertFalse(results.exists())

    def test_time_receipt_requires_user_system_and_elapsed_markers(self):
        valid = """\
User time (seconds): 0.01
System time (seconds): 0.00
Elapsed (wall clock) time (h:mm:ss or m:ss): 0:00.02
Maximum resident set size (kbytes): 128
Exit status: 0
"""
        with TemporaryDirectory(prefix="litchi-docx-styles-time-") as directory:
            time_path = Path(directory) / "sample.time.txt"
            stderr_path = Path(directory) / "sample.stderr.log"
            time_path.write_text(valid)
            stderr_path.write_bytes(b"")
            verify_profile.verify_time_sidecar(time_path, stderr_path)
            time_path.write_text(valid.replace("System time (seconds): 0.00\n", ""))
            with self.assertRaises(AssertionError):
                verify_profile.verify_time_sidecar(time_path, stderr_path)

    def test_generated_fixture_verifier_binds_exact_external_files(self):
        with TemporaryDirectory(prefix="litchi-docx-styles-generated-") as directory:
            root = Path(directory)
            generated = root / "generated-fixtures"
            generated.mkdir()
            rows = []
            for owner in ("main", "glossary"):
                member = "word/glossary/stylesWithEffects.xml" if owner == "glossary" else "word/stylesWithEffects.xml"
                for scale, target in (("64k", 64 * 1024), ("1m", 1024 * 1024)):
                    resource = generated / f"{owner}-{scale}.stylesWithEffects.xml"
                    package = generated / f"{owner}-{scale}.docx"
                    xml = (b"<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                           + b"x" * (target - 90) + b"</w:styles>")
                    resource.write_bytes(xml)
                    with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                        archive.writestr(member, xml)
                    digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
                    rows.append({
                        "schema": "docx-styles-effects-generated-fixture-v1",
                        "owner": owner,
                        "scale": scale,
                        "member": member,
                        "resource_bytes": len(xml),
                        "resource_sha256": digest(resource),
                        "package_bytes": package.stat().st_size,
                        "package_sha256": digest(package),
                        "xml_events": 3,
                        "xml_depth": 1,
                        "style_count": 1,
                        "opaque_marker_bytes": 1,
                        "resource_path": str(resource),
                        "package_path": str(package),
                    })
            manifest = root / "generated-fixtures.json"
            manifest.write_text(json.dumps({"schema": verify_profile.GENERATED_SCHEMA, "source_commit": verify_profile.SOURCE_COMMIT, "fixtures": rows}))
            verify_profile.verify_generated(manifest)
            resource = generated / "main-64k.stylesWithEffects.xml"
            resource.write_bytes(resource.read_bytes() + b"tamper")
            with self.assertRaises(AssertionError):
                verify_profile.verify_generated(manifest)

    def test_profile_runner_retention_policy_executes_failure_closed_cleanup(self):
        """Exercise the sourced EXIT cleanup used by the real profile runner."""
        helper = HERE / "cleanup_target.sh"
        self.assertTrue(helper.is_file())

        def run_cleanup(results, target):
            command = " ".join(
                (
                    "set -euo pipefail;",
                    f"source {shlex.quote(str(helper))};",
                    f"RESULTS={shlex.quote(str(results))};",
                    f"TARGET={shlex.quote(str(target))};",
                    f"SUCCESS_SENTINEL={shlex.quote(str(results / 'profile-success.sentinel'))};",
                    "cleanup_profile_target",
                )
            )
            return subprocess.run(
                ["bash", "-c", command],
                capture_output=True,
                text=True,
                check=False,
            )

        with TemporaryDirectory(prefix="litchi-docx-styles-cleanup-") as directory:
            root = Path(directory)
            results = root / "results"
            target = root / "target"
            results.mkdir()
            target.mkdir()
            (target / "build-marker").write_text("built")

            # A build/sample failure has no success receipt. The target must
            # remain available for diagnosis.
            run = run_cleanup(results, target)
            self.assertEqual(run.returncode, 0, run.stderr)
            self.assertTrue(target.exists())
            self.assertIn("retained", run.stderr)

            # A verifier failure or tampered receipt must also retain it.
            verification = results / "verification.json"
            verification.write_text('{"passed": false}\n')
            sentinel = results / "profile-success.sentinel"
            sentinel.write_text("verification_sha256=not-the-verification-digest\n")
            run = run_cleanup(results, target)
            self.assertEqual(run.returncode, 0, run.stderr)
            self.assertTrue(target.exists())
            self.assertIn("retained", run.stderr)

            # Only a digest matching the completed verifier output authorizes
            # cleanup of the target directory.
            digest = hashlib.sha256(verification.read_bytes()).hexdigest()
            sentinel.write_text(f"verification_sha256={digest}\n")
            run = run_cleanup(results, target)
            self.assertEqual(run.returncode, 0, run.stderr)
            self.assertFalse(target.exists())

    def test_replay_paths_must_be_external_and_disjoint(self):
        with TemporaryDirectory(prefix="litchi-docx-styles-paths-") as directory:
            root = Path(directory) / "checkout"
            evidence = root / "docs"
            root.mkdir()
            evidence.mkdir()
            external = Path(directory) / "results"
            target = Path(directory) / "target"
            verify_profile.verify_external_paths(root, evidence, external, target)
            with self.assertRaises(AssertionError):
                verify_profile.verify_external_paths(root, evidence, evidence / "results", target)
            with self.assertRaises(AssertionError):
                verify_profile.verify_external_paths(root, evidence, external, external / "target")

    def test_profile_runner_rejects_inherited_flags_before_output(self):
        with TemporaryDirectory(prefix="litchi-docx-styles-profile-flags-") as directory:
            results = Path(directory) / "results"
            environment = dict(os.environ)
            environment["PROFILE_FROZEN"] = "1"
            environment["DOCX_STYLES_EFFECTS_PROFILE_API_WIRED"] = "1"
            environment["RUSTFLAGS"] = "-Copt-level=0"
            environment["DOCX_STYLES_EFFECTS_PROFILE_RESULTS"] = str(results)
            run = subprocess.run(
                ["bash", str(HERE / "run_profile.sh")],
                cwd=HERE,
                env=environment,
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(run.returncode, 2, run.stderr)
            self.assertFalse(results.exists())


if __name__ == "__main__":
    unittest.main()
