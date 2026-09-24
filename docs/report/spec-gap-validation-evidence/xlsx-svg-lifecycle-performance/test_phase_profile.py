"""Adversarial checks for the exploratory phase receipt contract."""

import json
import hashlib
from pathlib import Path
import re
import shlex
import shutil
import subprocess
from tempfile import TemporaryDirectory
import unittest

import phase_summarize
import phase_verify


class HostProbeTests(unittest.TestCase):
    @unittest.skipUnless(shutil.which("awk"), "host metadata reader requires awk")
    def test_phase_host_metadata_awk_programs_execute_fixture_fields(self):
        script = Path(__file__).with_name("run_phase_profile.sh").read_text()
        cases = [
            (r"awk -F: '([^']+)'", ["-F:"], "model name : Fixture CPU\n", "Fixture CPU"),
            (r"awk '([^']+)'", [], "MemTotal: 12345 kB\n", "12345 kB"),
        ]
        for pattern, options, source, expected in cases:
            with self.subTest(expected=expected):
                program = re.search(pattern, script)
                self.assertIsNotNone(program, "phase host metadata program is missing")
                run = subprocess.run(
                    ["awk", *options, program.group(1)],
                    input=source,
                    capture_output=True,
                    text=True,
                )
                self.assertEqual(run.returncode, 0, run.stderr)
                self.assertEqual(run.stdout.strip(), expected)

    @unittest.skipUnless(
        shutil.which("df") and shutil.which("tail"),
        "storage metadata readers require df and tail",
    )
    def test_phase_storage_probe_executes_against_fixture_path(self):
        script = Path(__file__).with_name("run_phase_profile.sh").read_text()
        line = next(
            line for line in script.splitlines() if 'storage=$(df -P "$ROOT"' in line
        )
        with TemporaryDirectory(prefix="litchi-xlsx-phase-host-") as directory:
            run = subprocess.run(
                ["bash", "-c", f"ROOT={shlex.quote(directory)}\n{line}\n"],
                capture_output=True,
                text=True,
            )
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertRegex(run.stdout, r"^storage=\S.*\n$")


class PhaseReceiptTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="litchi-xlsx-phase-")
        self.addCleanup(temporary.cleanup)
        self.results = Path(temporary.name)

    def phase(self):
        return {
            "elapsed_ns": 1,
            "requested_alloc_bytes": 5,
            "direct_allocated_bytes": 2,
            "realloc_new_bytes": 3,
            "realloc_old_bytes": 0,
            "deallocated_bytes": 0,
            "live_before_bytes": 10,
            "live_after_bytes": 15,
            "live_delta_bytes": 5,
            "retained_live_bytes_after": 15,
            "peak_live_delta_bytes": 5,
            "alloc_balance_ok": True,
            "alloc_invalid": False,
            "alloc_failed": 0,
        }

    def write_receipt(self, pictures, process):
        phases = {name: self.phase() for name in phase_verify.PHASES}
        payload = {
            "schema": phase_verify.SCHEMA,
            "lane": f"multi_picture_same_drawing_{pictures}",
            "picture_count": pictures,
            "input_bytes": 100 + pictures,
            "input_hash_fnv1a64": 200 + pictures,
            "input_sha256": "a" * 64,
            "warmup": 2,
            "sample_count": 20,
            "expected_success": True,
            "phase_order": list(phase_verify.PHASES),
            "allocation_note": phase_verify.ALLOCATION_NOTE,
            "semantic_checks": list(phase_verify.SEMANTIC_CHECKS),
            "samples": [
                {"semantic_ok": True, "output_exact": True, "phases": phases}
                for _ in range(20)
            ],
        }
        path = self.results / f"phase_{pictures}-p{process}.json"
        path.write_text(json.dumps(payload))
        path.with_suffix(".stderr.log").write_bytes(b"")
        path.with_suffix(".time.txt").write_text(
            "Maximum resident set size (kbytes): 100\nExit status: 0\n"
        )

    def write_bundle(self):
        for pictures in phase_verify.PICTURE_COUNTS:
            for process in phase_verify.PROCESSES:
                self.write_receipt(pictures, process)
        for name in ("binary.sha256", "binary-after.sha256"):
            (self.results / name).write_text("a" * 64 + "  /disposable/binary\n")
        for name in ("metadata-before.json", "metadata-after.json"):
            (self.results / name).write_text('{"packages": []}\n')
        metadata_digest = hashlib.sha256((self.results / "metadata-before.json").read_bytes()).hexdigest()
        manifest = (
            "format=xlsx-svg-lifecycle-profile-build-source-v2\n"
            + "git_commit="
            + "a" * 40
            + "\nmetadata_sha256="
            + metadata_digest
            + "\npackage=fixture\t0.0.0\tpath\tfixture/Cargo.toml\t"
            + "a" * 64
            + "\t0\t"
            + "a" * 64
            + "\nfile=fixture\t0.0.0\tfixture/Cargo.toml\t"
            + "a" * 64
            + "\n"
        )
        for name in ("source-manifest-before.txt", "source-manifest-after.txt"):
            (self.results / name).write_text(manifest)
        manifest_digest = hashlib.sha256(manifest.encode()).hexdigest()
        (self.results / "source-provenance.txt").write_text(
            f"source_manifest_before_sha256={manifest_digest}\n"
            f"source_manifest_after_sha256={manifest_digest}\n"
            f"git_head={'a' * 40}\n"
            f"approved_source_pin={'b' * 40}\n"
        )
        (self.results / "commands.txt").write_text(
            "\n".join(
                f"run=/usr/bin/time -v /disposable/binary --pictures {pictures} --warmup 2 --samples 20 (fresh_process={process})"
                for pictures in phase_verify.PICTURE_COUNTS
                for process in phase_verify.PROCESSES
            )
            + "\n"
        )
        (self.results / "build.log").write_text("build\n")
        (self.results / "build-provenance.txt").write_text(
            "\n".join(
                [
                    "binary=/disposable/binary",
                    "a" * 64 + "  /disposable/binary",
                    "rustc -vV:",
                    "rustc 1.89.0 (fixture)",
                    "cargo=cargo 1.89.0 (fixture)",
                    "target=/disposable/target",
                    "cargo_incremental=0",
                    "compiler_flags=RUSTFLAGS=unset CARGO_ENCODED_RUSTFLAGS=unset RUSTC_BOOTSTRAP=unset",
                    "allocator=CountingAllocator (process-local GlobalAlloc observer)",
                    "phase_scope=open,stages,commit,firstsave,reopen_secondsave,validation",
                    f"git_head={'a' * 40}",
                    f"approved_source_pin={'b' * 40}",
                    "os=fixture",
                    "cpu_model=Fixture CPU",
                    "core_count=3",
                    "memory_total=1 kB",
                    "storage=fixture",
                    "environment=CARGO_TARGET_DIR=/disposable/target CARGO_INCREMENTAL=0",
                ]
            )
            + "\n"
        )
    def test_valid_bundle_and_summary(self):
        self.write_bundle()
        report = self.results / "report.md"
        report.write_text(phase_summarize.summarize(self.results))
        verified = phase_verify.verify_results(self.results, report)
        self.assertEqual(verified["receipts"], 9)
        summary = phase_summarize.summarize(self.results)
        self.assertIn("phase decomposition", summary)
        self.assertIn("`reopen_secondsave`", summary)
        self.assertIn("must not be summed", summary)
        self.assertIn("Largest observed phase medians", summary)
        self.assertIn("Largest phase-local requested-allocation medians", summary)
        report.write_text(report.read_text() + "tampered\n")
        with self.assertRaisesRegex(ValueError, "not derived from raw receipts"):
            phase_verify.verify_results(self.results, report)

    def test_unknown_numeric_lane_is_rejected(self):
        self.write_bundle()
        (self.results / "phase_17-p1.json").write_text("{}")
        with self.assertRaisesRegex(ValueError, "unknown phase picture count"):
            phase_verify.verify_results(self.results)

    def test_live_equation_is_rejected(self):
        self.write_bundle()
        path = self.results / "phase_256-p2.json"
        payload = json.loads(path.read_text())
        payload["samples"][0]["phases"]["commit"]["live_delta_bytes"] = 4
        path.write_text(json.dumps(payload))
        with self.assertRaisesRegex(ValueError, "live delta failed"):
            phase_verify.verify_results(self.results)
        payload = json.loads(path.read_text())
        del payload["input_hash_fnv1a64"]
        path.write_text(json.dumps(payload))
        with self.assertRaisesRegex(ValueError, "receipt fields changed"):
            phase_verify.verify_results(self.results)

    def test_command_flags_are_rejected_when_changed(self):
        self.write_bundle()
        path = self.results / "commands.txt"
        path.write_text(path.read_text().replace("--warmup 2", "--warmup 1", 1))
        with self.assertRaisesRegex(ValueError, "command flags or process invocation changed"):
            phase_verify.verify_results(self.results)

        self.write_bundle()
        path.write_text(path.read_text().replace("--samples 20", "--samples 19", 1))
        with self.assertRaisesRegex(ValueError, "command flags or process invocation changed"):
            phase_verify.verify_results(self.results)

        self.write_bundle()
        path.write_text(path.read_text().replace("--samples 20", "--samples 20 --unexpected", 1))
        with self.assertRaisesRegex(ValueError, "command flags or process invocation changed"):
            phase_verify.verify_results(self.results)

    def test_build_provenance_contract_is_rejected_when_changed(self):
        mutations = (
            ("cargo_incremental=", "cargo incremental setting changed"),
            ("compiler_flags=", "compiler flags changed"),
            ("allocator=", "allocator provenance changed"),
            ("phase_scope=", "phase scope changed"),
        )
        for prefix, message in mutations:
            with self.subTest(prefix=prefix):
                self.write_bundle()
                path = self.results / "build-provenance.txt"
                lines = path.read_text().splitlines()
                index = next(index for index, line in enumerate(lines) if line.startswith(prefix))
                lines[index] = prefix + "tampered"
                path.write_text("\n".join(lines) + "\n")
                with self.assertRaisesRegex(ValueError, message):
                    phase_verify.verify_results(self.results)


if __name__ == "__main__":
    unittest.main()
