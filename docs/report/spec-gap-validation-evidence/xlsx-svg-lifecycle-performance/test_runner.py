"""Runner refusal tests use a failing Cargo stub, never a profiling build."""

import os
from pathlib import Path
import re
import shutil
import subprocess
from tempfile import TemporaryDirectory
import unittest


class RunnerInputTests(unittest.TestCase):
    @unittest.skipUnless(shutil.which("awk"), "host metadata reader requires awk")
    def test_embedded_host_metadata_programs_decode_proc_fields(self):
        script = Path(__file__).with_name("run_profile.sh").read_text()
        cases = [
            (r"awk -F: '([^']+)'", ["-F:"], "model name : Fixture CPU\n", "Fixture CPU"),
            (r"awk '([^']+)'", [], "MemTotal: 12345 kB\n", "12345 kB"),
        ]
        for pattern, options, source, expected in cases:
            with self.subTest(expected=expected):
                program = re.search(pattern, script)
                self.assertIsNotNone(program, "host metadata program is missing")
                run = subprocess.run(
                    ["awk", *options, program.group(1)], input=source,
                    capture_output=True, text=True,
                )
                self.assertEqual(run.returncode, 0, run.stderr)
                self.assertEqual(run.stdout.strip(), expected)

    def test_declared_manifest_extras_exist_in_committed_checkout(self):
        here = Path(__file__).resolve().parent
        root = here.parents[3]
        script = (here / "run_profile.sh").read_text()
        variables = {"ROOT": root, "HERE": here}
        for name, relative in re.findall(r'^([A-Z_]+)="\$HERE/([^"\n]+)"$', script, re.M):
            variables[name] = here / relative
        arguments = re.findall(r'--extra "([^"\n]+)"', script)
        self.assertTrue(arguments, "runner has no declared manifest inputs")
        paths = set()
        for argument in arguments:
            parsed = re.fullmatch(r'\$([A-Z_]+)(/.*)?', argument)
            self.assertIsNotNone(parsed, f"unhandled input expression: {argument}")
            variable, suffix = parsed.groups()
            self.assertIn(variable, variables, f"unresolved input variable: {argument}")
            path = variables[variable] / (suffix or "").lstrip("/")
            paths.add(path.relative_to(root).as_posix())
        for path in sorted(paths):
            with self.subTest(path=path):
                run = subprocess.run(
                    ["git", "--no-replace-objects", "-C", str(root),
                     "cat-file", "-e", f"HEAD:{path}"],
                    capture_output=True, text=True,
                )
                self.assertEqual(run.returncode, 0, f"input absent from Git HEAD: {path}")


@unittest.skipUnless(Path("/usr/bin/time").is_file(), "runner requires GNU time")
class RunnerOutputTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="litchi-xlsx-runner-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.checkout = self.directory / "checkout"
        self.here = self.checkout / "docs/report/spec-gap-validation-evidence/profile"
        self.here.mkdir(parents=True)
        self.script = self.here / "run_profile.sh"
        self.script.write_bytes(Path(__file__).with_name("run_profile.sh").read_bytes())
        self.target = self.directory / "target"
        self.output = self.directory / "new-receipts"
        self.called = self.directory / "cargo-called"
        commands = self.directory / "bin"
        commands.mkdir()
        cargo = commands / "cargo"
        cargo.write_text('#!/bin/sh\nprintf called > "$CARGO_STUB_CALLED"\nexit 81\n')
        cargo.chmod(0o755)
        self.env = os.environ.copy()
        for key in (
            "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP",
            "ALLOW_EXISTING_TARGET", "XLSX_SVG_PROFILE_RESULTS",
        ):
            self.env.pop(key, None)
        self.env.update({
            "PATH": str(commands) + os.pathsep + self.env.get("PATH", ""),
            "PROFILE_FROZEN": "1", "XLSX_SVG_PROFILE_API_WIRED": "1",
            "PROCESSES": "3", "WARMUP": "2", "SAMPLES": "20",
            "CARGO_TARGET_DIR": str(self.target),
            "CARGO_STUB_CALLED": str(self.called),
            "XLSX_SVG_PROFILE_RESULTS": str(self.output),
        })
        self.retained = {
            self.here / "results/capture_native_fixture-p1.json": b"prior lane",
            self.here / "results/capture_native_fixture-p1.stderr.log": b"",
            self.here / "results/exploratory-review.json": b"prior review",
            self.here / "report.md": b"prior report",
            self.here / "verification.json": b"prior verification",
        }
        for path, value in self.retained.items():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(value)

    def run_script(self):
        return subprocess.run(
            ["bash", str(self.script)], env=self.env, capture_output=True, text=True
        )

    def assert_retained(self):
        for path, value in self.retained.items():
            self.assertEqual(path.read_bytes(), value)

    def assert_refusal(self):
        run = self.run_script()
        self.assertEqual(run.returncode, 2, run.stderr)
        self.assertFalse(self.called.exists(), "refusal must precede Cargo")
        self.assertFalse(self.target.exists(), "refusal must not create build output")
        self.assert_retained()

    def test_missing_explicit_output_refuses_without_changes(self):
        del self.env["XLSX_SVG_PROFILE_RESULTS"]
        self.assert_refusal()
        self.assertFalse(self.output.exists())

    def test_existing_output_even_if_empty_is_never_reused(self):
        self.output.mkdir()
        self.assert_refusal()
        self.assertEqual(list(self.output.iterdir()), [])
        sentinel = self.output / "capture_native_fixture-p1.json"
        sentinel.write_bytes(b"accepted evidence")
        self.assert_refusal()
        self.assertEqual(sentinel.read_bytes(), b"accepted evidence")

    def test_checkout_output_is_refused(self):
        self.env["XLSX_SVG_PROFILE_RESULTS"] = str(self.here / "new-output")
        self.assert_refusal()
        self.assertFalse((self.here / "new-output").exists())

    def test_output_cannot_overlap_disposable_target(self):
        for output, target in (
            (self.output, self.output),
            (self.output, self.output / "target"),
            (self.output / "receipts", self.output),
        ):
            with self.subTest(output=output, target=target):
                self.env["XLSX_SVG_PROFILE_RESULTS"] = str(output)
                self.env["CARGO_TARGET_DIR"] = str(target)
                self.assert_refusal()
                self.assertFalse(self.output.exists())

    def test_symlink_to_checkout_is_refused(self):
        link = self.directory / "checkout-link"
        link.symlink_to(self.checkout, target_is_directory=True)
        self.env["XLSX_SVG_PROFILE_RESULTS"] = str(link / "new-receipts")
        self.assert_refusal()
        self.assertFalse((self.checkout / "new-receipts").exists())

    def test_symlinked_script_still_recognizes_its_real_checkout(self):
        link = self.directory / "invocation-link"
        link.symlink_to(self.checkout, target_is_directory=True)
        self.script = link / self.script.relative_to(self.checkout)
        self.env["XLSX_SVG_PROFILE_RESULTS"] = str(self.checkout / "new-receipts")
        self.assert_refusal()
        self.assertFalse((self.checkout / "new-receipts").exists())

    def test_cargo_failure_retains_prior_receipts_and_new_diagnostics(self):
        run = self.run_script()
        self.assertEqual(run.returncode, 81, run.stderr)
        self.assertTrue(self.called.is_file())
        self.assertFalse(self.target.exists(), "owned target must be cleaned")
        self.assertTrue((self.output / "metadata-before.json").is_file())
        self.assert_retained()


if __name__ == "__main__":
    unittest.main()
