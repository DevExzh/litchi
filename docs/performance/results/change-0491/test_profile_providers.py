import tempfile
from pathlib import Path
import unittest

import profile_providers as profile


class ProfileParsingTests(unittest.TestCase):
    def test_perf_keeps_unsupported_and_rejects_duplicates_or_missing_events(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "perf.txt"
            lines = [f"10,,{event},100,100.00,,\n" for event in profile.PERF_EVENTS]
            lines[0] = "<not supported>,,cycles,0,0,,\n"
            path.write_text("".join(lines))
            self.assertEqual(profile._parse_perf(path)["events"]["cycles"]["status"], "unsupported")
            for text in ["".join(lines[:-1]), "".join(lines + lines[:1]), "# empty capture\n"]:
                path.write_text(text)
                with self.assertRaises(profile.ProfileError):
                    profile._parse_perf(path)

    def test_perf_rejects_negative_counter_values(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "perf.txt"
            path.write_text(
                "".join(f"10,,{event},100,100.00,,\n" for event in profile.PERF_EVENTS)
                .replace("10,,cycles", "-1,,cycles", 1)
            )
            with self.assertRaises(profile.ProfileError):
                profile._parse_perf(path)

    def test_strace_requires_summary_and_valid_counts(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "strace.txt"
            valid = "% time seconds usecs/call calls errors syscall\n100.00 0.001000 10 100 2 pread64\n100.00 0.001000 10 100 2 total\n"
            path.write_text(valid)
            self.assertEqual(profile._parse_strace(path)["syscalls"]["pread64"]["calls"], 100)
            for text in [valid.replace("100 2 pread64", "1 2 pread64"), valid.splitlines()[0], valid.replace("total", "missing")]:
                path.write_text(text)
                with self.assertRaises(profile.ProfileError):
                    profile._parse_strace(path)

    def test_profile_argv_binds_resolved_profiler_executable(self):
        build_path, build = profile._load_build()
        del build_path
        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "report.json"
            resource = Path(directory) / "resource.txt"
            profiler = Path(directory) / "perf.txt"
            binding = profile._tool_executable("perf")
            argv = profile._benchmark_argv(
                profile._spec("file"), build["binary"], report, resource,
                build["git_revision"], profiler, "perf", binding["path"],
            )
        self.assertEqual(argv[0], binding["path"])
        self.assertTrue(Path(argv[0]).is_absolute())

    def test_private_cleanup_requires_empty_owned_roots(self):
        with tempfile.TemporaryDirectory() as directory:
            run_root = Path(directory) / "run"
            tmp_root = run_root / "tmp"
            cleanup = {
                "schema": "docx-provider-profile-private-cleanup-v1",
                "status": "pass",
                "run_root": str(run_root),
                "tmpdir": str(tmp_root),
                "removed": [str(tmp_root), str(run_root)],
                "remaining": [],
            }
            self.assertEqual(profile._validate_cleanup(cleanup, run_root, tmp_root, "unit"), cleanup)
            tampered = dict(cleanup, remaining=[str(tmp_root)])
            with self.assertRaises(profile.ProfileError):
                profile._validate_cleanup(tampered, run_root, tmp_root, "unit")


if __name__ == "__main__":
    unittest.main()
