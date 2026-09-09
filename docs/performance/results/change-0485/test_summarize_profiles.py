"""Focused parser checks for the read-only profile summary helper."""

from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest


HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
import summarize_profiles as summary  # noqa: E402


class ProfileSummaryParserTests(unittest.TestCase):
    def test_perf_keeps_unsupported_and_missing_counters_unavailable(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "perf.txt"
            path.write_text(
                "1234,,cycles,\n"
                "<not supported>,,instructions,\n"
                "<not counted>,,branches,\n"
                "5678,,branch-misses,\n",
                encoding="utf-8",
            )
            parsed = summary._parse_perf(path)
        self.assertEqual(parsed["events"]["cycles"]["value"], 1234)
        self.assertEqual(parsed["events"]["cycles"]["status"], "measured")
        self.assertEqual(parsed["events"]["instructions"]["status"], "unsupported")
        self.assertIsNone(parsed["events"]["page-faults"]["value"])
        self.assertEqual(parsed["events"]["page-faults"]["status"], "missing")

    def test_strace_accepts_implicit_zero_errors_column(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "strace.txt"
            path.write_text(
                "% time     seconds  usecs/call     calls    errors syscall\n"
                "------ ----------- ----------- --------- --------- ----------------\n"
                " 90.00    0.090000           1        10           statx\n"
                " 10.00    0.010000           2         5         1 pread64\n"
                "------ ----------- ----------- --------- --------- ----------------\n"
                "100.00    0.100000           1        15           total\n",
                encoding="utf-8",
            )
            parsed = summary._parse_strace(path)
        self.assertEqual(parsed["selected"]["statx"]["calls"], 10)
        self.assertEqual(parsed["selected"]["statx"]["errors"], 0)
        self.assertEqual(parsed["selected"]["pread64"]["calls"], 5)
        self.assertEqual(parsed["selected"]["pread64"]["errors"], 1)

    def test_zero_baseline_percentage_is_explicitly_undefined(self) -> None:
        change = summary._percent_change(0, 5)
        self.assertTrue(change["zero_baseline"])
        self.assertIsNone(change["percent_change"])
        self.assertIsNone(summary._review(change, 5.0))


if __name__ == "__main__":
    unittest.main()
