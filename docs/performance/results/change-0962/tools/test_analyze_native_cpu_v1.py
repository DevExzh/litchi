#!/usr/bin/env python3
"""Focused synthetic controls for analyze_native_cpu_v1.py."""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import analyze_native_cpu_v1 as analyzer  # noqa: E402


SOURCE = "/retained/tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs"


def frame(ip: str, symbol: str, dso: str = "", *, inline: bool = False, source: str | None = None) -> str:
    marker = " (inlined)" if inline else ""
    suffix = f" ({dso})" if dso else ""
    result = f"\t{ip} {symbol}{marker}{suffix}\n"
    if source is not None:
        result += f"  {source}\n"
    return result


def sample(period: int, frames: list[str], ordinal: int) -> str:
    body = "".join(frames)
    return f"bench 99/99 1.{ordinal:06d}: {period:>10} cycles:u: \n{body}\n"


def qualified(period: int, leaf: str, ordinal: int) -> str:
    return sample(
        period,
        [
            frame("100", leaf, "/retained/benchmark"),
            frame("101", "stage_timed_commit+0x2", inline=True, source=f"{SOURCE}:398 (inlined)"),
            frame("102", "timed_iteration+0x3", "/retained/benchmark", source=f"{SOURCE}:428"),
        ],
        ordinal,
    )


def paired_record(period: int, ordinal: int, *, physical_timed: bool = True, unknown: bool = False) -> tuple[str, str]:
    """Return inline and no-inline records sharing one exact event key."""
    inline_frames = [
        frame("100", "leaf+0x1", "/retained/benchmark"),
        frame("102", "stage_timed_commit+0x2", inline=True, source=f"{SOURCE}:398 (inlined)"),
        frame("102", "timed_iteration+0x3", inline=True, source=f"{SOURCE}:428 (inlined)"),
    ]
    physical_frames = [
        frame("100", "leaf+0x1", "/retained/benchmark"),
    ]
    if physical_timed:
        physical_frames.append(
            frame(
                "102",
                analyzer.DEFAULT_PHYSICAL_TIMED_PREFIX + "timed_iteration+0x3",
                "/retained/benchmark",
                source="mod.rs:547",
            )
        )
    if unknown:
        inline_frames.append(frame("103", "[unknown]", "[unknown]"))
        physical_frames.append(frame("103", "[unknown]", "[unknown]"))
    return sample(period, inline_frames, ordinal), sample(period, physical_frames, ordinal)


class AnalyzerControls(unittest.TestCase):
    def run_fixture(self, text: str) -> dict:
        with tempfile.TemporaryDirectory(prefix="litchi-0962-analyzer-") as directory:
            root = Path(directory)
            script = root / "fixture.txt"
            output = root / "result.json"
            script.write_text(text, encoding="utf-8")
            arguments = argparse.Namespace(
                perf_data=None,
                script=script,
                physical_script=None,
                output=output,
                event="cycles:u",
                source_file_suffix="xlsx_managed_compression_benchmark.rs",
                stage_lines="382-402",
                timed_lines="404-449",
                physical_timed_prefix=analyzer.DEFAULT_PHYSICAL_TIMED_PREFIX,
                binary=None,
            )
            analyzer.run(arguments)
            return json.loads(output.read_text(encoding="utf-8"))

    def test_conservation_and_disjoint_qualification(self) -> None:
        text = "".join(
            [
                qualified(10, "zlib_rs::deflate_medium+0x1", 1),
                qualified(20, "litchi_xlsx::cell_values::source::commit+0x1", 2),
                sample(
                    30,
                    [
                        frame("200", "leaf+0x1", "/retained/benchmark"),
                        frame("201", "stage_timed_commit+0x2", inline=True, source=f"{SOURCE}:398 (inlined)"),
                        frame("202", "timed_iteration+0x3", inline=True, source=f"{SOURCE}:428 (inlined)"),
                    ],
                    3,
                ),
                sample(
                    40,
                    [
                        frame("300", "leaf+0x1", "/retained/benchmark"),
                        frame("301", "stage_timed_commit+0x2", inline=True, source=f"{SOURCE}:398 (inlined)"),
                        frame("302", "timed_iteration+0x3", "/retained/benchmark", source=f"{SOURCE}:428"),
                        frame("303", "[unknown]", "[unknown]"),
                    ],
                    4,
                ),
                sample(50, [frame("400", "unrelated+0x1", "/retained/benchmark")], 5),
                sample(
                    60,
                    [
                        frame("500", "leaf+0x1", "/retained/benchmark"),
                        frame("501", "stage_timed_commit+0x2", inline=True, source=f"{SOURCE}:398 (inlined)"),
                        frame("502", "stage_timed_commit+0x3", inline=True, source=f"{SOURCE}:399 (inlined)"),
                        frame("503", "timed_iteration+0x3", "/retained/benchmark", source=f"{SOURCE}:428"),
                    ],
                    6,
                ),
            ]
        )
        result = self.run_fixture(text)
        self.assertEqual(result["qualification"], "qualified_below_admission_gate")
        self.assertEqual(result["whole_process"]["samples"], 6)
        self.assertEqual(result["whole_process"]["period"], 210)
        self.assertEqual(result["denominators"]["qualified_ancestor"]["samples"], 2)
        self.assertEqual(result["denominators"]["qualified_ancestor"]["period"], 30)
        self.assertEqual(result["denominators"]["inline_only"]["period"], 30)
        self.assertEqual(result["classification"]["unknown_frames"]["period"], 40)
        self.assertEqual(result["classification"]["ambiguous_phase"]["period"], 60)
        self.assertTrue(result["conservation"]["sample_count_pass"])
        self.assertTrue(result["conservation"]["period_sum_pass"])
        self.assertFalse(result["qualification_gate"]["admitted"])
        owners = {item["key"]: item["period"] for item in result["qualified_attribution"]["exclusive_owner_category"]}
        self.assertEqual(owners, {"compression": 10, "litchi_xlsx": 20})
        tags = {item["key"]: item["period"] for item in result["qualified_attribution"]["inclusive_tags"]}
        self.assertEqual(tags["timed_iteration"], 30)
        self.assertEqual(tags["stage_timed_commit"], 30)
        self.assertGreater(sum(tags.values()), result["denominators"]["qualified_ancestor"]["period"])

    def test_exact_matching_order_and_source_gate(self) -> None:
        text = "".join(
            [
                sample(
                    10,
                    [
                        frame("100", "not_stage_timed_commit+0x1", "/retained/benchmark"),
                        frame("101", "stage_timed_commit_extra+0x2", inline=True, source=f"{SOURCE}:398 (inlined)"),
                        frame("102", "timed_iteration_extra+0x3", "/retained/benchmark", source=f"{SOURCE}:428"),
                    ],
                    1,
                ),
                sample(
                    20,
                    [
                        frame("200", "leaf+0x1", "/retained/benchmark"),
                        frame("201", "timed_iteration+0x2", "/retained/benchmark", source=f"{SOURCE}:428"),
                        frame("202", "stage_timed_commit+0x3", inline=True, source=f"{SOURCE}:398 (inlined)"),
                    ],
                    2,
                ),
                sample(
                    30,
                    [
                        frame("300", "leaf+0x1", "/retained/benchmark"),
                        frame("301", "stage_timed_commit+0x2", source=f"{SOURCE}:398"),
                        frame("302", "timed_iteration+0x3", "/retained/benchmark", source=f"{SOURCE}:428"),
                    ],
                    3,
                ),
                sample(
                    40,
                    [
                        frame("400", "leaf+0x1", "/retained/benchmark"),
                        frame("401", "stage_timed_commit+0x2", inline=True, source=f"{SOURCE}:999 (inlined)"),
                        frame("402", "timed_iteration+0x3", "/retained/benchmark", source=f"{SOURCE}:428"),
                    ],
                    4,
                ),
            ]
        )
        result = self.run_fixture(text)
        self.assertEqual(result["qualification"], "no_qualified_ancestry")
        self.assertEqual(result["denominators"]["qualified_ancestor"]["samples"], 0)
        self.assertEqual(result["classification"]["no_phase_marker"]["period"], 10)
        self.assertEqual(result["classification"]["unqualified_phase"]["period"], 50)
        self.assertEqual(result["classification"]["source_invalid"]["period"], 40)
        self.assertTrue(result["conservation"]["period_sum_pass"])

    def test_mixed_inline_and_physical_markers_fail_closed(self) -> None:
        result = self.run_fixture(
            sample(
                10,
                [
                    frame("100", "leaf+0x1", "/retained/benchmark"),
                    frame("101", "stage_timed_commit+0x2", inline=True, source=f"{SOURCE}:398 (inlined)"),
                    frame("102", "stage_timed_commit+0x3", "/retained/benchmark", source=f"{SOURCE}:399"),
                    frame("103", "timed_iteration+0x4", "/retained/benchmark", source=f"{SOURCE}:428"),
                ],
                1,
            )
        )
        self.assertEqual(result["classification"]["ambiguous_phase"]["period"], 10)
        self.assertEqual(result["reasons"]["mixed_stage_marker_forms"]["period"], 10)

        result = self.run_fixture(
            sample(
                11,
                [
                    frame("100", "leaf+0x1", "/retained/benchmark"),
                    frame("102", "timed_iteration+0x4", inline=True, source=f"{SOURCE}:428 (inlined)"),
                ],
                2,
            )
        )
        self.assertEqual(result["classification"]["unqualified_phase"]["period"], 11)
        self.assertEqual(result["reasons"]["stage_marker_missing"]["period"], 11)

    def test_owner_matching_does_not_treat_compress_as_sha(self) -> None:
        ordinary = analyzer.Frame(
            index=0,
            ip=1,
            symbol="compression::compress_block",
            dso="/retained/benchmark",
            inline=False,
            unknown=False,
        )
        sha = analyzer.Frame(
            index=0,
            ip=1,
            symbol="sha2::compress256",
            dso="/retained/benchmark",
            inline=False,
            unknown=False,
        )
        self.assertNotEqual(analyzer.owner_category(ordinary), "sha256")
        zlib_frame = analyzer.Frame(
            index=0,
            ip=1,
            symbol="zlib_rs::compress_block",
            dso="/retained/benchmark",
            inline=False,
            unknown=False,
        )
        self.assertEqual(analyzer.owner_category(zlib_frame), "compression")
        self.assertEqual(analyzer.owner_category(sha), "sha256")

    def test_wrong_event_fails_closed(self) -> None:
        text = sample(1, [frame("100", "leaf", "/retained/benchmark")], 1).replace("cycles:u:", "instructions:")
        with self.assertRaises(analyzer.AnalysisError):
            analyzer.parse_samples(text.splitlines(keepends=True), "cycles:u")

    def run_paired_fixture(self, inline_text: str, physical_text: str) -> dict:
        with tempfile.TemporaryDirectory(prefix="litchi-0962-paired-") as directory:
            root = Path(directory)
            inline_path = root / "inline.txt"
            physical_path = root / "physical.txt"
            output = root / "result.json"
            inline_path.write_text(inline_text, encoding="utf-8")
            physical_path.write_text(physical_text, encoding="utf-8")
            arguments = argparse.Namespace(
                perf_data=None,
                script=inline_path,
                physical_script=physical_path,
                output=output,
                event="cycles:u",
                source_file_suffix="xlsx_managed_compression_benchmark.rs",
                stage_lines="382-402",
                timed_lines="404-449",
                physical_timed_prefix=analyzer.DEFAULT_PHYSICAL_TIMED_PREFIX,
                binary=None,
            )
            analyzer.run(arguments)
            return json.loads(output.read_text(encoding="utf-8"))

    def test_paired_physical_ip_gate_and_conservation(self) -> None:
        inline_rows = []
        physical_rows = []
        for row in (
            paired_record(10, 1),
            paired_record(20, 2, physical_timed=False),
            paired_record(30, 3, unknown=True),
        ):
            inline_rows.append(row[0])
            physical_rows.append(row[1])
        result = self.run_paired_fixture("".join(inline_rows), "".join(physical_rows))
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["qualification"], "qualified_below_admission_gate")
        self.assertEqual(result["denominators"]["qualified_ancestor"]["period"], 10)
        self.assertEqual(result["denominators"]["inline_only"]["period"], 20)
        self.assertEqual(result["classification"]["unknown_frames"]["period"], 30)
        self.assertEqual(result["pairing"]["paired_samples"], 3)
        self.assertTrue(result["conservation"]["paired_stream_population_pass"])
        self.assertTrue(result["conservation"]["inline_period_sum_pass"])
        self.assertFalse(result["qualification_gate"]["admitted"])

    def test_physical_source_and_inline_evidence_fail_closed(self) -> None:
        inline, physical = paired_record(10, 1)
        missing_source = physical.replace("  mod.rs:547\n", "")
        result = self.run_paired_fixture(inline, missing_source)
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["classification"]["source_invalid"]["period"], 10)
        self.assertEqual(result["reasons"]["physical_timed_source_check_failed"]["period"], 10)
        self.assertEqual(result["denominators"]["qualified_ancestor"]["samples"], 0)

        inline, physical = paired_record(10, 2)
        physical_inline = physical.replace(
            "timed_iteration+0x3 (/retained/benchmark)",
            "timed_iteration+0x3 (inlined) (/retained/benchmark)",
        )
        result = self.run_paired_fixture(inline, physical_inline)
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["classification"]["ambiguous_phase"]["period"], 10)
        self.assertEqual(result["reasons"]["physical_stream_contains_inline_evidence"]["period"], 10)
        self.assertEqual(result["denominators"]["qualified_ancestor"]["samples"], 0)

    def test_duplicate_header_payloads_fail_closed_and_are_reported(self) -> None:
        inline_a, physical_a = paired_record(10, 1)
        inline_b, physical_b = paired_record(10, 1)
        # The event headers are identical while the callchain payloads differ.
        # No join can prove which physical payload belongs to which inline row.
        inline_b = inline_b.replace("leaf+0x1", "other_leaf+0x1")
        physical_a = physical_a.replace("leaf+0x1", "other_physical+0x1")
        result = self.run_paired_fixture(inline_a + inline_b, physical_a + physical_b)
        self.assertEqual(result["status"], "failed_closed")
        self.assertTrue(result["pairing"]["header_multiset_equal"])
        self.assertEqual(result["pairing"]["duplicate_inline_keys"], 1)
        self.assertEqual(result["pairing"]["duplicate_physical_keys"], 1)
        self.assertEqual(result["pairing"]["duplicate_inline_records"], 1)
        self.assertEqual(result["pairing"]["duplicate_physical_records"], 1)
        self.assertEqual(result["classification"]["ambiguous_pairing"]["period"], 20)
        self.assertEqual(result["denominators"]["qualified_ancestor"]["samples"], 0)

    def test_header_count_and_period_mismatches_are_not_paired(self) -> None:
        inline, physical = paired_record(10, 1)
        _, extra_physical = paired_record(20, 2)
        result = self.run_paired_fixture(inline, physical + extra_physical)
        self.assertEqual(result["status"], "failed_closed")
        self.assertFalse(result["pairing"]["header_count_equal"])
        self.assertFalse(result["pairing"]["header_multiset_equal"])
        self.assertEqual(result["pairing"]["unmatched_physical_samples"], 1)

        inline, physical = paired_record(10, 3)
        # Same timestamp and task identity, but a changed period is a different
        # event key and must remain unmatched rather than being zip-joined.
        _, changed_period = paired_record(11, 3)
        result = self.run_paired_fixture(inline, changed_period)
        self.assertEqual(result["status"], "failed_closed")
        self.assertFalse(result["pairing"]["header_period_equal"])
        self.assertFalse(result["pairing"]["header_multiset_equal"])
        self.assertEqual(result["pairing"]["paired_samples"], 0)

    def test_pairing_does_not_join_by_position_or_ignore_unmatched(self) -> None:
        inline, _ = paired_record(10, 1)
        _, physical = paired_record(11, 2)
        result = self.run_paired_fixture(inline, physical)
        self.assertEqual(result["status"], "failed_closed")
        self.assertEqual(result["pairing"]["paired_samples"], 0)
        self.assertEqual(result["classification"]["unpaired_inline"]["period"], 10)
        self.assertFalse(result["conservation"]["paired_stream_population_pass"])

    def test_malformed_orphan_fails_closed(self) -> None:
        with self.assertRaises(analyzer.AnalysisError):
            analyzer.parse_samples(["not a perf record\n"], "cycles:u")


if __name__ == "__main__":
    unittest.main()
