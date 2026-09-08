#!/usr/bin/env python3
"""Focused tests for the standalone 0475 CPU profile evidence helpers."""

from __future__ import annotations

import gzip
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

import cpu_analyze
import cpu_parser


def frame(address: str, symbol: str, dso: str = "/tmp/profile") -> str:
    return f"\t{address} {symbol}+0x0 ({dso})"


def block(header: str, *frames: str) -> str:
    return header + "\n" + "\n".join(frames) + "\n\n"


class ParseSamplesTests(unittest.TestCase):
    def parse(self, text: str):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "perf-script.txt"
            path.write_text(text, encoding="utf-8")
            return cpu_parser.parse_samples(path)

    def test_pid_tid_forms_and_event_colon_are_accepted(self):
        text = (
            block(
                "profile 10/11 1.0: 1,000 cycles:u:",
                frame("1000", "leaf"),
            )
            + block(
                "profile 12 13 2.0: 2 cycles:u:",
                frame("2000", "other"),
            )
        )
        samples, stats = self.parse(text)
        self.assertEqual([sample.period for sample in samples], [1000, 2])
        self.assertEqual([sample.event for sample in samples], ["cycles:u", "cycles:u"])
        self.assertEqual([sample.tid for sample in samples], [11, 13])
        self.assertEqual(stats["cycle_headers"], 2)
        self.assertEqual(stats["accepted_cycle_period"], 1002)

    def test_wrong_event_malformed_and_nonpositive_periods_are_accounted(self):
        text = (
            block(
                "profile 10/10 1.0: 7 instructions:u:",
                frame("1000", "wrong_event"),
            )
            + "profile 10/10 2.0: nope cycles:u:\n\n"
            + "profile 10/10 3.0: cycles:u:\n\n"
            + "profile 10/10 4.0: 0 cycles:u:\n\n"
            + block(
                "profile 10/10 5.0: 5 cycles:u:",
                frame("1000", "valid"),
            )
        )
        samples, stats = self.parse(text)
        self.assertEqual([sample.period for sample in samples], [5])
        self.assertEqual(stats["invalid_cycle_event_headers"], 1)
        self.assertEqual(stats["malformed_period_headers"], 1)
        self.assertEqual(stats["invalid_cycle_period_headers"], 1)
        self.assertEqual(stats["invalid_sample_blocks"], 3)
        self.assertEqual(stats["invalid_sample_period"], 7)

    def test_unknown_unparsed_empty_lost_explicit_and_eof_truncation_are_visible(self):
        text = (
            block(
                "profile 1/1 1.0: 3 cycles:u:",
                "\t1000 [unknown] ([unknown])",
                "\tNOT A FRAME",
            )
            + block("profile 1/1 2.0: 4 cycles:u:")
            + block(
                "profile 1/1 3.0: 5 cycles:u:",
                frame("3000", "explicit_leaf"),
                "\t...",
            )
            + "lost 2 samples\n"
            + "profile 1/1 4.0: 6 cycles:u:\n"
            + frame("4000", "eof_leaf")
            + "\n"
        )
        samples, stats = self.parse(text)
        self.assertEqual([sample.period for sample in samples], [3, 4, 5, 6])
        self.assertEqual(stats["unknown_frame_period"], 3)
        self.assertEqual(stats["unparsed_frame_period"], 3)
        self.assertEqual(stats["empty_stack_period"], 4)
        self.assertEqual(stats["lost_sample_count"], 2)
        self.assertEqual(stats["truncation_markers"], 1)
        self.assertEqual(stats["truncated_period"], 11)
        self.assertTrue(samples[2].truncated)
        self.assertTrue(samples[-1].truncated)

    def test_unresolved_address_symbols_are_unknown(self):
        text = block(
            "profile 1/1 1.0: 9 cycles:u:",
            "\t0x7fff1234",
            frame("1000", "leaf"),
        )
        samples, stats = self.parse(text)
        self.assertEqual(samples[0].symbols[0], "[unknown]")
        self.assertEqual(stats["unknown_frame_period"], 9)


class AttributionTests(unittest.TestCase):
    def sample(self, period: int, *symbols: str) -> cpu_parser.Sample:
        return cpu_parser.Sample(period, cpu_parser.EXPECTED_EVENT, tuple(symbols))

    def test_inclusive_ranking_deduplicates_repeated_ancestor(self):
        samples = [
            self.sample(10, "leaf", "ancestor", "ancestor", "root"),
            self.sample(5, "other", "ancestor", "root"),
        ]
        rows = cpu_parser.rank_symbols(samples, lambda _sample: True, 15, 15, 10, leaf=False)
        by_symbol = {row["symbol"]: row for row in rows}
        self.assertEqual(by_symbol["ancestor"]["weighted_event_period"], 15)
        self.assertEqual(by_symbol["ancestor"]["raw_stack_blocks"], 2)
        self.assertEqual(by_symbol["root"]["weighted_event_period"], 15)

    def test_writer_and_preflight_contexts_partition_without_losing_samples(self):
        run = cpu_analyze.DEFAULT_CONTEXT_MARKERS["writer_under_run"][0]
        preflight = cpu_analyze.DEFAULT_CONTEXT_MARKERS["materialized_preflight"][0]
        samples = [
            self.sample(2, "leaf", run),
            self.sample(3, "leaf", preflight),
            self.sample(5, "leaf", "other"),
        ]
        result = cpu_parser.context_summaries(
            samples, 10, cpu_analyze.DEFAULT_CONTEXT_MARKERS
        )
        categories = result["categories"]
        self.assertEqual(categories["writer_under_run"]["scope"]["weighted_event_period"], 2)
        self.assertEqual(categories["materialized_preflight"]["scope"]["weighted_event_period"], 3)
        self.assertEqual(categories["unclassified"]["scope"]["weighted_event_period"], 5)
        self.assertTrue(result["partition"]["periods_partition_exact"])
        self.assertTrue(result["partition"]["blocks_partition_exact"])

    def test_same_frame_context_ambiguity_remains_overlap(self):
        markers = {"left": ("same",), "right": ("same",)}
        self.assertEqual(cpu_parser.classify_scope(("same",), markers), "overlap")

    def test_writer_marker_takes_precedence_over_preflight_marker(self):
        run = cpu_analyze.DEFAULT_CONTEXT_MARKERS["writer_under_run"][0]
        preflight = cpu_analyze.DEFAULT_CONTEXT_MARKERS["materialized_preflight"][0]
        self.assertEqual(
            cpu_parser.classify_context_scope((run, preflight), cpu_analyze.DEFAULT_CONTEXT_MARKERS),
            "writer_under_run",
        )

    def test_owner_families_are_independent_overlapping_lenses(self):
        samples = [
            self.sample(
                7,
                "flate2::zio::Writer<...>::write_with_status",
                "soapberry_zip::office::StreamingArchiveWriter::<...>::start_entry",
                "litchi_perf_baseline::HashingDiscardSink::write",
            ),
            self.sample(
                2,
                "<flate2::deflate::write::DeflateEncoder<target>>::new",
            ),
            self.sample(1, "zlib_rs::stable::Deflate::new"),
            self.sample(1, "<flate2::mem::Compress>::new"),
            self.sample(3, "ordinary_frame"),
        ]
        result = cpu_parser.owner_family_summaries(
            samples, 10, cpu_analyze.DEFAULT_OWNER_FAMILIES
        )
        families = result["families"]
        self.assertEqual(families["deflate_compression"]["scope"]["weighted_event_period"], 7)
        self.assertEqual(families["deflate_initialization"]["scope"]["weighted_event_period"], 4)
        self.assertEqual(families["zip_entry_directory"]["scope"]["weighted_event_period"], 7)
        self.assertEqual(families["hashing_discard_sink"]["scope"]["weighted_event_period"], 7)
        self.assertTrue(result["overlap_allowed"])
        self.assertTrue(result["period_sum_is_not_a_partition"])

    def test_gzip_script_binding_uses_decompressed_identity(self):
        text = block("profile 1/1 1.0: 3 cycles:u:", frame("1000", "leaf"))
        with tempfile.TemporaryDirectory() as directory:
            bundle = Path(directory)
            path = bundle / "samples" / "perf-script.txt.gz"
            path.parent.mkdir()
            with gzip.open(path, "wt", encoding="utf-8") as stream:
                stream.write(text)
            samples, _stats = cpu_parser.parse_samples(path)
            self.assertEqual([sample.period for sample in samples], [3])
            binding = cpu_parser.file_binding(path, bundle)
            self.assertEqual(binding["path"], "samples/perf-script.txt.gz")
            self.assertEqual(binding["compression"], "gzip")
            self.assertEqual(binding["bytes"], len(text.encode("utf-8")))
            self.assertEqual(
                binding["sha256"], hashlib.sha256(text.encode("utf-8")).hexdigest()
            )

    def test_analyze_binds_report_without_claiming_phase_time(self):
        text = block(
            "profile 1/1 1.0: 3 cycles:u:",
            frame("1000", "leaf"),
        ) + "lost 2 samples\n"
        report = {
            "schema_version": 1,
            "configuration": {"samples_per_case": 30},
            "results": [{"case": "pptx_streaming_create", "elapsed_ns": {"samples": [10]}}],
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "perf-script.txt"
            report_path = root / "report.json"
            script.write_text(text, encoding="utf-8")
            report_path.write_text(json.dumps(report), encoding="utf-8")
            summary = cpu_analyze.analyze_script(script, [report_path], bundle_root=root)
            self.assertEqual(summary["whole_process"]["weighted_event_period"], 3)
            self.assertEqual(summary["inputs"]["normal_reports"][0]["status"], "parsed")
            self.assertFalse(summary["timing_semantics"]["phase_latency"])
            self.assertFalse(summary["timing_semantics"]["sample_period_to_elapsed_conversion"])
            self.assertEqual(
                summary["writer_context"]["scope"]["lost_sample_count"], 2
            )


if __name__ == "__main__":
    unittest.main()
