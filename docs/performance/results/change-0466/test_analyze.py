#!/usr/bin/env python3
"""Focused parser tests for the 0466 sampled CPU attribution helper."""

from __future__ import annotations

import json
import gzip
import hashlib
from pathlib import Path
import tempfile
import unittest

import analyze


def frame(address: str, symbol: str, dso: str = "/tmp/normal") -> str:
    return f"\t{address} {symbol}+0x0 ({dso})"


def block(header: str, *frames: str) -> str:
    return header + "\n" + "\n".join(frames) + "\n\n"


class ParseSamplesTests(unittest.TestCase):
    def parse(self, text: str):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "perf-script.txt"
            path.write_text(text, encoding="utf-8")
            return analyze.parse_samples(path)

    def test_pid_tid_forms_and_event_colon_are_accepted(self):
        text = (
            block(
                "normal 10/11 1.0: 1,000 cycles:u:",
                frame("1000", "leaf"),
            )
            + block(
                "normal 12 13 2.0: 2 cycles:u:",
                frame("2000", "other"),
            )
        )
        samples, stats = self.parse(text)
        self.assertEqual([sample.period for sample in samples], [1000, 2])
        self.assertEqual([sample.event for sample in samples], ["cycles:u", "cycles:u"])
        self.assertEqual(stats["cycle_headers"], 2)
        self.assertEqual(stats["accepted_cycle_period"], 1002)

    def test_wrong_event_malformed_and_missing_period_are_accounted(self):
        text = (
            block(
                "normal 10/10 1.0: 7 instructions:u:",
                frame("1000", "wrong_event"),
            )
            + "normal 10/10 2.0: nope cycles:u:\n\n"
            + "normal 10/10 3.0: cycles:u:\n\n"
            + block(
                "normal 10/10 4.0: 5 cycles:u:",
                frame("1000", "valid"),
            )
        )
        samples, stats = self.parse(text)
        self.assertEqual([sample.period for sample in samples], [5])
        self.assertEqual(stats["invalid_cycle_event_headers"], 1)
        self.assertEqual(stats["malformed_cycle_headers"], 2)
        self.assertEqual(stats["invalid_sample_blocks"], 1)
        self.assertEqual(stats["invalid_sample_period"], 7)

    def test_unknown_unparsed_empty_lost_and_eof_truncated_are_not_silent(self):
        text = (
            block(
                "normal 1/1 1.0: 3 cycles:u:",
                "\t1000 [unknown] ([unknown])",
                "\tNOT A FRAME",
            )
            + block("normal 1/1 2.0: 4 cycles:u:")
            + "lost 2 samples\n"
            + "normal 1/1 3.0: 5 cycles:u:\n"
            + frame("3000", "eof_leaf")
            + "\n"
        )
        samples, stats = self.parse(text)
        self.assertEqual([sample.period for sample in samples], [3, 4, 5])
        self.assertEqual(stats["unknown_frame_period"], 3)
        self.assertEqual(stats["unknown_frame_lines"], 1)
        self.assertEqual(stats["unparsed_frame_period"], 3)
        self.assertEqual(stats["empty_stack_period"], 4)
        self.assertEqual(stats["lost_sample_count"], 2)
        # A final newline is not a block separator.  Retain the sample but
        # expose its incomplete coverage in both the sample and denominator.
        self.assertTrue(samples[-1].truncated)
        self.assertEqual(stats["truncated_period"], 5)

    def test_unresolved_address_symbols_are_unknown(self):
        text = block(
            "normal 1/1 1.0: 9 cycles:u:",
            "\t0x7fff1234",
            frame("1000", "leaf"),
        )
        samples, stats = self.parse(text)
        self.assertEqual(samples[0].symbols[0], "[unknown]")
        self.assertEqual(stats["unknown_frame_period"], 9)


class AttributionTests(unittest.TestCase):
    def sample(self, period: int, *symbols: str) -> analyze.Sample:
        return analyze.Sample(period, analyze.EXPECTED_EVENT, tuple(symbols))

    def test_inclusive_ranking_deduplicates_repeated_ancestor(self):
        samples = [
            self.sample(10, "leaf", "ancestor", "ancestor", "root"),
            self.sample(5, "other", "ancestor", "root"),
        ]
        rows = analyze.rank_symbols(
            samples,
            lambda _sample: True,
            15,
            15,
            10,
            leaf=False,
        )
        by_symbol = {row["symbol"]: row for row in rows}
        self.assertEqual(by_symbol["ancestor"]["weighted_event_period"], 15)
        self.assertEqual(by_symbol["ancestor"]["raw_stack_blocks"], 2)
        self.assertEqual(by_symbol["root"]["weighted_event_period"], 15)

    def test_exact_markers_and_disjoint_partition(self):
        commit = "<litchi_xlsx::workbook::edit::semantic::transaction::Edit>::commit"
        write = "litchi_xlsx::writer::write_to::<&mut litchi_perf_baseline::CountingSink>"
        prepare = "litchi_perf_baseline::prepare_xlsx_updates"
        nested_open = [commit, "<litchi_xlsx::workbook::model::Workbook>::from_package_with_styles", "leaf"]
        samples = [
            self.sample(2, commit, "caller"),
            self.sample(3, write, "caller"),
            self.sample(5, prepare, "caller"),
            self.sample(7, *nested_open),
            self.sample(11, "commit_name_part", "caller"),
        ]
        scopes = analyze.scope_summaries(samples, 10)
        categories = scopes["categories"]
        # The leaf-most marker owns a stack when an open/prepare marker is an
        # ancestor of the exact Edit::commit frame.
        self.assertEqual(categories["edit_commit"]["scope"]["weighted_event_period"], 9)
        self.assertEqual(categories["workbook_write_to"]["scope"]["weighted_event_period"], 3)
        self.assertEqual(categories["prepare_open_or_oracle"]["scope"]["weighted_event_period"], 5)
        self.assertEqual(categories["overlap"]["scope"]["weighted_event_period"], 0)
        self.assertEqual(categories["unclassified"]["scope"]["weighted_event_period"], 11)
        self.assertTrue(scopes["partition"]["periods_partition_exact"])
        self.assertTrue(scopes["partition"]["blocks_partition_exact"])
        self.assertTrue(scopes["mixed_marker_stacks"]["nearest_marker_resolution"])
        self.assertFalse(analyze.marker_names(("litchi_xlsx::workbook::edit::semantic::transaction::Edit::commit_name_part",)))

        writer = "<litchi_opc::pkgwriter::PackageWriter>::write_to_stream::<&mut litchi_perf_baseline::CountingSink>"
        self.assertEqual(analyze.classify_scope((writer,)), "writer_write_to_stream")

    def test_equal_weight_ancestors_have_canonical_order(self):
        rows = analyze.rank_symbols([self.sample(5, 'zeta', 'alpha', 'middle')],
                                    lambda _sample: True, 5, 5, 10, leaf=False)
        self.assertEqual([row['symbol'] for row in rows], ['alpha', 'middle', 'zeta'])

    def test_exact_contexts_deduplicate_and_reject_other_writer(self):
        commit = analyze.MARKERS["edit_commit"][0]
        attribute = "litchi_ooxml_common::xml::unqualified_attribute_value"
        samples = [self.sample(7, attribute, attribute, commit, commit),
                   self.sample(3, attribute, commit + "::{closure#0}"),
                   self.sample(5, "<litchi_opc::pkgwriter::PackageWriter>::write_to_stream::<Vec<u8>>")]
        rows = analyze.exact_contexts(samples)["rows"]
        self.assertEqual(rows["commit"]["weighted_event_period"], 7)
        self.assertEqual(rows["attribute_lookup_under_commit"]["weighted_event_period"], 7)
        self.assertEqual(rows["counting_sink_writer"]["weighted_event_period"], 0)
        self.assertEqual(analyze.classify_scope(samples[-1].symbols), "unclassified")

    def test_same_frame_marker_ambiguity_remains_overlap(self):
        markers = {"left": ("same",), "right": ("same",)}
        self.assertEqual(analyze.classify_scope(("same",), markers), "overlap")

    def test_gzip_script_hashes_decompressed_bytes_and_uses_relative_identity(self):
        text = block("normal 1/1 1.0: 3 cycles:u:", frame("1000", "leaf"))
        with tempfile.TemporaryDirectory() as directory:
            bundle = Path(directory)
            plain = bundle / "samples" / "perf-script.txt"
            compressed = bundle / "samples" / "perf-script.txt.gz"
            plain.parent.mkdir()
            plain.write_text(text, encoding="utf-8")
            with gzip.open(compressed, "wt", encoding="utf-8") as stream:
                stream.write(text)
            samples, _stats = analyze.parse_samples(compressed)
            self.assertEqual([sample.period for sample in samples], [3])
            binding = analyze.file_binding(compressed, bundle)
            self.assertEqual(binding["path"], "samples/perf-script.txt.gz")
            self.assertEqual(binding["compression"], "gzip")
            self.assertEqual(binding["bytes"], len(text.encode("utf-8")))
            self.assertEqual(binding["sha256"], hashlib.sha256(text.encode("utf-8")).hexdigest())
            self.assertFalse(binding["outside_bundle"])

    def test_analyze_binds_reports_without_claiming_phase_time(self):
        text = block(
            "normal 1/1 1.0: 3 cycles:u:",
            frame("1000", "leaf"),
        )
        report = {
            "schema_version": 1,
            "configuration": {"samples_per_case": 30},
            "results": [{"case": "xlsx_one_percent_commit_save", "elapsed_ns": {"samples": [10]}}],
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "perf-script.txt"
            report_path = root / "report.json"
            script.write_text(text, encoding="utf-8")
            report_path.write_text(json.dumps(report), encoding="utf-8")
            summary = analyze.analyze_script(script, [report_path], top=3)
            self.assertEqual(summary["whole_process"]["weighted_event_period"], 3)
            self.assertEqual(summary["inputs"]["normal_reports"][0]["status"], "parsed")
            self.assertEqual(
                summary["inputs"]["normal_reports"][0]["results"]["elapsed_statistics"][0]["sample_count"],
                1,
            )
            self.assertFalse(summary["timing_semantics"]["phase_latency"])
        self.assertFalse(summary["timing_semantics"]["operation_counter"])


if __name__ == "__main__":
    unittest.main()
