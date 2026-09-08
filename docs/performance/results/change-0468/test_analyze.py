#!/usr/bin/env python3
"""Focused tests for the 0468 analyzer wrapper."""

from __future__ import annotations

import gzip
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

import analyze


def frame(address: str, symbol: str, dso: str = "/tmp/profile") -> str:
    return f"\t{address} {symbol}+0x0 ({dso})"


def block(header: str, *frames: str) -> str:
    return header + "\n" + "\n".join(frames) + "\n\n"


class WrapperImportTests(unittest.TestCase):
    def test_shared_analyzer_is_loaded_by_explicit_path(self):
        self.assertEqual(analyze.legacy.SCHEMA, "litchi-0466-xlsx-cpu-attribution-v1")
        self.assertEqual(analyze.LEGACY_PATH.name, "analyze.py")
        self.assertNotEqual(analyze.SCHEMA, analyze.legacy.SCHEMA)


class MaterialViewTests(unittest.TestCase):
    def sample(self, period: int, *symbols: str) -> analyze.Sample:
        return analyze.Sample(period, analyze.EXPECTED_EVENT, tuple(symbols))

    def test_commit_context_and_material_families_are_inclusive_and_deduplicated(self):
        commit = analyze.COMMIT_MARKER
        samples = [
            self.sample(
                10,
                "litchi_xlsx::raw::worksheet::model::Parser::parse",
                commit,
                "litchi_xlsx::raw::worksheet::model::Parser::parse",
            ),
            self.sample(
                5,
                "litchi_xlsx::raw::worksheet::edit::codec::wire::tag",
                commit,
            ),
            self.sample(3, "quick_xml::events::attributes::IterState::next", commit),
            self.sample(7, "outside::caller"),
        ]
        summary = analyze.material_symbol_views(samples, top=10)
        self.assertEqual(summary["whole_process"]["weighted_event_period"], 25)
        self.assertEqual(summary["exact_commit_context"]["weighted_event_period"], 18)
        parser = summary["categories"]["parser"]["exact_commit_context"]
        parser_rows = {row["symbol"]: row for row in parser["inclusive_period_weighted_ranking"]}
        # The repeated parser frame is one inclusive frame per stack block.
        self.assertEqual(parser_rows["litchi_xlsx::raw::worksheet::model::Parser::parse"]["weighted_event_period"], 10)
        self.assertEqual(parser["matched_weighted_event_period"], 10)
        attribute = summary["categories"]["attribute"]["exact_commit_context"]
        self.assertEqual(attribute["matched_weighted_event_period"], 3)

    def test_family_rows_can_overlap_without_being_presented_as_additive(self):
        sample = self.sample(
            4,
            "litchi_xlsx::raw::worksheet::model::Parser::parse",
            "quick_xml::events::attributes::IterState::next",
            analyze.COMMIT_MARKER,
        )
        summary = analyze.material_symbol_views([sample])
        self.assertTrue(summary["families_overlap"])
        self.assertEqual(
            summary["categories"]["parser"]["exact_commit_context"][
                "matched_weighted_event_period"
            ],
            4,
        )
        self.assertEqual(
            summary["categories"]["attribute"]["exact_commit_context"][
                "matched_weighted_event_period"
            ],
            4,
        )

    def test_profile_context_keeps_direct_web_caller_and_compact_views(self):
        commit = analyze.COMMIT_MARKER
        samples = [
            self.sample(
                9,
                "quick_xml::reader::Reader::read_event_impl",
                "litchi_xlsx::raw::web::read",
                commit,
            ),
            self.sample(
                6,
                "quick_xml::events::Event>::into_owned",
                "litchi_xlsx::raw::compact::changed",
                commit,
            ),
            self.sample(3, "outside::caller"),
        ]
        with tempfile.TemporaryDirectory() as directory:
            script = Path(directory) / "perf-script.txt"
            script.write_text(
                block(
                    "normal 1/1 1.0: 9 cycles:u:",
                    frame("1000", samples[0].symbols[0]),
                    frame("2000", samples[0].symbols[1]),
                    frame("3000", samples[0].symbols[2]),
                )
                + block(
                    "normal 1/1 2.0: 6 cycles:u:",
                    frame("4000", samples[1].symbols[0]),
                    frame("5000", samples[1].symbols[1]),
                    frame("6000", samples[1].symbols[2]),
                )
                + block(
                    "normal 1/1 3.0: 3 cycles:u:",
                    frame("7000", samples[2].symbols[0]),
                ),
                encoding="utf-8",
            )
            detail = analyze.profile_context(script, top=5)
        web = detail["raw_web_read_context"]
        self.assertEqual(web["matched_sample_blocks"], 1)
        self.assertEqual(web["nearest_retained_caller"][0]["symbol"], commit)
        self.assertEqual(web["nearest_retained_child"][0]["symbol"], samples[0].symbols[0])
        compact = detail["compact_changed_subtree"]
        self.assertEqual(compact["raw_stack_blocks"], 1)
        self.assertEqual(compact["leaf_period_weighted_ranking"][0]["symbol"], samples[1].symbols[0])
        self.assertTrue(detail["sample_period_amdahl_proxy"]["rows_are_overlapping"])


class AnalyzeScriptTests(unittest.TestCase):
    def test_plain_and_gzip_inputs_bind_relative_hashes_and_no_phase_claim(self):
        text = block(
            "normal 1/1 1.0: 9 cycles:u:",
            frame("1000", "leaf"),
            frame("2000", analyze.COMMIT_MARKER),
        )
        report = {
            "schema_version": 1,
            "configuration": {"samples_per_case": 50},
            "results": [{"case": "xlsx_one_percent_commit_save", "elapsed_ns": {"samples": [10]}}],
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "samples-fp" / "perf-script.stdout"
            compressed = root / "samples-fp" / "perf-script.stdout.gz"
            report_path = root / "normal-r1" / "report.json"
            script.parent.mkdir()
            report_path.parent.mkdir()
            script.write_text(text, encoding="utf-8")
            with gzip.open(compressed, "wt", encoding="utf-8") as stream:
                stream.write(text)
            report_path.write_text(json.dumps(report), encoding="utf-8")
            plain = analyze.analyze_script(script, [report_path], top=5)
            zipped = analyze.analyze_script(compressed, [report_path], top=5)
        self.assertEqual(plain["whole_process"], zipped["whole_process"])
        self.assertEqual(plain["inputs"]["perf_script"]["path"], "perf-script.stdout")
        self.assertFalse(plain["timing_semantics"]["phase_latency"])
        self.assertFalse(plain["timing_semantics"]["operation_counter"])
        self.assertEqual(
            plain["inputs"]["normal_reports"][0]["file"]["path"], "report.json"
        )
        self.assertEqual(
            plain["inputs"]["perf_script"]["sha256"],
            hashlib.sha256(text.encode("utf-8")).hexdigest(),
        )
        self.assertEqual(plain["inputs"]["parser"]["path"], "analyze.py")
        self.assertEqual(plain["inputs"]["shared_analyzer"]["path"], "../change-0466/analyze.py")


if __name__ == "__main__":
    raise SystemExit(unittest.main())
