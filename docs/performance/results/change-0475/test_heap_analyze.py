#!/usr/bin/env python3
"""Focused tests for the checked interpreted-Heaptrack attribution reader."""

from __future__ import annotations

import gzip
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest


HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("heap_analyze", HERE / "heap_analyze.py")
assert SPEC is not None and SPEC.loader is not None
heap_analyze = importlib.util.module_from_spec(SPEC)
sys.modules["heap_analyze"] = heap_analyze
SPEC.loader.exec_module(heap_analyze)


def sized(value: str) -> str:
    return f"s {len(value.encode('utf-8')):x} {value}"


def fixture_trace(*, malformed: str | None = None) -> bytes:
    """Make a tiny v3 trace with build, run, OPC, ZIP, and Deflate stacks."""

    strings = [
        "perf-binary",
        "litchi_perf_baseline::pptx_streaming_create::build_corpus",
        "litchi_pptx::StreamingPresentationWriter::write_text_box",
        "litchi_perf_baseline::pptx_streaming_create::run",
        "litchi_opc::PackURI::validate_part_name",
        "soapberry_zip::StreamingArchiveWriter::central_directory",
        "miniz::deflate_init",
        "fixture.rs",
    ]
    lines = [
        "v 10500 3",
        "X /tmp/litchi --case pptx_streaming_create --semantic-shape large",
        "I 1000 1",
    ]
    lines.extend(sized(value) for value in strings)

    # Each i record has IP, module id, then function/file/line triples.  The
    # stack trace points at one direct frame and one parent node.  The frame
    # names intentionally resemble the mangled-symbol substrings exposed by
    # the real Rust binary, while remaining easy to inspect in this fixture.
    lines.extend(
        [
            "i 1 1 2 8 1",  # build_corpus
            "i 2 1 3 8 2",  # writer
            "i 3 1 4 8 3",  # run
            "i 4 1 5 8 4",  # OPC
            "i 5 1 6 8 5",  # ZIP
            "i 6 1 7 8 6",  # Deflate
            "t 1 0",
            "t 2 1",
            "t 3 0",
            "t 4 3",
            "t 5 3",
            "t 6 3",
            "a 10 2",  # build writer: 0x10
            "a 20 4",  # run OPC: 0x20
            "a 30 5",  # run ZIP: 0x30
            "a 40 6",  # run Deflate: 0x40
            "+ 0",
            "+ 1",
            "+ 2",
            "+ 3",
            "c 1",
            "- 1",
            "c 2",
            "- 0",
            "- 2",
            "- 3",
            "R 100",
        ]
    )
    if malformed is not None:
        lines.append(malformed)
    return ("\n".join(lines) + "\n").encode("utf-8")


class HeapAnalyzeTests(unittest.TestCase):
    def write(self, directory: Path, name: str, data: bytes) -> Path:
        path = directory / name
        path.write_bytes(data)
        return path

    def test_phase_requires_exact_run_or_build_function_ancestry(self) -> None:
        self.assertEqual(
            heap_analyze.classify_stack(
                ["litchi_perf_baseline::pptx_streaming_create::write_pptx_stream"]
            )[0],
            heap_analyze.PHASE_OTHER,
        )
        self.assertEqual(
            heap_analyze.classify_stack(
                ["_RNvNtCs2moOUUYrrBf_20litchi_perf_baseline21pptx_streaming_create3run"]
            )[0],
            heap_analyze.PHASE_RUN,
        )
        self.assertEqual(
            heap_analyze.classify_stack(
                ["_RNvNtCs2moOUUYrrBf_20litchi_perf_baseline21pptx_streaming_create12build_corpus"]
            )[0],
            heap_analyze.PHASE_BUILD,
        )

    def test_deflate_label_does_not_claim_initialization(self) -> None:
        phase, category = heap_analyze.classify_stack(["some::compress_payload"])
        self.assertEqual(phase, heap_analyze.PHASE_OTHER)
        self.assertEqual(category, heap_analyze.CATEGORY_DEFLATE)
        self.assertEqual(heap_analyze.CATEGORY_DEFLATE, "Deflate")

    def test_v3_parser_aligns_calls_requested_bytes_and_lifetimes(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = self.write(Path(raw), "heap.txt", fixture_trace())
            result = heap_analyze.analyze_trace(path, "H1", full=True)
        self.assertEqual(result.stats.allocation_events, 4)
        self.assertEqual(result.stats.deallocation_events, 4)
        self.assertEqual(result.stats.allocation_descriptors, 4)
        self.assertEqual(result.all_peak_live_bytes, 0x10 + 0x20 + 0x30 + 0x40)
        self.assertEqual(result.final_live_bytes, 0)
        self.assertEqual(result.final_outstanding_allocations, 0)
        self.assertEqual(result.phase[heap_analyze.PHASE_BUILD].calls, 1)
        self.assertEqual(result.phase[heap_analyze.PHASE_BUILD].requested_bytes, 0x10)
        self.assertEqual(result.phase[heap_analyze.PHASE_RUN].calls, 3)
        self.assertEqual(result.phase[heap_analyze.PHASE_RUN].requested_bytes, 0x20 + 0x30 + 0x40)
        self.assertEqual(result.category[heap_analyze.CATEGORY_WRITER].requested_bytes, 0x10)
        self.assertEqual(result.category[heap_analyze.CATEGORY_OPC].requested_bytes, 0x20)
        self.assertEqual(result.category[heap_analyze.CATEGORY_ZIP].requested_bytes, 0x30)
        self.assertEqual(result.category[heap_analyze.CATEGORY_DEFLATE].requested_bytes, 0x40)
        self.assertEqual(result.unresolved_trace_ids, set())

    def test_fast_scope_marks_whole_process_totals_unavailable(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = self.write(Path(raw), "heap.txt", fixture_trace())
            report = heap_analyze.analyze_trace(path, "H1").as_dict()
        self.assertFalse(report["records"]["complete_event_scan"])
        self.assertEqual(
            report["scope"]["whole_process"]["status"], "unavailable"
        )
        self.assertIsNone(report["scope"]["whole_process"]["requested_bytes"])
        self.assertEqual(report["timeline"]["peak_live_bytes"], 0x20 + 0x30 + 0x40)
        self.assertEqual(
            report["phase_scope"][heap_analyze.PHASE_BUILD]["status"], "unavailable"
        )

    def test_fast_selector_handles_chunk_boundary_and_prefix_ids(self) -> None:
        lines = [
            "v 10500 3",
            "X /tmp/profile --case pptx_streaming_create",
            "I 1000 1",
            sized("binary"),
            sized("litchi_perf_baseline::pptx_streaming_create::run"),
            sized("fixture.rs"),
            "i 1 1 2 3 1",
            "t 1 0",
        ]
        lines.extend("a 1 1" for _ in range(0xAC))
        lines.extend(["+ a", "+ ab", "- a", "- ab"])
        data = ("\n".join(lines) + "\n").encode()
        # A valid comment crosses the selector's 8 MiB read boundary.
        data = data[: data.index(b"+ a\n")] + b"#" + b"x" * (8 * 1024 * 1024) + b"\n" + data[data.index(b"+ a\n") :]
        with tempfile.TemporaryDirectory() as raw:
            path = self.write(Path(raw), "boundary.txt", data)
            result = heap_analyze.analyze_trace(path, "H1")
        self.assertEqual(result.stats.allocation_events, 2)
        self.assertEqual(result.stats.deallocation_events, 2)
        self.assertEqual(result.phase[heap_analyze.PHASE_RUN].requested_bytes, 2)

    def test_fast_selector_rejects_selected_malformed_and_unbalanced_records(self) -> None:
        cases = (
            (b"+ 1\n", b"+ 1 extra\n"),
            (b"+ 1\n", b"- 1\n"),
        )
        for original, replacement in cases:
            with self.subTest(replacement=replacement), tempfile.TemporaryDirectory() as raw:
                data = fixture_trace().replace(original, replacement, 1)
                path = self.write(Path(raw), "bad-selected.txt", data)
                with self.assertRaises(heap_analyze.AnalysisError):
                    heap_analyze.analyze_trace(path, "H1")

    def test_fast_run_projection_matches_full_writer_projection(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = self.write(Path(raw), "differential.txt", fixture_trace())
            full = heap_analyze.analyze_trace(path, "H1", full=True)
            fast = heap_analyze.analyze_trace(path, "H1")
        full_run = full.phase[heap_analyze.PHASE_RUN]
        fast_run = fast.phase[heap_analyze.PHASE_RUN]
        self.assertEqual(fast_run.calls, full_run.calls)
        self.assertEqual(fast_run.requested_bytes, full_run.requested_bytes)
        self.assertEqual(fast_run.peak_live_bytes, full_run.peak_live_bytes)
        for category in (
            heap_analyze.CATEGORY_OPC,
            heap_analyze.CATEGORY_ZIP,
            heap_analyze.CATEGORY_DEFLATE,
        ):
            self.assertEqual(
                fast.category[category].requested_bytes,
                full.category[category].requested_bytes,
            )

    def test_frames_include_inlined_groups_and_stack_rows_are_aligned(self) -> None:
        data = fixture_trace().replace(
            b"i 3 1 4 8 3\n", b"i 3 1 4 8 3 3 8 30\n"
        )
        with tempfile.TemporaryDirectory() as raw:
            path = self.write(Path(raw), "heap.txt", data)
            result = heap_analyze.analyze_trace(path, "H1")
        names = result.stack_names[4]
        self.assertEqual(names[0], "litchi_opc::PackURI::validate_part_name in perf-binary")
        self.assertIn("litchi_perf_baseline::pptx_streaming_create::run in perf-binary", names)
        opc_rows = [
            row for key, row in result.stacks.items() if key.category == heap_analyze.CATEGORY_OPC
        ]
        self.assertEqual(len(opc_rows), 1)
        self.assertEqual(opc_rows[0].requested_bytes, 0x20)

    def test_gzip_input_is_supported_and_binding_is_compressed(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "decoded.stdout.gz"
            with gzip.open(path, "wb") as stream:
                stream.write(fixture_trace())
            result = heap_analyze.analyze_trace(path, "H2")
        self.assertEqual(result.file_version, 3)
        self.assertEqual(result.lane, "H2")
        self.assertGreater(result.compressed_bytes, 0)
        self.assertEqual(len(result.compressed_sha256), 64)

    def test_rejects_unknown_record_and_unaligned_deallocation(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            directory = Path(raw)
            unknown = self.write(directory, "unknown.txt", fixture_trace(malformed="q 1"))
            with self.assertRaises(heap_analyze.AnalysisError):
                heap_analyze.analyze_trace(unknown, "H1", full=True)
            unaligned = self.write(directory, "unaligned.txt", fixture_trace().replace(b"- 1\n", b"- 3\n", 1))
            with self.assertRaises(heap_analyze.AnalysisError):
                heap_analyze.analyze_trace(unaligned, "H1", full=True)

    def test_rejects_missing_trace_metadata_instead_of_inventing_phase(self) -> None:
        data = fixture_trace().replace(b"t 4 3\n", b"t 4 99\n")
        with tempfile.TemporaryDirectory() as raw:
            path = self.write(Path(raw), "missing-parent.txt", data)
            result = heap_analyze.analyze_trace(path, "H1")
        self.assertIn(4, result.unresolved_trace_ids)
        self.assertNotIn(4, {key.trace_id for key in result.stacks})
        self.assertIn("unresolved", " ".join(result.limitations))

    def test_heaptrack_print_fallback_does_not_claim_requested_bytes(self) -> None:
        printed = (
            b"MOST CALLS TO ALLOCATION FUNCTIONS\n"
            b"12 calls to allocation functions with 3,456B peak consumption from\n"
            b"    writer_symbol\n"
            b"      in /tmp/binary\n"
        )
        with tempfile.TemporaryDirectory() as raw:
            directory = Path(raw)
            trace = self.write(directory, "heap.txt", fixture_trace())
            report = self.write(directory, "print.txt", printed)
            result = heap_analyze.analyze_trace(trace, "H1", print_path=report)
        self.assertIsNotNone(result.fallback)
        assert result.fallback is not None
        self.assertEqual(result.fallback["rows"][0]["allocation_calls"], 12)
        self.assertIsNone(result.fallback["rows"][0]["requested_bytes"])


if __name__ == "__main__":
    unittest.main()
