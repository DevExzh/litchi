#!/usr/bin/env python3
"""Small mutation tests for the 0479 evidence boundary.

These tests use synthetic, self-contained reports.  They exercise the
independent checks that must reject plausible but altered evidence before the
formal 24-process bundle exists.
"""

from __future__ import annotations

import copy
import hashlib
from pathlib import Path
import tempfile
import unittest
import zlib

import analyze
import verify


def _digest(seed: str) -> str:
    return hashlib.sha256(seed.encode()).hexdigest()


def _process() -> dict[str, int]:
    return {field: (100 if field == "clock_ticks_per_second" else 0) for field in analyze.PROCESS_FIELDS}


def _reads(calls: int = 1, requested: int = 100, returned: int = 100) -> dict[str, object]:
    return {
        "calls": calls,
        "requested_bytes": requested,
        "returned_bytes": returned,
        "request_histogram": {
            "bytes_0": 0,
            "bytes_1_to_512": calls,
            "bytes_513_to_4096": 0,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": 0,
            "bytes_over_65536": 0,
        },
    }


def _allocation(measured: bool, *, before: int = 100, allocated: int = 200, deallocated: int = 200, after: int = 100, peak: int = 200) -> dict[str, object] | None:
    if not measured:
        return None
    return {
        "status": "measured",
        "scope": "operation_global_system_allocator",
        "allocation_calls": 3,
        "deallocation_calls": 2,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": allocated,
        "deallocated_bytes": deallocated,
        "live_bytes_before": before,
        "live_bytes_after": after,
        "peak_live_bytes_before": before,
        "peak_live_bytes_after": peak,
        "region_peak_live_bytes": peak,
    }


def _member(path: str, decoded: int, decoded_hash: str, seed: str) -> dict[str, object]:
    return {
        "path": path,
        "compression_method": "Deflated",
        "data_descriptor": False,
        "crc32": 1,
        "decoded_bytes": decoded,
        "decoded_sha256": decoded_hash,
        "compressed_bytes": max(1, decoded // 2),
        "compressed_sha256": _digest(seed),
    }


def _corpus(count: int) -> dict[str, object]:
    source_xml = analyze._source_xml(count)
    candidate_xml = analyze._candidate_xml(count)
    source_hash = _digest(f"source-archive-{count}")
    candidate_hash = _digest(f"candidate-archive-{count}")
    main_source = _member(analyze.MAIN_PATH, len(source_xml), hashlib.sha256(source_xml).hexdigest(), f"main-source-{count}")
    main_candidate = _member(analyze.MAIN_PATH, len(candidate_xml), hashlib.sha256(candidate_xml).hexdigest(), f"main-candidate-{count}")
    opaque = _member(analyze.OPAQUE_PATH, analyze.OPAQUE_BYTES, analyze.OPAQUE_SHA256, "opaque-compressed")
    opaque["crc32"] = analyze.OPAQUE_CRC32
    main_source["crc32"] = zlib.crc32(source_xml) & 0xFFFF_FFFF
    main_candidate["crc32"] = zlib.crc32(candidate_xml) & 0xFFFF_FFFF
    content_types = _member("[Content_Types].xml", 50, _digest("content-types"), "content-types-compressed")
    rels = _member("_rels/.rels", 50, _digest("rels"), "rels-compressed")
    source_members = [content_types, rels, main_source, opaque]
    candidate_members = [content_types, rels, main_candidate, opaque]
    return {
        "generator": analyze.GENERATOR,
        "format": analyze.FORMAT,
        "count": count,
        "opaque_path": analyze.OPAQUE_PATH,
        "opaque_bytes": analyze.OPAQUE_BYTES,
        "source_archive_bytes": len(source_xml) + analyze.OPAQUE_BYTES + 100_000,
        "source_archive_sha256": source_hash,
        "candidate_archive_bytes": len(source_xml) + analyze.OPAQUE_BYTES + 100_100,
        "candidate_archive_sha256": candidate_hash,
        "source_main_xml_bytes": len(source_xml),
        "source_main_xml_sha256": hashlib.sha256(source_xml).hexdigest(),
        "candidate_main_xml_bytes": len(candidate_xml),
        "candidate_main_xml_sha256": hashlib.sha256(candidate_xml).hexdigest(),
        "source_main_xml_archive_verified": True,
        "candidate_main_xml_archive_verified": True,
        "candidate_main_xml_oracle_verified": True,
        "source_members": source_members,
        "candidate_members": candidate_members,
        "source_member_count": 4,
        "candidate_member_count": 4,
        "source_semantic": analyze._semantic(count),
        "candidate_semantic": analyze._semantic(count, True),
        "source_unchanged_verified": True,
        "source_semantic_reopen_verified": True,
        "candidate_semantic_reopen_verified": True,
        "tail_copy_exactly_one_verified": True,
        "untouched_members_verified": True,
        "opaque_member_exact_verified": True,
        "physical_order_verified": True,
        "limits": {
            "max_xml_bytes": len(source_xml) + 1024 * 1024,
            "max_paragraphs": count + 1,
            "max_events": (count + 1) * 8 + 128,
            "max_depth": 16,
            "max_output_bytes": len(source_xml) + 2 * 1024 * 1024,
            "max_durable_bytes": analyze.MAX_DURABLE_BYTES,
        },
        "patch": {
            "copied_source_position": 0,
            "copied_before_position": count,
            "copied_paragraphs": 1,
            "copied_bytes": len(analyze._fragment(0)),
            "durable_bytes": 100,
            "durable_canonical_verified": True,
            "replay_verified": True,
            "inverse_verified": True,
            "stale_source_refusal_verified": True,
            "publication_inverse_verified": True,
        },
    }


def _sink(corpus: dict[str, object]) -> dict[str, object]:
    calls = (int(corpus["candidate_archive_bytes"]) + (16 * 1024 - 1)) // (16 * 1024)
    return {
        "accepted_bytes": corpus["candidate_archive_bytes"],
        "write_calls": calls,
        "largest_write": 16 * 1024,
        "histogram": {
            "bytes_0": 0,
            "bytes_1_to_512": 0,
            "bytes_513_to_4096": 0,
            "bytes_4097_to_16384": calls,
            "bytes_16385_to_65536": 0,
            "bytes_over_65536": 0,
        },
        "sha256": corpus["candidate_archive_sha256"],
    }


def _phase(instrumentation: str, reads: dict[str, object] | None = None, *, before: int = 100, allocated: int = 200, deallocated: int = 200, after: int = 100, peak: int = 200) -> dict[str, object]:
    return {
        "elapsed_ns": 10,
        "source_reads": reads if reads is not None else _reads(0, 0, 0),
        "allocation": _allocation(instrumentation == "allocator", before=before, allocated=allocated, deallocated=deallocated, after=after, peak=peak),
        "process": _process(),
    }


def _report(spec: dict[str, object]) -> dict[str, object]:
    count = int(spec["count"])
    instrumentation = str(spec["instrumentation"])
    mode = str(spec["mode"])
    corpus = _corpus(count)
    binary = {
        "binary": "litchi-perf-baseline" if instrumentation == "normal" else "litchi-perf-baseline-alloc",
        "allocator": "Rust system allocator" if instrumentation == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "instrumentation": "none" if instrumentation == "normal" else "system_allocator_operation_scoped",
        "counter_revision": None if instrumentation == "normal" else "serialized_region_peak_v3",
    }
    total = []
    phases = []
    for sample in range(analyze.SAMPLES):
        reads = _reads()
        sink = _sink(corpus)
        if mode == "total":
            total.append({
                "sample": sample,
                "elapsed_ns": 10,
                "source_reads": reads,
                "sink": sink,
                "allocation": _allocation(instrumentation == "allocator"),
                "process": _process(),
            })
        else:
            phases.append({
                "sample": sample,
                "open": _phase(instrumentation, _reads(), before=100, allocated=500, deallocated=100, after=500, peak=500),
                "snapshot": _phase(instrumentation, before=500, allocated=300, deallocated=0, after=800, peak=800),
                "stage": _phase(instrumentation, before=800, allocated=100, deallocated=0, after=900, peak=900),
                "commit": _phase(instrumentation, before=900, allocated=0, deallocated=200, after=700, peak=900),
                "publish": _phase(instrumentation, before=700, allocated=100, deallocated=0, after=800, peak=800),
                "drop": _phase(instrumentation, before=800, allocated=0, deallocated=700, after=100, peak=800),
                "source_reads": reads,
                "sink": sink,
            })
    return {
        "schema": analyze.REPORT_SCHEMA,
        "version": 1,
        "binary": binary,
        "config": {
            "counts": [count],
            "samples": analyze.SAMPLES,
            "warmups": analyze.WARMUPS,
            "mode": mode,
            "lifecycle_phases": list(analyze.PHASES),
            "sink": analyze.SINK_ID,
            "source": analyze.SOURCE_ID,
        },
        "cases": [{"count": count, "corpus": corpus, "total_samples": total or None, "phase_samples": phases or None}],
    }


class EvidenceBoundaryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.spec = analyze.expected_captures()[0]

    def test_valid_total_report_is_accepted(self) -> None:
        checked = analyze.validate_report(_report(self.spec), self.spec)
        self.assertEqual(len(checked["samples"]), analyze.SAMPLES)

    def test_valid_phase_report_is_accepted(self) -> None:
        spec = analyze.expected_captures()[1]
        checked = analyze.validate_report(_report(spec), spec)
        self.assertEqual(checked["samples"][0]["publish"]["source_reads"]["calls"], 0)

    def test_allocator_phase_growth_and_release_is_accepted(self) -> None:
        # Each synthetic phase deliberately retains a different live-byte
        # boundary.  This guards against requiring every phase to return to
        # zero while still checking continuity and final release.
        spec = analyze.expected_captures()[7]  # allocator/64/phases
        checked = analyze.validate_report(_report(spec), spec)
        allocation = checked["samples"][0]["snapshot"]["allocation"]
        self.assertEqual(allocation["live_bytes_before"], 500)
        self.assertEqual(allocation["live_bytes_after"], 800)

    def test_explicit_normal_unavailable_sample_is_accepted(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["total_samples"][0]["allocation"] = {
            "status": "unavailable",
            "scope": "operation_global_system_allocator",
        }
        checked = analyze.validate_report(report, self.spec)
        self.assertEqual(checked["samples"][0]["allocation"]["status"], "unavailable")

    def test_external_rss_parser_requires_one_gnu_time_observation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "captures").mkdir()
            label = self.spec["label"]
            path = root / "captures" / f"{label}.resource"
            path.write_text("Maximum resident set size (kbytes): 42\n", encoding="utf-8")
            original = analyze.ROOT
            try:
                analyze.ROOT = root
                self.assertEqual(analyze._resource_rss(self.spec), 42)
                path.write_text(
                    "Maximum resident set size (kbytes): 42\n"
                    "Maximum resident set size (kbytes): 43\n",
                    encoding="utf-8",
                )
                with self.assertRaises(analyze.AnalysisError):
                    analyze._resource_rss(self.spec)
            finally:
                analyze.ROOT = original

    def test_corpus_xml_oracle_rejects_mutation(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["corpus"]["candidate_main_xml_sha256"] = _digest("wrong")
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_read_histogram_rejects_unaccounted_call(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["total_samples"][0]["source_reads"]["request_histogram"]["bytes_1_to_512"] = 0
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_sink_rejects_oversized_write(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["total_samples"][0]["sink"]["largest_write"] = 16 * 1024 + 1
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_normal_allocation_numbers_are_not_silently_accepted(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["total_samples"][0]["allocation"] = _allocation(True)
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_sample_indices_must_be_exact(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["total_samples"][1]["sample"] = 0
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_protocol_order_and_mode_are_frozen(self) -> None:
        protocol = {
            "schema": analyze.PROTOCOL_SCHEMA,
            "samples": analyze.SAMPLES,
            "warmups": analyze.WARMUPS,
            "cpu": analyze.CPU,
            "append_count": 1,
            "normal_and_allocator_timings_separate": True,
            "phase_peaks_are_not_total_peaks": True,
            "regression_review_percent": 5,
            "performance_claim": "none",
            "comparison": "comparison",
            "scope": "scope",
            "environment": {key: "x" for key in ("RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "DEBUGINFOD_URLS", "LC_ALL")},
            "scripts": {"common.py": "a" * 64, "capture.py": "b" * 64},
            "captures": analyze.expected_captures(),
        }
        protocol["captures"][0] = dict(protocol["captures"][0], mode="phases")
        with self.assertRaises(analyze.AnalysisError):
            analyze.protocol_rows(protocol)

    def test_portable_path_rejects_escape(self) -> None:
        with self.assertRaises(verify.VerificationError):
            verify.safe_relative("../outside", "mutation")


if __name__ == "__main__":
    unittest.main()
