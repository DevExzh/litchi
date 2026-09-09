#!/usr/bin/env python3
"""Focused standard-library tests for the 0483 evidence boundary."""

from __future__ import annotations

import copy
import hashlib
import json
import unittest
import zlib

import analyze
import capture
import verify


def _digest(text: str) -> str:
    return hashlib.sha256(text.encode()).hexdigest()


def _member(path: str, decoded_bytes: int, decoded_sha: str) -> dict[str, object]:
    return {
        "path": path,
        "compression_method": "store",
        "data_descriptor": False,
        "crc32": 1,
        "decoded_bytes": decoded_bytes,
        "decoded_sha256": decoded_sha,
        "compressed_bytes": decoded_bytes,
        "compressed_sha256": decoded_sha,
    }


def _corpus(count: int, route: str = "materialized") -> dict[str, object]:
    source = analyze.source_xml(count)
    source_sha = hashlib.sha256(source).hexdigest()
    opaque = analyze.opaque_payload()
    opaque_sha = hashlib.sha256(opaque).hexdigest()
    members = [
        _member("[Content_Types].xml", 10, _digest("types")),
        _member("_rels/.rels", 10, _digest("rels")),
        _member(analyze.MAIN_PATH, len(source), source_sha),
        _member(analyze.OPAQUE_PATH, analyze.OPAQUE_BYTES, opaque_sha),
    ]
    members[3]["crc32"] = zlib.crc32(opaque) & 0xFFFF_FFFF

    def route_record(route_name: str) -> dict[str, object]:
        candidate = (
            analyze.candidate_xml(count)
            if route_name == "materialized"
            else analyze.bounded_candidate_xml(count)
        )
        candidate_sha = hashlib.sha256(candidate).hexdigest()
        candidate_members = copy.deepcopy(members)
        candidate_members[2] = _member(analyze.MAIN_PATH, len(candidate), candidate_sha)
        candidate_members[3]["crc32"] = members[3]["crc32"]
        output_members = candidate_members
        scalar_proof = None
        if route_name == "bounded":
            body_offset = source.index(b"</w:body>")
            empty_sha = _digest("")
            scalar_proof = {
                "source_version_id": 1,
                "source_version_revision": 0,
                "source_len": len(source),
                "source_sha256": source_sha,
                "insertion_offset": body_offset,
                "source_paragraph_count": count,
                "source_event_count": 100,
                "source_max_depth": 2,
                "source_strict_namespace": False,
                "source_sect_pr_len": 0,
                "source_sect_pr_sha256": empty_sha,
                "candidate_len": len(candidate),
                "candidate_sha256": candidate_sha,
                "candidate_paragraph_count": count + 1,
                "candidate_event_count": 101,
                "candidate_max_depth": 2,
                "generated_offset": body_offset,
                "generated_once": True,
                "candidate_sect_pr_len": 0,
                "candidate_sect_pr_sha256": empty_sha,
            }
        return {
            "route": analyze.ROUTE_NAMES[route_name],
            "output_archive_bytes": 110,
            "output_archive_sha256": _digest(f"{route_name} archive"),
            "output_main_xml_bytes": len(candidate),
            "output_main_xml_sha256": candidate_sha,
            "output_main_xml_expected_semantic_verified": True,
            "output_main_xml_expected_raw_verified": True,
            "scalar_proof": scalar_proof,
            "output_semantic": analyze.semantic(count, True),
            "output_members": output_members,
            "output_member_count": 4,
            "output_untouched_members_verified": True,
            "output_opaque_member_exact_verified": True,
            "output_physical_order_verified": True,
            "output_main_compressed_equal_source": False,
            "route_output_replay_verified": True,
            "route_inverse_verified": True,
            "stale_source_refusal_verified": True,
        }
    return {
        "generator": analyze.GENERATOR,
        "format": analyze.FORMAT,
        "count": count,
        "append_text": analyze.append_text(count),
        "append_text_sha256": hashlib.sha256(analyze.append_text(count).encode()).hexdigest(),
        "source_archive_bytes": 100,
        "source_archive_sha256": _digest("source archive"),
        "source_main_xml_bytes": len(source),
        "source_main_xml_sha256": source_sha,
        "source_main_xml_archive_verified": True,
        "source_members": members,
        "source_member_count": 4,
        "source_semantic": analyze.semantic(count),
        "source_unchanged_verified": True,
        "source_opaque_member_exact_verified": True,
        "route_semantics_equal_verified": True,
        "route_untouched_members_equal_verified": True,
        "route_outputs_physical_equal": False,
        "materialized": route_record("materialized"),
        "bounded": route_record("bounded"),
        "limits": {
            "copy_max_xml_bytes": len(source) + 1024 * 1024,
            "copy_max_paragraphs": count + 1,
            "copy_max_events": (count + 1) * 8 + 128,
            "copy_max_depth": 16,
            "copy_max_output_bytes": len(source) + 2 * 1024 * 1024,
            "bounded": "finite bounded tail policy",
        },
    }


def _sample(count: int, candidate_sha: str, allocation=None) -> dict[str, object]:
    return {
        "sample": 0,
        "elapsed_ns": 100,
        "source_reads": {
            "calls": 1,
            "requested_bytes": 16,
            "returned_bytes": 16,
            "request_histogram": {
                "bytes_0": 0,
                "bytes_1_to_512": 1,
                "bytes_513_to_4096": 0,
                "bytes_4097_to_16384": 0,
                "bytes_16385_to_65536": 0,
                "bytes_over_65536": 0,
            },
        },
        "sink": {
            "accepted_bytes": 110,
            "write_calls": 1,
            "largest_write": 110,
            "histogram": {
                "bytes_0": 0,
                "bytes_1_to_512": 1,
                "bytes_513_to_4096": 0,
                "bytes_4097_to_16384": 0,
                "bytes_16385_to_65536": 0,
                "bytes_over_65536": 0,
            },
            "sha256": candidate_sha,
        },
        "allocation": allocation,
        "process": None,
    }


def _report(spec: dict[str, object], *, sample_count: int = analyze.SAMPLES) -> dict[str, object]:
    corpus = _corpus(int(spec["count"]), str(spec["route"]))
    samples = []
    for index in range(sample_count):
        sample = _sample(
            int(spec["count"]),
            str(corpus[spec["route"]]["output_archive_sha256"]),
        )
        sample["sample"] = index
        sample["route"] = spec["route_name"]
        samples.append(sample)
    return {
        "schema": "docx-bounded-tail-append-comparison-v1",
        "version": 1,
        "binary": {
            "binary": "litchi-perf-baseline-alloc" if spec["instrumentation"] == "allocator" else "litchi-perf-baseline",
            "allocator": "CountingSystemAllocator(std::alloc::System)" if spec["instrumentation"] == "allocator" else "Rust system allocator",
            "instrumentation": "system_allocator_operation_scoped" if spec["instrumentation"] == "allocator" else "none",
            "counter_revision": "serialized_region_peak_v3" if spec["instrumentation"] == "allocator" else None,
        },
        "config": {
            "counts": [spec["count"]],
            "samples": sample_count,
            "warmups": analyze.WARMUPS,
            "route": spec["route_name"],
            "lifecycle": "timed lifecycle",
            "text_authoring": "bounded caller text",
            "sink": analyze.SINK_ID,
            "source": analyze.SOURCE_ID,
        },
        "cases": [{
            "count": spec["count"],
            "corpus": corpus,
            "routes": [{
                "route": spec["route_name"],
                "corpus": corpus[spec["route"]],
                "samples": samples,
            }],
        }],
    }


class EvidenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.spec = {
            "label": "b1-normal-64-bounded",
            "arm": "b1",
            "repeat": 1,
            "route": "bounded",
            "route_name": "bounded_plain_text_tail_append",
            "instrumentation": "normal",
            "count": 64,
            "attempt": "accepted",
        }

    def test_protocol_plan_has_reversed_second_repeat(self) -> None:
        rows = capture.expected_captures("accepted")
        self.assertEqual(len(rows), 24)
        self.assertEqual([row["count"] for row in rows if row["arm"] == "a1" and row["instrumentation"] == "normal"], [64, 8192, 131072])
        self.assertEqual([row["count"] for row in rows if row["arm"] == "b2" and row["instrumentation"] == "normal"], [131072, 8192, 64])
        self.assertEqual([row["route_name"] for row in rows if row["arm"] == "b1"], ["bounded_plain_text_tail_append"] * 6)

    def test_independent_xml_and_semantic_oracles(self) -> None:
        self.assertNotEqual(analyze.source_xml(64), analyze.candidate_xml(64))
        self.assertNotEqual(analyze.candidate_xml(64), analyze.bounded_candidate_xml(64))
        self.assertEqual(analyze.semantic(64)["paragraph_count"], 64)
        self.assertEqual(analyze.semantic(64, True)["paragraph_count"], 65)
        self.assertEqual(
            analyze.semantic(64, True)["text_bytes"],
            analyze.semantic(64)["text_bytes"] + len(analyze.append_text(64).encode()),
        )

    def test_normal_report_validates_and_route_is_explicit(self) -> None:
        report = _report(self.spec)
        checked = analyze.validate_report(report, self.spec)
        self.assertEqual(checked["route_name"], "bounded_plain_text_tail_append")
        self.assertEqual(len(checked["samples"]), analyze.SAMPLES)

    def test_pilot_report_uses_same_oracles_with_small_sample_shape(self) -> None:
        report = _report(self.spec, sample_count=1)
        report["config"]["warmups"] = 1
        checked = analyze.validate_report(
            report,
            self.spec,
            expected_samples=1,
            expected_warmups=1,
        )
        self.assertEqual(len(checked["samples"]), 1)

    def test_pilot_argv_binds_binary_route_shape_and_portable_report_suffix(self) -> None:
        pilot_spec = dict(self.spec, label="pilot-accepted")
        pilot = {
            "path": "pilots/pilot-accepted.report.json",
            "samples": 1,
            "warmups": 1,
            "spec": pilot_spec,
        }
        binary = {"path": "/tmp/litchi-goal-0483/accepted/normal/docx_bounded_tail_append_compare"}
        argv = [
            binary["path"], "--route", "bounded", "--counts", "64", "--samples", "1",
            "--warmups", "1", "--json", "/home/zhuhe/code/litchi/docs/performance/results/change-0483/pilots/pilot-accepted.report.json",
        ]
        verify.check_pilot_argv(
            argv,
            {"cwd": "/home/zhuhe/code/litchi"},
            pilot,
            binary,
            verify.ROOT / pilot["path"],
            "pilot-accepted",
        )
        timed_argv = [
            "/usr/bin/time", "-v", "-o", "/home/zhuhe/code/litchi/docs/performance/results/change-0483/pilots/pilot-accepted.resource",
            "/usr/bin/taskset", "-c", "2", binary["path"], "--route", "bounded", "--counts", "64",
            "--samples", "1", "--warmups", "1", "--json",
            "/home/zhuhe/code/litchi/docs/performance/results/change-0483/pilots/pilot-accepted.report.json",
        ]
        verify.check_pilot_argv(
            timed_argv,
            {"cwd": "/home/zhuhe/code/litchi"},
            pilot,
            binary,
            verify.ROOT / pilot["path"],
            "pilot-accepted",
        )
        with self.assertRaises(verify.VerificationError):
            verify.check_pilot_argv(
                argv[:-1] + ["/home/zhuhe/code/litchi/docs/performance/results/change-0483/pilots/wrong.report.json"],
                {"cwd": "/home/zhuhe/code/litchi"},
                pilot,
                binary,
                verify.ROOT / pilot["path"],
                "pilot-accepted",
            )

    def test_validation_plan_classifies_successful_developmental_receipt(self) -> None:
        digest = "a" * 64
        protocol = {
            "validation": {
                "required_labels": ["build-accepted"],
                "pilot_labels": ["pilot-accepted"],
                "argv": {"build-accepted": ["cargo"], "pilot-accepted": ["binary"]},
                "pilot_reports": {
                    "pilot-accepted": {
                        "path": "pilots/pilot-accepted.report.json",
                        "bytes": 1,
                        "sha256": digest,
                        "spec": {
                            "label": "pilot-accepted",
                            "route": "bounded",
                            "route_name": "bounded_plain_text_tail_append",
                            "instrumentation": "normal",
                            "count": 64,
                            "attempt": "accepted",
                        },
                        "samples": 1,
                        "warmups": 1,
                    }
                },
                "developmental": {
                    "xml-audit-dev-01": {
                        "classification": "developmental",
                        "reason": "source changed during concurrent edits",
                        "source_before_sha256": digest,
                        "source_after_sha256": "b" * 64,
                        "current_source_differs": True,
                    }
                },
            }
        }
        plan = verify.validation_plan(protocol)
        self.assertEqual(plan["developmental"]["xml-audit-dev-01"]["classification"], "developmental")

    def test_validation_plan_accepts_no_developmental_entries_when_none_exist(self) -> None:
        protocol = {"validation": {"required_labels": ["accepted"], "pilot_labels": [], "argv": {"accepted": []}, "pilot_reports": {}, "developmental": {}}}
        self.assertEqual(verify.validation_plan(protocol)["labels"], ["accepted"])

    def test_report_rejects_route_mismatch(self) -> None:
        report = _report(self.spec)
        report["config"]["route"] = "materialized_paragraph_copy"
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_sink_rejects_digest_mismatch(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["routes"][0]["samples"][0]["sink"]["sha256"] = _digest("wrong")
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_sink_rejects_histogram_weighted_minimum_mismatch(self) -> None:
        report = _report(self.spec)
        sink = report["cases"][0]["routes"][0]["samples"][0]["sink"]
        sink["histogram"]["bytes_1_to_512"] = 0
        sink["histogram"]["bytes_513_to_4096"] = 1
        sink["largest_write"] = 600
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_sink_rejects_histogram_weighted_maximum_mismatch(self) -> None:
        sink = copy.deepcopy(_sample(64, _digest("candidate"))["sink"])
        sink["accepted_bytes"] = 5_000
        sink["largest_write"] = 500
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_sink(sink, 5_000, sink["sha256"], "corrupt.sink")

    def test_sink_rejects_nonzero_bucket_above_short_write_limit(self) -> None:
        report = _report(self.spec)
        sink = report["cases"][0]["routes"][0]["samples"][0]["sink"]
        sink["histogram"]["bytes_1_to_512"] = 0
        sink["histogram"]["bytes_16385_to_65536"] = 1
        sink["largest_write"] = 16_384
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_sink_rejects_largest_write_inconsistent_with_accepted_bytes(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["routes"][0]["samples"][0]["sink"]["largest_write"] = 111
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_opaque_member_is_authenticated_independently(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["corpus"]["source_members"][3]["decoded_sha256"] = _digest("wrong")
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_bounded_scalar_proof_is_authenticated_independently(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["corpus"]["bounded"]["scalar_proof"]["insertion_offset"] += 1
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_read_histogram_rejects_unaccounted_call(self) -> None:
        report = _report(self.spec)
        report["cases"][0]["routes"][0]["samples"][0]["source_reads"]["request_histogram"]["bytes_1_to_512"] = 0
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, self.spec)

    def test_allocator_conservation_is_checked(self) -> None:
        allocator_spec = dict(self.spec, instrumentation="allocator")
        report = _report(allocator_spec)
        allocation = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            **{field: 0 for field in analyze.ALLOC_FIELDS},
        }
        allocation["allocated_bytes"] = 10
        allocation["live_bytes_after"] = 0
        report["cases"][0]["routes"][0]["samples"][0]["allocation"] = allocation
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(report, allocator_spec)

    def test_allocator_absolute_peak_cannot_regress(self) -> None:
        allocation = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            **{field: 0 for field in analyze.ALLOC_FIELDS},
        }
        allocation["peak_live_bytes_before"] = 1
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_allocation(allocation, "allocator", "corrupt.allocation")

    def test_allocator_region_peak_cannot_exceed_absolute_peak_after(self) -> None:
        allocation = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            **{field: 0 for field in analyze.ALLOC_FIELDS},
        }
        allocation["peak_live_bytes_after"] = 1
        allocation["region_peak_live_bytes"] = 2
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_allocation(allocation, "allocator", "corrupt.allocation")

    def test_percentiles_are_derived(self) -> None:
        result = analyze._stats([1, 2, 3, 4, 5])
        self.assertEqual(result["p50"], 3.0)
        self.assertEqual(result["p95"], 4.8)
        self.assertEqual(result["p99"], 4.96)

    def test_safe_relative_rejects_escape(self) -> None:
        with self.assertRaises(verify.VerificationError):
            verify.safe_relative("../outside", "path")


if __name__ == "__main__":
    unittest.main()
