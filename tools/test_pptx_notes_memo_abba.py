#!/usr/bin/env python3
"""Meaningful synthetic protocol tests for ``pptx_notes_memo_abba``."""

from __future__ import annotations

import copy
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import pptx_notes_memo_abba as validator  # noqa: E402


def statistics(values: list[int]) -> dict[str, object]:
    ordered = sorted(enumerate(values), key=lambda item: (item[1], item[0]))
    sorted_values = [value for _index, value in ordered]
    sample_order = [index for index, _value in ordered]
    count = len(values)
    p50 = (
        sorted_values[(count - 1) // 2] // 2
        + sorted_values[count // 2] // 2
        + (
            sorted_values[(count - 1) // 2] % 2
            + sorted_values[count // 2] % 2
        )
        // 2
    )
    return {
        "unit": "ns",
        "samples": sorted_values,
        "sample_order": sample_order,
        "min": min(sorted_values),
        "p50": p50,
        "p95": sorted_values[max(1, (count * 95 + 99) // 100) - 1],
        "p99": sorted_values[max(1, (count * 99 + 99) // 100) - 1],
        "max": max(sorted_values),
        "mean": float(sum(values)) / count,
        "standard_deviation": 1.0,
        "confidence_interval_95": {
            "method": "synthetic test evidence",
            "lower": 1.0,
            "upper": 2.0,
        },
    }


def fidelity(case: str) -> dict[str, object]:
    changed_slide = case == validator.CASES[2]
    warm_notes = case == validator.CASES[1]
    changed_members = (
        []
        if case == validator.CASES[0]
        else [
            "ppt/notesSlides/notesSlide2.xml"
            if warm_notes
            else "ppt/slides/slide1.xml"
        ]
    )
    raw = {
        "source_member_count": 20,
        "output_member_count": 20,
        "source_payload_sha256": "a" * 64,
        "output_payload_sha256": ("a" if case == validator.CASES[0] else "b") * 64,
        "member_names_equal": True,
        "untouched_payloads_equal": True,
        "relationship_member_count": 8,
        "relationship_names_equal": True,
        "relationship_payloads_equal": True,
        "master_theme_member_count": 4,
        "master_theme_payloads_equal": True,
        "changed_members": changed_members,
        "expected_changed_members": changed_members,
        "exact_changed_member_set": True,
        "changed_payloads_nonempty": True,
        "slide_root_count": 3,
        "notes_root_count": 3,
        "changed_slide_index": 0 if changed_slide else None,
        "changed_notes_slide_index": 1 if warm_notes else None,
        "untouched_slide_root_count": 2 if changed_slide else 3,
    }
    result = {
        "semantic": {
            "verified": True,
            "slide_count": 3,
            "notes_slide_count": 3,
            "slide_text_sha256": "c" * 64,
            "notes_text_sha256": "d" * 64,
        },
        "changed_commit_reported": case != validator.CASES[0],
        "output_reopened": True,
        "changed_slide_index": 0 if changed_slide else None,
        "untouched_slide_root_count": 2 if changed_slide else 3,
        "raw": raw,
    }
    if case != validator.CASES[0]:
        result["source_immutable"] = True
    return result


def report(revision: str, binary_sha: str) -> dict[str, object]:
    rows = []
    for row_index, case in enumerate(validator.CASES):
        elapsed = [100 + row_index * 20 + ((index * 7) % 16) for index in range(16)]
        observations = [
            {
                "sample_index": index,
                "elapsed_ns": value,
                "output_bytes": 512 + row_index,
                "output_sha256": ("e" if row_index == 0 else "f") * 64,
                "fidelity": fidelity(case),
                "allocation": {
                    "status": "unavailable",
                    "scope": "operation_global_system_allocator",
                },
            }
            for index, value in enumerate(elapsed)
        ]
        rows.append(
            {
                "case": case,
                "operation": validator.OPERATIONS[case],
                "timing_scope": validator.TIMING_SCOPES[case],
                "elapsed_ns": statistics(elapsed),
                "observations": observations,
            }
        )
    return {
        "schema": validator.SCHEMA,
        "tool": "litchi-perf-baseline",
        "binary_identity": {
            "path": f"/tmp/{binary_sha}",
            "binary_sha256": binary_sha,
            "binary_bytes": 100,
            "mode_bits": 493,
            "executable": True,
            "profile": "release",
        },
        "source_revision": revision,
        "instrumentation": "none",
        "allocator": "Rust system allocator",
        "allocator_counter_revision": None,
        "timing_claim": "native_wall_clock_samples; compare only under the approved quiet ABBA protocol",
        "allocation_scope": "one non-nested operation-scoped global-allocator region around each timed API operation",
        "fixture": {
            "generator": "pptx-notes-memo-three-slide-v1",
            "slides": 3,
            "notes_slides": 3,
            "authored_shapes_per_slide": 1,
            "bytes": 100,
            "sha256": "1" * 64,
            "recipe": "three authored slides; title notes slide {i}, one body text box, speaker notes {i}",
        },
        "samples": 16,
        "warmup": 2,
        "checked_iteration_count": 18,
        "memo_scope": "synthetic memo scope",
        "fidelity_scope": "synthetic fidelity scope",
        "results": rows,
    }


class NotesMemoAbbaTests(unittest.TestCase):
    def setUp(self) -> None:
        self.reports = {
            "a1": report("a" * 40, "1" * 64),
            "b1": report("b" * 40, "2" * 64),
            "b2": report("b" * 40, "2" * 64),
            "a2": report("a" * 40, "1" * 64),
        }

    def test_full_four_leg_four_round_packet_passes(self) -> None:
        summary = validator.validate_abba(self.reports)
        self.assertEqual(summary["rounds"], 4)
        self.assertEqual(summary["samples_per_round"], 4)
        self.assertEqual(len(summary["rows"]), 3)

    def test_short_packet_is_rejected(self) -> None:
        short = copy.deepcopy(self.reports)
        short["a1"]["samples"] = 12
        short["a1"]["checked_iteration_count"] = 14
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(short)

    def test_false_raw_oracle_is_rejected(self) -> None:
        false_oracle = copy.deepcopy(self.reports)
        false_oracle["b1"]["results"][0]["observations"][0]["fidelity"]["raw"][
            "untouched_payloads_equal"
        ] = False
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(false_oracle)

    def test_source_identity_pairing_is_rejected(self) -> None:
        wrong_pair = copy.deepcopy(self.reports)
        wrong_pair["b2"]["source_revision"] = "c" * 40
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(wrong_pair)

    def test_binary_identity_pairing_is_rejected(self) -> None:
        wrong_pair = copy.deepcopy(self.reports)
        wrong_pair["a2"]["binary_identity"]["binary_sha256"] = "9" * 64
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(wrong_pair)

    def test_control_and_candidate_binary_hashes_must_differ(self) -> None:
        same_binary = copy.deepcopy(self.reports)
        for leg in ("b1", "b2"):
            same_binary[leg]["binary_identity"]["binary_sha256"] = "1" * 64
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(same_binary)

    def test_unknown_keys_are_rejected_at_each_fixed_layer(self) -> None:
        paths = (
            (),
            ("binary_identity",),
            ("fixture",),
            ("results", 0),
            ("results", 0, "elapsed_ns"),
            ("results", 0, "elapsed_ns", "confidence_interval_95"),
            ("results", 0, "observations", 0),
            ("results", 0, "observations", 0, "fidelity"),
            ("results", 0, "observations", 0, "fidelity", "semantic"),
            ("results", 0, "observations", 0, "fidelity", "raw"),
            ("results", 0, "observations", 0, "allocation"),
        )
        for path in paths:
            candidate = copy.deepcopy(self.reports)
            target: object = candidate["a1"]
            for key in path:
                target = target[key]  # type: ignore[index]
            target["unexpected"] = True  # type: ignore[index]
            with self.subTest(path=path), self.assertRaises(validator.ValidationError):
                validator.validate_abba(candidate)

    def test_scope_and_fidelity_identity_are_required(self) -> None:
        wrong_scope = copy.deepcopy(self.reports)
        wrong_scope["a1"]["results"][0]["timing_scope"] = "Package::from_bytes plus opened_presentation"
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(wrong_scope)

        drifting_oracle = copy.deepcopy(self.reports)
        drifting_oracle["a1"]["results"][0]["observations"][1]["fidelity"]["semantic"][
            "slide_text_sha256"
        ] = "9" * 64
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(drifting_oracle)

    def test_allocator_packet_requires_measured_regions(self) -> None:
        allocator = copy.deepcopy(self.reports)
        measured = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            **{field: 0 for field in validator.ALLOCATION_NUMERIC_FIELDS},
        }
        for packet in allocator.values():
            packet["instrumentation"] = "system_allocator_operation_scoped"
            packet["allocator"] = "CountingSystemAllocator(std::alloc::System)"
            packet["allocator_counter_revision"] = "serialized_region_peak_v3"
            packet["timing_claim"] = "allocator-instrumented wall-clock samples; not a latency claim"
            for row in packet["results"]:
                for observation in row["observations"]:
                    observation["allocation"] = copy.deepcopy(measured)
        summary = validator.validate_abba(allocator, mode="allocator")
        self.assertEqual(summary["instrumentation"], "system_allocator_operation_scoped")
        allocator["a1"]["results"][0]["observations"][0]["allocation"].pop("allocated_bytes")
        with self.assertRaises(validator.ValidationError):
            validator.validate_abba(allocator, mode="allocator")


if __name__ == "__main__":
    unittest.main()
