#!/usr/bin/env python3
"""Pure synthetic tests for :mod:`heap_reader`.

These tests exercise the decoded line protocol directly.  They intentionally
do not build the Rust probe, launch Heaptrack, or depend on a retained capture;
the capture driver can therefore run them before any profiled workload exists.
"""

from __future__ import annotations

import copy
from pathlib import Path
import sys
import unittest


PACKET = Path(__file__).resolve().parent
if str(PACKET) not in sys.path:
    sys.path.insert(0, str(PACKET))

import heap_reader as reader


OWNER = reader.PROFILE_OWNER


def _sized(text: str) -> str:
    return f"s {len(text.encode('utf-8')):x} {text}"


def synthetic_trace() -> bytes:
    """Return a valid trace with exact/hash/phase/unknown examples."""

    strings = [
        "/tmp/opc-preservation-probe",
        OWNER,
        "soapberry_zip::PreservationIndex::new",
        "std::vec::Vec::push",
        "soapberry_zip::PreservationIndex::prepare_with",
        "soapberry_zip::PreservationIndex::write_prepared_local",
        "/src/probe.rs",
        OWNER + "::h0123456789abcdef",
        "opc_preservation_probe::profile_publish_extra",
    ]
    lines = ["v 10500 3", "I 1000 1"]
    lines.extend(_sized(value) for value in strings)
    # One instruction pointer carries an inline frame.  Both frames must
    # survive ancestry expansion in the result.
    lines.extend(
        [
            "i 1 1 3 7 1 4 7 2",
            "i 2 1 2",
            "i 3 1 5",
            "i 4 1 6",
            "i 5 1 8",
            "i 6 1 9",
            "i 7 0 0",
            "t 1 0",
            "t 2 1",
            "t 3 0",
            "t 2 3",
            "t 4 0",
            "t 2 5",
            "t 5 0",
            "t 6 0",
            "t 7 0",
            "a a 2",
            "+ 0",
            "+ 0",
            "- 0",
            "a 14 4",
            "+ 1",
            "a 1e 6",
            "+ 2",
            "a 28 7",
            "+ 3",
            "a 32 8",
            "+ 4",
            "a 3c 9",
            "+ 5",
            "a 46 0",
            "+ 6",
            "a 50 3",
            "+ 7",
            "a 5a 5",
            "+ 8",
        ]
    )
    return ("\n".join(lines) + "\n").encode("utf-8")


class HeapReaderTests(unittest.TestCase):
    def test_exact_stacks_owner_totals_and_disjoint_phases(self) -> None:
        report = reader.parse(synthetic_trace())

        self.assertEqual(report["schema"], reader.SCHEMA)
        self.assertEqual(report["header"], {"heaptrack_version": 0x10500, "file_version": 3})
        whole = report["whole_process"]
        owner = report["owner_attribution"]
        self.assertEqual(whole["allocation_events"], 10)
        self.assertEqual(whole["requested_bytes"], 2 * 10 + 20 + 30 + 40 + 50 + 60 + 70 + 80 + 90)
        self.assertEqual(owner["allocation_events"], 5)
        self.assertEqual(owner["requested_bytes"], 2 * 10 + 20 + 30 + 40)
        self.assertEqual(
            whole["allocation_events"],
            report["conservation"]["allocation_events"]["matched_owner"]
            + report["conservation"]["allocation_events"]["unmatched"],
        )
        self.assertEqual(
            whole["requested_bytes"],
            report["conservation"]["requested_bytes"]["matched_owner"]
            + report["conservation"]["requested_bytes"]["unmatched"],
        )

        rows = report["allocation_rows"]
        self.assertEqual([row["allocation_info_index"] for row in rows], list(range(9)))
        first = rows[0]
        self.assertTrue(first["owner_match"])
        self.assertEqual(first["owner_match_count"], 1)
        self.assertEqual(first["phase"], "preservation_index")
        self.assertEqual(first["innermost_owner_frame_index"], 0)
        self.assertEqual(first["frames_leaf_to_root"][0]["function"], OWNER)
        self.assertTrue(first["frames_leaf_to_root"][2]["inlined"])
        self.assertEqual(
            first["frames_leaf_to_root"][2]["function"], "std::vec::Vec::push"
        )

        groups = report["phase_groups"]["groups"]
        self.assertTrue(report["phase_groups"]["disjoint"])
        self.assertEqual(
            {phase: groups[phase]["allocation_info_indexes"] for phase in reader.PHASES},
            {
                "preservation_index": [0],
                "topology_preparation": [],
                "plan": [1],
                "emission": [2],
                "unknown": [3],
            },
        )
        self.assertEqual(
            sum(groups[phase]["requested_bytes"] for phase in reader.PHASES),
            owner["requested_bytes"],
        )
        self.assertEqual(
            sum(groups[phase]["allocation_events"] for phase in reader.PHASES),
            owner["allocation_events"],
        )

    def test_hash_suffix_is_supported_but_near_miss_is_not(self) -> None:
        report = reader.parse(synthetic_trace())
        hash_row = report["allocation_rows"][3]
        self.assertTrue(hash_row["owner_match"])
        self.assertEqual(hash_row["owner_match_kinds"], ["rust_hash"])
        self.assertEqual(report["diagnostics"]["owner_hash_matches"]["allocation_events"], 1)
        near_miss = report["allocation_rows"][6]
        self.assertFalse(near_miss["owner_match"])
        self.assertEqual(near_miss["phase"], "unmatched")

    def test_unknown_frames_and_empty_traces_are_retained(self) -> None:
        report = reader.parse(synthetic_trace())
        unknown = report["allocation_rows"][5]
        self.assertFalse(unknown["owner_match"])
        self.assertEqual(unknown["trace_id"], 9)
        self.assertIsNone(unknown["frames_leaf_to_root"][0]["module"])
        self.assertIsNone(unknown["frames_leaf_to_root"][0]["function"])
        empty = report["allocation_rows"][6]
        self.assertEqual(empty["trace_id"], 0)
        self.assertEqual(empty["frames_leaf_to_root"], [])
        self.assertGreaterEqual(
            report["diagnostics"]["unknown_frames"]["missing_module"]["allocation_events"],
            1,
        )
        self.assertGreaterEqual(
            report["diagnostics"]["unknown_frames"]["empty_trace"]["allocation_events"],
            1,
        )

    def test_hash_and_inline_frames_are_not_dropped_from_full_stack(self) -> None:
        report = reader.parse(synthetic_trace())
        stack = report["allocation_rows"][0]["frames_leaf_to_root"]
        self.assertEqual(
            [frame["function"] for frame in stack],
            [OWNER, "soapberry_zip::PreservationIndex::new", "std::vec::Vec::push"],
        )
        self.assertEqual([frame["index"] for frame in stack], [0, 1, 2])

    def test_phase_buckets_use_concrete_symbols_only(self) -> None:
        self.assertEqual(
            reader._phase_matches("SourceBackedPackage::prepare_topology_state"),
            ("topology_preparation",),
        )
        self.assertEqual(
            reader._phase_matches(
                "soapberry_zip::PreservationIndex$LT$R$GT$::new_with_limits"
            ),
            ("preservation_index",),
        )
        self.assertEqual(
            reader._phase_matches("SourceBackedPackage::write_topology_to_stream"),
            (),
        )
        self.assertEqual(
            reader._phase_matches("PreservationIndex::validate_output_layout"),
            ("plan",),
        )
        self.assertEqual(
            reader._phase_matches("PreservationIndex::write_prepared_central"),
            ("emission",),
        )

    def assert_rejected(self, payload: bytes, text: str, **kwargs: object) -> None:
        with self.assertRaises(reader.HeapReaderError) as context:
            reader.parse(payload, **kwargs)
        self.assertIn(text, str(context.exception))

    def test_mutations_reject_version_and_unknown_references(self) -> None:
        base = synthetic_trace()
        self.assert_rejected(base.replace(b"v 10500 3", b"v 10400 3", 1), "unsupported")
        self.assert_rejected(base + b"+ ffff\n", "unknown + allocation")
        self.assert_rejected(base.replace(b"t 1 0", b"t 1 1", 1), "forward or cyclic")
        self.assert_rejected(base.replace(b"a a 2", b"a a ff", 1), "unknown allocation trace")
        self.assert_rejected(base.replace(b"i 2 1 2", b"i 2 1 ff", 1), "unknown function string")
        self.assert_rejected(base + b"q 0\n", "unknown Heaptrack record mode")

    def test_mutations_reject_malformed_strings_and_cycles(self) -> None:
        base = synthetic_trace()
        original_text = "/tmp/opc-preservation-probe"
        original = _sized(original_text).encode("utf-8")
        mutated = f"s {len(original_text.encode('utf-8')) + 1:x} {original_text}".encode("utf-8")
        self.assert_rejected(base.replace(original, mutated, 1), "string byte length")
        self.assert_rejected(base.replace(b"+ 0", b"+ 0 0", 1), "needs one field")
        self.assert_rejected(base.replace(b"a a 2", b"a -1 2", 1), "invalid allocation size")
        self.assert_rejected(base.replace(b"t 2 1", b"t 2 2", 1), "forward or cyclic")

    def test_limits_are_fail_closed(self) -> None:
        base = synthetic_trace()
        self.assert_rejected(base, "decoded trace exceeds", max_bytes=len(base) - 1)
        self.assert_rejected(base, "more than 2 records", max_records=2)
        self.assert_rejected(base, "more than 1 frames", max_stack_frames=1)
        self.assert_rejected(base, "exceeds 1 trace nodes", max_depth=1)
        self.assert_rejected(base, "max_records must be positive", max_records=0)

    def test_parse_heap_trace_alias_accepts_decoded_bytes(self) -> None:
        report = reader.parse_heap_trace(synthetic_trace())
        self.assertEqual(report["whole_process"]["allocation_events"], 10)


if __name__ == "__main__":
    unittest.main(verbosity=2)
