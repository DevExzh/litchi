#!/usr/bin/env python3
"""Lightweight source-only tests for the generalized 0795 Callgrind reader."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from callgrind_parser import (
    CallgrindError,
    aggregate_edges,
    parse_callgrind,
    vector_mapping,
)


EVENTS = ("Ir", "Bc", "Bcm", "Bi", "Bim")


def write_profile(directory: Path, name: str, body: str) -> Path:
    path = directory / name
    path.write_text(body, encoding="utf-8")
    return path


class CallgrindParserTests(unittest.TestCase):
    def test_sparse_multi_event_vectors_and_branch_only_edge(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = write_profile(
                Path(temporary),
                "positive.callgrind.1",
                """# callgrind format
version: 1
cmd: probe
part: 1
desc: Trigger: --dump-after=owner
positions: line
events: Ir Bc Bcm Bi Bim
summary: 2 1 1 0 0
fn=(1) owner
0 2 1
cfn=(2) child
calls=3 4
* 8 4 1 2 7
fn=(2) child
0 0 0 1
totals: 2 1 1 0 0
""",
            )
            parsed = parse_callgrind(path, expected_events=EVENTS)
            self.assertEqual(parsed["self_total"], (2, 1, 1, 0, 0))
            self.assertEqual(parsed["statistics"]["cost_records"], 3)
            self.assertEqual(parsed["statistics"]["position_kinds"],
                             {"absolute": 2, "wildcard": 1})
            children = aggregate_edges(parsed, 1)
            self.assertEqual(len(children), 1)
            self.assertEqual(children[0]["calls"], 3)
            self.assertEqual(vector_mapping(children[0]["cost"], EVENTS),
                             {"Ir": 8, "Bc": 4, "Bcm": 1, "Bi": 2, "Bim": 7})

    def test_empty_termination_part_is_accepted_only_when_requested(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = write_profile(
                Path(temporary),
                "termination.callgrind",
                """# callgrind format
version: 1
cmd: probe
part: 2
desc: Trigger: Program termination
events: Ir Bc Bcm Bi Bim
summary: 0 0 0 0 0
totals: 0 0 0 0 0
""",
            )
            parsed = parse_callgrind(path, expected_events=EVENTS, allow_empty=True)
            self.assertEqual(parsed["functions"], {})
            self.assertEqual(parsed["self_total"], (0, 0, 0, 0, 0))
            with self.assertRaises(CallgrindError):
                parse_callgrind(path, expected_events=EVENTS, allow_empty=False)

    def test_conflicting_compressed_names_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = write_profile(
                Path(temporary),
                "conflict.callgrind",
                """# callgrind format
part: 1
events: Ir Bc Bcm Bi Bim
summary: 1 0 0 0 0
fn=(1) first
0 1
fn=(1) second
0 0
totals: 1 0 0 0 0
""",
            )
            with self.assertRaises(CallgrindError):
                parse_callgrind(path, expected_events=EVENTS)


if __name__ == "__main__":
    unittest.main()
