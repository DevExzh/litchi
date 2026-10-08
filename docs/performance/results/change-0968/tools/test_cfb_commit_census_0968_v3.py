#!/usr/bin/env python3
"""Read-only v3 schema tests, including the retained v5 raw packets."""
from __future__ import annotations

import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS))
import cfb_census_schema_0968_v3 as schema  # noqa: E402


ROOT = Path("/home/zhuhe/litchi-performance-0968")
SMOKE = ROOT / "smoke-docnohf-e1-zero-v4.json"
SOURCE = ROOT / "source-diagnostic"
RETAINED_V5 = ROOT / "capture/census-v5"
RETAINED_V5_PLAN = ROOT / "plans/plan-0968-prospective-v5.json"
INCIDENT = ROOT / "review/census-v2-incident-v1.json"


def fixture_descriptor(path: Path) -> dict[str, object]:
    return schema.descriptor(path, "fixture")


def smoke_plan(path: Path = SMOKE) -> tuple[dict, dict, dict]:
    report = schema.load_json(path)
    fixtures = {
        case: fixture_descriptor(SOURCE / relative)
        for case, relative in schema.CASE_FILES.items()
    }
    oracle = {}
    report_case = report["case"]
    report_edits = report["edits"]
    for cell in schema.expected_cells():
        if cell["case"] == report_case and cell["edits"] == report_edits:
            route = next(
                route for route in report["route_outputs"]
                if route["retention"] == schema.REPORT_RETENTION[cell["retention"]]
            )
            oracle[cell["id"]] = {"sha256": route["output_sha256"], "bytes": route["output_bytes"]}
        else:
            oracle[cell["id"]] = {"sha256": "0" * 64, "bytes": 0}
    plan = {"fixtures": fixtures}
    return report, plan, oracle


class CensusV3Tests(unittest.TestCase):
    def test_real_v2_smoke_packet_passes(self) -> None:
        report, plan, oracle = smoke_plan()
        schema.check_report(report, schema.expected_rows()[0], plan, oracle)

    def test_real_v4_smoke_packets_pass(self) -> None:
        for path, case in (
            (ROOT / "smoke-docnohf-e1-zero-v4.json", "docnohf"),
            (ROOT / "smoke-docfloat-e1-zero-v4.json", "docfloat"),
        ):
            report, plan, oracle = smoke_plan(path)
            row = next(row for row in schema.expected_rows() if row["case"] == case and row["edits"] == 1 and row["retention"] == "zero")
            schema.check_report(report, row, plan, oracle)

    def test_retained_v5_raw_reports_pass_per_edit_cardinality(self) -> None:
        """The failed v2 parse must be recoverable offline without rerunning work."""
        plan = schema.load_json(RETAINED_V5_PLAN)
        oracle = plan["ordinary_output_oracle_cells"]
        rows = {row["child_id"]: row for row in schema.expected_rows()}
        packets = sorted(RETAINED_V5.glob("*/stdout.raw"))
        self.assertEqual(len(packets), 48)
        seen_edits = set()
        for packet in packets:
            child_id = packet.parent.name
            report = json.loads(packet.read_bytes())
            row = rows[child_id]
            schema.check_report(report, row, plan, oracle)
            census = report["samples"][0]["census"]
            self.assertEqual(census["commit_attempts"], row["edits"])
            self.assertEqual(census["capture_attempted"], row["edits"])
            seen_edits.add(row["edits"])
        self.assertEqual(seen_edits, {1, 2, 8, 32})

    def test_incident_binds_retained_failed_epoch(self) -> None:
        schema.check_incident(schema.load_json(INCIDENT), "incident")

    def test_exact_fixed_matrix_has_two_directions(self) -> None:
        rows = schema.expected_rows()
        self.assertEqual(len(rows), 48)
        self.assertEqual({row["repeat"] for row in rows}, {1, 2})
        self.assertEqual({row["direction"] for row in rows}, {"forward", "reverse"})
        self.assertEqual(len({row["child_id"] for row in rows}), 48)

    def test_extra_elapsed_field_rejected(self) -> None:
        report, plan, oracle = smoke_plan()
        report["elapsed_ns"] = 1
        with self.assertRaises(schema.CensusError):
            schema.check_report(report, schema.expected_rows()[0], plan, oracle)

    def test_missing_census_field_rejected(self) -> None:
        report, plan, oracle = smoke_plan()
        del report["samples"][0]["census"]["layout_error"]
        with self.assertRaises(schema.CensusError):
            schema.check_report(report, schema.expected_rows()[0], plan, oracle)

    def test_repeat_change_is_detectable(self) -> None:
        report, _plan, _oracle = smoke_plan()
        second = copy.deepcopy(report)
        second["samples"][0]["census"]["capture_payload_bytes"] += 1
        self.assertNotEqual(schema.normalized_report(report), schema.normalized_report(second))

    def test_output_oracle_mismatch_rejected(self) -> None:
        report, plan, oracle = smoke_plan()
        oracle = copy.deepcopy(oracle)
        oracle["docnohf/e1/zero"]["bytes"] += 1
        with self.assertRaises(schema.CensusError):
            schema.check_report(report, schema.expected_rows()[0], plan, oracle)

    def test_error_counter_rejected(self) -> None:
        report, plan, oracle = smoke_plan()
        report["samples"][0]["census"]["layout_error"] = 1
        with self.assertRaises(schema.CensusError):
            schema.check_report(report, schema.expected_rows()[0], plan, oracle)

    def test_duplicate_json_keys_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"schema": "v1", "schema": "v2"}\n', encoding="utf-8")
            with self.assertRaises(schema.CensusError):
                schema.load_json(path)


if __name__ == "__main__":
    unittest.main()
