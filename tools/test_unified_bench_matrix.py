"""Unit tests for tools/unified_bench_matrix.py (records 0337 and 0344)."""

from __future__ import annotations

import csv
import io
import json
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path
from unittest.mock import patch

from tools import unified_bench_matrix as matrix


def row(record: str, status: str = "meets", explanation: str = "", **extra) -> dict:
    r = matrix.row(record, "xlsx", "set", "scenario", "", 1.0, 1.2, [1.2], matrix.SINGLE,
                   f"series:{record}", "parsed", explanation=explanation)
    r["status"] = status
    r.update(extra)
    return r


class EvidenceCoverage(unittest.TestCase):
    def evidence(self, names: list[str], empty_dirs: tuple[str, ...] = ()) -> dict:
        with tempfile.TemporaryDirectory() as tmp:
            res = Path(tmp)
            for name in names:
                path = res / name
                if "." in name:
                    path.write_text("{}", encoding="utf-8")
                else:
                    path.mkdir()
                    (path / "run-1.json").write_text("{}", encoding="utf-8")
            for name in empty_dirs:
                (res / name).mkdir()
            return matrix.evidence_records(res)

    def test_entries_belong_to_the_record_their_name_starts_with(self):
        found = self.evidence(["0328", "0255-python-overhead.json", "0251a-docx-element-writes",
                               "0209-xls-batch.json", "notes"], empty_dirs=("0343",))
        # Records before 0249 are outside the matrix; an empty directory is no evidence.
        self.assertEqual(found, {"0251a": ["0251a-docx-element-writes"],
                                 "0255": ["0255-python-overhead.json"], "0328": ["0328"]})

    def test_evidence_without_rows_fails(self):
        problems = matrix.coverage_problems([row("0328")], {}, {"0328": ["0328"], "0342": ["0342"]})
        self.assertEqual(len(problems), 1)
        self.assertIn("record 0342", problems[0])
        self.assertIn("contributes no rows", problems[0])

    def test_follow_up_suffixes_cannot_hide_unregistered_evidence(self):
        found = self.evidence(["0360ab", "0361fb-receipt.json", "0362lab-native"])
        self.assertEqual(found, {
            "0360ab": ["0360ab"],
            "0361fb": ["0361fb-receipt.json"],
            "0362lab": ["0362lab-native"],
        })
        problems = matrix.coverage_problems([], {}, found)
        self.assertEqual(len(problems), 3)
        self.assertTrue(all("contributes no rows" in problem for problem in problems))

    def test_numbered_follow_up_ids_keep_the_full_id_and_filename_boundary(self):
        found = self.evidence(["0362ul3", "0362ul10f-receipt.json", "0362ul48",
                               "0362ul3-extra.json", "0362ul3_extra.json", "10362ul3.json",
                               "10362.json", "03621", "03621-receipt.json"])
        self.assertEqual(found, {
            "0362ul3": ["0362ul3", "0362ul3-extra.json"],
            "0362ul10f": ["0362ul10f-receipt.json"],
            "0362ul48": ["0362ul48"],
        })
        self.assertEqual(matrix.coverage_problems(
            [], {"evidence_without_rows": [{"record": "0362ul48", "reason": "tooling-only"}]},
            {"0362ul48": ["0362ul48"]}), [])

    def test_an_exemption_needs_a_reason_and_must_not_go_stale(self):
        evidence = {"0339": ["0339"], "0328": ["0328"]}
        ok = {"evidence_without_rows": [{"record": "0339", "reason": "a probe, not a measurement"}]}
        self.assertEqual(matrix.coverage_problems([row("0328")], ok, evidence), [])
        no_reason = {"evidence_without_rows": [{"record": "0339", "reason": " "}]}
        self.assertIn("has no reason", matrix.coverage_problems([row("0328")], no_reason, evidence)[0])
        has_rows = {"evidence_without_rows": [{"record": "0328", "reason": "x"}]}
        problems = matrix.coverage_problems([row("0328")], has_rows, evidence)
        self.assertTrue(any("contributes rows" in p for p in problems))
        gone = {"evidence_without_rows": [{"record": "0400", "reason": "x"}]}
        problems = matrix.coverage_problems([row("0328")], gone, evidence)
        self.assertTrue(any("has no evidence" in p for p in problems))


class MissDispositions(unittest.TestCase):
    def test_explicit_and_worded_dispositions(self):
        self.assertEqual(matrix.disposition(row("0338", "miss", "x", disposition="residual")), "residual")
        self.assertEqual(matrix.disposition(row("0338", "miss", "Follow-up 0338-P1.")), "follow-up")
        self.assertEqual(matrix.disposition(row("0335", "miss", "cold handle; F6")), "follow-up")
        self.assertEqual(matrix.disposition(row("0338", "miss", "a fixed per-call cost")), "")
        with self.assertRaises(SystemExit):
            matrix.disposition(row("0338", "miss", "x", disposition="later"))


class IWorkChartTextEvidence(unittest.TestCase):
    def parse(self, routes):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "0375b" / "integrated-callgrind.json"
            path.parent.mkdir()
            path.write_text(json.dumps({
                "source_commit": "4aa57b3c19c24ce625f31cdd5489ff0a1d39bc68",
                "routes": routes,
            }))
            with patch.object(matrix, "RES", root):
                return matrix.p0375b()

    def test_instruction_counts_keep_the_host_baseline_and_historical_scope(self):
        rows = self.parse([{
            "test": "numbers_chart_title_read_cost_abba",
            "host_ir": 100, "leaf_ir": 95,
        }])
        self.assertEqual(len(rows), 1)
        item = rows[0]
        self.assertEqual((item["leaf"], item["unified"], item["ratio"]), (100, 95, .95))
        self.assertEqual(item["unit"], "Ir")
        self.assertEqual(item["comparison"], "leaf/host (Callgrind Ir)")
        self.assertEqual(item["budget_value"], 1.05)
        self.assertEqual(item["stage"], "integrated 4aa57b3c19")
        self.assertIn("current-trunk", item["explanation"])

    def test_invalid_or_duplicate_instruction_routes_are_rejected(self):
        valid = {"test": "pages_chart_title_set_cost_abba", "host_ir": 100, "leaf_ir": 95}
        for routes in ([valid, valid], [{**valid, "host_ir": 0}],
                       [{**valid, "leaf_ir": -1}], [{**valid, "test": "unknown_cost_abba"}]):
            with self.subTest(routes=routes), self.assertRaises(SystemExit):
                self.parse(routes)


class Committed(unittest.TestCase):
    def test_committed_outputs_are_current_and_complete(self):
        """`--check`: outputs regenerate byte for byte, every open miss has a
        disposition, every record's evidence contributes rows, and every
        output stays below 1 MiB."""
        err = io.StringIO()
        with redirect_stderr(err):
            rc = matrix.main(["--check"])
        self.assertEqual(rc, 0, err.getvalue())

    def test_row_tables_are_per_format(self):
        """Record 0348: one TSV per format, each holding that format's rows
        under the shared header; no single row table."""
        rows, conflicts, agreements, cur, _ = matrix.build()
        files = matrix.outputs(rows, conflicts, agreements, cur)
        tables = {p.stem: t for p, t in files.items() if p.suffix == ".tsv"}
        self.assertNotIn(matrix.OLD_TSV, files)
        self.assertEqual(set(tables), {r["format"] for r in rows})
        count = 0
        for f, text in tables.items():
            table = list(csv.reader(io.StringIO(text), delimiter="\t"))
            self.assertEqual(table[0], matrix.TSV_COLUMNS, f)
            self.assertTrue(all(line[0] == f for line in table[1:]), f)
            count += len(table) - 1
        self.assertEqual(count, len(rows))

    def test_audit_records_contribute_rows(self):
        rows, _, _, _, _ = matrix.build()
        records = {r["record"] for r in rows}
        for rec in ("0307", "0308", "0314", "0316", "0317", "0318", "0323", "0324", "0325",
                    "0327", "0328", "0329", "0330", "0333", "0335", "0336", "0338", "0342"):
            self.assertIn(rec, records)

    def test_assembled_performance_receipts_contribute_diagnostic_rows(self):
        rows, _, _, _, _ = matrix.build()
        records = {r["record"] for r in rows}
        measured = ("0359qb", "0359qc", "0359qd", "0359r", "0359ra", "0359rb",
                    "0359rc", "0359rd", "0359re", "0359qe", "0362ya", "0362yb", "0362yc",
                    "0362yd", "0362ye")
        for rec in measured:
            self.assertIn(rec, records)
        measured_rows = [r for r in rows if r["record"] in measured]
        self.assertTrue(all(r["source"].startswith(f"parsed:{r['record']}/")
                            and r["stage"] for r in measured_rows))
        diagnostic = [r for r in measured_rows if r["unit"] in
                      ("bytes", "Ir", "count", "instructions")]
        self.assertTrue(diagnostic)
        self.assertTrue(all(r["status"] == "no-budget" for r in diagnostic))
        native = [r for r in measured_rows if r["record"] in ("0359qb", "0359qe", "0362ye")
                  and r["status"] == "unqualified"]
        self.assertGreaterEqual(len(native), 2)
        qe = [r for r in measured_rows if r["record"] == "0359qe"]
        self.assertEqual(len(qe), 12)
        self.assertEqual({r["comparison"] for r in qe}, {"candidate/baseline"})
        self.assertTrue(all("quiet-host" in r["explanation"] or r["status"] == "no-budget"
                            for r in qe))
        final_only = [r for r in measured_rows
                      if r["record"] == "0359re" and r["stage"] == "Heaptrack final-only"]
        self.assertEqual({r["series"].rsplit(":", 1)[-1] for r in final_only},
                         {"temporary", "peak_heap", "leaked"})
        self.assertTrue(all(r["leaf"] is None and r["ratio"] is None
                            and r["unified"] is not None for r in final_only))

    def test_summary_total_matches_its_eight_counter_header(self):
        rows, conflicts, agreements, cur, _ = matrix.build()
        text = matrix.outputs(rows, conflicts, agreements, cur)[matrix.OUT_MD]
        lines = text.splitlines()
        header = ("| format | rows | parsed | curated | meets | miss (open) | "
                  "miss (superseded) | unqualified | no-budget |")
        header_index = lines.index(header)
        total = next(line for line in lines[header_index + 2:] if line.startswith("| **all** |"))
        cells = [cell.strip() for cell in total.strip("|").split("|")]
        self.assertEqual(len(cells), 9)
        expected = [len(rows),
                    sum(r["source"].startswith("parsed") for r in rows),
                    sum(r["source"].startswith("curated") for r in rows),
                    sum(r["status"] == "meets" for r in rows),
                    sum(r["status"] == "miss" for r in rows),
                    sum(r["status"].startswith("miss (") for r in rows),
                    sum(r["status"] == "unqualified" for r in rows),
                    sum(r["status"] == "no-budget" for r in rows)]
        self.assertEqual([int(cell.strip("*")) for cell in cells[1:]], expected)

    def test_the_two_audit_misses_have_follow_ups(self):
        rows, _, _, _, _ = matrix.build()
        wanted = {("0327", "sheet_features:xls"), ("0327", "sheet_features:ods"),
                  ("0338", "p0_residue:xlsx:set_run_bold")}
        hits = [r for r in rows if (r["record"], r["series"]) in wanted]
        self.assertEqual(len(hits), 3)
        for r in hits:
            self.assertEqual(r["status"], "miss")
            self.assertEqual(matrix.disposition(r), "follow-up")


if __name__ == "__main__":
    unittest.main()
