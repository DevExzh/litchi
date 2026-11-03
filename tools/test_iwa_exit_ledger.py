from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools import iwa_exit_ledger as ledger

LEAF = """\
pub struct Package;
impl Package {
    pub fn edit_table_merges(&self) {}
}
"""

QUEUE = """\
# queue

| Wave | Format | Route | Reason queued | Fixture | Result | Artifact SHA-256 |
|---|---|---|---|---|---|---|
| W0-C | Numbers | `litchi_numbers::Package::edit_table_merges` | none | f | pending | - |
"""


def row(name: str, occurrence: int = 1, status: str = "host-only", **fields: str) -> ledger.Row:
    values = {
        "path": "numbers/editor.rs",
        "name": name,
        "occurrence": occurrence,
        "format": "numbers",
        "family": "table.merge",
        "status": status,
        "destination": "-",
        "parity_test": "-",
        "tier": "-",
        "native_evidence": "-",
        "reason": "-",
    }
    values.update(fields)
    return ledger.Row(**values)


def migrated(name: str, occurrence: int = 1, **fields: str) -> ledger.Row:
    values = {
        "destination": "litchi_numbers::Package::edit_table_merges",
        "parity_test": "numbers_merge_parity",
        "tier": "A",
    }
    values.update(fields)
    return row(name, occurrence, "migrated", **values)


class LedgerTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.write("crates/litchi-numbers/src/lib.rs", LEAF)
        self.write("docs/unified-ops/0363-mac-queue.md", QUEUE)
        self.write("test-data/receipt.json", "{}\n")
        self.rows = [
            migrated("merge_cells", 1),
            row("merge_cells", 2, "retired", reason="duplicate of merge_cells #1"),
            migrated(
                "table_cell_merges",
                destination="litchi_numbers::Package",
                parity_test="numbers_read_parity",
                tier="B",
                native_evidence="test-data/receipt.json",
            ),
        ]

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def write(self, relative: str, text: str) -> None:
        path = self.root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")

    def errors(self) -> list[str]:
        self.write(str(ledger.LEDGER), ledger.render_ledger(self.rows))
        return ledger.check(self.root)

    def assert_error(self, fragment: str) -> None:
        errors = self.errors()
        self.assertTrue(any(fragment in error for error in errors), errors)

    def test_valid_ledger_passes(self) -> None:
        self.assertEqual(self.errors(), [])

    def test_host_only_rows_fail_after_the_deletion(self) -> None:
        self.rows.append(row("gone"))
        self.assert_error("host-only row, but litchi-iwa was deleted")

    def test_status_and_retired_reason(self) -> None:
        self.rows[1] = row("merge_cells", 2, "moved")
        self.assert_error("is not one of")
        self.rows[1] = row("merge_cells", 2, "retired")
        self.assert_error("retired rows need a reason")
        self.rows[1] = row("merge_cells", 2, "retired", reason="duplicate of merge_cells #1")
        self.assertEqual(self.errors(), [])

    def test_migrated_requires_a_parity_test_and_tier(self) -> None:
        self.rows[0] = migrated("merge_cells", parity_test="-")
        self.assert_error("has no parity test")
        self.rows[0] = migrated("merge_cells", tier="C")
        self.assert_error("tier 'C' is not A or B")

    def test_migrated_destination_must_exist(self) -> None:
        self.rows[0] = migrated("merge_cells", destination="-")
        self.assert_error("has no destination")
        self.rows[0] = migrated("merge_cells", destination="litchi_numbers::Package::edit_gone")
        self.assert_error("edit_gone is not a public item of litchi_numbers")
        self.rows[0] = migrated("merge_cells", destination="litchi_missing::Package")
        self.assert_error("crate litchi_missing has no source directory")

    def test_native_evidence_must_exist(self) -> None:
        (self.root / "test-data/receipt.json").unlink()
        self.assert_error("native evidence test-data/receipt.json does not exist")

    def test_unqueued_route_without_native_evidence_fails(self) -> None:
        self.write("docs/unified-ops/0363-mac-queue.md", QUEUE.replace("Numbers", "Pages"))
        self.assert_error("has no host native evidence and no docs/unified-ops/0363-mac-queue.md row")
        report = ledger.queue_report(self.root)
        self.assertIn("[MISSING] Numbers litchi_numbers::Package::edit_table_merges", report)
        self.assertIn("| <wave> | Numbers | `litchi_numbers::Package::edit_table_merges` |", report)

    def test_queue_report_lists_queued_routes(self) -> None:
        self.errors()
        report = ledger.queue_report(self.root)
        self.assertIn("[queued] Numbers litchi_numbers::Package::edit_table_merges", report)
        self.assertNotIn("Rows to append", report)

    def test_cited_mac_queue_inputs_must_be_indexed(self) -> None:
        self.write(
            "docs/unified-ops/0363-mac-queue.md",
            QUEUE.replace(
                "| f |",
                "| `test-data/iwork/mac-queue/inputs/a.numbers`, `.coord/iwa-w3/mac-inputs/b.numbers` |",
            ),
        )
        self.assert_error("index.tsv does not exist")
        self.write(
            "test-data/iwork/mac-queue/index.tsv",
            "test\tcase\tfixture_label\tinput\tlocation\tbytes\tsha256\n"
            "t\tc\tl\ttest-data/iwork/mac-queue/inputs/a.numbers\tarchive\t1\t00\n",
        )
        self.assert_error("inputs.tar.xz does not exist")
        self.write("test-data/iwork/mac-queue/inputs.tar.xz", "x")
        self.assert_error("cited macOS input .coord/iwa-w3/mac-inputs/b.numbers has no external row")
        self.write(
            "test-data/iwork/mac-queue/index.tsv",
            "test\tcase\tfixture_label\tinput\tlocation\tbytes\tsha256\n"
            "t\tc\tl\ttest-data/iwork/mac-queue/inputs/a.numbers\tarchive\t1\t00\n"
            "t\tc\tl\t.coord/iwa-w3/mac-inputs/b.numbers\texternal\t1\t00\n",
        )
        self.assertEqual(self.errors(), [])


class RepositoryLedgerTest(unittest.TestCase):
    def test_tracked_ledger_passes(self) -> None:
        self.assertEqual(ledger.check(), [])


if __name__ == "__main__":
    unittest.main()
