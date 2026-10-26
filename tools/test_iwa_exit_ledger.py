from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools import iwa_exit_ledger as ledger

HOST = """\
impl Draft {
    pub fn merge_cells(&mut self) {}
    pub fn table_cell_merges(&self) {}
    pub(crate) fn hidden(&self) {}
}
mod nested {
    pub fn merge_cells() {}
}
"""

LEAF = """\
pub struct Package;
impl Package {
    pub fn edit_table_merges(&self) {}
}
"""

PARITY = """\
#[test]
fn numbers_merge_parity() -> TestResult {
    Case::new("numbers_merge_parity", Format::Numbers, Tier::A, Evidence::Queued, &input)
        .run()?;
    Ok(())
}

#[test]
fn numbers_read_parity() -> TestResult {
    Case::new(
        "numbers_read_parity",
        Format::Numbers,
        Tier::B,
        Evidence::Native("test-data/receipt.json"),
        &input,
    )
    .run()?;
    Ok(())
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
        self.write("crates/litchi-iwa/src/numbers/editor.rs", HOST)
        self.write("crates/litchi-iwa/src/private.rs", "pub fn not_scanned() {}\n")
        self.write("crates/litchi-numbers/src/lib.rs", LEAF)
        self.write("crates/litchi-iwa/tests/parity/numbers_cases.rs", PARITY)
        self.write("docs/unified-ops/0363-mac-queue.md", QUEUE)
        self.write("test-data/receipt.json", "{}\n")
        self.rows = [
            migrated("merge_cells", 1),
            row("merge_cells", 2),
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

    def test_scan_counts_occurrences_and_skips_unscanned_modules(self) -> None:
        found = ledger.scan_inventory(self.root)
        self.assertEqual(
            [(fn.name, fn.occurrence, fn.line) for fn in found],
            [("merge_cells", 1, 2), ("table_cell_merges", 1, 3), ("merge_cells", 2, 7)],
        )

    def test_valid_ledger_passes(self) -> None:
        self.assertEqual(self.errors(), [])

    def test_unmapped_host_function_fails(self) -> None:
        del self.rows[1]
        self.assert_error("unmapped host function numbers/editor.rs:7 merge_cells")

    def test_stale_rows(self) -> None:
        self.rows.append(row("gone"))
        self.assert_error("stale host-only row")
        self.rows[-1] = migrated("gone")
        self.assertEqual(self.errors(), [])

    def test_status_and_retired_reason(self) -> None:
        self.rows[1] = row("merge_cells", 2, "moved")
        self.assert_error("is not one of")
        self.rows[1] = row("merge_cells", 2, "retired")
        self.assert_error("retired rows need a reason")
        self.rows[1] = row("merge_cells", 2, "retired", reason="duplicate of merge_cells #1")
        self.assertEqual(self.errors(), [])

    def test_migrated_requires_an_existing_parity_test(self) -> None:
        self.rows[0] = migrated("merge_cells", parity_test="-")
        self.assert_error("has no parity test")
        self.rows[0] = migrated("merge_cells", parity_test="no_such_test")
        self.assert_error("parity test no_such_test not found")

    def test_migrated_destination_must_exist(self) -> None:
        self.rows[0] = migrated("merge_cells", destination="-")
        self.assert_error("has no destination")
        self.rows[0] = migrated("merge_cells", destination="litchi_numbers::Package::edit_gone")
        self.assert_error("edit_gone is not a public item of litchi_numbers")
        self.rows[0] = migrated("merge_cells", destination="litchi_missing::Package")
        self.assert_error("crate litchi_missing has no source directory")

    def test_tier_and_evidence_must_match_the_test(self) -> None:
        self.rows[0] = migrated("merge_cells", tier="B")
        self.assert_error("ledger tier 'B' but numbers_merge_parity declares ['A']")
        self.rows[0] = migrated("merge_cells", native_evidence="test-data/receipt.json")
        self.assert_error("declares ['-']")

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

    def test_update_adds_new_and_drops_stale_host_only_rows(self) -> None:
        self.rows = [migrated("merge_cells", 1), row("gone"), migrated("removed_after_migration")]
        self.write(str(ledger.LEDGER), ledger.render_ledger(self.rows))
        self.assertEqual(ledger.update(self.root), (2, 1))
        names = {(r.name, r.occurrence, r.status) for r in ledger.read_ledger(self.root / ledger.LEDGER)}
        self.assertEqual(
            names,
            {
                ("merge_cells", 1, "migrated"),
                ("merge_cells", 2, "host-only"),
                ("table_cell_merges", 1, "host-only"),
                ("removed_after_migration", 1, "migrated"),
            },
        )


class RepositoryLedgerTest(unittest.TestCase):
    def test_tracked_ledger_passes(self) -> None:
        self.assertEqual(ledger.check(), [])


if __name__ == "__main__":
    unittest.main()
