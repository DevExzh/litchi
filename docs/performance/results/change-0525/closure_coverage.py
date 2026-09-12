"""Translate the frozen Rust corpus selectors to assess reconstruction closure size.

This is a static count, not a runtime performance estimate. The source digest
binds the translation to the generator inspected during candidate selection.
"""
import hashlib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SOURCE = REPO / "tools/perf-baseline/src/lib.rs"


def analyze():
    rows = []
    for shape in ("medium", "dense-sparse"):
        inventory = []
        for sheet in range(4):
            if shape == "medium":
                cells = [(sheet, row, col) for row in range(48) for col in range(48)]
            elif sheet == 0:
                cells = [(sheet, row, col) for row in range(128) for col in range(128)]
            elif sheet in (1, 2):
                step = 4 if sheet == 1 else 8
                cells = [(sheet, row, col) for row in range(0, 128, step)
                         for col in range(0, 128, step)]
            else:
                cells = [(sheet, index, index) for index in range(128)]
            inventory.extend(cells)
        count = (len(inventory) + 99) // 100
        updates = [inventory[index * len(inventory) // count] for index in range(count)]
        touched_rows = {(sheet, row) for sheet, row, _ in updates}
        touched_cells = sum((sheet, row) in touched_rows for sheet, row, _ in inventory)
        prior = json.loads((HERE.parent / "change-0522" / "baseline" /
                            f"native-r1-primary-{shape}.json").read_text())["results"][0]["corpus"]
        assert len(inventory) == prior["entry_count"]
        assert count == prior["xlsx"]["one_percent_update_count"]
        rows.append(dict(
            shape=shape, cells=len(inventory), updates=count,
            touched_rows=len(touched_rows), cells_in_touched_rows=touched_cells,
            row_only_skipped_cells=len(inventory) - touched_cells,
            row_only_skipped_fraction=(len(inventory) - touched_cells) / len(inventory),
            cell_span_skipped_cells=len(inventory) - count,
            cell_span_skipped_fraction=(len(inventory) - count) / len(inventory),
            updates_by_sheet=[sum(sheet == index for sheet, _, _ in updates) for index in range(4)],
        ))
    return {
        "schema": "litchi-0525-source-derived-closure-v1",
        "source": "tools/perf-baseline/src/lib.rs",
        "source_sha256": hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
        "owners": ["xlsx_cell_crud_inventory", "xlsx_cell_crud_updates"],
        "scope": "Exact deterministic generator translation for static mechanism selection; counts are not instruction, allocation, or latency estimates. Current source is bound to retained0522 baseline.",
        "results": rows,
    }


if __name__ == "__main__":
    expected = json.loads((HERE / "closure-coverage.json").read_text())
    assert analyze() == expected, "source-derived closure evidence changed"
    print("Closure coverage replay passed for both primary shapes")
