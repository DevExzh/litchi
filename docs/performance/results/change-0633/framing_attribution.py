#!/usr/bin/env python3
"""The framing attribution the frozen design of change 0633 cites.

Reads the retained before-leg `callgrind_annotate --inclusive=yes
--separate-callers=2` pair for one scenario, differences it, divides by the
extra operations, and prints the rows that price the two framing passes of
`Snapshot::from_bytes` against the semantic work each pass does around them.

Usage: framing_attribution.py [<packet dir>] [<leg>] [<stem>] [<operation>]
"""
import re
import sys
from pathlib import Path

here = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent)
leg = sys.argv[2] if len(sys.argv) > 2 else "before"
stem = sys.argv[3] if len(sys.argv) > 3 else "54016"
operation = sys.argv[4] if len(sys.argv) > 4 else "open"

# callgrind_annotate right-aligns the percentage, so a row under 10% prints
# "( 4.83%)" with a leading space. Change 0633 widened this pattern; the
# 0620 copy matched only rows at or above 10% and silently dropped the rest.
ANNOTATION = re.compile(r"^\s*([\d,]+) \(\s*([\d.]+)%\)\s+(.*)$")


def read(path):
    """{function-with-caller-chain: inclusive Ir}, read exactly as analyze.py does."""
    totals = {}
    for line in path.read_text(errors="replace").splitlines():
        if line.startswith(" ") and "=>" in line:
            continue
        match = ANNOTATION.match(line)
        if not match:
            continue
        name = match.group(3).strip()
        if name.startswith("=>"):
            continue
        name = re.sub(r"\s+\[[^\]]*\]$", "", name)
        if ":" in name:
            name = name.split(":", 1)[1]
        value = int(match.group(1).replace(",", ""))
        totals[name] = max(totals.get(name, 0), value)
    return totals


pairs = {}
for line in (here / "callgrind" / f"pairs-{leg}.txt").read_text().splitlines():
    fields = line.split()
    if len(fields) == 4:
        pairs[(fields[0], fields[1])] = (int(fields[2]), int(fields[3]))

small_n, large_n = pairs[(stem, operation)]
small = read(here / "callgrind" / f"ann-{leg}-{stem}-{operation}-small.txt")
large = read(here / "callgrind" / f"ann-{leg}-{stem}-{operation}-large.txt")
per = {k: (v - small.get(k, 0)) / (large_n - small_n) for k, v in large.items()}

# (label, substring that identifies the function-and-caller-chain row)
ROWS = [
    ("whole operation", "PROGRAM TOTALS"),
    ("Snapshot::from_bytes", "Snapshot>::from_bytes'xls_edit_probe::main"),
    ("  Editor::open (CFB capture)", "editor::Editor>::open'<litchi_xls::cell_values::Snapshot>::from_bytes"),
    ("  constructor", "from_package_editor'<litchi_xls::cell_values::Snapshot>::from_bytes"),
    ("    PASS 1 parse_worksheet (inventory)", "cell_values::parse_worksheet'<litchi_xls::cell_values::Snapshot>::from_package_editor"),
    ("      framing: Records::next", "Iterator>::next'litchi_xls::cell_values::parse_worksheet"),
    ("      push_entry", "cell_values::push_entry'litchi_xls::cell_values::parse_worksheet"),
    ("      parse_reference", "cell_values::parse_reference'litchi_xls::cell_values::parse_worksheet"),
    ("    resolve_shared_strings", "cell_values::resolve_shared_strings'<litchi_xls::cell_values::Snapshot>::from_package_editor"),
    ("    PASS 2 Workbook::new", "Workbook<core::io::cursor::Cursor<&[u8]>>>::new'<litchi_xls::cell_values::Snapshot>::from_package_editor"),
    ("      OleFile::open_stream (2nd Workbook extraction)", "OleFile<core::io::cursor::Cursor<&[u8]>>>::open_stream'<litchi_xls::workbook::model::Workbook"),
    ("      SharedStringTable::parse_from_records", "SharedStringTable>::parse_from_records'<litchi_xls::workbook::model::Workbook"),
    ("      parse_worksheet_records_with_compatibility", "parse_worksheet_records_with_compatibility'<litchi_xls::workbook::model::Workbook"),
    ("        framing: Records::next", "Iterator>::next'<litchi_xls::workbook::model::Workbook<core::io::cursor::Cursor<alloc::vec::Vec<u8>>>>::parse_worksheet_records_with_compatibility"),
    ("        add_cell", "semantic::worksheet::add_cell'<litchi_xls::workbook::model::Workbook"),
    ("          BTreeMap::insert", "litchi_xls::cell::Cell>>::insert'litchi_xls::workbook::codec::semantic::worksheet::add_cell"),
    ("    drop_glue::<Workbook>", "drop_glue::<litchi_xls::workbook::model::Workbook<core::io::cursor::Cursor<&[u8]>>>'<litchi_xls::cell_values::Snapshot>::from_package_editor"),
]

total = max((v for k, v in per.items() if "PROGRAM TOTALS" in k), default=1)
print(f"# change 0633 framing attribution: {leg} leg, {stem}, --operation {operation}")
print(f"# isolation pair N={small_n} and N={large_n}, differenced and divided by {large_n - small_n}")
print(f"# callgrind --inclusive=yes --separate-callers=2; Ir per operation\n")
print(f"{'phase':52s} {'Ir/op':>14s} {'share':>8s}")
for label, needle in ROWS:
    hits = [v for k, v in per.items() if needle in k]
    if not hits:
        print(f"{label:52s} {'(absent)':>14s}")
        continue
    value = max(hits)
    print(f"{label:52s} {value:14,.0f} {value / total * 100:7.2f}%")
