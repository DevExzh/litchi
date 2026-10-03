#!/usr/bin/env python3
"""Read canonical inclusive rows without summing overlapping call trees."""
import argparse
import hashlib
import json
from pathlib import Path
import re

P = Path(__file__).resolve().parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--output", type=Path, required=True)
args = parser.parse_args()
selectors = {
    "runner": "tools/perf-baseline/src/producer_shape.rs:litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan [",
    "mce_process": "crates/litchi-ooxml-common/src/mce/codec.rs:litchi_ooxml_common::mce::codec::process_markup_compatibility [",
    "mce_start": "crates/litchi-ooxml-common/src/mce/codec.rs:litchi_ooxml_common::mce::codec::start [",
    "x14ac_capture": "crates/litchi-xlsx/src/raw/worksheet/x14ac.rs:litchi_xlsx::raw::worksheet::x14ac::capture_inner [",
}
rows = []
for folder in (P / "callgrind-00", P / "callgrind-01"):
    raw = (folder / "callgrind.out").read_text()
    annotation = (folder / "inclusive-canonical.txt").read_text()
    stderr = (folder / "stderr.txt").read_text()
    total = int(re.search(r"^totals: (\d+)", raw, re.M)[1])
    if total <= 0 or total != int(re.search(r"^summary: (\d+)", raw, re.M)[1]):
        raise ValueError("instruction totals disagree")
    pid = int(re.search(r"^pid: (\d+)", raw, re.M)[1])
    if pid != int(re.search(r"^==(\d+)== Callgrind", stderr, re.M)[1]):
        raise ValueError("raw profile is not the original parent process")
    selected = {}
    for name, selector in selectors.items():
        matches = [line for line in annotation.splitlines() if selector in line]
        if len(matches) != 1:
            raise ValueError("canonical function row is missing or ambiguous: " + name)
        value = int(matches[0].split()[0].replace(",", ""))
        selected[name] = {"inclusive_ir": value, "percent_of_runner": 100 * value / total, "selector": selector}
    if selected["runner"]["inclusive_ir"] != total:
        raise ValueError("selected runner does not cover collected instructions")
    rows.append({"child": folder.name, "total_ir": total, "selected_inclusive_rows": selected, "parent_pid": pid,
        "brk_warning": "brk segment overflow" in stderr,
        "callgrind_sha256": hashlib.sha256(raw.encode()).hexdigest(),
        "annotation_sha256": hashlib.sha256(annotation.encode()).hexdigest()})
with args.output.open("x") as out:
    json.dump({"status": "pass", "rows": rows}, out, indent=2)
    out.write("\n")
