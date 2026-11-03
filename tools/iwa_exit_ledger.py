#!/usr/bin/env python3
"""Checked historical ledger of the `litchi-iwa` exit (records 0363, 0369, 0429).

`litchi-iwa` was deleted by record 0429 (W3). Its public-function inventory
was frozen at commit 27b4840684, the last commit that contains the crate, in
`docs/unified-ops/results/0363/iwa-exit-ledger.tsv`. The ledger is kept as the
historical record of where every host function went; this tool checks it
without the crate and fails when:

- a row has an unknown status, a `host-only` row remains (the exit is
  complete), or a `retired` row has no reason;
- a `migrated` row has no destination, or its destination path
  (`crate::Item::item`) does not resolve to public items of that crate;
- a `migrated` row has no parity test or a tier other than A or B (the
  parity harness was deleted with the host; the test names are history);
- a `migrated` row cites native evidence that does not exist, or has no
  native evidence and no row in `docs/unified-ops/0363-mac-queue.md` for its
  format and destination (owner decision D1);
- a macOS input cited by the queue (`test-data/iwork/mac-queue/inputs/...`,
  packed in `inputs.tar.xz`, or the host-local `.coord/iwa-w3/mac-inputs/...`)
  has no row in `test-data/iwork/mac-queue/index.tsv`, or the archive is
  missing.

Usage:
  python3 tools/iwa_exit_ledger.py --check            # the gate
  python3 tools/iwa_exit_ledger.py --print-mac-queue  # routes that need a macOS run, and missing queue rows
  python3 tools/iwa_exit_ledger.py --summary          # counts per format and status
"""

from __future__ import annotations

import argparse
import collections
import csv
import dataclasses
import io
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LEDGER = Path("docs/unified-ops/results/0363/iwa-exit-ledger.tsv")
MAC_QUEUE = Path("docs/unified-ops/0363-mac-queue.md")
MAC_QUEUE_INPUTS = ("test-data/iwork/mac-queue/inputs/", ".coord/iwa-w3/mac-inputs/")
MAC_QUEUE_INDEX = Path("test-data/iwork/mac-queue/index.tsv")
MAC_QUEUE_ARCHIVE = Path("test-data/iwork/mac-queue/inputs.tar.xz")
HOST_COMMIT = "27b4840684"

COLUMNS = (
    "path", "name", "occurrence", "format", "family", "status",
    "destination", "parity_test", "tier", "native_evidence", "reason",
)
STATUSES = ("host-only", "migrated", "retired")
FORMAT_NAMES = {"numbers": "Numbers", "pages": "Pages", "keynote": "Keynote"}
EMPTY = ("", "-")


@dataclasses.dataclass
class Row:
    path: str
    name: str
    occurrence: int
    format: str
    family: str
    status: str
    destination: str
    parity_test: str
    tier: str
    native_evidence: str
    reason: str

    @property
    def key(self) -> tuple[str, str, int]:
        return (self.path, self.name, self.occurrence)


def _blank(value: str) -> bool:
    return value.strip() in EMPTY


def read_ledger(path: Path) -> list[Row]:
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t", quoting=csv.QUOTE_NONE)
        if tuple(reader.fieldnames or ()) != COLUMNS:
            raise SystemExit(f"{path}: header must be " + "\t".join(COLUMNS))
        rows = []
        for record in reader:
            values = {column: (record.get(column) or "").strip() for column in COLUMNS}
            values["occurrence"] = int(values["occurrence"] or 0)
            rows.append(Row(**values))
        return rows


def render_ledger(rows: list[Row]) -> str:
    out = io.StringIO()
    out.write("\t".join(COLUMNS) + "\n")
    for row in sorted(rows, key=lambda r: r.key):
        values = dataclasses.asdict(row)
        out.write("\t".join(str(values[column]) or "-" for column in COLUMNS) + "\n")
    return out.getvalue()


def mac_queue(root: Path = ROOT) -> list[dict[str, str]]:
    """Rows of the macOS verification queue table."""
    path = root / MAC_QUEUE
    if not path.is_file():
        return []
    rows = []
    header: list[str] | None = None
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.startswith("|"):
            header = None if not line.strip() else header
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if header is None:
            header = [cell.lower() for cell in cells]
            continue
        if all(set(cell) <= set("-: ") for cell in cells):
            continue
        rows.append(dict(zip(header, cells)))
    return rows


def _strip_code(text: str) -> str:
    return text.replace("`", "").strip()


def queued(row: Row, queue: list[dict[str, str]]) -> bool:
    fmt = FORMAT_NAMES.get(row.format, row.format).lower()
    for entry in queue:
        if entry.get("format", "").strip().lower() != fmt:
            continue
        if _strip_code(entry.get("route", "")).split(" ")[0] == row.destination:
            return True
    return False


_ITEM = r"pub\s+(?:const\s+|async\s+|unsafe\s+)*(?:fn|struct|enum|trait|type|mod|const|static)\s+{name}\b"


def destination_exists(destination: str, root: Path = ROOT, cache: dict | None = None) -> str | None:
    """Return None when `crate::A::b` resolves, else the reason it does not."""
    parts = [part for part in destination.split("::") if part]
    if len(parts) < 2:
        return f"destination {destination!r} must be a crate path such as litchi_numbers::Package::edit_x"
    crate_dir = root / "crates" / parts[0].replace("_", "-") / "src"
    if not crate_dir.is_dir():
        return f"crate {parts[0]} has no source directory {crate_dir.relative_to(root)}"
    if cache is None:
        cache = {}
    text = cache.get(crate_dir)
    if text is None:
        text = "\n".join(file.read_text(encoding="utf-8") for file in sorted(crate_dir.rglob("*.rs")))
        cache[crate_dir] = text
    for segment in parts[1:]:
        declared = re.search(_ITEM.format(name=re.escape(segment)), text)
        exported = re.search(r"pub\s+use\s+[^;]*\b" + re.escape(segment) + r"\b[^;]*;", text)
        if not declared and not exported:
            return f"{segment} is not a public item of {parts[0]}"
    return None


def queue_inputs(queue: list[dict[str, str]]) -> list[str]:
    """macOS input paths cited by the queue's fixture column."""
    found: list[str] = []
    pattern = re.compile("(?:" + "|".join(map(re.escape, MAC_QUEUE_INPUTS)) + r")[A-Za-z0-9_.-]+")
    for entry in queue:
        for match in pattern.finditer(entry.get("fixture", "")):
            path = match.group(0).rstrip(".,;:)")
            if path not in found:
                found.append(path)
    return found


def check(root: Path = ROOT, ledger_path: Path | None = None) -> list[str]:
    errors: list[str] = []
    rows = read_ledger(ledger_path or root / LEDGER)
    queue = mac_queue(root)
    seen: set[tuple[str, str, int]] = set()
    for row in rows:
        if row.key in seen:
            errors.append(f"duplicate ledger row {row.path} {row.name} #{row.occurrence}")
        seen.add(row.key)
    cache: dict = {}
    for row in rows:
        where = f"{row.path} {row.name} #{row.occurrence}"
        if row.status not in STATUSES:
            errors.append(f"{where}: status {row.status!r} is not one of {', '.join(STATUSES)}")
            continue
        if row.status == "host-only":
            errors.append(
                f"{where}: host-only row, but litchi-iwa was deleted (record 0429); "
                "the row must be migrated or retired"
            )
            continue
        if row.status == "retired":
            if _blank(row.reason):
                errors.append(f"{where}: retired rows need a reason")
            continue
        if _blank(row.destination):
            errors.append(f"{where}: migrated row has no destination")
        else:
            reason = destination_exists(row.destination, root, cache)
            if reason:
                errors.append(f"{where}: destination {row.destination} does not exist: {reason}")
        if _blank(row.parity_test):
            errors.append(f"{where}: migrated row has no parity test")
        if row.tier not in ("A", "B"):
            errors.append(f"{where}: migrated row tier {row.tier!r} is not A or B")
        evidence = "-" if _blank(row.native_evidence) else row.native_evidence
        if evidence != "-":
            if not (root / evidence).is_file():
                errors.append(f"{where}: native evidence {evidence} does not exist")
        elif not queued(row, queue):
            errors.append(
                f"{where}: migrated tier-{row.tier} route {row.destination} has no host native evidence "
                f"and no {MAC_QUEUE} row (format {FORMAT_NAMES.get(row.format, row.format)})"
            )
    cited = queue_inputs(queue)
    if cited:
        index = root / MAC_QUEUE_INDEX
        listed: dict[str, str] = {}
        if index.is_file():
            with index.open(encoding="utf-8", newline="") as handle:
                for record in csv.DictReader(handle, delimiter="\t", quoting=csv.QUOTE_NONE):
                    listed[record.get("input", "")] = record.get("location", "")
        else:
            errors.append(f"{MAC_QUEUE}: cites macOS inputs but {MAC_QUEUE_INDEX} does not exist")
        if any(path.startswith(MAC_QUEUE_INPUTS[0]) for path in cited) and not (root / MAC_QUEUE_ARCHIVE).is_file():
            errors.append(f"{MAC_QUEUE}: cites archived macOS inputs but {MAC_QUEUE_ARCHIVE} does not exist")
        for path in cited:
            expected = "archive" if path.startswith(MAC_QUEUE_INPUTS[0]) else "external"
            if listed.get(path) != expected:
                errors.append(f"{MAC_QUEUE}: cited macOS input {path} has no {expected} row in {MAC_QUEUE_INDEX}")
    return errors


def queue_report(root: Path = ROOT) -> str:
    rows = read_ledger(root / LEDGER)
    queue = mac_queue(root)
    routes: dict[tuple[str, str], list[Row]] = collections.defaultdict(list)
    for row in rows:
        if row.status == "migrated" and _blank(row.native_evidence):
            routes[(row.format, row.destination)].append(row)
    lines = ["Migrated routes without host native evidence (owner decision D1):"]
    missing = []
    for (fmt, destination), members in sorted(routes.items()):
        state = "queued" if queued(members[0], queue) else "MISSING"
        tests = sorted({member.parity_test for member in members})
        lines.append(
            f"  [{state}] {FORMAT_NAMES.get(fmt, fmt)} {destination} "
            f"(host: {', '.join(sorted({m.name for m in members}))}; tests: {', '.join(tests)})"
        )
        if state == "MISSING":
            missing.append(
                f"| <wave> | {FORMAT_NAMES.get(fmt, fmt)} | `{destination}` | "
                f"no host native evidence; tier {members[0].tier} ({', '.join(tests)}) | <fixture> | pending | - |"
            )
    if not routes:
        lines.append("  none")
    if missing:
        lines.append(f"Rows to append to {MAC_QUEUE}:")
        lines.extend(missing)
    return "\n".join(lines) + "\n"


def summary(root: Path = ROOT) -> str:
    counts: collections.Counter[tuple[str, str]] = collections.Counter()
    for row in read_ledger(root / LEDGER):
        counts[(row.format, row.status)] += 1
    lines = [f"{fmt:8} {status:10} {count}" for (fmt, status), count in sorted(counts.items())]
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--check", action="store_true", help="validate the ledger (the gate)")
    group.add_argument("--print-mac-queue", action="store_true", help="list routes that need a macOS run")
    group.add_argument("--summary", action="store_true", help="counts per format and status")
    args = parser.parse_args(argv)
    if args.print_mac_queue:
        sys.stdout.write(queue_report())
        return 0
    if args.summary:
        sys.stdout.write(summary())
        return 0
    errors = check()
    for error in errors:
        print(f"error: {error}", file=sys.stderr)
    if errors:
        print(f"{len(errors)} ledger error(s)", file=sys.stderr)
        return 1
    print(f"{LEDGER}: ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
