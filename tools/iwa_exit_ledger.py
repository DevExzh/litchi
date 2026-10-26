#!/usr/bin/env python3
"""Checked ledger of the `litchi-iwa` exit (records 0363 and 0369).

The tool regenerates the host's public-function inventory from
`crates/litchi-iwa/src` (the scan of record 0363 section 1.4), joins it with
the tracked ledger `docs/unified-ops/results/0363/iwa-exit-ledger.tsv`, and
fails when:

- a host function has no ledger row, or a `host-only` row names a function
  that no longer exists (stale);
- a row has an unknown status, or a `retired` row has no reason;
- a `migrated` row has no destination, or its destination path
  (`crate::Item::item`) does not resolve to public items of that crate;
- a `migrated` row has no parity test, the test does not exist under
  `crates/litchi-iwa/tests/parity/`, or the tier / native evidence the row
  records differ from what the test declares (`Tier::A|B`,
  `Evidence::Native("path")` or `Evidence::Queued` in the test body);
- a `migrated` row cites native evidence that does not exist, or has no
  native evidence and no row in `docs/unified-ops/0363-mac-queue.md` for its
  format and destination (owner decision D1).

Usage:
  python3 tools/iwa_exit_ledger.py --check            # the gate
  python3 tools/iwa_exit_ledger.py --update           # add new host fns as host-only, drop stale host-only rows
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
HOST_SRC = Path("crates/litchi-iwa/src")
PARITY_DIR = Path("crates/litchi-iwa/tests/parity")
LEDGER = Path("docs/unified-ops/results/0363/iwa-exit-ledger.tsv")
MAC_QUEUE = Path("docs/unified-ops/0363-mac-queue.md")

# The module scan of record 0363 section 1.4, in that order.
SCANNED_MODULES = (
    "numbers", "keynote", "pages", "charts", "text", "shapes", "media",
    "comments.rs", "comments", "document.rs", "drawable_order.rs",
    "application.rs", "lib.rs", "package.rs",
)
PUB_FN = re.compile(r"^\s*pub fn ([a-z0-9_]+)")

COLUMNS = (
    "path", "name", "occurrence", "format", "family", "status",
    "destination", "parity_test", "tier", "native_evidence", "reason",
)
STATUSES = ("host-only", "migrated", "retired")
FORMAT_NAMES = {"numbers": "Numbers", "pages": "Pages", "keynote": "Keynote"}
EMPTY = ("", "-")


@dataclasses.dataclass(frozen=True)
class HostFn:
    path: str
    name: str
    occurrence: int
    line: int

    @property
    def key(self) -> tuple[str, str, int]:
        return (self.path, self.name, self.occurrence)


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


@dataclasses.dataclass(frozen=True)
class ParityTest:
    name: str
    file: str
    tiers: frozenset[str]
    evidence: frozenset[str]  # "-" for Evidence::Queued, else the cited path


def _blank(value: str) -> bool:
    return value.strip() in EMPTY


def format_of(path: str) -> str:
    head = path.split("/")[0].split(".")[0]
    return head if head in FORMAT_NAMES else "shared"


def scan_inventory(root: Path = ROOT) -> list[HostFn]:
    """Every `pub fn` of the scanned host modules, with per-file occurrence."""
    src = root / HOST_SRC
    found: list[HostFn] = []
    seen: set[Path] = set()
    for module in SCANNED_MODULES:
        target = src / module
        if target.is_file():
            files = [target]
        elif target.is_dir():
            files = sorted(target.rglob("*.rs"))
        else:
            continue
        for file in files:
            if file in seen:
                continue
            seen.add(file)
            relative = file.relative_to(src).as_posix()
            counts: collections.Counter[str] = collections.Counter()
            for number, text in enumerate(file.read_text(encoding="utf-8").splitlines(), 1):
                match = PUB_FN.match(text)
                if match:
                    counts[match.group(1)] += 1
                    found.append(HostFn(relative, match.group(1), counts[match.group(1)], number))
    return found


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


def parity_tests(root: Path = ROOT) -> dict[str, ParityTest]:
    """`#[test]` functions of the parity harness with their declarations."""
    tests: dict[str, ParityTest] = {}
    directory = root / PARITY_DIR
    if not directory.is_dir():
        return tests
    start = re.compile(r"#\[test\]\s*(?:#\[[^\]]*\]\s*)*fn\s+([a-z0-9_]+)\s*\(")
    for file in sorted(directory.rglob("*.rs")):
        text = file.read_text(encoding="utf-8")
        for match in start.finditer(text):
            end = text.find("\n}\n", match.end())
            body = text[match.end(): end if end >= 0 else len(text)]
            evidence = set(re.findall(r'Evidence::Native\(\s*"([^"]+)"\s*\)', body))
            if "Evidence::Queued" in body:
                evidence.add("-")
            tests[match.group(1)] = ParityTest(
                match.group(1),
                file.relative_to(root).as_posix(),
                frozenset(re.findall(r"Tier::([AB])\b", body)),
                frozenset(evidence),
            )
    return tests


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


def check(root: Path = ROOT, ledger_path: Path | None = None) -> list[str]:
    errors: list[str] = []
    inventory = scan_inventory(root)
    rows = read_ledger(ledger_path or root / LEDGER)
    tests = parity_tests(root)
    queue = mac_queue(root)
    by_key: dict[tuple[str, str, int], Row] = {}
    for row in rows:
        if row.key in by_key:
            errors.append(f"duplicate ledger row {row.path} {row.name} #{row.occurrence}")
        by_key[row.key] = row
    present = {fn.key: fn for fn in inventory}
    for fn in inventory:
        if fn.key not in by_key:
            errors.append(
                f"unmapped host function {fn.path}:{fn.line} {fn.name} (occurrence {fn.occurrence}); "
                "run --update and set its status"
            )
    cache: dict = {}
    for row in rows:
        where = f"{row.path} {row.name} #{row.occurrence}"
        if row.status not in STATUSES:
            errors.append(f"{where}: status {row.status!r} is not one of {', '.join(STATUSES)}")
            continue
        if row.key not in present and row.status == "host-only":
            errors.append(f"{where}: stale host-only row (the host function no longer exists)")
        if row.status == "retired" and _blank(row.reason):
            errors.append(f"{where}: retired rows need a reason")
        if row.status != "migrated":
            continue
        if _blank(row.destination):
            errors.append(f"{where}: migrated row has no destination")
        else:
            reason = destination_exists(row.destination, root, cache)
            if reason:
                errors.append(f"{where}: destination {row.destination} does not exist: {reason}")
        if _blank(row.parity_test):
            errors.append(f"{where}: migrated row has no parity test")
            continue
        test = tests.get(row.parity_test)
        if test is None:
            errors.append(f"{where}: parity test {row.parity_test} not found under {PARITY_DIR}")
            continue
        if test.tiers != {row.tier}:
            errors.append(
                f"{where}: ledger tier {row.tier!r} but {row.parity_test} declares {sorted(test.tiers) or 'none'}"
            )
        evidence = "-" if _blank(row.native_evidence) else row.native_evidence
        if test.evidence != {evidence}:
            errors.append(
                f"{where}: ledger native evidence {evidence!r} but {row.parity_test} declares "
                f"{sorted(test.evidence) or 'none'}"
            )
        if evidence != "-":
            if not (root / evidence).is_file():
                errors.append(f"{where}: native evidence {evidence} does not exist")
        elif not queued(row, queue):
            errors.append(
                f"{where}: migrated tier-{row.tier} route {row.destination} has no host native evidence "
                f"and no {MAC_QUEUE} row (format {FORMAT_NAMES.get(row.format, row.format)})"
            )
    return errors


def update(root: Path = ROOT) -> tuple[int, int]:
    path = root / LEDGER
    rows = read_ledger(path)
    present = {fn.key: fn for fn in scan_inventory(root)}
    kept = [row for row in rows if row.key in present or row.status != "host-only"]
    known = {row.key for row in kept}
    added = 0
    for key, fn in present.items():
        if key not in known:
            kept.append(Row(fn.path, fn.name, fn.occurrence, format_of(fn.path), "-", "host-only",
                            "-", "-", "-", "-", "-"))
            added += 1
    path.write_text(render_ledger(kept), encoding="utf-8")
    return added, len(rows) - (len(kept) - added)


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
    group.add_argument("--update", action="store_true", help="add new host functions, drop stale host-only rows")
    group.add_argument("--print-mac-queue", action="store_true", help="list routes that need a macOS run")
    group.add_argument("--summary", action="store_true", help="counts per format and status")
    args = parser.parse_args(argv)
    if args.update:
        added, dropped = update()
        print(f"{LEDGER}: added {added} host-only rows, dropped {dropped} stale host-only rows")
        return 0
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
