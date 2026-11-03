"""The macOS verification queue and its mapping to Litchi output packages.

Shared by `generate.py`, `apply_results.py` and their unit tests (record
0430). The queue is `docs/unified-ops/0363-mac-queue.md`; every row has to be
mapped to the parity cases (or `litchi_ops::File` routes) whose output
packages the native check opens, or to an explicit grouping that says why no
generated package exists.
"""

from __future__ import annotations

import dataclasses
import hashlib
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
MAC_QUEUE = Path("docs/unified-ops/0363-mac-queue.md")
PARITY_DIR = Path("crates/litchi-iwa/tests/parity")
RECORDS_DIR = Path("docs/unified-ops")
LEDGER = Path("docs/unified-ops/results/0363/iwa-exit-ledger.tsv")

# Columns of the queue table, lower-cased.
COLUMNS = ("wave", "format", "route", "reason queued", "fixture", "result", "artifact sha-256")
# Columns that identify a row (the two result columns are filled by the run).
IDENTITY = COLUMNS[:5]

APPS = {"keynote": "Keynote", "numbers": "Numbers", "pages": "Pages"}
EXTENSIONS = {"keynote": "key", "numbers": "numbers", "pages": "pages"}
FORMAT_PREFIX = r"(?:keynote|numbers|pages)(?:_(?:slide|sheet|body))?"

# The checks every generated package goes through on the Mac.
NATIVE_CHECKS = ("open", "save", "close", "reopen", "export-pdf")


@dataclasses.dataclass(frozen=True)
class Row:
    """One row of the queue table."""

    id: str
    line: int  # 1-based line number in the queue file
    cells: dict[str, str]

    @property
    def wave(self) -> str:
        return self.cells["wave"]

    @property
    def format(self) -> str:
        return self.cells["format"]

    @property
    def route(self) -> str:
        return self.cells["route"].replace("`", "").strip()

    def text(self) -> str:
        return " ".join(self.cells[column] for column in IDENTITY)


def row_id(cells: dict[str, str]) -> str:
    """Content identifier of a row: stable while its five identity cells are."""
    digest = hashlib.sha256("\x1f".join(cells[c].strip() for c in IDENTITY).encode("utf-8"))
    return "q-" + digest.hexdigest()[:12]


def split_cells(line: str) -> list[str]:
    """Cells of a Markdown table line (no escaped pipes occur in the queue)."""
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def parse_queue(text: str) -> list[Row]:
    """Rows of the queue table, in file order, with content identifiers."""
    rows: list[Row] = []
    header: list[str] | None = None
    seen: dict[str, int] = {}
    for number, line in enumerate(text.splitlines(), start=1):
        if not line.startswith("|"):
            header = None
            continue
        cells = split_cells(line)
        if header is None:
            header = [cell.lower() for cell in cells]
            continue
        if all(set(cell) <= set("-: ") for cell in cells):
            continue
        if tuple(header) != COLUMNS:
            raise ValueError(f"line {number}: unexpected queue header {header}")
        if len(cells) != len(COLUMNS):
            raise ValueError(f"line {number}: {len(cells)} cells, expected {len(COLUMNS)}")
        mapping = dict(zip(COLUMNS, cells))
        base = row_id(mapping)
        seen[base] = seen.get(base, 0) + 1
        identifier = base if seen[base] == 1 else f"{base}-{seen[base]}"
        rows.append(Row(identifier, number, mapping))
    return rows


def read_queue(root: Path = ROOT) -> list[Row]:
    return parse_queue((root / MAC_QUEUE).read_text(encoding="utf-8"))


_TEST = re.compile(r"#\[test\]\s*(?:#\[[^\]]*\]\s*)*fn\s+([a-z0-9_]+)\s*\(")
# Tests a `macro_rules!` table declares: a bare `name_parity` opening a line
# of the invocation, followed by `:` or `,` (shape_style_cases.rs,
# series_write_cases.rs).
_MACRO_TEST = re.compile(r"^\s*([a-z0-9_]+_parity)\s*[:,]", re.MULTILINE)


def parity_tests(root: Path = ROOT) -> dict[str, str]:
    """`#[test]` names of the parity harness (including macro-declared
    ones), mapped to their source file."""
    tests: dict[str, str] = {}
    for file in sorted((root / PARITY_DIR).rglob("*.rs")):
        text = file.read_text(encoding="utf-8")
        relative = file.relative_to(root).as_posix()
        for pattern in (_TEST, _MACRO_TEST):
            for match in pattern.finditer(text):
                tests.setdefault(match.group(1), relative)
    return tests


def expand_braces(token: str) -> list[str]:
    """`a_{b,c}_d` -> [`a_b_d`, `a_c_d`], recursively."""
    match = re.search(r"\{([^{}]*)\}", token)
    if not match:
        return [token]
    out: list[str] = []
    for alternative in match.group(1).split(","):
        out.extend(expand_braces(token[: match.start()] + alternative.strip() + token[match.end():]))
    return out


_TOKEN = re.compile(r"[A-Za-z0-9_*]+(?:\{[A-Za-z0-9_, ]+\}[A-Za-z0-9_*]*)*")
_SIBLINGS = re.compile(r"([a-z0-9_]+_parity) and siblings")


def cited_tests(text: str, tests: dict[str, str]) -> list[str]:
    """Parity tests a row cites, in citation order.

    Accepts exact names, brace alternatives (`x_{set,clear}_parity`), a `*`
    standing for the format prefix (`*_chart_pie_start_angle_parity`) and
    "`name` and siblings" (every format's tests of `name`'s family).
    """
    found: list[str] = []

    def add(name: str) -> None:
        if name in tests and name not in found:
            found.append(name)

    for token in _TOKEN.findall(text):
        for name in expand_braces(token):
            if "*" in name:
                pattern = re.compile("^" + re.escape(name).replace(r"\*", FORMAT_PREFIX) + "$")
                for test in sorted(tests):
                    if pattern.match(test):
                        add(test)
            else:
                add(name)
    for name in _SIBLINGS.findall(text):
        stem = re.sub("^" + FORMAT_PREFIX + "_", "", name)
        stem = re.sub(r"_[a-z0-9]+_parity$", "", stem)
        pattern = re.compile("^" + FORMAT_PREFIX + "_" + re.escape(stem) + r"_[a-z0-9_]*parity$")
        for test in sorted(tests):
            if pattern.match(test):
                add(test)
    return found


_RECORD = re.compile(r"(?:\brecords?\s+|\()(0[0-9]{3}[a-z]{0,3})\b")


def cited_records(text: str) -> list[str]:
    """Record numbers a row cites (`record 0372`, `(0375b)`), in order."""
    out: list[str] = []
    for match in _RECORD.finditer(text):
        if match.group(1) not in out:
            out.append(match.group(1))
    for match in re.finditer(r"records?\s+0[0-9]{3}[a-z]{0,3}(?:(?:,\s*|\s+and\s+)(0[0-9]{3}[a-z]{0,3}))+", text):
        for extra in re.findall(r"0[0-9]{3}[a-z]{0,3}", match.group(0)):
            if extra not in out:
                out.append(extra)
    return out


def record_path(record: str, root: Path = ROOT) -> str | None:
    """The tracked record file of a record number, if one exists."""
    matches = sorted((root / RECORDS_DIR).glob(f"{record}-*.md"))
    return matches[0].relative_to(root).as_posix() if matches else None


def is_host_route(row: Row) -> bool:
    """Rows whose route is the host's (`litchi_iwa::...`), not a leaf's."""
    route = row.route
    return route.startswith("host ") or route.startswith("litchi_iwa::")


def formats_of(row: Row) -> list[str]:
    fmt = row.format.strip().lower()
    return [fmt] if fmt in APPS else sorted(APPS)


# ---------------------------------------------------------------------------
# Explicit mappings: rows the generic citation rules cannot resolve, or must
# not (the `File` rows cite the leaf differential, but the native check must
# open what `File` wrote).

# `litchi_ops::File` routes: stems written by the `mac_queue_file_routes`
# example, `<format>-<operation>`.
FILE_ROUTE_STEMS = {
    ("get", "cell.comment"): "cell-comment-set",
    ("set", "cell.comment"): "cell-comment-set",
    ("set", "cell.comment=null"): "cell-comment-clear",
    ("get", "chart.font.size"): "chart-font-size-set",
    ("set", "chart.font.size"): "chart-font-size-set",
    ("set", "chart.font.size=null"): "chart-font-size-reset",
}

# Pages documents built by the same example (`Identity::generate()`).
BUILDER_STEMS = ("pages-builder-identity-a", "pages-builder-identity-b")

# Host routes retired from the public host API whose fixed writer the leaf
# now runs: route prefix -> example stems. The host `numbers::Draft::set_cell`
# is crate-private since W1; `File::set` reaches the same shared TST writer.
EXAMPLE_ROUTES = {
    "litchi_iwa::numbers::Draft::{set_cell, set_cells}": ("numbers-cell-value-set",),
}

# Host-route rows: route prefix -> parity cases whose host output is checked.
HOST_CASES = {
    "litchi_iwa::keynote::Draft slide-table set-cell routes": ["keynote_slide_table_cell_set_parity"],
    "litchi_iwa::pages::Draft::{set_table_cell, set_table_cells}": ["pages_body_table_cell_write_corpus_parity"],
    "litchi_iwa::pages::Draft::remove_text_box": ["pages_text_box_remove_parity"],
    "litchi_iwa::pages::Draft::duplicate_body_shape": ["pages_body_shape_duplicate_parity"],
    "litchi_iwa::pages::Draft::remove_body_chart": ["pages_body_chart_remove_parity"],
    "litchi_iwa::keynote::Draft::clear_slide_table_cell_comment": ["keynote_slide_table_comment_clear_with_replies_parity"],
    "litchi_iwa::keynote::Draft::add_slide_table_cell_comment_reply": ["keynote_slide_table_comment_reply_add_parity"],
    "litchi_iwa::keynote::Draft::set_slide_table_cell_comment": ["keynote_slide_table_comment_set_with_replies_parity"],
    "litchi_iwa::keynote::Draft::set_slide_table_cell_comment_reply": ["keynote_slide_table_comment_reply_set_parity"],
    "litchi_iwa::keynote::Draft::remove_slide_table_cell_comment_reply": ["keynote_slide_table_comment_reply_remove_parity"],
    "litchi_iwa::pages::Draft::clear_table_cell_comment": ["pages_body_table_comment_clear_with_replies_parity"],
    "litchi_iwa::pages::Draft::set_table_cell_comment": ["pages_body_table_comment_set_with_replies_parity"],
    "litchi_iwa::pages::Draft::add_table_cell_comment_reply": ["pages_body_table_comment_reply_add_parity"],
    "litchi_iwa::pages::Draft::set_table_cell_comment_reply": ["pages_body_table_comment_reply_set_parity"],
    "litchi_iwa::pages::Draft::remove_table_cell_comment_reply": ["pages_body_table_comment_reply_remove_parity"],
}

# Leaf rows whose cited test is not a harness case: route prefix -> the
# harness cases that write the same mutation.
LEAF_CASES = {
    # numbers_owner_allocation.rs compares engines directly; the leaf table
    # duplicate cases write the packages.
    "litchi_numbers_wire::formula_graph::next_owner_id": [
        "numbers_table_duplicate_parity",
        "numbers_table_duplicate_corpus_parity",
    ],
}

# Rows with no Litchi-generated package: the operator authors the input in
# the native application, as the row describes.
NATIVE_AUTHORING = {
    "native-created volatile highlight": "the row asks for a Numbers-created volatile-highlight document; author it as its Fixture column says and commit it under test-data/iwork/numbers/ (record 0376b)",
    "litchi_numbers::table::pivot::Aggregate": "the row asks for a native pivot workbook with one value column per Summarize-by function; author it as its Reason column says",
}


@dataclasses.dataclass(frozen=True)
class Source:
    """One package a row needs: a parity case's input/host/leaf output, or a
    package written by the `mac_queue_file_routes` example."""

    kind: str  # "parity" or "example"
    name: str  # case name, or example stem
    role: str  # "leaf", "host" (parity), "file" (example)


@dataclasses.dataclass(frozen=True)
class Mapping:
    sources: tuple[Source, ...]
    group: str | None = None  # set when no package is generated
    reason: str = ""
    rule: str = ""


def _file_route(row: Row) -> tuple[str, str] | None:
    match = re.match(r"litchi_ops::File::(get|set)\s+\((.*)\)\s*$", row.route)
    if not match:
        return None
    detail = match.group(2)
    prop = re.match(r"([a-z.]+(?:=null)?)", detail)
    return (match.group(1), prop.group(1)) if prop else None


def ledger_cases(root: Path = ROOT) -> dict[str, list[str]]:
    """Exit-ledger destinations mapped to their parity tests."""
    path = root / LEDGER
    out: dict[str, list[str]] = {}
    if not path.is_file():
        return out
    lines = path.read_text(encoding="utf-8").splitlines()
    header = lines[0].split("\t")
    for line in lines[1:]:
        cells = dict(zip(header, line.split("\t")))
        destination = cells.get("destination", "-")
        test = cells.get("parity_test", "-")
        if destination in ("", "-") or test in ("", "-"):
            continue
        out.setdefault(destination, [])
        if test not in out[destination]:
            out[destination].append(test)
    return out


def map_row(row: Row, tests: dict[str, str], ledger: dict[str, list[str]] | None = None) -> Mapping:
    """Map one queue row to the packages its native check opens."""
    route = row.route
    for prefix, reason in NATIVE_AUTHORING.items():
        if route.startswith(prefix):
            return Mapping((), "native-authoring", reason, "explicit")
    file_route = _file_route(row)
    if file_route is not None:
        stem = FILE_ROUTE_STEMS.get(file_route)
        if stem is None:
            return Mapping((), None, f"unknown File route {file_route}", "unmapped")
        return Mapping(
            tuple(Source("example", f"{fmt}-{stem}", "file") for fmt in formats_of(row)),
            rule="file-route",
        )
    if route.startswith("litchi_pages::builder::Identity::from_parts"):
        return Mapping(tuple(Source("example", stem, "file") for stem in BUILDER_STEMS), rule="builder")
    host_route = route.removeprefix("host ").strip()
    for prefix, stems in EXAMPLE_ROUTES.items():
        if host_route.startswith(prefix):
            return Mapping(tuple(Source("example", stem, "file") for stem in stems), rule="example-explicit")
    for prefix, cases in HOST_CASES.items():
        if host_route.startswith(prefix):
            return Mapping(tuple(Source("parity", case, "host") for case in cases), rule="host-explicit")
    for prefix, cases in LEAF_CASES.items():
        if route.startswith(prefix):
            return Mapping(tuple(Source("parity", case, "leaf") for case in cases), rule="leaf-explicit")
    role = "host" if is_host_route(row) else "leaf"
    cited = cited_tests(row.text(), tests)
    if cited:
        return Mapping(tuple(Source("parity", case, role) for case in cited), rule="cited")
    destination = route.split(" ")[0]
    from_ledger = [case for case in (ledger or {}).get(destination, []) if case in tests]
    if from_ledger:
        return Mapping(tuple(Source("parity", case, role) for case in from_ledger), rule="ledger")
    return Mapping((), None, "no parity case cited and no explicit mapping", "unmapped")


def explicit_cases() -> list[str]:
    """Every parity case an explicit mapping names (validated by the tests)."""
    names = [case for cases in HOST_CASES.values() for case in cases]
    names += [case for cases in LEAF_CASES.values() for case in cases]
    return names


def example_stems() -> list[str]:
    """Every stem the `mac_queue_file_routes` example writes."""
    stems = [f"{fmt}-{stem}" for fmt in sorted(APPS) for stem in sorted(set(FILE_ROUTE_STEMS.values()))]
    stems += list(BUILDER_STEMS)
    stems += [stem for group in EXAMPLE_ROUTES.values() for stem in group]
    return stems
