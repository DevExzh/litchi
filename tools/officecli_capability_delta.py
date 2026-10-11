#!/usr/bin/env python3
"""Generate the OfficeCLI capability delta table (record 0010 §7.2, item 6).

The table compares OfficeCLI's help schemas (``schemas/help/<format>/*.json``,
with their ``extends`` bases merged the way OfficeCLI's ``SchemaHelpLoader``
merges them) with Litchi's golden capability reports
(``crates/litchi-ops/tests/golden/capabilities/<format>.json``, the JSON of
``Capabilities::of(format)``) for the three formats OfficeCLI describes:
DOCX, XLSX and PPTX, plus XLS and XLSB against the XLSX vocabulary as a parity
target (not a claim of OfficeCLI XLS/XLSB file support). DOCX's optional
``docx-encryption.json`` baseline contributes feature-gated operation gaps
that are absent from the default report. It writes
``docs/unified-ops/officecli-capability-delta.md``.

OfficeCLI is not tracked in git, so the script runs locally::

    python3 tools/officecli_capability_delta.py --officecli /path/to/OfficeCLI

``--officecli`` defaults to ``$LITCHI_OFFICECLI``, then to ``3rdparty/OfficeCLI``
in the repository. The table records the SHA-256 of every golden report it
read and of the schema tree, and ``tools/test_officecli_capability_delta.py``
fails when a golden report no longer matches the hash recorded here. Rerun
the script whenever the goldens change: after an adapter lands (R4a, R4c,
...) and its goldens are blessed with ``LITCHI_BLESS=1``. ``--check``
regenerates the table in memory and fails if the committed one differs.

The output is deterministic: no timestamps, sorted where OfficeCLI's order is
not meaningful, so rerunning on the same inputs reproduces it byte for byte.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
GOLDEN_DIR = Path("crates/litchi-ops/tests/golden/capabilities")
OUTPUT = Path("docs/unified-ops/officecli-capability-delta.md")
FORMATS = ("docx", "xlsx", "xlsb", "xls", "pptx")


def reference_format(fmt: str) -> str:
    """Binary Excel formats use SpreadsheetML vocabulary as a comparison target."""
    return "xlsx" if fmt in ("xls", "xlsb") else fmt


# OfficeCLI elements Litchi models inside another element: their properties
# and their status follow that element. Each entry is curated, with the reason.
FOLDED: dict[str, dict[str, tuple[str, str]]] = {
    "xlsx": {
        **{
            rule: ("conditionalFormatting", "a conditional-format rule type")
            for rule in (
                "aboveaverage",
                "cellis",
                "cfextended",
                "colorscale",
                "containstext",
                "databar",
                "dateoccurring",
                "duplicatevalues",
                "formulacf",
                "iconset",
                "topn",
                "uniquevalues",
            )
        },
        "pagebreak": ("rowbreak", "OfficeCLI's dispatcher to rowbreak and colbreak"),
    },
    "docx": {},
    "pptx": {},
}

# Keys OfficeCLI's node builders emit under another name than its schema
# declares, so `get` returns the first: (format, element) -> {schema key:
# emitted key}. Each entry cites the builder.
NODE_KEYS: dict[tuple[str, str], dict[str, str]] = {
    # ExcelHandler.Helpers.Node.cs `TableToNode` writes `firstCol` and
    # `lastCol`; table.json lists neither as an alias.
    ("xlsx", "table"): {"showFirstColumn": "firstCol", "showLastColumn": "lastCol"},
}

# OfficeCLI elements outside the unified read surface, with the reason.
OUT_OF_SCOPE: dict[str, str] = {
    "raw": "the raw XML layer (`raw`/`raw-set` commands), not an element (design §3)",
    "sort": "a set-only action that reorders rows; M2 write surface",
    "markdown": "an add-time import format, not an element of the document",
}

ELEMENT_STATUSES = (
    "supported",
    "gap",
    "folded",
    "pending",
    "missing",
    "out of scope",
    "litchi-only",
)
PROPERTY_STATUSES = (
    "supported",
    "gap",
    "element gap",
    "pending",
    "write-only",
    "raw-only",
    "missing",
    "litchi-only",
)


def normalize(name: str) -> str:
    """Element and key names compare case-insensitively, without dashes."""
    return name.lower().replace("-", "")


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def merge(base: dict, override: dict) -> dict:
    """OfficeCLI ``SchemaHelpLoader.MergeSchemaJson``: top-level members of the
    override replace the base's; ``properties`` are the override's, in order,
    then the base's that the override does not redeclare."""
    merged = {key: value for key, value in base.items() if key != "shared_base"}
    for key, value in override.items():
        if key == "extends":
            continue
        if key == "properties":
            combined = dict(value)
            for name, prop in base.get("properties", {}).items():
                combined.setdefault(name, prop)
            merged["properties"] = combined
        else:
            merged[key] = value
    return merged


def load_schema(help_dir: Path, path: Path) -> dict:
    """A schema with its ``extends`` bases composed, as OfficeCLI loads it."""
    schema = json.loads(path.read_text(encoding="utf-8"))
    bases = schema.get("extends")
    if not bases:
        return schema
    if isinstance(bases, str):
        bases = [bases]
    composed = json.loads((help_dir / f"{bases[0]}.json").read_text(encoding="utf-8"))
    for base in bases[1:]:
        composed = merge(
            composed, json.loads((help_dir / f"{base}.json").read_text(encoding="utf-8"))
        )
    return merge(composed, schema)


def schema_tree_hash(help_dir: Path) -> tuple[str, int]:
    """One SHA-256 over every schema file, by relative path and content."""
    digest = hashlib.sha256()
    files = sorted(path for path in help_dir.rglob("*.json") if path.is_file())
    for path in files:
        digest.update(path.relative_to(help_dir).as_posix().encode())
        digest.update(b"\0")
        digest.update(path.read_bytes())
        digest.update(b"\0")
    return digest.hexdigest(), len(files)


def officecli_commit(root: Path) -> str:
    try:
        result = subprocess.run(
            ["git", "-C", str(root), "rev-parse", "--short=8", "HEAD"],
            check=True,
            capture_output=True,
            text=True,
        )
    except (OSError, subprocess.CalledProcessError):
        return "unknown"
    return result.stdout.strip() or "unknown"


@dataclass
class OcElement:
    """One OfficeCLI element schema."""

    stem: str
    name: str
    aliases: list[str]
    parents: list[str]
    operations: dict
    properties: dict

    def names(self) -> set[str]:
        return {normalize(self.stem), normalize(self.name), *map(normalize, self.aliases)}


@dataclass
class LitchiElement:
    """One element of a Litchi capability report."""

    name: str
    aliases: list[str]
    gap: dict | None
    properties: dict
    predicate_keys: list[str]

    def names(self) -> set[str]:
        return {normalize(self.name), *map(normalize, self.aliases)}


@dataclass
class FormatDelta:
    """The rows of one format's delta."""

    name: str
    pending: dict | None = None
    elements: list[tuple[str, str, str, str]] = field(default_factory=list)
    properties: list[tuple[str, str, str, str, str, str]] = field(default_factory=list)
    operation_gaps: list[tuple[str, str, str, str, str]] = field(default_factory=list)

    def counts(self, rows, statuses):
        counts = {status: 0 for status in statuses}
        for row in rows:
            counts[row[-2]] += 1
        return counts


def path_names(schema: dict) -> list[str]:
    """The step names OfficeCLI's own path forms address the element by:
    the last step of each form that selects with a predicate
    (`/SheetName/chart[N]/axis[@role=ROLE]` gives `axis`)."""
    forms = [*schema.get("paths", {}).get("stable", []), *schema.get("paths", {}).get("positional", [])]
    addressing = schema.get("addressing") or {}
    if "pathForm" in addressing:
        forms.append(addressing["pathForm"])
    names = []
    for form in forms:
        last = form.rsplit("/", 1)[-1]
        if "[" in last:
            name = last.split("[", 1)[0]
            if name.isalnum() and name not in names:
                names.append(name)
    return names


def officecli_elements(help_dir: Path, fmt: str) -> list[OcElement]:
    schemas = []
    for path in sorted((help_dir / reference_format(fmt)).glob("*.json")):
        schema = load_schema(help_dir, path)
        parent = schema.get("parent") or []
        if isinstance(parent, str):
            parent = [part.strip() for part in parent.split("|") if part.strip()]
        aliases = list(schema.get("elementAliases", []))
        for name in path_names(schema):
            if name not in aliases:
                aliases.append(name)
        schemas.append(
            OcElement(
                stem=path.stem,
                name=schema.get("element", path.stem),
                aliases=aliases,
                parents=list(parent),
                operations=schema.get("operations", {}),
                properties=schema.get("properties", {}),
            )
        )
    # A parent's child path segment names the child too (`axis` for
    # `chart-axis`), when it is a plain word.
    by_name = {normalize(schema.name): schema for schema in schemas}
    for path in sorted((help_dir / reference_format(fmt)).glob("*.json")):
        for child in load_schema(help_dir, path).get("children", []):
            segment = child.get("pathSegment", "")
            target = by_name.get(normalize(child.get("element", "")))
            if target and segment and segment.replace("-", "").isalnum():
                if normalize(segment) not in target.names():
                    target.aliases.append(segment)
    return schemas


def litchi_elements(report: dict) -> list[LitchiElement]:
    return [
        LitchiElement(
            name=element["element"],
            aliases=list(element.get("elementAliases", [])),
            gap=element.get("gap"),
            properties=element.get("properties", {}),
            predicate_keys=list(element.get("predicates", {}).get("keys", [])),
        )
        for element in report.get("elements", [])
    ]


def ops_letters(spec: dict) -> str:
    """OfficeCLI's `asg` flags of a property, `-` for each verb it lacks."""
    return "".join(
        letter if spec.get(verb) else "-" for letter, verb in (("a", "add"), ("s", "set"), ("g", "get"))
    )


def follow(gap: dict | None) -> str:
    return gap["followUp"] if gap else ""


def operation_gap_rows(report: dict) -> list[tuple[str, str, str, str, str]]:
    """Return contextual operation gaps in report order.

    These gaps describe an operation that is available for an element in the
    ordinary case but refused for a particular selection or leaf context.
    They do not participate in the OfficeCLI element/property status counts,
    so keep them as a separate report section.
    """
    return [
        (
            gap["verb"],
            gap["element"],
            gap["context"],
            gap["reason"],
            gap["followUp"],
        )
        for gap in report.get("gaps", [])
        if gap.get("subject") == "operation"
    ]


def _operation_gap_key(gap: dict) -> tuple:
    """The serialized fields that identify one contextual operation gap."""
    return tuple(
        gap.get(field)
        for field in ("verb", "element", "context", "code", "followUp", "reason")
    )


def merge_optional_operation_gaps(report: dict, optional_reports: list[dict]) -> dict:
    """Add operation gaps from feature-specific reports to a base report.

    Capability goldens are format baselines, while a feature such as DOCX
    encryption can add a refusal without changing the ordinary element or
    property surface. Keep the default report authoritative for those rows and
    union only operation gaps from optional reports, deduplicated by their
    serialized identity.
    """
    merged = dict(report)
    gaps = list(report.get("gaps", []))
    seen = {
        _operation_gap_key(gap)
        for gap in gaps
        if gap.get("subject") == "operation"
    }
    for optional in optional_reports:
        for gap in optional.get("gaps", []):
            if gap.get("subject") != "operation":
                continue
            key = _operation_gap_key(gap)
            if key not in seen:
                gaps.append(gap)
                seen.add(key)
    merged["gaps"] = gaps
    return merged


def compare(fmt: str, officecli: list[OcElement], report: dict) -> FormatDelta:
    delta = FormatDelta(fmt, operation_gaps=operation_gap_rows(report))
    verbs = report.get("verbs", {})
    if not verbs.get("get", {}).get("supported", False):
        delta.pending = verbs.get("get", {})
        for schema in officecli:
            reason = OUT_OF_SCOPE.get(schema.stem) or OUT_OF_SCOPE.get(schema.name)
            if reason:
                delta.elements.append((schema.name, "", "out of scope", reason))
                continue
            delta.elements.append(
                (schema.name, "", "pending", delta.pending.get("followUp", ""))
            )
            for key, spec in schema.properties.items():
                delta.properties.append(
                    (schema.name, key, spec.get("type", ""), ops_letters(spec), "pending", "")
                )
        return delta

    litchi = litchi_elements(report)
    index: dict[str, LitchiElement] = {}
    for element in litchi:
        for name in element.names():
            index.setdefault(name, element)
    property_gaps: dict[tuple[str, str], dict] = {}
    for gap in report.get("gaps", []):
        if gap.get("subject") == "property":
            property_gaps[(normalize(gap["element"]), normalize(gap["key"]))] = gap

    matched: set[str] = set()
    folded = FOLDED.get(reference_format(fmt), {})
    for schema in officecli:
        target = next((index[name] for name in sorted(schema.names()) if name in index), None)
        note = ""
        if schema.stem in OUT_OF_SCOPE or schema.name in OUT_OF_SCOPE:
            reason = OUT_OF_SCOPE.get(schema.stem) or OUT_OF_SCOPE[schema.name]
            delta.elements.append((schema.name, "", "out of scope", reason))
            continue
        if fmt == "docx" and schema.name in {"fieldChar", "instrText", "ptab"}:
            target = None
        if target is None and schema.stem in folded:
            into, why = folded[schema.stem]
            target = index.get(normalize(into))
            note = f"{why}; in `{into}`"
            if target is None:
                raise SystemExit(f"{fmt}: {schema.stem} is folded into {into}, which the report lacks")
            status = "folded"
        elif target is None:
            delta.elements.append((schema.name, "", "missing", ""))
            for key, spec in schema.properties.items():
                status = "missing" if spec.get("get") else "write-only"
                delta.properties.append(
                    (schema.name, key, spec.get("type", ""), ops_letters(spec), status, "")
                )
            continue
        else:
            status = "gap" if target.gap else "supported"
            matched.add(target.name)
        follow_up = follow(target.gap)
        if status == "folded" and target.gap:
            note = f"{note} (gap)"
        delta.elements.append((schema.name, target.name, status, follow_up or note))

        litchi_keys = {normalize(key): key for key in target.properties}
        seen_litchi: set[str] = set()
        emitted = NODE_KEYS.get((reference_format(fmt), schema.name), {})
        for key, spec in schema.properties.items():
            names = [key, *(spec.get("aliases") or [])] if isinstance(spec.get("aliases"), list) else [key]
            if key in emitted:
                names.append(emitted[key])
            names = [normalize(name) for name in names]
            hit = next((litchi_keys[name] for name in names if name in litchi_keys), None)
            gap = next(
                (property_gaps[(normalize(target.name), name)] for name in names
                 if (normalize(target.name), name) in property_gaps),
                None,
            )
            if (reference_format(fmt), schema.name, key) in {("docx", "paragraph", "paraId"), ("docx", "sdt", "id"), ("xlsx", "cell", "numFmtId")}:
                row_status, row_follow = "raw-only", "ADR 0034: native selection requires explicit raw admission"
            elif target.gap:
                row_status, row_follow = "element gap", follow(target.gap)
            elif hit is not None:
                seen_litchi.add(hit)
                row_status, row_follow = "supported", ""
            elif any(name in map(normalize, target.predicate_keys) for name in names):
                row_status, row_follow = "supported", "the node's text"
            elif gap is not None:
                row_status, row_follow = "gap", gap["followUp"]
            elif not spec.get("get"):
                row_status, row_follow = "write-only", "M2"
            else:
                row_status, row_follow = "missing", ""
            delta.properties.append(
                (schema.name, key, spec.get("type", ""), ops_letters(spec), row_status, row_follow)
            )
        if status != "folded":
            for key, spec in target.properties.items():
                if key not in seen_litchi:
                    delta.properties.append(
                        (schema.name, key, spec.get("type", ""), "", "litchi-only", "")
                    )

    for element in litchi:
        if element.name not in matched and not any(
            row[1] == element.name for row in delta.elements
        ):
            delta.elements.append(("", element.name, "litchi-only", follow(element.gap)))
    return delta


def render(
    deltas: list[FormatDelta],
    goldens: list[tuple[str, str]],
    tree: tuple[str, int],
    commit: str,
) -> str:
    lines = [
        "# OfficeCLI capability delta (generated)",
        "",
        "Generated by `tools/officecli_capability_delta.py`; do not edit by hand.",
        "It compares OfficeCLI's help schemas with Litchi's golden capability",
        "reports (`Capabilities::of(format)`, record 0010 §7), for the formats",
        "OfficeCLI describes. M1 is read-only, so a property counts as supported",
        "when Litchi reads it; OfficeCLI's add/set/get flags are shown as `asg`.",
        "XLS and XLSB are compared against the XLSX help vocabulary as Litchi's parity",
        "target; this does not claim OfficeCLI itself supports XLS or XLSB files.",
        "DOCX operation gaps also include the feature-specific encryption baseline",
        "when that golden is present.",
        "",
        "Rerun the script after blessing the goldens (`LITCHI_BLESS=1`), for",
        "example when an adapter lands; `tools/test_officecli_capability_delta.py`",
        "fails while a golden differs from the hash recorded below.",
        "",
        "## Inputs",
        "",
        "| Input | File | SHA-256 |",
        "|---|---|---|",
    ]
    for file, digest in goldens:
        lines.append(f"| golden | `{file}` | `{digest}` |")
    lines.append(
        f"| officecli | `schemas/help` ({tree[1]} files, OfficeCLI commit `{commit}`) | `{tree[0]}` |"
    )
    lines += [
        "",
        "## Statuses",
        "",
        "| Status | Meaning |",
        "|---|---|",
        "| supported | Litchi reads it (a property, or the node's text for a predicate key such as a cell's `value`) |",
        "| gap | Litchi lists it in `capabilities()` as a gap, with the follow-up shown |",
        "| element gap | a property of an element Litchi lists as a gap |",
        "| folded | an OfficeCLI element Litchi models inside another element (named in the follow-up) |",
        "| pending | the format's adapter has not landed (the follow-up names its M1 step) |",
        "| write-only | OfficeCLI cannot read it either (no `get`); M2's write surface decides |",
        "| raw-only | native selection metadata excluded from ordinary APIs by ADR 0034 |",
        "| operation gap | an element operation refused for a particular selection or leaf context; listed in the contextual gap table |",
        "| missing | OfficeCLI reads it and Litchi neither reads nor lists it: untriaged |",
        "| out of scope | not an element of the unified read surface |",
        "| litchi-only | Litchi reads it and OfficeCLI's schema lacks it |",
        "",
        "## Summary",
        "",
        "Elements:",
        "",
        "| Format | " + " | ".join(ELEMENT_STATUSES) + " |",
        "|---|" + "---|" * len(ELEMENT_STATUSES),
    ]
    for delta in deltas:
        counts = delta.counts(delta.elements, ELEMENT_STATUSES)
        lines.append(f"| {delta.name} | " + " | ".join(str(counts[s]) for s in ELEMENT_STATUSES) + " |")
    lines += [
        "",
        "Properties:",
        "",
        "| Format | " + " | ".join(PROPERTY_STATUSES) + " |",
        "|---|" + "---|" * len(PROPERTY_STATUSES),
    ]
    for delta in deltas:
        counts = delta.counts(delta.properties, PROPERTY_STATUSES)
        lines.append(f"| {delta.name} | " + " | ".join(str(counts[s]) for s in PROPERTY_STATUSES) + " |")

    for delta in deltas:
        lines += ["", f"## {delta.name}", ""]
        if delta.pending is not None:
            lines += [
                f"The {delta.name} adapter has not landed: `get` is "
                f"`{delta.pending.get('code', '')}` ({delta.pending.get('reason', '')}; follow-up "
                f"{delta.pending.get('followUp', '')}). OfficeCLI's elements, with their property counts:",
                "",
                "| OfficeCLI element | Properties | Status | Follow-up |",
                "|---|---|---|---|",
            ]
            for name, _, status, follow_up in delta.elements:
                count = sum(1 for row in delta.properties if row[0] == name)
                lines.append(f"| {name} | {count} | {status} | {follow_up} |")
        else:
            lines += [
                "### Elements",
                "",
                "| OfficeCLI element | Litchi element | Status | Follow-up |",
                "|---|---|---|---|",
            ]
            for name, litchi, status, follow_up in delta.elements:
                lines.append(f"| {name} | {litchi} | {status} | {follow_up} |")
            lines += [
                "",
                "### Properties",
                "",
                "| Element | Property | Type | OfficeCLI `asg` | Status | Follow-up |",
                "|---|---|---|---|---|---|",
            ]
            for element, key, ty, ops, status, follow_up in delta.properties:
                lines.append(f"| {element} | `{key}` | {ty} | {ops} | {status} | {follow_up} |")
        lines += [
            "",
            "### Contextual operation gaps",
            "",
            "These operation gaps apply only to the named selection or leaf context; ordinary element operation support remains in the capability report.",
            "",
            "| Verb | Element | Context | Reason | Follow-up |",
            "|---|---|---|---|---|",
        ]
        if delta.operation_gaps:
            for verb, element, context, reason, follow_up in delta.operation_gaps:
                lines.append(f"| `{verb}` | `{element}` | `{context}` | {reason} | {follow_up} |")
        else:
            lines.append("| — | — | — | no contextual operation gaps registered | — |")
    lines.append("")
    return "\n".join(lines)


def generate(officecli: Path, repo: Path) -> str:
    help_dir = officecli / "schemas" / "help"
    if not help_dir.is_dir():
        raise SystemExit(f"no OfficeCLI help schemas at {help_dir}; pass --officecli")
    deltas = []
    goldens = []
    for fmt in FORMATS:
        golden = repo / GOLDEN_DIR / f"{fmt}.json"
        data = golden.read_bytes()
        goldens.append(((GOLDEN_DIR / f"{fmt}.json").as_posix(), sha256(data)))
        report = json.loads(data)
        if fmt == "docx":
            encrypted_golden = repo / GOLDEN_DIR / "docx-encryption.json"
            if encrypted_golden.is_file():
                encrypted_data = encrypted_golden.read_bytes()
                goldens.append(
                    (
                        (GOLDEN_DIR / "docx-encryption.json").as_posix(),
                        sha256(encrypted_data),
                    )
                )
                report = merge_optional_operation_gaps(
                    report, [json.loads(encrypted_data)]
                )
        deltas.append(compare(fmt, officecli_elements(help_dir, fmt), report))
    return render(deltas, goldens, schema_tree_hash(help_dir), officecli_commit(officecli))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "--officecli",
        type=Path,
        default=Path(os.environ.get("LITCHI_OFFICECLI", REPO / "3rdparty" / "OfficeCLI")),
        help="the OfficeCLI checkout (default: $LITCHI_OFFICECLI, then 3rdparty/OfficeCLI)",
    )
    parser.add_argument("--output", type=Path, default=REPO / OUTPUT)
    parser.add_argument(
        "--check",
        action="store_true",
        help="fail if the committed table differs from a fresh one",
    )
    arguments = parser.parse_args(argv)
    table = generate(arguments.officecli, REPO)
    if len(table.encode()) >= 1 << 20:
        raise SystemExit("the delta table would reach 1 MiB; split it before committing")
    if arguments.check:
        current = arguments.output.read_text(encoding="utf-8") if arguments.output.exists() else ""
        if current != table:
            print(f"{arguments.output} is stale; rerun {Path(__file__).name}", file=sys.stderr)
            return 1
        return 0
    arguments.output.write_text(table, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
