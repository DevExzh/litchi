#!/usr/bin/env python3
"""Generate the OfficeCLI capability delta table (record 0010 §7.2, item 6).

The table compares OfficeCLI's help schemas (``schemas/help/<format>/*.json``,
with their ``extends`` bases merged the way OfficeCLI's ``SchemaHelpLoader``
merges them) with Litchi's property registry export
(``crates/litchi-ops/tests/golden/registry/<format>.json``, the golden JSON
of ``litchi_ops::schema(format)``; record 0263) for every format the
registry covers. OfficeCLI describes DOCX, XLSX and PPTX; the other formats
are compared against the vocabulary of their family as a parity target (not
a claim of OfficeCLI support for them): DOC and RTF against DOCX, XLS and
XLSB against XLSX, PPT against PPTX. An optional ``<format>-encryption.json``
snapshot contributes feature-gated operation gaps that are absent from the
default export. Each property row carries a read status (Litchi's ``get``
against OfficeCLI's) and a write status (Litchi's ``add``/``set`` against
OfficeCLI's). It writes ``docs/unified-ops/officecli-capability-delta.md``.

The OfficeCLI help schemas it reads are vendored, unchanged, under
``test-data/officecli/`` (record 0257; provenance, license and the pinned
upstream commit sit next to them), so the script runs anywhere::

    python3 tools/officecli_capability_delta.py

``--officecli`` (default ``$LITCHI_OFFICECLI``, then ``test-data/officecli``)
names another OfficeCLI tree, such as a newer upstream checkout; refresh the
vendored copy from it before regenerating the committed table. The table
records the SHA-256 of every registry snapshot it read and of the schema
tree, and ``tools/test_officecli_capability_delta.py`` fails when a snapshot
no longer matches the hash recorded here. Rerun the script whenever the
registry changes and its snapshots are blessed with ``LITCHI_BLESS=1``
(``cargo test -p litchi-ops --test registry``). ``--check`` regenerates the
table in memory and fails if the committed one differs.

Every gap row carries a disposition from the curated
``tools/officecli_dispositions.json`` (record 0332): ``planned`` (with a work
package or record id), ``deferred-P1``/``deferred-P2`` (by design §13 tier),
``wontfix`` or ``out-of-scope`` (with a reason). The property tables are
long, so each format's sits in its own page under
``docs/unified-ops/officecli-capability-delta/``. ``--check`` also fails on
an untriaged (``missing``) row, a gap without a disposition, a P0 row that is
deferred or wontfix, a milestone-only follow-up without a current id, and a
disposition or tier key that matches nothing. Record 0344 (audit WP13)
extends the milestone rule to every ``followUp`` of the registry and
capability goldens (``crates/litchi-ops/tests/golden/capabilities``): each
one that names only a milestone must match a ``followUps`` entry of the
dispositions file, which gives its current record or work-package id (the
tables show it after the text), and every entry must match one.

The output is deterministic: no timestamps, sorted where OfficeCLI's order is
not meaningful, so rerunning on the same inputs reproduces it byte for byte.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass, field
from fnmatch import fnmatchcase
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
VENDORED_OFFICECLI = Path("test-data/officecli")
# The vendored tree has no git metadata; this file pins its upstream commit.
UPSTREAM_COMMIT = "UPSTREAM_COMMIT"
GOLDEN_DIR = Path("crates/litchi-ops/tests/golden/registry")
OUTPUT = Path("docs/unified-ops/officecli-capability-delta.md")
# Per-format property tables, next to the main table.
PROPERTY_DIR = Path("docs/unified-ops/officecli-capability-delta")
DISPOSITIONS = Path("tools/officecli_dispositions.json")
# The capability goldens (`File::capabilities()` per format); record 0344
# reads their follow-ups for the milestone rule.
CAPABILITY_DIR = Path("crates/litchi-ops/tests/golden/capabilities")
# Generated files stay below the repository's 1 MiB evidence limit.
SIZE_LIMIT = 1 << 20
# Table order: each OfficeCLI format, then the formats compared against its
# vocabulary. A registry snapshot of another format is appended, sorted.
FORMATS = ("docx", "doc", "rtf", "xlsx", "xlsb", "xls", "pptx", "ppt")
OPTIONAL_SUFFIX = "-encryption"

REFERENCE = {
    "doc": "docx",
    "rtf": "docx",
    "odt": "docx",
    "xls": "xlsx",
    "xlsb": "xlsx",
    "ods": "xlsx",
    "ppt": "pptx",
    "odp": "pptx",
}


def reference_format(fmt: str) -> str:
    """The OfficeCLI format whose vocabulary ``fmt`` is compared against:
    its own for DOCX, XLSX and PPTX, its family's otherwise."""
    return REFERENCE.get(fmt, fmt)


def formats(repo: Path) -> list[str]:
    """Every format with a registry snapshot, in table order."""
    present = sorted(
        path.stem
        for path in (repo / GOLDEN_DIR).glob("*.json")
        if not path.stem.endswith(OPTIONAL_SUFFIX)
    )
    known = [fmt for fmt in FORMATS if fmt in present]
    return known + [fmt for fmt in present if fmt not in FORMATS]


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

# OfficeCLI elements whose aliases or path forms name a Litchi element that
# is not the same thing: a DOCX `fieldChar`, `instrText` or `ptab` is a run
# subtype, not a run; a PPTX `diagram` (a Mermaid rendering placed on a
# slide, path `/slide[N]/...`) is not the slide.
UNMATCHED: dict[str, set[str]] = {
    "docx": {"fieldChar", "instrText", "ptab"},
    "pptx": {"diagram"},
}

# OfficeCLI elements outside the unified read surface, with the reason.
OUT_OF_SCOPE: dict[str, str] = {
    "raw": "the raw XML layer: the `raw`, `raw-set`, `add-part` and `validate` commands (record 0336), not an element (design §12)",
    "sort": "a set-only action that reorders rows, not an element (design §13 P1: sort)",
    "markdown": "an add-time import format, not an element of the document",
}

ELEMENT_STATUSES = (
    "supported",
    "gap",
    "folded",
    "pending",
    "missing",
    "unlisted",
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
    "unlisted",
    "litchi-only",
)
WRITE_STATUSES = (
    "supported",
    "partial",
    "gap",
    "pending",
    "litchi-only",
    "—",
)

# Property row fields: element, key, type, OfficeCLI asg, read status,
# follow-up, Litchi asg, write status.
READ_STATUS = 4
WRITE_STATUS = 7

# Element and read statuses that are gaps, and write statuses that are gaps.
ELEMENT_GAPS = ("gap", "missing", "unlisted", "pending", "out of scope")
READ_GAPS = ("gap", "element gap", "missing", "unlisted", "pending", "raw-only")
WRITE_GAPS = ("gap", "partial", "pending")
# A `missing` row (untriaged) that the dispositions classify is `unlisted`:
# still absent from `capabilities()`, but with a recorded decision.
TRIAGED = "unlisted"


# ---------------------------------------------------------------------------
# Dispositions (record 0332)
# ---------------------------------------------------------------------------

DISPOSITION_VALUES = ("planned", "deferred-P1", "deferred-P2", "wontfix", "out-of-scope")
# `deferred` in the file resolves to the row's tier (`deferred-P1`/`-P2`).
DISPOSITION_INPUTS = (*DISPOSITION_VALUES, "deferred")
TIER_VALUES = ("P0", "v1", "P1", "P2")
# Tiers design §13 puts in v1: never deferred and never wontfix.
V1_TIERS = ("P0", "v1")
# Dispositions that account for a gap in the v1-scope corpus subset.
ACCOUNTED = ("wontfix", "deferred-P1", "deferred-P2", "out-of-scope")
# A current work-package or record id: `WP4`, `0328`, `0324a`.
REF = re.compile(r"(WP\d+|\d{4}[a-z]?)\b")
# A follow-up that names only a closed milestone or a step of one
# ("M2 (validate)", "M1-extended (R4x)", "R4x", "M1-Δ1"): its row's
# disposition must name a current id (audit WP13).
MILESTONE = re.compile(r"(^|[\s(])(M\d[\w-]*|R\d+[a-z]*)\b")
RECORD = re.compile(r"\b(record \d{4}|\d{4}[a-z]?\b|WP\d+|L-[A-Z]+-[A-Z]?\d)")


def milestone_only(follow_up: str) -> bool:
    """Whether a follow-up names a milestone but no record or work item."""
    return bool(MILESTONE.search(follow_up)) and not RECORD.search(follow_up)


def _literal(segment: str) -> int:
    return sum(1 for char in segment if char not in "*?[]")


def specificity(pattern: str) -> tuple[int, ...]:
    """How specific a key pattern is, most significant part first: a
    pattern that names the property beats one that names only the element,
    which beats one that names only the format. Delta and tier keys are
    ``[delta:]<format>:<element>[.<property>]``; corpus keys are
    ``corpus:<format>|<verb>|<element>|<refusal>|<property>`` (or a named
    class, ``corpus:<id>[|<verb>]``)."""
    if pattern.startswith("corpus:"):
        fields = pattern[len("corpus:"):].split("|")
        if len(fields) != 5:
            return (_literal(pattern),)
        fmt, verb, element, refusal, prop = fields
        return tuple(_literal(part) for part in (prop, element, verb, refusal, fmt))
    fmt, element, prop = key_parts(pattern)
    return (int(prop is not None), _literal(prop or ""), _literal(element), _literal(fmt))


def key_parts(pattern: str) -> tuple[str, str, str | None]:
    """A delta or tier key's format, element and property (None when the
    key names no property)."""
    body = pattern[len("delta:"):] if pattern.startswith("delta:") else pattern
    fmt, _, rest = body.partition(":")
    element, dot, prop = rest.partition(".")
    return fmt, element, prop if dot else None


def names_element(pattern: str) -> bool:
    """Whether a key's element part is literal (``delta:docx:comment.*``,
    not ``delta:docx:*.*``)."""
    return not any(char in key_parts(pattern)[1] for char in "*?[")


class Patterns:
    """``fnmatch`` key patterns with a payload each. The most specific
    matching pattern wins; equal specificity goes to the earlier entry.
    Every winning pattern is recorded, so the unused ones can be reported
    as stale."""

    def __init__(self, entries: list[tuple[str, dict]], what: str):
        seen: set[str] = set()
        for pattern, _ in entries:
            if pattern in seen:
                raise SystemExit(f"{what}: key `{pattern}` is listed twice")
            seen.add(pattern)
        self.entries = entries
        self.used: set[str] = set()

    def best(self, key: str, exclude: str | None = None, mark: bool = True, where=None):
        """The winning ``(pattern, payload)`` for ``key``, or None. With
        ``exclude``, a pattern that also matches ``exclude`` is skipped (a
        property key then sees only patterns written for properties); with
        ``where``, only patterns it accepts compete."""
        found = None
        for index, (pattern, payload) in enumerate(self.entries):
            if not fnmatchcase(key, pattern):
                continue
            if exclude is not None and fnmatchcase(exclude, pattern):
                continue
            if where is not None and not where(pattern):
                continue
            rank = (*specificity(pattern), -index)
            if found is None or rank > found[0]:
                found = (rank, pattern, payload)
        if found is None:
            return None
        if mark:
            self.used.add(found[1])
        return found[1], found[2]

    def unused(self, prefix: str | None = None, exclude_prefix: str | None = None) -> list[str]:
        return [
            pattern
            for pattern, _ in self.entries
            if pattern not in self.used
            and (prefix is None or pattern.startswith(prefix))
            and (exclude_prefix is None or not pattern.startswith(exclude_prefix))
        ]


@dataclass(frozen=True)
class Resolution:
    """The disposition of one gap row, or of one corpus gap class."""

    disposition: str | None  # a DISPOSITION_VALUES member, `deferred` unresolved, None: none
    tier: str  # the effective tier, "" when not tiered
    ref: str = ""
    entry: str = ""
    reason: str = ""
    via: str = ""  # where an inherited disposition comes from

    @property
    def accounted(self) -> bool:
        return self.disposition in ACCOUNTED

    def cell(self) -> str:
        """The disposition as a table cell: value, id, entry."""
        if self.disposition is None:
            return "**none**"
        parts = [self.disposition]
        if self.ref:
            parts.append(self.ref)
        if self.entry:
            parts.append(f"({self.entry})")
        return " ".join(parts)


NONE = Resolution(None, "")


class Dispositions:
    """The curated dispositions file: design §13 tiers by OfficeCLI element
    and property, and the disposition entries keyed by row."""

    def __init__(self, data: dict, origin: str = str(DISPOSITIONS)):
        self.origin = origin
        self.data = data
        policy = data.get("policy", {})
        self.p0_formats = tuple(policy.get("p0Formats", ("docx", "xlsx", "pptx")))
        self.breadth_tier = policy.get("propertyBreadthTier", "P1")
        if self.breadth_tier not in TIER_VALUES:
            raise SystemExit(f"{origin}: propertyBreadthTier `{self.breadth_tier}` is not a tier")
        tiers = []
        for entry in data.get("tiers", []):
            if entry.get("tier") not in TIER_VALUES:
                raise SystemExit(f"{origin}: tier entry {entry.get('keys')} has tier `{entry.get('tier')}`")
            if not entry.get("basis"):
                raise SystemExit(f"{origin}: tier entry {entry.get('keys')} has no basis")
            tiers += [(key, entry) for key in entry["keys"]]
        self.tiers = Patterns(tiers, f"{origin} tiers")
        entries = []
        ids: set[str] = set()
        for entry in data.get("dispositions", []):
            label = entry.get("id", "")
            if not label or label in ids:
                raise SystemExit(f"{origin}: disposition id `{label}` is empty or repeated")
            ids.add(label)
            value = entry.get("disposition")
            if value not in DISPOSITION_INPUTS:
                raise SystemExit(f"{origin}: `{label}` has disposition `{value}`")
            ref = entry.get("ref", "")
            if value == "planned" and not REF.fullmatch(ref.split(" ", 1)[0]):
                raise SystemExit(f"{origin}: planned `{label}` needs a work-package or record id, not `{ref}`")
            if ref and not REF.match(ref):
                raise SystemExit(f"{origin}: `{label}` ref `{ref}` is not a work-package or record id")
            if value in ("wontfix", "out-of-scope") and not entry.get("reason"):
                raise SystemExit(f"{origin}: {value} `{label}` needs a reason")
            if value != "planned" and not ref and not entry.get("reason"):
                raise SystemExit(f"{origin}: `{label}` needs a reason or a ref")
            entries += [(key, entry) for key in entry.get("keys", [])]
        self.entries = Patterns(entries, f"{origin} dispositions")
        self.scripts = Patterns(
            [(key, entry) for entry in data.get("corpusScripts", []) for key in entry["keys"]],
            f"{origin} corpusScripts",
        )
        for entry in data.get("corpusScripts", []):
            if entry.get("tier") not in TIER_VALUES or not entry.get("basis"):
                raise SystemExit(f"{origin}: corpus script entry {entry.get('keys')} needs a tier and a basis")
        self.by_id = {entry["id"]: entry for entry in data.get("dispositions", [])}
        self.follow_ups = []
        follow_ids: set[str] = set()
        for entry in data.get("followUps", []):
            label = entry.get("id", "")
            if not label or label in follow_ids:
                raise SystemExit(f"{origin}: follow-up map id `{label}` is empty or repeated")
            follow_ids.add(label)
            ref = entry.get("ref", "")
            if not REF.fullmatch(ref.split("/", 1)[0]) or milestone_only(ref):
                raise SystemExit(f"{origin}: follow-up map `{label}` needs a work-package or record id, not `{ref}`")
            if not entry.get("followUps") or not entry.get("note"):
                raise SystemExit(f"{origin}: follow-up map `{label}` needs `followUps` and a `note`")
            self.follow_ups.append(entry)
        self.follow_up_uses: dict[str, int] = {}
        # (format, follow-up) of every milestone-only golden follow-up, set
        # by `classified`.
        self.milestone_follow_ups: list[tuple[str, str]] = []

    @classmethod
    def load(cls, path: Path) -> "Dispositions":
        return cls(json.loads(path.read_text(encoding="utf-8")), path.as_posix())

    def effective(self, fmt: str, tier: str) -> str:
        """Design §13's v1 boundary binds the formats the goal's definition of
        done names (``p0Formats``); elsewhere a P0 or v1 feature falls in the
        P1 wave."""
        if tier in V1_TIERS and fmt not in self.p0_formats:
            return "P1"
        return tier

    def tier(self, fmt: str, element: str, prop: str | None = None, mark: bool = True) -> str | None:
        """The effective tier of an OfficeCLI element (or of one of its
        properties) in ``fmt``. Element keys are ``<vocabulary>:<element>``;
        property keys ``<vocabulary>:<element>.<property>``. A property
        without its own tier takes a P1 or P2 element's tier; a property of a
        P0 or v1 element is formatting breadth (``propertyBreadthTier``). A
        property's own tier never comes before its element's."""
        vocab = reference_format(fmt)
        element_key = f"{vocab}:{element}"
        hit = self.tiers.best(element_key, mark=mark)
        if hit is None:
            return None
        tier = hit[1]["tier"]
        if prop is not None:
            own = self.tiers.best(f"{element_key}.{prop}", exclude=element_key, mark=mark)
            if own is not None:
                # A property is never in an earlier tier than its element.
                tier = max(tier, own[1]["tier"], key=TIER_VALUES.index)
            elif tier in V1_TIERS:
                tier = self.breadth_tier
        return self.effective(fmt, tier)

    def keyed_tier(self, key: str, mark: bool = True) -> str | None:
        """The tier of a key outside the OfficeCLI vocabulary (corpus classes)."""
        hit = self.tiers.best(key, mark=mark)
        return hit[1]["tier"] if hit else None

    def resolve(self, key: str, tier: str, exclude: str | None = None, mark: bool = True,
                where=None) -> Resolution | None:
        """The winning entry's disposition for ``key``, or None."""
        hit = self.entries.best(key, exclude=exclude, mark=mark, where=where)
        if hit is None:
            return None
        entry = hit[1]
        value = entry["disposition"]
        if value == "deferred" and tier in ("P1", "P2"):
            value = f"deferred-{tier}"
        return Resolution(value, tier, entry.get("ref", ""), entry["id"], entry.get("reason", ""))

    def follow_up_entry(self, fmt: str, text: str, mark: bool = True) -> dict | None:
        """The follow-up map entry (record 0344, audit WP13) that gives a
        milestone-only follow-up of ``fmt`` its current id: an entry that
        names the format beats one that does not, then the pattern with
        the most literal characters, then the earlier entry."""
        best, rank = None, None
        for order, entry in enumerate(self.follow_ups):
            fmts = entry.get("formats")
            if fmts and fmt not in fmts:
                continue
            for pattern in entry["followUps"]:
                if fnmatchcase(text, pattern):
                    key = (bool(fmts), _literal(pattern), -order)
                    if rank is None or key > rank:
                        best, rank = entry, key
        if best is not None and mark:
            self.follow_up_uses[best["id"]] = self.follow_up_uses.get(best["id"], 0) + 1
        return best

    def current(self, fmt: str, text: str) -> str:
        """A follow-up as the tables show it: a milestone-only one gains its
        current id (`M2 (validate) → 0336`)."""
        if not text or not milestone_only(text):
            return text
        entry = self.follow_up_entry(fmt, text, mark=False)
        return f"{text} → {entry['ref']}" if entry else text

    def script_tier(self, script: str) -> str | None:
        hit = self.scripts.best(script)
        return hit[1]["tier"] if hit else None


def resolution_problems(label: str, resolution: Resolution, follow_up: str = "") -> list[str]:
    """What ``--check`` refuses about one gap row's disposition."""
    if resolution.disposition is None:
        return [f"{label}: no disposition"]
    problems = []
    tier = resolution.tier
    value = resolution.disposition
    if tier in V1_TIERS and (value.startswith("deferred") or value == "wontfix"):
        problems.append(f"{label}: a {tier} row is `{value}` ({resolution.entry}); P0 is supported or planned")
    elif value == "deferred":
        problems.append(f"{label}: `deferred` without a P1/P2 tier ({resolution.entry})")
    elif value.startswith("deferred-") and value != f"deferred-{tier}":
        problems.append(f"{label}: `{value}` on a {tier or 'untiered'} row ({resolution.entry})")
    # A wontfix or out-of-scope reason supersedes the follow-up.
    if (follow_up and milestone_only(follow_up) and not resolution.ref
            and value not in ("wontfix", "out-of-scope")):
        problems.append(f"{label}: milestone-only follow-up `{follow_up}` needs a current id ({resolution.entry})")
    return problems


def element_key(fmt: str, element: str) -> str:
    return f"delta:{fmt}:{element}"


def property_key(fmt: str, element: str, key: str) -> str:
    return f"delta:{fmt}:{element}.{key}"


def property_is_gap(row: tuple) -> bool:
    return row[READ_STATUS] in READ_GAPS or row[WRITE_STATUS] in WRITE_GAPS


def apply_dispositions(delta: "FormatDelta", dispositions: Dispositions) -> list[str]:
    """Resolve every gap row of one format; a classified ``missing`` row
    becomes ``unlisted``. Returns what ``--check`` refuses."""
    fmt = delta.name
    problems: list[str] = []
    delta.element_dispositions = [None] * len(delta.elements)
    by_element: dict[str, Resolution] = {}
    for index, (name, litchi, status, follow_up) in enumerate(delta.elements):
        if not name or status not in ELEMENT_GAPS:
            continue
        tier = dispositions.tier(fmt, name)
        if tier is None:
            problems.append(f"{fmt} element {name}: no tier")
            tier = ""
        resolution = dispositions.resolve(element_key(fmt, name), tier) or Resolution(None, tier)
        delta.element_dispositions[index] = resolution
        by_element[name] = resolution
        problems += resolution_problems(f"{fmt} element {name}", resolution, follow_up)
        if status == "missing" and resolution.disposition is not None:
            delta.elements[index] = (name, litchi, TRIAGED, follow_up)
    delta.property_dispositions = [None] * len(delta.properties)
    for index, row in enumerate(delta.properties):
        if not property_is_gap(row):
            continue
        name, key = row[0], row[1]
        label = f"{fmt} {name}.{key}"
        parent = by_element.get(name)
        own_key = property_key(fmt, name, key)
        resolution = None
        if parent is not None:
            # A gap element's properties share its disposition unless an
            # entry names the element and the property.
            only = dict(exclude=element_key(fmt, name), where=names_element)
            if dispositions.entries.best(own_key, mark=False, **only):
                tier = dispositions.tier(fmt, name, key) or parent.tier
                resolution = dispositions.resolve(own_key, tier, **only)
            else:
                resolution = Resolution(parent.disposition, parent.tier, parent.ref, parent.entry,
                                        parent.reason, via="element")
        else:
            tier = dispositions.tier(fmt, name, key)
            if tier is None:
                problems.append(f"{label}: no tier")
                tier = ""
            resolution = dispositions.resolve(own_key, tier) or Resolution(None, tier)
        delta.property_dispositions[index] = resolution
        if resolution.via != "element":
            problems += resolution_problems(label, resolution, row[5])
        if row[READ_STATUS] == "missing" and resolution.disposition is not None:
            delta.properties[index] = (*row[:READ_STATUS], TRIAGED, *row[READ_STATUS + 1:])
    return problems


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
    pinned = root / UPSTREAM_COMMIT
    if pinned.is_file():
        return pinned.read_text(encoding="utf-8").strip()[:8] or "unknown"
    if not (root / ".git").exists():
        # `git -C` would otherwise report an enclosing repository's commit.
        return "unknown"
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

    def lookup_order(self) -> list[str]:
        """The names to look a Litchi element up by: the element's own name
        and file stem first, then its aliases and path words, sorted. A path
        word (`shape` in `/slidemaster[N]/shape[K]`) must not win over the
        element's own name (record 0339)."""
        own = [normalize(self.name), normalize(self.stem)]
        return list(dict.fromkeys([*own, *sorted(self.names() - set(own))]))


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
    properties: list[tuple[str, str, str, str, str, str, str, str]] = field(default_factory=list)
    operation_gaps: list[tuple[str, str, str, str, str]] = field(default_factory=list)
    # Parallel to `elements` and `properties` once dispositions are applied:
    # the gap rows' resolutions, None for rows that are not gaps.
    element_dispositions: list | None = None
    property_dispositions: list | None = None

    def counts(self, rows, statuses, column):
        counts = {status: 0 for status in statuses}
        for row in rows:
            counts[row[column]] += 1
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


def element_gaps(report: dict) -> dict[str, dict]:
    """Gaps for elements the format has but its table does not declare
    (subject ``element`` in the export's top-level ``gaps``), by name."""
    return {
        normalize(gap["element"]): gap
        for gap in report.get("gaps", [])
        if gap.get("subject") == "element"
    }


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


def write_status(officecli: dict, litchi: dict | None) -> str:
    """Litchi's add/set flags against OfficeCLI's for one property.

    ``supported`` when Litchi offers every write verb OfficeCLI offers,
    ``partial`` when it offers some, ``gap`` when it offers none, and ``—``
    when OfficeCLI offers no write verb (``litchi-only`` if Litchi does).
    """
    wanted = [verb for verb in ("add", "set") if officecli.get(verb)]
    litchi = litchi or {}
    offered = [verb for verb in ("add", "set") if litchi.get(verb)]
    if not wanted:
        return "litchi-only" if offered else "—"
    have = [verb for verb in wanted if verb in offered]
    if len(have) == len(wanted):
        return "supported"
    return "partial" if have else "gap"


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
    if "pending" in report:
        delta.pending = report["pending"]
        for schema in officecli:
            reason = OUT_OF_SCOPE.get(schema.stem) or OUT_OF_SCOPE.get(schema.name)
            if reason:
                delta.elements.append((schema.name, "", "out of scope", reason))
                continue
            delta.elements.append(
                (schema.name, "", "pending", delta.pending.get("followUp", ""))
            )
            for key, spec in schema.properties.items():
                write = "pending" if spec.get("add") or spec.get("set") else "—"
                delta.properties.append(
                    (schema.name, key, spec.get("type", ""), ops_letters(spec), "pending", "", "", write)
                )
        return delta

    litchi = litchi_elements(report)
    index: dict[str, LitchiElement] = {}
    for element in litchi:
        for name in element.names():
            index.setdefault(name, element)
    property_gaps: dict[tuple[str, str], dict] = {}
    for element in report.get("elements", []):
        for key, gap in element.get("propertyGaps", {}).items():
            property_gaps[(normalize(element["element"]), normalize(key))] = gap
    undeclared = element_gaps(report)

    matched: set[str] = set()
    folded = FOLDED.get(reference_format(fmt), {})
    for schema in officecli:
        target = next((index[name] for name in schema.lookup_order() if name in index), None)
        note = ""
        if schema.stem in OUT_OF_SCOPE or schema.name in OUT_OF_SCOPE:
            reason = OUT_OF_SCOPE.get(schema.stem) or OUT_OF_SCOPE[schema.name]
            delta.elements.append((schema.name, "", "out of scope", reason))
            continue
        if schema.name in UNMATCHED.get(reference_format(fmt), set()):
            target = None
        if target is None and schema.stem in folded:
            into, why = folded[schema.stem]
            target = index.get(normalize(into))
            note = f"{why}; in `{into}`"
            if target is None:
                raise SystemExit(f"{fmt}: {schema.stem} is folded into {into}, which the report lacks")
            status = "folded"
        elif target is None:
            gap = next((undeclared[name] for name in schema.lookup_order() if name in undeclared), None)
            if gap is not None:
                # Listed in `capabilities()` as an element gap without a table entry.
                delta.elements.append((schema.name, "", "gap", gap["followUp"]))
                for key, spec in schema.properties.items():
                    write = "gap" if spec.get("add") or spec.get("set") else "—"
                    delta.properties.append(
                        (schema.name, key, spec.get("type", ""), ops_letters(spec), "element gap",
                         gap["followUp"], "", write)
                    )
                continue
            delta.elements.append((schema.name, "", "missing", ""))
            for key, spec in schema.properties.items():
                status = "missing" if spec.get("get") else "write-only"
                delta.properties.append(
                    (schema.name, key, spec.get("type", ""), ops_letters(spec), status, "",
                     "", write_status(spec, None))
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
            # A `<N>` key (`level<N>.format`, `startOverride.<N>`) matches the
            # registry's numbered spellings (`level0.format`, ...).
            numbered = [
                litchi_keys[key]
                for name in names
                if "<n>" in name
                for key in sorted(litchi_keys)
                if re.fullmatch(re.escape(name).replace("<n>", r"\d+"), key)
            ]
            if hit is None and numbered:
                hit = numbered[0]
                seen_litchi.update(numbered)
            gap = next(
                (property_gaps[(normalize(target.name), name)] for name in names
                 if (normalize(target.name), name) in property_gaps),
                None,
            )
            litchi_spec = target.properties[hit] if hit is not None else None
            if (reference_format(fmt), schema.name, key) in {("docx", "paragraph", "paraId"), ("docx", "sdt", "id"), ("xlsx", "cell", "numFmtId")}:
                row_status, row_follow = "raw-only", "ADR 0034: native selection requires explicit raw admission"
            elif target.gap:
                row_status, row_follow = "element gap", follow(target.gap)
            elif hit is not None:
                seen_litchi.add(hit)
                # A property Litchi registers only for add/set is written, not read.
                row_status = "supported" if litchi_spec.get("get") else "write-only"
                row_follow = ""
            elif any(name in map(normalize, target.predicate_keys) for name in names):
                row_status, row_follow = "supported", "the node's text"
            elif gap is not None:
                row_status, row_follow = "gap", gap["followUp"]
            elif not spec.get("get"):
                row_status, row_follow = "write-only", ""
            else:
                row_status, row_follow = "missing", ""
            write = "gap" if target.gap and (spec.get("add") or spec.get("set")) else write_status(spec, litchi_spec)
            delta.properties.append(
                (schema.name, key, spec.get("type", ""), ops_letters(spec), row_status, row_follow,
                 ops_letters(litchi_spec) if litchi_spec else "", write)
            )
        if status != "folded":
            for key, spec in target.properties.items():
                if key not in seen_litchi:
                    delta.properties.append(
                        (schema.name, key, spec.get("type", ""), "", "litchi-only", "",
                         ops_letters(spec), write_status({}, spec))
                    )

    for element in litchi:
        if element.name not in matched and not any(
            row[1] == element.name for row in delta.elements
        ):
            delta.elements.append(("", element.name, "litchi-only", follow(element.gap)))
    return delta


def disposition_counts(rows: list) -> dict[str, int]:
    """Gap rows per disposition (``none`` for an unclassified one)."""
    counts = {value: 0 for value in (*DISPOSITION_VALUES, "none")}
    for resolution in rows or []:
        if resolution is None:
            continue
        value = resolution.disposition
        counts[value if value in counts else "none"] += 1
    return counts


def disposition_cell(resolutions: list | None, index: int) -> str:
    if resolutions is None or resolutions[index] is None:
        return ""
    return resolutions[index].cell()


def tier_cell(resolutions: list | None, index: int) -> str:
    if resolutions is None or resolutions[index] is None:
        return ""
    return resolutions[index].tier


def property_page(fmt: str) -> Path:
    return PROPERTY_DIR / f"{fmt}.md"


def render(
    deltas: list[FormatDelta],
    goldens: list[tuple[str, str]],
    tree: tuple[str, int],
    commit: str,
    dispositions: Dispositions | None = None,
) -> str:
    """The main table: inputs, statuses, summaries, dispositions, and each
    format's elements and contextual operation gaps. Property tables are
    separate pages (``render_properties``)."""
    lines = [
        "# OfficeCLI capability delta (generated)",
        "",
        "Generated by `tools/officecli_capability_delta.py`; do not edit by hand.",
        "It compares OfficeCLI's help schemas with Litchi's property registry",
        "export (`litchi_ops::schema(format)`, golden snapshots under",
        "`crates/litchi-ops/tests/golden/registry/`, record 0263), for every",
        "format the registry covers. Each property has a read status (Litchi's",
        "`get` against OfficeCLI's) and a write status (Litchi's `add`/`set`",
        "against OfficeCLI's); both sides' add/set/get flags are shown as `asg`.",
        "DOC and RTF are compared against the DOCX help vocabulary, XLS and XLSB",
        "against XLSX, and PPT against PPTX, as Litchi's parity target; this does",
        "not claim OfficeCLI itself supports those files.",
        "Operation gaps also include a format's feature-specific encryption",
        "snapshot when it is present.",
        "Every gap row carries a disposition from",
        f"`{DISPOSITIONS.as_posix()}` (record 0332). Each format's property table is",
        f"a separate page under `{PROPERTY_DIR.as_posix()}/`.",
        "",
        "Rerun the script after blessing the registry snapshots (`LITCHI_BLESS=1`);",
        "`tools/test_officecli_capability_delta.py` fails while a snapshot",
        "differs from the hash recorded below, and `--check` fails while a gap",
        "row has no disposition.",
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
        "Read statuses:",
        "",
        "| Status | Meaning |",
        "|---|---|",
        "| supported | Litchi reads it (a property, or the node's text for a predicate key such as a cell's `value`) |",
        "| gap | Litchi lists it in `capabilities()` as a gap (an element gap may have no table entry), with the follow-up shown |",
        "| element gap | a property of an element Litchi lists as a gap |",
        "| folded | an OfficeCLI element Litchi models inside another element (named in the follow-up) |",
        "| pending | the format's adapter has not landed (the follow-up names the step that lands it) |",
        "| write-only | not read: OfficeCLI has no `get`, or Litchi registers it for add/set only; see the write status |",
        "| raw-only | native selection metadata excluded from ordinary APIs by ADR 0034 |",
        "| operation gap | an element operation refused for a particular selection or leaf context; listed in the contextual gap table |",
        "| missing | OfficeCLI reads it and Litchi neither reads nor lists it, and no disposition classifies it: untriaged (`--check` fails) |",
        "| unlisted | OfficeCLI reads it and Litchi neither reads nor lists it in `capabilities()`; its disposition records the decision |",
        "| out of scope | not an element of the unified read surface |",
        "| litchi-only | Litchi reads it and OfficeCLI's schema lacks it |",
        "",
        "Write statuses (OfficeCLI's `add`/`set` flags against Litchi's):",
        "",
        "| Status | Meaning |",
        "|---|---|",
        "| supported | Litchi offers every write verb OfficeCLI offers for it |",
        "| partial | Litchi offers some of OfficeCLI's write verbs (compare the two `asg` columns) |",
        "| gap | OfficeCLI writes it and Litchi does not (including properties of gap or missing elements) |",
        "| pending | the format's adapter has not landed |",
        "| litchi-only | Litchi writes it and OfficeCLI offers no write verb, or its schema lacks it |",
        "| — | neither side writes it |",
        "",
        "## Summary",
        "",
        "Elements:",
        "",
        "| Format | " + " | ".join(ELEMENT_STATUSES) + " |",
        "|---|" + "---|" * len(ELEMENT_STATUSES),
    ]
    for delta in deltas:
        counts = delta.counts(delta.elements, ELEMENT_STATUSES, 2)
        lines.append(f"| {delta.name} | " + " | ".join(str(counts[s]) for s in ELEMENT_STATUSES) + " |")
    lines += [
        "",
        "Property reads:",
        "",
        "| Format | " + " | ".join(PROPERTY_STATUSES) + " |",
        "|---|" + "---|" * len(PROPERTY_STATUSES),
    ]
    for delta in deltas:
        counts = delta.counts(delta.properties, PROPERTY_STATUSES, READ_STATUS)
        lines.append(f"| {delta.name} | " + " | ".join(str(counts[s]) for s in PROPERTY_STATUSES) + " |")
    lines += [
        "",
        "Property writes:",
        "",
        "| Format | " + " | ".join(WRITE_STATUSES) + " |",
        "|---|" + "---|" * len(WRITE_STATUSES),
    ]
    for delta in deltas:
        counts = delta.counts(delta.properties, WRITE_STATUSES, WRITE_STATUS)
        lines.append(f"| {delta.name} | " + " | ".join(str(counts[s]) for s in WRITE_STATUSES) + " |")
    if dispositions is not None:
        lines += render_disposition_summary(deltas, dispositions)

    for delta in deltas:
        lines += ["", f"## {delta.name}", ""]
        if delta.pending is not None:
            lines += [
                f"The {delta.name} adapter has not landed: `get` is "
                f"`{delta.pending.get('code', '')}` ({delta.pending.get('reason', '')}; follow-up "
                f"{delta.pending.get('followUp', '')}). OfficeCLI's elements, with their property counts:",
                "",
                "| OfficeCLI element | Properties | Status | Tier | Disposition | Follow-up |",
                "|---|---|---|---|---|---|",
            ]
            for index, (name, _, status, follow_up) in enumerate(delta.elements):
                count = sum(1 for row in delta.properties if row[0] == name)
                lines.append(
                    f"| {name} | {count} | {status} | {tier_cell(delta.element_dispositions, index)} | "
                    f"{disposition_cell(delta.element_dispositions, index)} | {follow_up} |"
                )
        else:
            lines += [
                "### Elements",
                "",
                "| OfficeCLI element | Litchi element | Status | Tier | Disposition | Follow-up |",
                "|---|---|---|---|---|---|",
            ]
            for index, (name, litchi, status, follow_up) in enumerate(delta.elements):
                lines.append(
                    f"| {name} | {litchi} | {status} | {tier_cell(delta.element_dispositions, index)} | "
                    f"{disposition_cell(delta.element_dispositions, index)} | {follow_up} |"
                )
        if delta.properties:
            page = property_page(delta.name)
            lines += [
                "",
                f"Properties: {len(delta.properties)} rows, "
                f"{sum(1 for row in delta.properties if property_is_gap(row))} of them gaps, in "
                f"[`{page.name}`]({PROPERTY_DIR.name}/{page.name}).",
            ]
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


def render_disposition_summary(deltas: list[FormatDelta], dispositions: Dispositions) -> list[str]:
    columns = (*DISPOSITION_VALUES, "none")
    lines = [
        "",
        "## Dispositions",
        "",
        f"Every gap row (an element that is not supported or folded; a property whose read or",
        f"write status is a gap) is classified by `{DISPOSITIONS.as_posix()}`:",
        "",
        "| Disposition | Meaning |",
        "|---|---|",
        "| planned | to be closed; the id names the work package (`WPn`, coordinator audit 2026-10-17) or record that owns it |",
        "| deferred-P1 | after v1: design §13 tier P1 (or a P0/v1 feature of a format outside the v1 boundary) |",
        "| deferred-P2 | explicitly later: design §13 tier P2 |",
        "| wontfix | a deliberate difference from OfficeCLI, with the reason |",
        "| out-of-scope | not part of the unified surface (rendering, raw markup, an add-time format), with the reason |",
        "",
        "Tiers come from the same file (design §13, keyed by OfficeCLI element and property). "
        f"The v1 boundary binds {', '.join(f'`{fmt}`' for fmt in dispositions.p0_formats)} "
        "(the goal's definition of done: \"docx/xlsx/pptx additionally cover the P0 surface\"); "
        "in the other formats a P0 or v1 feature is P1. A property of a P0 or v1 element "
        f"is formatting breadth, tier {dispositions.breadth_tier}, unless the file names it. "
        "P0 and v1 rows are supported or planned, never deferred or wontfix; a deferred row's "
        "tier is its row's tier; a row whose follow-up names only a closed milestone carries a "
        "current work-package or record id (audit WP13). `--check` enforces all of this.",
        "",
        "Element gap rows:",
        "",
        "| Format | " + " | ".join(columns) + " |",
        "|---|" + "---:|" * len(columns),
    ]
    for delta in deltas:
        counts = disposition_counts(delta.element_dispositions)
        lines.append(f"| {delta.name} | " + " | ".join(str(counts[c]) for c in columns) + " |")
    lines += [
        "",
        "Property gap rows:",
        "",
        "| Format | " + " | ".join(columns) + " |",
        "|---|" + "---:|" * len(columns),
    ]
    for delta in deltas:
        counts = disposition_counts(delta.property_dispositions)
        lines.append(f"| {delta.name} | " + " | ".join(str(counts[c]) for c in columns) + " |")
    by_tier: dict[tuple[str, str], int] = {}
    for delta in deltas:
        for resolution in [*(delta.element_dispositions or []), *(delta.property_dispositions or [])]:
            if resolution is not None:
                key = (resolution.tier or "—", resolution.disposition or "none")
                by_tier[key] = by_tier.get(key, 0) + 1
    tiers = [tier for tier in (*TIER_VALUES, "—") if any(key[0] == tier for key in by_tier)]
    lines += [
        "",
        "All gap rows by tier:",
        "",
        "| Tier | " + " | ".join(columns) + " |",
        "|---|" + "---:|" * len(columns),
    ]
    for tier in tiers:
        lines.append(f"| {tier} | " + " | ".join(str(by_tier.get((tier, c), 0)) for c in columns) + " |")
    uses: dict[str, int] = {}
    for delta in deltas:
        for resolution in [*(delta.element_dispositions or []), *(delta.property_dispositions or [])]:
            if resolution is not None and resolution.entry:
                uses[resolution.entry] = uses.get(resolution.entry, 0) + 1
    lines += [
        "",
        "Entries (rows counts this table's element and property rows; a gap element's",
        "properties share its entry unless an entry names the property):",
        "",
        "| Entry | Disposition | Id | Rows | Keys | Reason |",
        "|---|---|---|---:|---|---|",
    ]
    for entry_id, entry in dispositions.by_id.items():
        keys = [key for key in entry.get("keys", []) if key.startswith("delta:")]
        if not keys:
            continue
        shown = ", ".join(f"`{key}`" for key in keys)
        lines.append(
            f"| {entry_id} | {entry['disposition']} | {entry.get('ref', '')} | {uses.get(entry_id, 0)} | "
            f"{shown} | {entry.get('reason', '')} |"
        )
    lines += follow_up_lines(dispositions)
    return lines


def follow_up_lines(dispositions: Dispositions) -> list[str]:
    """The milestone follow-up map (record 0344, audit WP13)."""
    counts: dict[str, int] = {}
    formats_of: dict[str, set[str]] = {}
    for fmt, text in dispositions.milestone_follow_ups:
        entry = dispositions.follow_up_entry(fmt, text, mark=False)
        if entry is not None:
            counts[entry["id"]] = counts.get(entry["id"], 0) + 1
            formats_of.setdefault(entry["id"], set()).add(fmt)
    total = len(dispositions.milestone_follow_ups)
    distinct = len(set(text for _, text in dispositions.milestone_follow_ups))
    lines = [
        "",
        "## Milestone follow-ups",
        "",
        f"{total} follow-ups in the registry and capability goldens ({distinct} distinct texts) "
        "name only a milestone or a step of one (`M2 (validate)`, `M1-extended (R4x)`). "
        f"`{DISPOSITIONS.as_posix()}` (`followUps`) maps each to the record or work package that owns it "
        "now; the tables above show the id after the text (`M2 (validate) → 0336`). `--check` fails on "
        "a milestone-only follow-up without a current id and on a map entry that maps none "
        "(audit WP13, record 0344).",
        "",
        "| Map entry | Follow-ups | Formats | Golden follow-ups | Current id | Now |",
        "|---|---|---|---:|---|---|",
    ]
    for entry in dispositions.follow_ups:
        shown = ", ".join(f"`{pattern}`" for pattern in entry["followUps"])
        fmts = ", ".join(sorted(formats_of.get(entry["id"], set())))
        lines.append(f"| {entry['id']} | {shown} | {fmts} | {counts.get(entry['id'], 0)} | "
                     f"{entry['ref']} | {entry['note']} |")
    return lines


def render_properties(delta: FormatDelta) -> str:
    """One format's property table, with each gap row's tier and disposition."""
    lines = [
        f"# OfficeCLI capability delta: {delta.name} properties (generated)",
        "",
        "Generated by `tools/officecli_capability_delta.py`; do not edit by hand.",
        f"The statuses, tiers and dispositions are explained in [`{OUTPUT.name}`](../{OUTPUT.name}).",
        "",
        "| Element | Property | Type | OfficeCLI `asg` | Litchi `asg` | Read | Write | Tier | Disposition | Follow-up |",
        "|---|---|---|---|---|---|---|---|---|---|",
    ]
    for index, (element, key, ty, ops, status, follow_up, litchi_ops, write) in enumerate(delta.properties):
        lines.append(
            f"| {element} | `{key}` | {ty} | {ops} | {litchi_ops} | {status} | {write} | "
            f"{tier_cell(delta.property_dispositions, index)} | "
            f"{disposition_cell(delta.property_dispositions, index)} | {follow_up} |"
        )
    lines.append("")
    return "\n".join(lines)


@dataclass
class Generated:
    """The generated files and what ``--check`` refuses about the inputs."""

    files: dict[Path, str]
    problems: list[str]
    deltas: list[FormatDelta]


def compute_deltas(officecli: Path, repo: Path) -> tuple[list[FormatDelta], list[tuple[str, str]]]:
    """Every format's delta (without dispositions) and the golden hashes."""
    help_dir = officecli / "schemas" / "help"
    if not help_dir.is_dir():
        raise SystemExit(f"no OfficeCLI help schemas at {help_dir}; pass --officecli")
    deltas = []
    goldens = []
    for fmt in formats(repo):
        golden = repo / GOLDEN_DIR / f"{fmt}.json"
        data = golden.read_bytes()
        goldens.append(((GOLDEN_DIR / f"{fmt}.json").as_posix(), sha256(data)))
        report = json.loads(data)
        optional = GOLDEN_DIR / f"{fmt}{OPTIONAL_SUFFIX}.json"
        if (repo / optional).is_file():
            optional_data = (repo / optional).read_bytes()
            goldens.append((optional.as_posix(), sha256(optional_data)))
            report = merge_optional_operation_gaps(report, [json.loads(optional_data)])
        deltas.append(compare(fmt, officecli_elements(help_dir, fmt), report))
    return deltas, goldens


def golden_follow_ups(repo: Path) -> list[tuple[str, str, str]]:
    """Every ``followUp`` of the registry and capability goldens, as
    (format, follow-up, golden path); an ``-encryption`` snapshot counts for
    its format."""
    found = []
    for directory in (GOLDEN_DIR, CAPABILITY_DIR):
        for path in sorted((repo / directory).glob("*.json")):
            fmt = path.stem.removesuffix(OPTIONAL_SUFFIX)
            stack = [json.loads(path.read_text(encoding="utf-8"))]
            while stack:
                node = stack.pop()
                if isinstance(node, dict):
                    text = node.get("followUp")
                    if isinstance(text, str) and text:
                        found.append((fmt, text, (directory / path.name).as_posix()))
                    stack.extend(node.values())
                elif isinstance(node, list):
                    stack.extend(node)
    return found


def milestone_problems(follow_ups: list[tuple[str, str, str]], dispositions: Dispositions) -> list[str]:
    """Record 0344 (audit WP13): every follow-up of the registry and
    capability goldens that names only a milestone maps to a current id in
    the dispositions file's ``followUps``, and every map entry is used."""
    problems = []
    unmapped: dict[tuple[str, str], list[str]] = {}
    for fmt, text, origin in follow_ups:
        if milestone_only(text) and dispositions.follow_up_entry(fmt, text) is None:
            unmapped.setdefault((fmt, text), []).append(origin)
    for (fmt, text), origins in sorted(unmapped.items()):
        problems.append(f"{fmt}: milestone-only follow-up `{text}` ({len(origins)} in "
                        f"{', '.join(sorted(set(origins)))}) has no current id in `followUps`")
    for entry in dispositions.follow_ups:
        if not dispositions.follow_up_uses.get(entry["id"]):
            problems.append(f"stale follow-up map `{entry['id']}`: it maps no golden follow-up")
    return problems


def map_follow_ups(delta: FormatDelta, dispositions: Dispositions) -> None:
    """Show each milestone-only follow-up of a delta with its current id."""
    fmt = delta.name
    delta.elements = [(*row[:3], dispositions.current(fmt, row[3])) for row in delta.elements]
    delta.properties = [(*row[:5], dispositions.current(fmt, row[5]), *row[6:]) for row in delta.properties]
    delta.operation_gaps = [(*row[:4], dispositions.current(fmt, row[4])) for row in delta.operation_gaps]


def classified(officecli: Path, repo: Path, dispositions: Dispositions) -> tuple[list[FormatDelta], list[tuple[str, str]], list[str]]:
    """Every format's delta with its dispositions applied, the golden
    hashes, and the problems (including stale keys)."""
    deltas, goldens = compute_deltas(officecli, repo)
    problems: list[str] = []
    follow_ups = golden_follow_ups(repo)
    dispositions.milestone_follow_ups = [(fmt, text) for fmt, text, _ in follow_ups if milestone_only(text)]
    problems += milestone_problems(follow_ups, dispositions)
    for delta in deltas:
        map_follow_ups(delta, dispositions)
        problems += apply_dispositions(delta, dispositions)
    problems += [f"stale disposition key `{key}`: it classifies no row"
                 for key in dispositions.entries.unused(prefix="delta:")]
    problems += [f"stale tier key `{key}`: it classifies no row"
                 for key in dispositions.tiers.unused(exclude_prefix="corpus:")]
    return deltas, goldens, problems


def generate(officecli: Path, repo: Path, dispositions_path: Path | None = None) -> Generated:
    dispositions = Dispositions.load(dispositions_path or repo / DISPOSITIONS)
    deltas, goldens, problems = classified(officecli, repo, dispositions)
    help_dir = officecli / "schemas" / "help"
    files = {
        OUTPUT: render(deltas, goldens, schema_tree_hash(help_dir), officecli_commit(officecli), dispositions)
    }
    for delta in deltas:
        if delta.properties:
            files[property_page(delta.name)] = render_properties(delta)
    for path, text in files.items():
        if len(text.encode()) >= SIZE_LIMIT:
            raise SystemExit(f"{path} would reach 1 MiB; split it before committing")
    return Generated(files, problems, deltas)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "--officecli",
        type=Path,
        default=Path(os.environ.get("LITCHI_OFFICECLI", REPO / VENDORED_OFFICECLI)),
        help="the OfficeCLI tree (default: $LITCHI_OFFICECLI, then test-data/officecli)",
    )
    parser.add_argument("--root", type=Path, default=REPO,
                        help="where the generated files go (default: the repository)")
    parser.add_argument("--dispositions", type=Path, default=REPO / DISPOSITIONS)
    parser.add_argument(
        "--check",
        action="store_true",
        help="fail if the committed tables differ from fresh ones, or a gap lacks a valid disposition",
    )
    arguments = parser.parse_args(argv)
    generated = generate(arguments.officecli, REPO, arguments.dispositions)
    root = arguments.root
    status = 0
    for problem in generated.problems:
        print(f"disposition: {problem}", file=sys.stderr)
    if generated.problems:
        status = 1
    expected = {root / path for path in generated.files}
    if arguments.check:
        for path, text in generated.files.items():
            target = root / path
            current = target.read_text(encoding="utf-8") if target.exists() else ""
            if current != text:
                print(f"{target} is stale; rerun {Path(__file__).name}", file=sys.stderr)
                status = 1
        directory = root / PROPERTY_DIR
        if directory.is_dir():
            for extra in sorted(set(directory.glob("*.md")) - expected):
                print(f"{extra} is not generated any more; delete it", file=sys.stderr)
                status = 1
        return status
    (root / PROPERTY_DIR).mkdir(parents=True, exist_ok=True)
    for extra in sorted(set((root / PROPERTY_DIR).glob("*.md")) - expected):
        extra.unlink()
    for path, text in generated.files.items():
        (root / path).write_text(text, encoding="utf-8")
    return status


if __name__ == "__main__":
    sys.exit(main())
