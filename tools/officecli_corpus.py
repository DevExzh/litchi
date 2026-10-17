#!/usr/bin/env python3
"""Classify the OfficeCLI example-corpus outcomes and write the delta table
(unified-ops record 0265).

``crates/litchi-cli/tests/officecli_corpus.rs`` replays every OfficeCLI
example script vendored under ``test-data/officecli/examples/`` through the
``litchi`` binary and records one row per ``officecli`` call in
``crates/litchi-cli/tests/officecli_corpus/outcomes.tsv``. This script reads
those rows and classifies every call that did not pass:

- **delta**: a recorded, deliberate or upstream-caused difference (the rules
  in ``RULES`` with kind ``delta``);
- **gap**: Litchi lacks the verb, option, element or property. Element and
  property refusals are cross-checked against the registry export
  (``crates/litchi-ops/tests/golden/registry/<format>.json``, the golden of
  ``litchi_ops::schema``, which also feeds ``capabilities()``): a refusal of
  something the registry claims is a **bug**, not a gap;
- **cascade**: a path that does not exist because an earlier ``add`` or
  ``batch`` in the same script was refused;
- **warned**: the call applied with ``--unsupported warn`` and skipped the
  properties it lists (record 0269); each skipped property is cross-checked
  like a refused one;
- a call that still failed after ``--unsupported warn`` skipped properties
  (detail ``<property>;skipped:<properties>``, record 0271) is classified
  by its own failure first (deltas, cascades, element refusals); a
  ``missing_property`` that no rule explains is the named gap F3b, because
  the skipped properties are what it lacked; each skipped property is
  cross-checked like a refused one, so skipping a property the registry
  claims is still a bug, and any other unexplained failure stays a bug;
- **bug**: anything else.

Every delta, named gap and element/property group carries a disposition from
``tools/officecli_dispositions.json`` (record 0332, shared with
``tools/officecli_capability_delta.py``): a group whose element or property
OfficeCLI's help schema names inherits the capability delta's disposition
for that row unless a corpus entry names the property. The file also gives
every script a design §13 tier; the scripts of tier P0 or v1 form the
v1-scope subset (design §14), reported green when every line passes or
carries only wontfix, deferred or out-of-scope classes, red otherwise.

It writes ``docs/unified-ops/officecli-corpus-deltas.md``: the summary per
script family, per format and per script, the v1-scope subset, the deltas,
the gaps with their registry follow-ups and dispositions, the bugs, and the
record 0256 follow-ups the corpus hits. ``--check`` regenerates it in memory
and fails when the committed file differs, when any call is a panic or
``internal_error`` (the Rust test refuses those before they can be
recorded), when a class has no valid disposition (a P0 or v1 class is never
deferred or wontfix, a recorded delta is wontfix or out-of-scope), or when
a ``corpus:`` disposition or tier key or a script key matches nothing::

    LITCHI_BLESS=1 cargo test -p litchi-cli --test officecli_corpus
    python3 tools/officecli_corpus.py           # or --check

The output is deterministic: no timestamps, stable orders.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from tools import officecli_capability_delta as capability_delta  # noqa: E402

REPO = Path(__file__).resolve().parent.parent
OUTCOMES = Path("crates/litchi-cli/tests/officecli_corpus/outcomes.tsv")
SHIM = Path("crates/litchi-cli/tests/officecli_corpus/officecli")
REGISTRY = Path("crates/litchi-ops/tests/golden/registry")
VENDORED = Path("test-data/officecli")
HELP = VENDORED / "schemas" / "help"
UPSTREAM_COMMIT = Path("test-data/officecli/UPSTREAM_COMMIT")
OUTPUT = Path("docs/unified-ops/officecli-corpus-deltas.md")

FORMATS = ("docx", "xlsx", "pptx")
LITCHI_VERBS = {
    "create", "get", "query", "view", "set", "add", "remove", "move", "swap",
    "batch", "capabilities", "schema", "import", "evaluate", "recalculate", "merge",
    "refresh", "dump", "raw", "raw-set", "add-part", "validate",
}
# Separates a failed call's own property from the properties it skipped.
SKIPPED_SEPARATOR = ";skipped:"
CASCADE_CODES = {"path_not_found", "missing_argument", "unsupported_path", "out_of_range"}
FATAL = {"panic", "internal_error", "timeout", "lost"}

# OfficeCLI element words (`--type`, path steps) that name a registry
# element differently.
ELEMENT_WORDS = {
    "docx": {
        "p": "paragraph", "r": "run", "tbl": "table", "tr": "row", "tc": "cell",
        "sectpr": "section", "validation": "datavalidation",
    },
    "xlsx": {
        "validation": "datavalidation", "col": "column", "conditionalformatting":
        "conditionalformatting", "cf": "conditionalformatting",
    },
    "pptx": {"p": "paragraph", "r": "run", "3dmodel": "model3d", "column": "col"},
}
CELL = re.compile(r"^\$?[A-Z]{1,3}\$?[0-9]+$")
RANGE = re.compile(r"^\$?[A-Z]{1,3}\$?[0-9]+:\$?[A-Z]{1,3}\$?[0-9]+$")


@dataclass
class Call:
    """One row of the outcomes file."""

    script: str
    line: int
    verb: str
    fmt: str
    kind: str
    result: str
    detail: str
    target: str
    # The properties a call that failed after skipping skipped (the part of
    # the recorded detail after `;skipped:`, record 0271).
    after_skip: str = ""

    @property
    def cite(self) -> str:
        return f"{self.script}:{self.line}"

    @property
    def readback(self) -> bool:
        return self.line == 0

    @property
    def passed(self) -> bool:
        return self.result == "ok"

    @property
    def warned(self) -> bool:
        return self.result == "warned"

    @property
    def refused(self) -> bool:
        return self.result not in ("ok", "skip", "warned")

    @property
    def failed_after_skipping(self) -> bool:
        """A refused call whose ``--unsupported warn`` run skipped properties
        before the rest failed (record 0271)."""
        return self.refused and bool(self.after_skip)

    def skipped(self) -> list[str]:
        """The properties a warned call, or a call that failed after
        skipping, skipped. The outcome lists them in full (record 0339);
        ``<head>.{<rest>,...}`` stands for consecutive names sharing a first
        dotted segment."""
        if self.warned:
            detail = self.detail
        elif self.failed_after_skipping:
            detail = self.after_skip
        else:
            return []
        return expand_skipped(detail, self.cite)


def expand_skipped(detail: str, cite: str = "") -> list[str]:
    """The names of a recorded skipped list: top-level commas separate
    entries, and ``head.{a,b}`` is ``head.a,head.b``."""
    names: list[str] = []
    entry, depth = "", 0
    for char in detail + ",":
        if char == "," and depth == 0:
            if entry.startswith("+"):
                raise SystemExit(f"{cite}: a truncated skipped list (`{entry}`) predates record "
                                 "0339; re-record the corpus")
            head, brace, rest = entry.partition(".{")
            if brace and rest.endswith("}"):
                names += [f"{head}.{name}" for name in rest[:-1].split(",") if name]
            elif entry:
                names.append(entry)
            entry = ""
            continue
        depth += {"{": 1, "}": -1}.get(char, 0)
        entry += char
    return names


def parse(text: str) -> list[Call]:
    calls: list[Call] = []
    script = None
    for raw in text.splitlines():
        if not raw or raw.startswith("#"):
            continue
        if raw.startswith("@ "):
            script = raw[2:]
            continue
        fields = raw.split("\t")
        if len(fields) != 7 or script is None:
            raise SystemExit(f"malformed outcome row: {raw!r}")
        line, verb, fmt, kind, result, detail, target = (
            "" if value == "-" else value for value in fields
        )
        detail, _, after_skip = detail.partition(SKIPPED_SEPARATOR)
        calls.append(Call(script, int(line), verb, fmt, kind, result, detail, target, after_skip))
    return calls


def family(script: str) -> str:
    directory = script.rsplit("/", 1)[0] if "/" in script else ""
    return "ppt/templates" if directory.startswith("ppt/templates") else directory


def word(name: str) -> str:
    return capability_delta.normalize(name)


class Registry:
    """The registry export of one format, keyed by normalized names."""

    def __init__(self, fmt: str, report: dict, officecli: list):
        self.fmt = fmt
        self.elements = {}
        for element in report.get("elements", []):
            for name in [element["element"], *element.get("elementAliases", [])]:
                self.elements.setdefault(word(name), element)
        self.top_gaps = capability_delta.element_gaps(report)
        roots = report.get("elements", [])
        self.root = roots[0]["element"] if roots else ""
        # An element's own name wins over another element's alias.
        self.officecli = {}
        for element in officecli:
            for name in (word(element.stem), word(element.name)):
                self.officecli.setdefault(name, element)
        for element in officecli:
            for name in element.names():
                self.officecli.setdefault(name, element)

    def element_name(self, call: Call) -> str:
        """The registry element a call addresses, or the OfficeCLI word."""
        name = call.kind if call.verb == "add" else call.target
        if call.verb == "add" and call.kind == "(from)":
            return "(from)"
        if not name:
            return ""
        if name == "/":
            return self.root
        if self.fmt == "xlsx" and call.verb != "add":
            if CELL.match(name):
                return "cell"
            if RANGE.match(name):
                return "range"
            mapped = ELEMENT_WORDS["xlsx"].get(word(name), word(name))
            if mapped not in self.elements:
                return "sheet"  # `/SheetName`
        mapped = ELEMENT_WORDS.get(self.fmt, {}).get(word(name), word(name))
        element = self.elements.get(mapped)
        return element["element"] if element else name

    def element(self, name: str) -> dict | None:
        return self.elements.get(word(name))

    def property(self, element: dict, key: str) -> tuple[str, dict] | None:
        wanted = word(key)
        for name, spec in element.get("properties", {}).items():
            aliases = spec.get("aliases", [])
            if not isinstance(aliases, list):
                aliases = []
            if wanted in {word(name), *map(word, aliases)}:
                return name, spec
        return None

    def listed_on(self, key: str, verb: str) -> list[str]:
        """The elements whose registry entry can write the property."""
        op = "add" if verb == "add" else "set"
        names = []
        for element in {id(e): e for e in self.elements.values()}.values():
            found = self.property(element, key)
            if found and found[1].get(op) and element["element"] not in names:
                names.append(element["element"])
        return sorted(names)

    def documented(self, element_name: str, key: str) -> bool:
        """Whether OfficeCLI's help schema lists the property on the element."""
        element = self.officecli.get(word(element_name))
        if element is None:
            return False
        wanted = word(key)
        for name, spec in element.properties.items():
            aliases = spec.get("aliases", []) if isinstance(spec, dict) else []
            if wanted in {word(name), *map(word, aliases if isinstance(aliases, list) else [])}:
                return True
        return False


@dataclass
class Verdict:
    kind: str  # passed, warned, resident, delta, gap, cascade, bug
    rule: str  # rule id or a group key
    note: str = ""
    # A cascade's root: the index of the refused `add` or `batch` it follows.
    root: int | None = None


@dataclass
class Rule:
    """A recorded delta or a named gap."""

    id: str
    kind: str
    title: str
    officecli: str
    litchi: str

    def matches(self, call: Call, earlier: list[Call]) -> bool:  # pragma: no cover
        raise NotImplementedError


class When(Rule):
    def __init__(self, id, kind, title, officecli, litchi, test):
        super().__init__(id, kind, title, officecli, litchi)
        self.test = test

    def matches(self, call, earlier):
        return self.test(call, earlier)


def failed_add(earlier: list[Call], kinds: set[str] | None = None) -> bool:
    return any(
        previous.refused
        and previous.verb in ("add", "batch")
        and (kinds is None or previous.kind.lower() in kinds)
        for previous in earlier
    )


RULES: list[Rule] = [
    When(
        "D1", "delta", "Created-path form captured by the script",
        "`add` prints the new paragraph as `/body/p[@paraId=...]`, which the script greps "
        "to address it later",
        "`add` prints the positional path `/body/p[N]` (record 0256 §3), so the "
        "capture is empty and the dependent `set` has no path (`missing_argument`)",
        lambda c, e: c.script == "word/revisions.sh" and c.result == "missing_argument",
    ),
    When(
        "D1b", "delta", "Created numbering-instance path captured by the script",
        "`add --type num` prints `/numbering/num[@id=N]`, which the script greps for "
        "the new `numId`",
        "`add` prints the positional path `/numbering/num[K]`: a numbering identifier "
        "is a native identity, addressed only by raw paths (record 0319), so the "
        "capture is empty and the dependent `numId=` is `invalid_value`",
        lambda c, e: c.script == "word/numbering.sh" and c.detail == "numId"
        and c.result == "invalid_value" and not failed_add(e, {"num", "abstractnum"}),
    ),
    When(
        "D11", "delta", "Re-adding a built-in style",
        "`add --type style` with the id of a built-in style the document already has "
        "(`Heading1`, `Normal`, ...) replaces that definition (upsert)",
        "an existing style id is refused (`duplicate_name`, record 0319): `add` never "
        "drops an existing definition; `set /styles/<id>` changes it",
        lambda c, e: c.verb == "add" and c.kind == "style" and c.result == "duplicate_name",
    ),
    When(
        "D4", "delta", "Bare slide insertion",
        "`add / --type slide` picks the deck's first layout; layout-type words "
        "(`blank`, `title`) also select a layout",
        "`layout` names a layout exactly (record 0256 D4); the harness adds "
        "`layout=Title Slide` (the first layout of Litchi's blank deck) or maps `title` "
        "and `blank` to `Title Slide` and `Blank`, and counts the line as translated",
        lambda c, e: c.detail == "translated:D4",
    ),
    When(
        "D4b", "delta", "Bare slide insertion inside a batch",
        "a batch `add` of a slide without `layout`, or with a layout-type word "
        "(`blank`), picks a layout as on the command line",
        "the harness translates only command-line slide insertions (D4), not batch "
        "items, so the batch refuses (`missing_property` layout, or `not_found` for "
        "`layout=blank`) once its other properties are admitted (record 0269)",
        lambda c, e: c.verb == "batch" and c.fmt == "pptx" and (
            (c.result == "missing_property" and c.detail == "layout")
            or (c.result == "not_found" and not c.detail)
        ),
    ),
    When(
        "D5", "delta", "Resident mode",
        "`open`/`close` start and stop a resident process that later commands reuse",
        "absent (record 0256 D5); every command opens and saves the file, so the "
        "harness treats both as no-ops (counted as resident, not run)",
        lambda c, e: c.result == "skip",
    ),
    When(
        "D7", "delta", "Stray words on the command line",
        "System.CommandLine leaves unmatched words unparsed and OfficeCLI ignores "
        "them; the examples pass unquoted values (`--prop categories=Product A,Product B`) "
        "that the shell splits",
        "the argument parser refuses the command line (usage, exit 2): Litchi is "
        "strict by default (record 0015 D4)",
        lambda c, e: c.result == "usage" and c.verb in LITCHI_VERBS
        and c.detail and not c.detail.startswith("-"),
    ),
    When(
        "U1", "delta", "Upstream script error: `--name=`",
        "`--name=VALUE` is not an OfficeCLI option either: "
        "`RejectUnknownOptionTokens` refuses it (invalid_argument)",
        "refused as an unknown option (usage, exit 2)",
        lambda c, e: c.result == "usage" and c.detail.startswith("--name"),
    ),
    When(
        "U2", "delta", "Upstream script error: batch input not shipped",
        "the script pipes `slideN.json` files that the upstream example directory "
        "does not contain, so OfficeCLI reads an empty batch too",
        "an empty batch payload is `corrupt_batch`",
        lambda c, e: c.result == "corrupt_batch"
        and c.script.endswith("tech--wildlife-company/build.sh"),
    ),
    When(
        "H1", "delta", "Harness: generated inputs need Pillow",
        "the script draws its input pictures with Pillow before adding them",
        "the harness runs `python3 -I -S` (standard library only) so every machine "
        "replays the same lines; the pictures are missing (`file_not_found`)",
        lambda c, e: c.result == "file_not_found",
    ),
    When(
        "D8", "delta", "Formula with a leading `=`",
        "the help schema documents `formula` \"without leading =\" (cell.json), but the "
        "cell handler also strips a leading `=`",
        "the documented contract: a leading `=` is `invalid_value`",
        lambda c, e: c.result == "invalid_value" and c.detail == "formula",
    ),
    When(
        "D9", "delta", "Chart formatting the chart type cannot carry",
        "a chart-level key the chart group's schema omits (`labelPos` on an area, "
        "doughnut, radar or stock chart) is reported as unsupported and the chart is "
        "still added (exit 2)",
        "the key is `invalid_value` and the whole `add` is refused, so no chart "
        "carries a dropped setting (record 0270)",
        lambda c, e: c.result == "invalid_value" and c.kind.lower() == "chart"
        and c.detail == "labelPos",
    ),
    When(
        "D10", "delta", "A property given twice in one chart `add`",
        "the later value wins (a shared option array like `DASH` sets `dataLabels=false` "
        "and the line sets `dataLabels=true` again)",
        "`invalid_argument`: one command assigns each property once, so no value is "
        "silently dropped (records 0270, 0283)",
        lambda c, e: c.verb == "add" and c.kind.lower() == "chart"
        and c.result == "invalid_argument",
    ),
    When(
        "G1", "gap", "Missing verbs",
        "an OfficeCLI verb the `litchi` CLI does not have (counted per verb in the "
        "dispositions below)",
        "absent from `litchi-ops`, the CLI and Python. The raw layer (`validate`, `raw`, "
        "`raw-set`, `add-part`) and `dump` landed in record 0336; `import`, `evaluate`, "
        "`recalculate` (record 0266), `merge` (record 0274) and `refresh` (record 0279) "
        "earlier; none of them is in this class",
        lambda c, e: c.result == "usage" and c.verb not in LITCHI_VERBS,
    ),
    When(
        "G2", "gap", "Copying an element (`add --from`)",
        "`add <file> <parent> --from <path>` clones an element",
        "the CLI passes `--from` (record 0265); `File::apply` refuses it: "
        "\"add-from copying is not integrated\" (record 0256 F6)",
        lambda c, e: c.verb == "add" and c.kind == "(from)",
    ),
    When(
        "G3", "gap", "Find and replace on `set`",
        "`set --find T --replace R` (or `--prop find=`/`replace=`) rewrites matching text",
        "the CLI maps the options to `find=`/`replace=` props (record 0265); record 0273 "
        "integrates the engine for DOCX, PPTX, XLSX and RTF, and record 0285 extends it "
        "to DOCX headers, footers, notes, comments and field results and PPTX table cells "
        "and notes; regular expressions (`regex=true`, record 0010 Q3), tracked-change "
        "finds, XLSX rich-text cells and the other formats remain refused",
        lambda c, e: c.verb == "set" and c.detail.lower() in ("find", "replace", "regex"),
    ),
    When(
        "G4", "gap", "Revision-tracked removal (`remove --prop`)",
        "`remove` takes `--prop revision.author=...` to record a tracked deletion",
        "the CLI passes the props (record 0265); `remove` admits no props "
        "(invalid_argument)",
        lambda c, e: c.verb == "remove" and c.result == "invalid_argument",
    ),
    When(
        "G5", "gap", "Path grammar: OfficeCLI-only steps",
        "`cell[R,C]` table-cell steps, `@id`/`@author` revision identities and other "
        "OfficeCLI path forms",
        "`invalid_path`: the selector grammar (record 0005) has no such step or "
        "identity yet",
        lambda c, e: c.result == "invalid_path",
    ),
    When(
        "G6", "gap", "Value vocabulary",
        "documented value forms such as table-cell `fill=none` and gradient fills "
        "`COLOR1-COLOR2[-ANGLE]` (pptx table-cell.json)",
        "`invalid_value`: the property's coercion accepts colors only",
        lambda c, e: c.result == "invalid_value" and c.fmt == "pptx" and c.detail == "fill",
    ),
    When(
        "G7", "gap", "Undocumented OfficeCLI alias: cell `text`",
        "the cell handler takes `text` as another name of `value` "
        "(`ExcelHandler.Set.Cells.cs`, `case \"value\" or \"text\"`); the help schema "
        "lists only `value`",
        "closed by record 0269: `text` is a registry alias of the cell `value` "
        "(XLSX, XLSB, XLS); a refusal here would be a regression",
        lambda c, e: c.fmt == "xlsx" and c.verb == "set" and c.detail == "text"
        and c.result == "unsupported_property",
    ),
    When(
        "C1", "cascade", "Numbering definitions refused",
        "`add --type num`/`abstractNum` creates the numbering that `numId` names",
        "numbering is a registry gap, so the later `numId` is `invalid_value`",
        lambda c, e: c.detail == "numId" and failed_add(e, {"num", "abstractnum"}),
    ),
    When(
        "F3b", "gap", "Failed after skipping unsupported properties",
        "applies the supported properties and warns about the rest; a command whose "
        "remaining properties are not enough fails",
        "`--unsupported warn` skipped properties, then the rest of the call failed with "
        "`missing_property`, which no other rule explains; the error carries the skipped "
        "properties "
        "(`Error::warnings`, record 0271), which are the gap (each is listed in the "
        "property gaps with refusal `skipped`)",
        lambda c, e: False,  # a fallback in `classify`, after the other rules
    ),
]


def cross_check(call: Call, registry: Registry | None) -> Verdict:
    """A gap verdict for an element or property refusal, or a bug when the
    registry claims what was refused."""
    verb = call.verb
    if registry is None:
        return Verdict("gap", f"{call.fmt or '?'}|{verb}|{call.kind or call.target}|{call.result}|"
                       f"{call.detail}", "no registry export for the format")
    name = registry.element_name(call)
    element = registry.element(name) if name else None
    key = f"{registry.fmt}|{verb}|{element['element'] if element else name}|{call.result}|"
    if call.result == "unsupported_property":
        key += call.detail
        if element is None:
            listed = registry.listed_on(call.detail, verb)
            where = "a batch item" if verb == "batch" else f"`{name}`" if name else "the target"
            if listed:
                return Verdict("gap", key, f"refused on {where}; the registry lists it on "
                               + ", ".join(listed))
            return Verdict("gap", key, f"refused on {where}; not in the registry")
        found = registry.property(element, call.detail)
        documented = "OfficeCLI documents it" if registry.documented(
            name, call.detail) else "not in OfficeCLI's help schema"
        if found is None:
            gaps = element.get("propertyGaps", {})
            gap = next((spec for prop, spec in gaps.items() if word(prop) == word(call.detail)), None)
            if gap:
                return Verdict("gap", key, f"registry gap {gap.get('followUp', '')}; {documented}")
            return Verdict("gap", key, f"not in the registry; {documented}")
        spec_name, spec = found
        op = "add" if verb == "add" else "set" if verb in ("set", "batch") else verb
        if spec.get(op):
            return Verdict("bug", key, f"the registry lists `{spec_name}` as {op}-able")
        return Verdict("gap", key, f"registry: `{spec_name}` is not {op}-able; {documented}")
    if element is None:
        return Verdict("gap", key, "element not in the registry")
    gap = element.get("gap") or registry.top_gaps.get(word(name))
    if gap:
        return Verdict("gap", key, f"registry gap {gap.get('followUp', '')}")
    op = {"add": "add", "set": "set", "get": "get", "query": "query", "remove": "remove"}.get(verb)
    if op and element.get("operations", {}).get(op):
        return Verdict("gap", key, f"registry lists {op}; refused for this target (operation gap)")
    return Verdict("gap", key, f"registry: no {op or verb} on {element['element']}")


def last_failed_add(indices: list[int], calls: list[Call]) -> int | None:
    """The index of the latest refused ``add`` or ``batch`` among ``indices``."""
    for index in reversed(indices):
        if calls[index].refused and calls[index].verb in ("add", "batch"):
            return index
    return None


def classify(calls: list[Call], registries: dict[str, Registry]) -> list[Verdict]:
    verdicts = []
    earlier_by_script: dict[str, list[Call]] = defaultdict(list)
    indices_by_script: dict[str, list[int]] = defaultdict(list)
    for position, call in enumerate(calls):
        earlier = earlier_by_script[call.script]
        verdict = None
        if call.result in FATAL:
            verdict = Verdict("bug", "fatal", call.result)
        elif call.passed and call.detail != "translated:D4":
            verdict = Verdict("passed", "")
        elif call.warned:
            verdict = Verdict("warned", "F3")
        else:
            for rule in RULES:
                if rule.matches(call, earlier):
                    verdict = Verdict(rule.kind, rule.id)
                    break
        if verdict is None and call.result in CASCADE_CODES and failed_add(earlier):
            verdict = Verdict("cascade", "after a refused add or batch",
                              root=last_failed_add(indices_by_script[call.script], calls))
        if verdict is None and (
            call.result.startswith("unsupported_") or call.result == "unsupported_property"
        ):
            verdict = cross_check(call, registries.get(call.fmt))
        if verdict is None and call.failed_after_skipping and call.result == "missing_property":
            # Record 0271: the skipped properties carried what is missing.
            verdict = Verdict("gap", "F3b")
        if verdict is None:
            verdict = Verdict("bug", "untriaged", f"{call.result} {call.detail}".strip())
        verdicts.append(verdict)
        earlier.append(call)
        indices_by_script[call.script].append(position)
    return verdicts


def registries(repo: Path) -> dict[str, Registry]:
    import json

    help_dir = repo / HELP
    result = {}
    for fmt in FORMATS:
        report = json.loads((repo / REGISTRY / f"{fmt}.json").read_text(encoding="utf-8"))
        officecli = capability_delta.officecli_elements(help_dir, fmt)
        result[fmt] = Registry(fmt, report, officecli)
    return result


COLUMNS = (
    "lines", "passed", "translated", "warned", "gap", "delta", "cascade", "bug", "resident",
)


def tally(pairs) -> dict[str, int]:
    counts = dict.fromkeys(COLUMNS, 0)
    readback = [0, 0]
    for call, verdict in pairs:
        if call.readback:
            readback[1] += 1
            readback[0] += call.passed
            continue
        if verdict.kind == "resident" or verdict.rule == "D5":
            counts["resident"] += 1
            continue
        counts["lines"] += 1
        if call.passed:
            counts["passed"] += 1
            if verdict.rule == "D4":
                counts["translated"] += 1
        elif verdict.kind in counts:
            counts[verdict.kind] += 1
    counts["readback"] = f"{readback[0]}/{readback[1]}"
    return counts


def table_rows(groups: dict[str, list]) -> list[str]:
    rows = []
    totals = defaultdict(int)
    readback = [0, 0]
    for name, pairs in groups.items():
        counts = tally(pairs)
        rows.append(
            f"| {name} | " + " | ".join(str(counts[column]) for column in COLUMNS)
            + f" | {counts['readback']} |"
        )
        for column in COLUMNS:
            totals[column] += counts[column]
        done, total = map(int, counts["readback"].split("/"))
        readback[0] += done
        readback[1] += total
    rows.append(
        "| **total** | " + " | ".join(f"**{totals[column]}**" for column in COLUMNS)
        + f" | **{readback[0]}/{readback[1]}** |"
    )
    return rows


HEADER_ROW = (
    "| {} | Lines | Passed | of which translated | Warned | Gap | Delta | Cascade | Bug | "
    "Resident (not run) | Read-back |"
)
RULE_ROW = "|---|" + "---:|" * 10


def cites(calls: list[Call], limit: int = 3) -> str:
    seen = []
    for call in calls:
        if call.cite not in seen:
            seen.append(call.cite)
        if len(seen) == limit:
            break
    return ", ".join(f"`{cite}`" for cite in seen)


FOLLOW_UPS = [
    ("F3", "property-granularity \"apply the rest and warn\" (OfficeCLI applies the others "
     "and exits 2): `--unsupported warn` (record 0269); calls it applied without some "
     "properties", lambda c, v: c.warned),
    ("F3", "calls still refused for an unsupported property (a property the warn policy "
     "cannot drop: an alias of a refused canonical name, or a verb without props)",
     lambda c, v: c.result == "unsupported_property"),
    ("F3", "calls that failed after skipping unsupported properties (named gap F3b; the "
     "error carries the skipped properties, record 0271)",
     lambda c, v: c.failed_after_skipping),
    ("F4", "view options (`--start`, `--end`, `--max-lines`, ...)",
     lambda c, v: c.verb == "view" and c.result == "usage"),
    ("F6", "`add --from`, find/replace on `set`, `remove --shift`, `get --save`",
     lambda c, v: v.rule in ("G2", "G3")),
    ("D6", "missing verbs (the raw layer and `dump` landed in record 0336)",
     lambda c, v: v.rule == "G1"),
]


def unclaimed(calls: list[Call], registries: dict[str, Registry]) -> list[tuple[Call, str]]:
    """Passed writes and reads whose registry element does not claim the verb:
    the registry (and so `capabilities()`) under-claims them."""
    found = []
    for call in calls:
        if not call.passed or call.readback or call.verb not in (
            "add", "set", "remove", "move", "swap", "get", "query"
        ):
            continue
        registry = registries.get(call.fmt)
        if registry is None:
            continue
        name = registry.element_name(call)
        element = registry.element(name) if name else None
        if element is None:
            if name and name != "(from)":
                found.append((call, f"`{name}` is not a registry element"))
            continue
        op = {"move": "set", "swap": "set"}.get(call.verb, call.verb)
        if not element.get("operations", {}).get(op):
            found.append((call, f"the registry has no {op} on `{element['element']}`"))
    return found


def call_skipped(call: Call, loaded: dict[str, Registry] | None) -> list[tuple[Call, Verdict]]:
    """One verdict per property one warned call (or a call that failed after
    skipping) skipped, cross-checked as if the call had been refused for it;
    the group key's refusal is `warned` (or `skipped`)."""
    found = []
    refusal = "|skipped|" if call.failed_after_skipping else "|warned|"
    for name in call.skipped():
        refused = Call(call.script, call.line, call.verb, call.fmt, call.kind,
                       "unsupported_property", name, call.target)
        verdict = cross_check(refused, (loaded or {}).get(call.fmt))
        if verdict.kind == "bug":
            verdict = Verdict("bug", verdict.rule, f"skipped {name}: {verdict.note}")
            found.append((refused, verdict))
            continue
        key = verdict.rule.replace("|unsupported_property|", refusal, 1)
        found.append((call, Verdict(verdict.kind, key, verdict.note)))
    return found


def skipped_verdicts(calls: list[Call],
                     loaded: dict[str, Registry] | None) -> list[tuple[Call, Verdict]]:
    """``call_skipped`` for every call, in order."""
    return [pair for call in calls for pair in call_skipped(call, loaded)]


Resolution = capability_delta.Resolution


class CorpusDispositions:
    """Dispositions of the corpus gap classes (record 0332).

    A named class (a ``RULES`` entry; ``G1`` per verb) is keyed
    ``corpus:<id>`` (``corpus:G1|<verb>``). An element or property group is
    keyed ``corpus:<format>|<verb>|<element>|<refusal>|<property>``. A group
    whose property (or element) OfficeCLI's schema names takes the
    disposition of that row of the capability delta, unless a corpus entry
    names the property literally; otherwise a corpus entry classifies it."""

    def __init__(self, dispositions: "capability_delta.Dispositions",
                 deltas: list, loaded: dict[str, Registry]):
        self.dispositions = dispositions
        self.registries = loaded
        self.elements: dict[tuple[str, str], Resolution] = {}
        self.properties: dict[tuple[str, str, str], Resolution] = {}
        for delta in deltas:
            for index, row in enumerate(delta.elements):
                resolution = (delta.element_dispositions or [None] * len(delta.elements))[index]
                if row[0] and resolution is not None:
                    self.elements[(delta.name, row[0])] = resolution
            for index, row in enumerate(delta.properties):
                resolution = (delta.property_dispositions or [None] * len(delta.properties))[index]
                if resolution is not None:
                    self.properties[(delta.name, row[0], row[1])] = resolution
        self.cache: dict[str, Resolution] = {}

    def officecli_names(self, fmt: str, element: str, prop: str) -> tuple[str | None, str | None]:
        """The OfficeCLI element and canonical property a group names."""
        registry = self.registries.get(fmt)
        schema = registry.officecli.get(word(element)) if registry and element else None
        if schema is None:
            return None, None
        if not prop:
            return schema.name, None
        wanted = word(prop)
        for name, spec in schema.properties.items():
            aliases = spec.get("aliases", []) if isinstance(spec, dict) else []
            if wanted in {word(name), *map(word, aliases if isinstance(aliases, list) else [])}:
                return schema.name, name
        return schema.name, None

    def tier(self, key: str, fmt: str, element: str | None, prop: str | None) -> str:
        tier = self.dispositions.keyed_tier(f"corpus:{key}")
        if tier is None and element:
            tier = self.dispositions.tier(fmt, element, prop, mark=False)
        return tier or ""

    def group(self, key: str) -> Resolution:
        if key in self.cache:
            return self.cache[key]
        fmt, verb, element, refusal, prop = key.split("|", 4)
        schema, canonical = self.officecli_names(fmt, element, prop)
        literal = dict(where=lambda pattern: capability_delta.specificity(pattern)[0] > 0)
        own = self.dispositions.entries.best(f"corpus:{key}", mark=False, **literal) if prop else None
        inherited = None
        if own is None and schema is not None:
            if canonical is not None:
                inherited = self.properties.get((fmt, schema, canonical))
            elif not prop:
                inherited = self.elements.get((fmt, schema))
        if inherited is not None and inherited.disposition is not None:
            where = f"{schema}.{canonical}" if canonical else schema
            resolution = Resolution(inherited.disposition, inherited.tier, inherited.ref,
                                    inherited.entry, inherited.reason, via=f"delta `{where}`")
        else:
            tier = self.tier(key, fmt, schema or element, canonical or prop or None)
            resolution = (self.dispositions.resolve(f"corpus:{key}", tier)
                          or Resolution(None, tier))
        self.cache[key] = resolution
        return resolution

    def rule(self, rule: Rule, call: Call | None = None) -> Resolution:
        key = f"{rule.id}|{call.verb}" if rule.id == "G1" and call is not None else rule.id
        if key in self.cache:
            return self.cache[key]
        tier = "" if rule.kind == "delta" else (self.dispositions.keyed_tier(f"corpus:{key}") or "")
        resolution = self.dispositions.resolve(f"corpus:{key}", tier) or Resolution(None, tier)
        self.cache[key] = resolution
        return resolution


RULES_BY_ID = {rule.id: rule for rule in RULES}


def line_classes(index: int, calls: list[Call], verdicts: list[Verdict],
                 classes: CorpusDispositions, loaded: dict[str, Registry] | None,
                 seen: frozenset = frozenset()) -> list[tuple[str, Resolution | None]]:
    """The gap classes one call that did not pass carries, with their
    dispositions (None: a bug). A cascade carries its root's classes."""
    call, verdict = calls[index], verdicts[index]
    found: list[tuple[str, Resolution | None]] = []
    if verdict.kind == "bug":
        return [("bug", None)]
    if verdict.kind == "cascade":
        if verdict.root is None or verdict.root in seen:
            return [("cascade without a root", None)]
        return line_classes(verdict.root, calls, verdicts, classes, loaded, seen | {index})
    if verdict.rule in RULES_BY_ID and verdict.rule != "F3b":
        rule = RULES_BY_ID[verdict.rule]
        name = f"G1|{call.verb}" if rule.id == "G1" else rule.id
        found.append((name, classes.rule(rule, call)))
    elif "|" in verdict.rule:
        found.append((verdict.rule, classes.group(verdict.rule)))
    if verdict.kind == "warned" or call.failed_after_skipping:
        for _, skipped in call_skipped(call, loaded):
            if skipped.kind == "bug":
                found.append(("bug", None))
            else:
                found.append((skipped.rule, classes.group(skipped.rule)))
    return found or [("unclassified", None)]


@dataclass
class Subset:
    """The v1-scope corpus subset (record 0332)."""

    scripts: dict[str, str]  # script -> tier
    lines: list[tuple[int, str, list]]  # (call index, state, classes)

    @property
    def green(self) -> bool:
        return all(state != "open" for _, state, _ in self.lines)


def v1_subset(calls: list[Call], verdicts: list[Verdict], classes: CorpusDispositions,
              loaded: dict[str, Registry] | None) -> Subset:
    """Every counted line of a script whose tier is P0 or v1: passed, or
    accounted (each class it carries is wontfix, deferred or out-of-scope),
    or open."""
    scripts: dict[str, str] = {}
    for call in calls:
        if call.script not in scripts:
            scripts[call.script] = classes.dispositions.script_tier(call.script) or ""
    lines = []
    for index, (call, verdict) in enumerate(zip(calls, verdicts)):
        if scripts[call.script] not in capability_delta.V1_TIERS:
            continue
        if call.readback or verdict.kind == "resident" or verdict.rule == "D5":
            continue
        if call.passed:
            lines.append((index, "passed", []))
            continue
        found = line_classes(index, calls, verdicts, classes, loaded)
        accounted = all(resolution is not None and resolution.accounted for _, resolution in found)
        lines.append((index, "accounted" if accounted else "open", found))
    return Subset(scripts, lines)


def blocker(found: list[tuple[str, Resolution | None]]) -> str:
    """Why an open line is open: the first class that does not account for it."""
    for name, resolution in found:
        if resolution is None:
            return "bug"
        if resolution.accounted:
            continue
        if resolution.disposition is None:
            return f"no disposition (`{name}`)"
        return f"{resolution.disposition} {resolution.ref}".strip()
    return "unclassified"


def account(found: list[tuple[str, Resolution | None]]) -> str:
    """The column an accounted line counts in: its most deferred class."""
    values = {resolution.disposition for _, resolution in found if resolution is not None}
    for value in ("deferred-P2", "deferred-P1", "wontfix", "out-of-scope"):
        if value in values:
            return value
    return "out-of-scope"


def corpus_problems(calls: list[Call], verdicts: list[Verdict], classes: CorpusDispositions,
                    loaded: dict[str, Registry] | None) -> list[str]:
    """What ``--check`` refuses about the corpus dispositions."""
    problems: list[str] = []
    checked: set[str] = set()

    def check(name: str, resolution: Resolution, kind: str) -> None:
        if name in checked:
            return
        checked.add(name)
        if kind == "delta":
            if resolution.disposition is None:
                problems.append(f"corpus {name}: no disposition")
            elif resolution.disposition not in ("wontfix", "out-of-scope"):
                problems.append(f"corpus {name}: a recorded delta is wontfix or out-of-scope, "
                                f"not `{resolution.disposition}`")
            return
        if kind == "gap" and not resolution.tier and resolution.disposition is not None:
            problems.append(f"corpus {name}: no tier")
        problems.extend(capability_delta.resolution_problems(f"corpus {name}", resolution))

    for index, (call, verdict) in enumerate(zip(calls, verdicts)):
        if verdict.rule in RULES_BY_ID and verdict.rule != "F3b":
            rule = RULES_BY_ID[verdict.rule]
            name = f"G1|{call.verb}" if rule.id == "G1" else rule.id
            check(name, classes.rule(rule, call), rule.kind)
        elif verdict.kind == "gap" and "|" in verdict.rule:
            check(verdict.rule, classes.group(verdict.rule), "gap")
        if call.detail == "translated:D4":
            check("D4", classes.rule(RULES_BY_ID["D4"]), "delta")
        for _, skipped in call_skipped(call, loaded):
            if skipped.kind == "gap":
                check(skipped.rule, classes.group(skipped.rule), "gap")
    dispositions = classes.dispositions
    for script in sorted({call.script for call in calls}):
        if dispositions.script_tier(script) is None:
            problems.append(f"corpus script {script}: no tier")
    problems += [f"stale disposition key `{key}`: it classifies no corpus class"
                 for key in dispositions.entries.unused(prefix="corpus:")]
    problems += [f"stale tier key `{key}`: it classifies no corpus class"
                 for key in dispositions.tiers.unused(prefix="corpus:")]
    problems += [f"stale corpus script key `{key}`: it matches no script"
                 for key in dispositions.scripts.unused()]
    return problems


def cell(resolution: Resolution | None) -> str:
    if resolution is None:
        return ""
    text = resolution.cell()
    if resolution.via:
        text += f" via {resolution.via}"
    return text


def render_subset(subset: Subset, calls: list[Call], dispositions) -> list[str]:
    """The v1-scope subset: its definition, its status and its open lines."""
    out = ["", "## v1-scope subset (design §14)", ""]
    open_lines = [line for line in subset.lines if line[1] == "open"]
    status = "green" if subset.green else "red"
    out += [
        "Design §14 requires the v1-scope subset of the corpus green, with every remaining "
        "delta recorded. The subset is every counted line of the scripts whose tier "
        f"(`corpusScripts` in `{capability_delta.DISPOSITIONS.as_posix()}`) is P0 or v1 "
        "(design §13). A line is **accounted** when every gap class it carries (its delta, "
        "named gap or element/property group; each property a warned call skipped; for a "
        "cascade, the refused call it follows) is `wontfix`, `deferred-P1`, `deferred-P2` "
        "or `out-of-scope`, and **open** otherwise (a `planned` class, a bug or a class without "
        "a disposition). The subset is green when no line is open.",
        "",
        f"**Status: {status}.** {len(subset.lines)} v1 lines: "
        f"{sum(1 for line in subset.lines if line[1] == 'passed')} passed, "
        f"{sum(1 for line in subset.lines if line[1] == 'accounted')} accounted, "
        f"{len(open_lines)} open.",
        "",
        "Script tiers:",
        "",
        "| Scripts | Tier | Basis | Scripts matched |",
        "|---|---|---|---:|",
    ]
    for entry in dispositions.data.get("corpusScripts", []):
        keys = ", ".join(f"`{key}`" for key in entry["keys"])
        matched = sum(1 for script in subset.scripts
                      if dispositions.scripts.best(script, mark=False)
                      and dispositions.scripts.best(script, mark=False)[1] is entry)
        out.append(f"| {keys} | {entry['tier']} | {entry['basis']} | {matched} |")
    columns = ("passed", "wontfix", "deferred-P1", "deferred-P2", "out-of-scope", "open")
    by_family: dict[str, dict[str, int]] = {}
    for index, state, found in subset.lines:
        counts = by_family.setdefault(family(calls[index].script), dict.fromkeys(columns, 0))
        column = state if state in ("passed", "open") else account(found)
        counts[column] += 1
    out += ["", "| Family | v1 lines | " + " | ".join(columns) + " |",
            "|---|---:|" + "---:|" * len(columns)]
    totals = dict.fromkeys(columns, 0)
    for name in sorted(by_family):
        counts = by_family[name]
        out.append(f"| {name} | {sum(counts.values())} | "
                   + " | ".join(str(counts[c]) for c in columns) + " |")
        for column in columns:
            totals[column] += counts[column]
    out.append(f"| **total** | **{sum(totals.values())}** | "
               + " | ".join(f"**{totals[c]}**" for c in columns) + " |")
    blockers: dict[str, list[Call]] = defaultdict(list)
    for index, _, found in open_lines:
        blockers[blocker(found)].append(calls[index])
    out += ["", "Open v1 lines by what blocks them (the first class that does not account "
            "for the line):", ""]
    if not blockers:
        out.append("None.")
    else:
        out += ["| Blocked by | Lines | Examples |", "|---|---:|---|"]
        for name in sorted(blockers, key=lambda key: (-len(blockers[key]), key)):
            out.append(f"| {name} | {len(blockers[name])} | {cites(blockers[name])} |")
    return out


def render(calls: list[Call], verdicts: list[Verdict], commit: str,
           under: list[tuple[Call, str]] | None = None,
           registries_for_render: dict[str, Registry] | None = None,
           classes: CorpusDispositions | None = None,
           subset: Subset | None = None) -> str:
    pairs = list(zip(calls, verdicts))
    out = [
        "# OfficeCLI example-corpus deltas",
        "",
        "Generated by `python3 tools/officecli_corpus.py` from",
        f"`{OUTCOMES.as_posix()}` (record 0265); do not edit by hand.",
        "Regenerate after re-recording the outcomes:",
        "",
        "```sh",
        "LITCHI_BLESS=1 cargo test -p litchi-cli --test officecli_corpus",
        "python3 tools/officecli_corpus.py      # --check verifies this file",
        "```",
        "",
        f"The corpus is every OfficeCLI example script (`examples/{{word,excel,ppt}}/**/*.sh`, "
        f"upstream commit `{commit}`), vendored under `test-data/officecli/examples/` and "
        "replayed by `crates/litchi-cli/tests/officecli_corpus.rs` with `officecli` "
        f"translated to `litchi` by `{SHIM.as_posix()}`. A line counts once per call, so a "
        "line inside a loop or a shell function counts each time it runs.",
        "",
        "Kinds:",
        "",
        "- **passed**: `litchi` exited 0. *Translated* lines needed a recorded translation "
        "first (delta D4).",
        "- **gap**: Litchi lacks the verb, option, path form, element or property. Element "
        "and property refusals are cross-checked against the registry export "
        f"(`{REGISTRY.as_posix()}/<format>.json`, the golden of `schema()`, which "
        "`capabilities()` shares): the note says what the registry records.",
        "- **delta**: a recorded difference, deliberate or caused upstream (table below).",
        "- **cascade**: a path that does not exist because an earlier `add` or `batch` of "
        "the same script was refused.",
        "- **warned**: the harness passes `--unsupported warn` to `set`, `add` and `batch`, "
        "because OfficeCLI applies the supported properties of a command and warns about "
        "the rest by default (record 0269). A warned call applied without the properties "
        "it lists (exit 9); each skipped property is listed in the gap tables below with "
        "refusal `warned`. A call that still failed after skipping is the named gap F3b; "
        "its skipped properties are listed with refusal `skipped` (record 0271).",
        "- **bug**: a refusal of something the registry claims, or an unexplained failure; "
        "none may be a panic or `internal_error` (the test refuses those).",
        "- **resident**: `open`/`close`, not run (delta D5).",
        "- **read-back**: after each script the harness runs `view <file> text` and "
        "`get <file> / --depth 2` on every file the script created.",
        "",
        "Every delta, named gap and element/property group carries a disposition "
        f"(`{capability_delta.DISPOSITIONS.as_posix()}`, record 0332): `planned` with a work "
        "package or record id, `deferred-P1`/`deferred-P2` by design §13 tier, `wontfix` or "
        "`out-of-scope` with a reason. A group OfficeCLI's schema names inherits the "
        "capability delta's disposition for that row (`via delta`). `--check` fails on a "
        "class without a disposition, on a P0 or v1 class that is deferred or wontfix, and "
        "on a disposition key that matches nothing.",
        "",
    ]
    by_family: dict[str, list] = defaultdict(list)
    by_format: dict[str, list] = defaultdict(list)
    by_script: dict[str, list] = defaultdict(list)
    for call, verdict in pairs:
        by_family[family(call.script)].append((call, verdict))
        by_format[call.fmt or "(none)"].append((call, verdict))
        by_script[call.script].append((call, verdict))
    out += ["## Summary by script family", "", HEADER_ROW.format("Family"), RULE_ROW]
    out += table_rows(dict(sorted(by_family.items())))
    out += ["", "## Summary by format", "",
            "The format is the extension of the document each call names.", "",
            HEADER_ROW.format("Format"), RULE_ROW]
    out += table_rows(dict(sorted(by_format.items())))
    if subset is not None and classes is not None:
        out += render_subset(subset, calls, classes.dispositions)

    def rule_cell(rule: Rule, hits: list[Call]) -> str:
        if classes is None or not hits:
            return ""
        if rule.id == "G1":
            verbs = sorted({call.verb for call in hits})
            return "; ".join(
                f"`{verb}` ({sum(1 for call in hits if call.verb == verb)}): "
                f"{cell(classes.rule(rule, next(c for c in hits if c.verb == verb)))}"
                for verb in verbs)
        return cell(classes.rule(rule))

    def rule_reason(rule: Rule, hits: list[Call]) -> str:
        if classes is None or not hits or rule.id == "G1":
            return ""
        return classes.rule(rule).reason

    out += ["", "## Deltas", "",
            "| Id | Delta | OfficeCLI | Litchi | Lines | Examples | Disposition | Reason |",
            "|---|---|---|---|---:|---|---|---|"]
    for rule in RULES:
        if rule.kind != "delta":
            continue
        hits = [call for call, verdict in pairs if verdict.rule == rule.id]
        if rule.id == "D4":
            hits = [call for call, verdict in pairs if call.detail == "translated:D4"]
        out.append(f"| {rule.id} | {rule.title} | {rule.officecli} | {rule.litchi} | "
                   f"{len(hits)} | {cites(hits)} | {rule_cell(rule, hits)} | "
                   f"{rule_reason(rule, hits)} |")

    out += ["", "## Named gaps", "",
            "| Id | Gap | OfficeCLI | Litchi | Lines | Examples | Disposition | Reason |",
            "|---|---|---|---|---:|---|---|---|"]
    for rule in RULES:
        if rule.kind not in ("gap", "cascade"):
            continue
        hits = [call for call, verdict in pairs if verdict.rule == rule.id]
        out.append(f"| {rule.id} | {rule.title} | {rule.officecli} | {rule.litchi} | "
                   f"{len(hits)} | {cites(hits)} | {rule_cell(rule, hits)} | "
                   f"{rule_reason(rule, hits)} |")

    groups: dict[str, list] = defaultdict(list)
    notes: dict[str, str] = {}
    for call, verdict in pairs:
        if verdict.kind == "gap" and "|" in verdict.rule:
            groups[verdict.rule].append(call)
            notes[verdict.rule] = verdict.note
    for call, skipped in skipped_verdicts(calls, registries_for_render):
        if skipped.kind == "gap":
            groups[skipped.rule].append(call)
            notes[skipped.rule] = skipped.note
    for fmt in sorted({key.split("|")[0] for key in groups}):
        out += ["", f"## Element and property gaps: {fmt}", "",
                "| Verb | Element | Refusal | Property | Lines | Registry cross-check | Examples "
                "| Tier | Disposition |",
                "|---|---|---|---|---:|---|---|---|---|"]
        keys = sorted(
            (key for key in groups if key.split("|")[0] == fmt),
            key=lambda key: (-len(groups[key]), key),
        )
        for key in keys:
            _, verb, element, result, prop = key.split("|", 4)
            resolution = classes.group(key) if classes is not None else None
            out.append(
                f"| {verb} | {element or '-'} | `{result}` | {f'`{prop}`' if prop else '-'} | "
                f"{len(groups[key])} | {notes[key]} | {cites(groups[key], 2)} | "
                f"{resolution.tier if resolution else ''} | {cell(resolution)} |"
            )

    out += ["", "## Bugs", ""]
    bugs = [(call, verdict) for call, verdict in pairs if verdict.kind == "bug"]
    bugs += [
        (call, verdict) for call, verdict in skipped_verdicts(calls, registries_for_render)
        if verdict.kind == "bug"
    ]
    if not bugs:
        out.append("None: every refusal is a recorded delta, a gap the registry agrees with, "
                   "or a cascade.")
    else:
        out += ["| Script line | Verb | Format | Element | Refusal | Detail | Note |",
                "|---|---|---|---|---|---|---|"]
        for call, verdict in bugs:
            detail = call.detail or "-"
            if call.after_skip:
                detail += f" (after skipping {call.after_skip})"
            out.append(f"| `{call.cite}` | {call.verb} | {call.fmt} | {call.kind or call.target} "
                       f"| `{call.result}` | {detail} | {verdict.note} |")

    out += ["", "## Passed calls the registry does not claim", "",
            "A passed call on an element whose registry entry lacks the verb means "
            "`schema()` and `capabilities()` under-claim it.", ""]
    if not under:
        out.append("None.")
    else:
        groups: dict[str, list] = defaultdict(list)
        for call, note in under:
            groups[f"{call.fmt} {call.verb}: {note}"].append(call)
        out += ["| Finding | Lines | Examples |", "|---|---:|---|"]
        for key in sorted(groups):
            out.append(f"| {key} | {len(groups[key])} | {cites(groups[key])} |")

    out += ["", "## Record 0256 follow-ups hit by the corpus", "",
            "| Follow-up | What | Lines |", "|---|---|---:|"]
    for follow_up, what, test in FOLLOW_UPS:
        out.append(f"| {follow_up} | {what} | {sum(1 for c, v in pairs if test(c, v))} |")

    out += ["", "## Summary by script", "", HEADER_ROW.format("Script"), RULE_ROW]
    out += table_rows({f"`{name}`": by_script[name] for name in sorted(by_script)})
    out.append("")
    return "\n".join(out)


def generate(repo: Path, outcomes: Path, dispositions_path: Path | None = None,
             problems: list[str] | None = None) -> tuple[str, list[Call], list[Verdict]]:
    """The table; ``problems`` (when given) collects what ``--check`` refuses
    about the dispositions."""
    calls = parse(outcomes.read_text(encoding="utf-8"))
    loaded = registries(repo)
    verdicts = classify(calls, loaded)
    commit = (repo / UPSTREAM_COMMIT).read_text(encoding="utf-8").strip()[:10]
    dispositions = capability_delta.Dispositions.load(
        dispositions_path or repo / capability_delta.DISPOSITIONS)
    deltas, _, _ = capability_delta.classified(repo / VENDORED, repo, dispositions)
    classes = CorpusDispositions(dispositions, deltas, loaded)
    found = corpus_problems(calls, verdicts, classes, loaded)
    subset = v1_subset(calls, verdicts, classes, loaded)
    if problems is not None:
        problems.extend(found)
    text = render(calls, verdicts, commit, unclaimed(calls, loaded), loaded, classes, subset)
    return text, calls, verdicts


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--outcomes", type=Path, default=REPO / OUTCOMES)
    parser.add_argument("--output", type=Path, default=REPO / OUTPUT)
    parser.add_argument("--dispositions", type=Path,
                        default=REPO / capability_delta.DISPOSITIONS)
    parser.add_argument("--check", action="store_true",
                        help="fail if the committed table differs from a fresh one, or a gap "
                        "class lacks a valid disposition")
    arguments = parser.parse_args(argv)
    problems: list[str] = []
    text, calls, verdicts = generate(REPO, arguments.outcomes, arguments.dispositions, problems)
    fatal = [call.cite for call, verdict in zip(calls, verdicts) if verdict.rule == "fatal"]
    if fatal:
        print("panics or internal errors recorded: " + ", ".join(fatal), file=sys.stderr)
        return 1
    if len(text.encode()) >= capability_delta.SIZE_LIMIT:
        raise SystemExit("the corpus table would reach 1 MiB; split it before committing")
    for problem in problems:
        print(f"disposition: {problem}", file=sys.stderr)
    status = 1 if problems else 0
    if arguments.check:
        current = arguments.output.read_text(encoding="utf-8") if arguments.output.exists() else ""
        if current != text:
            print(f"{arguments.output} is stale; rerun {Path(__file__).name}", file=sys.stderr)
            return 1
        return status
    arguments.output.write_text(text, encoding="utf-8")
    return status


if __name__ == "__main__":
    sys.exit(main())
