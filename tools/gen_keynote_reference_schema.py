#!/usr/bin/env python3
"""Generate the Keynote object-reference schema used by structural edits.

The focused ``litchi-keynote`` owner rewrites object graphs at the protobuf
wire level (no generated Prost messages in production code).  Duplicating a
slide, materializing a layout or removing a drawable needs to know *where*
``TSP.Reference`` and ``TSP.DataReference`` values live inside each native
message.  This tool derives that from the tracked iWork schema in
``crates/litchi-iwa-protos/src/protos`` and writes a static Rust table:

* one entry per message type that (transitively) holds a reference;
* per entry, only the fields that hold one: a direct reference, a data
  reference, or a nested message with its table index;
* a ``root`` function mapping the native IWA message-type identifiers that the
  Keynote owner admits to their message entries.

The message-type identifiers are the native Keynote registry values.  Each
was checked against every Keynote package under ``test-data/iwork/keynote``:
for every object, every identifier listed in its archive ``MessageInfo``
``object_references`` is found by walking the payload with this schema
(record 0298).  Re-run with ``--check`` in CI-like contexts to prove the
committed file is current.

Usage::

    python3 tools/gen_keynote_reference_schema.py            # rewrite
    python3 tools/gen_keynote_reference_schema.py --check    # verify only
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PROTOS = ROOT / "crates" / "litchi-iwa-protos" / "src" / "protos"
OUTPUT = ROOT / "crates" / "litchi-keynote" / "src" / "package" / "structure" / "schema.rs"

# Native Keynote message-type registry values admitted by structural edits.
# Every pair was verified against the tracked Keynote corpus (see docstring).
ROOT_TYPES: dict[int, str] = {
    1: ".KN.DocumentArchive",
    2: ".KN.ShowArchive",
    3: ".KN.UIStateArchive",
    4: ".KN.SlideNodeArchive",
    5: ".KN.SlideArchive",
    7: ".KN.PlaceholderArchive",
    8: ".KN.BuildArchive",
    9: ".KN.SlideStyleArchive",
    10: ".KN.ThemeArchive",
    15: ".KN.NoteArchive",
    19: ".KN.ClassicStylesheetRecordArchive",
    21: ".KN.Soundtrack",
    23: ".KN.DesktopUILayoutArchive",
    24: ".KN.CanvasSelectionArchive",
    25: ".KN.SlideCollectionSelectionArchive",
    26: ".KN.MotionBackgroundStyleArchive",
    401: ".TSS.StylesheetArchive",
    633: ".TSA.CaptionInfoArchive",
    2001: ".TSWP.StorageArchive",
    2011: ".TSWP.ShapeInfoArchive",
    2021: ".TSWP.CharacterStyleArchive",
    2022: ".TSWP.ParagraphStyleArchive",
    2023: ".TSWP.ListStyleArchive",
    2024: ".TSWP.ColumnStyleArchive",
    2025: ".TSWP.ShapeStyleArchive",
    2043: ".TSWP.NumberAttachmentArchive",
    3005: ".TSD.ImageArchive",
    3006: ".TSD.MaskArchive",
    3007: ".TSD.MovieArchive",
    3016: ".TSD.MediaStyleArchive",
    3047: ".TSD.GuideStorageArchive",
    3056: ".TSD.CommentStorageArchive",
    3097: ".TSD.StandinCaptionArchive",
    5020: ".TSCH.ChartStylePreset",
    5021: ".TSCH.ChartDrawableArchive",
    5022: ".TSCH.ChartStyleArchive",
    5023: ".TSCH.ChartNonStyleArchive",
    5024: ".TSCH.LegendStyleArchive",
    5025: ".TSCH.LegendNonStyleArchive",
    5026: ".TSCH.ChartAxisStyleArchive",
    5027: ".TSCH.ChartAxisNonStyleArchive",
    5028: ".TSCH.ChartSeriesStyleArchive",
    5029: ".TSCH.ChartSeriesNonStyleArchive",
    5030: ".TSCH.ReferenceLineStyleArchive",
    5031: ".TSCH.ReferenceLineNonStyleArchive",
    10024: ".TSWP.TextStylePresetArchive",
}

SCALARS = {
    "double", "float", "int32", "int64", "uint32", "uint64", "sint32", "sint64",
    "fixed32", "fixed64", "sfixed32", "sfixed64", "bool", "string", "bytes",
}
TOKEN = re.compile(r'"(?:[^"\\]|\\.)*"|[A-Za-z_.][A-Za-z0-9_.]*|-?\d+(?:\.\d+)?|[{}\[\];=<>(),]')
REFERENCE = ".TSP.Reference"
DATA_REFERENCE = ".TSP.DataReference"


def tokenize(text: str) -> list[str]:
    text = re.sub(r"//[^\n]*", "", text)
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    return TOKEN.findall(text)


class Schema:
    def __init__(self) -> None:
        self.raw: dict[str, list[tuple[int, str, str, str]]] = {}
        self.enums: set[str] = set()
        self.extends: list[tuple[str, str, list[tuple[int, str, str, str]]]] = []
        self.fields: dict[str, list[tuple[int, str]]] = {}

    def parse_file(self, path: Path) -> None:
        toks = tokenize(path.read_text(encoding="utf-8"))
        package = ""

        def skip_braced(i: int) -> int:
            depth = 0
            while True:
                if toks[i] == "{":
                    depth += 1
                elif toks[i] == "}":
                    depth -= 1
                    if depth == 0:
                        return i + 1
                i += 1

        def body(i: int, scope: str) -> tuple[int, list[tuple[int, str, str, str]]]:
            fields: list[tuple[int, str, str, str]] = []
            while toks[i] != "}":
                token = toks[i]
                if token == "message":
                    full = f"{scope}.{toks[i + 1]}"
                    self.raw.setdefault(full, [])
                    i, nested = body(i + 3, full)
                    self.raw[full] = nested + self.raw[full]
                elif token == "enum":
                    self.enums.add(f"{scope}.{toks[i + 1]}")
                    i = skip_braced(i)
                elif token == "extend":
                    target = toks[i + 1]
                    i, extension = body(i + 3, scope)
                    self.extends.append((target, scope, extension))
                elif token in {"extensions", "reserved", "option"}:
                    while toks[i] != ";":
                        i += 1
                    i += 1
                elif token in {"optional", "required", "repeated"}:
                    typ, number = toks[i + 1], int(toks[i + 4])
                    if toks[i + 3] != "=":
                        raise SystemExit(f"{path}: malformed field near {toks[i:i + 6]}")
                    i += 5
                    if toks[i] == "[":
                        while toks[i] != "]":
                            i += 1
                        i += 1
                    if toks[i] != ";":
                        raise SystemExit(f"{path}: malformed field terminator")
                    i += 1
                    fields.append((number, token, typ, scope))
                elif token == ";":
                    i += 1
                else:
                    raise SystemExit(f"{path}: unsupported token {token!r}")
            return i + 1, fields

        i = 0
        while i < len(toks):
            token = toks[i]
            if token in {"syntax", "import", "option"}:
                while toks[i] != ";":
                    i += 1
                i += 1
            elif token == "package":
                package = toks[i + 1]
                i += 3
            elif token == "message":
                full = f".{package}.{toks[i + 1]}"
                self.raw.setdefault(full, [])
                i, nested = body(i + 3, full)
                self.raw[full] = nested + self.raw[full]
            elif token == "enum":
                self.enums.add(f".{package}.{toks[i + 1]}")
                i = skip_braced(i)
            elif token == "extend":
                target = toks[i + 1]
                i, extension = body(i + 3, f".{package}")
                self.extends.append((target, f".{package}", extension))
            elif token == ";":
                i += 1
            else:
                raise SystemExit(f"{path}: unsupported top-level token {token!r}")

    def resolve(self, typ: str, scope: str) -> str:
        if typ in SCALARS:
            return typ
        if typ.startswith("."):
            if typ in self.raw or typ in self.enums:
                return typ
            raise SystemExit(f"unresolved type {typ}")
        parts = scope.split(".")
        while len(parts) > 1:
            candidate = ".".join(parts) + "." + typ
            if candidate in self.raw or candidate in self.enums:
                return candidate
            parts = parts[:-1]
        raise SystemExit(f"unresolved type {typ} in {scope}")

    def finish(self) -> None:
        for target, scope, extension in self.extends:
            if target.startswith(".google."):
                continue
            full = self.resolve(target, scope)
            self.raw[full] = self.raw[full] + extension
        for message, fields in self.raw.items():
            resolved = []
            seen: set[int] = set()
            for number, _label, typ, scope in fields:
                if number in seen:
                    raise SystemExit(f"duplicate field {number} in {message}")
                seen.add(number)
                resolved.append((number, self.resolve(typ, scope)))
            self.fields[message] = resolved

    def holders(self) -> set[str]:
        holds = {REFERENCE, DATA_REFERENCE}
        changed = True
        while changed:
            changed = False
            for message, fields in self.fields.items():
                if message in holds:
                    continue
                if any(typ in holds for _number, typ in fields):
                    holds.add(message)
                    changed = True
        return holds


def load_schema() -> Schema:
    schema = Schema()
    for path in sorted(PROTOS.glob("*.proto")):
        schema.parse_file(path)
    schema.finish()
    return schema


def generate() -> str:
    schema = load_schema()
    holds = schema.holders()
    order: list[str] = []
    index: dict[str, int] = {}

    def visit(message: str) -> None:
        if message in index or message in {REFERENCE, DATA_REFERENCE}:
            return
        index[message] = len(order)
        order.append(message)
        for _number, typ in schema.fields[message]:
            if typ in holds:
                visit(typ)

    for type_id in sorted(ROOT_TYPES):
        message = ROOT_TYPES[type_id]
        if message not in schema.fields:
            raise SystemExit(f"root message {message} is not in the schema")
        visit(message)

    lines = [
        "//! Object-reference locations of the native Keynote messages admitted by",
        "//! structural edits.",
        "//!",
        "//! Generated by `tools/gen_keynote_reference_schema.py` from the tracked",
        "//! iWork schema in `crates/litchi-iwa-protos/src/protos`; do not edit.",
        "//! Each entry lists only the fields that (transitively) hold a",
        "//! `TSP.Reference` or `TSP.DataReference`.",
        "",
        "/// What one schema field holds.",
        "#[derive(Debug, Clone, Copy, PartialEq, Eq)]",
        "pub(super) enum Slot {",
        "    /// A `TSP.Reference` (object identifier in field 1).",
        "    Reference,",
        "    /// A `TSP.DataReference` (data identifier in field 1).",
        "    DataReference,",
        "    /// A nested message, by its index in [`MESSAGES`].",
        "    Message(u16),",
        "}",
        "",
        "/// One reference-bearing field of a message.",
        "#[derive(Debug, Clone, Copy)]",
        "pub(super) struct Field {",
        "    /// The protobuf field number.",
        "    pub(super) number: u32,",
        "    /// What the field holds.",
        "    pub(super) slot: Slot,",
        "}",
        "",
        "const fn reference(number: u32) -> Field {",
        "    Field { number, slot: Slot::Reference }",
        "}",
        "",
        "const fn data(number: u32) -> Field {",
        "    Field { number, slot: Slot::DataReference }",
        "}",
        "",
        "const fn message(number: u32, index: u16) -> Field {",
        "    Field { number, slot: Slot::Message(index) }",
        "}",
        "",
        "/// The reference-bearing messages: (name, fields sorted by number).",
        f"pub(super) static MESSAGES: [(&str, &[Field]); {len(order)}] = [",
    ]
    for message in order:
        entries = []
        for number, typ in sorted(schema.fields[message]):
            if typ == REFERENCE:
                entries.append(f"reference({number})")
            elif typ == DATA_REFERENCE:
                entries.append(f"data({number})")
            elif typ in holds:
                entries.append(f"message({number}, {index[typ]})")
        lines.append(f'    ("{message[1:]}", &[{", ".join(entries)}]),')
    lines += [
        "];",
        "",
        "/// The [`MESSAGES`] index of an admitted native message type.",
        "pub(super) const fn root(message_type: u32) -> Option<u16> {",
        "    match message_type {",
    ]
    for type_id in sorted(ROOT_TYPES):
        lines.append(f"        {type_id} => Some({index[ROOT_TYPES[type_id]]}),")
    lines += [
        "        _ => None,",
        "    }",
        "}",
        "",
    ]
    return "\n".join(lines)


def rustfmt(text: str) -> str:
    """Format generated Rust with the repository's rustfmt configuration, so
    the committed file is also stable under ``cargo fmt``."""
    import subprocess

    result = subprocess.run(
        ["rustfmt", "--edition", "2024", "--config-path", str(ROOT / "rustfmt.toml"), "--emit", "stdout"],
        input=text,
        capture_output=True,
        text=True,
        check=True,
        cwd=ROOT,
    )
    return result.stdout


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="verify the committed file")
    args = parser.parse_args()
    text = rustfmt(generate())
    if args.check:
        current = OUTPUT.read_text(encoding="utf-8") if OUTPUT.is_file() else ""
        if current != text:
            print(f"{OUTPUT.relative_to(ROOT)} is stale; re-run the generator", file=sys.stderr)
            return 1
        return 0
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
