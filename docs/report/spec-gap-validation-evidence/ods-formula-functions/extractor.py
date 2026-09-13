#!/usr/bin/env python3
"""Extract the ODF 1.4 Part 4 chapter-6 function catalog.

This intentionally uses only the Python standard library.  The normative HTML
is inside the ODF specification ZIP, so the command is reproducible without a
network request or a third-party HTML parser::

    python3 extractor.py /path/to/OpenDocument-v1.4-os.zip \
      --output normative-functions.tsv

The chapter contains 15 ``General`` subsections across the 16 function
families.  A catalog row is emitted for every other 6.5--6.20 ``h3`` whose body contains a
``Syntax:`` paragraph.  The visible heading is the source of the function
name; this handles the HTML's split ``I``/``MSEC`` text while retaining the
published section anchor.  Explicit name anchors are included when present;
the source has two adjacent anchor omissions/placements, so the section anchor
is the authoritative link for those rows.
"""

from __future__ import annotations

import argparse
import html
import re
import sys
import zipfile
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path


HTML_MEMBER = "part4-formula/OpenDocument-v1.4-os-part4-formula.html"
SECTION_RE = re.compile(r"^(6\.(?P<family>\d+)\.(?P<item>\d+))\s+(?P<name>.+)$")
NAME_RE = re.compile(r"^[A-Z][A-Z0-9._]*$")
SECTION_ANCHOR_RE = re.compile(r"^a_6_\d+_\d+_.+$")
REF_ANCHOR_RE = re.compile(r"^__RefHeading")


@dataclass
class Heading:
    title: str
    ids: list[str]
    body_text: str


class ChapterHeadingParser(HTMLParser):
    """Collect h3 text, ids, and text up to the next h3."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._heading: dict[str, object] | None = None
        self._inside_h3 = False
        self._body: list[str] = []
        self._headings: list[Heading] = []

    @property
    def headings(self) -> list[Heading]:
        return self._headings

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "h3":
            if self._inside_h3:
                self._finish_heading()
            self._inside_h3 = True
            self._heading = {"text": [], "ids": []}
            self._body = []
            return

        if self._inside_h3 and self._heading is not None:
            for key, value in attrs:
                if key == "id" and value is not None:
                    ids = self._heading["ids"]
                    assert isinstance(ids, list)
                    ids.append(value)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        # The published HTML uses both ``<a ...>`` and XHTML-style ``<a .../>``.
        self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag: str) -> None:
        if tag == "h3" and self._inside_h3:
            self._finish_heading()

    def handle_data(self, data: str) -> None:
        if self._inside_h3 and self._heading is not None:
            text = self._heading["text"]
            assert isinstance(text, list)
            text.append(data)
        elif self._headings:
            self._body.append(data)

    def _finish_heading(self) -> None:
        assert self._heading is not None
        text = self._heading["text"]
        ids = self._heading["ids"]
        assert isinstance(text, list)
        assert isinstance(ids, list)
        title = re.sub(r"\s+", " ", "".join(text)).strip()
        self._headings.append(Heading(title, list(ids), ""))
        self._inside_h3 = False
        self._heading = None
        self._body = []


def _parse_headings(source: bytes) -> list[Heading]:
    parser = ChapterHeadingParser()
    parser.feed(source.decode("utf-8"))
    parser.close()

    # Search the text between adjacent h3 tags for the normative Syntax label.
    # This check is deliberately conservative: the extractor only needs to
    # reject pseudo-headings, not to interpret the signature.
    chunks = re.split(r"(?=<h3\b)", source.decode("utf-8"))
    syntax_by_title: dict[str, bool] = {}
    for chunk in chunks:
        if not chunk.startswith("<h3"):
            continue
        end = chunk.find("</h3>")
        if end < 0:
            continue
        head = chunk[: end + len("</h3>")]
        title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", head)).strip()
        # Search only the following chunk; the next split starts at the next
        # h3, so this text is the complete body for this heading.
        tail = chunk[end + len("</h3>") :]
        tail = re.split(r"(?=<h3\b)", tail, maxsplit=1)[0]
        plain = html.unescape(re.sub(r"<[^>]+>", " ", tail))
        syntax_by_title[title] = bool(re.search(r"\bSyntax\s*:", plain))

    for heading in parser.headings:
        syntax_by_title.setdefault(heading.title, False)
    return [
        Heading(h.title, h.ids, "1" if syntax_by_title.get(h.title, False) else "0")
        for h in parser.headings
    ]


def _canonical_name(raw: str) -> str:
    # IMSEC is split across two text nodes in the HTML.  Removing presentation
    # whitespace is safe here because a function name is restricted by NAME_RE.
    return raw.replace(" ", "")


def extract(zip_path: Path) -> list[dict[str, str]]:
    with zipfile.ZipFile(zip_path) as bundle:
        source = bundle.read(HTML_MEMBER)
    headings = _parse_headings(source)
    rows: list[dict[str, str]] = []
    for heading in headings:
        match = SECTION_RE.fullmatch(heading.title)
        if match is None:
            continue
        family = int(match.group("family"))
        if not 5 <= family <= 20:
            continue
        raw_name = match.group("name")
        name = _canonical_name(raw_name)
        if name == "General":
            continue
        if not NAME_RE.fullmatch(name):
            raise ValueError(f"unexpected chapter-6 heading {heading.title!r}")
        section = match.group(1)
        section_anchor = next(
            (anchor for anchor in heading.ids if SECTION_ANCHOR_RE.fullmatch(anchor)),
            "-",
        )
        expected_anchor = name.replace(".", "_")
        name_anchor = next(
            (
                anchor
                for anchor in heading.ids
                if anchor == name or anchor == expected_anchor
            ),
            "-",
        )
        if heading.body_text != "1":
            raise ValueError(f"function heading lacks Syntax paragraph: {heading.title!r}")
        if section_anchor == "-":
            raise ValueError(f"function heading lacks section anchor: {heading.title!r}")
        rows.append(
            {
                "section": section,
                "name": name,
                "html_section_anchor": section_anchor,
                "html_name_anchor": name_anchor,
                "syntax_present": heading.body_text,
                "catalog_status": (
                    "standard_legacy_prefixed" if name.startswith("LEGACY.") else "standard"
                ),
                "appendix_disposition": "new-v1.4"
                if name == "EASTERSUNDAY"
                else "changed-v1.4"
                if name
                in {
                    "CONVERT",
                    "COUNTA",
                    "INDEX",
                    "ISBLANK",
                    "ISFORMULA",
                    "ISLOGICAL",
                    "ISNONTEXT",
                    "ISNUMBER",
                    "ISREF",
                    "ISTEXT",
                    "NPER",
                    "PMT",
                }
                else "unchanged-v1.4",
                "normative_aliases": "none-listed",
            }
        )
    if len(rows) != 393:
        raise ValueError(f"expected 393 chapter-6 function rows, found {len(rows)}")
    names = [row["name"] for row in rows]
    if len(set(names)) != len(names):
        raise ValueError("duplicate chapter-6 function name")
    return rows


def write_tsv(rows: list[dict[str, str]], output: Path) -> None:
    fields = [
        "section",
        "name",
        "html_section_anchor",
        "html_name_anchor",
        "syntax_present",
        "catalog_status",
        "appendix_disposition",
        "normative_aliases",
    ]
    with output.open("w", encoding="utf-8", newline="") as stream:
        stream.write("\t".join(fields) + "\n")
        for row in rows:
            stream.write("\t".join(row[field] for field in fields) + "\n")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("zip_path", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    rows = extract(args.zip_path)
    write_tsv(rows, args.output)
    print(f"wrote {len(rows)} functions to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
