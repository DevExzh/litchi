#!/usr/bin/env python3
"""Offline Transitional worksheet drawing and MS-ODRAWXML SVG schema checks."""

import hashlib
import io
import json
from pathlib import Path
import re
import sys
import zipfile

from lxml import etree

ROOT = next(path for path in Path(__file__).resolve().parents if (path / "crates").is_dir())
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
ANCHORS = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")
MAX_XML_BYTES = 32 * 1024 * 1024


def digest(data):
    return hashlib.sha256(data).hexdigest()


def schema_context():
    archive = ROOT / "3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip"
    source = ROOT / "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.24 http---schemas.microsoft.com-office-drawing-2016-SVG-main Schema.md"
    with zipfile.ZipFile(archive) as outer:
        packed = outer.read("OfficeOpenXML-XMLSchema-Transitional.zip")
    with zipfile.ZipFile(io.BytesIO(packed)) as inner:
        schemas = {name: inner.read(name) for name in inner.namelist() if name.endswith(".xsd")}
    extension = "\n".join(re.sub(r"^\s*(?:\d+\. )?", "", line)
                          for line in source.read_text().splitlines() if "\\<" in line)
    extension = extension.replace("\\<", "<").replace("\\>", ">")
    extension = extension.replace("oartbasetypes.xsd", "dml-main.xsd")
    extension = extension.replace("orel.xsd", "shared-relationshipReference.xsd")

    class Resolver(etree.Resolver):
        def resolve(self, url, public_id, context):
            if url not in schemas:
                raise ValueError(f"unrecognized offline schema import: {url}")
            return self.resolve_string(schemas[url], context, base_url=url)

    parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
    parser.resolvers.add(Resolver())
    drawing_schema = etree.XMLSchema(etree.fromstring(schemas["dml-spreadsheetDrawing.xsd"], parser))
    svg_schema = etree.XMLSchema(etree.fromstring(extension.encode(), parser))
    provenance = {
        "ecma_archive_sha256": digest(archive.read_bytes()),
        "svg_schema_source_sha256": digest(source.read_bytes()),
        "lxml_version": etree.LXML_VERSION,
    }
    return parser, drawing_schema, svg_schema, provenance


def validate(data, name, context):
    if len(data) > MAX_XML_BYTES:
        raise ValueError("drawing XML exceeds the validation profile limit")
    parser, drawing_schema, svg_schema, _ = context
    document = etree.fromstring(data, parser)
    if document.getroottree().docinfo.doctype or document.tag != f"{{{XDR}}}wsDr":
        raise ValueError("expected a DTD-free Transitional worksheet drawing")
    drawing_schema.assertValid(document)
    anchor_counts = {kind: 0 for kind in ANCHORS}
    owners = []
    for anchor in document:
        kind = next((kind for kind in ANCHORS if anchor.tag == f"{{{XDR}}}{kind}"), None)
        if kind is None:
            continue
        anchor_counts[kind] += 1
        picture = anchor.find(f"{{{XDR}}}pic")
        if picture is None:
            continue
        blip = picture.find(f"{{{XDR}}}blipFill/{{{DML}}}blip")
        if blip is None:
            continue
        recognized = [ext for ext in blip.findall(f"{{{DML}}}extLst/{{{DML}}}ext")
                      if ext.get("uri", "").strip(" \t\r\n") == URI]
        if len(recognized) > 1:
            raise ValueError("picture has duplicate recognized SVG extensions")
        for ext in recognized:
            selected = ext.findall(f"{{{SVG}}}svgBlip")
            if len(selected) != 1:
                raise ValueError("recognized SVG extension must have one direct owner")
            svg_schema.assertValid(selected[0])
            owners.append({"anchor": kind, "attributes": dict(selected[0].attrib)})
    return {"path": name, "sha256": digest(data), "bytes": len(data), "valid": True,
            "anchor_counts": anchor_counts, "svg_owners": owners}


def main():
    context = schema_context()
    reports = []
    for name in sys.argv[1:]:
        path = Path(name)
        if path.stat().st_size > MAX_XML_BYTES:
            raise ValueError("drawing XML exceeds the validation profile limit")
        reports.append(validate(path.read_bytes(), name, context))
    if not reports:
        raise ValueError("provide generated worksheet drawing XML files")
    print(json.dumps({"scope": "Transitional drawing and direct SVG extension grammar; no package graph, Strict dialect, native acceptance or rendering proof",
                      **context[3], "reports": reports}, indent=2))


if __name__ == "__main__":
    main()
