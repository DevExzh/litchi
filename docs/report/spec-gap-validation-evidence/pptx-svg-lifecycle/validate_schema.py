#!/usr/bin/env python3
"""Offline Transitional slide + MS-ODRAWXML SVG extension schema validation."""

import hashlib
import io
import json
from pathlib import Path
import re
import sys
import zipfile

from lxml import etree

ROOT = next(p for p in Path(__file__).resolve().parents if (p / "crates").is_dir())
PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def pictures(container):
    for child in container:
        if child.tag == f"{{{PML}}}pic":
            yield child
        elif child.tag == f"{{{PML}}}grpSp":
            yield from pictures(child)


def main():
    archive = ROOT / "3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip"
    source = ROOT / "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.24 http---schemas.microsoft.com-office-drawing-2016-SVG-main Schema.md"
    with zipfile.ZipFile(archive) as outer:
        packed = outer.read("OfficeOpenXML-XMLSchema-Transitional.zip")
    with zipfile.ZipFile(io.BytesIO(packed)) as inner:
        schemas = {name: inner.read(name) for name in inner.namelist() if name.endswith(".xsd")}
    extension = "\n".join(re.sub(r"^\s*(?:\d+\. )?", "", line)
                          for line in source.read_text().splitlines() if "\\<" in line)
    extension = extension.replace("\\<", "<").replace("\\>", ">")
    # Microsoft splits the ECMA DrawingML module; AG_Blob has the same shape.
    extension = extension.replace("oartbasetypes.xsd", "dml-main.xsd")
    extension = extension.replace("orel.xsd", "shared-relationshipReference.xsd")

    class Resolver(etree.Resolver):
        def resolve(self, url, public_id, context):
            if url not in schemas:
                raise ValueError(f"unrecognized offline schema import: {url}")
            return self.resolve_string(schemas[url], context, base_url=url)

    parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
    parser.resolvers.add(Resolver())
    slide_schema = etree.XMLSchema(etree.fromstring(schemas["pml.xsd"], parser))
    svg_schema = etree.XMLSchema(etree.fromstring(extension.encode(), parser))
    reports = []
    for name in sys.argv[1:]:
        data = Path(name).read_bytes()
        document = etree.fromstring(data, parser)
        if document.getroottree().docinfo.doctype or document.tag != f"{{{PML}}}sld":
            raise ValueError("expected a DTD-free Transitional slide")
        slide_schema.assertValid(document)
        owners = []
        tree = document.find(f"{{{PML}}}cSld/{{{PML}}}spTree")
        if tree is None:
            raise ValueError("slide lacks direct shape tree")
        for picture in pictures(tree):
            blip = picture.find(f"{{{PML}}}blipFill/{{{DML}}}blip")
            if blip is None:
                continue
            for ext in blip.findall(f"{{{DML}}}extLst/{{{DML}}}ext"):
                if re.sub(r"[ \t\r\n]+", " ", ext.get("uri", "")).strip(" ") != URI:
                    continue
                selected = ext.findall(f"{{{SVG}}}svgBlip")
                if len(selected) != 1:
                    raise ValueError("recognized SVG extension must have one direct owner")
                svg_schema.assertValid(selected[0])
                owners.append(dict(selected[0].attrib))
        reports.append({"path": name, "sha256": digest(data), "bytes": len(data),
                        "valid": True, "svg_owners": owners})
    if not reports:
        raise ValueError("provide generated slide XML files")
    print(json.dumps({"scope": "Generated slide and direct SVG extension grammar; no native acceptance or rendering proof",
                      "ecma_archive_sha256": digest(archive.read_bytes()),
                      "schema_source_sha256": digest(source.read_bytes()),
                      "lxml_version": etree.LXML_VERSION, "reports": reports}, indent=2))


if __name__ == "__main__":
    main()
