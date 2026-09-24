#!/usr/bin/env python3
"""Independently validate retained theme XML against the vendored ECMA schemas.

Usage: verify-theme-schema.py [theme.xml | workbook.xlsb ...]
With no arguments, validate the two native profiling fixtures. Schema imports
resolve only from the nested vendored schema archive; no extraction or network
access is used. This validates XML, not workbook topology or Office rendering.
"""

import hashlib
import io
import json
from pathlib import Path
import sys
import zipfile

from lxml import etree

ROOT = Path(__file__).resolve().parents[4]
NAMESPACES = {
    "http://schemas.openxmlformats.org/drawingml/2006/main": (4, "Transitional"),
    "http://purl.oclc.org/ooxml/drawingml/main": (1, "Strict"),
}


def schema_for(namespace):
    part, conformance = NAMESPACES[namespace]
    archive = ROOT / f"3rdparty/specs/ECMA-376/ECMA-376-{part}_5th_edition_december_2016.zip"
    with zipfile.ZipFile(archive) as outer:
        encoded = outer.read(f"OfficeOpenXML-XMLSchema-{conformance}.zip")
    with zipfile.ZipFile(io.BytesIO(encoded)) as inner:
        sources = {name: inner.read(name) for name in inner.namelist() if name.endswith(".xsd")}

    class Resolver(etree.Resolver):
        def resolve(self, url, public_id, context):
            if url not in sources:
                raise ValueError(f"schema import outside vendored archive: {url}")
            return self.resolve_string(sources[url], context, base_url=url)

    parser = etree.XMLParser(no_network=True, resolve_entities=False)
    parser.resolvers.add(Resolver())
    schema = etree.XMLSchema(etree.fromstring(sources["dml-main.xsd"], parser, base_url="dml-main.xsd"))
    return schema, {
        "conformance": conformance,
        "schema_archive_sha256": hashlib.sha256(encoded).hexdigest(),
        "main_schema_sha256": hashlib.sha256(sources["dml-main.xsd"]).hexdigest(),
    }


def main():
    paths = sys.argv[1:] or [
        "test-data/poi/test-data/spreadsheet/testVarious.xlsb",
        "test-data/ooxml/xlsb/62815.xlsb",
    ]
    schemas = {}
    reports = []
    for name in paths:
        path = Path(name)
        if path.suffix.lower() == ".xlsb":
            with zipfile.ZipFile(path) as archive:
                parts = [(part, archive.read(part)) for part in archive.namelist()
                         if "/theme/" in part and part.endswith(".xml") and "/_rels/" not in part]
            if not parts:
                raise ValueError(f"no theme XML fixture parts in {name}")
        else:
            parts = [(None, path.read_bytes())]
        for part, data in parts:
            parser = etree.XMLParser(no_network=True, resolve_entities=False)
            document = etree.fromstring(data, parser)
            if document.getroottree().docinfo.doctype:
                raise ValueError("theme fixture contains a DTD")
            namespace = etree.QName(document).namespace
            if namespace not in NAMESPACES or etree.QName(document).localname != "theme":
                raise ValueError(f"unexpected theme root: {document.tag}")
            if namespace not in schemas:
                schemas[namespace] = schema_for(namespace)
            schema, provenance = schemas[namespace]
            schema.assertValid(document)
            ns = {"a": namespace}
            palette = document.find("a:themeElements/a:clrScheme", ns)
            fonts = document.find("a:themeElements/a:fontScheme", ns)
            semantic = {
                "name": document.get("name"),
                "palette_name": palette.get("name"),
                "colors": {
                    etree.QName(slot).localname: [
                        {"kind": etree.QName(value).localname, "attributes": dict(value.attrib)}
                        for value in slot if isinstance(value.tag, str)
                    ] for slot in palette if isinstance(slot.tag, str)
                },
                "font_scheme_name": fonts.get("name"),
                "major_latin": fonts.find("a:majorFont/a:latin", ns).get("typeface"),
                "minor_latin": fonts.find("a:minorFont/a:latin", ns).get("typeface"),
            }
            reports.append({"path": name, "part": part, "bytes": len(data),
                            "sha256": hashlib.sha256(data).hexdigest(),
                            "valid": True, "semantic": semantic, **provenance})
    print(json.dumps({"lxml_version": etree.LXML_VERSION, "reports": reports}, indent=2))


if __name__ == "__main__":
    main()
