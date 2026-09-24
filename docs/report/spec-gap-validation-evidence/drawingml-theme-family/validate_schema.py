#!/usr/bin/env python3
"""Validate supplied fragments against the vendored themeFamily and ECMA XSDs.

Requires lxml. No network access, schema download, or temporary extraction.
Microsoft's a:ST_Guid import is mapped to ECMA's shared ST_Guid owner;
the lexical type itself is unchanged. This is schema evidence, not Office
application acceptance or validation of all opaque extension semantics.
"""

import argparse
import hashlib
import io
import json
from pathlib import Path
import re
import zipfile

from lxml import etree


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fragments", nargs="+", type=Path)
    args = parser.parse_args()
    root = next(p for p in Path(__file__).resolve().parents if (p / "crates").is_dir())
    archive = root / "3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip"
    source = root / "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.17 http---schemas.microsoft.com-office-thememl-2012-main Schema.md"
    with zipfile.ZipFile(archive) as outer:
        with zipfile.ZipFile(io.BytesIO(outer.read("OfficeOpenXML-XMLSchema-Transitional.zip"))) as inner:
            schemas = {name: inner.read(name) for name in inner.namelist() if name.endswith(".xsd")}
    lines = source.read_text().splitlines()
    schema = "\n".join(re.sub(r"^\s*(?:\d+\. )?", "", line) for line in lines if "\\<" in line)
    schema = schema.replace("\\<", "<").replace("\\>", ">")
    schema = schema.replace('xmlns:a=', 'xmlns:s="http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes" xmlns:a=', 1)
    schema = schema.replace("a:ST_Guid", "s:ST_Guid")
    schema = schema.replace("oartbasestylesheet.xsd", "dml-main.xsd").replace("oartbasetypes.xsd", "dml-main.xsd").replace("orel.xsd", "shared-relationshipReference.xsd")
    schema = schema.replace("<xsd:complexType", '<xsd:import namespace="http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes" schemaLocation="shared-commonSimpleTypes.xsd"/><xsd:complexType', 1)

    class Resolver(etree.Resolver):
        def resolve(self, url, public_id, context):
            name = url.rsplit("/", 1)[-1]
            if name not in schemas:
                raise ValueError(f"Unrecognized schema import: {url}")
            return self.resolve_string(schemas[name], context)

    xml_parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
    xml_parser.resolvers.add(Resolver())
    validator = etree.XMLSchema(etree.fromstring(schema.encode(), xml_parser))
    results = []
    for path in args.fragments:
        data = path.read_bytes()
        document = etree.fromstring(data, xml_parser)
        if document.getroottree().docinfo.doctype:
            raise ValueError(f"DTD is outside the fragment profile: {path}")
        validator.assertValid(document)
        results.append({"path": str(path), "sha256": hashlib.sha256(data).hexdigest(), "valid": True})
    print(json.dumps({"schema_source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(), "ecma_archive_sha256": hashlib.sha256(archive.read_bytes()).hexdigest(), "fragments": results}, indent=2))


if __name__ == "__main__":
    main()
