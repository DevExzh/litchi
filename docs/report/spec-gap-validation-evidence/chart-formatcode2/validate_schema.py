#!/usr/bin/env python3
"""Offline MS-ODRAWXML 5.42 element validation; no host-placement claim."""

import argparse
import hashlib
import io
import json
from pathlib import Path
import re
import zipfile

from lxml import etree

ROOT = next(p for p in Path(__file__).resolve().parents if (p / "crates").is_dir())
NAMESPACE = "http://schemas.microsoft.com/office/drawing/2015/06/chart"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def main():
    arguments = argparse.ArgumentParser(description=__doc__)
    arguments.add_argument("--spec-root", type=Path, default=ROOT)
    arguments.add_argument("xml", nargs="+")
    args = arguments.parse_args()
    archive = args.spec_root / "3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip"
    source = args.spec_root / "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.42 http---schemas.microsoft.com-office-drawing-2015-06-chart Schema.md"
    with zipfile.ZipFile(archive) as outer:
        packed = outer.read("OfficeOpenXML-XMLSchema-Transitional.zip")
    with zipfile.ZipFile(io.BytesIO(packed)) as inner:
        shared = inner.read("shared-commonSimpleTypes.xsd")
    schema = "\n".join(
        re.sub(r"^\s*(?:\d+\. )?", "", line)
        for line in source.read_text().splitlines() if "\\<" in line
    ).replace("\\<", "<").replace("\\>", ">")
    # Microsoft chart.xsd names this shared type through c:. ECMA places the
    # same normative ST_Xstring in shared-commonSimpleTypes.xsd instead.
    schema = schema.replace(
        "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    ).replace('schemaLocation="chart.xsd"', 'schemaLocation="shared-commonSimpleTypes.xsd"')

    class Resolver(etree.Resolver):
        def resolve(self, url, public_id, context):
            if url != "shared-commonSimpleTypes.xsd":
                raise ValueError(f"unexpected offline import: {url}")
            return self.resolve_string(shared, context, base_url=url)

    parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
    parser.resolvers.add(Resolver())
    validator = etree.XMLSchema(etree.fromstring(schema.encode(), parser))
    reports = []
    for name in args.xml:
        data = Path(name).read_bytes()
        document = etree.fromstring(data, parser)
        if document.getroottree().docinfo.doctype:
            raise ValueError("DTD outside profile")
        if document.tag != f"{{{NAMESPACE}}}formatcode2":
            raise ValueError("unexpected element root")
        validator.assertValid(document)
        reports.append({"path": name, "bytes": len(data), "sha256": digest(data), "valid": True})
    if not reports:
        raise ValueError("provide at least one generated XML fragment")
    print(json.dumps({
        "scope": "Global formatcode2 element XSD only; no string-decoding or host-placement proof",
        "schema_source_sha256": digest(source.read_bytes()),
        "ecma_archive_sha256": digest(archive.read_bytes()),
        "shared_schema_sha256": digest(shared),
        "lxml_version": etree.LXML_VERSION,
        "reports": reports,
    }, indent=2))


if __name__ == "__main__":
    main()
