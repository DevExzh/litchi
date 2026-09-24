#!/usr/bin/env python3
"""Offline ECMA Theme and MS-ODRAWXML family XSD validation.

Accept XML files or XLSB ZIPs. No network, extraction, or Office application
automation is used. Both documented and explicitly fixture-backed extension
identifiers are inventoried. Unknown extension payloads retain ECMA lax rules.
"""

import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys
import zipfile

from lxml import etree

ROOT = next(p for p in Path(__file__).resolve().parents if (p / "crates").is_dir())
FAMILY = "http://schemas.microsoft.com/office/thememl/2012/main"
URIS = {FAMILY, "{05A4C25C-085E-4340-85A3-A5531E510DB2}"}
LEGACY = Path(__file__).resolve().parent.parent / "xlsb-theme/verify-theme-schema.py"
sys.dont_write_bytecode = True
spec = importlib.util.spec_from_file_location("theme_schema", LEGACY)
theme_schema = importlib.util.module_from_spec(spec)
spec.loader.exec_module(theme_schema)


def family_schema(namespace):
    import io

    part, conformance = theme_schema.NAMESPACES[namespace]
    archive = ROOT / f"3rdparty/specs/ECMA-376/ECMA-376-{part}_5th_edition_december_2016.zip"
    source = ROOT / "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.17 http---schemas.microsoft.com-office-thememl-2012-main Schema.md"
    with zipfile.ZipFile(archive) as outer:
        encoded = outer.read(f"OfficeOpenXML-XMLSchema-{conformance}.zip")
    with zipfile.ZipFile(io.BytesIO(encoded)) as inner:
        schemas = {n: inner.read(n) for n in inner.namelist() if n.endswith(".xsd")}
    schema = "\n".join(re.sub(r"^\s*(?:\d+\. )?", "", line)
                       for line in source.read_text().splitlines() if "\\<" in line)
    schema = schema.replace("\\<", "<").replace("\\>", ">")
    # Microsoft's imported a:ST_Guid maps to the equivalent ECMA shared type.
    shared = ("http://purl.oclc.org/ooxml/officeDocument/sharedTypes" if part == 1
              else "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes")
    schema = schema.replace('xmlns:a=', f'xmlns:s="{shared}" xmlns:a=', 1)
    schema = schema.replace("a:ST_Guid", "s:ST_Guid")
    schema = schema.replace("oartbasestylesheet.xsd", "dml-main.xsd").replace("oartbasetypes.xsd", "dml-main.xsd").replace("orel.xsd", "shared-relationshipReference.xsd")
    if part == 1:
        schema = schema.replace("http://schemas.openxmlformats.org/drawingml/2006/main", namespace)
        schema = schema.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships")
    schema = schema.replace("<xsd:complexType", f'<xsd:import namespace="{shared}" schemaLocation="shared-commonSimpleTypes.xsd"/><xsd:complexType', 1)

    class Resolver(etree.Resolver):
        def resolve(self, url, public_id, context):
            if url not in schemas:
                raise ValueError(f"unrecognized schema import: {url}")
            return self.resolve_string(schemas[url], context, base_url=url)

    parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
    parser.resolvers.add(Resolver())
    validator = etree.XMLSchema(etree.fromstring(schema.encode(), parser))
    return validator, hashlib.sha256(source.read_bytes()).hexdigest()


def main():
    reports = []
    cache = {}
    for name in sys.argv[1:]:
        path = Path(name)
        if path.suffix.lower() == ".xlsb":
            with zipfile.ZipFile(path) as archive:
                parts = [(n, archive.read(n)) for n in archive.namelist()
                         if "/theme/" in n and n.endswith(".xml") and "/_rels/" not in n]
            if not parts:
                raise ValueError(f"no Theme part in {name}")
        else:
            parts = [(None, path.read_bytes())]
        for member, data in parts:
            doc = etree.fromstring(data, etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False))
            if doc.getroottree().docinfo.doctype:
                raise ValueError("DTD outside profile")
            ns = etree.QName(doc).namespace
            if etree.QName(doc).localname != "theme" or ns not in theme_schema.NAMESPACES:
                raise ValueError("unexpected Theme root")
            if ns not in cache:
                cache[ns] = (*theme_schema.schema_for(ns), *family_schema(ns))
            theme, provenance, family, family_hash = cache[ns]
            theme.assertValid(doc)
            found = []
            for ext in doc.findall(f"{{{ns}}}extLst/{{{ns}}}ext"):
                if ext.get("uri", "").strip(" \t\r\n") not in URIS:
                    continue
                for child in ext.findall(f"{{{FAMILY}}}themeFamily"):
                    family.assertValid(child)
                    found.append({"extension_uri": ext.get("uri"), "attributes": dict(child.attrib)})
            if len(found) > 1:
                raise ValueError("ambiguous supported family owners")
            reports.append({"path": name, "member": member, "bytes": len(data),
                            "input_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                            "sha256": hashlib.sha256(data).hexdigest(), "valid": True,
                            "families": found, "family_schema_source_sha256": family_hash,
                            **provenance})
    if not reports:
        raise ValueError("at least one XML or XLSB input is required")
    print(json.dumps({"lxml_version": etree.LXML_VERSION, "reports": reports}, indent=2))


if __name__ == "__main__":
    main()
