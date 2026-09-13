#!/usr/bin/env python3
"""Validate generated DDE owners against the bundled ODF 1.4 schema."""

import argparse
import copy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spec", type=Path, required=True)
    parser.add_argument("--content", type=Path, required=True)
    args = parser.parse_args()
    with ZipFile(args.spec) as archive:
        schema_bytes = archive.read("schemas/OpenDocument-v1.4-schema.rng")
    source = args.content.read_bytes()
    xml_parser = etree.XMLParser(resolve_entities=False, no_network=True)
    document = etree.fromstring(source, xml_parser)
    schema = etree.fromstring(schema_bytes, xml_parser)
    namespaces = {
        "r": "http://relaxng.org/ns/structure/1.0",
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    }
    owners = []
    for definition, query in [
        ("office-dde-source", "//office:dde-source"),
        ("table-dde-links", "//table:dde-links"),
    ]:
        focused = copy.deepcopy(schema)
        start = focused.find("r:start", namespaces)
        assert start is not None
        start.clear()
        etree.SubElement(start, "{%s}ref" % namespaces["r"], name=definition)
        validator = etree.RelaxNG(focused)
        for ordinal, node in enumerate(document.xpath(query, namespaces=namespaces)):
            valid = validator.validate(node)
            owners.append({
                "definition": definition,
                "ordinal": ordinal,
                "valid": valid,
                "diagnostics": [str(error) for error in validator.error_log],
            })
    complete = etree.RelaxNG(schema)
    full_valid = complete.validate(document)
    cache_semantics = []
    for ordinal, table in enumerate(document.xpath(
        "//table:dde-link/table:table", namespaces=namespaces
    )):
        issues = []
        for node in table.iter():
            if not isinstance(node.tag, str):
                continue
            if any(etree.QName(name).localname.endswith("style-name") for name in node.attrib):
                issues.append("DDE cache contains style information")
            if node.tag == "{%s}table-cell" % namespaces["table"]:
                if len(node) or (node.text and node.text.strip()):
                    issues.append("DDE cache cell is not empty")
                if node.get("{%s}value-type" % namespaces["office"]) == "string" and node.get(
                    "{%s}string-value" % namespaces["office"]
                ) is None:
                    issues.append("authored string cache cell has no attribute value")
        cache_semantics.append({
            "ordinal": ordinal,
            "valid": not issues,
            "diagnostics": issues,
            "reference": "ODF 1.4 Part 3 14.7.4",
        })
    result = {
        "content_sha256": hashlib.sha256(source).hexdigest(),
        "schema_sha256": hashlib.sha256(schema_bytes).hexdigest(),
        "whole_content_valid": full_valid,
        "whole_content_diagnostics": [str(error) for error in complete.error_log],
        "owners": owners,
        "cache_semantics": cache_semantics,
    }
    print(json.dumps(result, indent=2))
    return 0 if owners and all(owner["valid"] for owner in owners + cache_semantics) else 1


if __name__ == "__main__":
    raise SystemExit(main())
