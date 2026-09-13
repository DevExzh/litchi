"""Validate the four native-input metadata owners against the bundled ODF RNG."""

import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spec", type=Path, required=True)
    parser.add_argument("--package", type=Path, required=True)
    args = parser.parse_args()
    member = "schemas/OpenDocument-v1.4-schema.rng"
    with ZipFile(args.spec) as archive:
        schema_bytes = archive.read(member)
    with ZipFile(args.package) as archive:
        content_bytes = archive.read("content.xml")
    schema = etree.fromstring(schema_bytes)
    content = etree.fromstring(content_bytes)
    namespaces = {
        "r": "http://relaxng.org/ns/structure/1.0",
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    }
    owners = []
    for name in ("consolidation", "label-ranges", "cell-range-source", "detective"):
        definitions = schema.xpath(
            "//r:define[r:element[@name=$name]]",
            namespaces=namespaces,
            name="table:" + name,
        )
        assert len(definitions) == 1, (name, len(definitions))
        owner_schema = deepcopy(schema)
        start = owner_schema.find("r:start", namespaces)
        start.clear()
        etree.SubElement(
            start,
            "{" + namespaces["r"] + "}ref",
            name=definitions[0].get("name"),
        )
        validator = etree.RelaxNG(owner_schema)
        nodes = content.findall(".//table:" + name, namespaces)
        checks = []
        for node in nodes:
            valid = validator.validate(node)
            checks.append(
                {"valid": valid, "errors": [str(error) for error in validator.error_log]}
            )
        owners.append({"owner": name, "count": len(nodes), "checks": checks})
    whole = etree.RelaxNG(schema)
    whole_valid = whole.validate(content)
    result = {
        "package": str(args.package),
        "package_sha256": hashlib.sha256(args.package.read_bytes()).hexdigest(),
        "content_xml_sha256": hashlib.sha256(content_bytes).hexdigest(),
        "schema_archive": str(args.spec),
        "schema_member": member,
        "schema_sha256": hashlib.sha256(schema_bytes).hexdigest(),
        "lxml_version": etree.LXML_VERSION,
        "owners": owners,
        "whole_content_xml_valid": whole_valid,
        "whole_content_xml_errors": [str(error) for error in whole.error_log],
        "scope": "Owner datatypes and structures; not whole-package validation.",
    }
    print(json.dumps(result, indent=2))
    return 0 if all(
        owner["count"] == 1 and all(check["valid"] for check in owner["checks"])
        for owner in owners
    ) else 1


if __name__ == "__main__":
    raise SystemExit(main())
