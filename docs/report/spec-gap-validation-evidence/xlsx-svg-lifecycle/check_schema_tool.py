#!/usr/bin/env python3
"""Check the offline validator using a native drawing and explicit mutations."""

import copy
import json
from pathlib import Path
import zipfile

from lxml import etree

from validate_schema import ROOT, XDR, DML, digest, schema_context, validate


def main():
    here = Path(__file__).resolve().parent
    fixture = ROOT / "3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx"
    raw = fixture.read_bytes()
    assert digest(raw) == "0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72"
    with zipfile.ZipFile(fixture) as archive:
        data = archive.read("xl/drawings/drawing1.xml")
    context = schema_context()
    original = etree.fromstring(data)
    reports = [validate(data, "native:xl/drawings/drawing1.xml", context)]
    assert len(reports[0]["svg_owners"]) == 2
    for kind in ("oneCellAnchor", "absoluteAnchor"):
        document = copy.deepcopy(original)
        anchor = document[0]
        anchor.tag = f"{{{XDR}}}{kind}"
        anchor.attrib.pop("editAs", None)
        anchor.remove(anchor.find(f"{{{XDR}}}to"))
        if kind == "absoluteAnchor":
            anchor.remove(anchor.find(f"{{{XDR}}}from"))
            anchor.insert(0, etree.Element(f"{{{XDR}}}pos", x="10", y="20"))
        anchor.insert(1, etree.Element(f"{{{XDR}}}ext", cx="100", cy="200"))
        report = validate(etree.tostring(document), f"synthetic:{kind}", context)
        assert report["anchor_counts"][kind] == 1
        assert len(report["svg_owners"]) == 2
        reports.append(report)

    rejected = []
    for kind in ("missing-marker", "container-text", "duplicate-svg-owner"):
        document = copy.deepcopy(original)
        if kind == "missing-marker":
            document[0].remove(document[0].find(f"{{{XDR}}}from"))
        else:
            extensions = document[0].find(
                f"{{{XDR}}}pic/{{{XDR}}}blipFill/{{{DML}}}blip/{{{DML}}}extLst")
            assert extensions is not None
            if kind == "container-text":
                extensions.text = "invalid element-only content"
            else:
                extensions.append(copy.deepcopy(extensions[0]))
        try:
            validate(etree.tostring(document), f"negative:{kind}", context)
        except (ValueError, etree.DocumentInvalid):
            rejected.append(kind)
        else:
            raise AssertionError(f"validator accepted {kind}")
    print(json.dumps({
        "passed": True,
        "scope": "Validator smoke checks; no XLSX lifecycle implementation evidence",
        "inputs": {str(path.relative_to(ROOT)): digest(path.read_bytes()) for path in
                   (fixture, here / "validate_schema.py", Path(__file__).resolve())},
        **context[3], "accepted": reports, "rejected": rejected,
    }, indent=2))


if __name__ == "__main__":
    main()
