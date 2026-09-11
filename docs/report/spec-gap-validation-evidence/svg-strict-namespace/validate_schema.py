#!/usr/bin/env python3
"""Validate Strict core XML beside the unmodified Transitional MS SVG child schema.

The MS-ODRAWXML §5.24 text is used byte-for-byte for its target namespace and
namespace declarations.  Only its two relative import locations are resolved
to the vendored Transitional ECMA schema member names.  This keeps the
negative Strict-attribute control meaningful: changing the MS schema namespace
to make the control pass would invalidate the evidence.
"""

import hashlib
import io
import json
from pathlib import Path
import re
import sys
import zipfile

from lxml import etree


HERE = Path(__file__).resolve().parent
ROOT = next(path for path in HERE.parents if (path / "crates").is_dir())
PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
STRICT_PML = "http://purl.oclc.org/ooxml/presentationml/main"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
STRICT_DML = "http://purl.oclc.org/ooxml/drawingml/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
STRICT_XDR = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships"
SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SVG_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
MAX_XML_BYTES = 32 * 1024 * 1024

STRICT_ARCHIVE = ROOT / "3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip"
TRANSITIONAL_ARCHIVE = ROOT / "3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip"
MS_SCHEMA = ROOT / (
    "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/"
    "5.24 http---schemas.microsoft.com-office-drawing-2016-SVG-main Schema.md"
)
XLSX_FIXTURE = ROOT / "3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx"


def digest(data):
    return hashlib.sha256(data).hexdigest()


def read_nested_schema(archive, member):
    with zipfile.ZipFile(archive) as outer:
        packed = outer.read(member)
    with zipfile.ZipFile(io.BytesIO(packed)) as inner:
        return {name: inner.read(name) for name in inner.namelist() if name.endswith(".xsd")}


def ms_svg_schema():
    # The Markdown source escapes XML punctuation.  Preserve every schema
    # namespace and declaration; only import *locations* are mapped to the
    # names in the vendored ECMA archive.
    text = "\n".join(
        re.sub(r"^\s*(?:\d+\. )?", "", line)
        for line in MS_SCHEMA.read_text().splitlines()
        if "\\<" in line
    )
    text = text.replace("\\<", "<").replace("\\>", ">").replace(
        "oartbasetypes.xsd", "dml-main.xsd"
    ).replace("orel.xsd", "shared-relationshipReference.xsd")
    assert f'targetNamespace="{SVG}"' in text
    assert f'namespace="{DML}"' in text
    assert f'namespace="{REL}"' in text
    assert STRICT_DML not in text and STRICT_REL not in text
    return text.encode()


class SchemaResolver(etree.Resolver):
    def __init__(self, schemas):
        super().__init__()
        self.schemas = schemas

    def resolve(self, url, public_id, context):
        key = url.rsplit("/", 1)[-1]
        if key not in self.schemas:
            raise ValueError(f"unrecognized offline schema import: {url}")
        return self.resolve_string(self.schemas[key], context, base_url=key)


def schema_context():
    strict = read_nested_schema(STRICT_ARCHIVE, "OfficeOpenXML-XMLSchema-Strict.zip")
    transitional = read_nested_schema(
        TRANSITIONAL_ARCHIVE, "OfficeOpenXML-XMLSchema-Transitional.zip"
    )
    strict_parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
    strict_parser.resolvers.add(SchemaResolver(strict))
    transitional_parser = etree.XMLParser(
        no_network=True, resolve_entities=False, load_dtd=False
    )
    transitional_parser.resolvers.add(SchemaResolver(transitional))
    strict_slide = etree.XMLSchema(etree.fromstring(strict["pml.xsd"], strict_parser))
    strict_xdr = etree.XMLSchema(
        etree.fromstring(strict["dml-spreadsheetDrawing.xsd"], strict_parser)
    )
    svg_schema = etree.XMLSchema(
        etree.fromstring(ms_svg_schema(), transitional_parser)
    )
    provenance = {
        "strict_ecma_archive_sha256": digest(STRICT_ARCHIVE.read_bytes()),
        "transitional_ecma_archive_sha256": digest(TRANSITIONAL_ARCHIVE.read_bytes()),
        "ms_odrawxml_5_24_schema_sha256": digest(MS_SCHEMA.read_bytes()),
        "ms_schema_namespace_policy": "unmodified target/a/r namespace values; relative import locations only mapped to vendored Transitional members",
        "strict_schema_members": ["pml.xsd", "dml-main.xsd", "dml-spreadsheetDrawing.xsd"],
        "ms_schema_import_members": ["dml-main.xsd", "shared-relationshipReference.xsd"],
        "lxml_version": etree.LXML_VERSION,
    }
    return {
        "strict_parser": strict_parser,
        "transitional_parser": transitional_parser,
        "strict_slide": strict_slide,
        "strict_xdr": strict_xdr,
        "svg_schema": svg_schema,
        "provenance": provenance,
    }


def strictify_native_xlsx(data):
    parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
    text = data.decode("utf-8")
    text = text.replace(XDR, STRICT_XDR).replace(DML, STRICT_DML).replace(REL, STRICT_REL)
    document = etree.fromstring(text.encode(), parser)
    # The unmodified MS child schema imports Transitional AG_Blob.  Restore its
    # relationship attributes after strictifying the native XDR/DML context.
    for svg_blip in document.iter(f"{{{SVG}}}svgBlip"):
        for local in ("embed", "link"):
            strict_name = f"{{{STRICT_REL}}}{local}"
            transitional_name = f"{{{REL}}}{local}"
            if strict_name in svg_blip.attrib:
                value = svg_blip.attrib.pop(strict_name)
                svg_blip.set(transitional_name, value)
    return etree.tostring(document, encoding="UTF-8", xml_declaration=True)


def ensure_native_xlsx_control(path):
    with zipfile.ZipFile(XLSX_FIXTURE) as archive:
        data = archive.read("xl/drawings/drawing1.xml")
    mixed = strictify_native_xlsx(data)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(mixed)
    return {
        "fixture": str(XLSX_FIXTURE.relative_to(ROOT)),
        "fixture_sha256": digest(XLSX_FIXTURE.read_bytes()),
        "fixture_member": "xl/drawings/drawing1.xml",
        "fixture_member_sha256": digest(data),
        "output_sha256": digest(mixed),
        "output_path": str(path.relative_to(ROOT)),
    }


def pictures(container, pml):
    for child in container:
        if child.tag == f"{{{pml}}}pic":
            yield child
        elif child.tag == f"{{{pml}}}grpSp":
            yield from pictures(child, pml)


def recognized_svg(blip, dml, svg_schema):
    owners = []
    ext_list = blip.find(f"{{{dml}}}extLst")
    if ext_list is None:
        return owners
    for ext in ext_list.findall(f"{{{dml}}}ext"):
        if ext.get("uri", "").strip(" \t\r\n") != SVG_URI:
            continue
        selected = ext.findall(f"{{{SVG}}}svgBlip")
        if len(selected) != 1:
            raise ValueError("recognized SVG extension must have one direct owner")
        svg_schema.assertValid(selected[0])
        owners.append({
            "attributes": {
                key: value for key, value in selected[0].attrib.items()
            },
            "relationship_attribute_namespace": {
                local: next(
                    (key.split("}", 1)[0][1:] for key in selected[0].attrib if key.endswith("}" + local)),
                    None,
                )
                for local in ("embed", "link")
            },
        })
    return owners


def validate_slide(path, context):
    data = path.read_bytes()
    if len(data) > MAX_XML_BYTES:
        raise ValueError("slide XML exceeds the validation profile limit")
    document = etree.fromstring(data, context["strict_parser"])
    if document.getroottree().docinfo.doctype or document.tag != f"{{{STRICT_PML}}}sld":
        raise ValueError("expected a DTD-free Strict PresentationML slide")
    context["strict_slide"].assertValid(document)
    tree = document.find(f"{{{STRICT_PML}}}cSld/{{{STRICT_PML}}}spTree")
    if tree is None:
        raise ValueError("slide lacks direct shape tree")
    owners = []
    for picture in pictures(tree, STRICT_PML):
        blip = picture.find(f"{{{STRICT_PML}}}blipFill/{{{STRICT_DML}}}blip")
        if blip is not None:
            owners.extend(recognized_svg(blip, STRICT_DML, context["svg_schema"]))
    relationship_path = path.with_name(f"{path.stem}.rels.xml")
    relationship_types = []
    if relationship_path.is_file():
        relationship_types = re.findall(
            r'Type="([^"]+)"', relationship_path.read_bytes().decode("utf-8")
        )
        if any(not value.startswith(f"{STRICT_REL}/") for value in relationship_types):
            raise AssertionError("Strict slide relationship member contains a non-Strict physical type")
    return {
        "kind": "strict-presentation-slide",
        "path": str(path.relative_to(ROOT)),
        "sha256": digest(data),
        "bytes": len(data),
        "core_strict_xsd_valid": True,
        "core_namespace_bindings": {
            prefix or "default": uri for prefix, uri in document.nsmap.items()
            if uri in (STRICT_PML, STRICT_DML, STRICT_REL)
        },
        "physical_relationship_types": relationship_types,
        "svg_owners": owners,
    }


def validate_control(path, expected_valid, context):
    data = path.read_bytes()
    document = etree.fromstring(data, context["transitional_parser"])
    valid = True
    error = None
    try:
        context["svg_schema"].assertValid(document)
    except etree.DocumentInvalid as exc:
        valid = False
        error = str(exc.error_log.last_error or exc)
    if valid != expected_valid:
        raise AssertionError(
            f"{path.name}: expected valid={expected_valid}, observed valid={valid}: {error}"
        )
    return {
        "kind": "direct-ms-odrawxml-svg-child",
        "path": str(path.relative_to(ROOT)),
        "sha256": digest(data),
        "bytes": len(data),
        "schema_valid": valid,
        "expected_valid": expected_valid,
        "error": error,
    }


def validate_xlsx(path, context):
    data = path.read_bytes()
    document = etree.fromstring(data, context["strict_parser"])
    if document.getroottree().docinfo.doctype or document.tag != f"{{{STRICT_XDR}}}wsDr":
        raise ValueError("expected a DTD-free Strict SpreadsheetDrawing root")
    context["strict_xdr"].assertValid(document)
    owners = []
    anchors = {local: 0 for local in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")}
    for anchor in document:
        local = etree.QName(anchor).localname
        if local in anchors:
            anchors[local] += 1
        picture = anchor.find(f"{{{STRICT_XDR}}}pic")
        if picture is None:
            continue
        blip = picture.find(f"{{{STRICT_XDR}}}blipFill/{{{STRICT_DML}}}blip")
        if blip is not None:
            owners.extend(recognized_svg(blip, STRICT_DML, context["svg_schema"]))
    return {
        "kind": "strict-native-xlsx-mixed-namespace-drawing",
        "path": str(path.relative_to(ROOT)),
        "sha256": digest(data),
        "bytes": len(data),
        "core_strict_xsd_valid": True,
        "anchor_counts": anchors,
        "svg_owners": owners,
    }


def main():
    context = schema_context()
    paths = [Path(name).resolve() for name in sys.argv[1:]]
    if not paths:
        raise ValueError("provide retained generated XML paths")
    native_path = HERE / "outputs/xlsx-mixed-strict.xml"
    native_provenance = ensure_native_xlsx_control(native_path)
    if native_path.resolve() not in paths:
        paths.append(native_path.resolve())
    reports = []
    for path in paths:
        if path.name == "direct-transitional-valid.xml":
            reports.append(validate_control(path, True, context))
        elif path.name == "direct-strict-invalid.xml":
            reports.append(validate_control(path, False, context))
        elif path.name == "xlsx-mixed-strict.xml":
            reports.append(validate_xlsx(path, context))
        else:
            reports.append(validate_slide(path, context))
    slide_reports = [row for row in reports if row["kind"] == "strict-presentation-slide"]
    controls = [row for row in reports if row["kind"] == "direct-ms-odrawxml-svg-child"]
    xlsx_reports = [
        row for row in reports if row["kind"] == "strict-native-xlsx-mixed-namespace-drawing"
    ]
    assert len(slide_reports) == 3
    assert [len(row["svg_owners"]) for row in slide_reports] == [0, 1, 0]
    assert [row["schema_valid"] for row in controls] == [True, False]
    assert len(xlsx_reports) == 1 and len(xlsx_reports[0]["svg_owners"]) == 2
    print(json.dumps({
        "passed": True,
        "scope": "Strict PresentationML and native Strict SpreadsheetDrawing core XSDs with direct unmodified MS-ODRAWXML 5.24 child validation; no native Office acceptance or rendering proof",
        **context["provenance"],
        "native_xlsx_control": native_provenance,
        "reports": reports,
    }, indent=2))


if __name__ == "__main__":
    main()
