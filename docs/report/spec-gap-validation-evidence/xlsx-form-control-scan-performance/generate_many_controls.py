#!/usr/bin/env python3
"""Create a deterministic accepted many-control XLSX from the native button fixture."""

from __future__ import annotations

import pathlib
import re
import sys
import zipfile


FIXED_DATE = (1980, 1, 1, 0, 0, 0)
CONTROL_PROPERTIES_CONTENT_TYPE = "application/vnd.ms-excel.controlproperties+xml"


def replacement(text: str, old: str, new: str) -> str:
    if old not in text:
        raise RuntimeError(f"missing template text {old!r}")
    return text.replace(old, new)


def clone_info(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    clone = zipfile.ZipInfo(info.filename, FIXED_DATE)
    clone.compress_type = zipfile.ZIP_DEFLATED
    clone.create_system = 0
    clone.external_attr = 0
    return clone


def first_control_fragment(sheet: str) -> tuple[str, str]:
    controls_start = sheet.index("<controls>") + len("<controls>")
    controls_end = sheet.index("</controls>", controls_start)
    controls = sheet[controls_start:controls_end]
    alt_start = controls.index("<mc:AlternateContent")
    alt_end = controls.index("</mc:AlternateContent>", alt_start) + len(
        "</mc:AlternateContent>"
    )
    alt = controls[alt_start:alt_end]
    control_start = alt.index("<control shapeId=")
    control_end = alt.index("</control>", control_start) + len("</control>")
    return alt, alt[control_start:control_end]


def build_sheet(source: str, count: int) -> str:
    alt, control = first_control_fragment(source)
    controls_start = source.index("<controls>") + len("<controls>")
    controls_end = source.index("</controls>", controls_start)
    additions = []
    for position in range(1, count):
        shape_id = 1025 + position
        relationship_id = 3 + position
        name = f"Button {position + 1}"
        item = replacement(alt, control, replacement(
            replacement(
                replacement(
                    control,
                    'shapeId="1025"',
                    f'shapeId="{shape_id}"',
                ),
                'r:id="rId3"',
                f'r:id="rId{relationship_id}"',
            ),
            'name="Button 1"',
            f'name="{name}"',
        ))
        additions.append(item)
    return source[:controls_end] + "".join(additions) + source[controls_end:]


def build_drawing(source: str, count: int) -> str:
    first_start = source.index("<mc:AlternateContent")
    first_end = source.index("</mc:AlternateContent>", first_start) + len(
        "</mc:AlternateContent>"
    )
    template = source[first_start:first_end]
    additions = []
    for position in range(1, count):
        shape_id = 1025 + position
        item = replacement(
            replacement(
                replacement(template, "_x0000_s1025", f"_x0000_s{shape_id}"),
                'id="1025"',
                f'id="{shape_id}"',
            ),
            "Button 1",
            f"Button {position + 1}",
        )
        additions.append(item)
    return source.replace("</xdr:wsDr>", "".join(additions) + "</xdr:wsDr>", 1)


def build_vml(source: str, count: int) -> str:
    first_start = source.index('<v:shape id="_x0000_s1025"')
    first_end = source.index("</v:shape>", first_start) + len("</v:shape>")
    template = source[first_start:first_end]
    additions = []
    for position in range(1, count):
        shape_id = 1025 + position
        item = replacement(template, "_x0000_s1025", f"_x0000_s{shape_id}")
        additions.append(item)
    return source.replace("</xml>", "".join(additions) + "</xml>", 1)


def main() -> None:
    if len(sys.argv) != 4:
        raise SystemExit("usage: generate_many_controls.py BASE OUTPUT COUNT")
    base = pathlib.Path(sys.argv[1])
    output = pathlib.Path(sys.argv[2])
    count = int(sys.argv[3])
    if count < 3 or count > 64:
        raise SystemExit("count must be between 3 and 64")

    with zipfile.ZipFile(base, "r") as source_zip:
        members = {info.filename: source_zip.read(info.filename) for info in source_zip.infolist()}

    sheet_name = "xl/worksheets/sheet1.xml"
    drawing_name = "xl/drawings/drawing1.xml"
    vml_name = "xl/drawings/vmlDrawing1.vml"
    rels_name = "xl/worksheets/_rels/sheet1.xml.rels"
    content_types_name = "[Content_Types].xml"
    members[sheet_name] = build_sheet(members[sheet_name].decode("utf-8"), count).encode()
    members[drawing_name] = build_drawing(members[drawing_name].decode("utf-8"), count).encode()
    members[vml_name] = build_vml(members[vml_name].decode("utf-8"), count).encode()

    rels = members[rels_name].decode("utf-8")
    rel_additions = "".join(
        f'<Relationship Id="rId{3 + position}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" '
        f'Target="../ctrlProps/ctrlProp{position + 1}.xml"/>'
        for position in range(1, count)
    )
    members[rels_name] = rels.replace("</Relationships>", rel_additions + "</Relationships>", 1).encode()

    content_types = members[content_types_name].decode("utf-8")
    overrides = "".join(
        f'<Override PartName="/xl/ctrlProps/ctrlProp{position + 1}.xml" '
        f'ContentType="{CONTROL_PROPERTIES_CONTENT_TYPE}"/>'
        for position in range(1, count)
    )
    members[content_types_name] = content_types.replace("</Types>", overrides + "</Types>", 1).encode()
    properties = members["xl/ctrlProps/ctrlProp1.xml"]
    for position in range(1, count):
        members[f"xl/ctrlProps/ctrlProp{position + 1}.xml"] = properties

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w") as destination:
        for name, payload in members.items():
            destination.writestr(clone_info(zipfile.ZipInfo(name)), payload)


if __name__ == "__main__":
    main()
