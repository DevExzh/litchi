//! Image, drawing and rich-value regressions on unmodified native Excel files.
use litchi_xlsx::Package;
use litchi_xlsx::drawing::{DrawingAnchor, EditAs};
use litchi_xlsx::rich_values::package::{Document, Kind};
use sha2::{Digest, Sha256};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn digest(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = reader.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}
#[test]
fn native_picture_and_shape_share_a_drawing_without_losing_anchor_geometry() {
    let bytes = native("excel-conditional-drawings.xlsx");
    let package = Package::from_bytes(&bytes).unwrap();
    let workbook = package.workbook();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let drawing = sheet.drawing(0).unwrap();
    assert_eq!(drawing.typed().len(), 2);
    assert_eq!(drawing.picture_count(), 1);
    assert_eq!(drawing.content_part_count(), 0);
    let picture = drawing.picture(0).unwrap();
    assert_eq!(picture.source().c_nv_pr_id(), Some("3"));
    assert_eq!(picture.raster().part_uri().as_str(), "/xl/media/image1.png");
    assert_eq!(picture.raster().content_type(), "image/png");
    assert_eq!(
        picture.typed().description.as_deref(),
        Some("CF_PICTURE_005: synthetic blue and yellow stripes anchored over worksheet cells.")
    );
    let DrawingAnchor::TwoCell { from, to, edit_as } = picture.anchor() else {
        panic!("native two-cell anchor")
    };
    assert_eq!(*edit_as, EditAs::OneCell);
    assert_eq!(
        (
            from.column,
            from.row,
            from.column_offset.0,
            from.row_offset.0
        ),
        (0, 11, 0, 0)
    );
    assert_eq!(
        (to.column, to.row, to.column_offset.0, to.row_offset.0),
        (1, 16, 596900, 127000)
    );
    assert_eq!(
        digest(&members(&bytes)["xl/media/image1.png"]),
        "d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8"
    );
    let opc = Package::from_bytes(&bytes).unwrap().into_plain_opc();
    let shapes = litchi_xlsx::shapes::load_sheet_shapes(&opc, "Sheet1").unwrap();
    assert_eq!(shapes.objects.len(), 1);
    let litchi_xlsx::shapes::Object::Shape(shape) = &shapes.objects[0].object else {
        panic!("native rectangle")
    };
    assert_eq!(shape.non_visual.id, Some(4));
    assert_eq!(
        shape.non_visual.name.as_deref(),
        Some("Rounded Rectangle 3")
    );
    assert_eq!(shape.preset(), Some(litchi_xlsx::shapes::Preset::RoundRect));
    assert_eq!(shape.text_body.as_ref().unwrap().text(), "CF_SHAPE_005");
    let DrawingAnchor::TwoCell { from, to, edit_as } = &shapes.objects[0].anchor else {
        panic!("shape anchor")
    };
    assert_eq!(*edit_as, EditAs::TwoCell);
    assert_eq!(
        (
            from.column,
            from.row,
            from.column_offset.0,
            from.row_offset.0
        ),
        (5, 15, 1123950, 82550)
    );
    assert_eq!(
        (to.column, to.row, to.column_offset.0, to.row_offset.0),
        (7, 19, 1179830, 184150)
    );
    assert!(package.to_bytes().unwrap() == bytes);
}
#[test]
fn native_in_cell_image_rich_values_resolve_the_actual_lowercase_parts() {
    let bytes = native("excel-links-objects.xlsx");
    let opc = Package::from_bytes(&bytes).unwrap().into_plain_opc();
    let rich = litchi_xlsx::load_rich_values(&opc).unwrap();
    assert_eq!(rich.parts().len(), 4);
    let refresh =
        litchi_xlsx::load_rich_value_refresh(&opc, &litchi_xlsx::SourceLimits::default()).unwrap();
    assert_eq!(
        refresh.part_name().unwrap().as_str(),
        "/xl/richData/rdRichValueTypes.xml"
    );
    assert!(refresh.types().is_empty());
    assert_eq!(
        refresh.source_xml().unwrap(),
        members(&bytes)["xl/richData/rdRichValueTypes.xml"]
    );
    for part in rich.parts() {
        let xml = part.xml().unwrap();
        use litchi_xlsx::rich_values::codec;
        match part.document() {
            Document::Data(value) => assert_eq!(codec::parse_data(&xml).unwrap(), *value),
            Document::Structures(value) => {
                assert_eq!(codec::parse_structures(&xml).unwrap(), *value)
            },
            Document::Relationships(value) => {
                assert_eq!(codec::parse_rich_value_rels(&xml).unwrap(), *value)
            },
            Document::Opaque(value) => assert_eq!(xml, value.xml()),
            other => panic!("unexpected native part: {other:?}"),
        }
    }
    let Document::Data(data) = rich.part(Kind::Data).unwrap().document() else {
        panic!("data")
    };
    assert_eq!(data.values.len(), 1);
    assert_eq!(data.values[0].structure, 0);
    assert_eq!(data.values[0].values, ["0", "5"]);
    let Document::Structures(structures) = rich.part(Kind::Structures).unwrap().document() else {
        panic!("structures")
    };
    assert_eq!(structures.values.len(), 1);
    assert_eq!(structures.values[0].type_name, "_localImage");
    assert_eq!(
        structures.values[0]
            .keys
            .iter()
            .map(|k| k.name.as_str())
            .collect::<Vec<_>>(),
        ["_rvRel:LocalImageIdentifier", "CalcOrigin"]
    );
    assert!(
        structures.values[0]
            .keys
            .iter()
            .all(|k| k.value_type == litchi_xlsx::rich_values::ValueType::Integer)
    );
    assert_eq!(
        rich.part(Kind::Relationships)
            .unwrap()
            .relationships()
            .len(),
        1
    );
    let Document::Relationships(ids) = rich.part(Kind::Relationships).unwrap().document() else {
        panic!("rich relationships")
    };
    assert_eq!(ids.ids, ["rId1"]);
    let rel = &rich.part(Kind::Relationships).unwrap().relationships()[0];
    assert_eq!(rel.id, "rId1");
    assert_eq!(rel.mode, litchi_xlsx::rich_values::Mode::Internal);
    assert_eq!(
        rel.relationship_type,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    );
    assert_eq!(rel.resolved_target.as_deref(), Some("/xl/media/image1.png"));
    assert_eq!(
        digest(&members(&bytes)["xl/media/image1.png"]),
        "d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8"
    );
}
#[test]
fn native_svg_picture_keeps_its_vector_and_raster_fallback() {
    let bytes = native("excel-chart-families.xlsx");
    let package = Package::from_bytes(&bytes).unwrap();
    let book = package.workbook();
    let sheet = book.sheet("DrawingSVG_013").unwrap().unwrap();
    let picture = sheet
        .picture(litchi_xlsx::drawing::PictureSelector::new(0, 0))
        .unwrap();
    assert_eq!(
        picture.typed().description.as_deref(),
        Some("EXCEL_SVG_ALT_013: blue and yellow vector blocks with the SVG_MARKER_007 label.")
    );
    let DrawingAnchor::TwoCell { from, to, edit_as } = picture.anchor() else {
        panic!("native SVG anchor")
    };
    assert_eq!(*edit_as, EditAs::OneCell);
    assert_eq!(
        (
            from.column,
            from.row,
            from.column_offset.0,
            from.row_offset.0
        ),
        (0, 0, 0, 0)
    );
    assert_eq!(
        (to.column, to.row, to.column_offset.0, to.row_offset.0),
        (3, 8, 571500, 88900)
    );
    let svg = picture.read_svg_image().unwrap();
    assert_eq!(svg.descriptor().part_uri().as_str(), "/xl/media/image2.svg");
    assert_eq!(
        digest(svg.bytes()),
        "91bdc6ffe9ddcd4ddda5c1e4872f0003b6b0864a0523c19945a4917da6858a95"
    );
    assert_eq!(picture.raster().part_uri().as_str(), "/xl/media/image1.png");
    assert_eq!(
        digest(&members(&bytes)["xl/media/image1.png"]),
        "97fefe9283fc468417bf2b339b48197711947bd05cc835e40a31088b4e1e7990"
    );
    let parts = members(&bytes);
    let source = litchi_xlsx::drawing::SourceDrawing::scan(
        &parts["xl/drawings/drawing8.xml"],
        0,
        litchi_xlsx::drawing::ScanLimits::default(),
    )
    .unwrap();
    assert_eq!(source.pictures().len(), 1);
    assert_eq!(source.pictures()[0].anchor(), picture.anchor());
    assert!(package.to_bytes().unwrap() == bytes);
}

#[test]
fn native_svg_replacement_preserves_the_raster_geometry_and_all_chart_parts() {
    let bytes = native("excel-chart-families.xlsx");
    let before = members(&bytes);
    let replacement = String::from_utf8(before["xl/media/image2.svg"].clone())
        .unwrap()
        .replace("SVG_MARKER_007", "SVG_EDIT_019")
        .into_bytes();
    assert_ne!(replacement, before["xl/media/image2.svg"]);
    let package = Package::from_bytes(&bytes).unwrap();
    let selector = litchi_xlsx::drawing::PictureSelector::new(0, 0);
    let original_book = package.workbook();
    let original_picture = original_book.picture("DrawingSVG_013", selector).unwrap();
    let mut edit = package.edit().unwrap();
    edit.sheet("DrawingSVG_013")
        .unwrap()
        .unwrap()
        .detach_svg(selector)
        .unwrap()
        .attach_svg(selector, replacement.as_slice().into())
        .unwrap();
    let commit = edit.commit().unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("svg-edit.xlsx");
    commit.package().save(&path).unwrap();
    let saved = std::fs::read(path).unwrap();
    let reopened = Package::from_bytes(&saved).unwrap();
    let book = reopened.workbook();
    let picture = book.picture("DrawingSVG_013", selector).unwrap();
    assert_eq!(picture.read_svg_image().unwrap().bytes(), replacement);
    assert_eq!(picture.anchor(), original_picture.anchor());
    assert_eq!(
        picture.typed().description,
        original_picture.typed().description
    );
    assert_eq!(
        picture.raster().part_uri(),
        original_picture.raster().part_uri()
    );
    let after = members(&saved);
    let changed = [
        "[Content_Types].xml",
        "xl/drawings/drawing8.xml",
        "xl/drawings/_rels/drawing8.xml.rels",
        "xl/media/image2.svg",
    ];
    for (name, data) in &before {
        if !changed.contains(&name.as_str()) {
            assert!(after.get(name) == Some(data), "changed unrelated {name}");
        }
    }
    let new_svg = picture
        .svg()
        .unwrap()
        .unwrap()
        .part_uri()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    for name in after.keys() {
        assert!(
            before.contains_key(name) || name == &new_svg,
            "unexpected new member {name}"
        );
    }
    assert_eq!(after["xl/media/image1.png"], before["xl/media/image1.png"]);
    assert!(
        commit
            .patch()
            .inverse()
            .durable()
            .unwrap()
            .apply(&reopened)
            .unwrap()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(package.to_bytes().unwrap() == bytes);
}
