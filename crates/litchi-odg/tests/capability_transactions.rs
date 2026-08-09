#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::{
    Drawing,
    layer::Layer,
    page::Page,
    shape::{Shape, ShapeKind},
};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:drawing><draw:page draw:name="One"><draw:layer-set><draw:layer draw:name="base"/></draw:layer-set><draw:rect draw:name="Box" draw:layer="base" draw:style-name="gr1" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><text:p>old</text:p></draw:rect><draw:frame draw:name="Image"><draw:image xlink:href="Pictures/pixel.bin" xlink:type="simple"/></draw:frame></draw:page></office:drawing></office:body></office:document-content>"#;

fn package() -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer
        .add_file_with_media_type("Pictures/pixel.bin", b"before", "image/png")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn one_transaction_composes_structure_geometry_style_layers_and_resources() {
    let source = Drawing::from_bytes(package()).unwrap();
    assert_eq!(source.resources().len(), 1);
    assert_eq!(
        source.resource_bytes(0).unwrap().as_deref(),
        Some(&b"before"[..])
    );

    let mut edit = source.edit();
    edit.set_shape_geometry(0, 0, "2cm", "3cm", "5cm", "6cm")
        .unwrap();
    edit.set_shape_style_name(0, 0, "gr2").unwrap();
    edit.set_shape_text(0, 0, "new & compact").unwrap();
    edit.add_layer(0, Layer::new("notes").with_display("always"))
        .unwrap();
    edit.add_shape(
        0,
        Shape::new(ShapeKind::Ellipse)
            .with_name("Circle")
            .with_layer("notes")
            .with_geometry("1cm", "1cm", "2cm", "2cm"),
    )
    .unwrap();
    edit.add_group(0, "Group").unwrap();
    edit.add_page(Page::new("Two")).unwrap();
    edit.add_shape(
        1,
        Shape::new(ShapeKind::Rectangle)
            .with_name("Second")
            .with_geometry("0cm", "0cm", "1cm", "1cm"),
    )
    .unwrap();
    edit.set_resource(0, "image/png", b"after".to_vec())
        .unwrap();

    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().pages().len(), 2);
    assert_eq!(commit.snapshot().pages()[0].layers().len(), 2);
    assert_eq!(commit.snapshot().pages()[0].shapes().len(), 4);
    assert_eq!(commit.snapshot().pages()[0].shapes()[0].x(), Some("2cm"));
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].style_name(),
        Some("gr2")
    );
    assert_eq!(
        commit.snapshot().pages()[1].shapes()[0].name(),
        Some("Second")
    );
    assert_eq!(
        commit.snapshot().resource_bytes(0).unwrap().as_deref(),
        Some(&b"after"[..])
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    assert!(commit.patch().changes().len() >= 8);
    assert_eq!(commit.patch().resource_changes().len(), 1);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn removals_are_checked_and_group_subtrees_are_owned() {
    let group_content = CONTENT.replace(
        "<draw:frame draw:name=\"Image\">",
        "<draw:g draw:name=\"Group\"><draw:rect draw:name=\"Child\"/></draw:g><draw:frame draw:name=\"Image\">",
    );
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer
        .add_file("content.xml", group_content.as_bytes())
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/pixel.bin", b"before", "image/png")
        .unwrap();
    let source = Drawing::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

    let mut blocked = source.edit();
    assert!(blocked.remove_layer(0, "base").is_err());

    let mut edit = source.edit();
    let removed = edit.remove_shape(0, 1).unwrap();
    assert_eq!(removed.kind(), ShapeKind::Group);
    let commit = edit.commit().unwrap();
    assert!(
        commit.snapshot().pages()[0]
            .shapes()
            .iter()
            .all(|shape| shape.name() != Some("Child") && shape.name() != Some("Group"))
    );
}

#[test]
fn cross_drawing_group_transfer_carries_layers_resources_and_durable_inverse() {
    const SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Source"><draw:layer-set><draw:layer draw:name="media"/></draw:layer-set><draw:g draw:name="Transferred" draw:layer="media"><draw:frame draw:name="Picture" draw:layer="media"><draw:image xlink:href="Pictures/transfer.bin" xlink:type="simple"/></draw:frame></draw:g></draw:page></office:drawing></office:body></office:document-content>"#;
    const DESTINATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Destination"/></office:drawing></office:body></office:document-content>"#;

    let source = drawing_with_resource(SOURCE, Some(b"transfer bytes"));
    let destination = drawing_with_resource(DESTINATION, None);
    let transfer = source.snapshot().prepare_shape_transfer(0, 0).unwrap();
    assert_eq!(transfer.shape().kind(), ShapeKind::Group);
    assert_eq!(transfer.layers()[0].name(), "media");
    assert_eq!(transfer.resources()[0].path(), "Pictures/transfer.bin");

    let mut edit = destination.edit();
    edit.insert_shape_transfer(0, 0, &transfer).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().pages()[0].layers()[0].name(), "media");
    assert_eq!(commit.snapshot().pages()[0].shapes().len(), 2);
    assert_eq!(
        commit.snapshot().resource_bytes(0).unwrap().as_deref(),
        Some(&b"transfer bytes"[..])
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(destination.snapshot()).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        destination.as_bytes()
    );

    let limits = litchi_core::CompositionLimits::new(4, 4, 8, 8);
    let mut competing_edit = destination.edit();
    competing_edit
        .add_page(Page::new("Competing page"))
        .unwrap();
    let competing = competing_edit.commit().unwrap();
    let mut joined = destination.joined_edits(limits);
    joined
        .join(commit.patch().prepare("transfer", limits).unwrap())
        .unwrap();
    assert!(
        joined
            .join(competing.patch().prepare("structure", limits).unwrap())
            .is_err()
    );

    let collision = drawing_with_resource(DESTINATION, Some(b"different"));
    let mut remapped = collision.edit();
    remapped.insert_shape_transfer(0, 0, &transfer).unwrap();
    let remapped = remapped.commit().unwrap();
    assert!(remapped.snapshot().content_xml().contains("_litchi_"));
    assert_eq!(
        remapped.snapshot().resource_bytes(0).unwrap().as_deref(),
        Some(&b"transfer bytes"[..])
    );
}

#[test]
fn cross_drawing_transfer_copies_and_remaps_arbitrary_style_definitions() {
    const SOURCE: &str = r##"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Source"><draw:rect draw:name="Styled" draw:style-name="gr1"/></draw:page></office:drawing></office:body></office:document-content>"##;
    const DESTINATION: &str = r##"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#0000ff"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Destination"/></office:drawing></office:body></office:document-content>"##;
    let source = drawing_with_resource(SOURCE, None);
    let destination = drawing_with_resource(DESTINATION, None);
    let transfer = source.snapshot().prepare_shape_transfer(0, 0).unwrap();
    assert_eq!(transfer.style_definitions()[0].name(), "gr1");

    let mut edit = destination.edit();
    edit.insert_shape_transfer(0, 0, &transfer).unwrap();
    let commit = edit.commit().unwrap();
    let shape_style = commit.snapshot().pages()[0].shapes()[0]
        .style_name()
        .unwrap();
    assert!(shape_style.starts_with("gr1_litchi_"));
    assert!(commit.snapshot().style_definitions().iter().any(|style| {
        style.name() == shape_style && style.property("draw:fill-color") == Some("#ff0000")
    }));
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        destination.as_bytes()
    );
}

#[test]
fn unreferenced_resource_add_remove_is_safe_and_exactly_reversible() {
    let source = Drawing::from_bytes(package()).unwrap();
    let mut addition = source.edit();
    addition
        .add_resource("Media/sound.bin", "audio/basic", b"sound".to_vec())
        .unwrap();
    assert!(
        addition
            .add_resource("../escape.bin", "audio/basic", Vec::new())
            .is_err()
    );
    let added = addition.commit().unwrap();
    assert!(
        added
            .snapshot()
            .files()
            .unwrap()
            .contains(&"Media/sound.bin".to_string())
    );
    let restored = added
        .patch()
        .durable()
        .unwrap()
        .inverse()
        .apply(added.snapshot())
        .unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());

    let mut removal = added.snapshot().edit();
    removal
        .remove_unreferenced_resource("Media/sound.bin")
        .unwrap();
    let removed = removal.commit().unwrap();
    assert!(
        !removed
            .snapshot()
            .files()
            .unwrap()
            .contains(&"Media/sound.bin".to_string())
    );
}

fn drawing_with_resource(content: &str, resource: Option<&[u8]>) -> Drawing {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(bytes) = resource {
        writer
            .add_file_with_media_type("Pictures/transfer.bin", bytes, "application/octet-stream")
            .unwrap();
    }
    Drawing::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
}
