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
