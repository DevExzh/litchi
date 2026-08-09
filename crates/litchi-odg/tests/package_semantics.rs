#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::{Drawing, shape::ShapeKind};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><office:body><office:drawing><draw:layer-set><draw:layer draw:name="Foreground"/></draw:layer-set><draw:page draw:name="Page 1"><draw:rect draw:name="Label" draw:layer="Foreground"><text:p>Old label</text:p></draw:rect><draw:frame draw:name="Photo" svg:width="2cm" svg:height="1cm"/></draw:page></office:drawing></office:body></office:document-content>"#;

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn semantic_package_views_use_layers_shapes_and_shared_frame_context() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    assert_eq!(drawing.pages().len(), 1);
    assert_eq!(drawing.pages()[0].name(), Some("Page 1"));
    assert_eq!(drawing.layers()[0].name(), "Foreground");
    let shape = &drawing.pages()[0].shapes()[0];
    assert_eq!(shape.name(), Some("Label"));
    assert_eq!(shape.layer(), Some("Foreground"));
    assert_eq!(shape.kind(), ShapeKind::Rectangle);
    assert_eq!(shape.text(), "Old label");
    assert_eq!(
        drawing.pages()[0].shapes()[1]
            .frame()
            .unwrap()
            .width
            .as_deref(),
        Some("2cm")
    );
}

#[test]
fn package_edit_is_source_checked_reversible_and_compact() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_text(0, 0, "New <label>").unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].text(),
        "New <label>"
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("<text:p>New &lt;label&gt;</text:p>")
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    assert!(!commit.snapshot().content_xml().contains(">\r\n<"));
    assert!(commit.patch().is_applicable_to(drawing.snapshot()));
    let reapplied = commit.patch().apply(drawing.snapshot()).unwrap();
    assert_eq!(reapplied.as_bytes(), commit.snapshot().as_bytes());
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.as_bytes(), drawing.as_bytes());
}

#[test]
fn package_shape_name_edit_is_source_checked_reversible_and_compact() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_name(0, 0, "Renamed & exact").unwrap();
    let commit = transaction.commit().unwrap();

    let shape = &commit.snapshot().pages()[0].shapes()[0];
    assert_eq!(shape.name(), Some("Renamed & exact"));
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:name=\"Renamed &amp; exact\"")
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:layer=\"Foreground\"")
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:name=\"Photo\"")
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    assert!(!commit.snapshot().content_xml().contains(">\r\n<"));

    let change = commit.patch().name_change().unwrap();
    assert_eq!(change.before(), "Label");
    assert_eq!(change.after(), "Renamed & exact");
    assert!(commit.patch().is_applicable_to(drawing.snapshot()));
    let different =
        Drawing::from_bytes(package(&CONTENT.replace("Photo", "Different photo"))).unwrap();
    assert!(!commit.patch().is_applicable_to(different.snapshot()));
    assert!(commit.patch().apply(different.snapshot()).is_err());
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.as_bytes(), drawing.as_bytes());
}

#[test]
fn dtd_and_noncompact_rewrite_are_refused_without_execution() {
    let dtd = CONTENT.replacen("<office:body>", "<!DOCTYPE drawing><office:body>", 1);
    assert!(Drawing::from_bytes(package(&dtd)).is_err());

    let noncompact = CONTENT.replacen("<office:body>", "\n<office:body>", 1);
    let drawing = Drawing::from_bytes(package(&noncompact)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_text(0, 0, "New label").unwrap();
    assert!(transaction.commit().is_err());
}
