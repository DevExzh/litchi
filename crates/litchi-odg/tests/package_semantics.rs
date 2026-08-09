#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::{
    compact_xml,
    core::{OwnedPackage, PackageWriter},
};
use litchi_odg::{Drawing, shape::ShapeKind};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.3"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1" draw:master-page-name="Default" xml:id="page1"><draw:layer-set><draw:layer draw:name="Foreground" draw:display="always" draw:protected="false"/><draw:layer draw:name="Background"/></draw:layer-set><draw:rect draw:name="Label" draw:layer="Foreground" draw:style-name="gr1" draw:text-style-name="P1" draw:z-index="7" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><svg:title>Label title</svg:title><svg:desc>Label description</svg:desc><text:p>Old label</text:p></draw:rect><draw:frame draw:name="Photo" draw:layer="Background" svg:width="2cm" svg:height="1cm"><svg:title>Photo title</svg:title><svg:desc>Photo description</svg:desc></draw:frame></draw:page></office:drawing></office:body></office:document-content>"#;

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
    assert_eq!(
        drawing.page(0usize).unwrap().unwrap().name(),
        Some("Page 1")
    );
    assert_eq!(
        drawing.page("Page 1").unwrap().unwrap().name(),
        Some("Page 1")
    );
    assert!(drawing.page(9usize).unwrap().is_none());
    assert_eq!(drawing.pages().len(), 1);
    assert_eq!(drawing.pages()[0].name(), Some("Page 1"));
    assert_eq!(drawing.pages()[0].xml_id(), Some("page1"));
    assert_eq!(drawing.pages()[0].style_name(), Some("dp1"));
    assert_eq!(drawing.pages()[0].master_page_name(), Some("Default"));
    assert!(drawing.layers().is_empty());
    assert_eq!(drawing.pages()[0].layers().len(), 2);
    assert_eq!(drawing.pages()[0].layers()[0].name(), "Foreground");
    assert_eq!(drawing.pages()[0].layers()[0].display(), Some("always"));
    assert_eq!(drawing.pages()[0].layers()[0].protected(), Some(false));
    let shape = &drawing.pages()[0].shapes()[0];
    assert_eq!(drawing.pages()[0].shape("Label").unwrap(), Some(shape));
    assert_eq!(drawing.pages()[0].shape(0usize).unwrap(), Some(shape));
    assert!(drawing.pages()[0].shape("Missing").unwrap().is_none());
    assert_eq!(shape.name(), Some("Label"));
    assert_eq!(shape.layer(), Some("Foreground"));
    assert_eq!(shape.kind(), ShapeKind::Rectangle);
    assert_eq!(shape.style_name(), Some("gr1"));
    assert_eq!(shape.text_style_name(), Some("P1"));
    assert_eq!(shape.z_index(), Some(7));
    assert_eq!(shape.x(), Some("1cm"));
    assert_eq!(shape.y(), Some("2cm"));
    assert_eq!(shape.width(), Some("3cm"));
    assert_eq!(shape.height(), Some("4cm"));
    assert_eq!(shape.title(), Some("Label title"));
    assert_eq!(shape.description(), Some("Label description"));
    assert_eq!(shape.text(), "Old label");
    assert_eq!(
        drawing.pages()[0].shapes()[1]
            .frame()
            .unwrap()
            .width
            .as_deref(),
        Some("2cm")
    );
    assert_eq!(
        drawing.pages()[0].shapes()[1]
            .frame()
            .unwrap()
            .title
            .as_deref(),
        Some("Photo title")
    );
}

#[test]
fn exact_shape_name_selector_rejects_ambiguity() {
    let duplicate = CONTENT.replace("draw:name=\"Photo\"", "draw:name=\"Label\"");
    let drawing = Drawing::from_bytes(package(&duplicate)).unwrap();
    assert!(drawing.pages()[0].shape("Label").is_err());
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
    let output = OwnedPackage::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    let archive = output.package().unwrap();
    for path in ["content.xml", "META-INF/manifest.xml"] {
        compact_xml::validate(&archive.get_file(path).unwrap()).unwrap();
    }
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
fn package_layer_edit_is_declared_source_checked_and_reversible() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_layer(0, 0, "Background").unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].layer(),
        Some("Background")
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:layer=\"Background\"")
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    let change = commit.patch().layer_change().unwrap();
    assert_eq!(change.before(), "Foreground");
    assert_eq!(change.after(), "Background");
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.as_bytes(), drawing.as_bytes());

    let mut invalid = drawing.edit();
    assert!(invalid.set_shape_layer(0, 0, "Missing").is_err());
}

#[test]
fn ended_page_scope_does_not_capture_shapes_outside_a_page() {
    let malformed = CONTENT.replace(
        "</draw:page></office:drawing>",
        "</draw:page><draw:g><draw:rect draw:name=\"outside\"/></draw:g></office:drawing>",
    );
    assert!(Drawing::from_bytes(package(&malformed)).is_err());
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
