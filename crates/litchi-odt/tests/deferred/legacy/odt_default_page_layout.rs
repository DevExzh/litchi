//! Tests for `style:default-page-layout` against real documents.

use litchi_odt::FlatOpenDocument;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/odf/odt")
        .join(name)
}

const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

#[test]
fn reads_default_page_layout_from_a_flat_document() {
    let document = FlatOpenDocument::open(fixture("note-tracked-changes.fodt")).unwrap();
    let layout = document
        .default_page_layout()
        .unwrap()
        .expect("fixture declares a default page layout");
    assert!(layout.name.is_empty(), "default layout is unnamed");
    let properties = layout.properties.as_ref().expect("layout properties");
    assert_eq!(
        properties.attribute(Some(STYLE_NS), "writing-mode"),
        Some("lr-tb")
    );
    assert!(layout.xml.starts_with("<style:default-page-layout>"));
}

#[test]
fn documents_without_default_page_layout_report_none() {
    let document = FlatOpenDocument::open(fixture("drawing-style-resources.fodt")).unwrap();
    assert!(document.default_page_layout().unwrap().is_none());
}
