#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentView, DocumentViewKind, DocumentZoomKind, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_view_and_zoom_metadata_and_round_trips_in_stable_order() {
    let document =
        RtfDocument::parse(r"{\rtf1\viewkind4\viewscale135\viewzk3\viewbksp1\viewnobound Body}")
            .unwrap();
    assert_eq!(
        *document.document_view(),
        DocumentView {
            kind: Some(DocumentViewKind::Normal),
            scale_percent: Some(135),
            zoom_kind: Some(DocumentZoomKind::TextWidth),
            background_shapes: Some(true),
            hide_page_boundaries: true,
        }
    );
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\viewkind4").unwrap() < serialized.find("\\viewscale135").unwrap());
    assert!(serialized.find("\\viewscale135").unwrap() < serialized.find("\\viewzk3").unwrap());
    assert!(serialized.find("\\viewzk3").unwrap() < serialized.find("\\viewbksp1").unwrap());
    assert!(serialized.find("\\viewbksp1").unwrap() < serialized.find("\\viewnobound").unwrap());
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.document_view(), document.document_view());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn typed_api_preserves_absence_and_clear() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    assert!(document.document_view().is_empty());
    document
        .set_document_view(DocumentView {
            kind: Some(DocumentViewKind::PageLayout),
            scale_percent: Some(250),
            zoom_kind: None,
            background_shapes: Some(false),
            hide_page_boundaries: false,
        })
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.document_view(), document.document_view());
    document.clear_document_view();
    assert!(document.document_view().is_empty());
}

#[test]
fn rejects_invalid_values_duplicates_and_bad_placement_or_group_shape() {
    for source in [
        r"{\rtf1\viewkind Body}",
        r"{\rtf1\viewkind6 Body}",
        r"{\rtf1\viewscale Body}",
        r"{\rtf1\viewscale0 Body}",
        r"{\rtf1\viewscale10001 Body}",
        r"{\rtf1\viewzk Body}",
        r"{\rtf1\viewzk4 Body}",
        r"{\rtf1\viewbksp Body}",
        r"{\rtf1\viewbksp2 Body}",
        r"{\rtf1\viewnobound1 Body}",
        r"{\rtf1\viewbksp1\viewbksp0 Body}",
        r"{\rtf1\viewnobound\viewnobound Body}",
        r"{\rtf1\viewkind1\viewkind4 Body}",
        r"{\rtf1\viewkind1{\*\viewkind4}Body}",
        r"{\rtf1{\viewkind1 nested}Body}",
        r"{\rtf1 Body\viewscale100}",
        r"{\rtf1{\*\viewscale100 extra}Body}",
        r"{\rtf1{\*\viewkind1\b}Body}",
        r"{\rtf1{{\*\viewkind1}}Body}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn parses_bundled_libreoffice_starred_view_fixture() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf114303.rtf"
    ))
    .unwrap();
    let document = RtfDocument::parse_bytes(&bytes).unwrap();
    assert_eq!(
        document.document_view().kind,
        Some(DocumentViewKind::PageLayout)
    );
    assert_eq!(document.document_view().scale_percent, Some(100));
    assert_eq!(document.document_view().background_shapes, None);
    assert!(!document.document_view().hide_page_boundaries);
}

#[test]
fn parses_bundled_libreoffice_background_shapes_fixture() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/page-background.rtf"
    ))
    .unwrap();
    let document = RtfDocument::parse_bytes(&bytes).unwrap();
    assert_eq!(document.document_view().background_shapes, Some(true));
}
