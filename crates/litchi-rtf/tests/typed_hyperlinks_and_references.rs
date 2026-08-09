#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{Field, FieldOwner, FieldStatus, ReferenceFieldKind, RtfDocument, RtfWriter};

const FIXTURE: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/rtf/typed_hyperlinks_and_references.rtf"
));

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn inspects_hyperlink_and_cross_reference_metadata_without_resolving_targets() {
    let document = RtfDocument::parse(FIXTURE).unwrap();

    assert_eq!(document.text(), "BeforeAfter");
    assert_eq!(document.hyperlink_count(), 1);
    assert_eq!(document.reference_field_count(), 3);

    let hyperlinks = document.hyperlinks();
    let hyperlink = &hyperlinks[0];
    assert_eq!(
        hyperlink.external_target(),
        Some("https://example.invalid/document")
    );
    assert_eq!(hyperlink.bookmark(), Some("chapter-1"));
    assert_eq!(hyperlink.screen_tip(), Some("Read preview"));
    assert_eq!(hyperlink.target_frame(), Some("_blank"));
    assert_eq!(hyperlink.coordinates(), Some("12,34"));
    assert!(hyperlink.opens_in_new_window());
    assert_eq!(hyperlink.unknown_switches().len(), 2);
    assert_eq!(hyperlink.cached_result(), Some("Link text"));
    assert_eq!(
        hyperlink.status(),
        FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        }
    );
    assert_eq!(hyperlink.owner(), FieldOwner::Body);
    assert_eq!(hyperlink.position(), "Before".len());

    let references = document.reference_fields();
    assert_eq!(references[0].kind(), ReferenceFieldKind::Reference);
    assert_eq!(references[0].bookmark(), "bookmark-alpha");
    assert!(references[0].has_hyperlink());
    assert!(references[0].includes_relative_position());
    assert!(references[0].includes_footnote_mark());
    assert_eq!(references[0].unknown_switches().len(), 1);
    assert_eq!(references[0].cached_result(), Some("Reference text"));
    assert_eq!(
        references[0].status(),
        FieldStatus {
            edited: true,
            ..FieldStatus::default()
        }
    );

    assert_eq!(references[1].kind(), ReferenceFieldKind::PageReference);
    assert_eq!(references[1].bookmark(), "bookmark-beta");
    assert!(references[1].has_hyperlink());
    assert!(!references[1].includes_relative_position());
    assert!(!references[1].includes_footnote_mark());
    assert_eq!(references[1].unknown_switches().len(), 1);
    assert_eq!(references[1].cached_result(), Some("Page text"));

    assert_eq!(references[2].kind(), ReferenceFieldKind::NoteReference);
    assert_eq!(references[2].bookmark(), "note-1");
    assert!(!references[2].has_hyperlink());
    assert!(!references[2].includes_relative_position());
    assert!(references[2].includes_footnote_mark());
    assert_eq!(references[2].cached_result(), Some("Note text"));
    assert_eq!(
        references[2].status(),
        FieldStatus {
            private: true,
            ..FieldStatus::default()
        }
    );

    assert!(Field::parse_instruction("HYPERLINK").hyperlink().is_none());
    assert!(Field::parse_instruction("REF").reference_field().is_none());
}

#[test]
fn typed_hyperlink_and_cross_reference_metadata_round_trips() {
    let document = RtfDocument::parse(FIXTURE).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();

    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.hyperlink_count(), document.hyperlink_count());
    assert_eq!(
        reparsed.reference_field_count(),
        document.reference_field_count()
    );

    let hyperlink = reparsed.hyperlinks().remove(0);
    assert_eq!(
        hyperlink.external_target(),
        Some("https://example.invalid/document")
    );
    assert_eq!(hyperlink.bookmark(), Some("chapter-1"));
    assert_eq!(hyperlink.unknown_switches().len(), 2);

    let references = reparsed.reference_fields();
    assert_eq!(references[0].kind(), ReferenceFieldKind::Reference);
    assert_eq!(references[1].kind(), ReferenceFieldKind::PageReference);
    assert_eq!(references[2].kind(), ReferenceFieldKind::NoteReference);
}
