#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{
    DocumentReadOnlyRecommendation, DocumentSavePreferences, DocumentThumbnailPreference,
    RtfDocument, RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_both_parameterless_passive_flags() {
    let document = RtfDocument::parse(r"{\rtf1\readonlyrecommended\saveprevpict Body}").unwrap();
    assert_eq!(
        *document.save_preferences(),
        DocumentSavePreferences {
            read_only: DocumentReadOnlyRecommendation::Recommended,
            thumbnail: DocumentThumbnailPreference::RequiredIfSupported,
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_preserves_unspecified_semantics() {
    let document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    assert!(document.save_preferences().is_empty());
    assert_eq!(
        document.save_preferences().thumbnail,
        DocumentThumbnailPreference::Unspecified
    );
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    document.set_save_preferences(DocumentSavePreferences {
        read_only: DocumentReadOnlyRecommendation::Recommended,
        thumbnail: DocumentThumbnailPreference::RequiredIfSupported,
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(
        serialized.find("\\readonlyrecommended").unwrap()
            < serialized.find("\\saveprevpict").unwrap()
    );
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.save_preferences(), document.save_preferences());
    assert_eq!(reparsed.text(), "Body");
    document.clear_save_preferences();
    assert!(document.save_preferences().is_empty());
}

#[test]
fn coexists_after_existing_passive_root_metadata_in_deterministic_order() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\xform transform.xsl}\usexform"#,
        r#"{\*\wgrffmtfilter 2002}\stylesortmethod4"#,
        r#"\saveprevpict\readonlyrecommended Body}"#,
    ))
    .unwrap();
    let serialized = String::from_utf8(write(&document)).unwrap();
    for (first, second) in [
        ("\\xform", "\\usexform"),
        ("\\usexform", "\\wgrffmtfilter"),
        ("\\wgrffmtfilter", "\\stylesortmethod4"),
        ("\\stylesortmethod4", "\\readonlyrecommended"),
        ("\\readonlyrecommended", "\\saveprevpict"),
    ] {
        assert!(serialized.find(first).unwrap() < serialized.find(second).unwrap());
    }
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_flags() {
    for source in [
        r"{\rtf1\readonlyrecommended0 Body}",
        r"{\rtf1\readonlyrecommended1 Body}",
        r"{\rtf1\saveprevpict0 Body}",
        r"{\rtf1\saveprevpict-1 Body}",
        r"{\rtf1\readonlyrecommended\readonlyrecommended Body}",
        r"{\rtf1\saveprevpict\saveprevpict Body}",
        r"{\rtf1{\*\readonlyrecommended}Body}",
        r"{\rtf1{\*\saveprevpict}Body}",
        r"{\rtf1{\readonlyrecommended}Body}",
        r"{\rtf1{\saveprevpict}Body}",
        r"{\rtf1 Body\readonlyrecommended}",
        r"{\rtf1 Body\saveprevpict}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
