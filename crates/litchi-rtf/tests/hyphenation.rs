#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentHyphenation, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_all_document_hyphenation_controls_and_round_trips() {
    let document =
        RtfDocument::parse(r"{\rtf1\ansi\hyphhotz425\hyphconsec3\hyphcaps0\hyphauto Body}")
            .unwrap();
    assert_eq!(
        *document.hyphenation(),
        DocumentHyphenation {
            automatic: Some(true),
            capitalized_words: Some(false),
            consecutive_line_limit: Some(3),
            hot_zone_twips: Some(425),
        }
    );
    assert_eq!(document.text(), "Body");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.hyphenation(), document.hyphenation());
    assert_eq!(reparsed.text(), document.text());
}

#[test]
fn typed_api_preserves_absence_and_explicit_defaults() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    assert!(document.hyphenation().is_empty());
    document
        .set_hyphenation(DocumentHyphenation {
            automatic: Some(false),
            capitalized_words: Some(true),
            consecutive_line_limit: Some(0),
            hot_zone_twips: Some(0),
        })
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.hyphenation(), document.hyphenation());

    document.clear_hyphenation();
    assert!(document.hyphenation().is_empty());
}

#[test]
fn numeric_controls_round_trip_the_complete_rtf_nonnegative_range() {
    for value in [0, 1, 425, i32::MAX as u32] {
        let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
        document
            .set_hyphenation(DocumentHyphenation {
                consecutive_line_limit: Some(value),
                hot_zone_twips: Some(value),
                ..DocumentHyphenation::default()
            })
            .unwrap();
        let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
        assert_eq!(reparsed.hyphenation(), document.hyphenation());
        assert_eq!(reparsed.text(), "Body");
    }
}

#[test]
fn rejects_bad_values_duplicates_and_non_root_or_late_placement() {
    for source in [
        r"{\rtf1\hyphauto2 Body}",
        r"{\rtf1\hyphcaps-1 Body}",
        r"{\rtf1\hyphconsec Body}",
        r"{\rtf1\hyphconsec-1 Body}",
        r"{\rtf1\hyphconsec2147483648 Body}",
        r"{\rtf1\hyphhotz Body}",
        r"{\rtf1\hyphhotz-1 Body}",
        r"{\rtf1\hyphhotz2147483648 Body}",
        r"{\rtf1\hyphauto1\hyphauto0 Body}",
        r"{\rtf1{\hyphauto1 nested}Body}",
        r"{\rtf1 Body\hyphauto1}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }

    for control in ["hyphauto1", "hyphcaps0", "hyphconsec3", "hyphhotz425"] {
        for source in [
            format!(r"{{\rtf1\{control}\{control} Body}}"),
            format!(r"{{\rtf1{{\*\{control}}}Body}}"),
            format!(r"{{\rtf1{{\{control}}}Body}}"),
            format!(r"{{\rtf1 Body\{control}}}"),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }

    let invalid = DocumentHyphenation {
        hot_zone_twips: Some(i32::MAX as u32 + 1),
        ..DocumentHyphenation::default()
    };
    assert!(invalid.validate().is_err());
    let invalid = DocumentHyphenation {
        consecutive_line_limit: Some(i32::MAX as u32 + 1),
        ..DocumentHyphenation::default()
    };
    assert!(invalid.validate().is_err());
}

#[test]
fn parses_bundled_document_hyphenation_fixture() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf163003.rtf"
    ))
    .unwrap();
    let document = RtfDocument::parse_bytes(&bytes).unwrap();
    assert_eq!(document.hyphenation().hot_zone_twips, Some(142));
    assert_eq!(document.hyphenation().capitalized_words, Some(false));
    assert_eq!(document.hyphenation().automatic, Some(true));
}
