#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{
    MAX_PAGE_NUMBER_HEADING_LEVEL, PageNumberHeadingSeparator, RtfDocument, RtfWriter, Section,
    SectionPageNumberHeading,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_heading_level_and_separator_and_round_trips_in_canonical_order() {
    let source = r"{\rtf1\sectd\pgnhn2{\pgnhnsh}Body}";
    let document = RtfDocument::parse(source).unwrap();
    let heading = document.sections()[0].properties.page_number_heading;
    assert_eq!(
        heading,
        SectionPageNumberHeading {
            level: Some(2),
            separator: Some(PageNumberHeadingSeparator::Hyphen),
        }
    );
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r"\pgnhn2\pgnhnsh"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.sections()[0].properties.page_number_heading,
        heading
    );
}

#[test]
fn parses_all_separator_kinds() {
    for (control, expected) in [
        ("pgnhnsh", PageNumberHeadingSeparator::Hyphen),
        ("pgnhnsp", PageNumberHeadingSeparator::Period),
        ("pgnhnsc", PageNumberHeadingSeparator::Colon),
        ("pgnhnsm", PageNumberHeadingSeparator::EmDash),
        ("pgnhnsn", PageNumberHeadingSeparator::EnDash),
    ] {
        let source = format!(r"{{\rtf1\pgnhn1\{control} X}}");
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(
            document.sections()[0]
                .properties
                .page_number_heading
                .separator,
            Some(expected),
            "wrong separator for {control}"
        );
        let output = write(&document);
        let serialized = String::from_utf8(output).unwrap();
        assert!(
            serialized.contains(&format!(r"\{control}")),
            "missing {control} in {serialized}"
        );
    }
}

#[test]
fn preserves_explicit_disable_missing_parameter_and_sectd_reset() {
    let disabled = RtfDocument::parse(r"{\rtf1\pgnhn0 X}").unwrap();
    assert_eq!(
        disabled.sections()[0].properties.page_number_heading.level,
        Some(0)
    );

    let bare = RtfDocument::parse(r"{\rtf1\pgnhn X}").unwrap();
    assert_eq!(
        bare.sections()[0].properties.page_number_heading.level,
        Some(0)
    );

    let separator_only = RtfDocument::parse(r"{\rtf1\pgnhnsm X}").unwrap();
    assert_eq!(
        separator_only.sections()[0].properties.page_number_heading,
        SectionPageNumberHeading {
            level: None,
            separator: Some(PageNumberHeadingSeparator::EmDash),
        }
    );

    let reset = RtfDocument::parse(r"{\rtf1\pgnhn9\pgnhnsc X\sectd Y}").unwrap();
    assert!(
        reset.sections()[0]
            .properties
            .page_number_heading
            .is_empty()
    );
}

#[test]
fn inert_destinations_do_not_change_page_number_heading() {
    for source in [
        r"{\rtf1{\*\unknown\pgnhn4\pgnhnsh}Body}",
        r"{\rtf1{\header\pgnhn3\pgnhnsp Header}Body}",
        r"{\rtf1{\footer\pgnhn3 Footer}Body}",
        r"{\rtf1{\field{\*\fldinst IF \pgnhn2 1 1}{\fldrslt visible}}Body}",
        r"{\rtf1 A{\footnote\pgnhn1\pgnhnsn Note}Body}",
        r"{\rtf1{\object\pgnhn5}Body}",
    ] {
        let document = RtfDocument::parse(source).unwrap();
        assert!(
            document
                .sections()
                .iter()
                .all(|section| section.properties.page_number_heading.is_empty()),
            "destination leaked page-number heading: {source}"
        );
    }
}

#[test]
fn rejects_out_of_range_levels_and_invalid_public_writer_state() {
    for source in [
        r"{\rtf1\pgnhn-1 X}".to_string(),
        format!(r"{{\rtf1\pgnhn{} X}}", MAX_PAGE_NUMBER_HEADING_LEVEL + 1),
    ] {
        assert!(RtfDocument::parse(&source).is_err(), "accepted {source}");
    }

    let mut section = Section::new();
    section.properties.page_number_heading.level = Some(u8::MAX);
    let mut output = Vec::new();
    assert!(RtfWriter::new(&mut output).write_section(&section).is_err());
}
