#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{Alignment, Paragraph, RtfDocument, RtfWriter, StyleType};

fn paragraph_for<'a>(document: &'a RtfDocument<'a>, text: &str) -> Paragraph {
    document
        .blocks()
        .iter()
        .find(|block| block.text == text)
        .unwrap_or_else(|| panic!("missing block {text:?}"))
        .paragraph
}

#[test]
fn retains_scoped_references_without_materializing_style_properties() {
    let source = concat!(
        r#"{\rtf1{\stylesheet"#,
        r#"{\s5\qr\sb120 Heading;}"#,
        r#"{\s6\qc Independent Style;}"#,
        r#"}{\s5\ql A{\s6\qc B}C\par}{\pard D}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();

    let independent = document
        .stylesheet()
        .get_typed(StyleType::Paragraph, 6)
        .unwrap();
    let properties = independent.paragraph.unwrap();
    assert_eq!(properties.paragraph_style, None);
    assert_eq!(properties.alignment, Alignment::Center);

    let a = paragraph_for(&document, "A");
    assert_eq!(a.paragraph_style, Some(5));
    assert_eq!(a.alignment, Alignment::Left);
    let b = paragraph_for(&document, "B");
    assert_eq!(b.paragraph_style, Some(6));
    assert_eq!(b.alignment, Alignment::Center);
    assert_eq!(paragraph_for(&document, "C"), a);
    assert_eq!(paragraph_for(&document, "D").paragraph_style, None);

    let declared = document
        .stylesheet()
        .get_typed(StyleType::Paragraph, 5)
        .unwrap();
    assert_eq!(declared.paragraph.unwrap().alignment, Alignment::Right);
    assert_eq!(declared.paragraph.unwrap().paragraph_style, None);
}

#[test]
fn defaults_mutation_and_writer_preserve_zero_maximum_and_omission() {
    let document =
        RtfDocument::parse(r"{\rtf1{\*\defpap\s0\qc}{\s65535\qr Maximum\par}{\pard Reset}}")
            .unwrap();
    assert_eq!(
        document
            .default_formatting()
            .paragraph()
            .unwrap()
            .paragraph
            .paragraph_style,
        Some(0)
    );
    assert_eq!(
        paragraph_for(&document, "Maximum").paragraph_style,
        Some(65_535)
    );
    assert_eq!(paragraph_for(&document, "Reset").paragraph_style, None);

    let mut paragraph = Paragraph::default();
    paragraph.set_paragraph_style(Some(0));
    assert_eq!(paragraph.paragraph_style, Some(0));
    paragraph.set_paragraph_style(None);
    assert_eq!(paragraph, Paragraph::default());

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let serialized = String::from_utf8(first.clone()).unwrap();
    assert!(serialized.contains(r"{\*\defpap\s0"));
    assert!(serialized.contains(r"\s65535\qr"));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(
        reparsed.default_formatting().paragraph(),
        document.default_formatting().paragraph()
    );
    assert_eq!(
        paragraph_for(&reparsed, "Maximum"),
        paragraph_for(&document, "Maximum")
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_missing_out_of_range_duplicate_and_misplaced_references() {
    for source in [
        r"{\rtf1\s X}",
        r"{\rtf1\s-1 X}",
        r"{\rtf1\s65536 X}",
        r"{\rtf1{\stylesheet{\s Missing;}}}",
        r"{\rtf1{\stylesheet{\s-1 Negative;}}}",
        r"{\rtf1{\stylesheet{\s65536 Overflow;}}}",
        r"{\rtf1{\stylesheet{\b\s1 Late;}}}",
        r"{\rtf1{\*\defpap\s1\s2}X}",
        r"{\rtf1{\*\defpap\s}X}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let inert =
        RtfDocument::parse(r"{\rtf1{\field{\*\fldinst TEST \s65536}{\fldrslt Result}}Body}")
            .unwrap();
    assert_eq!(inert.fields().len(), 1);
    assert!(
        inert
            .blocks()
            .iter()
            .all(|block| block.paragraph.paragraph_style.is_none())
    );
}

#[test]
fn stable_round_trip_covers_body_style_and_default_owners() {
    let source = concat!(
        r#"{\rtf1{\*\defpap\s3\qc}{\stylesheet"#,
        r#"{\s3\qr Declared;}"#,
        r#"{\s4\ql Independent;}"#,
        r#"}{\s3\qc Styled\par}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.stylesheet(), document.stylesheet());
    assert_eq!(
        reparsed.default_formatting().paragraph(),
        document.default_formatting().paragraph()
    );
    assert_eq!(
        paragraph_for(&reparsed, "Styled"),
        paragraph_for(&document, "Styled")
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn parses_libreoffice_paragraph_style_producer() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/fdo82071.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert!(
        document
            .stylesheet()
            .get_typed(StyleType::Paragraph, 19)
            .is_some()
    );
}
