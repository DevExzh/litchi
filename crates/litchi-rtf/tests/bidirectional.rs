use litchi_rtf::{RtfDocument, RtfWriter, StyleBlock, TextDirection};
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

fn block<'rtf>(
    document: &'rtf RtfDocument<'rtf>,
    text: &str,
) -> StyleBlock<'rtf> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.trim() == text)
        .cloned()
        .unwrap_or_else(|| panic!("missing block {text:?}"))
}

#[test]
fn preserves_group_scoped_run_direction_and_boundaries() {
    let document = RtfDocument::parse(r#"{\rtf1\ltrch Left {\rtlch Right} LeftAgain}"#)
        .unwrap();
    assert_eq!(
        block(&document, "Left").formatting.direction,
        Some(TextDirection::LeftToRight)
    );
    assert_eq!(
        block(&document, "Right").formatting.direction,
        Some(TextDirection::RightToLeft)
    );
    assert_eq!(
        block(&document, "LeftAgain").formatting.direction,
        Some(TextDirection::LeftToRight)
    );

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        block(&reparsed, "Right").formatting.direction,
        Some(TextDirection::RightToLeft)
    );
}

#[test]
fn plain_and_pard_reset_character_and_paragraph_direction() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\rtlpar\rtlch RTL\par "#,
        r#"\pard\plain LTR\par}"#,
    ))
    .unwrap();
    let rtl = block(&document, "RTL");
    assert_eq!(rtl.formatting.direction, Some(TextDirection::RightToLeft));
    assert_eq!(rtl.paragraph.direction, Some(TextDirection::RightToLeft));

    let ltr = block(&document, "LTR");
    assert_eq!(ltr.formatting.direction, None);
    assert_eq!(ltr.paragraph.direction, None);
}

#[test]
fn writer_preserves_explicit_ltr_and_rtl_directions() {
    let document = RtfDocument::parse(
        r#"{\rtf1\rtlpar{\rtlch rtl}{\ltrch ltr}\par}"#,
    )
    .unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"\rtlpar"#));
    assert!(serialized.contains(r#"\rtlch"#));
    assert!(serialized.contains(r#"\ltrch"#));

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        block(&reparsed, "rtl").formatting.direction,
        Some(TextDirection::RightToLeft)
    );
    assert_eq!(
        block(&reparsed, "ltr").formatting.direction,
        Some(TextDirection::LeftToRight)
    );
    assert_eq!(
        block(&reparsed, "rtl").paragraph.direction,
        Some(TextDirection::RightToLeft)
    );
}

#[test]
fn parses_bundled_libreoffice_bidirectional_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/extras/rtfexport/data/tdf86182.rtf",
        "sw/qa/extras/rtfexport/data/tdf126309.rtf",
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
    ];
    let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../../3rdparty/libreoffice-core/");
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(
            document.blocks().iter().any(|block| {
                block.formatting.direction.is_some() || block.paragraph.direction.is_some()
            }),
            "fixture exposed no bidirectional run or paragraph metadata: {fixture}"
        );
    }
}
