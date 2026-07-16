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

#[test]
fn round_trips_document_section_table_and_row_direction() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\rtldoc\rtlgutter\sectd\rtlsect "#,
        r#"\trowd\taprtl\rtlrow\cellx2000\cellx4000\pard\intbl A\cell B\cell\row "#,
        r#"\trowd\ltrrow\cellx2000\cellx4000\pard\intbl C\cell D\cell\row}"#,
    ))
    .unwrap();
    assert_eq!(document.document_direction(), Some(TextDirection::RightToLeft));
    assert!(document.gutter_on_right());
    assert_eq!(
        document.sections()[0].properties.direction,
        Some(TextDirection::RightToLeft)
    );
    let table = &document.tables()[0];
    assert_eq!(table.direction(), Some(TextDirection::RightToLeft));
    assert_eq!(table.rows()[0].direction(), Some(TextDirection::RightToLeft));
    assert_eq!(table.rows()[1].direction(), Some(TextDirection::LeftToRight));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    for control in ["rtldoc", "rtlgutter", "rtlsect", "taprtl", "rtlrow", "ltrrow"] {
        assert!(serialized.contains(&format!(r#"\{control}"#)));
    }
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert!(!reparsed.tables().is_empty(), "writer emitted no reparsable table: {serialized}");
    assert_eq!(reparsed.document_direction(), document.document_direction());
    assert_eq!(reparsed.gutter_on_right(), document.gutter_on_right());
    assert_eq!(
        reparsed.sections()[0].properties.direction,
        document.sections()[0].properties.direction
    );
    assert_eq!(reparsed.tables()[0].direction(), table.direction());
    assert_eq!(reparsed.tables()[0].rows()[0].direction(), table.rows()[0].direction());
    assert_eq!(reparsed.tables()[0].rows()[1].direction(), table.rows()[1].direction());
}

#[test]
fn explicit_ltr_and_reset_controls_are_typed() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ltrdoc\sectd\rtlsect\sectd\ltrsect "#,
        r#"\trowd\taprtl0\rtlrow\cellx1000\pard\intbl A\cell\row "#,
        r#"\trowd\cellx1000\pard\intbl B\cell\row}"#,
    ))
    .unwrap();
    assert_eq!(document.document_direction(), Some(TextDirection::LeftToRight));
    assert_eq!(
        document.sections()[0].properties.direction,
        Some(TextDirection::LeftToRight)
    );
    let table = &document.tables()[0];
    assert_eq!(table.direction(), Some(TextDirection::LeftToRight));
    assert_eq!(table.rows()[0].direction(), Some(TextDirection::RightToLeft));
    assert_eq!(table.rows()[1].direction(), None);
}

#[test]
fn parses_bundled_libreoffice_scope_direction_fixtures() {
    let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../../3rdparty/libreoffice-core/");

    let gutter = RtfDocument::parse_bytes(
        &fs::read(format!("{root}sw/qa/extras/rtfexport/data/rtl-gutter.rtf")).unwrap(),
    )
    .unwrap();
    assert!(gutter.gutter_on_right());

    let document_ltr = RtfDocument::parse_bytes(
        &fs::read(format!("{root}sw/qa/extras/rtfexport/data/dplinehollow.rtf")).unwrap(),
    )
    .unwrap();
    assert_eq!(document_ltr.document_direction(), Some(TextDirection::LeftToRight));

    let row_rtl = RtfDocument::parse_bytes(
        &fs::read(format!("{root}sw/qa/extras/rtfexport/data/table-rtl.rtf")).unwrap(),
    )
    .unwrap();
    assert_eq!(
        row_rtl.tables()[0].rows()[0].direction(),
        Some(TextDirection::RightToLeft)
    );

    let section_ltr = RtfDocument::parse_bytes(
        &fs::read(format!("{root}sw/qa/core/data/rtf/pass/tdf116851.rtf")).unwrap(),
    )
    .unwrap();
    assert!(section_ltr
        .sections()
        .iter()
        .any(|section| section.properties.direction == Some(TextDirection::LeftToRight)));
}
