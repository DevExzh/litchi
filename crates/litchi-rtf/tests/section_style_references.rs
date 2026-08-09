#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{RtfDocument, RtfWriter, SectionProperties, StyleType};

#[test]
fn retains_grouped_references_inherits_sections_and_sectd_resets() {
    let inherited = RtfDocument::parse(concat!(
        r#"{\rtf1{\stylesheet{\*\ds5 Declared Section;}}"#,
        r#"\sectd\ds5{\ds6}First\sect\lndscpsxn Second}"#,
    ))
    .unwrap();

    assert!(
        inherited
            .stylesheet()
            .get_typed(StyleType::Section, 5)
            .is_some()
    );
    assert_eq!(inherited.sections().len(), 2);
    assert_eq!(inherited.sections()[0].properties.section_style, Some(6));
    assert_eq!(inherited.sections()[1].properties.section_style, Some(6));

    let reset = RtfDocument::parse(r"{\rtf1\sectd\ds6 First\sect\sectd Second}").unwrap();
    assert_eq!(reset.sections().len(), 2);
    assert_eq!(reset.sections()[0].properties.section_style, Some(6));
    assert_eq!(reset.sections()[1].properties.section_style, None);
}

#[test]
fn preserves_zero_maximum_omission_and_public_mutation() {
    let document =
        RtfDocument::parse(r"{\rtf1\sectd\ds0 First\sect\sectd\ds65535 Second\sect\sectd Third}")
            .unwrap();
    assert_eq!(document.sections()[0].properties.section_style, Some(0));
    assert_eq!(
        document.sections()[1].properties.section_style,
        Some(65_535)
    );
    assert_eq!(document.sections()[2].properties.section_style, None);

    let mut properties = SectionProperties::default();
    properties.set_section_style(Some(0));
    assert_eq!(properties.section_style, Some(0));
    properties.set_section_style(None);
    assert_eq!(properties, SectionProperties::default());
}

#[test]
fn rejects_malformed_body_and_stylesheet_handles_and_duplicate_selectors() {
    for source in [
        r"{\rtf1\sectd\ds Body}",
        r"{\rtf1\sectd\ds-1 Body}",
        r"{\rtf1\sectd\ds65536 Body}",
        r"{\rtf1{\stylesheet{\*\ds Missing;}}}",
        r"{\rtf1{\stylesheet{\*\ds-1 Negative;}}}",
        r"{\rtf1{\stylesheet{\*\ds65536 Overflow;}}}",
        r"{\rtf1{\stylesheet{\ds1 Unstarred;}}}",
        r"{\rtf1{\stylesheet{\*\b\ds1 Late;}}}",
        r"{\rtf1{\stylesheet{\*\ds1\ds2 Duplicate;}}}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let inert =
        RtfDocument::parse(r"{\rtf1{\field{\*\fldinst TEST \ds65536}{\fldrslt Result}}Body}")
            .unwrap();
    assert!(
        inert
            .sections()
            .iter()
            .all(|section| section.properties.section_style.is_none())
    );
}

#[test]
fn writer_is_canonical_and_round_trip_is_stable_without_resolving_styles() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\stylesheet{\*\ds3 Declared Section;}}"#,
        r#"\sectd\ds3\lndscpsxn Body}"#,
    ))
    .unwrap();
    let original = document.sections()[0].properties.clone();

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let serialized = String::from_utf8(first.clone()).unwrap();
    assert!(serialized.contains(r"\sectd\ds3"));

    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.stylesheet(), document.stylesheet());
    assert_eq!(reparsed.sections()[0].properties, original);
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn parses_libreoffice_section_fixture_without_materializing_a_reference() {
    let source = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf100507.rtf"
    ));
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert!(
        document
            .sections()
            .iter()
            .all(|section| section.properties.section_style.is_none())
    );
}
