use litchi_rtf::{
    PageBorder, PageBorderAppliesTo, PageBorderDepth, PageBorderOffset, PageBorderStyle,
    RtfDocument, RtfWriter,
};

#[test]
fn parses_libreoffice_page_border_export_and_round_trips_deterministically() {
    let fixture = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/page-border.rtf"
    ));
    let document = RtfDocument::parse(fixture).unwrap();
    let borders = document.sections()[0].properties.page_borders;
    assert_eq!(borders.top.unwrap().width, 10);
    assert_eq!(borders.left.unwrap().width, 20);
    assert_eq!(borders.bottom.unwrap().width, 30);
    assert_eq!(borders.right.unwrap().width, 40);
    assert_eq!(borders.top.unwrap().space, 480);
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let text = String::from_utf8(output).unwrap();
    let expected = r#"\pgbrdrt\brdrs\brdrw10\brdrcf0\brsp480\pgbrdrl\brdrs\brdrw20\brdrcf0\brsp480\pgbrdrb\brdrs\brdrw30\brdrcf0\brsp480\pgbrdrr\brdrs\brdrw40\brdrcf0\brsp480"#;
    assert!(text.contains(expected), "{text}");
    let reparsed = RtfDocument::parse(&text).unwrap();
    assert_eq!(reparsed.sections()[0].properties.page_borders, borders);
}

#[test]
fn parses_options_group_boundaries_destinations_and_sectd_reset() {
    let rtf = r#"{\rtf1{\*\unknown\pgbrdrt\brdrs\brdrw75}\sectd\pgbrdropt43\pgbrdrhead\pgbrdrfoot\pgbrdrsnap{\pgbrdrt\brdrdb\brdrw25\brdrcf7\brsp30\brdrsh}\sectd\pgbrdrb\brdrdot\brdrw5 Body}"#;
    let document = RtfDocument::parse(rtf).unwrap();
    let borders = document.sections()[0].properties.page_borders;
    assert!(borders.top.is_none());
    assert_eq!(borders.bottom.unwrap().style, PageBorderStyle::Dotted);
    assert_eq!(borders.applies_to, PageBorderAppliesTo::AllSectionPages);
    assert_eq!(borders.depth, PageBorderDepth::InFront);
    assert_eq!(borders.offset, PageBorderOffset::Text);
    assert!(!borders.surround_header && !borders.surround_footer && !borders.snap_to_text_borders);
}

#[test]
fn writer_emits_typed_options_and_art_in_canonical_order() {
    let mut section = litchi_rtf::Section::new();
    section.properties.page_borders.applies_to = PageBorderAppliesTo::WholeDocument;
    section.properties.page_borders.depth = PageBorderDepth::Behind;
    section.properties.page_borders.offset = PageBorderOffset::Page;
    section.properties.page_borders.surround_header = true;
    section.properties.page_borders.top = Some(PageBorder {
        art: Some(42),
        width: 12,
        color_ref: 3,
        space: 20,
        ..PageBorder::default()
    });
    let mut output = Vec::new();
    RtfWriter::new(&mut output).write_section(&section).unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(
        text.contains(r#"\pgbrdropt43\pgbrdrhead\pgbrdrt\brdrart42\brdrw12\brdrcf3\brsp20"#),
        "{text}"
    );
}

#[test]
fn rejects_malformed_page_borders_and_negative_libreoffice_fixture() {
    for rtf in [
        r#"{\rtf1\pgbrdrt Body}"#,
        r#"{\rtf1\pgbrdrt\brdrw10\brdrs Body}"#,
        r#"{\rtf1\pgbrdrt\brdrs\brdrdb Body}"#,
        r#"{\rtf1\pgbrdrt\brdrs\brdrw Body}"#,
        r#"{\rtf1\pgbrdrt\brdrs\brdrw76 Body}"#,
        r#"{\rtf1\pgbrdrt\brdrs\brsp1441 Body}"#,
        r#"{\rtf1\pgbrdrt\brdrart0 Body}"#,
        r#"{\rtf1\pgbrdrt\brdrart166 Body}"#,
        r#"{\rtf1\pgbrdropt Body}"#,
        r#"{\rtf1\pgbrdropt4 Body}"#,
        r#"{\rtf1\pgbrdropt16 Body}"#,
        r#"{\rtf1\pgbrdropt64 Body}"#,
        r#"{\rtf1\pgbrdrt\brdrs\brdrw10\pgbrdrt\brdrs Body}"#,
    ] {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted {rtf}");
    }
    let negative = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/writerfilter/rtftok/data/negative-page-border.rtf"
    ));
    assert!(RtfDocument::parse(negative).is_err());
}
