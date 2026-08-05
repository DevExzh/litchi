use super::*;
use crate::section::Start;
use std::mem::size_of;

#[test]
fn typed_section_round_trips_and_preserves_unknown_children() {
    let xml = r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test"><w:type w:val="continuous"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="100" w:right="200" w:bottom="300" w:left="400" w:header="500" w:footer="600" w:gutter="50"/><w:cols w:num="2" w:space="720"/><x:ext x:v="keep"/></w:sectPr>"#;
    let mut section = SectionProperties::from_xml(xml).unwrap();
    assert_eq!(section.start_type, Some(Start::Continuous));
    section.margin_left = 900;
    let mut output = String::new();
    section.write_xml(&mut output, None).unwrap();
    assert!(output.contains("w:left=\"900\""));
    assert!(output.contains("<x:ext x:v=\"keep\"/>"));
}

#[test]
fn rejects_duplicate_and_invalid_section_properties() {
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:type w:val=\"nextPage\"/><w:type w:val=\"continuous\"/></w:sectPr>"
        )
        .is_err()
    );
    let section = SectionProperties {
        columns: Some(SectionColumns {
            count: 0,
            ..SectionColumns::default()
        }),
        ..SectionProperties::default()
    };
    assert!(section.validate().is_err());
}

#[test]
fn page_layout_properties_round_trip() {
    let xml = r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:paperSrc w:first="1" w:other="260"/><w:pgBorders w:offsetFrom="text" w:zOrder="front" w:display="firstPage"><w:top w:val="double" w:sz="8" w:space="24" w:color="ff0000" w:shadow="1"/><w:bottom w:val="starsTop" w:sz="120" w:space="4" w:color="auto" w:frame="true"/></w:pgBorders><w:lnNumType w:countBy="5" w:start="0" w:distance="240" w:restart="newSection"/><w:pgNumType w:fmt="lowerRoman" w:start="3"/><w:cols w:num="2"/><w:vAlign w:val="both"/><w:titlePg/><w:bidi/><w:rtlGutter/><w:printerSettings r:id="rId9"/></w:sectPr>"#;
    let section = SectionProperties::from_xml(xml).unwrap();
    assert_eq!(
        section.paper_source,
        Some(SectionPaperSource {
            first: Some(1),
            other: Some(260),
        })
    );
    let borders = section.page_borders.as_ref().unwrap();
    assert_eq!(borders.offset_from, OffsetFrom::Text);
    assert_eq!(borders.z_order, ZOrder::Front);
    assert_eq!(borders.display, Display::FirstPage);
    let top = borders.top.as_ref().unwrap();
    assert_eq!(top.style, Style::Double);
    assert_eq!(top.size, Some(8));
    assert_eq!(top.space, Some(24));
    assert_eq!(top.color, Some(Color::rgb(255, 0, 0)));
    assert!(top.shadow);
    assert!(!top.frame);
    let bottom = borders.bottom.as_ref().unwrap();
    assert_eq!(bottom.style, Style::Art(Art::StarsTop));
    assert_eq!(bottom.size, Some(120));
    assert_eq!(bottom.color, Some(Color::Auto));
    assert!(bottom.frame);
    assert!(borders.left.is_none() && borders.right.is_none());
    assert_eq!(
        section.line_numbering,
        Some(SectionLineNumbering {
            count_by: Some(5),
            start: Some(0),
            distance: Some(240),
            restart: Some(LineNumberRestart::NewSection),
        })
    );
    assert_eq!(
        section.vertical_alignment,
        Some(SectionVerticalAlignment::Justified)
    );
    assert!(section.title_page);
    assert!(section.bidirectional);
    assert!(section.rtl_gutter);
    assert_eq!(
        section.printer_settings_relationship_id.as_deref(),
        Some("rId9")
    );

    let mut output = String::new();
    section.write_xml(&mut output, None).unwrap();
    assert!(output.contains("w:color=\"FF0000\""));
    let reparsed = SectionProperties::from_xml(&output).unwrap();
    assert_eq!(reparsed.paper_source, section.paper_source);
    assert_eq!(reparsed.page_borders, section.page_borders);
    assert_eq!(reparsed.line_numbering, section.line_numbering);
    assert_eq!(reparsed.vertical_alignment, section.vertical_alignment);
    assert!(reparsed.title_page && reparsed.bidirectional && reparsed.rtl_gutter);
    assert_eq!(
        reparsed.printer_settings_relationship_id,
        section.printer_settings_relationship_id
    );
}

#[test]
fn page_layout_defaults_and_empty_edges() {
    let xml = r#"<w:sectPr><w:pgBorders/><w:lnNumType w:countBy="2"/></w:sectPr>"#;
    let section = SectionProperties::from_xml(xml).unwrap();
    let borders = section.page_borders.as_ref().unwrap();
    assert_eq!(borders.offset_from, OffsetFrom::Page);
    assert_eq!(borders.z_order, ZOrder::Back);
    assert_eq!(borders.display, Display::AllPages);
    assert_eq!(
        section.line_numbering,
        Some(SectionLineNumbering {
            count_by: Some(2),
            ..SectionLineNumbering::default()
        })
    );
    let mut output = String::new();
    section.write_xml(&mut output, None).unwrap();
    assert!(
        output.contains(
            "<w:pgBorders w:offsetFrom=\"page\" w:zOrder=\"back\" w:display=\"allPages\"/>"
        )
    );
    assert!(output.contains("<w:lnNumType w:countBy=\"2\"/>"));
}

#[test]
fn typed_chapter_and_note_domains_round_trip() {
    let xml = r#"<w:sectPr><w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="2"/><w:numRestart w:val="eachPage"/><w:pos w:val="beneathText"/></w:footnotePr><w:endnotePr><w:numFmt w:val="upperLetter"/><w:pos w:val="docEnd"/></w:endnotePr><w:pgNumType w:fmt="decimal" w:chapStyle="1" w:chapSep="emDash"/></w:sectPr>"#;
    let section = SectionProperties::from_xml(xml).unwrap();
    assert_eq!(
        section.footnotes,
        Some(Footnotes {
            format: PageNumberFormat::LowerRoman,
            start: Some(2),
            restart: Some(NoteNumberRestart::EachPage),
            position: Some(FootnotePos::BeneathText),
        })
    );
    assert_eq!(
        section.endnotes,
        Some(Endnotes {
            format: PageNumberFormat::UpperLetter,
            position: Some(EndnotePos::DocumentEnd),
            ..Endnotes::default()
        })
    );
    assert_eq!(
        section
            .page_numbering
            .as_ref()
            .and_then(|numbering| numbering.chapter_separator),
        Some(ChapterSep::EmDash)
    );

    let mut output = String::new();
    section.write_xml(&mut output, None).unwrap();
    let reparsed = SectionProperties::from_xml(&output).unwrap();
    assert_eq!(reparsed.footnotes, section.footnotes);
    assert_eq!(reparsed.endnotes, section.endnotes);
    assert_eq!(reparsed.page_numbering, section.page_numbering);
}

#[test]
fn rejects_unknown_chapter_and_note_domain_values() {
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:pgNumType w:fmt=\"decimal\" w:chapSep=\"slash\"/></w:sectPr>"
        )
        .is_err()
    );
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:footnotePr><w:pos w:val=\"middle\"/></w:footnotePr></w:sectPr>"
        )
        .is_err()
    );
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:endnotePr><w:pos w:val=\"pageBottom\"/></w:endnotePr></w:sectPr>"
        )
        .is_err()
    );
}

#[test]
fn page_border_style_enum_round_trips() {
    let styles = [
        Style::Nil,
        Style::None,
        Style::Single,
        Style::Thick,
        Style::Double,
        Style::Dotted,
        Style::Dashed,
        Style::DotDash,
        Style::DotDotDash,
        Style::Triple,
        Style::ThinThickSmallGap,
        Style::ThinThickMediumGap,
        Style::ThinThickLargeGap,
        Style::ThickThinSmallGap,
        Style::ThickThinMediumGap,
        Style::ThickThinLargeGap,
        Style::ThinThickThinSmallGap,
        Style::ThinThickThinMediumGap,
        Style::ThinThickThinLargeGap,
        Style::Wave,
        Style::DoubleWave,
        Style::DashSmallGap,
        Style::DashDotStroked,
        Style::ThreeDEmboss,
        Style::ThreeDEngrave,
        Style::Outset,
        Style::Inset,
    ];
    for style in &styles {
        assert_eq!(&Style::parse(style.as_str()).unwrap(), style);
    }
    assert_eq!(size_of::<Art>(), 1);
    assert_eq!(Art::ALL.len(), 164);
    for (index, art) in Art::ALL.iter().enumerate() {
        assert_eq!(art.token().parse::<Art>().unwrap(), *art);
        assert_eq!(Style::parse(art.token()).unwrap(), (*art).into());
        assert_eq!(art.to_string(), art.token());
        let code = 0x40 + u8::try_from(index).unwrap();
        assert_eq!(art.code(), code);
        assert_eq!(Art::try_from(code).unwrap(), *art);
    }
    assert!(Art::try_from(0x3F).is_err());
    assert!(Art::try_from(0xE4).is_err());
    assert_eq!(Style::parse("apples").unwrap(), Style::Art(Art::Apples));
    for invalid in [
        "custom",
        "earth3",
        "triangle1",
        "triangle2",
        "triangleCircle1",
        "triangleCircle2",
        "shapes1",
        "shapes2",
        "unknownArt",
    ] {
        assert!(Style::parse(invalid).is_err(), "accepted {invalid}");
    }
    assert!(Style::parse("not a style!").is_err());
    assert!(Style::parse("").is_err());
}

#[test]
fn rejects_malformed_page_layout_properties() {
    // Unknown enum tokens.
    assert!(
        SectionProperties::from_xml("<w:sectPr><w:vAlign w:val=\"diagonal\"/></w:sectPr>").is_err()
    );
    assert!(
        SectionProperties::from_xml("<w:sectPr><w:lnNumType w:restart=\"weekly\"/></w:sectPr>")
            .is_err()
    );
    assert!(
        SectionProperties::from_xml("<w:sectPr><w:pgBorders w:offsetFrom=\"margin\"/></w:sectPr>")
            .is_err()
    );
    assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"single\"/><w:top w:val=\"thick\"/></w:pgBorders></w:sectPr>").is_err());
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:pgBorders><w:diagonal w:val=\"single\"/></w:pgBorders></w:sectPr>"
        )
        .is_err()
    );
    assert!(
        SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top/></w:pgBorders></w:sectPr>")
            .is_err()
    );
    // Out-of-bounds values rejected through validation.
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:pgBorders><w:top w:val=\"single\" w:sz=\"97\"/></w:pgBorders></w:sectPr>"
        )
        .is_err()
    );
    assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"single\" w:space=\"32\"/></w:pgBorders></w:sectPr>").is_err());
    assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"starsTop\" w:sz=\"1639\"/></w:pgBorders></w:sectPr>").is_err());
    assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"single\" w:color=\"FFF\"/></w:pgBorders></w:sectPr>").is_err());
    // Schema-order violations.
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:pgNumType w:fmt=\"decimal\"/><w:lnNumType w:countBy=\"5\"/></w:sectPr>"
        )
        .is_err()
    );
    assert!(
        SectionProperties::from_xml(
            "<w:sectPr><w:printerSettings r:id=\"rId1\"/><w:docGrid/></w:sectPr>"
        )
        .is_err()
    );
    // Empty relationship ID rejected through validation.
    let section = SectionProperties {
        printer_settings_relationship_id: Some(String::new()),
        ..SectionProperties::default()
    };
    assert!(section.validate().is_err());
}
