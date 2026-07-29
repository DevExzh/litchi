use litchi_rtf::{BorderStyle, RtfDocument, RtfWriter, StyleType};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_full_paragraph_borders_and_round_trips() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\colortbl;\red255\green0\blue0;}"#,
        r#"\pard\brdrt\brdrs\brdrw20\brsp10\brdrcf1\brdrb\brdrdot"#,
        r#"\brdrl\brdrhair\brdrw5\brdrr\brdrdb\brdrsh\brdrframe Text\par}"#,
    ))
    .unwrap();
    let borders = &document.blocks()[0].paragraph.borders;
    assert_eq!(borders.top.style, BorderStyle::Single);
    assert_eq!(borders.top.width, 20);
    assert_eq!(borders.top.space, 10);
    assert_eq!(borders.top.color_ref, 1);
    assert_eq!(borders.bottom.style, BorderStyle::Dotted);
    assert_eq!(borders.left.style, BorderStyle::Hairline);
    assert_eq!(borders.left.width, 5);
    assert_eq!(borders.right.style, BorderStyle::Double);
    assert!(borders.right.shadow);
    assert!(borders.right.frame);

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.blocks()[0].paragraph.borders, *borders);
}

#[test]
fn normalizes_box_segment_and_retains_bar_and_between() {
    let document = RtfDocument::parse(
        r"{\rtf1\pard\box\brdrs\brdrw20\brdrbar\brdrth\brdrw40\brdrbtw\brdrdot Text\par}",
    )
    .unwrap();
    let borders = &document.blocks()[0].paragraph.borders;
    for side in [&borders.top, &borders.bottom, &borders.left, &borders.right] {
        assert_eq!(side.style, BorderStyle::Single);
        assert_eq!(side.width, 20);
    }
    assert_eq!(borders.bar.style, BorderStyle::Thick);
    assert_eq!(borders.bar.width, 40);
    assert_eq!(borders.between.style, BorderStyle::Dotted);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\brdrbar"));
    assert!(serialized.contains("\\brdrbtw"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.blocks()[0].paragraph.borders, *borders);
}

#[test]
fn parses_style_entry_borders_and_keeps_character_borders_distinct() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\stylesheet{\s1\brdrt\brdrs\brdrw30 Bordered;}}"#,
        r#"{\*\defchp\chbrdr\brdrhair\brdrw5}\pard\chbrdr\brdrs\brdrw10 Text\par}"#,
    ))
    .unwrap();
    let style = document
        .stylesheet()
        .styles()
        .iter()
        .find(|style| style.style_type == StyleType::Paragraph)
        .unwrap();
    assert_eq!(
        style.paragraph.unwrap().borders.top.style,
        BorderStyle::Single
    );
    assert_eq!(style.paragraph.unwrap().borders.top.width, 30);

    let paragraph = &document.blocks()[0].paragraph;
    assert!(!paragraph.borders.has_any_border());
    let run = &document.runs()[0];
    assert_eq!(
        run.formatting.character_border.unwrap().style,
        litchi_rtf::CharacterBorderStyle::Single
    );

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(
        reparsed.stylesheet().styles()[0].paragraph.unwrap().borders.top.style,
        BorderStyle::Single
    );
}

#[test]
fn table_style_border_syntax_stays_out_of_paragraph_borders() {
    // Word's built-in table styles carry \trbrdr* row borders with shared
    // style components; those belong to the table-decoration machinery and
    // must not trip the paragraph-border segment rules.
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\stylesheet{\ql Normal;}"#,
        r#"{\*\ts20\tsrowd\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10"#,
        r#" \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 Table Grid;}}"#,
        r#"\pard\brdrt\brdrhair Text\par}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "Text\n");
    assert_eq!(
        document.blocks()[0].paragraph.borders.top.style,
        BorderStyle::Hairline
    );
}

#[test]
fn rejects_malformed_paragraph_borders() {
    let cases = [
        // Duplicate segment.
        r"{\rtf1\pard\brdrt\brdrt Text\par}",
        // Style without a segment.
        r"{\rtf1\pard\brdrs Text\par}",
        // Duplicate style on one segment.
        r"{\rtf1\pard\brdrt\brdrs\brdrdot Text\par}",
        // Width/color/space/shadow/frame without a segment.
        r"{\rtf1\pard\brdrw20 Text\par}",
        r"{\rtf1\pard\brdrcf1 Text\par}",
        r"{\rtf1\pard\brsp10 Text\par}",
        r"{\rtf1\pard\brdrsh Text\par}",
        r"{\rtf1\pard\brdrframe Text\par}",
        // Width without a parameter.
        r"{\rtf1\pard\brdrt\brdrw Text\par}",
        // Duplicate box segment.
        r"{\rtf1\pard\box\brdrs\box\brdrdot Text\par}",
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}
