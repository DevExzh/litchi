use litchi_rtf::{ParagraphFontAlignment, ParagraphWrapping, RtfDocument, RtfWriter, StyleBlock};

fn block<'a>(document: &'a RtfDocument<'a>, needle: &str) -> &'a StyleBlock<'a> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.contains(needle))
        .unwrap()
}

#[test]
fn parses_group_inheritance_resets_destinations_and_all_selectors() {
    let source = concat!(
        r#"{\rtf1\ansi\pard\hyphpar\aspalpha\aspnum\adjustright\nocwrap\fahang Outer\par "#,
        r#"{\hyphpar0\aspalpha0\aspnum0\adjustright0\nowwrap\facenter Inner\par }"#,
        r#"Tail\par {\pard Reset\par }"#,
        r#"{\*\unknown\hyphpar0\nooverflow\fafixed Ignored}"#,
        r#"\nooverflow\faroman Overflow\par "#,
        r#"\wrapdefault\favar Variable\par \fafixed Fixed\par }"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let outer = block(&document, "Outer").paragraph.line_breaking;
    assert!(outer.automatic_hyphenation && outer.auto_space_alphabetic && outer.auto_space_numbers);
    assert!(outer.adjust_right_indent);
    assert_eq!(outer.wrapping, ParagraphWrapping::NoCharacterWrap);
    assert_eq!(outer.font_alignment, ParagraphFontAlignment::Hanging);
    let inner = block(&document, "Inner").paragraph.line_breaking;
    assert_eq!(inner.wrapping, ParagraphWrapping::NoWordWrap);
    assert_eq!(inner.font_alignment, ParagraphFontAlignment::Center);
    assert!(
        !inner.automatic_hyphenation
            && !inner.auto_space_alphabetic
            && !inner.auto_space_numbers
            && !inner.adjust_right_indent
    );
    assert_eq!(block(&document, "Tail").paragraph.line_breaking, outer);
    assert_eq!(
        block(&document, "Reset").paragraph.line_breaking,
        Default::default()
    );
    assert_eq!(
        block(&document, "Overflow")
            .paragraph
            .line_breaking
            .wrapping,
        ParagraphWrapping::NoOverflow
    );
    assert_eq!(
        block(&document, "Overflow")
            .paragraph
            .line_breaking
            .font_alignment,
        ParagraphFontAlignment::Roman
    );
    assert_eq!(
        block(&document, "Variable")
            .paragraph
            .line_breaking
            .font_alignment,
        ParagraphFontAlignment::Variable
    );
    assert_eq!(
        block(&document, "Fixed")
            .paragraph
            .line_breaking
            .font_alignment,
        ParagraphFontAlignment::Fixed
    );
}

#[test]
fn deterministic_writer_and_stylesheet_round_trip() {
    let document = RtfDocument::parse(
        r#"{\rtf1{\stylesheet{\s7\hyphpar\aspalpha\aspnum\nooverflow\facenter\adjustright Body;}}\pard\hyphpar\aspalpha\aspnum\nooverflow\facenter\adjustright Body}"#,
    ).unwrap();
    let expected = document
        .stylesheet()
        .get(7)
        .unwrap()
        .paragraph
        .unwrap()
        .line_breaking;
    assert_eq!(expected, block(&document, "Body").paragraph.line_breaking);
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let text = String::from_utf8(first.clone()).unwrap();
    assert!(text.contains(r#"\hyphpar\nooverflow\aspalpha\aspnum\facenter\adjustright"#));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(
        reparsed
            .stylesheet()
            .get(7)
            .unwrap()
            .paragraph
            .unwrap()
            .line_breaking,
        expected
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn parses_real_libreoffice_fixture() {
    let bytes = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/data/rtf/pass/forcepoint-1.rtf"
    );
    let marker = br"{\stylesheet";
    let start = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .unwrap();
    let mut depth = 0usize;
    let mut end = None;
    for (offset, byte) in bytes[start..].iter().enumerate() {
        match byte {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    end = Some(start + offset + 1);
                    break;
                }
            },
            _ => {},
        }
    }
    let mut isolated = br"{\rtf1\ansi".to_vec();
    isolated.extend_from_slice(&bytes[start..end.unwrap()]);
    isolated.push(b'}');
    let document = RtfDocument::parse_bytes(&isolated).unwrap();
    let heading = document
        .stylesheet()
        .get(3)
        .unwrap()
        .paragraph
        .unwrap()
        .line_breaking;
    assert!(!heading.automatic_hyphenation);
    assert!(
        heading.auto_space_alphabetic && heading.auto_space_numbers && heading.adjust_right_indent
    );
    assert_eq!(heading.font_alignment, ParagraphFontAlignment::Auto);
}

#[test]
fn rejects_invalid_parameters() {
    for source in [
        r#"{\rtf1\hyphpar2 X}"#,
        r#"{\rtf1\aspalpha-1 X}"#,
        r#"{\rtf1\aspnum9 X}"#,
        r#"{\rtf1\adjustright3 X}"#,
        r#"{\rtf1\wrapdefault0 X}"#,
        r#"{\rtf1\nocwrap1 X}"#,
        r#"{\rtf1\nowwrap0 X}"#,
        r#"{\rtf1\nooverflow1 X}"#,
        r#"{\rtf1\faauto0 X}"#,
        r#"{\rtf1\fahang1 X}"#,
        r#"{\rtf1\facenter1 X}"#,
        r#"{\rtf1\faroman1 X}"#,
        r#"{\rtf1\favar1 X}"#,
        r#"{\rtf1\fafixed1 X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}
