use litchi_rtf::{
    MAX_PARAGRAPH_DROP_CAP_LINES, ParagraphDropCap, ParagraphDropCapKind, RtfDocument, RtfWriter,
    StyleBlock,
};

fn block<'a>(document: &'a RtfDocument<'a>, needle: &str) -> &'a StyleBlock<'a> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.contains(needle))
        .unwrap()
}

#[test]
fn parses_group_inheritance_and_pard_reset() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi\pard\dropcapli3\dropcapt1 Outer\par "#,
        r#"{\dropcapt2 Inner\par }Tail\par "#,
        r#"{\pard Reset\par }}"#,
    ))
    .unwrap();
    let outer = block(&document, "Outer").paragraph.drop_cap.unwrap();
    assert_eq!(outer.kind(), ParagraphDropCapKind::InText);
    assert_eq!(outer.line_count(), 3);
    let inner = block(&document, "Inner").paragraph.drop_cap.unwrap();
    assert_eq!(inner.kind(), ParagraphDropCapKind::Margin);
    assert_eq!(inner.line_count(), 3);
    assert_eq!(block(&document, "Tail").paragraph.drop_cap, Some(outer));
    assert_eq!(block(&document, "Reset").paragraph.drop_cap, None);
}

#[test]
fn body_style_and_default_round_trip_canonically() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi"#,
        r#"{\stylesheet{\s7\dropcapli2\dropcapt2 Drop;}}"#,
        r#"{\*\defpap\dropcapli4\dropcapt1}"#,
        r#"\pard\dropcapt2\dropcapli2 Body}"#,
    ))
    .unwrap();
    let expected = ParagraphDropCap::new(ParagraphDropCapKind::Margin, 2).unwrap();
    assert_eq!(block(&document, "Body").paragraph.drop_cap, Some(expected));
    assert_eq!(
        document
            .stylesheet()
            .get(7)
            .unwrap()
            .paragraph
            .unwrap()
            .drop_cap,
        Some(expected)
    );
    assert_eq!(
        document
            .default_formatting()
            .paragraph()
            .unwrap()
            .paragraph
            .drop_cap,
        Some(ParagraphDropCap::new(ParagraphDropCapKind::InText, 4).unwrap())
    );
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let text = String::from_utf8(first.clone()).unwrap();
    assert!(text.contains(r#"\dropcapli2\dropcapt2"#));
    assert!(text.contains(r#"\dropcapli4\dropcapt1"#));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(block(&reparsed, "Body").paragraph.drop_cap, Some(expected));
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_missing_partial_invalid_and_oversized_values() {
    for source in [
        r#"{\rtf1\dropcapli X}"#,
        r#"{\rtf1\dropcapt X}"#,
        r#"{\rtf1\dropcapli0\dropcapt1 X}"#,
        r#"{\rtf1\dropcapli-1\dropcapt1 X}"#,
        r#"{\rtf1\dropcapli256\dropcapt1 X}"#,
        r#"{\rtf1\dropcapli2\dropcapt0 X}"#,
        r#"{\rtf1\dropcapli2\dropcapt3 X}"#,
        r#"{\rtf1\dropcapli2 X}"#,
        r#"{\rtf1\dropcapt1 X}"#,
        r#"{\rtf1{\stylesheet{\s1\dropcapli2 Bad;}}X}"#,
        r#"{\rtf1{\*\defpap\dropcapt1}X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
    assert!(ParagraphDropCap::new(ParagraphDropCapKind::InText, 0).is_err());
    assert!(
        ParagraphDropCap::new(
            ParagraphDropCapKind::Margin,
            MAX_PARAGRAPH_DROP_CAP_LINES + 1,
        )
        .is_err()
    );
}
