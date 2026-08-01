use litchi_rtf::{RtfDocument, RtfWriter, SoftBreakKind};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_zero_width_break_characters_and_round_trips() {
    let document = RtfDocument::parse(r"{\rtf1\ansi a\zwbo b\zwnbo c\zwj d\zwnj e}").unwrap();
    assert_eq!(document.text(), "a\u{200B}b\u{FEFF}c\u{200D}d\u{200C}e");

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\zwbo "));
    assert!(serialized.contains("\\zwnbo "));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
}

#[test]
fn parses_soft_break_markers_and_round_trips() {
    let document =
        RtfDocument::parse(r"{\rtf1 a\softline b\softpage c\softcol d\softlheight240 e}").unwrap();
    assert_eq!(document.text(), "abcde");
    let breaks: Vec<_> = document.soft_breaks().collect();
    assert_eq!(breaks.len(), 4);
    assert_eq!(breaks[0].kind, SoftBreakKind::Line);
    assert_eq!(breaks[0].position, 1);
    assert_eq!(breaks[1].kind, SoftBreakKind::Page);
    assert_eq!(breaks[1].position, 2);
    assert_eq!(breaks[2].kind, SoftBreakKind::Column);
    assert_eq!(breaks[2].position, 3);
    assert_eq!(breaks[3].kind, SoftBreakKind::LineHeight(240));
    assert_eq!(breaks[3].position, 4);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\softline "));
    assert!(serialized.contains("\\softpage "));
    assert!(serialized.contains("\\softcol "));
    assert!(serialized.contains("\\softlheight240 "));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    let reparsed_breaks: Vec<_> = reparsed.soft_breaks().collect();
    assert_eq!(reparsed_breaks.len(), 4);
    assert_eq!(reparsed_breaks[3].kind, SoftBreakKind::LineHeight(240));
}

#[test]
fn rejects_malformed_soft_break_controls() {
    let cases = [
        // Parameterized parameterless marks.
        r"{\rtf1 a\softpage1 b}",
        r"{\rtf1 a\softcol0 b}",
        r"{\rtf1 a\softline2 b}",
        // Missing line-height parameter.
        r"{\rtf1 a\softlheight b}",
        // Out-of-range line height.
        r"{\rtf1 a\softlheight32768 b}",
        r"{\rtf1 a\softlheight-32769 b}",
        // Marks outside the main body story.
        r"{\rtf1 body{\footnote\softpage note}}",
        r"{\rtf1 body{\footnote\softlheight240 note}}",
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}
