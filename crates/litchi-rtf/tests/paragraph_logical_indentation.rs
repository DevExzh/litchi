#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{ParagraphLogicalIndentation, RtfDocument, RtfWriter, StyleBlock};
fn block<'a>(d: &'a RtfDocument<'a>, s: &str) -> &'a StyleBlock<'a> {
    d.blocks().iter().find(|b| b.text.contains(s)).unwrap()
}
#[test]
fn parses_inherits_resets_destinations() {
    let d=RtfDocument::parse(r"{\rtf1\pard\lin120\rin240\cufi-10\culi20\curi30\indmirror Outer\par {\lin0\rin0 Inner\par }Tail\par {\pard Reset\par }{\*\unknown\lin9 Ignored}Visible\par}").unwrap();
    let x = block(&d, "Outer").paragraph.logical_indentation;
    assert_eq!(
        x,
        ParagraphLogicalIndentation {
            start: Some(120),
            end: Some(240),
            first_line_character_units: Some(-10),
            left_character_units: Some(20),
            right_character_units: Some(30),
            mirrored: true
        }
    );
    assert_eq!(block(&d, "Tail").paragraph.logical_indentation, x);
    assert_eq!(
        block(&d, "Reset").paragraph.logical_indentation,
        ParagraphLogicalIndentation::default()
    );
    assert_eq!(block(&d, "Visible").paragraph.logical_indentation, x);
}
#[test]
fn stylesheet_writer_round_trip() {
    let d = RtfDocument::parse(
        r"{\rtf1{\stylesheet{\s9\lin120\rin240\cufi-10\culi20\curi30\indmirror Logical;}}Body}",
    )
    .unwrap();
    let x = d
        .stylesheet()
        .get(9)
        .unwrap()
        .paragraph
        .unwrap()
        .logical_indentation;
    let mut a = Vec::new();
    RtfWriter::new(&mut a).write_document(&d).unwrap();
    let d2 = RtfDocument::parse_bytes(&a).unwrap();
    assert_eq!(
        d2.stylesheet()
            .get(9)
            .unwrap()
            .paragraph
            .unwrap()
            .logical_indentation,
        x
    );
    let mut b = Vec::new();
    RtfWriter::new(&mut b).write_document(&d2).unwrap();
    assert_eq!(a, b);
}
#[test]
fn real_libreoffice() {
    let b = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/data/rtf/pass/tdf116851.rtf"
    );
    let d = RtfDocument::parse_bytes(b).unwrap();
    assert!(
        d.blocks()
            .iter()
            .any(|x| x.paragraph.logical_indentation.start.is_some())
    );
}
#[test]
fn malformed() {
    for s in [
        r"{\rtf1\lin X}",
        r"{\rtf1\rin X}",
        r"{\rtf1\cufi X}",
        r"{\rtf1\culi10000001 X}",
        r"{\rtf1\curi-10000001 X}",
        r"{\rtf1\indmirror0 X}",
    ] {
        assert!(RtfDocument::parse(s).is_err());
    }
}
