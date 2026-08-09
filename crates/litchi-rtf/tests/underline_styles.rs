#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

//! Round-trip tests for extended underline styles and `\ulcN` underline color.

use litchi_rtf::{RtfDocument, RtfWriter};

const STYLES: &[(&str, &str)] = &[
    (r"\ulhair", "Hairline"),
    (r"\ulthd", "ThickDotted"),
    (r"\ulthdash", "ThickDashed"),
    (r"\ulthdashd", "ThickDashDot"),
    (r"\ulthdashdd", "ThickDashDotDot"),
    (r"\ulthldash", "ThickLongDash"),
    (r"\ulldash", "LongDash"),
    (r"\ulhwave", "HeavyWave"),
    (r"\ululdbwave", "DoubleWave"),
];

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn extended_underline_styles_round_trip() {
    for (control, variant) in STYLES {
        let source = format!(r"{{\rtf1\ansi{control} Underlined\par}}");
        let document = RtfDocument::parse(&source).unwrap();
        let block = &document.blocks()[0];
        assert_eq!(
            format!("{:?}", block.formatting.underline),
            *variant,
            "parsed {control}"
        );

        let output = write(&document);
        let serialized = String::from_utf8(output).unwrap();
        assert!(
            serialized.contains(control),
            "missing {control} in {serialized}"
        );

        let reparsed = RtfDocument::parse(&serialized).unwrap();
        assert_eq!(
            format!("{:?}", reparsed.blocks()[0].formatting.underline),
            *variant,
            "round-tripped {control}"
        );
    }
}

#[test]
fn underline_color_round_trips() {
    let source =
        r"{\rtf1\ansi{\colortbl;\red0\green0\blue255;\red255\green0\blue0;}\uldb\ulc2 Text\par}";
    let document = RtfDocument::parse(source).unwrap();
    let block = &document.blocks()[0];
    assert_eq!(block.formatting.underline_color, Some(2));

    let output = write(&document);
    let serialized = String::from_utf8(output).unwrap();
    assert!(serialized.contains(r"\ulc2"), "missing ulc in {serialized}");

    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(reparsed.blocks()[0].formatting.underline_color, Some(2));
}

#[test]
fn underline_color_defaults_to_text_color() {
    let document = RtfDocument::parse(r"{\rtf1\ansi\ulwave Text\par}").unwrap();
    assert_eq!(document.blocks()[0].formatting.underline_color, None);
}
