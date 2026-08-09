#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{RtfDocument, RtfWriter, ShadingPattern};

fn write(document: &RtfDocument<'_>) -> String {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    String::from_utf8(output).unwrap()
}

#[test]
fn smart_quote_symbols_have_exact_unicode_semantics_and_canonical_output() {
    let document =
        RtfDocument::parse(r"{\rtf1\lquote left\rquote  \ldblquote double\rdblquote}").unwrap();
    assert_eq!(document.text(), "‘left’ “double”");

    let output = write(&document);
    for control in [r"\lquote ", r"\rquote ", r"\ldblquote ", r"\rdblquote "] {
        assert!(
            output.contains(control),
            "canonical output omitted {control}: {output}"
        );
    }
    assert_eq!(RtfDocument::parse(&output).unwrap().text(), document.text());
}

#[test]
fn paragraph_and_character_pattern_shading_round_trip_typed() {
    let document = RtfDocument::parse(
        r"{\rtf1\bgdkhoriz\cfpat2\cbpat3 Paragraph {\chbgdkdcross\chcfpat4\chcbpat5 Character}}",
    )
    .unwrap();
    let paragraph = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Paragraph"))
        .unwrap();
    assert_eq!(
        paragraph.paragraph.shading.pattern,
        Some(ShadingPattern::DarkHorizontal)
    );
    let character = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Character"))
        .unwrap()
        .formatting
        .character_shading
        .unwrap();
    assert_eq!(character.amount, None);
    assert_eq!(character.pattern, Some(ShadingPattern::DarkDiagonalCross));
    assert_eq!(character.foreground_color, Some(4));
    assert_eq!(character.background_color, Some(5));

    let output = write(&document);
    assert!(output.contains(r"\bgdkhoriz\cfpat2\cbpat3"));
    assert!(output.contains(r"\chbgdkdcross\chcfpat4\chcbpat5"));
    let reparsed = RtfDocument::parse(&output).unwrap();
    assert_eq!(write(&reparsed), output);
}

#[test]
fn shading_pattern_controls_reject_numeric_parameters() {
    for source in [r"{\rtf1\bghoriz1 invalid}", r"{\rtf1\chbgvert0 invalid}"] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn html_conversion_controls_are_typed_but_inert_for_rtf_display() {
    let source = r"{\rtf1\htmlrtf Conversion-only\htmlrtf0 Visible{\*\htmltag7 <b>}}";
    let document = RtfDocument::parse(source).unwrap();

    assert_eq!(document.text(), "Conversion-onlyVisible");
    assert_eq!(document.opaque_nodes().len(), 3);
    assert_eq!(document.opaque_nodes()[0].source(), br"\htmlrtf ");
    assert_eq!(document.opaque_nodes()[1].source(), br"\htmlrtf0 ");
    assert_eq!(document.opaque_nodes()[2].source(), br"{\*\htmltag7 <b>}");
}

#[test]
fn html_tag_requires_a_parameter_and_starred_destination() {
    for source in [
        r"{\rtf1{\*\htmltag missing}}",
        r"{\rtf1\htmltag7 misplaced}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}
