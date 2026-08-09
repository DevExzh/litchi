#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{RtfDocument, RtfWriter};

const SPECIAL_CHARACTERS: &[(&str, &str)] = &[
    (r"\emdash", "\u{2014}"),
    (r"\endash", "\u{2013}"),
    (r"\emspace", "\u{2003}"),
    (r"\enspace", "\u{2002}"),
    (r"\qmspace", "\u{2005}"),
    (r"\bullet", "\u{2022}"),
    (r"\ltrmark", "\u{200E}"),
    (r"\rtlmark", "\u{200F}"),
    (r"\zwj", "\u{200D}"),
    (r"\zwnj", "\u{200C}"),
];

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn special_character_control_words_extract_their_unicode_text() {
    for (control, expected) in SPECIAL_CHARACTERS {
        let source = format!(r"{{\rtf1\ansi A{control} B}}");
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(document.text(), format!("A{expected}B"), "parsed {source}");
    }
}

#[test]
fn special_character_control_words_round_trip_through_the_writer() {
    for (control, expected) in SPECIAL_CHARACTERS {
        let source = format!(r"{{\rtf1\ansi A{control} B}}");
        let document = RtfDocument::parse(&source).unwrap();

        let output = write(&document);
        let serialized = String::from_utf8(output.clone()).unwrap();
        assert!(
            serialized.contains(&format!("{control} ")),
            "serialized {control} as {serialized}"
        );
        let reparsed = RtfDocument::parse_bytes(&output).unwrap();
        assert_eq!(
            reparsed.text(),
            format!("A{expected}B"),
            "reparsed {control}"
        );
    }
}

#[test]
fn dynamic_date_time_stamps_parse_without_extracted_text() {
    let document = RtfDocument::parse(r"{\rtf1\ansi A\chdate B\chdpl C\chdpa D\chtime E}").unwrap();
    assert_eq!(document.text(), "ABCDE");
}

#[test]
fn special_character_control_words_survive_inside_generated_list_markers() {
    let document = RtfDocument::parse(r"{\rtf1\ansi{\listtext\pard\plain\bullet\tab}B}").unwrap();
    let marker = &document.generated_list_markers()[0];
    assert_eq!(marker.text, "\u{2022}\t");

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.generated_list_markers()[0].text, marker.text);
}
