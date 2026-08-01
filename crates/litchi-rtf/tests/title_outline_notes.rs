//! Round-trip tests for `\titlepg`, `\endnhere`, and `\outlinelevelN`.

use litchi_rtf::{RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn title_page_and_endnote_here_round_trip() {
    let source = r#"{\rtf1\ansi\titlepg\endnhere Body\par}"#;
    let document = RtfDocument::parse(source).unwrap();
    let properties = &document.sections()[0].properties;
    assert!(properties.title_page);
    assert!(properties.note_options.endnote_here);

    let output = write(&document);
    let serialized = String::from_utf8(output).unwrap();
    assert!(
        serialized.contains(r"\titlepg"),
        "missing titlepg in {serialized}"
    );
    assert!(
        serialized.contains(r"\endnhere"),
        "missing endnhere in {serialized}"
    );

    let reparsed = RtfDocument::parse(&serialized).unwrap();
    let properties = &reparsed.sections()[0].properties;
    assert!(properties.title_page);
    assert!(properties.note_options.endnote_here);
}

#[test]
fn outline_level_round_trips() {
    let source = r#"{\rtf1\ansi\outlinelevel2 Heading\par Body\par}"#;
    let document = RtfDocument::parse(source).unwrap();
    let heading = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Heading"))
        .expect("heading block");
    assert_eq!(heading.paragraph.outline_level, Some(2));

    let output = write(&document);
    let serialized = String::from_utf8(output).unwrap();
    assert!(
        serialized.contains(r"\outlinelevel2"),
        "missing outlinelevel in {serialized}"
    );

    let reparsed = RtfDocument::parse(&serialized).unwrap();
    let heading = reparsed
        .blocks()
        .iter()
        .find(|block| block.text.contains("Heading"))
        .expect("heading block");
    assert_eq!(heading.paragraph.outline_level, Some(2));
}

#[test]
fn outline_level_out_of_range_is_rejected() {
    assert!(RtfDocument::parse(r"{\rtf1\ansi\outlinelevel10 X\par}").is_err());
    assert!(RtfDocument::parse(r"{\rtf1\ansi\outlinelevel-1 X\par}").is_err());
    assert!(RtfDocument::parse(r"{\rtf1\ansi\outlinelevel9 X\par}").is_ok());
}
