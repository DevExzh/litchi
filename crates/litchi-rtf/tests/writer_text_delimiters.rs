//! Body text must survive a write/re-read cycle.
//!
//! RTF control words are terminated by a delimiter, so a break emitted directly
//! against the text that follows it (`\partwo`) is read back as one unknown
//! control word and the text is lost. Likewise, readers discard bare control
//! bytes such as a carriage return as line-ending noise, so they have to be
//! written as hex escapes. These tests pin both behaviours.

use litchi_rtf::{RtfDocument, RtfWriter};

/// Serialize `doc` and return the emitted RTF.
fn write(doc: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(doc)
        .expect("writing the document failed");
    String::from_utf8(bytes).expect("writer emits UTF-8")
}

/// Parse `source`, write it back out, and re-parse the result.
fn round_trip(source: &str) -> (String, String, String) {
    let doc = RtfDocument::parse(source).expect("source failed to parse");
    let emitted = write(&doc);
    let reparsed =
        RtfDocument::from_bytes(emitted.as_bytes()).expect("emitted RTF failed to parse");
    (doc.text(), emitted, reparsed.text())
}

#[test]
fn paragraph_breaks_are_delimited_from_following_text() {
    let (original, emitted, reparsed) = round_trip(r"{\rtf1 one\par two\par three}");
    assert_eq!(original, "one\ntwo\nthree");
    assert!(
        !emitted.contains("\\partwo"),
        "the break fused with its text: {emitted}"
    );
    assert_eq!(reparsed, original, "text changed across a write/read cycle");
}

#[test]
fn paragraph_breaks_are_delimited_from_following_digits() {
    // Without a delimiter `\par2024` parses as `\par` with parameter 2024, so
    // the year silently disappears from the document.
    let (original, _, reparsed) = round_trip(r"{\rtf1 year\par 2024}");
    assert_eq!(original, "year\n2024");
    assert_eq!(reparsed, original);
}

#[test]
fn line_breaks_and_tabs_are_delimited_from_following_text() {
    let (original, _, reparsed) = round_trip("{\\rtf1 foo\\line bar\\tab baz}");
    assert_eq!(original, "foo\nbar\tbaz");
    assert_eq!(reparsed, original);
}

#[test]
fn carriage_returns_survive_as_hex_escapes() {
    let (original, emitted, reparsed) = round_trip(r"{\rtf1 foo\'0dbar}");
    assert_eq!(original, "foo\rbar");
    assert!(
        emitted.contains("\\'0d"),
        "the carriage return was not escaped: {emitted}"
    );
    assert_eq!(reparsed, original);
}

#[test]
fn libreoffice_fixtures_round_trip_their_body_text() {
    for fixture in ["paragraph-break-then-text.rtf", "hex-crlf-text.rtf"] {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/rtf")
            .join(fixture);
        let bytes = std::fs::read(&path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        let doc = RtfDocument::from_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let emitted = write(&doc);
        let reparsed = RtfDocument::from_bytes(emitted.as_bytes())
            .unwrap_or_else(|error| panic!("failed to re-parse {fixture}: {error}"));
        assert_eq!(reparsed.text(), doc.text(), "{fixture} lost body text");
        assert!(
            !doc.text().is_empty(),
            "{fixture} has no body text to check"
        );
    }
}
