//! Round-trip tests for East Asian `\acc*` emphasis-mark character controls.

use litchi_rtf::{EmphasisMark, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

const MARKS: &[(EmphasisMark, &str)] = &[
    (EmphasisMark::Dot, r"\accdot"),
    (EmphasisMark::Comma, r"\acccomma"),
    (EmphasisMark::UnderDot, r"\accunderdot"),
    (EmphasisMark::Circle, r"\acccircle"),
];

#[test]
fn emphasis_marks_round_trip() {
    for (mark, control) in MARKS {
        let source = format!(r"{{\rtf1\ansi{control} Emphasized\par}}");
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(
            document.blocks()[0].formatting.emphasis_mark,
            *mark,
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
            reparsed.blocks()[0].formatting.emphasis_mark,
            *mark,
            "round-tripped {control}"
        );
    }
}

#[test]
fn explicit_accnone_matches_default_state() {
    let document = RtfDocument::parse(r"{\rtf1\ansi\accnone Plain\par}").unwrap();
    assert_eq!(
        document.blocks()[0].formatting.emphasis_mark,
        EmphasisMark::None
    );

    let output = write(&document);
    let reparsed = RtfDocument::parse(&String::from_utf8(output).unwrap()).unwrap();
    assert_eq!(
        reparsed.blocks()[0].formatting.emphasis_mark,
        EmphasisMark::None
    );
}

#[test]
fn emphasis_marks_are_mutually_exclusive_last_wins() {
    let document = RtfDocument::parse(r"{\rtf1\ansi\accdot\acccomma Marked\par}").unwrap();
    assert_eq!(
        document.blocks()[0].formatting.emphasis_mark,
        EmphasisMark::Comma
    );
}

#[test]
fn plain_resets_emphasis_mark() {
    let document = RtfDocument::parse(r"{\rtf1\ansi\accdot Marked\plain Unmarked\par}").unwrap();
    assert_eq!(
        document.blocks()[0].formatting.emphasis_mark,
        EmphasisMark::Dot
    );
    assert_eq!(
        document.blocks()[1].formatting.emphasis_mark,
        EmphasisMark::None
    );
}

#[test]
fn emphasis_mark_controls_reject_parameters() {
    for source in [
        r"{\rtf1\ansi\accdot1 Text\par}",
        r"{\rtf1\ansi\acccomma0 Text\par}",
        r"{\rtf1\ansi\accunderdot-1 Text\par}",
        r"{\rtf1\ansi\acccircle2 Text\par}",
        r"{\rtf1\ansi\accnone1 Text\par}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
