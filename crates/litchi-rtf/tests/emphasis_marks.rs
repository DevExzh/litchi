#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

//! Round-trip tests for East Asian `\acc*` emphasis-mark character controls.

mod repair_support;

use litchi_rtf::raw::Tree;
use litchi_rtf::{EmphasisMark, write::Writer};

const MARKS: &[(EmphasisMark, &str)] = &[
    (EmphasisMark::Dot, r"\accdot"),
    (EmphasisMark::Comma, r"\acccomma"),
    (EmphasisMark::UnderDot, r"\accunderdot"),
    (EmphasisMark::Circle, r"\acccircle"),
];

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn emphasis_marks_round_trip() {
    for (mark, control) in MARKS {
        let source = format!(r"{{\rtf1\ansi{control} Emphasized\par}}");
        let document = Tree::parse(&source, litchi_rtf::read::OpenOptions::default()).unwrap();
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

        let reparsed = Tree::parse(&serialized, litchi_rtf::read::OpenOptions::default()).unwrap();
        assert_eq!(
            reparsed.blocks()[0].formatting.emphasis_mark,
            *mark,
            "round-tripped {control}"
        );
    }
}

#[test]
fn explicit_accnone_matches_default_state() {
    let document = Tree::parse(
        r"{\rtf1\ansi\accnone Plain\par}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        document.blocks()[0].formatting.emphasis_mark,
        EmphasisMark::None
    );

    let output = write(&document);
    let reparsed = Tree::parse(
        &String::from_utf8(output).unwrap(),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        reparsed.blocks()[0].formatting.emphasis_mark,
        EmphasisMark::None
    );
}

#[test]
fn emphasis_marks_are_mutually_exclusive_last_wins() {
    let document = Tree::parse(
        r"{\rtf1\ansi\accdot\acccomma Marked\par}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        document.blocks()[0].formatting.emphasis_mark,
        EmphasisMark::Comma
    );
}

#[test]
fn plain_resets_emphasis_mark() {
    let document = Tree::parse(
        r"{\rtf1\ansi\accdot Marked\plain Unmarked\par}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
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
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                source,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed {source}"
        );
    }
}
