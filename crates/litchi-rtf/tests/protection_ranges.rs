#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{ProtectionRange, write::Writer};
use std::borrow::Cow;

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_overlapping_id_paired_ranges_and_round_trips() {
    // Layout adapted from the Word 2003 RTF specification example.
    let document = Tree::parse(
        concat!(
            r"{\rtf1 This is {\*\protstart 0300010003000000}SECTION 2. ",
            r"{\*\protstart 0200010004000000}This is SECTI{\*\protend 0300010003000000}",
            r"ON 3{\*\protend 0200010004000000}}",
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "This is SECTION 2. This is SECTION 3");
    let ranges = document.protection_ranges();
    assert_eq!(ranges.len(), 2);
    assert_eq!(ranges[0].id, "0300010003000000");
    assert_eq!(ranges[0].position, 8);
    assert_eq!(ranges[0].content, "SECTION 2. This is SECTI");
    assert_eq!(ranges[1].id, "0200010004000000");
    assert_eq!(ranges[1].position, 19);
    assert_eq!(ranges[1].content, "This is SECTION 3");

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.protection_ranges(), ranges);
}

#[test]
fn unclosed_range_extends_to_body_end_and_unmatched_end_is_ignored() {
    let document = Tree::parse(
        r"{\rtf1 ab{\*\protend ff}{\*\protstart 01}cd}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "abcd");
    let ranges = document.protection_ranges();
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].id, "01");
    assert_eq!(ranges[0].position, 2);
    assert_eq!(ranges[0].content, "cd");

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.protection_ranges(), ranges);
}

#[test]
fn coexists_with_bookmarks_and_body_markup() {
    let document = Tree::parse(
        r"{\rtf1 a{\*\bkmkstart bm}b{\*\bkmkend bm}{\*\protstart 0a0b}cd{\*\protend 0a0b}e}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "abcde");
    assert_eq!(document.bookmarks().bookmarks()[0].content, "b");
    let ranges = document.protection_ranges();
    assert_eq!(ranges.len(), 1);
    assert_eq!(ranges[0].position, 2);
    assert_eq!(ranges[0].content, "cd");

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.protection_ranges(), ranges);
    assert_eq!(
        reparsed.bookmarks().bookmarks()[0].content,
        document.bookmarks().bookmarks()[0].content
    );
}

#[test]
fn typed_constructor_validates_identifiers() {
    assert!(ProtectionRange::new(Cow::Borrowed(""), 0, Cow::Borrowed("")).is_err());
    assert!(ProtectionRange::new(Cow::Borrowed("zz"), 0, Cow::Borrowed("")).is_err());
    assert!(ProtectionRange::new(Cow::Borrowed("abc"), 0, Cow::Borrowed("")).is_err());
    let oversized = "0".repeat(66);
    assert!(ProtectionRange::new(Cow::Owned(oversized), 0, Cow::Borrowed("")).is_err());
    assert!(ProtectionRange::new(Cow::Borrowed("0300010003000000"), 0, Cow::Borrowed("x")).is_ok());
}

#[test]
fn rejects_malformed_protection_range_destinations() {
    let cases = [
        // Unstarred destinations.
        r"{\rtf1 a{\protstart 01}b{\protend 01}}",
        // Empty identifier.
        r"{\rtf1 a{\*\protstart }b}",
        // Non-hexadecimal identifier.
        r"{\rtf1 a{\*\protstart xy}b}",
        // Odd-length identifier.
        r"{\rtf1 a{\*\protstart 012}b}",
        // Grouped data inside the destination.
        r"{\rtf1 a{\*\protstart 01{x}}b}",
        // Binary data inside the destination.
        r"{\rtf1 a{\*\protstart 01\bin2 xx}b}",
        // Unterminated destination.
        r"{\rtf1 a{\*\protstart 01",
    ];
    for rtf in cases {
        assert!(
            Tree::parse(rtf, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {rtf}"
        );
    }
}
