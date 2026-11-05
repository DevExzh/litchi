#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

mod repair_support;

use litchi_rtf::raw::Tree;
use litchi_rtf::write::Writer;

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_nested_editable_regions_and_round_trips() {
    let document = Tree::parse(
        r"{\rtf1 a\ebcstart b\ebcstart c\ebcend d\ebcend e}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "abcde");
    let regions = document.editable_regions();
    assert_eq!(regions.len(), 2);
    assert_eq!(regions[0].position, 1);
    assert_eq!(regions[0].content, "bcd");
    assert_eq!(regions[1].position, 2);
    assert_eq!(regions[1].content, "c");

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.editable_regions(), regions);
}

#[test]
fn parses_adjacent_regions_and_coexists_with_body_markup() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1{\*\protstart 01}a{\*\protend 01}"#,
            r#"\ebcstart b\ebcend\ebcstart cd\ebcend{\*\bkmkstart bm}e{\*\bkmkend bm}}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "abcde");
    let regions = document.editable_regions();
    assert_eq!(regions.len(), 2);
    assert_eq!(regions[0].position, 1);
    assert_eq!(regions[0].content, "b");
    assert_eq!(regions[1].position, 2);
    assert_eq!(regions[1].content, "cd");
    assert_eq!(document.protection_ranges().len(), 1);
    assert_eq!(document.bookmarks().bookmarks()[0].content, "e");

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.editable_regions(), regions);
    assert_eq!(reparsed.protection_ranges(), document.protection_ranges());
}

#[test]
fn rejects_unbalanced_or_misplaced_editable_region_marks() {
    let cases = [
        // End without a matching start.
        r"{\rtf1 a\ebcend b}",
        // Unclosed start.
        r"{\rtf1 a\ebcstart b}",
        // Improper nesting is impossible to express, but crossed reuse of an
        // already-closed mark is.
        r"{\rtf1 a\ebcstart b\ebcend\ebcend}",
        // Parameterized marks.
        r"{\rtf1 a\ebcstart1 b\ebcend}",
        r"{\rtf1 a\ebcstart b\ebcend0}",
        // Marks outside the main body story.
        r"{\rtf1 body{\footnote\ebcstart note\ebcend}}",
    ];
    for rtf in cases {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                rtf,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed {rtf}"
        );
    }
}
