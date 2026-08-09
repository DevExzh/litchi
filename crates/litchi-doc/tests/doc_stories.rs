#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

//! Integration tests for DOC story subdocuments read from real Apache POI and
//! LibreOffice fixtures: headers/footers (`PlcfHdd`), footnotes
//! (`PlcffndRef`/`PlcffndTxt`), endnotes (`PlcfendRef`/`PlcfendTxt`), comments
//! (`PlcfandRef`/`PlcfandTxt` plus `GrpXstAtnOwners`), and HYPERLINK fields.

use litchi_doc::header_footer::HeaderFooter;
use litchi_doc::parts::headers::HeaderFooterType;
use litchi_doc::{Document, Package};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name)
}

fn open(path: PathBuf) -> Package {
    Package::open(&path).unwrap_or_else(|error| panic!("open {}: {error}", path.display()))
}

fn story_text(stories: &[HeaderFooter], story_type: HeaderFooterType) -> Option<&str> {
    stories
        .iter()
        .find(|story| story.header_footer_type == story_type)
        .map(HeaderFooter::text)
}

// ──────────────────────────────────────────────────────────────────
// Headers and footers (PlcfHdd)
// ──────────────────────────────────────────────────────────────────

#[test]
fn reads_odd_page_header_from_poi_three_col_head() {
    let mut package = open(fixture("ThreeColHead.doc"));
    let document = package.document().unwrap();

    let headers = document.headers().unwrap();
    assert_eq!(headers.len(), 1);
    assert_eq!(
        headers[0].header_footer_type,
        HeaderFooterType::OddPageHeader
    );
    assert_eq!(
        headers[0].text(),
        "First header column!\tMid header Right header!\r\r"
    );
    assert!(document.footers().unwrap().is_empty());
}

#[test]
fn reads_odd_page_footer_from_poi_three_col_foot() {
    let mut package = open(fixture("ThreeColFoot.doc"));
    let document = package.document().unwrap();

    let footers = document.footers().unwrap();
    assert_eq!(footers.len(), 1);
    assert_eq!(
        footers[0].header_footer_type,
        HeaderFooterType::OddPageFooter
    );
    assert_eq!(
        footers[0].text(),
        "Footer Left\tFooter Middle Footer Right\r\r"
    );
}

#[test]
fn reads_first_page_stories_from_poi_diff_first_page() {
    let mut package = open(fixture("DiffFirstPageHeadFoot.doc"));
    let document = package.document().unwrap();

    let stories = document.headers_footers().unwrap();
    assert_eq!(stories.len(), 4);
    assert_eq!(
        story_text(&stories, HeaderFooterType::FirstPageHeader),
        Some("I am the header on the first page, and I\u{2019}m nice and simple\r\r")
    );
    assert_eq!(
        story_text(&stories, HeaderFooterType::FirstPageFooter),
        Some("The footer of the first page\r\r")
    );
    assert_eq!(
        story_text(&stories, HeaderFooterType::OddPageHeader),
        Some("First header column!\tMid header Right header!\r\r")
    );
    assert_eq!(
        story_text(&stories, HeaderFooterType::OddPageFooter),
        Some("Footer Left\tFooter Middle Footer Right\r\r")
    );
}

#[test]
fn returns_no_stories_when_document_has_no_headers() {
    let mut package = open(fixture("NoHeadFoot.doc"));
    let document = package.document().unwrap();

    assert!(document.headers_footers().unwrap().is_empty());
    assert!(document.headers().unwrap().is_empty());
    assert!(document.footers().unwrap().is_empty());
}

#[test]
fn reads_all_six_story_kinds_per_section_from_libreoffice_fixture() {
    let mut package = open(fixture("first-header-footer.doc"));
    let document = package.document().unwrap();

    // Two sections, each with all six header/footer stories populated.
    let stories = document.headers_footers().unwrap();
    assert_eq!(stories.len(), 12);
    assert_eq!(document.headers().unwrap().len(), 6);
    assert_eq!(document.footers().unwrap().len(), 6);
    assert_eq!(
        story_text(&stories, HeaderFooterType::FirstPageHeader),
        Some("First page header\r\r")
    );
    assert_eq!(
        story_text(&stories, HeaderFooterType::EvenPageFooter),
        Some("Even page footer\r\r")
    );
    assert_eq!(stories[11].text(), "First page footer 2\r\r");
}

// ──────────────────────────────────────────────────────────────────
// Footnotes and endnotes (PlcffndRef/Txt, PlcfendRef/Txt)
// ──────────────────────────────────────────────────────────────────

#[test]
fn reads_footnotes_from_libreoffice_fixture() {
    let mut package = open(fixture("tdf71749_with_footnote.doc"));
    let document = package.document().unwrap();

    let footnotes = document.footnotes().unwrap();
    assert_eq!(footnotes.len(), 2);
    assert_eq!(footnotes[0].reference_position, 472);
    assert_eq!(footnotes[0].text(), "\u{2} Dummy.\r");
    assert_eq!(footnotes[1].reference_position, 485);
    assert!(
        footnotes[1]
            .text()
            .starts_with("\u{2} Production  details\r")
    );
    assert!(!footnotes[0].paragraphs().unwrap().is_empty());
    assert!(document.endnotes().unwrap().is_empty());
}

#[test]
fn reads_endnote_from_poi_endingnote() {
    let mut package = open(fixture("endingnote.doc"));
    let document = package.document().unwrap();

    let endnotes = document.endnotes().unwrap();
    assert_eq!(endnotes.len(), 1);
    assert_eq!(endnotes[0].reference_position, 10);
    assert_eq!(endnotes[0].number, 1);
    assert_eq!(endnotes[0].text(), "\u{2}\tEnding note text\r");
    assert!(document.footnotes().unwrap().is_empty());
}

#[test]
fn reads_multiple_endnotes_in_reference_order() {
    let mut package = open(fixture("3endnotes.doc"));
    let document = package.document().unwrap();

    let endnotes = document.endnotes().unwrap();
    assert_eq!(endnotes.len(), 3);
    let positions: Vec<u32> = endnotes
        .iter()
        .map(|endnote| endnote.reference_position)
        .collect();
    assert_eq!(positions, [784, 5450, 7466]);
    assert!(endnotes[0].text().contains("the first Reform Act"));
    assert!(endnotes[2].text().contains("Victorious Century"));
}

#[test]
fn reads_footnote_and_endnote_from_the_same_document() {
    let mut package = open(fixture("inline-endnote-and-footnote.doc"));
    let document = package.document().unwrap();

    let footnotes = document.footnotes().unwrap();
    let endnotes = document.endnotes().unwrap();
    assert_eq!(footnotes.len(), 1);
    assert_eq!(footnotes[0].text(), "\u{2} This is a footnote\r");
    assert_eq!(endnotes.len(), 1);
    assert_eq!(endnotes[0].text(), "\u{2} This is an endnote\r");
}

// ──────────────────────────────────────────────────────────────────
// Comments (PlcfandRef/Txt, GrpXstAtnOwners, annotation bookmarks)
// ──────────────────────────────────────────────────────────────────

#[test]
fn reads_range_comment_with_author_and_metadata() {
    let mut package = open(fixture("commented-table.doc"));
    let document = package.document().unwrap();

    let comments = document.comments().unwrap();
    assert_eq!(comments.len(), 1);
    let comment = &comments[0];
    assert_eq!(comment.author, "vmiklos");
    assert_eq!(comment.initials, "v");
    assert_eq!(comment.text, "hello\r");
    assert_eq!(
        (comment.range_start, comment.range_end),
        (Some(2), Some(19))
    );
    assert!(!comment.paragraphs.is_empty());

    let metadata = comment.extended_metadata.expect("ATRDPost10 metadata");
    assert_eq!(metadata.depth, 0);
    assert_eq!(metadata.parent_index, None);
    assert!(!metadata.is_ink);
    let modified = metadata.modified_at.expect("comment timestamp");
    assert_eq!(
        (modified.year, modified.month, modified.day),
        (2014, 12, 29)
    );
}

#[test]
fn rejects_comment_references_that_are_not_annotation_characters() {
    // LibreOffice anchors this comment on an inline image character instead
    // of the U+0005 annotation reference that [MS-DOC] 2.8.7 requires, so a
    // strict reader must report corruption instead of panicking.
    let mut package = open(fixture("image-comment-at-char.doc"));
    let document = package.document().unwrap();
    assert!(document.comments().is_err());
}

// ──────────────────────────────────────────────────────────────────
// Hyperlinks (HYPERLINK fields)
// ──────────────────────────────────────────────────────────────────

#[test]
fn reads_url_hyperlink_with_display_text() {
    let mut package = open(fixture("hyperlink.doc"));
    let document = package.document().unwrap();

    let hyperlinks = document.hyperlinks().unwrap();
    assert_eq!(hyperlinks.len(), 1);
    let hyperlink = &hyperlinks[0];
    assert_eq!(hyperlink.destination(), "http://testuri.org/");
    assert_eq!(hyperlink.display_text(), "Hyperlink text");
    assert!(hyperlink.is_url());
    assert!(!hyperlink.is_bookmark());

    let inside = (hyperlink.start_position + hyperlink.end_position) / 2;
    assert_eq!(document.hyperlinks_at_position(inside).len(), 1);
    assert!(document.hyperlinks_at_position(u32::MAX).is_empty());
}

// ──────────────────────────────────────────────────────────────────
// Malformed input
// ──────────────────────────────────────────────────────────────────

fn document_with_truncated_table_stream() -> Option<Document> {
    // Rewrite the fixture with its table stream cut short so every
    // table-anchored structure (PlcfHdd, note and comment PLCs) runs past
    // the stream bounds.
    let path = fixture("DiffFirstPageHeadFoot.doc");
    let mut package = Package::open(path).unwrap();
    let document = package.document().unwrap();
    let word_document = document.word_document().to_vec();

    let mut writer = litchi_cfb::OleWriter::new();
    writer
        .create_stream(&["WordDocument"], &word_document)
        .unwrap();
    writer.create_stream(&["1Table"], &[0u8; 16]).unwrap();
    let mut cursor = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();

    let mut reopened = Package::from_reader(std::io::Cursor::new(cursor.into_inner())).ok()?;
    reopened.document().ok()
}

#[test]
fn truncated_table_stream_is_rejected_without_panicking() {
    // Either the package parse fails outright or the affected lookups fail;
    // no code path may panic on the truncated stream.
    if let Some(document) = document_with_truncated_table_stream() {
        let _ = document.headers_footers();
        let _ = document.footnotes();
        let _ = document.endnotes();
        let _ = document.comments();
        let _ = document.hyperlinks();
    }
}

#[test]
fn garbage_bytes_are_rejected_without_panicking() {
    let garbage = vec![0xA5u8; 4096];
    assert!(Package::from_reader(std::io::Cursor::new(garbage)).is_err());

    let path = fixture("ThreeColHead.doc");
    let mut truncated = std::fs::read(path).unwrap();
    truncated.truncate(truncated.len() / 2);
    if let Ok(mut package) = Package::from_reader(std::io::Cursor::new(truncated)) {
        // If CFB parsing tolerates the truncation, document parsing must
        // still fail gracefully instead of panicking.
        let _ = package.document();
    }
}
