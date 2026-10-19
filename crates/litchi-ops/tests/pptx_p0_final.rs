//! The last PPTX P0 rows (record 0346): legacy comment author, initials and
//! date writes and `direction` on add; modern comment add and set; notes
//! and placeholder text on add; slide title, text and layout writes. Every
//! case runs on both owners.
#![cfg(feature = "pptx")]

use std::io::{Cursor, Read as _};

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode};

#[path = "support/pptx_presentation.rs"]
#[allow(
    dead_code,
    reason = "the shared fixture module serves several test targets"
)]
mod presentation;

fn files(bytes: &[u8]) -> [File; 2] {
    let source = litchi_ops::open_bytes(bytes.to_vec()).expect("source-backed PPTX");
    let eager = litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::default().source_mode(SourceMode::Eager),
    )
    .expect("eager PPTX");
    [source, eager]
}

fn value(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap_or_default()
        .into_owned()
}

/// One package member's text.
fn member(file: &File, name: &str) -> Option<String> {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut member = archive.by_name(name).ok()?;
    let mut text = String::new();
    member.read_to_string(&mut text).unwrap();
    Some(text)
}

#[test]
fn legacy_comment_date_and_attribution_are_set() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let dated = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("date", "2026-10-19T08:00:00Z"),
            )
            .unwrap();
        assert_eq!(
            value(&dated, "/slide[1]/comment[1]", "date").as_deref(),
            Some("2026-10-19T08:00:00Z")
        );
        // Ada's only comment: her author record is renamed in place.
        let renamed = dated
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new()
                    .with("author", "Ada King")
                    .with("initials", "AK"),
            )
            .unwrap();
        assert_eq!(
            value(&renamed, "/slide[1]/comment[1]", "author").as_deref(),
            Some("Ada King")
        );
        assert_eq!(
            value(&renamed, "/slide[1]/comment[1]", "initials").as_deref(),
            Some("AK")
        );
        let authors = member(&renamed, "ppt/commentAuthors.xml").unwrap();
        assert!(authors.contains(r#"name="Ada King""#), "{authors}");
        assert!(!authors.contains("Ada Lovelace"), "{authors}");
        // The text and the other comment are untouched.
        assert_eq!(text(&renamed, "/slide[1]/comment[1]"), "Check the numbers");
        assert_eq!(
            value(&renamed, "/slide[1]/comment[2]", "author").as_deref(),
            Some("Alan Turing")
        );
        // An equal assignment is an exact no-op.
        let same = renamed
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("author", "Ada King"),
            )
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), renamed.to_bytes().unwrap());
    }
}

#[test]
fn a_shared_author_is_forked_not_renamed() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        // A second comment by Ada makes her author record shared.
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "Another"),
            )
            .unwrap()
            .file;
        let comments = |file: &File| member(file, "ppt/comments/comment1.xml").unwrap();
        assert!(comments(&added).contains(r#"authorId="0""#));
        let forked = added
            .set(
                "/slide[1]/comment[3]",
                litchi_ops::Props::new().with("initials", "A."),
            )
            .unwrap();
        // The first comment keeps Ada's record; the third moves to a new
        // author of the same name with the new initials, at its index 1.
        assert_eq!(
            value(&forked, "/slide[1]/comment[1]", "initials").as_deref(),
            Some("AL")
        );
        assert_eq!(
            value(&forked, "/slide[1]/comment[3]", "initials").as_deref(),
            Some("A.")
        );
        assert_eq!(
            value(&forked, "/slide[1]/comment[3]", "author").as_deref(),
            Some("Ada Lovelace")
        );
        let xml = comments(&forked);
        assert_eq!(xml.matches(r#"authorId="2""#).count(), 1, "{xml}");
        assert_eq!(text(&forked, "/slide[1]/comment[3]"), "Another");
        let authors = member(&forked, "ppt/commentAuthors.xml").unwrap();
        assert_eq!(authors.matches("<p:cmAuthor ").count(), 3, "{authors}");
        // The forked author is now the comment's own: a later change
        // renames it in place, as OfficeCLI does.
        let back = forked
            .set(
                "/slide[1]/comment[3]",
                litchi_ops::Props::new().with("initials", "AL"),
            )
            .unwrap();
        assert_eq!(
            value(&back, "/slide[1]/comment[3]", "initials").as_deref(),
            Some("AL")
        );
        assert_eq!(comments(&back), comments(&forked));
    }
}

#[test]
fn legacy_comment_writes_refuse_bad_values() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let error = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("date", "yesterday"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("author", ""),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        let error = file
            .set(
                "/slide[1]/comment[1]",
                litchi_ops::Props::new().with("direction", "rtl"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
    }
}

#[test]
fn comment_direction_leads_right_to_left_text_with_a_mark() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "مرحبا")
                    .prop("dir", "rtl"),
            )
            .unwrap()
            .file;
        assert_eq!(text(&added, "/slide[1]/comment[3]"), "\u{200F}مرحبا");
        let ltr = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "plain")
                    .prop("direction", "ltr"),
            )
            .unwrap()
            .file;
        assert_eq!(text(&ltr, "/slide[1]/comment[3]"), "plain");
        let error = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("direction", "up"),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
    }
}
