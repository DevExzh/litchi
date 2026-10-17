//! Unified PPTX comment removal (record 0333): legacy comments, modern
//! threaded comments and their replies, on both owners. Authors are kept;
//! a slide's comment part goes with its last comment.
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

fn paths(file: &File, query: &str) -> Vec<String> {
    file.query(query)
        .expect("query")
        .map(|node| node.expect("node").path().to_string())
        .collect()
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

/// The package member names and one member's text.
fn members(file: &File) -> Vec<(String, String)> {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    (0..archive.len())
        .map(|index| {
            let mut member = archive.by_index(index).unwrap();
            let mut text = String::new();
            let _ = member.read_to_string(&mut text);
            (member.name().to_owned(), text)
        })
        .collect()
}

#[test]
fn legacy_comments_are_removed_and_authors_kept() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        assert_eq!(
            paths(&file, "comment"),
            ["/slide[1]/comment[1]", "/slide[1]/comment[2]"]
        );
        let second = text(&file, "/slide[1]/comment[2]");
        let removed = file.remove("/slide[1]/comment[1]").unwrap();
        assert_eq!(paths(&removed, "comment"), ["/slide[1]/comment[1]"]);
        assert_eq!(text(&removed, "/slide[1]/comment[1]"), second);
        let authors = |file: &File| {
            members(file)
                .into_iter()
                .find(|(name, _)| name == "ppt/commentAuthors.xml")
                .map(|(_, text)| text)
        };
        let before = authors(&file).expect("authors");
        // The last comment takes the slide's comments part with it.
        let empty = removed.remove("/slide[1]/comment[1]").unwrap();
        assert!(paths(&empty, "comment").is_empty());
        assert!(
            !members(&empty)
                .iter()
                .any(|(name, _)| name == "ppt/comments/comment1.xml")
        );
        assert_eq!(authors(&empty).as_deref(), Some(before.as_str()));
        // A later comment by a kept author takes a fresh index.
        let added = empty
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("author", "Ada Lovelace")
                    .prop("text", "again"),
            )
            .unwrap();
        assert_eq!(text(&added.file, "/slide[1]/comment[1]"), "again");
        assert_eq!(
            file.remove("/slide[1]/comment[3]").unwrap_err().code(),
            Code::PathNotFound
        );
        // The input handle is unchanged.
        assert_eq!(paths(&file, "comment").len(), 2);
    }
}

#[test]
fn modern_comments_and_replies_are_removed() {
    let bytes = presentation::bytes();
    for file in files(&bytes) {
        assert_eq!(
            paths(&file, "reply"),
            ["/slide[1]/moderncomment[1]/reply[1]"]
        );
        let reply = file.remove("/slide[1]/moderncomment[1]/reply[1]").unwrap();
        assert!(paths(&reply, "reply").is_empty());
        assert_eq!(paths(&reply, "modernComment").len(), 2);
        let first = reply.remove("/slide[1]/moderncomment[1]").unwrap();
        assert_eq!(
            paths(&first, "modernComment"),
            ["/slide[1]/moderncomment[1]"]
        );
        // A comment goes with its replies.
        let whole = file.remove("/slide[1]/modernComment[1]").unwrap();
        assert!(paths(&whole, "reply").is_empty());
        let last = first.remove("/slide[1]/moderncomment[1]").unwrap();
        assert!(paths(&last, "modernComment").is_empty());
        assert!(
            !members(&last)
                .iter()
                .any(|(name, _)| name == "ppt/comments/modernComment1.xml")
        );
        // The legacy comments are untouched.
        assert_eq!(paths(&last, "comment").len(), 2);
    }
}
