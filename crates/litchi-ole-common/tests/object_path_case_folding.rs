#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration tests use concise assertions on fixture-sized CFB files"
)]

//! Object target paths compare CFB names exactly as `litchi-cfb` does.
//!
//! `[MS-CFB]` 2.6.4 compares names as UTF-16 code points under a simple
//! uppercase mapping that leaves each surrogate unchanged. "𐐀" (U+10400) and
//! "𐐨" (U+10428) differ only by Deseret case, so they are two distinct
//! sibling storages. The object owner's path resolution and stream removal
//! must not treat them as one name.

use litchi_cfb::OleWriter;
use litchi_ole_common::object::{Editor, Limits, Target, Targets};
use std::io::Cursor;

const UPPER: &str = "\u{10400}";
const LOWER: &str = "\u{10428}";

fn siblings() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_storage(&[UPPER]).expect("upper storage");
    writer
        .create_stream(&[UPPER, "Payload"], b"upper")
        .expect("upper payload");
    writer.create_storage(&[LOWER]).expect("lower storage");
    writer
        .create_stream(&[LOWER, "Payload"], b"lower")
        .expect("lower payload");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("test CFB should write");
    output.into_inner()
}

fn editor_for(bytes: Vec<u8>, storage: &str) -> Editor {
    let targets = Targets::one(Target::new("object", [storage]).expect("target should validate"));
    Editor::open(bytes, targets, Limits::default()).expect("editor should open")
}

#[test]
fn case_distinct_supplementary_siblings_resolve_to_their_own_storage() {
    let bytes = siblings();
    for (storage, payload) in [(UPPER, b"upper"), (LOWER, b"lower")] {
        let editor = editor_for(bytes.clone(), storage);
        let object = editor
            .objects()
            .get("object")
            .expect("object should be discovered");
        assert_eq!(
            object.stream(&["Payload"]),
            Some(&payload[..]),
            "{storage:?} resolved to its sibling"
        );
    }
}

#[test]
fn removing_one_siblings_stream_leaves_the_other() {
    let mut editor = editor_for(siblings(), LOWER);
    let removed = editor
        .remove_stream(&[LOWER.to_owned(), "Payload".to_owned()])
        .expect("removal");
    assert_eq!(removed.as_deref(), Some(&b"lower"[..]));
    assert_eq!(
        editor.stream(&[UPPER.to_owned(), "Payload".to_owned()]),
        Some(&b"upper"[..])
    );
    assert_eq!(
        editor.stream(&[LOWER.to_owned(), "Payload".to_owned()]),
        None
    );
}
