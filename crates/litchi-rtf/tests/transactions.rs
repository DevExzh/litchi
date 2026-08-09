#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{Document, edit::Error};

#[test]
fn body_text_commit_is_atomic_reversible_and_source_checked() {
    let source = Document::parse(r"{\rtf1\ansi Plain\par Body}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_text("Changed\nbody").unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(source.text(), "Plain\nBody");
    assert_eq!(commit.snapshot().text(), "Changed\nbody");
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operation_count(), 1);

    let applied = commit.patch().apply(&source).unwrap();
    assert!(applied.same_snapshot(commit.snapshot()));
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert!(restored.same_snapshot(&source));

    let other = Document::parse(r"{\rtf1\ansi Other}").unwrap();
    assert!(matches!(
        commit.patch().apply(&other),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn body_text_noop_shares_the_source_snapshot() {
    let source = Document::parse(r"{\rtf1\ansi Same}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_text("Same").unwrap();
    let commit = edit.commit().unwrap();

    assert!(!commit.diagnostics().changed());
    assert!(commit.snapshot().same_snapshot(&source));
}

#[test]
fn changed_body_text_refuses_opaque_syntax_without_touching_the_source() {
    let source = Document::parse(r"{\rtf1\ansi A\future42 B}").unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.replace_body_text("Changed").unwrap();

    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(source.text(), "AB");
}

#[test]
fn transaction_accepts_only_one_semantic_operation() {
    let source = Document::parse(r"{\rtf1\ansi One}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_text("Two").unwrap();
    assert!(matches!(
        edit.replace_body_text("Three"),
        Err(Error::OperationAlreadyStaged)
    ));
}

#[test]
fn body_text_splice_preserves_modeled_header_bytes_exactly() {
    let prefix = br"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss Arial;}}{\colortbl;\red1\green2\blue3;}{\stylesheet{\s0 Normal;}}{\info{\title Preserved;}}\f0 ";
    let source = [prefix.as_slice(), b"Original", b"}"].concat();
    let document = Document::from_bytes(&source).unwrap();
    let mut edit = document.edit();
    edit.replace_body_text("Changed\nbody").unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.snapshot().to_bytes().unwrap();

    assert_eq!(&output[..prefix.len()], prefix);
    assert_eq!(&output[prefix.len()..], br"Changed\par body}");
    assert_eq!(commit.snapshot().text(), "Changed\nbody");
}

#[test]
fn paragraph_text_edit_is_checked_local_and_reversible() {
    let source = Document::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let mut edit = source.edit();
    edit.replace_paragraph_text(1, "Changed").unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(source.text(), "First\nSecond\nThird");
    assert_eq!(commit.snapshot().text(), "First\nChanged\nThird");
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(commit.diagnostics().changed());
    assert!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .same_snapshot(&source)
    );
}

#[test]
fn paragraph_text_edit_rejects_out_of_range_before_staging() {
    let source = Document::parse(r"{\rtf1 Only}").unwrap();
    let mut edit = source.edit();
    assert!(matches!(
        edit.replace_paragraph_text(1, "never"),
        Err(Error::ParagraphOutOfRange {
            position: 1,
            count: 1
        })
    ));
    edit.replace_paragraph_text(0, "Changed").unwrap();
    assert_eq!(edit.commit().unwrap().snapshot().text(), "Changed");
}
