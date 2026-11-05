//! The structural `Edit` verbs (record 0358) publish exactly what the
//! crate-private `MutableDocument` writes for the same verbs, chained verbs
//! share one working document, and a failing verb leaves the source
//! untouched.

use crate::header_footer::Kind;
use crate::mutable::MutableDocument;
use crate::{Builder, Package};

fn source() -> Package {
    let mut builder = Builder::new();
    builder.add_paragraph("first").unwrap();
    builder.build().unwrap()
}

#[test]
fn chained_document_verbs_match_the_mutable_document() {
    let source = source();
    let mut edit = source.edit();
    edit.add_heading("Title", 1).unwrap();
    edit.add_paragraph("second").unwrap();
    edit.add_master_page("Report", "pm1").unwrap();
    edit.set_header_footer_text("Report", Kind::Header, "Header text")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());

    let mut expected = MutableDocument::from_document(source.document().unwrap()).unwrap();
    expected.add_heading("Title", 1).unwrap();
    expected.add_paragraph("second").unwrap();
    expected.add_master_page("Report", "pm1").unwrap();
    expected
        .set_header_footer_text("Report", Kind::Header, "Header text")
        .unwrap();
    assert_eq!(commit.package().as_bytes(), expected.to_bytes().unwrap());

    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn a_failing_document_verb_refuses_the_commit_and_keeps_the_source() {
    let source = source();
    let before = source.as_bytes().to_vec();
    let mut edit = source.edit();
    edit.add_paragraph("kept only on success").unwrap();
    edit.set_header_footer_text("Missing", Kind::Footer, "x")
        .unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(source.as_bytes(), before.as_slice());
}

#[test]
fn an_empty_edit_is_an_exact_noop() {
    let source = source();
    let commit = source.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), source.as_bytes());
}
