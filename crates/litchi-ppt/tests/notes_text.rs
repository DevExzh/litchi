#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "speaker-notes edit tests use panic-driven assertions"
)]

//! Speaker-notes text through the canonical primary edit (record 0260).

use litchi_ppt::package::edit::{Error, Position};
use litchi_ppt::slide_order;
use litchi_ppt::text_edit::Refusal;
use litchi_ppt::{Builder, Package};
use std::path::PathBuf;

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

/// Every slide's notes text (`None` without a notes page) and shape texts.
fn state(package: &Package) -> Vec<(Option<String>, Vec<String>)> {
    let presentation = package.presentation().unwrap();
    presentation
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            let notes = slide
                .speaker_notes()
                .unwrap()
                .map(|notes| notes.text().unwrap().to_owned());
            let shapes = slide
                .shapes()
                .unwrap()
                .iter()
                .map(|shape| shape.text().unwrap())
                .collect();
            (notes, shapes)
        })
        .collect()
}

fn refusal(error: &Error) -> Option<Refusal> {
    let mut source: Option<&(dyn std::error::Error + 'static)> = Some(error);
    while let Some(current) = source {
        if let Some(slide_order::Error::Text(text)) = current.downcast_ref::<slide_order::Error>()
            && let litchi_ppt::text_edit::Error::Refused(refusal) = text.as_ref()
        {
            return Some(refusal.clone());
        }
        source = current.source();
    }
    None
}

const DECK: &str = "poi/test-data/slideshow/basic_test_ppt_file.ppt";

#[test]
fn notes_text_changes_length_and_preserves_everything_else() {
    let source_bytes = fixture(DECK);
    let source = Package::from_bytes(&source_bytes).unwrap();
    let before = state(&source);
    assert_eq!(
        before[0].0.as_deref(),
        Some("These are the notes for page 1")
    );

    let mut edit = source.edit().unwrap();
    edit.set_notes_text(
        Position::new(0),
        "Rewritten speaker notes, longer than before",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());

    let mut expected = before.clone();
    expected[0].0 = Some("Rewritten speaker notes, longer than before".to_owned());
    assert_eq!(state(commit.package()), expected);
    let published = commit.package().to_bytes().unwrap();
    assert_eq!(state(&Package::from_bytes(&published).unwrap()), expected);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    assert_eq!(
        source
            .apply(commit.patch())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        published
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );
}

#[test]
fn equal_notes_text_is_an_exact_noop() {
    let source = Package::from_bytes(&fixture(DECK)).unwrap();
    let mut edit = source.edit().unwrap();
    edit.set_notes_text(Position::new(1), state(&source)[1].0.clone().unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn a_slide_without_a_notes_page_is_a_typed_refusal() {
    let mut builder = Builder::new();
    builder.add_slide().unwrap();
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    let error = edit
        .set_notes_text(Position::new(0), "notes")
        .expect_err("no notes page");
    assert_eq!(refusal(&error), Some(Refusal::NoNotesPage));
    // The refused operation leaves the edit usable and unchanged.
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

#[test]
fn authored_notes_round_trip_through_the_edit() {
    let mut builder = Builder::new();
    let slide = builder.add_slide().unwrap();
    builder.set_slide_notes(slide, "Authored notes").unwrap();
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_notes_text(Position::new(0), "Edited notes")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        state(commit.package())[0].0.as_deref(),
        Some("Edited notes")
    );
}
