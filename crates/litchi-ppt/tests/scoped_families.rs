#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "scoped-family contract tests use panic-driven assertions"
)]

//! The record 0351 §2.3 contract of every scoped PPT family (record 0357):
//! `Package::edit_x` stages, `Edit::commit` publishes a new package with an
//! exact patch and diagnostics, an exact no-op publishes the source
//! unchanged, `Package::apply_x` replays a patch only onto its source, the
//! inverse restores the source bytes, and dropping an edit discards it.

use std::path::PathBuf;

use litchi_core::Position;
use litchi_core::edit::RefusalKind;
use litchi_ppt::font::Scope;
use litchi_ppt::{Package, broadcast, font};

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

fn built() -> Package {
    let mut builder = litchi_ppt::Builder::new();
    let slide = builder.add_slide().unwrap();
    builder
        .add_textbox(slide, 10, 10, 240, 40, "Scoped family")
        .unwrap();
    builder.build().unwrap()
}

fn slide_texts(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.text().unwrap().to_owned())
        .collect()
}

#[test]
fn font_edits_publish_replay_and_invert_exactly() {
    let source = built();
    let source_bytes = source.to_bytes().unwrap();

    let no_op = source.edit_fonts().unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_fonts().unwrap();
    let mut font = edit
        .fonts()
        .collection(Scope::Base)
        .unwrap()
        .get(0)
        .unwrap()
        .clone();
    font.name = "Scoped Font".into();
    edit.replace_font(Scope::Base, 0, font).unwrap();
    // Dropping a staged edit discards it.
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_fonts().unwrap();
    let mut font = edit
        .fonts()
        .collection(Scope::Base)
        .unwrap()
        .get(0)
        .unwrap()
        .clone();
    font.name = "Scoped Font".into();
    edit.replace_font(Scope::Base, 0, font).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().before_bytes(), source_bytes.as_slice());
    let fonts = commit.package().presentation().unwrap().fonts().unwrap();
    assert_eq!(fonts.get_base(0).unwrap().name, "Scoped Font");
    assert_eq!(slide_texts(commit.package()), slide_texts(&source));

    let replayed = source.apply_fonts(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_fonts(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_fonts(commit.patch()),
        Err(font::Error::PatchConflict)
    ));
}

#[test]
fn font_removal_is_a_typed_refusal() {
    let package = built();
    let mut edit = package.edit_fonts().unwrap();
    let Err(font::Error::Refused(refusal)) = edit.remove_font(Scope::Base, 0) else {
        panic!("font removal must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::UnsupportedOperation);
    assert!(edit.changes().is_empty());
}

#[test]
fn a_missing_broadcast_is_not_found() {
    let source = Package::from_bytes(&fixture("poi/test-data/slideshow/45543.ppt")).unwrap();
    assert!(
        source
            .presentation()
            .unwrap()
            .broadcasts()
            .unwrap()
            .broadcasts
            .is_empty()
    );
    let Err(broadcast::Error::Refused(refusal)) = source.edit_broadcast(Position::new(0)) else {
        panic!("a presentation without broadcasts must refuse the edit");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    assert!(matches!(
        built().edit_broadcast(Position::new(0)),
        Err(broadcast::Error::Refused(refusal)) if refusal.kind() == RefusalKind::NotFound
    ));
}
