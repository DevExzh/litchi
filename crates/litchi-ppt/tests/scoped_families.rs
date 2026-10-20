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
use litchi_ppt::{Package, broadcast, external_media, font};

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

/// The path of movie `id` in the live document. The edit's read view is used
/// because `Presentation::external_media` scans every `DocumentContainer` of
/// the stream and refuses an append-only republished package.
fn movie_path(package: &Package, id: u32) -> Option<String> {
    let edit = package.edit_external_media().unwrap();
    match edit
        .objects()
        .iter()
        .find(|object| object.id() == id)
        .expect("the fixture's movie")
    {
        external_media::Object::Movie(movie) => movie.video.path.clone(),
        _ => panic!("the fixture's media object is a movie"),
    }
}

#[test]
fn external_media_edits_publish_replay_and_invert_exactly() {
    let source = Package::from_bytes(&fixture("unified-ops/0294/video.ppt")).unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let id = 1;
    let original = movie_path(&source, id);

    let no_op = source.edit_external_media().unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_external_media().unwrap();
    edit.set_path(id, Some("dropped.avi".into())).unwrap();
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_external_media().unwrap();
    let Err(external_media::Error::Refused(refusal)) = edit.set_path(99, Some("x.avi".into()))
    else {
        panic!("a missing media id must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    let Err(external_media::Error::Refused(refusal)) = edit.remove(99) else {
        panic!("a missing media id must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    if edit.owner_count(id) > 0 {
        let Err(external_media::Error::Refused(refusal)) = edit.remove(id) else {
            panic!("an owned media object cannot be removed");
        };
        assert_eq!(refusal.kind(), RefusalKind::UnsupportedTarget);
    }
    assert!(edit.changes().is_empty());
    edit.set_path(id, Some("scoped.avi".into())).unwrap();
    assert_eq!(edit.changes().len(), 1);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().before_bytes(), source_bytes.as_slice());
    assert_eq!(
        movie_path(commit.package(), id).as_deref(),
        Some("scoped.avi")
    );
    assert_eq!(movie_path(&source, id), original);
    assert_eq!(slide_texts(commit.package()), slide_texts(&source));

    let replayed = source.apply_external_media(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_external_media(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_external_media(commit.patch()),
        Err(external_media::Error::PatchConflict)
    ));
}
