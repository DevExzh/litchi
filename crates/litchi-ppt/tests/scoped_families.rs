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
use litchi_ppt::embedded::object;
use litchi_ppt::font::Scope;
use litchi_ppt::master::Kind;
use litchi_ppt::master::metadata as master_metadata;
use litchi_ppt::{Package, broadcast, external_media, font, master_layout};

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

/// The path of movie `id` in the live document, read through the
/// presentation reader (which reads the live `DocumentContainer` of an
/// append-only republished package, record 0357).
fn movie_path(package: &Package, id: u32) -> Option<String> {
    let media = package
        .presentation()
        .unwrap()
        .external_media()
        .unwrap()
        .expect("the fixture's media collection");
    match media
        .objects
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

/// A deck with a main master and a handout master.
fn mastered() -> Package {
    Package::from_bytes(&fixture("poi/test-data/slideshow/headers_footers.ppt")).unwrap()
}

fn layout_record(package: &Package, kind: Kind) -> litchi_ppt::records::Record {
    package
        .edit_master_layout(kind, Position::new(0))
        .unwrap()
        .record()
        .clone()
}

#[test]
fn master_metadata_edits_publish_replay_and_invert_exactly() {
    let source = mastered();
    let source_bytes = source.to_bytes().unwrap();
    let main = Position::new(0);

    let no_op = source
        .edit_master_metadata(Kind::Main, main)
        .unwrap()
        .commit()
        .unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    // Setting and clearing an absent name is an exact no-op.
    let mut edit = source.edit_master_metadata(Kind::Main, main).unwrap();
    assert!(edit.name().unwrap().is_none());
    edit.set_name("Scoped master").unwrap();
    assert!(edit.clear_name().unwrap());
    let net_no_op = edit.commit().unwrap();
    assert!(!net_no_op.diagnostics().changed());
    assert_eq!(net_no_op.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_master_metadata(Kind::Main, main).unwrap();
    edit.set_name("Scoped master").unwrap();
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let handout_before = layout_record(&source, Kind::Handout);
    let mut edit = source.edit_master_metadata(Kind::Main, main).unwrap();
    edit.set_name("Scoped master").unwrap();
    edit.set_template_name("Scoped design").unwrap();
    assert_eq!(edit.name().unwrap().unwrap().as_str(), "Scoped master");
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().before_bytes(), source_bytes.as_slice());
    let published = commit
        .package()
        .edit_master_metadata(Kind::Main, main)
        .unwrap();
    assert_eq!(published.name().unwrap().unwrap().as_str(), "Scoped master");
    assert_eq!(
        published.template_name().unwrap().unwrap().as_str(),
        "Scoped design"
    );
    // The masters the edit does not name keep their records.
    assert_eq!(
        layout_record(commit.package(), Kind::Handout),
        handout_before
    );

    let replayed = source.apply_master_metadata(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_master_metadata(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_master_metadata(commit.patch()),
        Err(master_metadata::Error::PatchConflict)
    ));

    // Clearing the published name republishes the master without it.
    let mut clear = commit
        .package()
        .edit_master_metadata(Kind::Main, main)
        .unwrap();
    assert!(clear.clear_name().unwrap());
    let cleared = clear.commit().unwrap();
    assert!(cleared.diagnostics().changed());
    assert!(
        cleared
            .package()
            .edit_master_metadata(Kind::Main, main)
            .unwrap()
            .name()
            .unwrap()
            .is_none()
    );

    // A design name belongs to a main master only.
    let mut handout = source
        .edit_master_metadata(Kind::Handout, Position::new(0))
        .unwrap();
    let Err(master_metadata::Error::Refused(_)) = handout.set_template_name("x") else {
        panic!("a handout master cannot carry a design name");
    };
    assert!(handout.changes().is_empty());

    for (kind, position) in [
        (Kind::Main, Position::new(99)),
        (Kind::Handout, Position::new(1)),
        (Kind::Title, Position::new(0)),
    ] {
        let Err(master_metadata::Error::Refused(refusal)) =
            source.edit_master_metadata(kind, position)
        else {
            panic!("a missing master must be refused");
        };
        assert_eq!(refusal.kind(), RefusalKind::NotFound);
    }
}

#[test]
fn master_layout_edits_publish_replay_and_invert_exactly() {
    let source = mastered();
    let source_bytes = source.to_bytes().unwrap();
    let main = Position::new(0);

    let no_op = source
        .edit_master_layout(Kind::Main, main)
        .unwrap()
        .commit()
        .unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let handout_before = layout_record(&source, Kind::Handout);
    let mut edit = source.edit_master_layout(Kind::Main, main).unwrap();
    let original = edit.record().clone();
    let count = original.children.len();
    let last = original.children[count - 1].clone();
    edit.append(master_layout::Path::root(), last.clone())
        .unwrap();
    assert_eq!(edit.changes().len(), 1);
    assert_eq!(edit.view().record().children.len(), count + 1);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let published = commit
        .package()
        .edit_master_layout(Kind::Main, main)
        .unwrap();
    assert_eq!(published.record().children.len(), count + 1);
    assert_eq!(published.record().children[count], last);
    assert_eq!(
        layout_record(commit.package(), Kind::Handout),
        handout_before
    );

    let replayed = source.apply_master_layout(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_master_layout(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_master_layout(commit.patch()),
        Err(master_layout::Error::PatchConflict)
    ));

    // Removing the appended record again restores the master's tree.
    let mut remove = commit
        .package()
        .edit_master_layout(Kind::Main, main)
        .unwrap();
    assert_eq!(
        remove
            .remove(master_layout::Path::root().child(count))
            .unwrap(),
        last
    );
    let removed = remove.commit().unwrap();
    assert!(removed.diagnostics().changed());
    assert_eq!(
        removed
            .package()
            .edit_master_layout(Kind::Main, main)
            .unwrap()
            .record(),
        &original
    );

    // An out-of-range path is a refused value.
    let mut edit = source.edit_master_layout(Kind::Main, main).unwrap();
    let Err(master_layout::Error::Refused(refusal)) =
        edit.remove(master_layout::Path::root().child(count + 5))
    else {
        panic!("an out-of-range path must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::InvalidValue);
    assert!(edit.changes().is_empty());

    for (kind, position) in [
        (Kind::Title, Position::new(99)),
        (Kind::Handout, Position::new(1)),
    ] {
        let Err(master_layout::Error::Refused(refusal)) = source.edit_master_layout(kind, position)
        else {
            panic!("a missing master must be refused");
        };
        assert_eq!(refusal.kind(), RefusalKind::NotFound);
    }
}

fn object_ids(package: &Package) -> Vec<u32> {
    package
        .presentation()
        .unwrap()
        .ole_objects()
        .unwrap()
        .unwrap()
        .objects
        .iter()
        .map(object::ExternalObject::id)
        .collect()
}

fn refusal_kind<T: std::fmt::Debug>(result: object::Result<T>) -> RefusalKind {
    match result {
        Err(object::Error::Refused(refusal)) => refusal.kind(),
        other => panic!("expected a refusal, got {other:?}"),
    }
}

#[test]
fn embedded_object_edits_publish_replay_and_invert_exactly() {
    let source_bytes = fixture("poi/test-data/slideshow/ppt_with_embeded.ppt");
    let source = Package::from_bytes(&source_bytes).unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let ids = object_ids(&source);
    assert_eq!(ids, [1, 2, 3, 4, 5, 6]);

    // An exact no-op, including a reorder back to the source order.
    let mut edit = source.edit_objects().unwrap();
    edit.reorder(&ids).unwrap();
    let no_op = edit.commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    // Dropping a staged edit discards it.
    let mut edit = source.edit_objects().unwrap();
    edit.reorder(&[6, 5, 4, 3, 2, 1]).unwrap();
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    // A changed commit changes the read projection: a reordered list and a
    // replaced storage behind the first object.
    let persist_id = source
        .presentation()
        .unwrap()
        .ole_objects()
        .unwrap()
        .unwrap()
        .objects[0]
        .persist_id();
    let storage = object_storage(b"replacement OLE2 bytes");
    let mut edit = source.edit_objects().unwrap();
    edit.reorder(&[6, 5, 4, 3, 2, 1]).unwrap();
    edit.replace_storage(persist_id, storage.clone()).unwrap();
    assert_eq!(
        edit.objects()
            .objects
            .iter()
            .map(object::ExternalObject::id)
            .collect::<Vec<_>>(),
        [6, 5, 4, 3, 2, 1]
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    let changed = commit.package().to_bytes().unwrap();
    assert_eq!(object_ids(commit.package()), [6, 5, 4, 3, 2, 1]);
    assert_eq!(
        commit
            .package()
            .presentation()
            .unwrap()
            .ole_storage(persist_id)
            .unwrap(),
        Some(storage)
    );
    assert_eq!(object_ids(&source), ids);

    // The patch replays onto its source and its inverse restores the source.
    let replayed = source.apply_objects(commit.patch()).unwrap();
    assert_eq!(replayed.package().to_bytes().unwrap(), changed);
    let restored = commit
        .package()
        .apply_objects(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        built().apply_objects(commit.patch()),
        Err(object::Error::PatchConflict)
    ));

    // Absent targets are typed refusals that leave the edit unchanged.
    let mut edit = source.edit_objects().unwrap();
    assert_eq!(refusal_kind(edit.remove(99)), RefusalKind::NotFound);
    assert_eq!(
        refusal_kind(edit.replace_storage(0xF_FFFF, object_storage(b"x"))),
        RefusalKind::NotFound
    );
    assert_eq!(
        refusal_kind(edit.reorder(&[1, 2, 3, 4, 5, 99])),
        RefusalKind::NotFound
    );
    assert_eq!(
        refusal_kind(edit.reorder(&[1, 2, 3])),
        RefusalKind::InvalidValue
    );
    assert_eq!(edit.objects().objects.len(), ids.len());
    assert!(!edit.commit().unwrap().diagnostics().changed());

    // A presentation without an external-object list has no target.
    assert_eq!(
        refusal_kind(built().edit_objects()),
        RefusalKind::UnsupportedTarget
    );
}

fn object_storage(data: &[u8]) -> litchi_ppt::embedded::storage::Storage {
    litchi_ppt::embedded::storage::Storage::uncompressed(
        litchi_ppt::embedded::storage::Kind::OleObject,
        data.to_vec(),
    )
    .unwrap()
}
