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
        .unwrap()
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
        .unwrap()
        .collection(Scope::Base)
        .unwrap()
        .get(0)
        .unwrap()
        .clone();
    font.name = "Scoped Font".into();
    edit.replace_font(Scope::Base, 0, font).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().source_len(), source_bytes.len());
    assert_eq!(
        commit.patch().target_len(),
        commit.package().to_bytes().unwrap().len()
    );
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
    assert!(edit.changes().unwrap().is_empty());
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
    let Err(broadcast::Error::Refused(refusal)) = source
        .edit_broadcast(Position::new(0))
        .and_then(|edit| edit.broadcast().map(|_| ()))
    else {
        panic!("a presentation without broadcasts must refuse the edit");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    assert!(matches!(
        built()
            .edit_broadcast(Position::new(0))
            .and_then(|edit| edit.broadcast().map(|_| ())),
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
    if edit.owner_count(id).unwrap() > 0 {
        let Err(external_media::Error::Refused(refusal)) = edit.remove(id) else {
            panic!("an owned media object cannot be removed");
        };
        assert_eq!(refusal.kind(), RefusalKind::UnsupportedTarget);
    }
    assert!(edit.changes().unwrap().is_empty());
    edit.set_path(id, Some("scoped.avi".into())).unwrap();
    assert_eq!(edit.changes().unwrap().len(), 1);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().source_len(), source_bytes.len());
    assert_eq!(
        commit.patch().target_len(),
        commit.package().to_bytes().unwrap().len()
    );
    assert_eq!(
        movie_path(commit.package(), id).as_deref(),
        Some("scoped.avi")
    );
    assert_eq!(movie_path(&source, id), original);
    assert_eq!(slide_texts(commit.package()), slide_texts(&source));
    // The patch is record-sized (record 0357a): the append-only publisher
    // appends the rewritten document container and touches a few header,
    // FAT and directory sectors; the patch carries only those, never the
    // whole package.
    let changed = commit.patch().changed_bytes();
    let grown = commit.patch().target_len() - source_bytes.len();
    assert!(changed >= grown);
    assert!(
        changed <= grown + 8 * 512 && changed.saturating_mul(2) < source_bytes.len(),
        "a one-path patch spans {changed} of {} bytes (grown by {grown})",
        source_bytes.len()
    );
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.changed_bytes(), changed);
    assert_eq!(inverse.source_len(), commit.patch().target_len());
    assert_eq!(inverse.target_len(), source_bytes.len());

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
        .unwrap()
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
    assert_eq!(commit.patch().source_len(), source_bytes.len());
    assert_eq!(
        commit.patch().target_len(),
        commit.package().to_bytes().unwrap().len()
    );
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
    assert!(handout.changes().unwrap().is_empty());

    for (kind, position) in [
        (Kind::Main, Position::new(99)),
        (Kind::Handout, Position::new(1)),
        (Kind::Title, Position::new(0)),
    ] {
        let Err(master_metadata::Error::Refused(refusal)) = source
            .edit_master_metadata(kind, position)
            .and_then(|edit| edit.name().map(|_| ()))
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
    let original = edit.record().unwrap().clone();
    let count = original.children.len();
    let last = original.children[count - 1].clone();
    edit.append(master_layout::Path::root(), last.clone())
        .unwrap();
    assert_eq!(edit.changes().unwrap().len(), 1);
    assert_eq!(edit.view().unwrap().record().children.len(), count + 1);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let published = commit
        .package()
        .edit_master_layout(Kind::Main, main)
        .unwrap();
    assert_eq!(published.record().unwrap().children.len(), count + 1);
    assert_eq!(published.record().unwrap().children[count], last);
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
            .record()
            .unwrap(),
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
    assert!(edit.changes().unwrap().is_empty());

    for (kind, position) in [
        (Kind::Title, Position::new(99)),
        (Kind::Handout, Position::new(1)),
    ] {
        let Err(master_layout::Error::Refused(refusal)) = source
            .edit_master_layout(kind, position)
            .and_then(|edit| edit.record().map(|_| ()))
        else {
            panic!("a missing master must be refused");
        };
        assert_eq!(refusal.kind(), RefusalKind::NotFound);
    }
}

fn corpus_decks() -> Vec<PathBuf> {
    fn walk(dir: &std::path::Path, out: &mut Vec<PathBuf>) {
        for entry in std::fs::read_dir(dir).unwrap() {
            let path = entry.unwrap().path();
            if path.is_dir() {
                walk(&path, out);
            } else if path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("ppt"))
            {
                out.push(path);
            }
        }
    }
    let mut decks = Vec::new();
    walk(
        &PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data"),
        &mut decks,
    );
    decks.sort();
    decks
}

/// Every `test-data/**/*.ppt` deck that opens lists its masters, and every
/// listed master (notes masters included: most real decks carry a
/// `MasterId`-domain `NotesAtom.slideIdRef` there) opens both master
/// families, whose untouched commit is the source package byte for byte. A
/// renamed notes master publishes and the result lists its masters again.
#[test]
fn every_corpus_deck_lists_its_masters_and_opens_their_edits() {
    let mut notes_masters = 0;
    let mut decks = 0;
    for path in corpus_decks() {
        let bytes = std::fs::read(&path).unwrap();
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        let name = path.display().to_string();
        decks += 1;
        let source_bytes = package.to_bytes().unwrap();
        let presentation = package.presentation().unwrap();
        let masters = presentation
            .masters()
            .unwrap_or_else(|error| panic!("{name}: masters() failed: {error:?}"));
        let mut listed = Vec::new();
        listed.extend((0..masters.main().len()).map(|index| (Kind::Main, index)));
        listed.extend((0..masters.title().len()).map(|index| (Kind::Title, index)));
        listed.extend(masters.notes().map(|_| (Kind::Notes, 0)));
        listed.extend(masters.handout().map(|_| (Kind::Handout, 0)));
        let has_notes = masters.notes().is_some();
        drop(masters);
        for (kind, index) in listed {
            let position = Position::new(index);
            let metadata = package
                .edit_master_metadata(kind, position)
                .unwrap_or_else(|error| panic!("{name}: {kind:?} metadata: {error:?}"))
                .commit()
                .unwrap();
            assert!(!metadata.diagnostics().changed(), "{name}");
            assert_eq!(
                metadata.package().to_bytes().unwrap(),
                source_bytes,
                "{name}"
            );
            let layout = package
                .edit_master_layout(kind, position)
                .unwrap_or_else(|error| panic!("{name}: {kind:?} layout: {error:?}"))
                .commit()
                .unwrap();
            assert!(!layout.diagnostics().changed(), "{name}");
            assert_eq!(layout.package().to_bytes().unwrap(), source_bytes, "{name}");
        }
        if has_notes {
            notes_masters += 1;
            let mut edit = package
                .edit_master_metadata(Kind::Notes, Position::new(0))
                .unwrap();
            edit.set_name("Scoped notes master").unwrap();
            let commit = edit.commit().unwrap();
            assert!(commit.diagnostics().changed(), "{name}");
            let renamed = commit
                .package()
                .edit_master_metadata(Kind::Notes, Position::new(0))
                .unwrap();
            assert_eq!(
                renamed.name().unwrap().unwrap().as_str(),
                "Scoped notes master",
                "{name}"
            );
            let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
            assert!(
                reopened
                    .presentation()
                    .unwrap()
                    .masters()
                    .unwrap()
                    .notes()
                    .is_some(),
                "{name}"
            );
        }
    }
    assert!(decks >= 25, "the corpus shrank to {decks} decks");
    assert!(
        notes_masters >= 13,
        "only {notes_masters} decks have a notes master"
    );
}

/// Apache POI `slide_master.ppt`: two main masters and one title master.
fn titled() -> Package {
    Package::from_bytes(&fixture("poi/test-data/slideshow/slide_master.ppt")).unwrap()
}

#[test]
fn title_master_metadata_edits_publish_replay_and_invert_exactly() {
    let source = titled();
    let source_bytes = source.to_bytes().unwrap();
    let title = Position::new(0);
    let masters = source.presentation().unwrap();
    let masters = masters.masters().unwrap();
    assert_eq!((masters.main().len(), masters.title().len()), (2, 1));
    drop(masters);

    let no_op = source
        .edit_master_metadata(Kind::Title, title)
        .unwrap()
        .commit()
        .unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let main_before = layout_record(&source, Kind::Main);
    let mut edit = source.edit_master_metadata(Kind::Title, title).unwrap();
    let before = edit.name().unwrap().map(|name| name.as_str().to_owned());
    edit.set_name("Scoped title master").unwrap();
    // A design name belongs to a main master only.
    assert!(matches!(
        edit.set_template_name("x"),
        Err(master_metadata::Error::Refused(_))
    ));
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let published = commit
        .package()
        .edit_master_metadata(Kind::Title, title)
        .unwrap();
    assert_eq!(
        published.name().unwrap().unwrap().as_str(),
        "Scoped title master"
    );
    assert_ne!(before.as_deref(), Some("Scoped title master"));
    assert_eq!(layout_record(commit.package(), Kind::Main), main_before);
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .presentation()
            .unwrap()
            .masters()
            .unwrap()
            .title()
            .len(),
        1
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
    let Err(master_metadata::Error::Refused(refusal)) = source
        .edit_master_metadata(Kind::Title, Position::new(1))
        .and_then(|edit| edit.name().map(|_| ()))
    else {
        panic!("a missing title master must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
}

#[test]
fn title_master_layout_edits_publish_replay_and_invert_exactly() {
    let source = titled();
    let source_bytes = source.to_bytes().unwrap();
    let title = Position::new(0);

    let no_op = source
        .edit_master_layout(Kind::Title, title)
        .unwrap()
        .commit()
        .unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let main_before = layout_record(&source, Kind::Main);
    let mut edit = source.edit_master_layout(Kind::Title, title).unwrap();
    let original = edit.record().unwrap().clone();
    let count = original.children.len();
    let last = original.children[count - 1].clone();
    edit.append(master_layout::Path::root(), last.clone())
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let published = layout_record(commit.package(), Kind::Title);
    assert_eq!(published.children.len(), count + 1);
    assert_eq!(published.children[count], last);
    assert_eq!(layout_record(commit.package(), Kind::Main), main_before);

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

    // Removing the appended record restores the title master's tree.
    let mut remove = commit
        .package()
        .edit_master_layout(Kind::Title, title)
        .unwrap();
    assert_eq!(
        remove
            .remove(master_layout::Path::root().child(count))
            .unwrap(),
        last
    );
    let removed = remove.commit().unwrap();
    assert_eq!(layout_record(removed.package(), Kind::Title), original);
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
            .unwrap()
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
    assert_eq!(edit.objects().unwrap().objects.len(), ids.len());
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

fn embedded_object(id: u32) -> object::ExternalObject {
    object::ExternalObject::Object(object::Definition {
        kind: object::ContainerKind::Embedded(object::EmbedPreferences {
            color_follow: object::ColorFollow::EntireScheme,
            cannot_lock_server: false,
            dimension_policy: object::DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }),
        object: object::Metadata {
            draw_aspect: object::DrawAspect::Content,
            object_type: object::ObjectType::Embedded,
            id,
            subtype: object::ObjectSubtype::Default,
            persist_id: 1, // Reassigned by the edit.
            unused: [0; 4],
        },
        menu_name: None,
        program_id: Some("Package".into()),
        clipboard_name: None,
        metafile: None,
    })
}

/// A deck without an `ExObjListContainer` (a new package and a Builder deck)
/// edits an empty list; the first committed object creates the container,
/// and the patch replays and inverts exactly.
#[test]
fn embedded_objects_are_added_to_a_deck_without_an_object_list() {
    for source in [Package::new().unwrap(), built()] {
        let source_bytes = source.to_bytes().unwrap();
        assert_eq!(source.presentation().unwrap().ole_objects().unwrap(), None);

        // Nothing staged, or an add that is removed again: an exact no-op.
        let edit = source.edit_objects().unwrap();
        assert!(edit.objects().unwrap().objects.is_empty());
        assert_eq!(edit.objects().unwrap().id_seed, 1);
        let no_op = edit.commit().unwrap();
        assert!(!no_op.diagnostics().changed());
        assert!(no_op.patch().is_noop());
        assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);
        let mut edit = source.edit_objects().unwrap();
        edit.add(embedded_object(1), object_storage(b"x")).unwrap();
        edit.remove(1).unwrap();
        assert!(!edit.commit().unwrap().diagnostics().changed());

        let storage = object_storage(b"first OLE2 storage");
        let mut edit = source.edit_objects().unwrap();
        let persist_id = edit.add(embedded_object(1), storage.clone()).unwrap();
        edit.add(embedded_object(2), object_storage(b"second"))
            .unwrap();
        let commit = edit.commit().unwrap();
        assert!(commit.diagnostics().changed());

        // The read view and a reopened package both see the new list.
        let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
        for package in [commit.package(), &reopened] {
            let presentation = package.presentation().unwrap();
            let objects = presentation.ole_objects().unwrap().unwrap();
            assert_eq!(objects.id_seed, 2);
            assert_eq!(object_ids(package), [1, 2]);
            assert_eq!(objects.objects[0].persist_id(), persist_id);
            assert_eq!(
                presentation.ole_storage(persist_id).unwrap(),
                Some(storage.clone())
            );
            assert_eq!(slide_texts(package), slide_texts(&source));
            let edit = package.edit_objects().unwrap();
            assert_eq!(edit.objects().unwrap(), &objects);
        }

        let replayed = source.apply_objects(commit.patch()).unwrap();
        assert_eq!(
            replayed.package().to_bytes().unwrap(),
            commit.package().to_bytes().unwrap()
        );
        let restored = commit
            .package()
            .apply_objects(&commit.patch().inverse())
            .unwrap();
        assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
        assert_eq!(
            restored
                .package()
                .presentation()
                .unwrap()
                .ole_objects()
                .unwrap(),
            None
        );
    }
}

/// `source` with a digital-signature storage added (`_xmlsignatures`), the
/// shape a signed deck has.
fn signed(source: &Package) -> Package {
    use litchi_cfb::{OleFile, OleWriter};
    use std::io::Cursor;

    let bytes = source.to_bytes().unwrap();
    let mut source_file =
        OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let mut writer = OleWriter::new();
    for stream_path in source_file.list_streams() {
        let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = source_file.open_stream(&refs).unwrap();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer
        .create_stream(&["_xmlsignatures", "origin.sigs"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    Package::from_bytes(&output.into_inner()).unwrap()
}

/// The signed-package rule every family shares: the edit opens, an exact
/// no-op commit keeps the signed bytes, and a changed commit is `Signed`.
macro_rules! assert_signed_rule {
    ($package:expr, $edit:expr, $change:expr, $signed:pat) => {{
        let package: &Package = $package;
        let bytes = package.to_bytes().unwrap();
        let no_op = $edit(package).commit().unwrap();
        assert!(!no_op.diagnostics().changed());
        assert!(no_op.patch().is_noop());
        assert_eq!(no_op.package().to_bytes().unwrap(), bytes);
        let mut edit = $edit(package);
        $change(&mut edit);
        assert!(matches!(edit.commit(), Err($signed)));
        assert_eq!(package.to_bytes().unwrap(), bytes);
    }};
}

#[test]
fn a_signed_deck_admits_object_edits_and_refuses_only_a_changed_commit() {
    let embedded =
        Package::from_bytes(&fixture("poi/test-data/slideshow/ppt_with_embeded.ppt")).unwrap();
    let source = signed(&embedded);
    let ids = object_ids(&source);
    // A reorder back to the source order is an exact no-op too.
    let mut edit = source.edit_objects().unwrap();
    edit.reorder(&ids).unwrap();
    assert_eq!(
        edit.commit().unwrap().package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
    assert_signed_rule!(
        &source,
        |package: &Package| package.edit_objects().unwrap(),
        |edit: &mut object::Edit| edit.reorder(&[6, 5, 4, 3, 2, 1]).unwrap(),
        object::Error::Signed
    );
    // A signed deck without an object list: adding the first object would
    // invalidate the signature.
    assert_signed_rule!(
        &signed(&built()),
        |package: &Package| package.edit_objects().unwrap(),
        |edit: &mut object::Edit| {
            edit.add(embedded_object(1), object_storage(b"x")).unwrap();
        },
        object::Error::Signed
    );
}

fn object_storage(data: &[u8]) -> litchi_ppt::embedded::storage::Storage {
    litchi_ppt::embedded::storage::Storage::uncompressed(
        litchi_ppt::embedded::storage::Kind::OleObject,
        data.to_vec(),
    )
    .unwrap()
}

#[test]
fn a_signed_deck_admits_external_media_edits_and_refuses_only_a_changed_commit() {
    let source = signed(&Package::from_bytes(&fixture("unified-ops/0294/video.ppt")).unwrap());
    assert_signed_rule!(
        &source,
        |package: &Package| package.edit_external_media().unwrap(),
        |edit: &mut external_media::Edit| {
            edit.set_path(1, Some("scoped.avi".into())).unwrap();
        },
        external_media::Error::Signed
    );
}

#[test]
fn a_signed_deck_admits_master_edits_and_refuses_only_a_changed_commit() {
    // Main and handout masters, a notes master with a MasterId-domain
    // slideIdRef, and a title master.
    for (source, kind) in [
        (mastered(), Kind::Main),
        (mastered(), Kind::Handout),
        (mastered(), Kind::Notes),
        (titled(), Kind::Title),
    ] {
        let source = signed(&source);
        let position = Position::new(0);
        assert_signed_rule!(
            &source,
            |package: &Package| package.edit_master_metadata(kind, position).unwrap(),
            |edit: &mut master_metadata::Edit| edit.set_name("Scoped master").unwrap(),
            master_metadata::Error::Signed
        );
        assert_signed_rule!(
            &source,
            |package: &Package| package.edit_master_layout(kind, position).unwrap(),
            |edit: &mut master_layout::Edit| {
                let last = edit.record().unwrap().children.last().unwrap().clone();
                edit.append(master_layout::Path::root(), last).unwrap();
            },
            master_layout::Error::Signed
        );
    }
}
