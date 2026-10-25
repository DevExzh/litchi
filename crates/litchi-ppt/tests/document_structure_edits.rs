#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "document-structure edit tests use panic-driven assertions"
)]

//! Master reordering, custom table-style placement and `fNonOutlineData`
//! through the primary edit (record 0357a).

use std::io::Cursor;
use std::path::PathBuf;

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::edit::RefusalKind;
use litchi_ppt::package::edit::{CustomTableStylesPlacement, Error, Position};
use litchi_ppt::slide_order::{self, Refusal};
use litchi_ppt::{DocumentStructure, OpenOptions, Package};

fn read(path: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(path),
    )
    .expect("PPT fixture")
}

/// Two main masters and a title master (Apache POI `slide_master.ppt`).
fn masters_deck() -> Package {
    Package::from_bytes(&read("poi/test-data/slideshow/slide_master.ppt")).unwrap()
}

/// Two masters and a custom table-style atom after the `EndDocumentAtom`.
fn styles_deck() -> Package {
    Package::from_bytes(&read("poi/test-data/slideshow/45543.ppt")).unwrap()
}

fn bytes(package: &Package) -> Vec<u8> {
    package.to_bytes().unwrap()
}

fn structure(package: &Package) -> DocumentStructure {
    package
        .presentation()
        .unwrap()
        .document_structure()
        .unwrap()
}

fn master_ids(package: &Package) -> Vec<u32> {
    structure(package)
        .masters()
        .iter()
        .map(|master| master.master_id())
        .collect()
}

fn main_master_count(package: &Package) -> usize {
    package
        .presentation()
        .unwrap()
        .masters()
        .unwrap()
        .main()
        .len()
}

fn slide_texts(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.text().unwrap().to_string())
        .collect()
}

fn assert_replays_and_inverts(source: &Package, commit: &litchi_ppt::package::edit::Commit) {
    assert!(commit.diagnostics().changed());
    assert_eq!(
        bytes(source.apply(commit.patch()).unwrap().package()),
        bytes(commit.package())
    );
    assert_eq!(
        bytes(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
        ),
        bytes(source)
    );
    assert!(matches!(
        commit.package().apply(commit.patch()),
        Err(Error::PatchConflict)
    ));
}

fn refusal_kind(error: &Error) -> Option<RefusalKind> {
    match error {
        Error::Native(slide_order::Error::Refused(refusal)) => Some(refusal.kind()),
        _ => None,
    }
}

#[test]
fn masters_reorder_and_keep_every_reference() {
    let source = masters_deck();
    assert_eq!(main_master_count(&source), 2);
    let ids = master_ids(&source);
    assert_eq!(ids.len(), 3);
    let texts = slide_texts(&source);

    let mut edit = source.edit().unwrap();
    edit.move_master(Position::new(0), Position::new(1))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(master_ids(commit.package()), [ids[1], ids[0], ids[2]]);
    // Slides keep their payloads (the publisher checks every persisted
    // slide record) and every master still resolves.
    assert_eq!(slide_texts(commit.package()), texts);
    assert_eq!(main_master_count(commit.package()), 2);
    assert_eq!(
        commit
            .package()
            .presentation()
            .unwrap()
            .masters()
            .unwrap()
            .masters()
            .count(),
        source
            .presentation()
            .unwrap()
            .masters()
            .unwrap()
            .masters()
            .count()
    );
    assert_replays_and_inverts(&source, &commit);

    // By id and by a whole permutation.
    let mut edit = source.edit().unwrap();
    edit.move_master_id(ids[2], Position::new(0)).unwrap();
    let by_id = edit.commit().unwrap();
    assert_eq!(master_ids(by_id.package()), [ids[2], ids[0], ids[1]]);
    let mut edit = source.edit().unwrap();
    edit.reorder_masters(&[Position::new(2), Position::new(0), Position::new(1)])
        .unwrap();
    assert_eq!(
        bytes(edit.commit().unwrap().package()),
        bytes(by_id.package())
    );
}

#[test]
fn master_reorder_noops_and_refusals() {
    let source = masters_deck();
    let mut edit = source.edit().unwrap();
    edit.reorder_masters(&[Position::new(0), Position::new(1), Position::new(2)])
        .unwrap();
    edit.move_master(Position::new(1), Position::new(1))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package()), bytes(&source));

    let mut edit = source.edit().unwrap();
    let error = edit
        .reorder_masters(&[Position::new(0), Position::new(0), Position::new(1)])
        .unwrap_err();
    assert_eq!(refusal_kind(&error), Some(RefusalKind::InvalidValue));
    let error = edit
        .move_master(Position::new(3), Position::new(0))
        .unwrap_err();
    assert_eq!(refusal_kind(&error), Some(RefusalKind::NotFound));
    let error = edit
        .move_master_id(0x7fff_fff0, Position::new(0))
        .unwrap_err();
    assert_eq!(refusal_kind(&error), Some(RefusalKind::NotFound));
    // A refused call leaves the edit unchanged.
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

/// The fixture's custom table-style atom, read by replacing it with itself.
fn styles_atom(source: &Package) -> Vec<u8> {
    let mut edit = source.edit().unwrap();
    let atom = {
        let mut probe = source.edit().unwrap();
        // Any valid atom works as the probe replacement.
        probe
            .replace_custom_table_styles(&[0, 0, 0x28, 0x04, 0, 0, 0, 0])
            .unwrap()
    };
    assert_eq!(edit.replace_custom_table_styles(&atom).unwrap(), atom);
    assert!(!edit.commit().unwrap().diagnostics().changed());
    atom
}

#[test]
fn custom_table_styles_replace_move_remove_and_insert() {
    let source = styles_deck();
    assert_eq!(
        structure(&source).custom_table_styles,
        Some(CustomTableStylesPlacement::AfterEndDocument)
    );
    let atom = styles_atom(&source);
    assert_eq!(&atom[2..4], &[0x28, 0x04]);

    // Replace in place with a different payload of another length.
    let mut replacement = atom.clone();
    replacement.extend_from_slice(b"pad!");
    let length = u32::try_from(replacement.len() - 8).unwrap();
    replacement[4..8].copy_from_slice(&length.to_le_bytes());
    let mut edit = source.edit().unwrap();
    assert_eq!(
        edit.replace_custom_table_styles(&replacement).unwrap(),
        atom
    );
    let replaced = edit.commit().unwrap();
    assert_eq!(styles_atom(replaced.package()), replacement);
    assert_eq!(slide_texts(replaced.package()), slide_texts(&source));
    assert_replays_and_inverts(&source, &replaced);

    // Move before the EndDocumentAtom.
    let mut edit = source.edit().unwrap();
    edit.move_custom_table_styles(CustomTableStylesPlacement::BeforeEndDocument)
        .unwrap();
    let moved = edit.commit().unwrap();
    assert_eq!(
        structure(moved.package()).custom_table_styles,
        Some(CustomTableStylesPlacement::BeforeEndDocument)
    );
    assert_eq!(styles_atom(moved.package()), atom);
    assert_replays_and_inverts(&source, &moved);

    // Remove, then insert it again in the same edit: a net no-op.
    let mut edit = source.edit().unwrap();
    assert_eq!(edit.remove_custom_table_styles().unwrap(), atom);
    let error = edit.remove_custom_table_styles().unwrap_err();
    assert_eq!(refusal_kind(&error), Some(RefusalKind::NotFound));
    edit.insert_custom_table_styles(&atom, CustomTableStylesPlacement::AfterEndDocument)
        .unwrap();
    let round = edit.commit().unwrap();
    assert!(!round.diagnostics().changed());
    assert_eq!(bytes(round.package()), bytes(&source));

    // Remove alone, then insert into the deck that has none.
    let mut edit = source.edit().unwrap();
    edit.remove_custom_table_styles().unwrap();
    let removed = edit.commit().unwrap();
    assert_eq!(structure(removed.package()).custom_table_styles, None);
    assert_replays_and_inverts(&source, &removed);
    let mut edit = removed.package().edit().unwrap();
    edit.insert_custom_table_styles(&atom, CustomTableStylesPlacement::BeforeEndDocument)
        .unwrap();
    let inserted = edit.commit().unwrap();
    assert_eq!(
        structure(inserted.package()).custom_table_styles,
        Some(CustomTableStylesPlacement::BeforeEndDocument)
    );
    assert_replays_and_inverts(removed.package(), &inserted);
}

#[test]
fn custom_table_style_noops_and_refusals() {
    let source = styles_deck();
    let atom = styles_atom(&source);
    let mut edit = source.edit().unwrap();
    edit.move_custom_table_styles(CustomTableStylesPlacement::AfterEndDocument)
        .unwrap();
    edit.replace_custom_table_styles(&atom).unwrap();
    let error = edit
        .insert_custom_table_styles(&atom, CustomTableStylesPlacement::BeforeEndDocument)
        .unwrap_err();
    assert_eq!(refusal_kind(&error), Some(RefusalKind::UnsupportedContent));
    // Not exactly one RoundTripCustomTableStyles12Atom.
    for bad in [
        &atom[..atom.len() - 1],
        &[0u8, 0, 0xEA, 0x03, 0, 0, 0, 0][..],
    ] {
        let error = edit.replace_custom_table_styles(bad).unwrap_err();
        assert_eq!(refusal_kind(&error), Some(RefusalKind::InvalidValue));
    }
    let mut trailing = atom.clone();
    trailing.push(0);
    let error = edit.replace_custom_table_styles(&trailing).unwrap_err();
    assert_eq!(refusal_kind(&error), Some(RefusalKind::InvalidValue));
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(bytes(commit.package()), bytes(&source));

    // A deck without the atom.
    let plain = masters_deck();
    let mut edit = plain.edit().unwrap();
    for error in [
        edit.remove_custom_table_styles().unwrap_err(),
        edit.replace_custom_table_styles(&atom).unwrap_err(),
        edit.move_custom_table_styles(CustomTableStylesPlacement::BeforeEndDocument)
            .unwrap_err(),
    ] {
        assert_eq!(refusal_kind(&error), Some(RefusalKind::NotFound));
    }
}

#[test]
fn non_outline_data_flag_round_trips() {
    let source = styles_deck();
    let flags = |package: &Package| structure(package).slides()[2].flags() & (1 << 2) != 0;
    let before = flags(&source);
    let mut edit = source.edit().unwrap();
    edit.set_slide_non_outline_data(Position::new(2), before)
        .unwrap();
    assert!(!edit.commit().unwrap().diagnostics().changed());

    let mut edit = source.edit().unwrap();
    edit.set_slide_non_outline_data(Position::new(2), !before)
        .unwrap();
    let error = edit
        .set_slide_non_outline_data(Position::new(99), true)
        .unwrap_err();
    assert_eq!(refusal_kind(&error), Some(RefusalKind::NotFound));
    let commit = edit.commit().unwrap();
    assert_eq!(flags(commit.package()), !before);
    assert_eq!(slide_texts(commit.package()), slide_texts(&source));
    assert_replays_and_inverts(&source, &commit);
}

#[test]
fn document_edits_compose_with_slide_moves_and_plan() {
    let source = styles_deck();
    let ids = master_ids(&source);
    let mut expected = slide_texts(&source);
    expected.swap(0, 1);
    let mut edit = source.edit().unwrap();
    edit.move_slide(Position::new(0), Position::new(1)).unwrap();
    edit.move_master(Position::new(1), Position::new(0))
        .unwrap();
    edit.move_custom_table_styles(CustomTableStylesPlacement::BeforeEndDocument)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(slide_texts(commit.package()), expected);
    assert_eq!(master_ids(commit.package()), [ids[1], ids[0]]);
    assert_replays_and_inverts(&source, &commit);

    // These edits have a durable form (record 0357b; the round trips are
    // in tests/durable_patches.rs).
    let durable = commit
        .patch()
        .to_durable(litchi_core::patch::PatchLimits::new(
            litchi_core::patch::BlobLimits::new(0, 0, 0),
            1 << 20,
            64,
            8,
            4_096,
            8 << 20,
        ))
        .unwrap();
    assert_eq!(durable.operations().len(), 3);
    let mut other = source.edit().unwrap();
    other
        .move_master(Position::new(0), Position::new(1))
        .unwrap();
    let other = other.commit().unwrap();
    // Two master reorders do not join: their order decides the result.
    assert!(matches!(
        commit.patch().join(other.patch()),
        Err(Error::Native(slide_order::Error::Refused(Refusal::JoinConflict { target })))
            if target == "document/master-list"
    ));
    let plan = source
        .plan_three_way(commit.patch(), other.patch())
        .unwrap();
    assert!(
        plan.conflicts()
            .iter()
            .any(|conflict| conflict.target() == "document/master-list")
    );
}

fn signed(bytes: &[u8]) -> Vec<u8> {
    let mut source =
        OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let names: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = source.open_stream(&names).unwrap();
        writer.create_stream(&names, &data).unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"opaque synthetic signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn signed_decks_admit_document_noops_and_refuse_changes() {
    let signed_bytes = signed(&read("poi/test-data/slideshow/45543.ppt"));
    let source = Package::from_bytes(&signed_bytes).unwrap();
    // The signed twin carries the same document stream.
    let atom = styles_atom(&styles_deck());
    let mut edit = source.edit().unwrap();
    edit.move_master(Position::new(0), Position::new(0))
        .unwrap();
    edit.replace_custom_table_styles(&atom).unwrap();
    edit.move_custom_table_styles(CustomTableStylesPlacement::AfterEndDocument)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(bytes(commit.package()), signed_bytes);

    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.move_master(Position::new(0), Position::new(1)),
        Err(Error::Signed)
    ));
    assert!(matches!(
        edit.remove_custom_table_styles(),
        Err(Error::Signed)
    ));
    let present = structure(&source).slides()[0].flags() & (1 << 2) != 0;
    assert!(matches!(
        edit.set_slide_non_outline_data(Position::new(0), !present),
        Err(Error::Signed)
    ));
    let commit = edit.commit().unwrap();
    assert_eq!(bytes(commit.package()), signed_bytes);
}

/// A media-path change makes the commit rebase the structural edits onto
/// the working document; the document edits are replayed with them.
#[test]
fn document_edits_replay_through_the_structural_rebase() {
    let source = Package::from_bytes(&read("unified-ops/0294/video.ppt")).unwrap();
    let flags = |package: &Package| structure(package).slides()[0].flags() & (1 << 2) != 0;
    let before = flags(&source);
    let mut edit = source.edit().unwrap();
    edit.set_slide_non_outline_data(Position::new(0), !before)
        .unwrap();
    edit.set_external_media_path(1, Some("rebased.mp4".into()))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(flags(commit.package()), !before);
    assert_eq!(
        slide_order::Snapshot::from_bytes_with_options(
            bytes(commit.package()),
            OpenOptions::new(),
        )
            .unwrap()
            .external_media_path(1)
            .unwrap()
            .as_deref(),
        Some("rebased.mp4")
    );
    assert_replays_and_inverts(&source, &commit);
}
