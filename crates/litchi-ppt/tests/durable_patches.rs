#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "durable-patch, undo/redo, join and three-way tests use panic-driven assertions"
)]

//! Durable patches, undo/redo, joins and three-way planning of the primary
//! PPT edit (records 0357a, 0357b).
//!
//! Record 0357 kept these surfaces on the crate-private native slide-order
//! owner; they are `package::edit::Patch::{to_durable, join}` and
//! `Package::{apply_durable, plan_three_way}` now.

use std::path::PathBuf;

use litchi_core::patch::{BlobLimits, Patch as DurablePatch, PatchLimits, Reversible};
use litchi_ppt::comments::{DateTime, NewComment};
use litchi_ppt::package::edit::{
    CustomTableStylesPlacement, Error, Position, SlideTransitionVisual, Target,
};
use litchi_ppt::slide_order::{self, Refusal};
use litchi_ppt::{OpenOptions, Package, TransitionDirection, TransitionSpeed, TransitionType};

fn fixture(path: &str) -> Package {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    Package::from_bytes(&std::fs::read(path).expect("PPT fixture")).expect("open PPT fixture")
}

fn bytes(package: &Package) -> Vec<u8> {
    package.to_bytes().unwrap()
}

fn slide_count(package: &Package) -> usize {
    package.presentation().unwrap().slide_count()
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

fn shape_texts(package: &Package) -> Vec<Vec<String>> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            slide
                .shapes()
                .unwrap()
                .iter()
                .map(|shape| shape.text().unwrap())
                .collect()
        })
        .collect()
}

fn notes(package: &Package) -> Vec<Option<String>> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            slide
                .speaker_notes()
                .unwrap()
                .map(|notes| notes.text().unwrap().to_owned())
        })
        .collect()
}

fn comment_count(package: &Package, slide: usize) -> usize {
    package
        .presentation()
        .unwrap()
        .slide_at(slide)
        .unwrap()
        .unwrap()
        .comments()
        .unwrap()
        .len()
}

fn limits(blobs: BlobLimits) -> PatchLimits {
    PatchLimits::new(blobs, 1 << 20, 64, 8, 4_096, 8 << 20)
}

fn round_trip(durable: &DurablePatch<Reversible>, limits: PatchLimits) -> DurablePatch<Reversible> {
    let wire = durable.to_deterministic_json().unwrap();
    DurablePatch::<Reversible>::from_deterministic_json(&wire, limits).unwrap()
}

fn moved(source: &Package, from: usize, destination: usize) -> litchi_ppt::package::edit::Patch {
    let mut edit = source.edit().unwrap();
    edit.move_slide(Position::new(from), Position::new(destination))
        .unwrap();
    edit.commit().unwrap().patch().clone()
}

#[test]
fn durable_insertion_replays_and_reverts() {
    let source = fixture("ole/ppt/SampleShow.ppt");
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit
        .patch()
        .to_durable(limits(BlobLimits::new(4, 1024 * 1024, 4 * 1024 * 1024)))
        .unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert!(replayed.diagnostics().changed());
    assert_eq!(
        replayed.diagnostics().operation_count(),
        durable.operations().len()
    );
    assert_eq!(bytes(replayed.package()), bytes(commit.package()));
    let reverted = replayed
        .package()
        .apply_durable(&durable.inverse())
        .unwrap();
    assert_eq!(slide_count(reverted.package()), slide_count(&source));

    // The replay's own patch is exact: it replays and inverts by bytes.
    let exact = replayed.patch();
    assert_eq!(
        bytes(source.apply(exact).unwrap().package()),
        bytes(commit.package())
    );
    assert_eq!(
        bytes(
            replayed
                .package()
                .apply(&exact.inverse())
                .unwrap()
                .package()
        ),
        bytes(&source)
    );
    // It keeps no semantic operations, so it is not converted again.
    assert!(matches!(
        exact.to_durable(limits(BlobLimits::new(4, 1 << 20, 4 << 20))),
        Err(Error::Native(slide_order::Error::Refused(
            Refusal::UnsupportedPatch { .. }
        )))
    ));
}

#[test]
fn durable_notes_patch_replays_and_reverts() {
    let source = fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt");
    let mut edit = source.edit().unwrap();
    edit.set_notes_text(Position::new(1), "Second page notes, durable")
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit
        .patch()
        .to_durable(limits(BlobLimits::new(0, 0, 0)))
        .unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert_eq!(notes(replayed.package()), notes(commit.package()));
    let reverted = replayed
        .package()
        .apply_durable(&durable.inverse())
        .unwrap();
    assert_eq!(notes(reverted.package()), notes(&source));
}

#[test]
fn durable_comment_patch_replays_and_reverts() {
    let source = fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt");
    let mut edit = source.edit().unwrap();
    edit.add_comment(
        Position::new(1),
        &NewComment {
            author: "Reviewer".into(),
            initials: "R".into(),
            text: "Durable".into(),
            date: DateTime::parse("2026-10-17T09:30:00.250"),
            x: 120,
            y: 48,
        },
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit
        .patch()
        .to_durable(limits(BlobLimits::new(8, 4 << 20, 8 << 20)))
        .unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert_eq!(bytes(replayed.package()), bytes(commit.package()));
    let reverted = commit.package().apply_durable(&durable.inverse()).unwrap();
    assert_eq!(comment_count(reverted.package(), 1), 0);
}

#[test]
fn durable_outline_text_patch_replays_and_reverts() {
    let source = fixture("ole/ppt/SampleShow.ppt");
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(
        Target::new(Position::new(1), Position::new(0)),
        "A durable retitle",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit
        .patch()
        .to_durable(limits(BlobLimits::new(0, 0, 0)))
        .unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert_eq!(
        shape_texts(replayed.package()),
        shape_texts(commit.package())
    );
    let reverted = replayed
        .package()
        .apply_durable(&durable.inverse())
        .unwrap();
    assert_eq!(shape_texts(reverted.package()), shape_texts(&source));
}

#[test]
fn durable_patch_bounds_are_typed() {
    let source = fixture("ole/ppt/SampleShow.ppt");
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    // An insertion carries its slide payload as a blob; no blob fits.
    assert!(matches!(
        commit.patch().to_durable(limits(BlobLimits::new(0, 0, 0))),
        Err(Error::Durable(_))
    ));
}

#[test]
fn empty_durable_patch_publishes_the_source_unchanged() {
    let source = fixture("ole/ppt/SampleShow.ppt");
    let noop = source.edit().unwrap().commit().unwrap();
    let durable = noop
        .patch()
        .to_durable(limits(BlobLimits::new(0, 0, 0)))
        .unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert!(!replayed.diagnostics().changed());
    assert!(replayed.patch().is_noop());
    assert_eq!(bytes(replayed.package()), bytes(&source));
}

#[test]
fn undo_and_redo_are_inverse_and_apply() {
    let source = fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt");
    let original = slide_texts(&source);
    let mut edit = source.edit().unwrap();
    edit.move_slide(Position::new(0), Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        slide_texts(commit.package()),
        [original[1].clone(), original[0].clone()]
    );
    let undone = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(bytes(&undone), bytes(&source));
    let redone = undone.apply(commit.patch()).unwrap().into_package();
    assert_eq!(bytes(&redone), bytes(commit.package()));
    // A redo applies only to the exact package it was made from.
    assert!(matches!(
        commit.package().apply(commit.patch()),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn joined_move_patches_commute_and_replay_durably() {
    let source = fixture("poi/test-data/slideshow/45543.ppt");
    let first = moved(&source, 0, 1);
    let second = moved(&source, 4, 5);
    let joined = first.join(&second).unwrap();
    assert_eq!(joined.operation_count(), 2);
    let published = source.apply(&joined).unwrap().into_package();
    assert_eq!(slide_count(&published), slide_count(&source));
    let swapped = source
        .apply(&second.join(&first).unwrap())
        .unwrap()
        .into_package();
    assert_eq!(slide_texts(&published), slide_texts(&swapped));
    let mut expected = slide_texts(&source);
    expected.swap(0, 1);
    expected.swap(4, 5);
    assert_eq!(slide_texts(&published), expected);

    let durable = joined.to_durable(limits(BlobLimits::new(0, 0, 0))).unwrap();
    let replayed = source.apply_durable(&round_trip(&durable, limits(BlobLimits::new(0, 0, 0))));
    assert_eq!(bytes(replayed.unwrap().package()), bytes(&published));
    let restored = published.apply(&joined.inverse()).unwrap().into_package();
    assert_eq!(bytes(&restored), bytes(&source));
}

#[test]
fn joins_refuse_overlap_other_operations_and_other_sources() {
    let source = fixture("poi/test-data/slideshow/45543.ppt");
    let first = moved(&source, 0, 1);
    let overlapping = moved(&source, 1, 2);
    assert!(matches!(
        first.join(&overlapping),
        Err(Error::Native(slide_order::Error::Refused(
            Refusal::UnsupportedPatch { .. }
        )))
    ));

    // A slide insertion shifts every later position (record 0357b: other
    // operations join; insertions and removals do not).
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(3)).unwrap();
    let inserted = edit.commit().unwrap().patch().clone();
    let error = first.join(&inserted).unwrap_err();
    assert!(matches!(
        &error,
        Error::Native(slide_order::Error::Refused(refusal @ Refusal::UnsupportedPatch { .. }))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    // A slide-visibility change off the moved range joins.
    let mut edit = source.edit().unwrap();
    edit.set_slide_hidden(Position::new(3), true).unwrap();
    let hidden = edit.commit().unwrap().patch().clone();
    let joined = source.apply(&first.join(&hidden).unwrap()).unwrap();
    let native =
        slide_order::Snapshot::from_bytes_with_options(bytes(joined.package()), OpenOptions::new())
            .unwrap();
    assert!(native.slide_hidden(Position::new(3)).unwrap());

    let other = fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt");
    assert!(matches!(
        first.join(&moved(&other, 0, 1)),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn three_way_plans_list_common_targets() {
    let source = fixture("poi/test-data/slideshow/45543.ppt");
    let first = moved(&source, 0, 1);
    let overlapping = moved(&source, 1, 2);
    let disjoint = moved(&source, 4, 5);
    let plan = source.plan_three_way(&first, &disjoint).unwrap();
    assert!(plan.is_clean());
    let plan = source.plan_three_way(&first, &overlapping).unwrap();
    assert!(!plan.is_clean());
    assert!(
        plan.conflicts()
            .iter()
            .any(|conflict| conflict.target() == "slide-list/position:1")
    );
    let other = fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt");
    assert!(matches!(
        other.plan_three_way(&first, &disjoint),
        Err(Error::PatchConflict)
    ));
}

fn visual(
    kind: TransitionType,
    direction: TransitionDirection,
    speed: TransitionSpeed,
) -> SlideTransitionVisual {
    SlideTransitionVisual::new(kind, direction, speed).unwrap()
}

fn transition_visual(package: &Package, position: Position) -> SlideTransitionVisual {
    slide_order::Snapshot::from_bytes_with_options(bytes(package), OpenOptions::new())
        .unwrap()
        .slide_transition_visual(position)
        .unwrap()
}

/// The native-handoff visual-transition scenario of record 0357: exact
/// replay and inverse, a durable round trip, undo/redo and a three-way
/// conflict, all through the public route.
#[test]
fn genuine_visual_transition_reverses_merges_and_undoes() {
    let source = fixture("poi/test-data/slideshow/45543.ppt");
    let position = Position::new(0);
    let before = visual(
        TransitionType::Box,
        TransitionDirection::Out,
        TransitionSpeed::Slow,
    );
    let after = visual(
        TransitionType::Cover,
        TransitionDirection::FromLeft,
        TransitionSpeed::Medium,
    );
    assert_eq!(transition_visual(&source, position), before);
    let mut edit = source.edit().unwrap();
    edit.set_slide_transition_visual(position, after).unwrap();
    let commit = edit.commit().unwrap();
    let generated = commit.package();
    assert_eq!(transition_visual(generated, position), after);

    assert_eq!(
        bytes(source.apply(commit.patch()).unwrap().package()),
        bytes(generated)
    );
    let undone = generated
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(bytes(&undone), bytes(&source));
    assert_eq!(
        bytes(undone.apply(commit.patch()).unwrap().package()),
        bytes(generated)
    );

    let patch_limits =
        PatchLimits::new(BlobLimits::new(0, 0, 0), 64 * 1024, 64, 8, 4_096, 32 * 1024);
    let durable = commit.patch().to_durable(patch_limits).unwrap();
    let decoded = round_trip(&durable, patch_limits);
    let applied = source.apply_durable(&decoded).unwrap().into_package();
    assert_eq!(transition_visual(&applied, position), after);
    let restored = applied
        .apply_durable(&decoded.inverse())
        .unwrap()
        .into_package();
    assert_eq!(transition_visual(&restored, position), before);

    let mut competing = source.edit().unwrap();
    competing
        .set_slide_transition_visual(
            position,
            visual(
                TransitionType::Fade,
                TransitionDirection::None,
                TransitionSpeed::Fast,
            ),
        )
        .unwrap();
    let competing = competing.commit().unwrap();
    assert!(
        source
            .plan_three_way(commit.patch(), competing.patch())
            .unwrap()
            .conflicts()
            .iter()
            .any(|conflict| conflict.target().ends_with("/transition-visual"))
    );
}

// Record 0357b: the document-structure verbs of record 0357a (master order,
// custom table styles, `fNonOutlineData`) have durable operations and join.

fn masters_deck() -> Package {
    fixture("poi/test-data/slideshow/slide_master.ppt")
}

/// Two masters, a custom table-style atom after the `EndDocumentAtom`, and
/// enough slides for disjoint moves.
fn styles_deck() -> Package {
    fixture("poi/test-data/slideshow/45543.ppt")
}

fn structure(package: &Package) -> litchi_ppt::DocumentStructure {
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

fn non_outline_data(package: &Package, position: usize) -> bool {
    structure(package).slides()[position].flags() & (1 << 2) != 0
}

/// The custom table-style atom's exact bytes (an equal replacement is a
/// no-op that returns them).
fn styles_atom(package: &Package) -> Option<Vec<u8>> {
    structure(package).custom_table_styles?;
    let mut edit = package.edit().unwrap();
    let mut probe = package.edit().unwrap();
    let atom = probe
        .replace_custom_table_styles(&[0, 0, 0x28, 0x04, 0, 0, 0, 0])
        .unwrap();
    assert_eq!(edit.replace_custom_table_styles(&atom).unwrap(), atom);
    Some(atom)
}

/// Everything the document-structure verbs change, plus the slide texts.
fn projection(package: &Package) -> (Vec<u32>, Option<Vec<u8>>, Vec<bool>, Vec<String>) {
    let flags = (0..slide_count(package))
        .map(|position| non_outline_data(package, position))
        .collect();
    (
        master_ids(package),
        styles_atom(package),
        flags,
        slide_texts(package),
    )
}

fn blob_limits() -> PatchLimits {
    limits(BlobLimits::new(4, 1 << 20, 4 << 20))
}

/// `commit`'s durable patch, through deterministic JSON, replays on
/// `source` to the commit's bytes, its durable inverse restores the
/// source's projection, and it refuses another package.
fn assert_durable_round_trip(source: &Package, commit: &litchi_ppt::package::edit::Commit) {
    assert!(commit.diagnostics().changed());
    let durable = round_trip(
        &commit.patch().to_durable(blob_limits()).unwrap(),
        blob_limits(),
    );
    let replayed = source.apply_durable(&durable).unwrap();
    assert!(replayed.diagnostics().changed());
    assert_eq!(bytes(replayed.package()), bytes(commit.package()));
    let reverted = replayed
        .package()
        .apply_durable(&durable.inverse())
        .unwrap();
    assert_eq!(projection(reverted.package()), projection(source));
    // The operations bind the exact source artifact.
    assert!(commit.package().apply_durable(&durable).is_err());
}

#[test]
fn durable_master_reorders_replay_and_revert() {
    let source = masters_deck();
    let ids = master_ids(&source);
    assert_eq!(ids.len(), 3);
    let mut edit = source.edit().unwrap();
    edit.reorder_masters(&[Position::new(2), Position::new(0), Position::new(1)])
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(master_ids(commit.package()), [ids[2], ids[0], ids[1]]);
    // Master order needs no blob.
    let durable = commit
        .patch()
        .to_durable(limits(BlobLimits::new(0, 0, 0)))
        .unwrap();
    let operation = &durable.operations()[0];
    assert_eq!(operation.op, "document-structure.master-order.set");
    assert_eq!(operation.target, "document");
    assert_eq!(operation.value, serde_json::json!([2, 0, 1]));
    assert_eq!(
        operation.preconditions.keys().collect::<Vec<_>>(),
        ["artifact_sha256", "order_sha256"]
    );
    assert_durable_round_trip(&source, &commit);
}

#[test]
fn durable_custom_table_style_edits_replay_and_revert() {
    let source = styles_deck();
    let atom = styles_atom(&source).unwrap();

    let mut replacement = atom.clone();
    replacement.extend_from_slice(b"pad!");
    let length = u32::try_from(replacement.len() - 8).unwrap();
    replacement[4..8].copy_from_slice(&length.to_le_bytes());
    let mut edit = source.edit().unwrap();
    edit.replace_custom_table_styles(&replacement).unwrap();
    let replaced = edit.commit().unwrap();
    assert_eq!(styles_atom(replaced.package()).unwrap(), replacement);
    assert_durable_round_trip(&source, &replaced);
    // An atom travels as a blob; no blob fits.
    assert!(matches!(
        replaced
            .patch()
            .to_durable(limits(BlobLimits::new(0, 0, 0))),
        Err(Error::Durable(_))
    ));

    let mut edit = source.edit().unwrap();
    edit.move_custom_table_styles(CustomTableStylesPlacement::BeforeEndDocument)
        .unwrap();
    let moved_styles = edit.commit().unwrap();
    assert_durable_round_trip(&source, &moved_styles);

    let mut edit = source.edit().unwrap();
    edit.remove_custom_table_styles().unwrap();
    let removed = edit.commit().unwrap();
    assert_eq!(structure(removed.package()).custom_table_styles, None);
    assert_durable_round_trip(&source, &removed);

    let mut edit = removed.package().edit().unwrap();
    edit.insert_custom_table_styles(&atom, CustomTableStylesPlacement::BeforeEndDocument)
        .unwrap();
    let inserted = edit.commit().unwrap();
    assert_eq!(
        structure(inserted.package()).custom_table_styles,
        Some(CustomTableStylesPlacement::BeforeEndDocument)
    );
    assert_durable_round_trip(removed.package(), &inserted);
}

#[test]
fn durable_non_outline_data_and_mixed_structure_replay_and_revert() {
    let source = styles_deck();
    let before = non_outline_data(&source, 3);
    let mut edit = source.edit().unwrap();
    edit.set_slide_non_outline_data(Position::new(3), !before)
        .unwrap();
    let flagged = edit.commit().unwrap();
    assert_eq!(non_outline_data(flagged.package(), 3), !before);
    assert_durable_round_trip(&source, &flagged);

    // A slide move, a master reorder, a style move and a flag in one
    // structural segment, each bound to the slide order it runs at.
    let ids = master_ids(&source);
    let mut edit = source.edit().unwrap();
    edit.move_slide(Position::new(0), Position::new(1)).unwrap();
    edit.move_master(Position::new(1), Position::new(0))
        .unwrap();
    edit.move_custom_table_styles(CustomTableStylesPlacement::BeforeEndDocument)
        .unwrap();
    edit.set_slide_non_outline_data(Position::new(0), !non_outline_data(&source, 1))
        .unwrap();
    let mixed = edit.commit().unwrap();
    assert_eq!(master_ids(mixed.package()), [ids[1], ids[0]]);
    let durable = mixed.patch().to_durable(blob_limits()).unwrap();
    let ops: Vec<_> = durable
        .operations()
        .iter()
        .map(|operation| operation.op.as_str())
        .collect();
    assert_eq!(
        ops,
        [
            "slide-order.move",
            "document-structure.master-order.set",
            "document-structure.custom-table-styles.move",
            "document-structure.slide-non-outline-data.set",
        ]
    );
    assert_durable_round_trip(&source, &mixed);
}

/// The first non-empty shape text at or after slide `from`.
fn text_target(package: &Package, from: usize) -> Target {
    let texts = shape_texts(package);
    let (slide, shape) = texts
        .iter()
        .enumerate()
        .skip(from)
        .find_map(|(slide, shapes)| {
            shapes
                .iter()
                .position(|text| !text.trim().is_empty())
                .map(|shape| (slide, shape))
        })
        .expect("a slide with shape text");
    Target::new(Position::new(slide), Position::new(shape))
}

fn edited(
    source: &Package,
    stage: impl FnOnce(&mut litchi_ppt::package::edit::Edit),
) -> litchi_ppt::package::edit::Patch {
    let mut edit = source.edit().unwrap();
    stage(&mut edit);
    edit.commit().unwrap().patch().clone()
}

fn join_refusal(error: &Error) -> &Refusal {
    match error {
        Error::Native(slide_order::Error::Refused(refusal)) => refusal,
        other => panic!("expected a join refusal, got {other:?}"),
    }
}

#[test]
fn master_reorders_join_slide_moves_and_text_sets() {
    let source = styles_deck();
    let ids = master_ids(&source);
    let target = text_target(&source, 3);
    let masters = edited(&source, |edit| {
        edit.move_master(Position::new(1), Position::new(0))
            .unwrap();
    });
    let slides = moved(&source, 0, 1);
    let text = edited(&source, |edit| {
        edit.set_shape_text(target, "Joined text").unwrap();
    });

    let joined = masters.join(&slides).unwrap().join(&text).unwrap();
    assert_eq!(joined.operation_count(), 3);
    let published = source.apply(&joined).unwrap().into_package();
    assert_eq!(master_ids(&published), [ids[1], ids[0]]);
    let mut expected = shape_texts(&source);
    expected[target.slide().get()][target.shape().get()] = "Joined text".into();
    expected.swap(0, 1);
    assert_eq!(shape_texts(&published), expected);

    // Joins commute in effect.
    let swapped = text.join(&slides).unwrap().join(&masters).unwrap();
    let swapped = source.apply(&swapped).unwrap().into_package();
    assert_eq!(master_ids(&swapped), master_ids(&published));
    assert_eq!(shape_texts(&swapped), shape_texts(&published));

    // The joined patch inverts exactly and has a durable form.
    let restored = published.apply(&joined.inverse()).unwrap().into_package();
    assert_eq!(bytes(&restored), bytes(&source));
    let durable = round_trip(&joined.to_durable(blob_limits()).unwrap(), blob_limits());
    let replayed = source.apply_durable(&durable).unwrap().into_package();
    assert_eq!(master_ids(&replayed), master_ids(&published));
    assert_eq!(shape_texts(&replayed), shape_texts(&published));
    let reverted = replayed
        .apply_durable(&durable.inverse())
        .unwrap()
        .into_package();
    assert_eq!(projection(&reverted), projection(&source));
}

#[test]
fn document_edits_join_each_other_and_flags_off_the_moved_range() {
    let source = styles_deck();
    let flag = non_outline_data(&source, 4);
    let styles = edited(&source, |edit| {
        edit.move_custom_table_styles(CustomTableStylesPlacement::BeforeEndDocument)
            .unwrap();
    });
    let masters = edited(&source, |edit| {
        edit.move_master(Position::new(0), Position::new(1))
            .unwrap();
    });
    let flagged = edited(&source, |edit| {
        edit.set_slide_non_outline_data(Position::new(4), !flag)
            .unwrap();
    });
    let slides = moved(&source, 0, 1);
    let joined = styles
        .join(&masters)
        .unwrap()
        .join(&flagged)
        .unwrap()
        .join(&slides)
        .unwrap();
    let published = source.apply(&joined).unwrap().into_package();
    assert_eq!(
        structure(&published).custom_table_styles,
        Some(CustomTableStylesPlacement::BeforeEndDocument)
    );
    let ids = master_ids(&source);
    assert_eq!(master_ids(&published), [ids[1], ids[0]]);
    assert_eq!(non_outline_data(&published, 4), !flag);
    let mut texts = slide_texts(&source);
    texts.swap(0, 1);
    assert_eq!(slide_texts(&published), texts);
    assert_eq!(
        bytes(&published.apply(&joined.inverse()).unwrap().into_package()),
        bytes(&source)
    );
}

#[test]
fn conflicting_joins_are_typed_refusals() {
    let source = styles_deck();
    let conflict = |left: &litchi_ppt::package::edit::Patch,
                    right: &litchi_ppt::package::edit::Patch|
     -> String {
        let error = left.join(right).unwrap_err();
        let refusal = join_refusal(&error);
        assert_eq!(
            refusal.kind(),
            litchi_core::edit::RefusalKind::UnsupportedOperation
        );
        match refusal {
            Refusal::JoinConflict { target } => target.clone(),
            other => panic!("expected JoinConflict, got {other:?}"),
        }
    };

    // Two master reorders: their order would decide the master list.
    let first = edited(&source, |edit| {
        edit.move_master(Position::new(0), Position::new(1))
            .unwrap();
    });
    let second = edited(&source, |edit| {
        edit.reorder_masters(&[Position::new(1), Position::new(0)])
            .unwrap();
    });
    assert_eq!(conflict(&first, &second), "document/master-list");

    // Two custom table-style edits.
    let moved_styles = edited(&source, |edit| {
        edit.move_custom_table_styles(CustomTableStylesPlacement::BeforeEndDocument)
            .unwrap();
    });
    let removed_styles = edited(&source, |edit| {
        edit.remove_custom_table_styles().unwrap();
    });
    assert_eq!(
        conflict(&moved_styles, &removed_styles),
        "document/custom-table-styles"
    );

    // The same slide's flag, and the same shape's text.
    let flag = non_outline_data(&source, 4);
    let flagged = edited(&source, |edit| {
        edit.set_slide_non_outline_data(Position::new(4), !flag)
            .unwrap();
    });
    let flagged_too = flagged.clone();
    assert_eq!(
        conflict(&flagged, &flagged_too),
        "slide-list/position:4/non-outline-data"
    );
    let target = text_target(&source, 3);
    let text = edited(&source, |edit| edit.set_shape_text(target, "One").unwrap());
    let other_text = edited(&source, |edit| edit.set_shape_text(target, "Two").unwrap());
    assert!(conflict(&text, &other_text).ends_with("/text"));

    // A flag on a slide the other patch moves reads another slide there.
    let moved_slides = moved(&source, 3, 5);
    let error = flagged.join(&moved_slides).unwrap_err();
    assert!(matches!(
        join_refusal(&error),
        Refusal::UnsupportedPatch { .. }
    ));
    let error = moved_slides.join(&flagged).unwrap_err();
    assert!(matches!(
        join_refusal(&error),
        Refusal::UnsupportedPatch { .. }
    ));
}
