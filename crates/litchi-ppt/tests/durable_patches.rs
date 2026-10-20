#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "durable-patch, undo/redo, join and three-way tests use panic-driven assertions"
)]

//! Durable patches, undo/redo, joins and three-way planning of the primary
//! PPT edit (record 0357a).
//!
//! Record 0357 kept these surfaces on the crate-private native slide-order
//! owner; they are `package::edit::Patch::{to_durable, join}` and
//! `Package::{apply_durable, plan_three_way}` now.

use std::path::PathBuf;

use litchi_core::patch::{BlobLimits, Patch as DurablePatch, PatchLimits, Reversible};
use litchi_ppt::comments::{DateTime, NewComment};
use litchi_ppt::package::edit::{Error, Position, SlideTransitionVisual, Target};
use litchi_ppt::slide_order::{self, Refusal};
use litchi_ppt::{Package, TransitionDirection, TransitionSpeed, TransitionType};

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

    let mut edit = source.edit().unwrap();
    edit.set_slide_hidden(Position::new(3), true).unwrap();
    let hidden = edit.commit().unwrap().patch().clone();
    let error = first.join(&hidden).unwrap_err();
    assert!(matches!(
        &error,
        Error::Native(slide_order::Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));

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
    slide_order::Snapshot::parse(&bytes(package))
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
