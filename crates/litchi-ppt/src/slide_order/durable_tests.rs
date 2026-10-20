#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "durable-patch and history tests use panic-driven assertions"
)]

//! Durable patches, history and three-way planning of the native slide-order
//! owner.
//!
//! These scenarios moved here from the integration tests when the native
//! owner became crate-private (record 0357); [`crate::Package::edit`] is the
//! public route and does not expose these surfaces yet.

use std::path::PathBuf;

use litchi_core::patch::{BlobLimits, Patch as DurablePatch, PatchLimits, Reversible};

use super::{HistoryLimits, Position, SlideTransitionVisual, Snapshot};
use crate::comments::{DateTime, NewComment};
use crate::text_edit::Target;
use crate::{Package, TransitionDirection, TransitionSpeed, TransitionType};

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

fn shape_texts(snapshot: &Snapshot) -> Vec<Vec<String>> {
    let package = Package::from_bytes(snapshot.bytes()).unwrap();
    let presentation = package.presentation().unwrap();
    presentation
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

fn notes(snapshot: &Snapshot) -> Vec<Option<String>> {
    let package = Package::from_bytes(snapshot.bytes()).unwrap();
    let presentation = package.presentation().unwrap();
    presentation
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

fn comment_count(snapshot: &Snapshot, slide: usize) -> usize {
    Package::from_bytes(snapshot.bytes())
        .unwrap()
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

#[test]
fn durable_insertion_replays_and_reverts() {
    let source = Snapshot::from_bytes(fixture("ole/ppt/SampleShow.ppt")).unwrap();
    let mut transaction = source.edit().unwrap();
    transaction.insert_slide(Position::new(1)).unwrap();
    let (after, patch) = transaction.commit().unwrap().into_parts();
    let durable = patch
        .to_durable(limits(BlobLimits::new(4, 1024 * 1024, 4 * 1024 * 1024)))
        .unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert_eq!(replayed.bytes(), after.bytes());
    let reverted = replayed.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(reverted.slide_count(), source.slide_count());
}

#[test]
fn durable_notes_patch_replays_and_reverts() {
    let source =
        Snapshot::from_bytes(fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt")).unwrap();
    let mut transaction = source.edit().unwrap();
    transaction
        .set_notes_text(Position::new(1), "Second page notes, durable")
        .unwrap();
    let (target, patch) = transaction.commit().unwrap().into_parts();
    let durable = patch.to_durable(limits(BlobLimits::new(0, 0, 0))).unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert_eq!(notes(&replayed), notes(&target));
    let reverted = replayed.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(notes(&reverted), notes(&source));
}

#[test]
fn durable_comment_patch_replays_and_reverts() {
    let source =
        Snapshot::from_bytes(fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt")).unwrap();
    let mut transaction = source.edit().unwrap();
    transaction
        .add_comment(
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
    let (target, patch) = transaction.commit().unwrap().into_parts();
    let durable = patch
        .to_durable(limits(BlobLimits::new(8, 4 << 20, 8 << 20)))
        .unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert_eq!(replayed.bytes(), target.bytes());
    let reverted = target.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(comment_count(&reverted, 1), 0);
}

#[test]
fn durable_outline_text_patch_replays_and_reverts() {
    let source = Snapshot::from_bytes(fixture("ole/ppt/SampleShow.ppt")).unwrap();
    let mut transaction = source.edit().unwrap();
    transaction
        .set_shape_text(
            Target::new(Position::new(1), Position::new(0)),
            "A durable retitle",
        )
        .unwrap();
    let (after, patch) = transaction.commit().unwrap().into_parts();
    let durable = patch.to_durable(limits(BlobLimits::new(0, 0, 0))).unwrap();
    let replayed = source.apply_durable(&durable).unwrap();
    assert_eq!(shape_texts(&replayed), shape_texts(&after));
    let reverted = replayed.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(shape_texts(&reverted), shape_texts(&source));
}

fn visual(
    kind: TransitionType,
    direction: TransitionDirection,
    speed: TransitionSpeed,
) -> SlideTransitionVisual {
    SlideTransitionVisual::new(kind, direction, speed).unwrap()
}

#[test]
fn genuine_visual_transition_reverses_merges_and_uses_history() {
    let source = Snapshot::from_bytes(fixture("poi/test-data/slideshow/45543.ppt")).unwrap();
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
    assert_eq!(source.slide_transition_visual(position).unwrap(), before);
    let mut transaction = source.edit().unwrap();
    transaction
        .set_slide_transition_visual(position, after)
        .unwrap();
    let (generated, root_patch) = transaction.commit().unwrap().into_parts();
    assert_eq!(generated.slide_transition_visual(position).unwrap(), after);

    assert_eq!(root_patch.apply(&source).unwrap(), generated);
    assert_eq!(root_patch.inverse().apply(&generated).unwrap(), source);

    let patch_limits =
        PatchLimits::new(BlobLimits::new(0, 0, 0), 64 * 1024, 64, 8, 4_096, 32 * 1024);
    let durable = root_patch.to_durable(patch_limits).unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded = DurablePatch::<Reversible>::from_deterministic_json(&wire, patch_limits).unwrap();
    let applied = source.apply_durable(&decoded).unwrap();
    assert_eq!(applied.slide_transition_visual(position).unwrap(), after);
    let restored = applied.apply_durable(&decoded.inverse()).unwrap();
    assert_eq!(restored.slide_transition_visual(position).unwrap(), before);

    let mut history = source.history(HistoryLimits::new(4, 64 * 1024));
    history
        .record(generated.clone(), wire.len() as u64)
        .unwrap();
    assert!(history.undo());
    assert_eq!(history.current(), &source);
    assert!(history.redo());
    assert_eq!(history.current(), &generated);

    let mut competing_edit = source.edit().unwrap();
    competing_edit
        .set_slide_transition_visual(
            position,
            visual(
                TransitionType::Fade,
                TransitionDirection::None,
                TransitionSpeed::Fast,
            ),
        )
        .unwrap();
    let competing_commit = competing_edit.commit().unwrap();
    assert!(
        source
            .plan_three_way(&root_patch, competing_commit.patch())
            .unwrap()
            .conflicts()
            .iter()
            .any(|conflict| conflict.target().ends_with("/transition-visual"))
    );
}
