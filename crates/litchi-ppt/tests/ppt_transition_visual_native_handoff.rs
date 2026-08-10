#![allow(
    clippy::expect_used,
    reason = "test and explicitly invoked handoff assertions panic on failure by design"
)]

//! Exact public-API contract for an external same-format PPT resave.
//!
//! The mutation replaces only the canonical visual transition fields. Ignored
//! helpers stop at the filesystem boundary; an external harness owns any
//! native resave and evidence capture.

use std::path::{Path, PathBuf};

use litchi_core::patch::{BlobLimits, Patch, PatchLimits, Reversible};
use litchi_ppt::slide_order::{Error, Refusal};
use litchi_ppt::slide_order::{HistoryLimits, Position, SlideTransitionVisual, Snapshot};
use litchi_ppt::{TransitionDirection, TransitionSpeed, TransitionType};

const GENERATED_PATH_ENV: &str = "LITCHI_PPT_GENERATED_ARTIFACT";
const NATIVE_PATH_ENV: &str = "LITCHI_PPT_NATIVE_ARTIFACT";

fn source_fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow/45543.ppt")
}

fn source_position() -> Position {
    Position::new(0)
}

fn before_visual() -> SlideTransitionVisual {
    SlideTransitionVisual::new(
        TransitionType::Box,
        TransitionDirection::Out,
        TransitionSpeed::Slow,
    )
    .expect("fixture visual transition is exact")
}

fn after_visual() -> SlideTransitionVisual {
    SlideTransitionVisual::new(
        TransitionType::Cover,
        TransitionDirection::FromLeft,
        TransitionSpeed::Medium,
    )
    .expect("handoff visual transition is exact")
}

fn patch_limits() -> PatchLimits {
    PatchLimits::new(BlobLimits::new(0, 0, 0), 64 * 1024, 64, 8, 4_096, 32 * 1024)
}

fn generate(source: &Path) -> (Snapshot, litchi_ppt::slide_order::Patch) {
    let bytes = std::fs::read(source).expect("read genuine producer PPT fixture");
    let snapshot = Snapshot::from_bytes(bytes).expect("open genuine producer PPT fixture");
    assert_eq!(
        snapshot
            .slide_transition_visual(source_position())
            .expect("read fixture visual transition"),
        before_visual()
    );

    let mut transaction = snapshot.edit().expect("start root transaction");
    transaction
        .set_slide_transition_visual(source_position(), after_visual())
        .expect("stage fixed-width visual transition replacement");
    let commit = transaction.commit().expect("publish and fully reopen PPT");
    commit.into_parts()
}

fn assert_readback(path: &Path) {
    let bytes = std::fs::read(path).expect("read candidate PPT artifact");
    let snapshot = Snapshot::from_bytes(bytes).expect("fully reopen candidate PPT artifact");
    assert_eq!(
        snapshot
            .slide_transition_visual(source_position())
            .expect("read candidate visual transition"),
        after_visual()
    );
}

#[test]
fn genuine_visual_transition_reopens_reverses_merges_and_uses_history() {
    let source = Snapshot::from_bytes(
        std::fs::read(source_fixture()).expect("read genuine producer PPT fixture"),
    )
    .expect("open genuine producer PPT fixture");
    let (generated, root_patch) = generate(&source_fixture());
    assert_eq!(
        generated
            .slide_transition_visual(source_position())
            .expect("read generated visual transition"),
        after_visual()
    );
    let reopened = Snapshot::parse(generated.bytes()).expect("fully reopen generated PPT bytes");
    assert_eq!(
        reopened
            .slide_transition_visual(source_position())
            .expect("read reopened visual transition"),
        after_visual()
    );

    assert_eq!(
        root_patch.apply(&source).expect("apply exact patch"),
        generated
    );
    assert_eq!(
        root_patch
            .inverse()
            .apply(&generated)
            .expect("apply exact inverse"),
        source
    );

    let durable = root_patch
        .to_durable(patch_limits())
        .expect("create durable patch");
    let wire = durable
        .to_deterministic_json()
        .expect("serialize durable patch");
    let decoded = Patch::<Reversible>::from_deterministic_json(&wire, patch_limits())
        .expect("decode durable patch");
    let applied = source
        .apply_durable(&decoded)
        .expect("apply durable visual transition operation");
    assert_eq!(
        applied
            .slide_transition_visual(source_position())
            .expect("read durable result"),
        after_visual()
    );
    let restored = applied
        .apply_durable(&decoded.inverse())
        .expect("apply durable inverse");
    assert_eq!(
        restored
            .slide_transition_visual(source_position())
            .expect("read durable inverse result"),
        before_visual()
    );

    let mut history = source.history(HistoryLimits::new(4, 64 * 1024));
    history
        .record(generated.clone(), wire.len() as u64)
        .expect("record generated snapshot");
    assert!(history.undo());
    assert_eq!(history.current(), &source);
    assert!(history.redo());
    assert_eq!(history.current(), &generated);

    let competing_visual = SlideTransitionVisual::new(
        TransitionType::Fade,
        TransitionDirection::None,
        TransitionSpeed::Fast,
    )
    .expect("competing visual transition is exact");
    let mut competing_edit = source.edit().expect("start competing root transaction");
    competing_edit
        .set_slide_transition_visual(source_position(), competing_visual)
        .expect("stage competing visual transition");
    let competing_commit = competing_edit
        .commit()
        .expect("publish competing visual transition");
    assert!(
        source
            .plan_three_way(&root_patch, competing_commit.patch())
            .expect("plan visual-transition merge")
            .conflicts()
            .iter()
            .any(|conflict| conflict.target().ends_with("/transition-visual"))
    );
}

#[test]
fn visual_transition_refusals_are_typed_and_failure_atomic() {
    assert!(
        SlideTransitionVisual::new(
            TransitionType::Cover,
            TransitionDirection::None,
            TransitionSpeed::Medium,
        )
        .is_err()
    );

    let source = Snapshot::from_bytes(
        std::fs::read(
            PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/ppt/SampleShow.ppt"),
        )
        .expect("read genuine PPT without canonical visual owner"),
    )
    .expect("open genuine PPT without canonical visual owner");
    assert!(matches!(
        source.slide_transition_visual(source_position()),
        Err(Error::Refused(
            Refusal::UnsupportedSlideTransitionVisual { position }
        )) if position == source_position()
    ));
    let mut transaction = source.edit().expect("start root transaction");
    assert!(matches!(
        transaction.set_slide_transition_visual(source_position(), after_visual()),
        Err(Error::Refused(
            Refusal::UnsupportedSlideTransitionVisual { position }
        )) if position == source_position()
    ));
    assert!(transaction.slide_transition_visual_changes().is_empty());
    assert_eq!(
        transaction
            .commit()
            .expect("commit unchanged transaction")
            .snapshot(),
        &source
    );
}

#[test]
#[ignore = "writes the Litchi-generated artifact path requested by an external harness"]
fn generate_litchi_changed_artifact() {
    let output = std::env::var_os(GENERATED_PATH_ENV)
        .map(PathBuf::from)
        .expect("set LITCHI_PPT_GENERATED_ARTIFACT to an explicit output file");
    let (generated, _root_patch) = generate(&source_fixture());
    std::fs::write(output, generated.bytes()).expect("write Litchi-generated PPT artifact");
}

#[test]
#[ignore = "reads a native-resaved artifact supplied by an external harness"]
fn read_back_native_resaved_artifact() {
    let input = std::env::var_os(NATIVE_PATH_ENV)
        .map(PathBuf::from)
        .expect("set LITCHI_PPT_NATIVE_ARTIFACT to an explicit input file");
    assert_readback(&input);
}
