#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::Package;
use litchi_rtf::edit::{Error, History, HistoryLimits, RunFormatPatch, RunFormatUpdate};

const SOURCE: &[u8] = br"{\rtf1\ansi One\par Two}";

const MIXED_SOURCE: &[u8] = br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}{\info{\title Keep title}}{\header Keep header}\f0\cf1\i\fs24\uldb\strike\striked\vanish\caps\scaps\outl\up6\expndtw24\charscalex150\kerning28 Plain{\b Bold}\par Tail}";

fn bold_patch(value: bool) -> RunFormatPatch {
    RunFormatPatch::new().with_bold(value)
}

fn update(paragraph: usize, run: usize, patch: RunFormatPatch) -> RunFormatUpdate {
    RunFormatUpdate::new(Position::new(paragraph), Position::new(run), patch)
}

#[test]
fn native_run_format_batch_preflights_late_failure_before_staging() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let valid = update(0, 0, bold_patch(true));
    let invalid = update(1, 99, bold_patch(true));

    let mut edit = source.edit();
    assert!(matches!(
        edit.patch_body_run_formats(&[valid, invalid]),
        Err(Error::RunOutOfRange {
            paragraph: 1,
            run: 99,
            ..
        })
    ));
    assert_eq!(edit.operation_count(), 0);

    edit.patch_body_run_formats(&[valid]).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(
        commit
            .package()
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .runs()
            .next()
            .unwrap()
            .format()
            .bold()
    );
}

#[test]
fn native_run_format_batch_shape_and_typed_bounds_are_checked() {
    let source = Package::from_bytes(SOURCE).unwrap();

    let mut empty = source.edit();
    assert!(matches!(
        empty.patch_body_run_formats(&[]),
        Err(Error::EmptyRunFormatBatch)
    ));
    assert_eq!(empty.operation_count(), 0);

    let mut empty_patch = source.edit();
    assert!(matches!(
        empty_patch.patch_run_format(Position::new(0), Position::new(0), RunFormatPatch::new()),
        Err(Error::EmptyRunFormatPatch {
            paragraph: 0,
            run: 0
        })
    ));
    assert_eq!(empty_patch.operation_count(), 0);

    let mut unordered = source.edit();
    assert!(matches!(
        unordered.patch_body_run_formats(&[
            update(1, 0, bold_patch(true)),
            update(0, 0, bold_patch(true)),
        ]),
        Err(Error::RunFormatBatchOutOfOrder {
            previous_paragraph: 1,
            previous_run: 0,
            incoming_paragraph: 0,
            incoming_run: 0,
        })
    ));
    assert_eq!(unordered.operation_count(), 0);

    let mut paragraph_oob = source.edit();
    assert!(matches!(
        paragraph_oob.patch_run_format(Position::new(99), Position::new(0), bold_patch(true)),
        Err(Error::ParagraphOutOfRange {
            position: 99,
            count: 2
        })
    ));
    assert_eq!(paragraph_oob.operation_count(), 0);

    let mut run_oob = source.edit();
    assert!(matches!(
        run_oob.patch_run_format(Position::new(0), Position::new(99), bold_patch(true)),
        Err(Error::RunOutOfRange {
            paragraph: 0,
            run: 99,
            ..
        })
    ));
    assert_eq!(run_oob.operation_count(), 0);
}

#[test]
fn native_run_format_omitted_facets_preserve_mixed_run_properties() {
    let source = Package::from_bytes(MIXED_SOURCE).unwrap();
    let before = source
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .next()
        .unwrap()
        .format();

    assert!(!before.bold());
    assert!(before.italic());
    assert_eq!(before.size().get(), 24);
    assert!(before.strike());
    assert!(before.double_strike());
    assert_eq!(
        before.baseline(),
        litchi_rtf::CharacterBaseline::RaisedHalfPoints(6)
    );
    assert_eq!(before.horizontal_scale_percent(), 150);
    assert_eq!(before.kerning_half_points(), 28);

    let mut edit = source.edit();
    edit.patch_run_format(Position::new(0), Position::new(0), bold_patch(true))
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit
        .package()
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .next()
        .unwrap()
        .format();

    assert!(after.bold());
    assert_eq!(after.italic(), before.italic());
    assert_eq!(after.underline(), before.underline());
    assert_eq!(after.size(), before.size());
    assert_eq!(after.strike(), before.strike());
    assert_eq!(after.double_strike(), before.double_strike());
    assert_eq!(after.hidden(), before.hidden());
    assert_eq!(after.all_caps(), before.all_caps());
    assert_eq!(after.small_caps(), before.small_caps());
    assert_eq!(after.outline(), before.outline());
    assert_eq!(after.baseline(), before.baseline());
    assert_eq!(after.expansion(), before.expansion());
    assert_eq!(
        after.horizontal_scale_percent(),
        before.horizontal_scale_percent()
    );
    assert_eq!(after.kerning_half_points(), before.kerning_half_points());
    assert!(
        commit
            .package()
            .to_bytes()
            .unwrap()
            .windows(b"Keep title".len())
            .any(|window| window == b"Keep title")
    );
    assert!(
        commit
            .package()
            .to_bytes()
            .unwrap()
            .windows(b"Keep header".len())
            .any(|window| window == b"Keep header")
    );
}

#[test]
fn native_run_format_noop_and_inverse_preserve_exact_source_bytes() {
    let source = Package::from_bytes(MIXED_SOURCE).unwrap();
    let source_bytes = source.to_bytes().unwrap();

    let mut noop_edit = source.edit();
    noop_edit
        .patch_run_format(Position::new(0), Position::new(0), bold_patch(false))
        .unwrap();
    let noop = noop_edit.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().to_bytes().unwrap(), source_bytes);

    let mut changed_edit = source.edit();
    changed_edit
        .patch_run_format(Position::new(0), Position::new(0), bold_patch(true))
        .unwrap();
    let changed = changed_edit.commit().unwrap();
    assert!(changed.diagnostics().changed());
    let restored = changed
        .package()
        .apply(&changed.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);

    let mut history = History::new(source.clone(), HistoryLimits::new(2, 1024 * 1024));
    let mut history_edit = history.current().edit();
    history_edit
        .patch_run_format(Position::new(0), Position::new(0), bold_patch(true))
        .unwrap();
    let history_commit = history.commit(history_edit).unwrap();
    let changed_bytes = history_commit.package().to_bytes().unwrap();
    assert!(history.undo());
    assert_eq!(history.current().to_bytes().unwrap(), source_bytes);
    assert!(history.redo());
    assert_eq!(history.current().to_bytes().unwrap(), changed_bytes);
}

#[test]
fn native_run_format_refuses_mixed_structural_work() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let mut edit = source.edit();
    edit.patch_run_format(Position::new(0), Position::new(0), bold_patch(true))
        .unwrap();

    assert!(matches!(
        edit.replace_body_text("Changed\nwith a new paragraph"),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation && refusal.reason() == "paragraph-structure and property changes cannot compose safely"
    ));
    assert_eq!(edit.operation_count(), 1);
}
