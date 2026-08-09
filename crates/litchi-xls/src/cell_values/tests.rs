#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test fixtures and assertions intentionally fail fast"
)]

use super::*;
use litchi_cfb::{OleFile, OleWriter};

fn package() -> Vec<u8> {
    let mut writer = crate::Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 3, 2, 4.5).unwrap();
    writer.write_number(sheet, 4, 0, 1.0).unwrap();
    writer.write_number(sheet, 4, 1, 2.0).unwrap();
    writer.write_string(sheet, 6, 0, "alpha").unwrap();
    writer.write_string(sheet, 6, 1, "beta").unwrap();
    writer.write_string(sheet, 6, 2, "alpha").unwrap();
    writer.write_boolean(sheet, 7, 0, true).unwrap();
    writer.write_formula(sheet, 8, 0, "1+1").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package =
        PackageEditor::open(output.into_inner(), Targets::default(), Limits::default()).unwrap();
    package
        .add_stream(vec!["Opaque".to_string()], b"untouched".to_vec())
        .unwrap();
    package.finish().unwrap()
}

fn signed_package() -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(package())).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    writer
        .create_stream(&["DigitalSignature"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn stream(bytes: &[u8], name: &str) -> Vec<u8> {
    OleFile::open(Cursor::new(bytes))
        .unwrap()
        .open_stream(&[name])
        .unwrap()
}

#[test]
fn edits_only_one_number_field_and_round_trips_patch() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let worksheet = source.worksheet("Sheet1".into()).unwrap().unwrap();
    assert_eq!(worksheet.position(), 0);
    assert_eq!(worksheet.number(reference).unwrap().unwrap().value(), 4.5);

    let before_workbook = source.workbook_stream().to_vec();
    let mut edit = source.edit();
    edit.set_number(0usize.into(), reference, 9.25).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_number_fields(), 1);
    assert_eq!(commit.diagnostics().touched_streams(), 1);
    let after_workbook = commit.snapshot().workbook_stream();
    assert_eq!(before_workbook.len(), after_workbook.len());
    assert_eq!(
        before_workbook
            .iter()
            .zip(after_workbook)
            .filter(|(left, right)| left != right)
            .count(),
        2
    );
    assert_eq!(stream(source.bytes(), "Opaque"), b"untouched");
    assert_eq!(stream(commit.snapshot().bytes(), "Opaque"), b"untouched");
    assert_eq!(
        source
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .number(reference)
            .unwrap()
            .unwrap()
            .value(),
        4.5
    );

    let applied = commit.patch().apply(&source).unwrap();
    let value = applied
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .number(reference)
        .unwrap()
        .unwrap()
        .value();
    assert_eq!(value, 9.25);
    assert!(commit.patch().apply(&applied).is_err());
    assert_eq!(
        commit.patch().inverse().apply(&applied).unwrap().bytes(),
        source.bytes()
    );
}

#[test]
fn rejected_and_noop_edits_are_failure_atomic() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let mut edit = source.edit();
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::NAN)
            .is_err()
    );
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::INFINITY)
            .is_err()
    );
    assert!(edit.set_number("Sheet1".into(), reference, -0.0).is_err());
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::MIN_POSITIVE / 2.0)
            .is_err()
    );
    assert!(
        edit.set_number("Sheet1".into(), Reference::new(9, 9).unwrap(), 1.0)
            .is_err()
    );
    edit.set_number("sHeEt1".into(), reference, 0.0).unwrap();
    edit.set_number("Sheet1".into(), reference, 4.5).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
    assert_eq!(commit.diagnostics(), Diagnostics::default());
}

#[test]
fn references_enforce_the_biff8_grid() {
    let reference = Reference::new(u32::from(u16::MAX), u32::from(u8::MAX)).unwrap();
    assert_eq!(reference.row(), u16::MAX);
    assert_eq!(reference.column(), u8::MAX);
    assert!(Reference::new(u32::from(u16::MAX) + 1, 0).is_err());
    assert!(Reference::new(0, u32::from(u8::MAX) + 1).is_err());
}

#[test]
fn signed_packages_are_refused_before_editing() {
    assert!(Snapshot::from_bytes(signed_package()).is_err());
}

#[test]
fn edits_boolean_text_and_formula_cache() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let sheet = source.worksheet("Sheet1".into()).unwrap().unwrap();
    assert_eq!(
        sheet
            .cell(Reference::new(4, 0).unwrap())
            .unwrap()
            .unwrap()
            .storage(),
        Storage::Number
    );
    assert_eq!(
        sheet
            .cell(Reference::new(6, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text("alpha".to_string())
    );
    assert_eq!(
        sheet
            .cell(Reference::new(8, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::Empty)
    );

    let mut edit = source.edit();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(4, 0).unwrap(),
        Value::Number(11.0),
    )
    .unwrap();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(7, 0).unwrap(),
        Value::Boolean(false),
    )
    .unwrap();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(6, 0).unwrap(),
        Value::Text("beta".to_string()),
    )
    .unwrap();
    edit.set_value(
        "Sheet1".into(),
        Reference::new(8, 0).unwrap(),
        Value::FormulaCache(FormulaCache::Error(CellError::new(0x07).unwrap())),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_cells(), 4);
    let sheet = commit
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap();
    assert_eq!(
        sheet
            .cell(Reference::new(4, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Number(11.0)
    );
    assert_eq!(
        sheet
            .cell(Reference::new(7, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Boolean(false)
    );
    assert_eq!(
        sheet
            .cell(Reference::new(6, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::Text("beta".to_string())
    );
    assert_eq!(
        sheet
            .cell(Reference::new(8, 0).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        &Value::FormulaCache(FormulaCache::Error(CellError::new(0x07).unwrap()))
    );
}

#[test]
fn converts_fixed_width_labelsst_and_rk_with_sst_accounting() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(6, 0).unwrap();
    let mut to_number = source.edit();
    to_number
        .set_value("Sheet1".into(), reference, Value::Number(42.0))
        .unwrap();
    let number = to_number.commit().unwrap();
    let cell = number
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .cell(reference)
        .unwrap()
        .unwrap();
    assert_eq!(cell.storage(), Storage::Rk);
    assert_eq!(cell.value(), &Value::Number(42.0));

    let mut to_text = number.snapshot().edit();
    to_text
        .set_value("Sheet1".into(), reference, Value::Text("beta".to_string()))
        .unwrap();
    let text = to_text.commit().unwrap();
    let cell = text
        .snapshot()
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .cell(reference)
        .unwrap()
        .unwrap();
    assert_eq!(cell.storage(), Storage::LabelSst);
    assert_eq!(cell.value(), &Value::Text("beta".to_string()));
}

#[test]
fn durable_semantic_patch_round_trips_and_checks_preconditions() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(7, 0).unwrap();
    let mut edit = source.transaction();
    edit.set_value("Sheet1".into(), reference, Value::Boolean(false))
        .unwrap();
    let commit = edit.commit().unwrap();
    let json = commit.patch().semantic().to_deterministic_json().unwrap();
    let parsed = SemanticPatch::from_deterministic_json(&json).unwrap();
    assert_eq!(parsed.to_deterministic_json().unwrap(), json);
    let replay = parsed.apply(&source).unwrap();
    assert_eq!(replay.snapshot().bytes(), commit.snapshot().bytes());
    assert!(parsed.apply(commit.snapshot()).is_err());
    let restored = parsed.inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored
            .snapshot()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .cell(reference)
            .unwrap()
            .unwrap()
            .value(),
        &Value::Boolean(true)
    );
}

#[test]
fn joins_disjoint_work_and_reports_cell_conflicts() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let mut left = source.transaction();
    left.set_value(
        "Sheet1".into(),
        Reference::new(7, 0).unwrap(),
        Value::Boolean(false),
    )
    .unwrap();
    let mut right = source.transaction();
    right
        .set_value(
            "Sheet1".into(),
            Reference::new(6, 0).unwrap(),
            Value::Text("beta".to_string()),
        )
        .unwrap();
    left.join(right).unwrap();
    assert_eq!(left.commit().unwrap().diagnostics().changed_cells(), 2);

    let mut left = source.transaction();
    left.set_value(
        "Sheet1".into(),
        Reference::new(7, 0).unwrap(),
        Value::Boolean(false),
    )
    .unwrap();
    let mut right = source.transaction();
    right
        .set_value(
            "Sheet1".into(),
            Reference::new(7, 0).unwrap(),
            Value::Error(CellError::new(0x07).unwrap()),
        )
        .unwrap();
    let error = left.join(right).unwrap_err();
    let JoinError::Conflicts(conflicts) = error else {
        panic!("expected structured cell conflict")
    };
    assert_eq!(conflicts.len(), 1);
    assert_eq!(conflicts[0].reference(), Reference::new(7, 0).unwrap());
    assert_eq!(left.commit().unwrap().diagnostics().changed_cells(), 1);
}

#[test]
fn bounded_history_undoes_and_redoes_immutable_snapshots() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let mut history = source.history(HistoryLimits::new(2, u64::MAX));
    let mut edit = source.transaction();
    edit.set_number("Sheet1".into(), Reference::new(3, 2).unwrap(), 99.0)
        .unwrap();
    edit.commit().unwrap().record_in(&mut history).unwrap();
    assert!(history.can_undo());
    assert!(history.undo());
    assert_eq!(history.current().bytes(), source.bytes());
    assert!(history.redo());
    assert_eq!(
        history
            .current()
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .number(Reference::new(3, 2).unwrap())
            .unwrap()
            .unwrap()
            .value(),
        99.0
    );
}
