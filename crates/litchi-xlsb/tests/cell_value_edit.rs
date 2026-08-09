#![allow(
    clippy::pedantic,
    clippy::unwrap_used,
    clippy::wildcard_enum_match_arm,
    reason = "integration tests use panic-on-failure extraction and exhaustive fixture diagnostics"
)]

//! Real-workbook regression coverage for the bounded `BrtCellReal` seam.

#![allow(
    clippy::expect_used,
    reason = "the regression fixture must fail immediately when its checked corpus assumptions break"
)]

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_xlsb::{
    Workbook,
    cell_values::{CellError, CellFormula, Reference, StyleIndex, Value},
};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn real_fixture_cell_value_commit_is_source_checked_and_length_stable() {
    let (mut workbook, sheet, number) = [
        "test-data/ooxml/xlsb/universal-content.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/51519.xlsb",
        "test-data/ooxml/xlsb/62815.xlsb",
        "test-data/ooxml/xlsb/Simple.xlsb",
        "test-data/ooxml/xlsb/cond_format.xlsb",
        "test-data/ooxml/xlsb/comments.xlsb",
        "test-data/ooxml/xlsb/hyperlink.xlsb",
        "test-data/ooxml/xlsb/bug66682.xlsb",
        "test-data/ooxml/xlsb/sample.xlsb",
    ]
    .iter()
    .find_map(|relative| {
        let workbook = Workbook::new(File::open(fixture(relative)).ok()?).ok()?;
        let selected = (0..workbook.worksheet_count()).find_map(|sheet| {
            workbook
                .cell_values(sheet)
                .ok()
                .and_then(|snapshot| snapshot.numbers().find(|number| number.value().is_finite()))
                .map(|number| (sheet, number))
        });
        selected.map(|(sheet, number)| (workbook, sheet, number))
    })
    .expect("a real fixture contains a finite BrtCellReal value");
    let reference =
        Reference::new(number.reference().row(), number.reference().column()).expect("reference");
    let before = workbook.cell_values(sheet).expect("before snapshot");
    let sheet_name = workbook.worksheet_names()[sheet].clone();
    assert_eq!(
        workbook
            .cell_values_by_name(&sheet_name)
            .expect("snapshot by name")
            .source_bytes(),
        before.source_bytes()
    );

    let replacement = if number.value().to_bits() == 12_345.25_f64.to_bits() {
        54_321.5
    } else {
        12_345.25
    };
    let mut edit = before.edit();
    edit.set_number(reference, replacement).expect("set number");
    let commit = edit.commit().expect("commit");
    let after = commit.patch().after();
    assert_eq!(before.source_bytes().len(), after.len());
    assert!(
        before
            .source_bytes()
            .iter()
            .zip(after)
            .filter(|(left, right)| left != right)
            .count()
            <= 8,
        "only the selected f64 field may change"
    );

    let published = workbook
        .apply_cell_values_by_name(&sheet_name, &commit)
        .expect("publish");
    assert_eq!(
        published
            .number(reference)
            .expect("lookup")
            .expect("edited value")
            .value()
            .to_bits(),
        replacement.to_bits()
    );
    assert!(commit.patch().apply(before.source_bytes()).is_ok());
    assert!(commit.patch().apply(after).is_err());

    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).expect("save");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen");
    assert_eq!(
        reopened
            .cell_values(sheet)
            .expect("reopened snapshot")
            .number(reference)
            .expect("lookup")
            .expect("edited value")
            .value()
            .to_bits(),
        replacement.to_bits()
    );
}

#[test]
fn real_fixture_non_real_scalar_or_formula_cache_round_trips() {
    let fixtures = [
        "test-data/ooxml/xlsb/universal-content.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/51519.xlsb",
        "test-data/ooxml/xlsb/62815.xlsb",
        "test-data/ooxml/xlsb/Simple.xlsb",
        "test-data/ooxml/xlsb/cond_format.xlsb",
        "test-data/ooxml/xlsb/comments.xlsb",
        "test-data/ooxml/xlsb/hyperlink.xlsb",
        "test-data/ooxml/xlsb/bug66682.xlsb",
        "test-data/ooxml/xlsb/sample.xlsb",
    ];
    let (mut workbook, sheet, reference, replacement) = fixtures
        .iter()
        .find_map(|relative| {
            let workbook = Workbook::new(File::open(fixture(relative)).ok()?).ok()?;
            let selected = (0..workbook.worksheet_count()).find_map(|sheet| {
                let snapshot = workbook.cell_values(sheet).ok()?;
                snapshot.cells().find_map(|cell| {
                    replacement_for(cell.value())
                        .map(|replacement| (sheet, cell.reference(), replacement))
                })
            });
            selected
                .map(|(sheet, reference, replacement)| (workbook, sheet, reference, replacement))
        })
        .expect("a real fixture contains an editable non-BrtCellReal value");

    let mut edit = workbook.edit_cell_values(sheet).expect("edit");
    edit.set_value(reference, replacement.clone())
        .expect("set typed value");
    let commit = edit.commit().expect("commit");
    assert_eq!(commit.patch().before().len(), commit.patch().after().len());
    let published = workbook.apply_cell_values(sheet, &commit).expect("publish");
    assert_eq!(
        published
            .cell(reference)
            .expect("lookup")
            .expect("edited cell")
            .value(),
        &replacement
    );

    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).expect("save");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen");
    assert_eq!(
        reopened
            .cell_values(sheet)
            .expect("snapshot")
            .cell(reference)
            .expect("lookup")
            .expect("edited cell")
            .value(),
        &replacement
    );
}

#[test]
fn real_fixture_style_index_edit_is_contextually_validated() {
    let (mut workbook, sheet, reference, replacement) = [
        "test-data/ooxml/xlsb/universal-content.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/51519.xlsb",
        "test-data/ooxml/xlsb/62815.xlsb",
        "test-data/ooxml/xlsb/Simple.xlsb",
        "test-data/ooxml/xlsb/cond_format.xlsb",
        "test-data/ooxml/xlsb/comments.xlsb",
        "test-data/ooxml/xlsb/hyperlink.xlsb",
        "test-data/ooxml/xlsb/bug66682.xlsb",
        "test-data/ooxml/xlsb/sample.xlsb",
    ]
    .iter()
    .find_map(|relative| {
        let workbook = Workbook::new(File::open(fixture(relative)).ok()?).ok()?;
        if workbook.styles().cell_xfs.len() < 2 {
            return None;
        }
        let selected = (0..workbook.worksheet_count()).find_map(|sheet| {
            let snapshot = workbook.cell_values(sheet).ok()?;
            let cell = snapshot.cells().next()?;
            let replacement = StyleIndex::new(if cell.style().get() == 0 { 1 } else { 0 }).ok()?;
            Some((sheet, cell.reference(), replacement))
        });
        selected.map(|(sheet, reference, replacement)| (workbook, sheet, reference, replacement))
    })
    .expect("a real fixture contains at least two cell styles and one stored cell");

    let before = workbook.cell_values(sheet).expect("snapshot");
    let mut edit = before.edit();
    edit.set_style(reference, replacement).expect("set style");
    let commit = edit.commit().expect("commit");
    assert_eq!(before.source_bytes().len(), commit.patch().after().len());
    let published = workbook.apply_cell_values(sheet, &commit).expect("publish");
    assert_eq!(
        published
            .cell(reference)
            .expect("lookup")
            .expect("edited cell")
            .style(),
        replacement
    );

    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).expect("save");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen");
    assert_eq!(
        reopened
            .cell_values(sheet)
            .expect("snapshot")
            .cell(reference)
            .expect("lookup")
            .expect("edited cell")
            .style(),
        replacement
    );
}

#[test]
fn real_fixture_structural_crud_and_length_changing_formula_and_string_reopen() {
    let mut workbook =
        Workbook::new(File::open(fixture("test-data/ooxml/xlsb/Simple.xlsb")).expect("fixture"))
            .expect("workbook");
    let sheet = 0;
    let string_ref = Reference::new(10_000, 100).expect("string reference");
    let formula_ref = Reference::new(10_000, 101).expect("formula reference");
    let before = workbook.cell_values(sheet).expect("snapshot");
    assert!(before.cell(string_ref).expect("string lookup").is_none());
    assert!(before.cell(formula_ref).expect("formula lookup").is_none());

    let mut create = before.edit();
    create
        .insert(
            string_ref,
            StyleIndex::new(0).expect("style"),
            Value::InlineString("short".to_string()),
        )
        .expect("insert string");
    create
        .insert_formula(
            formula_ref,
            StyleIndex::new(0).expect("style"),
            Value::FormulaNumberCache(2.0),
            CellFormula::new(0, vec![0x1E, 2, 0], vec![]).expect("integer formula"),
        )
        .expect("insert formula");
    let created = create.commit().expect("create commit");
    workbook
        .apply_cell_values(sheet, &created)
        .expect("create publication");

    let mut change = workbook
        .edit_cell_values(sheet)
        .expect("edit created cells");
    change
        .set_inline_string(
            string_ref,
            "a substantially longer inline value".to_string(),
        )
        .expect("length-changing string");
    let mut number_formula = vec![0x1F];
    number_formula.extend_from_slice(&3.5_f64.to_le_bytes());
    change
        .set_formula(
            formula_ref,
            CellFormula::new(0x0002, number_formula, vec![]).expect("number formula"),
        )
        .expect("length-changing formula tokens");
    change
        .set_formula_number_cache(formula_ref, 3.5)
        .expect("formula cache");
    let changed = change.commit().expect("change commit");
    workbook
        .apply_cell_values(sheet, &changed)
        .expect("change publication");

    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).expect("save changed workbook");
    let mut reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen changed");
    let snapshot = reopened.cell_values(sheet).expect("changed snapshot");
    assert_eq!(
        snapshot
            .cell(string_ref)
            .expect("lookup")
            .expect("string")
            .value(),
        &Value::InlineString("a substantially longer inline value".to_string())
    );
    assert_eq!(
        snapshot
            .cell(formula_ref)
            .expect("lookup")
            .expect("formula")
            .value(),
        &Value::FormulaNumberCache(3.5)
    );

    let mut remove = snapshot.edit();
    remove.remove(string_ref).expect("remove string");
    remove.remove(formula_ref).expect("remove formula");
    let removed = remove.commit().expect("remove commit");
    reopened
        .apply_cell_values(sheet, &removed)
        .expect("remove publication");
    let mut bytes = Cursor::new(Vec::new());
    reopened.save(&mut bytes).expect("save removed workbook");
    let final_workbook = Workbook::new(Cursor::new(bytes.into_inner())).expect("final reopen");
    let final_snapshot = final_workbook.cell_values(sheet).expect("final snapshot");
    assert!(final_snapshot.cell(string_ref).expect("lookup").is_none());
    assert!(final_snapshot.cell(formula_ref).expect("lookup").is_none());
}

#[test]
fn real_fixture_publication_rejects_missing_style_and_shared_string_dependencies() {
    let workbook =
        Workbook::new(File::open(fixture("test-data/ooxml/xlsb/Simple.xlsb")).expect("fixture"))
            .expect("workbook");
    let sheet = 0;
    let reference = Reference::new(10_001, 100).expect("reference");

    let mut invalid_style_workbook =
        Workbook::new(File::open(fixture("test-data/ooxml/xlsb/Simple.xlsb")).expect("fixture"))
            .expect("workbook");
    let mut invalid_style = workbook.edit_cell_values(sheet).expect("edit");
    invalid_style
        .insert(
            reference,
            StyleIndex::new(0x00FF_FFFF).expect("wire style"),
            Value::Number(1.0),
        )
        .expect("insert");
    let commit = invalid_style.commit().expect("worksheet commit");
    assert!(
        invalid_style_workbook
            .apply_cell_values(sheet, &commit)
            .is_err()
    );

    let mut invalid_shared_workbook =
        Workbook::new(File::open(fixture("test-data/ooxml/xlsb/Simple.xlsb")).expect("fixture"))
            .expect("workbook");
    let mut invalid_shared = workbook.edit_cell_values(sheet).expect("edit");
    invalid_shared
        .insert(
            reference,
            StyleIndex::new(0).expect("style"),
            Value::SharedStringIndex(u32::MAX),
        )
        .expect("insert");
    let commit = invalid_shared.commit().expect("worksheet commit");
    assert!(
        invalid_shared_workbook
            .apply_cell_values(sheet, &commit)
            .is_err()
    );
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "Value is deliberately non-exhaustive across this public-API integration boundary"
)]
fn replacement_for(value: &Value) -> Option<Value> {
    match value {
        Value::RkNumber(number) => {
            Some(Value::RkNumber(if number.to_bits() == 1.0_f64.to_bits() {
                2.0
            } else {
                1.0
            }))
        },
        Value::Error(error) => Some(Value::Error(alternate_error(*error))),
        Value::Boolean(boolean) => Some(Value::Boolean(!*boolean)),
        Value::InlineString(string) => same_length_string(string).map(Value::InlineString),
        Value::FormulaStringCache(string) => {
            same_length_string(string).map(Value::FormulaStringCache)
        },
        Value::FormulaNumberCache(number) => Some(Value::FormulaNumberCache(
            if number.to_bits() == 1.0_f64.to_bits() {
                2.0
            } else {
                1.0
            },
        )),
        Value::FormulaBooleanCache(boolean) => Some(Value::FormulaBooleanCache(!*boolean)),
        Value::FormulaErrorCache(error) => Some(Value::FormulaErrorCache(alternate_error(*error))),
        Value::Blank | Value::Number(_) | Value::SharedStringIndex(_) => None,
        _ => None,
    }
}

fn same_length_string(value: &str) -> Option<String> {
    let units = value.encode_utf16().count();
    if units == 0 {
        return None;
    }
    let x = "X".repeat(units);
    Some(if x == value { "Y".repeat(units) } else { x })
}

const fn alternate_error(error: CellError) -> CellError {
    if matches!(error, CellError::Reference) {
        CellError::Value
    } else {
        CellError::Reference
    }
}
