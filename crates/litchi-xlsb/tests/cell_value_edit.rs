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
    Package, Workbook,
    cell_values::{CellError, CellFormula, Reference, StyleIndex, Value},
};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

#[path = "support/projection.rs"]
mod projection;
use projection::OwnedWorkbook as _;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn package_edit(workbook: &Workbook, sheet: usize) -> (Package, litchi_xlsb::cell_values::Edit) {
    let package = workbook.to_package();
    let name = workbook.worksheet_names()[sheet].clone();
    let edit = package
        .edit_cell_values(name)
        .expect("cell-value selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet selector");
    (package, edit)
}

fn publish(commit: litchi_xlsb::cell_values::Commit) -> Workbook {
    commit
        .into_package()
        .owned_workbook()
        .expect("published workbook")
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

    // The package patch owns the whole OPC graph.  Keep the worksheet codec
    // assertions on its explicit raw transaction so this regression still
    // proves the BrtCellReal replacement is length-stable and source-checked.
    let mut raw_edit = before.draft();
    raw_edit
        .set_number(reference, replacement)
        .expect("raw set number");
    let (_, raw_patch) = raw_edit.finish().expect("raw commit");
    let raw_after = raw_patch.after();
    assert_eq!(before.source_bytes().len(), raw_after.len());
    assert!(
        before
            .source_bytes()
            .iter()
            .zip(raw_after)
            .filter(|(left, right)| left != right)
            .count()
            <= 8,
        "only the selected f64 field may change"
    );
    assert!(raw_patch.apply(before.source_bytes()).is_ok());
    assert!(raw_patch.apply(raw_after).is_err());

    let (_, mut edit) = package_edit(&workbook, sheet);
    edit.set_number(reference, replacement).expect("set number");
    let commit = edit.commit().expect("commit");
    assert!(!commit.patch().is_noop());
    workbook = publish(commit);
    let published = workbook.cell_values(sheet).expect("published snapshot");
    assert_eq!(
        published
            .number(reference)
            .expect("lookup")
            .expect("edited value")
            .value()
            .to_bits(),
        replacement.to_bits()
    );
    let mut bytes = Cursor::new(Vec::new());
    workbook.to_package().write_to(&mut bytes).expect("save");
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

    let before = workbook.cell_values(sheet).expect("before snapshot");
    let mut raw_edit = before.draft();
    raw_edit
        .set_value(reference, replacement.clone())
        .expect("raw set typed value");
    let (_, raw_patch) = raw_edit.finish().expect("raw commit");
    assert_eq!(raw_patch.before().len(), raw_patch.after().len());

    let (_, mut edit) = package_edit(&workbook, sheet);
    edit.set_value(reference, replacement.clone())
        .expect("set typed value");
    let commit = edit.commit().expect("commit");
    assert!(!commit.patch().is_noop());
    workbook = publish(commit);
    let published = workbook.cell_values(sheet).expect("published snapshot");
    assert_eq!(
        published
            .cell(reference)
            .expect("lookup")
            .expect("edited cell")
            .value(),
        &replacement
    );

    let mut bytes = Cursor::new(Vec::new());
    workbook.to_package().write_to(&mut bytes).expect("save");
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
    let mut raw_edit = before.draft();
    raw_edit
        .set_style(reference, replacement)
        .expect("raw set style");
    let (_, raw_patch) = raw_edit.finish().expect("raw style commit");
    assert_eq!(before.source_bytes().len(), raw_patch.after().len());
    let (_, mut edit) = package_edit(&workbook, sheet);
    edit.set_style(reference, replacement).expect("set style");
    let commit = edit.commit().expect("commit");
    assert!(!commit.patch().is_noop());
    workbook = publish(commit);
    let published = workbook.cell_values(sheet).expect("published snapshot");
    assert_eq!(
        published
            .cell(reference)
            .expect("lookup")
            .expect("edited cell")
            .style(),
        replacement
    );

    let mut bytes = Cursor::new(Vec::new());
    workbook.to_package().write_to(&mut bytes).expect("save");
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

    let (_, mut create) = package_edit(&workbook, sheet);
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
            CellFormula::new(
                0,
                vec![0x1E, 2, 0],
                vec![],
                litchi_xlsb::raw::Limits::DEFAULT,
            )
            .expect("integer formula"),
        )
        .expect("insert formula");
    let created = create.commit().expect("create commit");
    workbook = publish(created);

    let (_, mut change) = package_edit(&workbook, sheet);
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
            CellFormula::new(
                0x0002,
                number_formula,
                vec![],
                litchi_xlsb::raw::Limits::DEFAULT,
            )
            .expect("number formula"),
        )
        .expect("length-changing formula tokens");
    change
        .set_formula_number_cache(formula_ref, 3.5)
        .expect("formula cache");
    let changed = change.commit().expect("change commit");
    workbook = publish(changed);

    let mut bytes = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut bytes)
        .expect("save changed workbook");
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

    let (_, mut remove) = package_edit(&reopened, sheet);
    remove.remove(string_ref).expect("remove string");
    remove.remove(formula_ref).expect("remove formula");
    let removed = remove.commit().expect("remove commit");
    reopened = publish(removed);
    let mut removed_bytes = Cursor::new(Vec::new());
    reopened
        .to_package()
        .write_to(&mut removed_bytes)
        .expect("save removed workbook");
    let final_workbook =
        Workbook::new(Cursor::new(removed_bytes.into_inner())).expect("final reopen");
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

    let (_, mut invalid_style) = package_edit(&workbook, sheet);
    invalid_style
        .insert(
            reference,
            StyleIndex::new(0x00FF_FFFF).expect("wire style"),
            Value::Number(1.0),
        )
        .expect("insert");
    assert!(invalid_style.commit().is_err());

    let (_, mut invalid_shared) = package_edit(&workbook, sheet);
    invalid_shared
        .insert(
            reference,
            StyleIndex::new(0).expect("style"),
            Value::SharedStringIndex(u32::MAX),
        )
        .expect("insert");
    assert!(invalid_shared.commit().is_err());
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

#[test]
fn real_fixture_exact_noop_publication_saves_byte_identical_output() {
    let mut workbook =
        Workbook::new(File::open(fixture("test-data/ooxml/xlsb/Simple.xlsb")).expect("fixture"))
            .expect("workbook");
    let sheet = 0;

    let mut before = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut before)
        .expect("save before");
    let before = before.into_inner();
    let before_cells: Vec<_> = workbook
        .cell_values(sheet)
        .expect("snapshot")
        .cells()
        .cloned()
        .collect();

    let (_, edit) = package_edit(&workbook, sheet);
    let commit = edit.commit().expect("commit");
    assert!(commit.patch().is_noop());
    workbook = publish(commit);

    let mut after = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut after)
        .expect("save after");
    assert_eq!(
        after.into_inner(),
        before,
        "an exact no-op publication must not change a single saved byte"
    );
    let after_cells: Vec<_> = workbook
        .cell_values(sheet)
        .expect("snapshot")
        .cells()
        .cloned()
        .collect();
    assert_eq!(after_cells, before_cells);
}

#[test]
fn real_fixture_refused_publication_saves_byte_identical_output() {
    let workbook =
        Workbook::new(File::open(fixture("test-data/ooxml/xlsb/Simple.xlsb")).expect("fixture"))
            .expect("workbook");
    let sheet = 0;

    let mut before = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut before)
        .expect("save before");
    let before = before.into_inner();

    let (_, mut edit) = package_edit(&workbook, sheet);
    edit.insert(
        Reference::new(10_001, 100).expect("reference"),
        StyleIndex::new(0x00FF_FFFF).expect("wire style"),
        Value::Number(1.0),
    )
    .expect("insert");
    assert!(edit.commit().is_err());

    let mut after = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut after)
        .expect("save after");
    assert_eq!(
        after.into_inner(),
        before,
        "a refused publication must leave the published workbook byte-identical"
    );
}
