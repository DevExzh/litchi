//! Real-workbook coverage for the bounded BIFF8 `Number` edit seam.

#![allow(
    clippy::expect_used,
    reason = "fixture assumptions are the assertions of this regression test"
)]

use litchi_cfb::OleFile;
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xls::Workbook;
use litchi_xls::cell_values::{Reference, Selector, Snapshot, Storage};
use std::io::Cursor;
use std::path::{Path, PathBuf};

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn real_mulrk_field_edit_remains_packed() {
    let path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/xls/FormulaSheetRange.xls");
    let source = Snapshot::from_bytes(std::fs::read(path).expect("read MulRk fixture"))
        .expect("parse MulRk fixture");
    let (sheet_position, reference) = source
        .worksheets()
        .find_map(|sheet| {
            sheet
                .cells()
                .find(|cell| cell.storage() == Storage::MulRk)
                .map(|cell| (sheet.position(), cell.reference()))
        })
        .expect("fixture contains MulRk");
    let mut edit = source.edit();
    edit.set_value(
        Selector::Position(sheet_position),
        reference,
        litchi_xls::cell_values::Value::Number(123.0),
    )
    .expect("stage exact RK value");
    let commit = edit.commit().expect("commit MulRk field");
    let cell = commit
        .snapshot()
        .worksheet(Selector::Position(sheet_position))
        .expect("resolve sheet")
        .expect("sheet exists")
        .cell(reference)
        .expect("unique cell")
        .expect("cell exists");
    assert_eq!(cell.storage(), Storage::MulRk);
    assert_eq!(cell.value(), &litchi_xls::cell_values::Value::Number(123.0));
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("inverse MulRk patch")
            .bytes(),
        source.bytes()
    );
}

#[test]
fn real_mulrk_edge_deletion_rebuilds_and_semantically_repacks() {
    let path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/xls/FormulaSheetRange.xls");
    let source = Snapshot::from_bytes(std::fs::read(path).expect("read MulRk fixture"))
        .expect("parse MulRk fixture");
    let (sheet_position, reference) = source
        .worksheets()
        .find_map(|sheet| {
            sheet
                .cells()
                .find(|cell| cell.storage() == Storage::MulRk)
                .map(|cell| (sheet.position(), cell.reference()))
        })
        .expect("fixture contains a MulRk edge");
    let mut transaction = source.transaction();
    transaction
        .remove_cell(Selector::Position(sheet_position), reference)
        .expect("stage packed edge removal");
    let commit = transaction.commit().expect("remove and reopen packed edge");
    assert!(
        commit
            .snapshot()
            .worksheet(Selector::Position(sheet_position))
            .expect("resolve edited sheet")
            .expect("edited sheet exists")
            .cell(reference)
            .expect("unique removed location")
            .is_none()
    );
    let restored = commit
        .patch()
        .semantic()
        .inverse()
        .apply(commit.snapshot())
        .expect("semantically restore packed edge");
    assert_eq!(
        restored
            .snapshot()
            .worksheet(Selector::Position(sheet_position))
            .expect("resolve restored sheet")
            .expect("restored sheet exists")
            .cell(reference)
            .expect("unique restored location")
            .expect("restored packed cell")
            .storage(),
        Storage::MulRk
    );
}

#[test]
fn real_fixture_number_commit_reopens_and_preserves_other_streams() {
    let source = Snapshot::from_bytes(
        std::fs::read(fixture("44010-TwoCharts.xls")).expect("read real XLS fixture"),
    )
    .expect("parse real XLS fixture");
    let (sheet_position, number) = source
        .worksheets()
        .find_map(|worksheet| {
            worksheet
                .numbers()
                .find(|number| number.value().is_finite())
                .map(|number| (worksheet.position(), number))
        })
        .expect("real fixture contains a finite BIFF8 Number record");
    let reference = Reference::new(
        u32::from(number.reference().row()),
        u32::from(number.reference().column()),
    )
    .expect("checked source reference");
    let replacement = if number.value().to_bits() == 12_345.25_f64.to_bits() {
        54_321.5
    } else {
        12_345.25
    };

    let mut edit = source.edit();
    edit.set_number(Selector::Position(sheet_position), reference, replacement)
        .expect("stage Number replacement");
    let commit = edit.commit().expect("publish Number replacement");
    assert_eq!(commit.diagnostics().changed_number_fields(), 1);
    assert_eq!(commit.diagnostics().touched_streams(), 1);
    assert_eq!(
        source.workbook_stream().len(),
        commit.snapshot().workbook_stream().len()
    );
    assert!(
        source
            .workbook_stream()
            .iter()
            .zip(commit.snapshot().workbook_stream())
            .filter(|(left, right)| left != right)
            .count()
            <= 8
    );

    assert_other_streams_equal(source.bytes(), commit.snapshot().bytes());
    let reopened = Workbook::new(Cursor::new(commit.snapshot().bytes()))
        .expect("committed package reopens through the public workbook reader");
    let parsed_index = reopened
        .sheet(sheet_position)
        .and_then(litchi_xls::SheetMetadata::parsed_worksheet_index)
        .expect("edited tab remains a parsed worksheet");
    let cell = reopened
        .xls_worksheet(parsed_index)
        .expect("edited worksheet")
        .get_cell(u32::from(reference.row()), u32::from(reference.column()))
        .expect("edited cell");
    let actual = match cell.value() {
        CellValue::Float(value) | CellValue::DateTime(value) => *value,
        value @ (CellValue::Empty
        | CellValue::Bool(_)
        | CellValue::Int(_)
        | CellValue::String(_)
        | CellValue::Error(_)
        | CellValue::Formula { .. }) => {
            panic!("edited Number reopened as unexpected value {value:?}")
        },
    };
    assert_eq!(actual.to_bits(), replacement.to_bits());

    let applied = commit.patch().apply(&source).expect("apply exact source");
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    assert!(commit.patch().apply(&applied).is_err());
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(&applied)
            .expect("apply inverse")
            .bytes(),
        source.bytes()
    );
}

#[test]
fn multiple_real_fixtures_rename_and_fully_reopen() {
    for (name, replacement) in [
        ("44010-TwoCharts.xls", "Chart Data Renamed"),
        ("unicodeNameRecord.xls", "Unicode Renamed"),
    ] {
        let source = Snapshot::from_bytes(std::fs::read(fixture(name)).expect("read real fixture"))
            .expect("open real fixture transaction");
        let position = source
            .worksheets()
            .next()
            .expect("fixture has a worksheet")
            .position();
        let mut transaction = source.transaction();
        transaction
            .rename_sheet(Selector::Position(position), replacement)
            .expect("stage real-fixture rename");
        let commit = transaction
            .commit()
            .expect("commit and reopen real fixture");
        assert_eq!(
            commit
                .snapshot()
                .worksheet(Selector::Position(position))
                .expect("resolve reopened worksheet")
                .expect("reopened worksheet exists")
                .name(),
            replacement
        );
        assert_other_streams_equal(source.bytes(), commit.snapshot().bytes());
        Workbook::new(Cursor::new(commit.snapshot().bytes()))
            .expect("renamed real fixture fully reopens");
    }
}

fn assert_other_streams_equal(before: &[u8], after: &[u8]) {
    let mut source_ole = OleFile::open(Cursor::new(before)).expect("source CFB");
    let mut target_ole = OleFile::open(Cursor::new(after)).expect("target CFB");
    let before_paths = source_ole.list_streams();
    let after_paths = target_ole.list_streams();
    assert_eq!(before_paths, after_paths);
    for path in before_paths {
        if path.len() == 1 && matches!(path[0].as_str(), "Workbook" | "Book") {
            continue;
        }
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        assert_eq!(
            source_ole.open_stream(&refs).expect("source stream"),
            target_ole.open_stream(&refs).expect("target stream"),
            "stream payload changed at {path:?}"
        );
    }
}
