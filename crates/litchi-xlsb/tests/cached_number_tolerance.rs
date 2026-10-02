#![allow(
    clippy::unwrap_used,
    clippy::float_cmp,
    reason = "test assertions use panic-on-failure extraction and exact literal values by design"
)]

//! Audit finding F3.10: cached `Xnum` values that are NaN, infinite,
//! denormalized, or negative zero. The repair is opt-in
//! (`Leniency::TolerateCachedNumbers`); by default the cell-value snapshot and
//! the indexed lookup still refuse such a value.

use std::io::Cursor;

use litchi_core::sheet::{Cell as _, CellValue};
use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsb::cell_values::{CellError, Reference, Value};
use litchi_xlsb::raw::{Kind, Records, Writer, kind};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::{
    CachedNumberDefect, Defect, Leniency, SourceBackedWorkbook, ToleranceReport, Workbook,
};

const SHEET1: &str = "/xl/worksheets/sheet1.bin";

/// One sheet with `1.5` at A1 and `2.5` at B1, both stored as `BrtCellReal`.
fn numbers_workbook() -> Vec<u8> {
    let mut sheet = MutableWorksheet::new("Numbers");
    sheet.set_cell(0, 0, 1.5);
    sheet.set_cell(0, 1, 2.5);
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).unwrap();
    output.into_inner()
}

/// Replace the cached `Xnum` of the first record of `record_type`.
fn with_cached_number(source: &[u8], record_type: Kind, value: f64) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(source)).unwrap();
    let uri = PackURI::new(SHEET1).unwrap();
    let mut stream: Vec<(Kind, Vec<u8>)> = Records::new(package.get_part(&uri).unwrap().blob())
        .map(|record| {
            let record = record.unwrap();
            (record.kind(), record.payload().to_vec())
        })
        .collect();
    let payload = &mut stream
        .iter_mut()
        .find(|(kind, _)| *kind == record_type)
        .unwrap()
        .1;
    payload[8..16].copy_from_slice(&value.to_le_bytes());
    let mut blob = Vec::new();
    let mut writer = Writer::new(&mut blob);
    for (kind, payload) in &stream {
        writer.write_record(*kind, payload).unwrap();
    }
    package.get_part_mut(&uri).unwrap().set_blob(blob);
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    output.into_inner()
}

fn assert_single(
    report: &ToleranceReport,
    defect: CachedNumberDefect,
    record_type: Kind,
    value: f64,
) {
    assert_eq!(report.total(), 1, "{:?}", report.defects());
    let entry = report.defects()[0];
    assert_eq!(entry.defect(), Defect::CachedNumber(defect));
    assert_eq!(entry.record_type(), record_type.get());
    assert_eq!(entry.location(), 0);
    assert_eq!(
        entry.observed(),
        u32::try_from(value.to_bits() >> 32).unwrap()
    );
}

fn a1() -> Reference {
    Reference::new(0, 0).unwrap()
}

fn b1() -> Reference {
    Reference::new(0, 1).unwrap()
}

#[test]
fn a_nan_cache_is_still_refused_by_default() {
    let bytes = with_cached_number(&numbers_workbook(), kind::CELL_REAL, f64::NAN);
    let workbook = Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(workbook.leniency(), Leniency::Strict);
    assert!(workbook.cell_values(0).is_err());

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
    let indexed = source
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .indexed_values()
        .unwrap()
        .unwrap();
    assert!(indexed.cached_value(0, 0).is_err());
    assert!(source.tolerance_report().is_clean());
}

#[test]
fn the_opt_in_reads_non_finite_caches_as_num_errors_and_records_them() {
    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let bytes = with_cached_number(&numbers_workbook(), kind::CELL_REAL, value);
        let workbook = Workbook::new_with_leniency(
            Cursor::new(bytes.clone()),
            Leniency::TolerateCachedNumbers,
        )
        .unwrap();
        assert!(workbook.tolerance_report().is_clean());

        let snapshot = workbook.cell_values(0).unwrap();
        assert_eq!(
            snapshot.cell(a1()).unwrap().unwrap().value().clone(),
            Value::Error(CellError::Number)
        );
        assert_eq!(
            snapshot.cell(b1()).unwrap().unwrap().value().clone(),
            Value::Number(2.5)
        );
        assert_single(
            snapshot.tolerance_report(),
            CachedNumberDefect::NonFinite,
            kind::CELL_REAL,
            value,
        );

        let worksheet = workbook.worksheet(0).unwrap();
        assert_eq!(
            *worksheet.get_cell(0, 0).unwrap().value(),
            CellValue::Error("#NUM!".to_string())
        );
        assert_eq!(
            *worksheet.get_cell(0, 1).unwrap().value(),
            CellValue::Float(2.5)
        );
        assert_single(
            worksheet.tolerance_report(),
            CachedNumberDefect::NonFinite,
            kind::CELL_REAL,
            value,
        );

        let source = SourceBackedWorkbook::from_reader_with_leniency(
            Cursor::new(bytes),
            Leniency::TolerateCachedNumbers,
        )
        .unwrap();
        let sheet = source.worksheet_by_index(0).unwrap().unwrap();
        let materialized = sheet.materialize().unwrap();
        assert_single(
            materialized.tolerance_report(),
            CachedNumberDefect::NonFinite,
            kind::CELL_REAL,
            value,
        );
        let indexed = sheet.indexed_values().unwrap().unwrap();
        assert_eq!(
            indexed.cached_value(0, 0).unwrap(),
            Some(CellValue::Error("#NUM!".to_string()))
        );
        assert_single(
            &source.tolerance_report(),
            CachedNumberDefect::NonFinite,
            kind::CELL_REAL,
            value,
        );
    }
}

#[test]
fn the_opt_in_reads_denormalized_caches_as_zero_and_records_them() {
    for value in [-0.0, f64::MIN_POSITIVE / 2.0] {
        let bytes = with_cached_number(&numbers_workbook(), kind::CELL_REAL, value);
        let workbook =
            Workbook::new_with_leniency(Cursor::new(bytes), Leniency::TolerateCachedNumbers)
                .unwrap();
        let snapshot = workbook.cell_values(0).unwrap();
        let Value::Number(read) = *snapshot.cell(a1()).unwrap().unwrap().value() else {
            panic!("expected a number");
        };
        assert_eq!(read.to_bits(), 0, "{value:e} must read as +0");
        assert_single(
            snapshot.tolerance_report(),
            CachedNumberDefect::Denormalized,
            kind::CELL_REAL,
            value,
        );
    }
}

#[test]
fn the_opt_in_keeps_a_formula_and_caches_num_for_a_nan_result() {
    let source = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/universal-content.xlsb"
    ))
    .unwrap();
    let bytes = with_cached_number(&source, kind::FMLA_NUM, f64::NAN);
    assert!(
        Workbook::new(Cursor::new(bytes.clone()))
            .unwrap()
            .cell_values(0)
            .is_err()
    );

    let workbook =
        Workbook::new_with_leniency(Cursor::new(bytes), Leniency::TolerateCachedNumbers).unwrap();
    let snapshot = workbook.cell_values(0).unwrap();
    let repaired: Vec<_> = snapshot
        .cells()
        .filter(|cell| *cell.value() == Value::FormulaErrorCache(CellError::Number))
        .collect();
    assert_eq!(repaired.len(), 1);
    assert!(repaired[0].formula().is_some());
    assert_eq!(snapshot.tolerance_report().total(), 1);
    assert_eq!(
        snapshot.tolerance_report().defects()[0].record_type(),
        kind::FMLA_NUM.get()
    );
    assert_eq!(workbook.worksheet(0).unwrap().tolerance_report().total(), 1);
}

#[test]
fn a_tolerant_snapshot_still_commits_edits_to_other_cells() {
    let bytes = with_cached_number(&numbers_workbook(), kind::CELL_REAL, f64::NAN);
    let workbook =
        Workbook::new_with_leniency(Cursor::new(bytes), Leniency::TolerateCachedNumbers).unwrap();
    let mut edit = workbook.cell_values(0).unwrap().edit();
    edit.set_value(b1(), Value::Number(3.5)).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .cell(b1())
            .unwrap()
            .unwrap()
            .value()
            .clone(),
        Value::Number(3.5)
    );
    assert_eq!(
        commit
            .snapshot()
            .cell(a1())
            .unwrap()
            .unwrap()
            .value()
            .clone(),
        Value::Error(CellError::Number)
    );
}

#[test]
fn the_opt_in_changes_nothing_on_a_conformant_workbook() {
    let workbook = Workbook::new_with_leniency(
        Cursor::new(numbers_workbook()),
        Leniency::TolerateCachedNumbers,
    )
    .unwrap();
    assert!(
        workbook
            .cell_values(0)
            .unwrap()
            .tolerance_report()
            .is_clean()
    );
    let worksheet = workbook.worksheet(0).unwrap();
    assert!(worksheet.tolerance_report().is_clean());
    assert_eq!(
        *worksheet.get_cell(0, 0).unwrap().value(),
        CellValue::Float(1.5)
    );
}
