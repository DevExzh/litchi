#![allow(
    clippy::unwrap_used,
    reason = "test assertions use panic-on-failure extraction by design"
)]

//! Audit finding F3.12: Boolean bytes other than 0/1, undefined `BErr`
//! codes, unpaired UTF-16 surrogates, and undefined data-validation enums or
//! reserved bits are repaired by default the way Excel reads them, and every
//! repair is recorded. The unmodified workbook is the control.

use std::io::Cursor;

use litchi_core::sheet::{Cell as _, CellValue};
use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsb::cell_values::{CellError, Reference, Value};
use litchi_xlsb::raw::{Kind, Records, Writer, kind};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::{
    Defect, SourceBackedWorkbook, ToleranceReport, Validation, ValueEncodingDefect, Workbook,
};

const SHEET1: &str = "/xl/worksheets/sheet1.bin";
const SST: &str = "/xl/sharedStrings.bin";

/// A1 TRUE, B1 #DIV/0!, C1 a shared string, D1 a Boolean formula, and one
/// classic whole-number validation on A2:A4.
fn fixture() -> Vec<u8> {
    let mut sheet = MutableWorksheet::new("Values");
    sheet.set_cell(0, 0, true);
    sheet.set_cell(0, 1, CellValue::Error("#DIV/0!".to_string()));
    sheet.set_cell(0, 2, "text\u{e9}");
    sheet.set_cell(
        0,
        3,
        CellValue::Formula {
            formula: "1=1".to_string(),
            cached_value: Some(Box::new(CellValue::Bool(true))),
            is_array: false,
            array_range: None,
        },
    );
    let mut rule = Validation::new(1, "A2:A4".to_string());
    rule.operator = 0;
    rule.formula1 = Some("1".to_string());
    rule.formula2 = Some("4".to_string());
    sheet.add_data_validation(rule);
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).unwrap();
    output.into_inner()
}

fn rewrite(
    source: &[u8],
    part: &str,
    record_type: Kind,
    edit: impl FnOnce(&mut Vec<u8>),
) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(source)).unwrap();
    let uri = PackURI::new(part).unwrap();
    let mut stream: Vec<(Kind, Vec<u8>)> = Records::new(package.get_part(&uri).unwrap().blob())
        .map(|record| {
            let record = record.unwrap();
            (record.kind(), record.payload().to_vec())
        })
        .collect();
    edit(
        &mut stream
            .iter_mut()
            .find(|(kind, _)| *kind == record_type)
            .unwrap()
            .1,
    );
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
    defect: ValueEncodingDefect,
    record_type: Kind,
    observed: u32,
) {
    assert_eq!(report.total(), 1, "{:?}", report.defects());
    let entry = report.defects()[0].clone();
    assert_eq!(entry.defect(), Defect::ValueEncoding(defect));
    assert_eq!(entry.record_type(), record_type.get());
    assert_eq!(entry.observed(), observed);
}

fn cell(row: u32, column: u32) -> Reference {
    Reference::new(row, column).unwrap()
}

/// The eager worksheet, its report, and the cell-value snapshot's report.
fn open(bytes: &[u8]) -> (Workbook, litchi_xlsb::Worksheet) {
    let workbook = Workbook::new(Cursor::new(bytes.to_vec())).unwrap();
    let worksheet = workbook.worksheet(0).unwrap();
    (workbook, worksheet)
}

fn indexed_value(bytes: &[u8], row: u32, column: u32) -> (Option<CellValue>, ToleranceReport) {
    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec())).unwrap();
    let value = source
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .indexed_values()
        .unwrap()
        .unwrap()
        .cached_value(row, column)
        .unwrap();
    (value, source.tolerance_report())
}

#[test]
fn the_unmodified_control_repairs_nothing() {
    let bytes = fixture();
    let (workbook, worksheet) = open(&bytes);
    assert!(workbook.tolerance_report().is_clean());
    assert!(worksheet.tolerance_report().is_clean());
    assert!(
        workbook
            .cell_values(0)
            .unwrap()
            .tolerance_report()
            .is_clean()
    );
    assert_eq!(worksheet.data_validations().len(), 1);
    assert!(indexed_value(&bytes, 0, 0).1.is_clean());
}

#[test]
fn a_nonzero_boolean_byte_reads_as_true_and_is_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, kind::CELL_BOOL, |payload| {
        payload[8] = 2
    });
    let (workbook, worksheet) = open(&bytes);
    assert_eq!(
        *worksheet.get_cell(0, 0).unwrap().value(),
        CellValue::Bool(true)
    );
    assert_single(
        worksheet.tolerance_report(),
        ValueEncodingDefect::BooleanByte,
        kind::CELL_BOOL,
        2,
    );

    let snapshot = workbook.cell_values(0).unwrap();
    assert_eq!(
        *snapshot.cell(cell(0, 0)).unwrap().unwrap().value(),
        Value::Boolean(true)
    );
    assert_single(
        snapshot.tolerance_report(),
        ValueEncodingDefect::BooleanByte,
        kind::CELL_BOOL,
        2,
    );

    let (value, report) = indexed_value(&bytes, 0, 0);
    assert_eq!(value, Some(CellValue::Bool(true)));
    assert_single(
        &report,
        ValueEncodingDefect::BooleanByte,
        kind::CELL_BOOL,
        2,
    );
}

#[test]
fn a_nonzero_formula_boolean_cache_reads_as_true_and_is_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, kind::FMLA_BOOL, |payload| {
        payload[8] = 0x7F
    });
    let (workbook, worksheet) = open(&bytes);
    assert_single(
        worksheet.tolerance_report(),
        ValueEncodingDefect::BooleanByte,
        kind::FMLA_BOOL,
        0x7F,
    );
    let snapshot = workbook.cell_values(0).unwrap();
    assert_eq!(
        *snapshot.cell(cell(0, 3)).unwrap().unwrap().value(),
        Value::FormulaBooleanCache(true)
    );
}

#[test]
fn an_undefined_error_code_reads_as_a_generic_error_and_is_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, kind::CELL_ERROR, |payload| {
        payload[8] = 0x55
    });
    let (workbook, worksheet) = open(&bytes);
    assert_eq!(
        *worksheet.get_cell(0, 1).unwrap().value(),
        CellValue::Error("#ERR!".to_string())
    );
    assert_single(
        worksheet.tolerance_report(),
        ValueEncodingDefect::ErrorCode,
        kind::CELL_ERROR,
        0x55,
    );

    // The typed cell-value model has no generic error; Excel shows #NULL!.
    let snapshot = workbook.cell_values(0).unwrap();
    assert_eq!(
        *snapshot.cell(cell(0, 1)).unwrap().unwrap().value(),
        Value::Error(CellError::Null)
    );
    assert_single(
        snapshot.tolerance_report(),
        ValueEncodingDefect::ErrorCode,
        kind::CELL_ERROR,
        0x55,
    );

    let (value, report) = indexed_value(&bytes, 0, 1);
    assert_eq!(value, Some(CellValue::Error("#NULL!".to_string())));
    assert_single(
        &report,
        ValueEncodingDefect::ErrorCode,
        kind::CELL_ERROR,
        0x55,
    );
}

#[test]
fn an_unpaired_surrogate_decodes_lossily_and_is_recorded() {
    // The SST item is a RichStr: flags byte, character count, UTF-16 units.
    let bytes = rewrite(&fixture(), SST, kind::SST_ITEM, |payload| {
        payload[5..7].copy_from_slice(&0xD800u16.to_le_bytes());
    });
    let (workbook, worksheet) = open(&bytes);
    assert_single(
        workbook.tolerance_report(),
        ValueEncodingDefect::Utf16,
        kind::SST_ITEM,
        1,
    );
    assert_eq!(
        *worksheet.get_cell(0, 2).unwrap().value(),
        CellValue::String("\u{FFFD}ext\u{e9}".to_string())
    );
    assert!(worksheet.tolerance_report().is_clean());

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
    assert!(source.tolerance_report().is_clean());
    source
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .materialize()
        .unwrap();
    assert_single(
        &source.tolerance_report(),
        ValueEncodingDefect::Utf16,
        kind::SST_ITEM,
        1,
    );
}

#[test]
fn reserved_data_validation_bits_are_ignored_and_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, kind::D_VAL, |payload| {
        payload[3] |= 0x80
    });
    let (_, worksheet) = open(&bytes);
    let report = worksheet.tolerance_report();
    assert_eq!(report.total(), 1, "{:?}", report.defects());
    assert_eq!(
        report.defects()[0].defect(),
        Defect::ValueEncoding(ValueEncodingDefect::DataValidation)
    );
    assert_eq!(report.defects()[0].observed() & 0xff00_0000, 0x8000_0000);
    let rule = &worksheet.data_validations()[0];
    assert_eq!(rule.validation_type, 1);
    assert_eq!(rule.formula2.as_deref(), Some("4"));
}

#[test]
fn an_undefined_validation_type_falls_back_to_any_value_and_is_recorded() {
    let bytes = rewrite(&fixture(), SHEET1, kind::D_VAL, |payload| {
        payload[0] = (payload[0] & 0xF0) | 0x09;
    });
    let (_, worksheet) = open(&bytes);
    assert_eq!(worksheet.tolerance_report().total(), 1);
    assert_eq!(
        worksheet.tolerance_report().defects()[0].record_type(),
        kind::D_VAL.get()
    );
    let rule = &worksheet.data_validations()[0];
    assert_eq!(rule.validation_type, 0);
    assert_eq!(rule.cell_ranges, "A2:A4");
}
