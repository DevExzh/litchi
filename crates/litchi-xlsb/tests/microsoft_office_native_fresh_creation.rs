//! Fresh creation from selected native values; this is an explicit semantic copy, not a source resave.
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait};
use litchi_xlsb::{
    Package, SourceBackedWorkbook,
    builder::{Builder, Worksheet},
};
use std::{io::Cursor, path::Path};

fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}

#[test]
fn native_values_seed_an_independent_fresh_xlsb_with_exact_eager_and_indexed_readback() {
    let metadata = native("excel-metadata.xlsb");
    let controls = native("excel-form-controls.xlsb");
    let source = Package::from_bytes(&metadata).unwrap();
    let control = Package::from_bytes(&controls).unwrap();
    let book = source.workbook().unwrap();
    let sheet = book.worksheet_by_name("Sheet1").unwrap();
    let label = sheet.get_cell(0, 0).unwrap().value().clone();
    let number = book
        .worksheet_by_name("Print_015")
        .unwrap()
        .get_cell(1, 1)
        .unwrap()
        .value()
        .clone();
    assert_eq!(label, CellValue::String("EXCEL_METADATA_015".into()));
    assert_eq!(number, CellValue::Int(42));
    let control_book = control.workbook().unwrap();
    let boolean = control_book
        .worksheet_by_name("Sheet1")
        .unwrap()
        .get_cell(1, 1)
        .unwrap()
        .value()
        .clone();
    assert_eq!(boolean, CellValue::Bool(true));
    let mut values = Worksheet::new("Native values");
    values.set_cell(0, 0, label);
    values.set_cell(1, 1, number);
    values.set_cell(2, 0, "日本語 🍎");
    let mut flags = Worksheet::new("Native controls");
    flags.set_cell(1, 1, boolean);
    let mut builder = Builder::new();
    builder.add_worksheet(values);
    builder.add_worksheet(flags);
    let created = builder.build().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    created.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let bytes = reopened.to_bytes().unwrap();
    assert_eq!(bytes, created.to_bytes().unwrap());
    let book = reopened.workbook().unwrap();
    assert_eq!(book.worksheet_names(), ["Native values", "Native controls"]);
    let lazy = SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
    for (sheet, row, col, expected) in [
        (0, 0, 0, CellValue::String("EXCEL_METADATA_015".into())),
        // Fresh numeric cells use BrtCellReal rather than the source's RK storage.
        (0, 1, 1, CellValue::Float(42.0)),
        (0, 2, 0, CellValue::String("日本語 🍎".into())),
        (1, 1, 1, CellValue::Bool(true)),
    ] {
        assert_eq!(
            book.worksheet_by_name(&book.worksheet_names()[sheet])
                .unwrap()
                .get_cell(row, col)
                .unwrap()
                .value(),
            &expected
        );
        assert_eq!(
            lazy.sheet(sheet)
                .unwrap()
                .unwrap()
                .indexed_values()
                .unwrap()
                .unwrap()
                .cached_value(row, col)
                .unwrap(),
            Some(expected)
        );
    }
    assert_eq!(source.to_bytes().unwrap(), metadata);
    assert_eq!(control.to_bytes().unwrap(), controls);
    assert_eq!(native("excel-metadata.xlsb"), metadata);
    assert_eq!(native("excel-form-controls.xlsb"), controls);
}
