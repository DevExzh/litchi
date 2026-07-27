//! Round-trip tests for the BIFF8 Table record (what-if data tables).

use litchi_core::sheet::Cell;
use litchi_ole::xls::writer::XlsWriter;
use litchi_ole::xls::{
    XlsDataTable, XlsDataTableInputCell, XlsDataTableKind, XlsDataTableRange, XlsWorkbook,
};
use std::io::Cursor;

#[test]
fn data_tables_round_trip_through_writer_and_reader() {
    let one_variable = XlsDataTable::one_variable(
        XlsDataTableRange::new(2, 8, 3, 5).unwrap(),
        true,
        XlsDataTableInputCell::Present { row: 0, col: 6 },
    );
    let mut two_variable = XlsDataTable::two_variable(
        XlsDataTableRange::new(2, 8, 7, 9).unwrap(),
        XlsDataTableInputCell::Present { row: 0, col: 10 },
        XlsDataTableInputCell::Deleted,
    );
    two_variable.set_always_calc(true);

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Tables").unwrap();
    writer.write_number(sheet, 3, 3, 1.5).unwrap();
    writer.add_data_table(sheet, 1, 2, one_variable).unwrap();
    writer.add_data_table(sheet, 1, 6, two_variable).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let tables = worksheet.data_tables();
    assert_eq!(tables, &[one_variable, two_variable]);

    // The anchor cell carries the PtgTbl formula.
    let anchor = worksheet.get_cell(1, 2).unwrap();
    assert!(anchor.is_formula());

    // The cell inside the table range keeps its ordinary value.
    assert!(matches!(
        worksheet.get_cell(3, 3).unwrap().value(),
        litchi_core::sheet::CellValue::Float(value) if *value == 1.5
    ));
}

#[test]
fn data_table_anchor_validation() {
    let range = XlsDataTableRange::new(2, 8, 3, 5).unwrap();
    let table =
        XlsDataTable::one_variable(range, false, XlsDataTableInputCell::Present { row: 0, col: 0 });

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Tables").unwrap();
    // Anchor inside the range.
    assert!(writer.add_data_table(sheet, 4, 4, table).is_err());
    // Anchor beyond the BIFF8 grid.
    assert!(writer.add_data_table(sheet, 1, 300, table).is_err());
    // Duplicate anchor.
    writer.add_data_table(sheet, 1, 2, table).unwrap();
    assert!(writer.add_data_table(sheet, 1, 2, table).is_err());
    // Anchor colliding with a valued cell.
    writer.write_number(sheet, 0, 0, 1.0).unwrap();
    assert!(writer.add_data_table(sheet, 0, 0, table).is_err());
    // Unknown sheet.
    assert!(writer.add_data_table(9, 1, 0, table).is_err());
}

#[test]
fn kind_accessors_expose_input_cells() {
    let table = XlsDataTable::two_variable(
        XlsDataTableRange::new(2, 8, 3, 5).unwrap(),
        XlsDataTableInputCell::Present { row: 1, col: 2 },
        XlsDataTableInputCell::Deleted,
    );
    let XlsDataTableKind::TwoVariable { row_input, column_input } = table.kind() else {
        panic!()
    };
    assert_eq!(*row_input, XlsDataTableInputCell::Present { row: 1, col: 2 });
    assert_eq!(*column_input, XlsDataTableInputCell::Deleted);
}
