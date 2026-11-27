//! Native Excel coverage of the shared one-based worksheet contract.
use litchi_core::sheet::{Cell, CellValue, WorkbookTrait, Worksheet};
use std::io::Cursor;

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}

fn grid<S: Worksheet>(sheet: &S) {
    assert_eq!(
        sheet.cell_value(2, 2).unwrap().as_ref(),
        &CellValue::Bool(true)
    );
    assert_eq!(sheet.row(1).unwrap()[1], CellValue::Bool(true));
    assert_eq!(sheet.dimensions(), Some((1, 1, 12, 4)));
    for cell in sheet.cells() {
        let cell = cell.unwrap();
        let (row, col) = (cell.row(), cell.column());
        assert!(row >= 1 && col >= 1);
        let numeric = sheet.cell(row, col).unwrap();
        assert_eq!(numeric.coordinate(), cell.coordinate());
        assert_eq!(numeric.value(), cell.value());
        assert_eq!(sheet.cell_value(row, col).unwrap().as_ref(), cell.value());
        let named = sheet.cell_by_coordinate(&cell.coordinate()).unwrap();
        assert_eq!((named.row(), named.column()), (row, col));
        assert_eq!(named.value(), cell.value());
    }
    for cell in [
        sheet.cell(8, 3).unwrap(),
        sheet.cell_by_coordinate("C8").unwrap(),
    ] {
        assert_eq!(
            (cell.row(), cell.column(), cell.coordinate()),
            (8, 3, "C8".into())
        );
        assert!(cell.is_empty());
    }
}

#[test]
fn native_shared_sheet_traits_use_one_based_cells_and_keep_native_storage_zero_based() {
    let bytes = native("excel-form-controls.xlsb");
    let book = litchi_xlsb::Workbook::new(Cursor::new(&bytes)).unwrap();
    let sheet = book.worksheet_by_name("Sheet1").unwrap();
    grid(&sheet);
    let stored = sheet.get_cell(1, 1).unwrap();
    assert_eq!(stored.value(), &CellValue::Bool(true));
    assert_eq!((stored.row(), stored.column()), (1, 1));
    assert_eq!((Cell::row(stored), Cell::column(stored)), (2, 2));
    assert!(bytes == native("excel-form-controls.xlsb"));
}

#[test]
fn native_shared_sheet_traits_refuse_out_of_grid_coordinates() {
    let bytes = native("excel-form-controls.xlsb");
    let book = litchi_xlsb::Workbook::new(Cursor::new(&bytes)).unwrap();
    let sheet = book.worksheet_by_name("Sheet1").unwrap();
    for (row, col) in [
        (0, 1),
        (1, 0),
        (1048576 + 1, 1),
        (1, 16384 + 1),
        (u32::MAX, u32::MAX),
    ] {
        assert!(matches!(
            sheet.cell(row, col),
            Err(litchi_xlsb::package::PackageError::InvalidCellReference(_))
        ));
        assert!(matches!(
            sheet.cell_value(row, col),
            Err(litchi_xlsb::package::PackageError::InvalidCellReference(_))
        ));
    }
    for label in [
        "A0",
        "",
        "XFE1",
        "A1048577",
        "ZZZZZZZZZZZZZZZZZZZZ1",
        "A4294967295",
    ] {
        assert!(matches!(
            sheet.cell_by_coordinate(label),
            Err(litchi_xlsb::package::PackageError::InvalidCellReference(_))
        ));
    }
    let last = sheet.cell(1048576, 16384).unwrap();
    assert_eq!((Cell::row(&last), Cell::column(&last)), (1048576, 16384));
    assert_eq!(last.coordinate(), "XFD1048576");
    assert!(last.is_empty());
    assert_eq!(
        sheet.cell_value(1048576, 16384).unwrap().as_ref(),
        &CellValue::Empty
    );
    assert!(bytes == native("excel-form-controls.xlsb"));
}
