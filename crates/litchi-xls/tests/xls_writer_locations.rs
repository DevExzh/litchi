use std::io::Cursor;

use litchi_xls::Workbook;
use litchi_xls::writer::{Column, FrozenPanes, Row, Writer};

#[test]
fn row_and_column_property_apis_round_trip_checked_locations() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Locations").unwrap();
    let row = Row::new(40).unwrap();
    let column = Column::new(12).unwrap();

    writer.write_number(sheet, 0, 0, 1.0).unwrap();
    writer.set_row_height(sheet, row, 18.0).unwrap();
    writer.set_column_width(sheet, column, 20.0).unwrap();
    writer.hide_row(sheet, row).unwrap();
    writer.hide_column(sheet, column).unwrap();
    writer.show_row(sheet, row).unwrap();
    writer.show_column(sheet, column).unwrap();
    writer
        .freeze_panes(
            sheet,
            FrozenPanes::new(Row::new(7).unwrap(), Column::new(5).unwrap()),
        )
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    bytes.set_position(0);
    let workbook = Workbook::new(bytes).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert_eq!(worksheet.row_layout(40).unwrap().height_twips(), 360);
    assert_eq!(worksheet.column_layout(12).unwrap().width_256ths(), 5120);
    assert!(!worksheet.row_layout(40).unwrap().is_hidden());
    assert!(!worksheet.column_layout(12).unwrap().is_hidden());
    assert!(
        workbook
            .xls_worksheet(0)
            .unwrap()
            .worksheet_view()
            .unwrap()
            .has_frozen_panes()
    );
}

#[test]
fn out_of_grid_locations_are_rejected_before_writer_mutation() {
    let result = std::panic::catch_unwind(|| {
        assert!(Row::new(65_536).is_err());
        assert!(Column::new(256).is_err());
    });
    assert!(result.is_ok());
}
