//! Create a Numbers table with a hidden body row and a configured sort order.

use std::path::PathBuf;

use litchi_iwa::numbers::{
    CellValue, NumbersDocumentBuilder, NumbersTableHeaderCount, NumbersTableHeaderSettings,
    NumbersTableSortColumnIndex, NumbersTableSortDirection, NumbersTableSortOrder,
    NumbersTableSortRule, TableCellUpdate,
};
use litchi_iwa::table_hidden_axes::{TableAxisIndex, TableHiddenAxes};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_numbers_hidden_row_sort <output.numbers>")?,
    );
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Hidden Forecast")
        .table_dimensions(5, 2)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_table_header_settings(
        table_id,
        NumbersTableHeaderSettings {
            header_rows: Some(NumbersTableHeaderCount::ONE),
            footer_rows: Some(NumbersTableHeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_cells(
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Region".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Q1".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(1, 1, CellValue::Number(120.0)),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, CellValue::Number(98.0)),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::Number(105.0)),
            TableCellUpdate::new(4, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(4, 1, CellValue::Number(323.0)),
        ],
    )?;
    editor.set_table_hidden_axes(table_id, &TableHiddenAxes::new([TableAxisIndex::row(2)])?)?;
    editor.set_table_sort_order(
        table_id,
        NumbersTableSortOrder::new([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(1)?,
            NumbersTableSortDirection::Descending,
        )])?,
    )?;
    if !editor.apply_table_sort_order(table_id)? {
        return Err("expected the hidden-row table to be reordered".into());
    }
    let hidden = TableHiddenAxes::new([TableAxisIndex::row(2)])?;
    if editor.table_hidden_axes(table_id)? != hidden {
        return Err("sort did not preserve the hidden row position".into());
    }
    editor.save(output)?;
    Ok(())
}
