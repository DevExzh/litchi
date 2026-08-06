//! Create a Numbers table and execute a selected-row sort without an input file.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRowRange, NumbersTableSortRule,
};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_selected_row_sort <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Selected Rows")
        .table_dimensions(6, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_table_header_settings(
        table_id,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_cells(
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Region".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Q1".to_owned())),
            TableCellUpdate::new(0, 2, CellValue::Text("Q2".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("Outside".to_owned())),
            TableCellUpdate::new(1, 1, CellValue::Number(50.0)),
            TableCellUpdate::new(1, 2, CellValue::Number(75.0)),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, CellValue::Number(98.0)),
            TableCellUpdate::new(2, 2, CellValue::Number(132.0)),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::Number(105.0)),
            TableCellUpdate::new(3, 2, CellValue::Number(139.0)),
            TableCellUpdate::new(4, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(4, 1, CellValue::Number(120.0)),
            TableCellUpdate::new(4, 2, CellValue::Number(145.0)),
            TableCellUpdate::new(5, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(5, 1, CellValue::Number(373.0)),
        ],
    )?;
    editor.set_table_sort_order(
        table_id,
        NumbersTableSortOrder::selected_rows([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(1)?,
            NumbersTableSortDirection::Descending,
        )])?,
    )?;
    let selected_rows = NumbersTableSortRowRange::new(1, 4)?;
    if !editor.apply_table_sort_order_to_rows(table_id, selected_rows)? {
        return Err("expected the selected rows to be reordered".into());
    }
    editor.save(output)?;
    Ok(())
}
