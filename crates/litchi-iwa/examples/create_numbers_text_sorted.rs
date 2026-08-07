//! Create and physically sort a plain-text Numbers table without an input document.

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRule,
};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_text_sorted <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Cities")
        .table_dimensions(5, 2)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cells(
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Name".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Marker".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("zebra".to_owned())),
            TableCellUpdate::new(1, 1, CellValue::Text("last".to_owned())),
            TableCellUpdate::new(2, 0, CellValue::Text("apple".to_owned())),
            TableCellUpdate::new(2, 1, CellValue::Text("first apple".to_owned())),
            TableCellUpdate::new(3, 0, CellValue::Text("banana".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::Text("middle".to_owned())),
            TableCellUpdate::new(4, 0, CellValue::Text("apple".to_owned())),
            TableCellUpdate::new(4, 1, CellValue::Text("second apple".to_owned())),
        ],
    )?;
    editor.set_table_sort_order(
        table_id,
        NumbersTableSortOrder::new([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(0)?,
            NumbersTableSortDirection::Ascending,
        )])?,
    )?;
    if !editor.apply_table_sort_order(table_id)? {
        return Err("expected the source table to be reordered".into());
    }
    editor.save(output)?;
    Ok(())
}
