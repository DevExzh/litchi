//! Create a source-built Numbers workbook with cells across sparse tile boundaries.

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_sparse_table <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Sparse inventory")
        .table_dimensions(257, 2)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cells(
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("First tile".to_owned())),
            TableCellUpdate::new(256, 0, CellValue::Text("Boundary".to_owned())),
            TableCellUpdate::new(256, 1, CellValue::Number(257.0)),
        ],
    )?;
    editor.save(output)?;
    Ok(())
}
