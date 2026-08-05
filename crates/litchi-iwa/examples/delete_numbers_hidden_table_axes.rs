//! Delete one hidden row and column from a native Numbers table.

use std::path::PathBuf;

use litchi_iwa::numbers::{NumbersEditor, TableColumnDeletion, TableRowDeletion};
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let source = PathBuf::from(
        arguments
            .next()
            .ok_or("usage: delete_numbers_hidden_table_axes <source.numbers> <output.numbers>")?,
    );
    let output = PathBuf::from(
        arguments
            .next()
            .ok_or("usage: delete_numbers_hidden_table_axes <source.numbers> <output.numbers>")?,
    );
    let initial = HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)])?;

    let mut editor = NumbersEditor::open(source)?;
    let table = editor.tables()?.remove(0);
    assert_eq!(editor.table_hidden_axes(table.object_id)?, initial);
    editor.remove_table_row(table.object_id, TableRowDeletion::body(1))?;
    editor.remove_table_column(table.object_id, TableColumnDeletion::body(0))?;
    assert!(editor.table_hidden_axes(table.object_id)?.is_empty());
    editor.save(output)?;
    Ok(())
}
