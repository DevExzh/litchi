//! Delete one hidden row and column from a native Numbers table.

use std::path::PathBuf;

use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_numbers::table::topology::{ColumnDeletion, RowDeletion};
use litchi_numbers::TableSelector;

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
    let table = TableSelector::index(0);
    assert_eq!(editor.table_hidden_axes(table)?, initial);
    editor.remove_table_row(table, RowDeletion::body(1))?;
    editor.remove_table_column(table, ColumnDeletion::body(0))?;
    assert!(editor.table_hidden_axes(table)?.is_empty());
    editor.save(output)?;
    Ok(())
}
