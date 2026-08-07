//! Insert visible axes before hidden axes in a native Numbers table.

use std::path::PathBuf;

use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_numbers::TableSelector;
use litchi_numbers::table::topology::{ColumnInsertion, RowInsertion};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let source =
        PathBuf::from(arguments.next().ok_or(
            "usage: edit_numbers_hidden_table_topology <source.numbers> <output.numbers>",
        )?);
    let output =
        PathBuf::from(arguments.next().ok_or(
            "usage: edit_numbers_hidden_table_topology <source.numbers> <output.numbers>",
        )?);
    let initial = HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)])?;
    let shifted = HiddenAxes::new([AxisIndex::row(3), AxisIndex::column(2)])?;

    let mut editor = NumbersEditor::open(source)?;
    let table = TableSelector::index(0);
    assert_eq!(editor.table_hidden_axes(table)?, initial);
    editor.insert_table_row(table, RowInsertion::body(0))?;
    editor.insert_table_column(table, ColumnInsertion::body(0))?;
    assert_eq!(editor.table_hidden_axes(table)?, shifted);
    editor.save(output)?;
    Ok(())
}
