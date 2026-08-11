//! Create a Numbers file with native hidden table axes.

use std::path::PathBuf;

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_numbers::TableSelector;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_numbers_table_hidden_axes <output.numbers>")?,
    );
    let hidden = HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)])?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Hidden Axes")
        .table_dimensions(6, 4)
        .build()?;
    editor
        .tables()?
        .first()
        .ok_or("the document has no tables")?;
    let table = TableSelector::index(0);
    editor.set_table_hidden_axes(table, &hidden)?;
    editor.save(&output)?;

    assert_eq!(
        NumbersEditor::open(output)?.table_hidden_axes(table)?,
        hidden
    );
    Ok(())
}
