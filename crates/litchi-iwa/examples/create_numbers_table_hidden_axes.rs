//! Create a Numbers file with native hidden table axes.

use std::path::PathBuf;

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::table_hidden_axes::{TableAxisIndex, TableHiddenAxes};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_numbers_table_hidden_axes <output.numbers>")?,
    );
    let hidden = TableHiddenAxes::new([TableAxisIndex::row(2), TableAxisIndex::column(1)])?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Hidden Axes")
        .table_dimensions(6, 4)
        .build()?;
    let table = editor.tables()?.remove(0);
    editor.set_table_hidden_axes(table.object_id, &hidden)?;
    editor.save(&output)?;

    assert_eq!(
        NumbersEditor::open(output)?.table_hidden_axes(table.object_id)?,
        hidden
    );
    Ok(())
}
