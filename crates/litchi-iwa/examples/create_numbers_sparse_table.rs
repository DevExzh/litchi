//! Create a source-built Numbers workbook with cells across sparse tile boundaries.

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_numbers::table::CellPosition;
use litchi_numbers::table::cells::{Change, Input};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_sparse_table <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Sparse inventory")
        .table_dimensions(257, 2)
        .build()?;
    editor = set_focused_table_cells(
        editor,
        [Change::set(
            CellPosition::new(256, 1),
            Input::number(257.0)?,
        )],
    )?;
    editor = set_focused_table_cells(
        editor,
        [
            Change::set(CellPosition::new(0, 0), Input::text("First tile")?),
            Change::set(CellPosition::new(256, 0), Input::text("Boundary")?),
        ],
    )?;
    editor.save(output)?;
    Ok(())
}

fn set_focused_table_cells(
    editor: NumbersEditor,
    changes: impl IntoIterator<Item = Change>,
) -> Result<NumbersEditor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_cells(SheetSelector::index(0), TableSelector::index(0))?
        .extend(changes)?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(NumbersEditor::from_bytes(&bytes)?)
}
