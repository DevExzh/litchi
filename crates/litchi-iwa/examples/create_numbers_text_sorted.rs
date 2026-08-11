//! Create and physically sort a plain-text Numbers table without an input document.

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRule,
};
use litchi_numbers::table::CellPosition;
use litchi_numbers::table::cells::{Change, Input};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_text_sorted <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Cities")
        .table_dimensions(5, 2)
        .build()?;
    let table = TableSelector::index(0);
    editor = set_focused_table_cells(
        editor,
        [
            Change::set(CellPosition::new(0, 0), Input::text("Name")?),
            Change::set(CellPosition::new(0, 1), Input::text("Marker")?),
            Change::set(CellPosition::new(1, 0), Input::text("zebra")?),
            Change::set(CellPosition::new(1, 1), Input::text("last")?),
            Change::set(CellPosition::new(2, 0), Input::text("apple")?),
            Change::set(CellPosition::new(2, 1), Input::text("first apple")?),
            Change::set(CellPosition::new(3, 0), Input::text("banana")?),
            Change::set(CellPosition::new(3, 1), Input::text("middle")?),
            Change::set(CellPosition::new(4, 0), Input::text("apple")?),
            Change::set(CellPosition::new(4, 1), Input::text("second apple")?),
        ],
    )?;
    editor.set_table_sort_order(
        table,
        NumbersTableSortOrder::new([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(0)?,
            NumbersTableSortDirection::Ascending,
        )])?,
    )?;
    if !editor.apply_table_sort_order(table)? {
        return Err("expected the source table to be reordered".into());
    }
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
