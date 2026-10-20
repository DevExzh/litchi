//! Create and physically sort a plain-text Numbers table without an input document.

use litchi_iwa::numbers::{self};
use litchi_numbers::table::At;
use litchi_numbers::table::cells::{Change, Input};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_text_sorted <output.numbers>")?;
    let mut editor = numbers::Builder::new()
        .table_name("Cities")
        .table_dimensions(5, 2)
        .build()?;
    let table = TableSelector::index(0);
    editor = set_focused_table_cells(
        editor,
        [
            Change::set(At::new(0, 0), Input::text("Name")?),
            Change::set(At::new(0, 1), Input::text("Marker")?),
            Change::set(At::new(1, 0), Input::text("zebra")?),
            Change::set(At::new(1, 1), Input::text("last")?),
            Change::set(At::new(2, 0), Input::text("apple")?),
            Change::set(At::new(2, 1), Input::text("first apple")?),
            Change::set(At::new(3, 0), Input::text("banana")?),
            Change::set(At::new(3, 1), Input::text("middle")?),
            Change::set(At::new(4, 0), Input::text("apple")?),
            Change::set(At::new(4, 1), Input::text("second apple")?),
        ],
    )?;
    let order = numbers::TableSortOrder::new([numbers::TableSortRule::new(
        numbers::TableSortColumnIndex::new(0)?,
        numbers::TableSortDirection::Ascending,
    )])?;
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_sort_order(SheetSelector::index(0), table)?
        .set(order)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    editor = numbers::Draft::from_bytes(&bytes)?;
    if !editor.apply_table_sort_order(table)? {
        return Err("expected the source table to be reordered".into());
    }
    editor.save(output)?;
    Ok(())
}

fn set_focused_table_cells(
    editor: numbers::Draft,
    changes: impl IntoIterator<Item = Change>,
) -> Result<numbers::Draft, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_cells(SheetSelector::index(0), TableSelector::index(0))?
        .extend(changes)?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(numbers::Draft::from_bytes(&bytes)?)
}
