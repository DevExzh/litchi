//! Create a Numbers table and execute a selected-row sort without an input file.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use litchi_iwa::numbers::{self};
use litchi_numbers::table::At;
use litchi_numbers::table::cells::{Change, Input};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_selected_row_sort <output.numbers>")?;
    let mut editor = numbers::Builder::new()
        .table_name("Selected Rows")
        .table_dimensions(6, 3)
        .build()?;
    let table = TableSelector::index(0);
    editor = set_focused_table_cells(
        editor,
        [
            Change::set(At::new(0, 0), Input::text("Region")?),
            Change::set(At::new(0, 1), Input::text("Q1")?),
            Change::set(At::new(0, 2), Input::text("Q2")?),
            Change::set(At::new(1, 0), Input::text("Outside")?),
            Change::set(At::new(1, 1), Input::number(50.0)?),
            Change::set(At::new(1, 2), Input::number(75.0)?),
            Change::set(At::new(2, 0), Input::text("South")?),
            Change::set(At::new(2, 1), Input::number(98.0)?),
            Change::set(At::new(2, 2), Input::number(132.0)?),
            Change::set(At::new(3, 0), Input::text("Central")?),
            Change::set(At::new(3, 1), Input::number(105.0)?),
            Change::set(At::new(3, 2), Input::number(139.0)?),
            Change::set(At::new(4, 0), Input::text("North")?),
            Change::set(At::new(4, 1), Input::number(120.0)?),
            Change::set(At::new(4, 2), Input::number(145.0)?),
            Change::set(At::new(5, 0), Input::text("Total")?),
            Change::set(At::new(5, 1), Input::number(373.0)?),
        ],
    )?;
    editor = set_focused_table_headers(
        editor,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    let order = numbers::TableSortOrder::selected_rows([numbers::TableSortRule::new(
        numbers::TableSortColumnIndex::new(1)?,
        numbers::TableSortDirection::Descending,
    )])?;
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_sort_order(SheetSelector::index(0), table)?
        .set(order)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    editor = numbers::Editor::from_bytes(&bytes)?;
    let selected_rows = numbers::TableSortRowRange::new(1, 4)?;
    if !editor.apply_table_sort_order_to_rows(table, selected_rows)? {
        return Err("expected the selected rows to be reordered".into());
    }
    editor.save(output)?;
    Ok(())
}

fn set_focused_table_headers(
    editor: numbers::Editor,
    settings: HeaderSettings,
) -> Result<numbers::Editor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_headers(SheetSelector::index(0), TableSelector::index(0))?
        .set(settings)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(numbers::Editor::from_bytes(&bytes)?)
}

fn set_focused_table_cells(
    editor: numbers::Editor,
    changes: impl IntoIterator<Item = Change>,
) -> Result<numbers::Editor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_cells(SheetSelector::index(0), TableSelector::index(0))?
        .extend(changes)?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(numbers::Editor::from_bytes(&bytes)?)
}
