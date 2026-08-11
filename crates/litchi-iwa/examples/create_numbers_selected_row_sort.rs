//! Create a Numbers table and execute a selected-row sort without an input file.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRowRange, NumbersTableSortRule,
};
use litchi_numbers::table::CellPosition;
use litchi_numbers::table::cells::{Change, Input};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_selected_row_sort <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Selected Rows")
        .table_dimensions(6, 3)
        .build()?;
    let table = TableSelector::index(0);
    editor = set_focused_table_cells(
        editor,
        [
            Change::set(CellPosition::new(0, 0), Input::text("Region")?),
            Change::set(CellPosition::new(0, 1), Input::text("Q1")?),
            Change::set(CellPosition::new(0, 2), Input::text("Q2")?),
            Change::set(CellPosition::new(1, 0), Input::text("Outside")?),
            Change::set(CellPosition::new(1, 1), Input::number(50.0)?),
            Change::set(CellPosition::new(1, 2), Input::number(75.0)?),
            Change::set(CellPosition::new(2, 0), Input::text("South")?),
            Change::set(CellPosition::new(2, 1), Input::number(98.0)?),
            Change::set(CellPosition::new(2, 2), Input::number(132.0)?),
            Change::set(CellPosition::new(3, 0), Input::text("Central")?),
            Change::set(CellPosition::new(3, 1), Input::number(105.0)?),
            Change::set(CellPosition::new(3, 2), Input::number(139.0)?),
            Change::set(CellPosition::new(4, 0), Input::text("North")?),
            Change::set(CellPosition::new(4, 1), Input::number(120.0)?),
            Change::set(CellPosition::new(4, 2), Input::number(145.0)?),
            Change::set(CellPosition::new(5, 0), Input::text("Total")?),
            Change::set(CellPosition::new(5, 1), Input::number(373.0)?),
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
    editor.set_table_sort_order(
        table,
        NumbersTableSortOrder::selected_rows([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(1)?,
            NumbersTableSortDirection::Descending,
        )])?,
    )?;
    let selected_rows = NumbersTableSortRowRange::new(1, 4)?;
    if !editor.apply_table_sort_order_to_rows(table, selected_rows)? {
        return Err("expected the selected rows to be reordered".into());
    }
    editor.save(output)?;
    Ok(())
}

fn set_focused_table_headers(
    editor: NumbersEditor,
    settings: HeaderSettings,
) -> Result<NumbersEditor, Box<dyn std::error::Error>> {
    let package = Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_headers(SheetSelector::index(0), TableSelector::index(0))?
        .set(settings)
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(NumbersEditor::from_bytes(&bytes)?)
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
