//! Create a Numbers table with a hidden body row and a configured sort order.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use std::path::PathBuf;

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRule,
};
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_numbers::table::CellPosition;
use litchi_numbers::table::cells::{Change, Input};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_numbers_hidden_row_sort <output.numbers>")?,
    );
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Hidden Forecast")
        .table_dimensions(5, 2)
        .build()?;
    let table = TableSelector::index(0);
    editor = set_focused_table_cells(
        editor,
        [
            Change::set(CellPosition::new(0, 0), Input::text("Region")?),
            Change::set(CellPosition::new(0, 1), Input::text("Q1")?),
            Change::set(CellPosition::new(1, 0), Input::text("North")?),
            Change::set(CellPosition::new(1, 1), Input::number(120.0)?),
            Change::set(CellPosition::new(2, 0), Input::text("South")?),
            Change::set(CellPosition::new(2, 1), Input::number(98.0)?),
            Change::set(CellPosition::new(3, 0), Input::text("Central")?),
            Change::set(CellPosition::new(3, 1), Input::number(105.0)?),
            Change::set(CellPosition::new(4, 0), Input::text("Total")?),
            Change::set(CellPosition::new(4, 1), Input::number(323.0)?),
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
    editor.set_table_hidden_axes(table, &HiddenAxes::new([AxisIndex::row(2)])?)?;
    editor.set_table_sort_order(
        table,
        NumbersTableSortOrder::new([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(1)?,
            NumbersTableSortDirection::Descending,
        )])?,
    )?;
    if !editor.apply_table_sort_order(table)? {
        return Err("expected the hidden-row table to be reordered".into());
    }
    let hidden = HiddenAxes::new([AxisIndex::row(2)])?;
    if editor.table_hidden_axes(table)? != hidden {
        return Err("sort did not preserve the hidden row position".into());
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
