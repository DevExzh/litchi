//! Create a Numbers table with a hidden body row and a configured sort order.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use std::path::PathBuf;

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRule,
};
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};
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
    let table_id = editor.tables()?.remove(0).id();
    let table = TableSelector::index(0);
    editor = set_focused_table_headers(
        editor,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_cells(
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Region".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Q1".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(1, 1, CellValue::number(120.0)?),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, CellValue::number(98.0)?),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::number(105.0)?),
            TableCellUpdate::new(4, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(4, 1, CellValue::number(323.0)?),
        ],
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
