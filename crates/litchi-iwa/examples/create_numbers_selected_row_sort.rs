//! Create a Numbers table and execute a selected-row sort without an input file.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRowRange, NumbersTableSortRule,
};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_selected_row_sort <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Selected Rows")
        .table_dimensions(6, 3)
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
            TableCellUpdate::new(0, 2, CellValue::Text("Q2".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("Outside".to_owned())),
            TableCellUpdate::new(1, 1, CellValue::number(50.0)?),
            TableCellUpdate::new(1, 2, CellValue::number(75.0)?),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, CellValue::number(98.0)?),
            TableCellUpdate::new(2, 2, CellValue::number(132.0)?),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::number(105.0)?),
            TableCellUpdate::new(3, 2, CellValue::number(139.0)?),
            TableCellUpdate::new(4, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(4, 1, CellValue::number(120.0)?),
            TableCellUpdate::new(4, 2, CellValue::number(145.0)?),
            TableCellUpdate::new(5, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(5, 1, CellValue::number(373.0)?),
        ],
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
