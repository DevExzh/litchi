//! Create and physically sort a Numbers table without an input document.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use litchi_iwa::numbers::{
    NumbersDocumentBuilder, NumbersEditor, NumbersTableSortColumnIndex, NumbersTableSortDirection,
    NumbersTableSortOrder, NumbersTableSortRule,
};
use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};
use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_numbers_sorted <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Forecast")
        .table_dimensions(5, 3)
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
            TableCellUpdate::new(1, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(1, 1, CellValue::number(120.0)?),
            TableCellUpdate::new(1, 2, CellValue::number(145.0)?),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, CellValue::number(98.0)?),
            TableCellUpdate::new(2, 2, CellValue::number(132.0)?),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::number(105.0)?),
            TableCellUpdate::new(3, 2, CellValue::number(139.0)?),
            TableCellUpdate::new(4, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(4, 1, CellValue::number(323.0)?),
        ],
    )?;
    editor.set_cell_comment(table_id, 2, 1, "South comment follows its sorted row")?;
    let reply_id =
        editor.add_cell_comment_reply(table_id, 2, 1, "Numbers keeps this thread intact")?;
    editor.set_table_sort_order(
        table,
        NumbersTableSortOrder::new([NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(1)?,
            NumbersTableSortDirection::Ascending,
        )])?,
    )?;
    if !editor.apply_table_sort_order(table)? {
        return Err("expected the source table to be reordered".into());
    }
    let moved = editor
        .cell_comment(table_id, 1, 1)?
        .ok_or("sorted row lost its comment")?;
    let moved_reply_id = editor
        .cell_comment_replies(table_id, 1, 1)?
        .first()
        .map(|reply| reply.storage_id.get());
    if moved.comment.text != "South comment follows its sorted row"
        || moved_reply_id != Some(reply_id)
    {
        return Err("sorted row did not preserve its comment thread".into());
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
