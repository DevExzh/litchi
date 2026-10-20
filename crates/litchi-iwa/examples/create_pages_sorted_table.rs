//! Create and physically sort a plain-text Pages table without an input document.
use std::env;

use litchi_iwa::pages::{self};
use litchi_pages::table::headers::{Count as HeaderCount, Settings as HeaderSettings};
use litchi_pages::table::sort::{ColumnIndex, Direction, Order, Rule};
use litchi_pages::{BodyTableSelector, Package};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .ok_or("usage: create_pages_sorted_table <output.pages>")?;
    let mut editor = pages::Builder::new()
        .body_text("Cities sorted by name\n")
        .body_table("Cities", 5, 2)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    let focused = Package::from_bytes(&editor.to_bytes()?)?;
    let header_commit = focused
        .edit_body_table_header_settings(BodyTableSelector::index(0))?
        .set(HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            ..Default::default()
        })
        .commit()?;
    let mut header_bytes = Vec::new();
    header_commit.package().write_to(&mut header_bytes)?;
    editor = litchi_iwa::pages::Editor::from_bytes(&header_bytes)?;
    editor.set_table_cells(
        table_id,
        [
            pages::TableCellUpdate::new(0, 0, pages::CellValue::Text("Name".to_owned())),
            pages::TableCellUpdate::new(0, 1, pages::CellValue::Text("Marker".to_owned())),
            pages::TableCellUpdate::new(1, 0, pages::CellValue::Text("zebra".to_owned())),
            pages::TableCellUpdate::new(1, 1, pages::CellValue::Text("last".to_owned())),
            pages::TableCellUpdate::new(2, 0, pages::CellValue::Text("apple".to_owned())),
            pages::TableCellUpdate::new(2, 1, pages::CellValue::Text("first apple".to_owned())),
            pages::TableCellUpdate::new(3, 0, pages::CellValue::Text("banana".to_owned())),
            pages::TableCellUpdate::new(3, 1, pages::CellValue::Text("middle".to_owned())),
            pages::TableCellUpdate::new(4, 0, pages::CellValue::Text("apple".to_owned())),
            pages::TableCellUpdate::new(4, 1, pages::CellValue::Text("second apple".to_owned())),
        ],
    )?;
    editor.set_table_cell_comment(table_id, 1, 1, "Zebra comment follows its sorted row")?;
    let reply_id =
        editor.add_table_cell_comment_reply(table_id, 1, 1, "Pages keeps this thread intact")?;
    let sort_commit = Package::from_bytes(&editor.to_bytes()?)?
        .edit_body_table_sort_order(BodyTableSelector::index(0))?
        .set(Order::new([Rule::new(
            ColumnIndex::new(0)?,
            Direction::Ascending,
        )])?)
        .commit()?;
    let mut sort_bytes = Vec::new();
    sort_commit.package().write_to(&mut sort_bytes)?;
    editor = litchi_iwa::pages::Editor::from_bytes(&sort_bytes)?;
    if !editor.apply_table_sort_order(table_id)? {
        return Err("expected the source table to be reordered".into());
    }
    let moved = editor
        .table_cell_comment(table_id, 4, 1)?
        .ok_or("sorted row lost its comment")?;
    let moved_reply_id = editor
        .table_cell_comment_replies(table_id, 4, 1)?
        .first()
        .map(|reply| reply.storage_id.get());
    if moved.comment.text != "Zebra comment follows its sorted row"
        || moved_reply_id != Some(reply_id)
    {
        return Err("sorted row did not preserve its comment thread".into());
    }
    editor.save(&output)?;
    println!("created {output}");
    Ok(())
}
