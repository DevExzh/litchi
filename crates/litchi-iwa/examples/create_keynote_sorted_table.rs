//! Create and physically sort a plain-text Keynote table without an input presentation.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use std::env;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteTableCellUpdate, KeynoteTableCellValue};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_numbers::table::sort::{ColumnIndex, Direction, Order, Rule};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .ok_or("usage: create_keynote_sorted_table <output.key>")?;
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Cities sorted by name")
        .subtitle("Created entirely by litchi-iwa")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Cities",
        5,
        2,
        DrawablePoint { x: 360.0, y: 380.0 },
        DrawableSize {
            width: 1_160.0,
            height: 470.0,
        },
    )?;
    editor.set_slide_table_header_settings(
        0,
        table.model_object_id,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )?;
    editor.set_slide_table_cells(
        0,
        table.model_object_id,
        [
            KeynoteTableCellUpdate::new(0, 0, KeynoteTableCellValue::Text("Name".to_owned())),
            KeynoteTableCellUpdate::new(0, 1, KeynoteTableCellValue::Text("Marker".to_owned())),
            KeynoteTableCellUpdate::new(1, 0, KeynoteTableCellValue::Text("zebra".to_owned())),
            KeynoteTableCellUpdate::new(1, 1, KeynoteTableCellValue::Text("last".to_owned())),
            KeynoteTableCellUpdate::new(2, 0, KeynoteTableCellValue::Text("apple".to_owned())),
            KeynoteTableCellUpdate::new(
                2,
                1,
                KeynoteTableCellValue::Text("first apple".to_owned()),
            ),
            KeynoteTableCellUpdate::new(3, 0, KeynoteTableCellValue::Text("banana".to_owned())),
            KeynoteTableCellUpdate::new(3, 1, KeynoteTableCellValue::Text("middle".to_owned())),
            KeynoteTableCellUpdate::new(4, 0, KeynoteTableCellValue::Text("apple".to_owned())),
            KeynoteTableCellUpdate::new(
                4,
                1,
                KeynoteTableCellValue::Text("second apple".to_owned()),
            ),
        ],
    )?;
    editor.set_slide_table_cell_comment(
        0,
        table.model_object_id,
        1,
        1,
        "Zebra comment follows its sorted row",
    )?;
    let reply_id = editor.add_slide_table_cell_comment_reply(
        0,
        table.model_object_id,
        1,
        1,
        "Keynote keeps this thread intact",
    )?;
    editor.set_slide_table_sort_order(
        0,
        table.model_object_id,
        Order::new([Rule::new(ColumnIndex::new(0)?, Direction::Ascending)])?,
    )?;
    if !editor.apply_slide_table_sort_order(0, table.model_object_id)? {
        return Err("expected the source table to be reordered".into());
    }
    let moved = editor
        .slide_table_cell_comment(0, table.model_object_id, 4, 1)?
        .ok_or("sorted row lost its comment")?;
    let moved_reply_id = editor
        .slide_table_cell_comment_replies(0, table.model_object_id, 4, 1)?
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
