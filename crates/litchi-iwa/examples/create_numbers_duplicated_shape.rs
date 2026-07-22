//! Create a Numbers spreadsheet with an independently editable duplicated shape.

use std::env;

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapePreset};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .ok_or("usage: create_numbers_duplicated_shape <output.numbers>")?;

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Duplicated shapes")
        .table_name("Data")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let source = editor.add_sheet_shape(
        sheet_id,
        "Source shape",
        DrawablePoint { x: 420.0, y: 300.0 },
        DrawableSize {
            width: 300.0,
            height: 150.0,
        },
        ShapePreset::RightArrow,
    )?;
    let duplicate = editor.duplicate_sheet_shape(sheet_id, source.drawable_object_id)?;
    editor.set_sheet_shape_text(sheet_id, duplicate.drawable_object_id, "Independent copy")?;
    editor.save(output)?;

    println!(
        "sheet={sheet_id} source={} clone={} source_storage={} clone_storage={}",
        source.drawable_object_id,
        duplicate.drawable_object_id,
        source.storage.object_id,
        duplicate.storage.object_id,
    );
    Ok(())
}
