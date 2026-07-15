//! Create a Numbers spreadsheet and editable preset shape without an input package.

use std::env;

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapePreset};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_shape <output.numbers> [text]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Scratch Shape")
        .table_name("Scratch Table")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let created = editor.add_sheet_shape(
        sheet_id,
        &text,
        DrawablePoint { x: 420.0, y: 300.0 },
        DrawableSize {
            width: 300.0,
            height: 150.0,
        },
        ShapePreset::ROUNDED_RECTANGLE,
    )?;
    editor.save(output)?;
    println!(
        "created Numbers {:?} {:?} {} with storage {} on sheet {}",
        created.kind,
        created.preset,
        created.drawable_object_id,
        created.storage.object_id,
        sheet_id
    );
    Ok(())
}
