//! Create a Numbers spreadsheet and text box without an input package.

use std::env;

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_text_box <output.numbers> [text]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Scratch Sheet")
        .table_name("Scratch Table")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let created = editor.add_sheet_text_box(
        sheet_id,
        &text,
        DrawablePoint { x: 40.0, y: 300.0 },
        DrawableSize {
            width: 300.0,
            height: 72.0,
        },
    )?;
    editor.save(output)?;
    println!(
        "created Numbers text box {} with storage {} on sheet {}",
        created.drawable_object_id, created.storage.object_id, sheet_id
    );
    Ok(())
}
