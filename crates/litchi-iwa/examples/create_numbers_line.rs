//! Create a Numbers spreadsheet and editable straight line without an input package.

use std::env;

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, LineEndpoint, LineEndpoints};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_line <output.numbers>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Scratch Line")
        .table_name("Source Data")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let created = editor.add_sheet_line_with_endpoints(
        sheet_id,
        DrawablePoint { x: 420.0, y: 300.0 },
        DrawablePoint { x: 720.0, y: 450.0 },
        LineEndpoints::new(LineEndpoint::FilledCircle, LineEndpoint::SimpleArrow),
    )?;
    editor.set_sheet_line_segment(
        sheet_id,
        created.drawable_object_id,
        DrawablePoint { x: 72.0, y: 180.0 },
        DrawablePoint { x: 432.0, y: 180.0 },
    )?;
    let segment = editor.sheet_line_segment(sheet_id, created.drawable_object_id)?;
    let endpoints = editor.sheet_line_endpoints(sheet_id, created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Numbers line {} from {:?} to {:?}, endpoints {endpoints:?}, on sheet {}",
        created.drawable_object_id,
        segment.start(),
        segment.end(),
        sheet_id
    );
    Ok(())
}
