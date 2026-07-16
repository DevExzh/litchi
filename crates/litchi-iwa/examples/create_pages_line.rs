//! Create a Pages document and editable straight line without an input package.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, LineEndpoint, LineEndpoints};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_line <output.pages>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let body = "Pages line created entirely by litchi-iwa";
    let mut editor = PagesEditor::create_with_text(body)?;
    let created = editor.add_body_line_with_endpoints(
        body.encode_utf16().count(),
        DrawablePoint { x: 180.0, y: 240.0 },
        DrawablePoint { x: 480.0, y: 390.0 },
        LineEndpoints::new(LineEndpoint::OpenCircle, LineEndpoint::FilledArrow),
    )?;
    editor.set_body_line_segment(
        created.drawable_object_id,
        DrawablePoint { x: 96.0, y: 180.0 },
        DrawablePoint { x: 456.0, y: 180.0 },
    )?;
    let segment = editor.body_line_segment(created.drawable_object_id)?;
    let endpoints = editor.body_line_endpoints(created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Pages line {} from {:?} to {:?}, endpoints {endpoints:?}, at UTF-16 anchor {}",
        created.drawable_object_id,
        segment.start(),
        segment.end(),
        created.anchor_character_index
    );
    Ok(())
}
