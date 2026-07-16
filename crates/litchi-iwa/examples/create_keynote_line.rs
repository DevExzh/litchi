//! Create a Keynote presentation and editable straight line without an input package.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::DrawablePoint;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_line <output.key>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Line built from typed IWA objects")
        .build()?;
    let created = editor.add_slide_line(
        0,
        DrawablePoint { x: 720.0, y: 660.0 },
        DrawablePoint {
            x: 1_200.0,
            y: 900.0,
        },
    )?;
    editor.set_slide_line_segment(
        0,
        created.drawable_object_id,
        DrawablePoint { x: 96.0, y: 108.0 },
        DrawablePoint { x: 456.0, y: 108.0 },
    )?;
    let segment = editor.slide_line_segment(0, created.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created Keynote line {} from {:?} to {:?}",
        created.drawable_object_id,
        segment.start(),
        segment.end()
    );
    Ok(())
}
