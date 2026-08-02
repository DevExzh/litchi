//! Create a Pages document with an independently editable duplicated shape.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapePreset};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .ok_or("usage: create_pages_duplicated_shape <output.pages>")?;

    let body = "Created entirely by litchi-iwa";
    let mut editor = PagesEditor::create_with_text(body)?;
    let source = editor.add_body_shape(
        body.encode_utf16().count(),
        "Source shape",
        DrawablePoint { x: 180.0, y: 240.0 },
        DrawableSize {
            width: 300.0,
            height: 150.0,
        },
        ShapePreset::RightArrow,
    )?;
    let anchor = editor.body_text()?.encode_utf16().count();
    let duplicate = editor.duplicate_body_shape(source.drawable_object_id, anchor)?;
    editor.set_body_shape_text(duplicate.drawable_object_id, "Independent copy")?;
    editor.save(output)?;

    println!(
        "source={} clone={} source_storage={} clone_storage={}",
        source.drawable_object_id,
        duplicate.drawable_object_id,
        source.storage.object_id,
        duplicate.storage.object_id,
    );
    Ok(())
}
