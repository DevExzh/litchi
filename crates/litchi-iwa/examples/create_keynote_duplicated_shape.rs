//! Create a Keynote presentation with an independently editable duplicated shape.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::shape::path::Preset;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args()
        .nth(1)
        .ok_or("usage: create_keynote_duplicated_shape <output.key>")?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Independent duplicated shapes")
        .build()?;
    let source = editor.add_slide_shape(
        0,
        "Source shape",
        DrawablePoint { x: 720.0, y: 660.0 },
        DrawableSize {
            width: 480.0,
            height: 240.0,
        },
        Preset::RightArrow,
    )?;
    let duplicate = editor.duplicate_slide_shape(0, source.drawable_object_id)?;
    editor.set_slide_shape_text(0, duplicate.drawable_object_id, "Independent copy")?;
    editor.save(output)?;

    println!(
        "source={} clone={} source_storage={} clone_storage={}",
        source.drawable_object_id,
        duplicate.drawable_object_id,
        source.storage.id,
        duplicate.storage.id,
    );
    Ok(())
}
