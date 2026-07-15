//! List ordinary shapes owned directly by Keynote slides.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_keynote_shapes <input.key>")?;
    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        for (shape_index, shape) in editor.slide_shapes(slide.index)?.into_iter().enumerate() {
            println!(
                "slide={} shape_index={shape_index} drawable={} kind={:?} preset={:?} storage={} text={:?} geometry={:?} properties={:?}",
                slide.index,
                shape.drawable_object_id,
                shape.kind,
                shape.preset,
                shape.storage.object_id,
                shape.storage.text,
                shape.geometry,
                shape.properties,
            );
        }
    }
    Ok(())
}
