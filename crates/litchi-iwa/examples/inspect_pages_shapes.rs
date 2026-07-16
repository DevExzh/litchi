//! Inspect typed ordinary shapes anchored to a Pages document body.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_shapes <input.pages>")?;
    let editor = PagesEditor::open(input)?;
    for (index, shape) in editor.body_shapes()?.iter().enumerate() {
        println!(
            "shape[{index}] drawable={} storage={} anchor={} kind={:?} preset={:?} line={:?} text={:?} geometry={:?} properties={:?}",
            shape.drawable_object_id,
            shape.storage.object_id,
            shape.anchor_character_index,
            shape.kind,
            shape.preset,
            shape.line_segment,
            shape.storage.text,
            shape.geometry,
            shape.properties
        );
    }
    Ok(())
}
