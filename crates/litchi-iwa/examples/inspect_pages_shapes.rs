//! Inspect typed ordinary shapes anchored to a Pages document body.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_shapes <input.pages>")?;
    let editor = PagesEditor::open(input)?;
    for (index, shape) in editor.body_shapes()?.iter().enumerate() {
        let fill = editor.body_shape_fill(shape.drawable_object_id)?;
        let stroke = editor.body_shape_stroke(shape.drawable_object_id)?;
        let effects = editor.body_shape_effects(shape.drawable_object_id)?;
        let shadow = editor.body_shape_shadow(shape.drawable_object_id)?;
        let text_layout = editor.body_shape_text_layout(shape.drawable_object_id)?;
        let title_caption = editor.body_shape_title_caption(shape.drawable_object_id)?;
        println!(
            "shape[{index}] drawable={} storage={} anchor={} kind={:?} preset={:?} line={:?} endpoints={:?} fill={fill:?} stroke={stroke:?} effects={effects:?} shadow={shadow:?} text_layout={text_layout:?} title_caption={title_caption:?} text={:?} geometry={:?} properties={:?}",
            shape.drawable_object_id,
            shape.storage.object_id,
            shape.anchor_character_index,
            shape.kind,
            shape.preset,
            shape.line_segment,
            shape.line_endpoints,
            shape.storage.storage.text(),
            shape.geometry,
            shape.properties
        );
    }
    Ok(())
}
