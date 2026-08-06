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
            let fill = editor.slide_shape_fill(slide.index, shape.drawable_object_id)?;
            let stroke = editor.slide_shape_stroke(slide.index, shape.drawable_object_id)?;
            let effects = editor.slide_shape_effects(slide.index, shape.drawable_object_id)?;
            let shadow = editor.slide_shape_shadow(slide.index, shape.drawable_object_id)?;
            let text_layout =
                editor.slide_shape_text_layout(slide.index, shape.drawable_object_id)?;
            let title_caption =
                editor.slide_shape_title_caption(slide.index, shape.drawable_object_id)?;
            println!(
                "slide={} shape_index={shape_index} drawable={} kind={:?} preset={:?} line={:?} endpoints={:?} fill={fill:?} stroke={stroke:?} effects={effects:?} shadow={shadow:?} text_layout={text_layout:?} title_caption={title_caption:?} storage={} text={:?} geometry={:?} properties={:?}",
                slide.index,
                shape.drawable_object_id,
                shape.kind,
                shape.preset,
                shape.line_segment,
                shape.line_endpoints,
                shape.storage.id,
                shape.storage.storage.text(),
                shape.geometry,
                shape.properties,
            );
        }
    }
    Ok(())
}
