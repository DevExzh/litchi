//! List ordinary shapes owned directly by Numbers sheets.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_numbers_shapes <input.numbers>")?;
    let editor = NumbersEditor::open(input)?;
    for sheet in editor.sheets()? {
        for (shape_index, shape) in editor
            .sheet_shapes(sheet.object_id)?
            .into_iter()
            .enumerate()
        {
            let fill = editor.sheet_shape_fill(sheet.object_id, shape.drawable_object_id)?;
            let stroke = editor.sheet_shape_stroke(sheet.object_id, shape.drawable_object_id)?;
            let effects = editor.sheet_shape_effects(sheet.object_id, shape.drawable_object_id)?;
            let shadow = editor.sheet_shape_shadow(sheet.object_id, shape.drawable_object_id)?;
            let text_layout =
                editor.sheet_shape_text_layout(sheet.object_id, shape.drawable_object_id)?;
            let title_caption =
                editor.sheet_shape_title_caption(sheet.object_id, shape.drawable_object_id)?;
            println!(
                "sheet={} shape_index={shape_index} drawable={} kind={:?} preset={:?} line={:?} endpoints={:?} fill={fill:?} stroke={stroke:?} effects={effects:?} shadow={shadow:?} text_layout={text_layout:?} title_caption={title_caption:?} storage={} text={:?} geometry={:?} properties={:?}",
                sheet.object_id,
                shape.drawable_object_id,
                shape.kind,
                shape.preset,
                shape.line_segment,
                shape.line_endpoints,
                shape.storage.object_id,
                shape.storage.storage.text(),
                shape.geometry,
                shape.properties,
            );
        }
    }
    Ok(())
}
