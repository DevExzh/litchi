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
            println!(
                "sheet={} shape_index={shape_index} drawable={} kind={:?} storage={} text={:?} geometry={:?} properties={:?}",
                sheet.object_id,
                shape.drawable_object_id,
                shape.kind,
                shape.storage.object_id,
                shape.storage.text,
                shape.geometry,
                shape.properties,
            );
        }
    }
    Ok(())
}
