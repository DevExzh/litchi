//! List ordinary text boxes owned by each Numbers sheet.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_numbers_text_boxes <input.numbers>")?;
    let editor = NumbersEditor::open(input)?;
    for sheet in editor.sheets()? {
        println!(
            "sheet={} object={} name={:?}",
            sheet.index, sheet.object_id, sheet.name
        );
        for (index, text_box) in editor
            .sheet_text_boxes(sheet.object_id)?
            .into_iter()
            .enumerate()
        {
            let geometry =
                editor.sheet_text_box_geometry(sheet.object_id, text_box.drawable_object_id)?;
            let properties =
                editor.sheet_text_box_properties(sheet.object_id, text_box.drawable_object_id)?;
            let comment =
                editor.sheet_drawable_comment(sheet.object_id, text_box.drawable_object_id)?;
            let columns =
                editor.sheet_text_box_columns(sheet.object_id, text_box.drawable_object_id)?;
            println!(
                "  text_box_index={index} drawable={} storage={} text={:?} geometry={geometry:?} properties={properties:?} columns={columns:?} comment={:?}",
                text_box.drawable_object_id,
                text_box.storage.object_id,
                text_box.storage.text,
                comment.as_ref().map(|value| value.comment.text.as_str())
            );
        }
    }
    Ok(())
}
