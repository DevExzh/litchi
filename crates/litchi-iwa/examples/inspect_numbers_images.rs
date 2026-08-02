//! List ordinary images owned by each Numbers sheet.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_numbers_images <input.numbers>")?;
    let editor = NumbersEditor::open(input)?;
    for sheet in editor.sheets()? {
        println!(
            "sheet={} object={} name={:?}",
            sheet.index, sheet.object_id, sheet.name
        );
        for (index, image) in editor
            .sheet_images(sheet.object_id)?
            .into_iter()
            .enumerate()
        {
            let comment =
                editor.sheet_drawable_comment(sheet.object_id, image.drawable_object_id)?;
            println!(
                "  image_index={index} drawable={} data={} thumbnail={:?} geometry={:?} original={:?} natural={:?} comment={:?}",
                image.drawable_object_id,
                image.image_data_identifier,
                image.thumbnail_data_identifier,
                image.geometry,
                image.original_size,
                image.natural_size,
                comment.as_ref().map(|value| value.comment.text.as_str())
            );
        }
    }
    Ok(())
}
