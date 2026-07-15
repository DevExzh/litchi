//! Replace one Numbers shape's path with a typed preset.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::shapes::ShapePreset;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_shape_preset <input.numbers> <output.numbers> <sheet-index> <shape-index> <rectangle|rounded-rectangle|ellipse|pentagon|star>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let shape_index: usize = arguments.next().ok_or("missing shape index")?.parse()?;
    let preset = match arguments.next().as_deref() {
        Some("rectangle") => ShapePreset::Rectangle,
        Some("rounded-rectangle") => ShapePreset::ROUNDED_RECTANGLE,
        Some("ellipse") => ShapePreset::Ellipse,
        Some("pentagon") => ShapePreset::PENTAGON,
        Some("star") => ShapePreset::STAR,
        Some(other) => return Err(format!("unsupported shape preset {other:?}").into()),
        None => return Err("missing shape preset".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet_id = editor
        .sheets()?
        .get(sheet_index)
        .ok_or("sheet index is out of bounds")?
        .object_id;
    let shape = editor
        .sheet_shapes(sheet_id)?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    editor.set_sheet_shape_preset(sheet_id, shape.drawable_object_id, preset)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_id} drawable={} preset={preset:?}",
        shape.drawable_object_id
    );
    Ok(())
}
