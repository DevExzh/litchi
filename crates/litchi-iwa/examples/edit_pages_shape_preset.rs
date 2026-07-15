//! Replace one Pages body shape's path with a typed preset.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::ShapePreset;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_shape_preset <input.pages> <output.pages> <shape-index> <rectangle|rounded-rectangle|ellipse|pentagon|star>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
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

    let mut editor = PagesEditor::open(input)?;
    let shape = editor
        .body_shapes()?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    editor.set_body_shape_preset(shape.drawable_object_id, preset)?;
    editor.save(output)?;
    println!(
        "drawable={} anchor={} preset={preset:?}",
        shape.drawable_object_id, shape.anchor_character_index
    );
    Ok(())
}
