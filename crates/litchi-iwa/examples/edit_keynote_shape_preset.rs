//! Replace one Keynote shape's path with a typed preset.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::shapes::ShapePreset;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_shape_preset <input.key> <output.key> <slide-index> <shape-index> <rectangle|rounded-rectangle|ellipse|pentagon|star>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
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

    let mut editor = KeynoteEditor::open(input)?;
    let shape = editor
        .slide_shapes(slide_index)?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    editor.set_slide_shape_preset(slide_index, shape.drawable_object_id, preset)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} preset={preset:?}",
        shape.drawable_object_id
    );
    Ok(())
}
