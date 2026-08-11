//! Update the position, size, and rotation of an ordinary Numbers text box.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_text_box_geometry <input.numbers> <output.numbers> <sheet-index> <drawable-id> <x> <y> <width> <height> <angle-degrees>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let drawable_id: u64 = arguments.next().ok_or("missing drawable ID")?.parse()?;
    let x: f32 = arguments.next().ok_or("missing x position")?.parse()?;
    let y: f32 = arguments.next().ok_or("missing y position")?.parse()?;
    let width: f32 = arguments.next().ok_or("missing width")?.parse()?;
    let height: f32 = arguments.next().ok_or("missing height")?.parse()?;
    let angle: f32 = arguments.next().ok_or("missing angle")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index out of range")?;
    let mut geometry = editor.sheet_text_box_geometry(sheet.id(), drawable_id)?;
    geometry.position = Some(DrawablePoint { x, y });
    geometry.size = Some(DrawableSize { width, height });
    geometry.angle = Some(angle);
    editor.set_sheet_text_box_geometry(sheet.id(), drawable_id, geometry)?;
    editor.save(output)?;
    println!("sheet={sheet_index} drawable={drawable_id} geometry={geometry:?}");
    Ok(())
}
