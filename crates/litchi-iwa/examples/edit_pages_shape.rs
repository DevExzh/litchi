//! Replace all text in one ordinary Pages body shape.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_shape <input.pages> <output.pages> <shape-index> <replacement> [title] [caption]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let shape_index: usize = arguments.next().ok_or("missing shape index")?.parse()?;
    let replacement = arguments.next().ok_or("missing replacement text")?;
    let title = arguments.next();
    let caption = arguments.next();
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let shape = editor
        .body_shapes()?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    editor.set_body_shape_text(shape.drawable_object_id, &replacement)?;
    if let Some(title) = title {
        editor.set_body_shape_title(shape.drawable_object_id, &title)?;
    }
    if let Some(caption) = caption {
        editor.set_body_shape_caption(shape.drawable_object_id, &caption)?;
    }
    editor.save(output)?;
    println!(
        "drawable={} storage={} anchor={} kind={:?} title_caption={:?}",
        shape.drawable_object_id,
        shape.storage.id,
        shape.anchor_character_index,
        shape.kind,
        editor.body_shape_title_caption(shape.drawable_object_id)?
    );
    Ok(())
}
