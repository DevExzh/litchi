//! Replace all text in one ordinary Keynote shape.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_shape <input.key> <output.key> <slide-index> <shape-index> <replacement> [title] [caption]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let shape_index: usize = arguments.next().ok_or("missing shape index")?.parse()?;
    let replacement = arguments.next().ok_or("missing replacement text")?;
    let title = arguments.next();
    let caption = arguments.next();
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let shape = editor
        .slide_shapes(slide_index)?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    editor.set_slide_shape_text(slide_index, shape.drawable_object_id, &replacement)?;
    if let Some(title) = title {
        editor.set_slide_shape_title(slide_index, shape.drawable_object_id, &title)?;
    }
    if let Some(caption) = caption {
        editor.set_slide_shape_caption(slide_index, shape.drawable_object_id, &caption)?;
    }
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} storage={} kind={:?} title_caption={:?}",
        shape.drawable_object_id,
        shape.storage.id,
        shape.kind,
        editor.slide_shape_title_caption(slide_index, shape.drawable_object_id)?
    );
    Ok(())
}
