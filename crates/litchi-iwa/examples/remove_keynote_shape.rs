//! Remove an ordinary Keynote shape and its private object graph.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: remove_keynote_shape <input.key> <output.key> <slide-index> <shape-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let shape_index: usize = arguments.next().ok_or("missing shape index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let shape = editor
        .slide_shapes(slide_index)?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    let removed = editor.remove_slide_shape(slide_index, shape.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} storage={} kind={:?}",
        removed.shape.drawable_object_id, removed.shape.storage.id, removed.shape.kind
    );
    Ok(())
}
