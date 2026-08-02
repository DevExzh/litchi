//! Remove one ordinary Pages body shape and its private graph.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_pages_shape <input.pages> <output.pages> <shape-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let shape_index: usize = arguments.next().ok_or("missing shape index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let shape = editor
        .body_shapes()?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    let removed = editor.remove_body_shape(shape.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "drawable={} storage={} anchor={} kind={:?}",
        removed.shape.drawable_object_id,
        removed.shape.storage.object_id,
        removed.shape.anchor_character_index,
        removed.shape.kind
    );
    Ok(())
}
