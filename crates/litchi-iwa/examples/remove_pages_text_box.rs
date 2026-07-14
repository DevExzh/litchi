//! Remove a body-anchored Pages text box and its private object graph.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_pages_text_box <input.pages> <output.pages> <text-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let text_index: usize = arguments.next().ok_or("missing text index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let target = editor
        .drawable_text_storages()?
        .get(text_index)
        .cloned()
        .ok_or("text index is out of bounds")?;
    let removed = editor.remove_text_box(target.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "drawable={} storage={} anchor={}",
        removed.text.drawable_object_id,
        removed.text.storage.object_id,
        removed.anchor_character_index
    );
    Ok(())
}
