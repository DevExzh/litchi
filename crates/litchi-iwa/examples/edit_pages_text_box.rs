//! Replace text in a reachable Pages text box or placeholder.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_text_box <input.pages> <output.pages> <text-index> <replacement>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let text_index: usize = arguments.next().ok_or("missing text index")?.parse()?;
    let replacement = arguments.next().ok_or("missing replacement text")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let text = editor
        .drawable_text_storages()?
        .get(text_index)
        .cloned()
        .ok_or("text index is out of bounds")?;
    editor.set_drawable_text(text.drawable_object_id, &replacement)?;
    editor.save(output)?;
    println!(
        "drawable={} storage={}",
        text.drawable_object_id, text.storage.id
    );
    Ok(())
}
