//! Replace text in any slide-owned Keynote text box or placeholder.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_text_box <input.key> <output.key> <slide-index> <text-index> <replacement>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let text_index: usize = arguments.next().ok_or("missing text index")?.parse()?;
    let replacement = arguments.next().ok_or("missing replacement text")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let text = editor
        .slide_text_storages(slide_index)?
        .get(text_index)
        .cloned()
        .ok_or("text index is out of bounds")?;
    editor.set_slide_text_storage(slide_index, text.drawable_object_id, &replacement)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} storage={} role={:?}",
        text.drawable_object_id, text.storage.id, text.role
    );
    Ok(())
}
