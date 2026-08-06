//! Remove an ordinary Keynote text box and its private object graph.

use std::env;

use litchi_iwa::keynote::{KeynoteEditor, KeynoteSlideTextRole};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: remove_keynote_text_box <input.key> <output.key> <slide-index> <text-box-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let text_box_index: usize = arguments.next().ok_or("missing text-box index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let target = editor
        .slide_text_storages(slide_index)?
        .into_iter()
        .filter(|item| item.role == KeynoteSlideTextRole::TextBox)
        .nth(text_box_index)
        .ok_or("text-box index is out of bounds")?;
    let removed = editor.remove_slide_text_box(slide_index, target.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} storage={}",
        removed.text.drawable_object_id, removed.text.storage.id
    );
    Ok(())
}
