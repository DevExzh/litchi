//! Duplicate an ordinary Keynote text box with independent text storage.

use std::env;

use litchi_iwa::keynote::{KeynoteEditor, KeynoteSlideTextRole};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_keynote_text_box <input.key> <output.key> <slide-index> <text-box-index> <text>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let text_box_index: usize = arguments.next().ok_or("missing text-box index")?.parse()?;
    let text = arguments.next().ok_or("missing text")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let source = editor
        .slide_text_storages(slide_index)?
        .into_iter()
        .filter(|item| item.role == KeynoteSlideTextRole::TextBox)
        .nth(text_box_index)
        .ok_or("text-box index is out of bounds")?;
    let created = editor.duplicate_slide_text_box(slide_index, source.drawable_object_id, &text)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} storage={} source={}",
        created.drawable_object_id, created.storage.object_id, source.drawable_object_id
    );
    Ok(())
}
