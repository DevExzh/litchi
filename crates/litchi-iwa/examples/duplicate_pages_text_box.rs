//! Duplicate a body-anchored Pages text box with independent text storage.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_pages_text_box <input.pages> <output.pages> <source-index> <body-utf16-index> <text>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let source_index: usize = arguments.next().ok_or("missing source index")?.parse()?;
    let body_index: usize = arguments
        .next()
        .ok_or("missing body UTF-16 index")?
        .parse()?;
    let text = arguments.next().ok_or("missing text")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let source = editor
        .drawable_text_storages()?
        .get(source_index)
        .cloned()
        .ok_or("source index is out of bounds")?;
    let created = editor.duplicate_text_box(source.drawable_object_id, body_index, &text)?;
    editor.save(output)?;
    println!(
        "drawable={} storage={} source={}",
        created.drawable_object_id, created.storage.id, source.drawable_object_id
    );
    Ok(())
}
