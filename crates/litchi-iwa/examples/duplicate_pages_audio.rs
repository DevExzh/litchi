//! Duplicate a body-anchored Pages audio control.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_pages_audio <input.pages> <output.pages> <audio-index> <utf16-anchor>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    let anchor: usize = arguments.next().ok_or("missing UTF-16 anchor")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let source = editor
        .body_audio()?
        .into_iter()
        .nth(audio_index)
        .ok_or("body audio index is out of bounds")?;
    let created = editor.duplicate_body_audio(source.drawable_object_id, anchor)?;
    editor.save(output)?;
    println!(
        "anchor={} drawable={} source={} audio={}",
        created.anchor_character_index,
        created.drawable_object_id,
        source.drawable_object_id,
        created.audio_data_identifier,
    );
    Ok(())
}
