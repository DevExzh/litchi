//! Remove one body-anchored Pages audio clip and its private graph.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_pages_audio <input.pages> <output.pages> <audio-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let audio = editor
        .body_audio()?
        .get(audio_index)
        .cloned()
        .ok_or("audio index is out of bounds")?;
    let removed = editor.remove_body_audio(audio.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "drawable={} removed_data={:?}",
        removed.audio.drawable_object_id, removed.removed_data_identifiers
    );
    Ok(())
}
