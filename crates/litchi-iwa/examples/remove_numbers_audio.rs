//! Remove one Numbers audio clip and its private graph.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: remove_numbers_audio <input.numbers> <output.numbers> <sheet-index> <audio-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet_id = editor
        .sheets()?
        .get(sheet_index)
        .ok_or("sheet index is out of bounds")?
        .id();
    let audio = editor
        .sheet_audio(sheet_id)?
        .get(audio_index)
        .cloned()
        .ok_or("audio index is out of bounds")?;
    let removed = editor.remove_sheet_audio(sheet_id, audio.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_id} drawable={} removed_data={:?}",
        removed.audio.drawable_object_id, removed.removed_data_identifiers
    );
    Ok(())
}
