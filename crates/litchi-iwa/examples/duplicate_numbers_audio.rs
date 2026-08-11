//! Duplicate a sheet-owned Numbers audio control.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_numbers_audio <input.numbers> <output.numbers> <sheet-index> <audio-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index is out of bounds")?;
    let source = editor
        .sheet_audio(sheet.id())?
        .into_iter()
        .nth(audio_index)
        .ok_or("sheet audio index is out of bounds")?;
    let created = editor.duplicate_sheet_audio(sheet.id(), source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "sheet={} drawable={} source={} audio={}",
        sheet.id(),
        created.drawable_object_id,
        source.drawable_object_id,
        created.audio_data_identifier,
    );
    Ok(())
}
