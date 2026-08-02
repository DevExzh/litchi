//! Move one Numbers sheet audio control.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::shapes::DrawablePoint;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_audio_position <input.numbers> <output.numbers> <sheet-index> <audio-index> <x> <y>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    let position = DrawablePoint {
        x: arguments.next().ok_or("missing x")?.parse()?,
        y: arguments.next().ok_or("missing y")?.parse()?,
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet_id = editor
        .sheets()?
        .get(sheet_index)
        .ok_or("sheet index is out of bounds")?
        .object_id;
    let audio = editor
        .sheet_audio(sheet_id)?
        .get(audio_index)
        .cloned()
        .ok_or("audio index is out of bounds")?;
    editor.set_sheet_audio_position(sheet_id, audio.drawable_object_id, position)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_id} drawable={} position={position:?}",
        audio.drawable_object_id
    );
    Ok(())
}
