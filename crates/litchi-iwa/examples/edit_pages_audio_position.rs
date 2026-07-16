//! Move one body-anchored Pages audio control.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::DrawablePoint;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_audio_position <input.pages> <output.pages> <audio-index> <x> <y>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    let position = DrawablePoint {
        x: arguments.next().ok_or("missing x")?.parse()?,
        y: arguments.next().ok_or("missing y")?.parse()?,
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let audio = editor
        .body_audio()?
        .get(audio_index)
        .cloned()
        .ok_or("audio index is out of bounds")?;
    editor.set_body_audio_position(audio.drawable_object_id, position)?;
    editor.save(output)?;
    println!(
        "drawable={} position={position:?}",
        audio.drawable_object_id
    );
    Ok(())
}
