//! Move an independently positioned Keynote slide-audio control.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::shapes::DrawablePoint;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_keynote_audio_position <input.key> <output.key> <slide-index> <audio-index> <x> <y>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    let x: f32 = arguments.next().ok_or("missing x")?.parse()?;
    let y: f32 = arguments.next().ok_or("missing y")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let source = editor
        .slide_audio(slide_index)?
        .get(audio_index)
        .cloned()
        .ok_or("audio index is out of bounds")?;
    let position = DrawablePoint { x, y };
    editor.set_slide_audio_position(slide_index, source.drawable_object_id, position)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} position={position:?}",
        source.drawable_object_id
    );
    Ok(())
}
