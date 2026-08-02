//! Duplicate a slide-owned Keynote audio control.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_keynote_audio <input.key> <output.key> <slide-index> <audio-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index: usize = arguments.next().ok_or("missing slide index")?.parse()?;
    let audio_index: usize = arguments.next().ok_or("missing audio index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    let source = editor
        .slide_audio(slide_index)?
        .into_iter()
        .nth(audio_index)
        .ok_or("slide audio index is out of bounds")?;
    let created = editor.duplicate_slide_audio(slide_index, source.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "slide={slide_index} drawable={} source={} audio={}",
        created.drawable_object_id, source.drawable_object_id, created.audio_data_identifier,
    );
    Ok(())
}
