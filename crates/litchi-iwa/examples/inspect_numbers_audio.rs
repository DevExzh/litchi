//! Inspect audio-only media controls owned by Numbers sheets.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_numbers_audio <input.numbers>")?;
    let editor = NumbersEditor::open(input)?;
    for sheet in editor.sheets()? {
        for (audio_index, audio) in editor.sheet_audio(sheet.id())?.iter().enumerate() {
            println!(
                "sheet={} audio_index={audio_index} drawable={} audio_data={} position={:?} duration={:?}",
                sheet.id(),
                audio.drawable_object_id,
                audio.audio_data_identifier,
                audio.position,
                audio.duration
            );
        }
    }
    Ok(())
}
