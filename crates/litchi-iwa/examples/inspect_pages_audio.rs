//! Inspect audio-only media controls anchored to the Pages body.

use std::env;

use litchi_iwa::pages;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_audio <input.pages>")?;
    let editor = pages::Editor::open(input)?;
    for (audio_index, audio) in editor.body_audio()?.iter().enumerate() {
        println!(
            "audio_index={audio_index} anchor={} drawable={} audio_data={} position={:?} duration={:?}",
            audio.anchor_character_index,
            audio.drawable_object_id,
            audio.audio_data_identifier,
            audio.position,
            audio.duration
        );
    }
    Ok(())
}
