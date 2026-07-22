//! Create a Pages document and body-anchored audio without an input package.

use std::env;
use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::pages::{PagesAudioOptions, PagesEditor};
use litchi_iwa::shapes::DrawablePoint;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_audio <output.pages> <audio> <duration-seconds>")?;
    let audio_path = arguments.next().ok_or("missing audio path")?;
    let duration_seconds: f64 = arguments.next().ok_or("missing duration")?.parse()?;
    if !duration_seconds.is_finite() || duration_seconds <= 0.0 {
        return Err("duration must be finite and greater than zero".into());
    }
    let duration = Duration::try_from_secs_f64(duration_seconds)
        .map_err(|_| "duration is too large for std::time::Duration")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let audio = fs::read(&audio_path)?;
    let preferred_filename = Path::new(&audio_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("audio path has no UTF-8 basename")?;
    let body = "Source-built Pages audio";
    let anchor = body.encode_utf16().count();
    let mut editor = PagesEditor::create_with_text(body)?;
    let created = editor.add_body_audio(
        anchor,
        preferred_filename,
        &audio,
        PagesAudioOptions::new(DrawablePoint { x: 180.0, y: 240.0 }, duration),
    )?;
    let mut properties = editor.body_audio_properties(created.drawable_object_id)?;
    properties.accessibility_description = Some(format!("Embedded audio: {preferred_filename}"));
    editor.set_body_audio_properties(created.drawable_object_id, properties)?;
    editor.save(output)?;
    println!(
        "anchor={} drawable={} audio_data={} duration={:?}",
        created.anchor_character_index,
        created.drawable_object_id,
        created.audio_data_identifier,
        created.duration
    );
    Ok(())
}
