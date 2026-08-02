//! Create a Numbers spreadsheet and sheet audio without an input package.

use std::env;
use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersSheetAudioOptions};
use litchi_iwa::shapes::DrawablePoint;
use litchi_iwa::{MediaLoopMode, MediaVolume};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_audio <output.numbers> <audio> <duration-seconds>")?;
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
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Source-built Audio")
        .table_name("Scratch Table")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let created = editor.add_sheet_audio(
        sheet_id,
        preferred_filename,
        &audio,
        NumbersSheetAudioOptions::new(DrawablePoint { x: 420.0, y: 180.0 }, duration),
    )?;
    let mut properties = editor.sheet_audio_properties(sheet_id, created.drawable_object_id)?;
    properties.accessibility_description = Some(format!("Embedded audio: {preferred_filename}"));
    editor.set_sheet_audio_properties(sheet_id, created.drawable_object_id, properties)?;
    editor.set_sheet_audio_playback_settings(
        sheet_id,
        created.drawable_object_id,
        created
            .playback
            .with_loop_mode(Some(MediaLoopMode::Repeat))
            .with_volume(Some(MediaVolume::new(0.75)?)),
    )?;
    editor.save(output)?;
    println!(
        "sheet={sheet_id} drawable={} audio_data={} duration={:?}",
        created.drawable_object_id, created.audio_data_identifier, created.duration
    );
    Ok(())
}
