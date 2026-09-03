//! Create a Keynote presentation with slide audio and no input package.

use std::env;
use std::fs;
use std::path::Path;
use std::time::Duration;

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa_common::media::playback::{MediaLoopMode, MediaVolume};
use litchi_iwa_common::shape::geometry::Point;
use litchi_keynote::slide::audio::Options as SlideAudioOptions;
use litchi_keynote::{MovieSelector, Package as KeynotePackage, SlideSelector};

const SLIDE_CENTER: Point = Point { x: 960.0, y: 540.0 };

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_audio <output.key> <audio> <duration-seconds>")?;
    let audio_path = arguments.next().ok_or("missing audio path")?;
    let duration_seconds: f64 = arguments.next().ok_or("missing audio duration")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let preferred_filename = Path::new(&audio_path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("audio path must end in a UTF-8 file name")?;
    let audio = fs::read(&audio_path)?;

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Audio built from typed IWA objects")
        .build()?;
    let created = editor.add_slide_audio(
        0,
        preferred_filename,
        &audio,
        SlideAudioOptions::new(SLIDE_CENTER, Duration::try_from_secs_f64(duration_seconds)?)?,
    )?;
    let mut properties = editor.slide_audio_properties(0, created.drawable_object_id)?;
    properties.accessibility_description = Some(format!("Embedded audio: {preferred_filename}"));
    editor.set_slide_audio_properties(0, created.drawable_object_id, properties)?;
    let playback = created
        .playback
        .with_loop_mode(Some(MediaLoopMode::Repeat))
        .with_volume(Some(MediaVolume::new(0.75)?));
    let package = KeynotePackage::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_slide_movie_playback_settings(SlideSelector::index(0), MovieSelector::index(0))?
        .set(playback.try_into()?)?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    editor = KeynoteEditor::from_bytes(&bytes)?;
    editor.save(output)?;
    println!(
        "created Keynote audio {} backed by data {}",
        created.drawable_object_id, created.audio_data_identifier
    );
    Ok(())
}
