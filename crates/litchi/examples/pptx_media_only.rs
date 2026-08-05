//! Minimal PresentationML media model and XML smoke test.
//!
//! The typed media item owns the payload and emits the slide-picture XML
//! fragment after the package writer assigns relationship IDs.
//!
//! Run with: cargo run --example pptx_media_only --features ooxml

use litchi_pptx::presentation::media::{Format, Item, Kind};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let audio_data = fs::read("file_example_MP3_700KB.mp3")?;
    let mut audio = Item::with_format(
        audio_data,
        Format::Mp3,
        914_400,
        1_828_800,
        914_400,
        914_400,
    )
    .with_auto_play();
    audio.set_name("Audio Test");
    audio.validate()?;

    // The package boundary supplies the relationship IDs; this is the
    // canonical typed XML fragment for the media-only slide shape.
    let shape_xml = audio.to_shape_xml(2, "rIdAudio", None)?;
    assert_eq!(audio.kind(), Kind::Audio);
    assert!(shape_xml.contains("audioFile"));
    assert!(shape_xml.contains("rIdAudio"));

    println!("Validated one MP3 media item ({} bytes)", audio.data.len());
    println!("Generated media shape XML ({} bytes)", shape_xml.len());

    Ok(())
}
