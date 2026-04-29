//! Minimal PPTX demo with ONLY media for testing.
//!
//! Run with: cargo run --example pptx_media_only

use litchi::ooxml::pptx::Package;
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal PPTX with only media...\n");

    let mut pkg = Package::new()?;
    let pres = pkg.presentation_mut()?;

    // Only one slide with one audio
    let slide = pres.add_slide()?;
    slide.set_title("Audio Test");

    // Add just the MP3 audio
    let audio_data = fs::read("file_example_MP3_700KB.mp3")?;
    slide.add_audio(audio_data, 914400, 1828800, 914400, 914400);
    println!("Added MP3 audio");

    // Save
    let output_path = "pptx_media_only.pptx";
    println!("\nSaving to {}...", output_path);
    pkg.save(output_path)?;
    println!("✓ Done!");

    Ok(())
}
