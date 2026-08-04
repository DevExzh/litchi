//! Read an OpenDocument Presentation (`.odp`) file, print the slide count,
//! and dump per-slide title and text content.
//!
//! If a path argument is supplied, that file is opened. Otherwise the example
//! creates a small ODP in a tempfile via [`odp::Builder`] and reads it
//! back.
//!
//! Run with:
//! ```bash
//! cargo run -p litchi-odf --example read_odp
//! cargo run -p litchi-odf --example read_odp -- path/to/file.odp
//! ```

use std::path::PathBuf;

use litchi_odf::{Presentation, odp};
use tempfile::NamedTempFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let (path, _tempfile_guard): (PathBuf, Option<NamedTempFile>) = match std::env::args().nth(1) {
        Some(arg) => (PathBuf::from(arg), None),
        None => {
            println!("No path provided; creating a fresh ODP via odp::Builder...");
            let tmp = NamedTempFile::with_suffix(".odp")?;
            let mut builder = odp::Builder::new();
            builder.add_slide_with_title(
                "litchi-odf example",
                "This presentation was created by the read_odp example.",
            )?;
            builder.add_slide_with_title(
                "Slide Two",
                "Demonstrates a build-then-read round trip\n\
                 with multiple lines of body text.",
            )?;
            builder.add_slide_with_title("Final Slide", "Thanks for trying litchi-odf!")?;
            let path = tmp.path().to_path_buf();
            builder.save(&path)?;
            (path, Some(tmp))
        },
    };

    println!("Opening: {}", path.display());
    let pres = Presentation::open(&path)?;

    let slide_count = pres.slide_count()?;
    println!("Slide count: {}", slide_count);

    let slides = pres.slides()?;
    for slide in &slides {
        let title = slide.title()?.unwrap_or("<untitled>");
        let body = slide.text()?;
        println!("\n--- Slide {} ---", slide.index() + 1);
        println!("  title: {}", title);
        if !body.is_empty() {
            println!("  text:  {}", body.replace('\n', "\n         "));
        }
        let shapes = slide.shapes()?;
        if !shapes.is_empty() {
            println!("  shapes: {}", shapes.len());
        }
    }

    // tempfile (if any) is dropped here.
    Ok(())
}
