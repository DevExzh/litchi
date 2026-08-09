#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Open a `.pptx` presentation and print its slide count plus per-slide name
//! (which corresponds to the slide title when one is set) and text content.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-pptx --example read_pptx
//! cargo run -p litchi-pptx --example read_pptx -- path/to/file.pptx
//! ```
//!
//! Default input: `test-data/ooxml/pptx/sample.pptx`.

use std::env;

use litchi_pptx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() >= 2 {
        args[1].clone()
    } else {
        "test-data/ooxml/pptx/sample.pptx".to_string()
    };

    println!("Opening PPTX: {path}");
    let pkg = Package::open(&path)?;
    let pres = pkg.presentation()?;

    let slide_count = pres.slide_count()?;
    println!("Slide count: {slide_count}");

    let (width, height) = pres.slide_size()?;
    println!("Slide size : {width}x{height} EMUs");

    // Iterate slides.
    let slides = pres.slides()?;
    for (idx, slide) in slides.iter().enumerate() {
        let name = slide.name().unwrap_or_else(|_| "<unnamed>".to_string());
        let text = slide.text().unwrap_or_default();
        let shape_count = slide.shape_count().unwrap_or(0);

        println!(
            "\n--- Slide {} of {} ({} shapes) ---",
            idx + 1,
            slide_count,
            shape_count
        );
        println!("Name : {name}");
        if text.is_empty() {
            println!("(no text content)");
        } else {
            println!("Text :\n{text}");
        }
    }

    Ok(())
}
