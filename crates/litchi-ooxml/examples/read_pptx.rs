//! Open a `.pptx` presentation and print its slide count plus per-slide name
//! (which corresponds to the slide title when one is set) and text content.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-ooxml --example read_pptx --all-features
//! cargo run -p litchi-ooxml --example read_pptx --all-features -- path/to/file.pptx
//! ```
//!
//! Default input: `test-data/ooxml/pptx/sample.pptx`.

use std::env;

use litchi_ooxml::pptx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() >= 2 {
        args[1].clone()
    } else {
        "test-data/ooxml/pptx/sample.pptx".to_string()
    };

    println!("Opening PPTX: {}", path);
    let pkg = Package::open(&path)?;
    let pres = pkg.presentation()?;

    let slide_count = pres.slide_count()?;
    println!("Slide count: {}", slide_count);

    if let (Some(w), Some(h)) = (pres.slide_width()?, pres.slide_height()?) {
        println!("Slide size : {}x{} EMUs", w, h);
    }

    // Core document properties (some PPTX files leave most of these unset).
    let props = pkg.properties();
    println!("\n--- Core properties ---");
    println!("Title   : {:?}", props.title);
    println!("Creator : {:?}", props.creator);
    println!("Subject : {:?}", props.subject);
    println!("Modified: {:?}", props.modified);

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
        println!("Name : {}", name);
        if text.is_empty() {
            println!("(no text content)");
        } else {
            println!("Text :\n{}", text);
        }
    }

    Ok(())
}
