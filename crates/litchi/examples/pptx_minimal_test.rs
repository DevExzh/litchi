//! Minimal PPTX test - creates a simple presentation without advanced features
//!
//! This is useful for verifying the basic PPTX structure is correct.
//!
//! ```bash
//! cargo run --example pptx_minimal_test
//! ```

use litchi::pptx::*;
use std::error::Error;

fn main() -> std::result::Result<(), Box<dyn Error>> {
    println!("Creating minimal PPTX...");

    let mut pkg = Package::new()?;

    {
        let pres = pkg.presentation_mut()?;
        pres.set_widescreen_slide_size();

        // Create just 2 simple slides
        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1");
        slide1.add_text_box("This is slide 1", 914400, 1828800, 7315200, 914400);

        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2");
        slide2.add_text_box("This is slide 2", 914400, 1828800, 7315200, 914400);

        println!("Created {} slides", pres.slide_count());
    }

    let output = "pptx_minimal_test.pptx";
    pkg.save(output)?;
    println!("Saved to: {}", output);

    Ok(())
}
