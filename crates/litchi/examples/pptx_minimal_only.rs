//! Minimal PPTX test - just one slide with title only

use litchi::pptx::Package;
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal PPTX with one slide...");

    let mut pkg = Package::new()?;
    let pres = pkg.presentation_mut()?;

    // Add single slide with just a title
    let slide = pres.add_slide()?;
    slide.set_title("Minimal Test");

    // Save
    pkg.save(Path::new("minimal_only.pptx"))?;
    println!("Saved: minimal_only.pptx");

    Ok(())
}
