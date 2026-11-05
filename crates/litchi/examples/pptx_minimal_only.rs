//! Minimal PPTX test - just one slide with title only

use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal PPTX with one slide...");

    let mut builder = litchi_pptx::builder::Builder::new()?;
    let pres = builder.presentation()?;

    // Add single slide with just a title
    let slide = pres.add_slide()?;
    slide.set_title("Minimal Test");

    // Save
    let pkg = builder.build()?;
    pkg.save(Path::new("minimal_only.pptx"))?;
    println!("Saved: minimal_only.pptx");

    Ok(())
}
