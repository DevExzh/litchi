// Minimal PPT writer example to verify structure
// Generates a minimal PowerPoint 97-2003 file named "minimal.ppt"

use litchi::ole::ppt::PptWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create empty presentation (master only, no slides) to match POI's empty.ppt
    let mut writer = PptWriter::new();
    writer.save("minimal.ppt")?;
    println!("Created minimal.ppt (empty, master only)");

    // Also test with a slide
    let mut writer2 = PptWriter::new();
    writer2.add_slide()?;
    writer2.save("with_slide.ppt")?;
    println!("Created with_slide.ppt (with 1 slide)");

    Ok(())
}
