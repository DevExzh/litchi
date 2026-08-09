#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Test PPT without any animations

use litchi_ppt::Writer;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let mut writer = Writer::new();

    let slide = writer.add_slide()?;
    writer.add_textbox(slide, 100, 100, 300, 100, "Test Without Animation")?;

    writer.save("output/test_no_animation.ppt")?;

    println!("✅ Created: output/test_no_animation.ppt (no animations)");

    Ok(())
}
