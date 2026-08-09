#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Minimal PPT test - single empty slide

use litchi_ppt::writer::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal PPT with empty slide...");

    let mut writer = Writer::new();
    writer.add_slide()?;
    writer.save("output_minimal.ppt")?;

    println!("Created output_minimal.ppt");
    Ok(())
}
