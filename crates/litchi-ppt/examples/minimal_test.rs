#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Minimal PPT test - single empty slide

use litchi_ppt::writer::Builder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal PPT with empty slide...");

    let mut writer = Builder::new();
    writer.add_slide()?;
    writer.build()?.save("output_minimal.ppt")?;

    println!("Created output_minimal.ppt");
    Ok(())
}
