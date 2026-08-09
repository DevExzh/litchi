#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Empty PPT test - no slides at all

use litchi_ppt::writer::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating empty PPT with no slides...");

    let mut writer = Writer::new();
    writer.save("output_empty.ppt")?;

    println!("Created output_empty.ppt");
    Ok(())
}
