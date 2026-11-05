//! List the worksheets of the XLSB workbook that the `write_xlsb` example
//! writes to `target/example.xlsb`, through the unified `litchi::open` API.
//!
//! Run with:
//! ```bash
//! cargo run --example write_xlsb --features xlsb
//! cargo run --example read_xlsb --features xlsb
//! ```

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // The format is detected from the content; the workbook view refuses a
    // file that is not a workbook.
    let file = litchi::open("target/example.xlsb")?;
    let workbook = file.workbook()?;
    println!("Sheets: {}", workbook.worksheet_count()?);
    for name in workbook.worksheet_names()? {
        println!("- {}", name);
    }
    Ok(())
}
