//! Minimal XLSB test - empty worksheet
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_minimal --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 0: Creating minimal XLSB with empty sheet...");

    let mut workbook = WorkbookWriter::new();
    let sheet = MutableWorksheet::new("Sheet1");
    workbook.add_worksheet(sheet);

    let file = File::create("test_00_empty.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created test_00_empty.xlsb");
    Ok(())
}
