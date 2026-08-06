//! Test XLSB with one cell
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_one_cell --features ooxml --no-default-features
//! ```

use litchi::sheet::CellValue;
use litchi::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 1: Creating XLSB with one string cell...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, CellValue::String("Hello".to_string()));
    workbook.add_worksheet(sheet);

    let file = File::create("test_01_one_cell.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created test_01_one_cell.xlsb");
    Ok(())
}
