//! Test XLSB with one number cell
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_one_number --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 2: Creating XLSB with one number cell...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, CellValue::Float(42.0));
    workbook.add_worksheet(sheet);

    let file = File::create("test_02_one_number.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created test_02_one_number.xlsb");
    Ok(())
}
