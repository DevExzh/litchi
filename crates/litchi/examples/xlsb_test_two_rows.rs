//! Test XLSB with cells in two rows
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_two_rows --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 4: Creating XLSB with two rows...");

    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, CellValue::String("A1".to_string()));
    sheet.set_cell(1, 0, CellValue::String("A2".to_string()));
    workbook.add_worksheet(sheet);

    let file = File::create("test_04_two_rows.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created test_04_two_rows.xlsb");
    Ok(())
}
