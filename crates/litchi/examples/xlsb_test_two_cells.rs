//! Test XLSB with two cells in same row
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_two_cells --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 3: Creating XLSB with two cells in same row...");

    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, CellValue::String("A".to_string()));
    sheet.set_cell(0, 1, CellValue::String("B".to_string()));
    workbook.add_worksheet(sheet);

    let file = File::create("test_03_two_cells.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created test_03_two_cells.xlsb");
    Ok(())
}
