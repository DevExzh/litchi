//! Test XLSB with one number cell
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_one_number --features ooxml --no-default-features
//! ```

use litchi::sheet::CellValue;
use litchi::xlsb::builder::{self, Builder};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 2: Creating XLSB with one number cell...");

    let mut workbook = Builder::new();
    let mut sheet = builder::Worksheet::new("Sheet1");
    sheet.set_cell(0, 0, CellValue::Float(42.0));
    workbook.add_worksheet(sheet);

    let file = File::create("test_02_one_number.xlsb")?;
    workbook.build()?.write_to(file)?;

    println!("✓ Created test_02_one_number.xlsb");
    Ok(())
}
