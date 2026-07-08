//! Minimal test - ONE named range only

#![allow(clippy::all)]

use litchi::ooxml::xlsb::named_ranges::NamedRange;
use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test: ONE named range...");

    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Data");

    // Simple data
    sheet.set_cell(0, 0, CellValue::String("Value".to_string()));
    sheet.set_cell(1, 0, CellValue::Float(100.0));
    sheet.set_cell(2, 0, CellValue::Float(200.0));

    workbook.add_worksheet(sheet);

    // ONE simple named range without formula bytes
    let test_range = NamedRange::new("TestRange".to_string(), Some(0));
    workbook.add_named_range(test_range);

    let file = File::create("xlsb_test_one_named_range.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_one_named_range.xlsb");
    Ok(())
}
