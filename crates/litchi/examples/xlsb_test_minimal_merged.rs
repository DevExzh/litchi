//! Minimal test for merged cells only

#![allow(clippy::all)]

use litchi::sheet::CellValue;
use litchi::xlsb::builder::{self, Builder};
use litchi::xlsb::merged_cells::MergedCell;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test: Minimal merged cells...");

    let mut workbook = Builder::new();
    let mut sheet = builder::Worksheet::new("Test");

    // Simple data with one merged cell
    sheet.set_cell(0, 0, CellValue::String("Merged Title".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 2)); // A1:C1

    sheet.set_cell(1, 0, CellValue::String("Data1".to_string()));
    sheet.set_cell(1, 1, CellValue::String("Data2".to_string()));

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_minimal_merged.xlsb")?;
    workbook.build()?.write_to(file)?;

    println!("✓ Created xlsb_test_minimal_merged.xlsb");
    Ok(())
}
