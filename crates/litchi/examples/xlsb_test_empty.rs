//! Absolutely minimal XLSB test - empty workbook with one blank sheet

#![allow(clippy::all)]

use litchi::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal empty XLSB...");

    let mut workbook = WorkbookWriter::new();
    let sheet = MutableWorksheet::new("Sheet1");
    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_empty.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_empty.xlsb");
    Ok(())
}
