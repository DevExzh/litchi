//! Minimal test for hyperlinks only

#![allow(clippy::all)]

use litchi::ooxml::xlsb::hyperlinks::Hyperlink;
use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test: Minimal hyperlink...");

    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Test");

    // Simple data with one hyperlink
    sheet.set_cell(0, 0, CellValue::String("Click here".to_string()));

    let link = Hyperlink::new_external(0, 0, 0, 0, "https://example.com".to_string())
        .with_display("Click here".to_string());
    sheet.add_hyperlink(link);

    sheet.set_cell(1, 0, CellValue::String("Data1".to_string()));

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_minimal_hyperlink.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_minimal_hyperlink.xlsb");
    Ok(())
}
