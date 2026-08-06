//! Test XLSB with merged cells + hyperlinks (NO named ranges)

#![allow(clippy::all)]

use litchi::sheet::CellValue;
use litchi::xlsb::hyperlinks::Hyperlink;
use litchi::xlsb::merged_cells::MergedCell;
use litchi::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Data");

    // Title row (merged)
    sheet.set_cell(0, 0, CellValue::String("Sales Report".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 3));

    // Data with hyperlink
    sheet.set_cell(1, 0, CellValue::String("Widget A".to_string()));
    sheet.set_cell(1, 1, CellValue::String("Details".to_string()));
    let link = Hyperlink::new_external(1, 1, 1, 1, "https://example.com/a".to_string())
        .with_display("Details".to_string());
    sheet.add_hyperlink(link);

    workbook.add_worksheet(sheet);
    // NO named ranges added

    let file = File::create("xlsb_test_merged_hyperlinks.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_merged_hyperlinks.xlsb");
    println!("  - Merged cells + Hyperlinks (NO named ranges)");
    Ok(())
}
