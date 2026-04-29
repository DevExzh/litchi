//! Test merged cells AND hyperlinks together

#![allow(clippy::all)]

use litchi::ooxml::xlsb::hyperlinks::Hyperlink;
use litchi::ooxml::xlsb::merged_cells::MergedCell;
use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test: Merged cells + hyperlinks combined...");

    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Dashboard");

    // Title row with merged cells
    sheet.set_cell(0, 0, CellValue::String("Sales Dashboard".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 3)); // A1:D1

    // Headers
    sheet.set_cell(2, 0, CellValue::String("Region".to_string()));
    sheet.set_cell(2, 1, CellValue::String("Sales".to_string()));
    sheet.set_cell(2, 2, CellValue::String("Details".to_string()));

    // Data with hyperlinks
    let regions = vec!["North", "South"];
    for (i, region) in regions.iter().enumerate() {
        let row = (i + 3) as u32;
        sheet.set_cell(row, 0, CellValue::String(region.to_string()));
        sheet.set_cell(row, 1, CellValue::Float(45000.0));
        sheet.set_cell(row, 2, CellValue::String("View Report".to_string()));

        let link = Hyperlink::new_external(
            row,
            row,
            2,
            2,
            format!("https://example.com/reports/{}", region.to_lowercase()),
        )
        .with_display("View Report".to_string());
        sheet.add_hyperlink(link);
    }

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_merged_and_hyperlinks.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_merged_and_hyperlinks.xlsb");
    Ok(())
}
