//! Test XLSB with a simple merged cell range only
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_merge_cells --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::merged_cells::MergedCell;
use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test M: Creating XLSB with one merged title row...");

    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Merged");

    // Title in A1, merge A1:D1
    sheet.set_cell(0, 0, CellValue::String("Sales Report".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 3));

    // Simple data row under the merged header so the sheet is non-empty beyond A1
    sheet.set_cell(1, 0, CellValue::String("October".to_string()));
    sheet.set_cell(1, 1, CellValue::Float(125_000.0));

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_merge_cells.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_merge_cells.xlsb");
    Ok(())
}
