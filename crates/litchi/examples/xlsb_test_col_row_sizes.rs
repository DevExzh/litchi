//! Test XLSB column widths and row heights
//!
//! This example creates a single worksheet with customized column widths and
//! row heights using `MutableWorksheet::set_column_width` and
//! `MutableWorksheet::set_row_height`.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_col_row_sizes --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating XLSB with custom column widths and row heights...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Sizes");

    // Header row
    sheet.set_cell(0, 0, CellValue::String("Wide Col".to_string()));
    sheet.set_cell(0, 1, CellValue::String("Default".to_string()));
    sheet.set_cell(0, 2, CellValue::String("Narrow".to_string()));

    // Sample data rows
    sheet.set_cell(1, 0, CellValue::String("Tall Row".to_string()));
    sheet.set_cell(1, 1, CellValue::String("Middle".to_string()));
    sheet.set_cell(1, 2, CellValue::String("Short".to_string()));

    sheet.set_cell(2, 0, CellValue::String("More data".to_string()));
    sheet.set_cell(2, 1, CellValue::String("...".to_string()));
    sheet.set_cell(2, 2, CellValue::String("...".to_string()));

    // Column widths (character units)
    sheet.set_column_width(0, 30.0); // A: wide
    sheet.set_column_width(2, 5.0); // C: narrow

    // Row heights (points)
    sheet.set_row_height(0, 30.0); // Row 1: tall header
    sheet.set_row_height(2, 10.0); // Row 3: short

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_col_row_sizes.xlsb")?;
    workbook.save(file)?;

    println!("\n[32m✓ Created xlsb_test_col_row_sizes.xlsb[0m");
    Ok(())
}
