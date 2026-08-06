//! Test XLSB AutoFilter skeleton
//!
//! This example creates a worksheet with a small table and configures a basic
//! auto-filter range via `MutableWorksheet::set_auto_filter`. The writer
//! emits a `BrtBeginAFilter`/`BrtEndAFilter` pair for the specified range.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_autofilter --features ooxml --no-default-features
//! ```

use litchi::sheet::CellValue;
use litchi::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating XLSB with a basic AutoFilter range...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("AutoFilter");

    // Header row (row 0)
    sheet.set_cell(0, 0, CellValue::String("Name".to_string()));
    sheet.set_cell(0, 1, CellValue::String("Region".to_string()));
    sheet.set_cell(0, 2, CellValue::String("Sales".to_string()));

    // Data rows (rows 1-5)
    let rows = [
        ("Alice", "North", 120_000.0),
        ("Bob", "South", 95_000.0),
        ("Carol", "East", 110_500.0),
        ("Dave", "West", 87_250.0),
        ("Eve", "North", 132_750.0),
    ];

    for (i, (name, region, sales)) in rows.iter().enumerate() {
        let r = (i + 1) as u32; // 0-based row index
        sheet.set_cell(r, 0, CellValue::String((*name).to_string()));
        sheet.set_cell(r, 1, CellValue::String((*region).to_string()));
        sheet.set_cell(r, 2, CellValue::Float(*sales));
    }

    // AutoFilter over A1:C6 (0-based: rows 0-5, cols 0-2)
    sheet.set_auto_filter(0, 5, 0, 2);

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_autofilter.xlsb")?;
    workbook.save(file)?;

    println!("\n[32m✓ Created xlsb_test_autofilter.xlsb[0m");
    Ok(())
}
