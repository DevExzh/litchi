//! Test XLSB with named ranges
//!
//! This example demonstrates the named ranges write functionality.
//! Creates a workbook with various named ranges (global and sheet-scoped).
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_named_ranges --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::named_ranges::{NamedRange, create_area3d_formula};
use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test NR: Creating XLSB with named ranges...");

    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Data");

    // Create some data
    sheet.set_cell(0, 0, CellValue::String("Product".to_string()));
    sheet.set_cell(0, 1, CellValue::String("Price".to_string()));
    sheet.set_cell(0, 2, CellValue::String("Quantity".to_string()));

    sheet.set_cell(1, 0, CellValue::String("Widget A".to_string()));
    sheet.set_cell(1, 1, CellValue::Float(19.99));
    sheet.set_cell(1, 2, CellValue::Float(100.0));

    sheet.set_cell(2, 0, CellValue::String("Widget B".to_string()));
    sheet.set_cell(2, 1, CellValue::Float(29.99));
    sheet.set_cell(2, 2, CellValue::Float(75.0));

    sheet.set_cell(3, 0, CellValue::String("Widget C".to_string()));
    sheet.set_cell(3, 1, CellValue::Float(39.99));
    sheet.set_cell(3, 2, CellValue::Float(50.0));

    // Add total row
    sheet.set_cell(4, 0, CellValue::String("Total".to_string()));
    sheet.set_cell(4, 2, CellValue::Float(225.0));

    workbook.add_worksheet(sheet);

    // Create named range for price column B2:B4 (rows 1-3, col 1, 0-indexed)
    // Formula bytes use Area3dPtg format: sheet_idx, row1, row2, col1, col2
    let prices_formula = create_area3d_formula(0, 1, 3, 1, 1)?;
    let prices_range = NamedRange::new("Prices".to_string(), Some(0)).with_formula(prices_formula);
    workbook.add_named_range(prices_range);

    // Create named range for quantity column C2:C4 (rows 1-3, col 2, 0-indexed)
    let quantities_formula = create_area3d_formula(0, 1, 3, 2, 2)?;
    let quantities_range =
        NamedRange::new("Quantities".to_string(), Some(0)).with_formula(quantities_formula);
    workbook.add_named_range(quantities_range);

    // Create global named range for total cell C5 (row 4, col 2)
    let total_formula = create_area3d_formula(0, 4, 4, 2, 2)?;
    let total_range = NamedRange::new("GrandTotal".to_string(), None).with_formula(total_formula);
    workbook.add_named_range(total_range);

    // Create a hidden named range for A1:C5 (entire data area)
    let hidden_formula = create_area3d_formula(0, 0, 4, 0, 2)?;
    let hidden_range = NamedRange::new("_HiddenCalc".to_string(), None)
        .with_formula(hidden_formula)
        .with_hidden(true);
    workbook.add_named_range(hidden_range);

    let file = File::create("xlsb_test_named_ranges.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_named_ranges.xlsb");
    println!("  - Open in Excel and use Name Manager (Formulas > Name Manager) to verify");
    println!("  - Named ranges: Prices, Quantities, GrandTotal, _HiddenCalc");
    Ok(())
}
