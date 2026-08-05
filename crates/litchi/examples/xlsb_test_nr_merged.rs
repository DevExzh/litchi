//! Test XLSB with named ranges + merged cells (no hyperlinks)

#![allow(clippy::all)]

use litchi::ooxml::xlsb::merged_cells::MergedCell;
use litchi::ooxml::xlsb::named_ranges::{Definition, area3d_formula};
use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Data");

    // Title row (merged)
    sheet.set_cell(0, 0, CellValue::String("Sales Report".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 3));

    // Data
    sheet.set_cell(1, 0, CellValue::String("Product".to_string()));
    sheet.set_cell(1, 1, CellValue::String("Sales".to_string()));
    sheet.set_cell(2, 0, CellValue::String("Widget A".to_string()));
    sheet.set_cell(2, 1, CellValue::Float(100.0));
    sheet.set_cell(3, 0, CellValue::String("Widget B".to_string()));
    sheet.set_cell(3, 1, CellValue::Float(200.0));

    workbook.add_worksheet(sheet);

    // Named range for sales data B3:B4
    let sales_formula = area3d_formula(0, 2, 3, 1, 1)?;
    let sales_range = Definition::new("SalesData".to_string(), Some(0)).with_formula(sales_formula);
    workbook.add_named_range(sales_range);

    let file = File::create("xlsb_test_nr_merged.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_nr_merged.xlsb");
    println!("  - Named ranges + Merged cells (NO hyperlinks)");
    Ok(())
}
