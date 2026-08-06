//! Test XLSB with named ranges + merged cells + hyperlinks

#![allow(clippy::all)]

use litchi::sheet::CellValue;
use litchi::xlsb::hyperlinks::Hyperlink;
use litchi::xlsb::merged_cells::MergedCell;
use litchi::xlsb::named_ranges::{Definition, area3d_formula};
use litchi::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Data");

    // Title row (merged)
    sheet.set_cell(0, 0, CellValue::String("Sales Report".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 3));

    // Headers
    sheet.set_cell(1, 0, CellValue::String("Product".to_string()));
    sheet.set_cell(1, 1, CellValue::String("Sales".to_string()));
    sheet.set_cell(1, 2, CellValue::String("Link".to_string()));

    // Data with hyperlinks
    sheet.set_cell(2, 0, CellValue::String("Widget A".to_string()));
    sheet.set_cell(2, 1, CellValue::Float(100.0));
    sheet.set_cell(2, 2, CellValue::String("Details".to_string()));
    // Hyperlink at cell (row=2, col=2): row_first=2, row_last=2, col_first=2, col_last=2
    let link = Hyperlink::new_external(2, 2, 2, 2, "https://example.com/a".to_string())
        .with_display("Details".to_string());
    sheet.add_hyperlink(link);

    sheet.set_cell(3, 0, CellValue::String("Widget B".to_string()));
    sheet.set_cell(3, 1, CellValue::Float(200.0));
    sheet.set_cell(3, 2, CellValue::String("Details".to_string()));
    // Hyperlink at cell (row=3, col=2): row_first=3, row_last=3, col_first=2, col_last=2
    let link2 = Hyperlink::new_external(3, 3, 2, 2, "https://example.com/b".to_string())
        .with_display("Details".to_string());
    sheet.add_hyperlink(link2);

    workbook.add_worksheet(sheet);

    // Named range for sales data B3:B4
    let sales_formula = area3d_formula(0, 2, 3, 1, 1)?;
    let sales_range = Definition::new("SalesData".to_string(), Some(0)).with_formula(sales_formula);
    workbook.add_named_range(sales_range);

    let file = File::create("xlsb_test_all_three.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_all_three.xlsb");
    println!("  - Named ranges + Merged cells + Hyperlinks");
    Ok(())
}
