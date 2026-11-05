//! Test XLSB with named ranges + hyperlinks (isolate issue)

#![allow(clippy::all)]

use litchi::sheet::CellValue;
use litchi::xlsb::builder::{self, Builder};
use litchi::xlsb::hyperlinks::Hyperlink;
use litchi::xlsb::named_ranges::{Definition, area3d_formula};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Builder::new();
    let mut sheet = builder::Worksheet::new("Data");

    // Data
    sheet.set_cell(0, 0, CellValue::String("Product".to_string()));
    sheet.set_cell(0, 1, CellValue::String("Sales".to_string()));
    sheet.set_cell(0, 2, CellValue::String("Link".to_string()));

    sheet.set_cell(1, 0, CellValue::String("Widget A".to_string()));
    sheet.set_cell(1, 1, CellValue::Float(100.0));
    sheet.set_cell(1, 2, CellValue::String("Details".to_string()));

    // Hyperlink on row 1
    let link = Hyperlink::new_external(1, 2, 1, 2, "https://example.com/widget-a".to_string())
        .with_display("Details".to_string());
    sheet.add_hyperlink(link);

    sheet.set_cell(2, 0, CellValue::String("Widget B".to_string()));
    sheet.set_cell(2, 1, CellValue::Float(200.0));
    sheet.set_cell(2, 2, CellValue::String("Details".to_string()));

    // Hyperlink on row 2
    let link2 = Hyperlink::new_external(2, 2, 2, 2, "https://example.com/widget-b".to_string())
        .with_display("Details".to_string());
    sheet.add_hyperlink(link2);

    workbook.add_worksheet(sheet);

    // Named range for sales data B2:B3
    let sales_formula = area3d_formula(0, 1, 2, 1, 1)?;
    let sales_range = Definition::new("SalesData".to_string(), Some(0)).with_formula(sales_formula);
    workbook.add_named_range(sales_range);

    let file = File::create("xlsb_test_nr_hyperlink.xlsb")?;
    workbook.build()?.write_to(file)?;

    println!("✓ Created xlsb_test_nr_hyperlink.xlsb");
    println!("  - Named ranges + Hyperlinks (testing combination)");
    Ok(())
}
