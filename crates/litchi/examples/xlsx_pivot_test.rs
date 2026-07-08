//! Pivot-table showcase for XLSX writer.
//!
//! This example creates two workbooks that contain fully wired pivot tables so
//! that you can open them in Microsoft Excel (or LibreOffice) for validation.
//!
//! ```bash
//! cargo run --example xlsx_pivot_test -- pivot_examples
//! # -> pivot_examples/pivot_basic.xlsx
//! # -> pivot_examples/pivot_with_filters.xlsx
//! ```

use litchi::ooxml::pivot::{
    PivotAxis, PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction,
};
use litchi::ooxml::xlsx::Workbook;
use std::env;
use std::fs;
use std::path::Path;

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    let output_dir = env::args()
        .nth(1)
        .unwrap_or_else(|| "pivot_examples".to_string());
    fs::create_dir_all(&output_dir)?;

    let basic_path = Path::new(&output_dir).join("pivot_basic.xlsx");
    let filtered_path = Path::new(&output_dir).join("pivot_with_filters.xlsx");

    generate_basic_example(&basic_path)?;
    generate_filter_example(&filtered_path)?;

    println!("Generated pivot workbooks:");
    println!("  - {}", basic_path.display());
    println!("  - {}", filtered_path.display());
    println!("Open them in Microsoft Excel to verify the pivot tables.");
    Ok(())
}

fn generate_basic_example(path: &Path) -> ExampleResult<()> {
    let mut wb = Workbook::create()?;
    let sheet = wb.worksheet_mut(0)?;
    sheet.set_name("SalesData".to_string());

    sheet.set_cell_value(1, 1, "Product");
    sheet.set_cell_value(1, 2, "Region");
    sheet.set_cell_value(1, 3, "Quarter");
    sheet.set_cell_value(1, 4, "Amount");

    let records = [
        ("Laptops", "North", "Q1", 120_000.0),
        ("Laptops", "North", "Q2", 98_500.0),
        ("Laptops", "South", "Q1", 76_200.0),
        ("Laptops", "South", "Q2", 88_450.0),
        ("Tablets", "North", "Q1", 54_300.0),
        ("Tablets", "North", "Q2", 61_125.0),
        ("Tablets", "South", "Q1", 42_000.0),
        ("Tablets", "South", "Q2", 39_875.0),
        ("Accessories", "North", "Q1", 15_200.0),
        ("Accessories", "North", "Q2", 12_950.0),
        ("Accessories", "South", "Q1", 10_400.0),
        ("Accessories", "South", "Q2", 9_875.0),
    ];
    for (idx, (product, region, quarter, amount)) in records.iter().enumerate() {
        let row = (idx + 2) as u32;
        sheet.set_cell_value(row, 1, *product);
        sheet.set_cell_value(row, 2, *region);
        sheet.set_cell_value(row, 3, *quarter);
        sheet.set_cell_value(row, 4, *amount);
    }

    wb.add_worksheet("PivotSummary");
    wb.add_pivot_table(PivotTable {
        name: "SalesByProductQuarter".into(),
        source_sheet: Some("SalesData".into()),
        source_ref: Some("A1:D13".into()),
        field_names: vec![
            "Product".into(),
            "Region".into(),
            "Quarter".into(),
            "Amount".into(),
        ],
        sheet_name: "PivotSummary".into(),
        cache_id: 0,
        location_ref: "A3".into(),
        row_fields: vec![PivotFieldRole {
            field_name: "Product".into(),
            axis: PivotAxis::Row,
            position: 0,
        }],
        column_fields: vec![PivotFieldRole {
            field_name: "Quarter".into(),
            axis: PivotAxis::Column,
            position: 0,
        }],
        filter_fields: vec![],
        data_fields: vec![PivotDataField {
            field_name: "Amount".into(),
            function: PivotValueFunction::Sum,
            display_name: Some("Total Amount".into()),
        }],
    })?;

    wb.save(path)?;
    Ok(())
}

fn generate_filter_example(path: &Path) -> ExampleResult<()> {
    let mut wb = Workbook::create()?;
    let sheet = wb.worksheet_mut(0)?;
    sheet.set_name("Opportunities".to_string());

    sheet.set_cell_value(1, 1, "Salesperson");
    sheet.set_cell_value(1, 2, "Channel");
    sheet.set_cell_value(1, 3, "Region");
    sheet.set_cell_value(1, 4, "Year");
    sheet.set_cell_value(1, 5, "Revenue");

    let rows = [
        ("Alice", "Online", "East", 2024, 45_000.0),
        ("Alice", "Retail", "East", 2025, 52_000.0),
        ("Bob", "Online", "West", 2024, 38_000.0),
        ("Bob", "Retail", "West", 2025, 41_500.0),
        ("Carol", "Online", "East", 2024, 33_750.0),
        ("Carol", "Retail", "East", 2025, 36_200.0),
        ("Dave", "Online", "West", 2024, 29_600.0),
        ("Dave", "Retail", "West", 2025, 31_400.0),
    ];
    for (idx, (salesperson, channel, region, year, revenue)) in rows.iter().enumerate() {
        let row = (idx + 2) as u32;
        sheet.set_cell_value(row, 1, *salesperson);
        sheet.set_cell_value(row, 2, *channel);
        sheet.set_cell_value(row, 3, *region);
        sheet.set_cell_value(row, 4, *year as i64);
        sheet.set_cell_value(row, 5, *revenue);
    }

    wb.add_worksheet("PivotWithFilters");
    wb.add_pivot_table(PivotTable {
        name: "RevenueByChannel".into(),
        source_sheet: Some("Opportunities".into()),
        source_ref: Some("A1:E9".into()),
        field_names: vec![
            "Salesperson".into(),
            "Channel".into(),
            "Region".into(),
            "Year".into(),
            "Revenue".into(),
        ],
        sheet_name: "PivotWithFilters".into(),
        cache_id: 0,
        location_ref: "A3".into(),
        row_fields: vec![PivotFieldRole {
            field_name: "Region".into(),
            axis: PivotAxis::Row,
            position: 0,
        }],
        column_fields: vec![PivotFieldRole {
            field_name: "Channel".into(),
            axis: PivotAxis::Column,
            position: 0,
        }],
        filter_fields: vec![PivotFieldRole {
            field_name: "Year".into(),
            axis: PivotAxis::Filter,
            position: 0,
        }],
        data_fields: vec![PivotDataField {
            field_name: "Revenue".into(),
            function: PivotValueFunction::Sum,
            display_name: Some("Revenue".into()),
        }],
    })?;

    wb.save(path)?;
    Ok(())
}
