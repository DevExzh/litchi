//! Transactional XLSX workbook authoring example.
//!
//! Cell values, formulas, merges, rows, columns, and sheet names use the
//! standalone immutable-snapshot/transaction facade. Formatting, validation,
//! named-range, and drawing values are shown as typed models where their
//! worksheet package attachment is not exposed by that facade yet.

use litchi_xlsx::Formula;
use litchi_xlsx::Workbook;
use litchi_xlsx::data_validation::{
    Collection, ListSource, Source, Sqref, Validation, ValidationType,
};
use litchi_xlsx::style::format::{CellFill, CellFillPatternType, CellFont, CellFormat};
use litchi_xlsx::style::stylesheet::alignment::{Alignment, Horizontal, Vertical};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let output_file = env::args()
        .nth(1)
        .unwrap_or_else(|| "output.xlsx".to_string());
    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;

    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("Sales Data")?;
    {
        let mut sales = edit
            .sheet("Sales Data")?
            .ok_or("Sales Data worksheet is missing")?;
        for (cell, value) in [
            ("A1", "Product"),
            ("B1", "Quantity"),
            ("C1", "Price"),
            ("D1", "Total"),
        ] {
            sales.set(cell, value)?;
        }
        for (row, product, quantity, price) in [
            (2, "Laptops", 150_i32, "999.99"),
            (3, "Mice", 500_i32, "25.50"),
            (4, "Keyboards", 300_i32, "79.99"),
        ] {
            sales.set((row, 0), product)?;
            sales.set((row, 1), quantity)?;
            sales.set((row, 2), litchi_xlsx::Number::new(price)?)?;
            sales.set((row, 3), Formula::new(format!("B{row}*C{row}"))?)?;
        }
        sales
            .set("A5", "TOTAL")?
            .set("D5", Formula::new("SUM(D2:D4)")?)?;
        sales.column("A")?.width(15.0)?;
        sales.column("B")?.width(10.0)?;
        sales.column("C")?.width(10.0)?;
        sales.column("D")?.width(12.0)?;
        sales.row(1)?.height(20.0)?.thick_bottom();
    }

    let mut report = edit.add("Quarterly Report")?;
    report.set("A1", "Q4 2024 Sales Report")?.merge("A1:D1")?;
    let mut form = edit.add("Input Form")?;
    form.set("A1", "Employee Name:")?
        .set("A2", "Department:")?
        .set("A3", "Status:")?
        .set("A4", "Rating (1-5):")?;
    let mut chart_data = edit.add("Chart Data")?;
    chart_data.set("A1", "Month")?.set("B1", "Sales")?;
    for (row, (month, sales_value)) in [
        ("Jan", 12_000_i32),
        ("Feb", 15_000),
        ("Mar", 18_000),
        ("Apr", 16_000),
        ("May", 21_000),
        ("Jun", 24_000),
    ]
    .iter()
    .enumerate()
    {
        chart_data.set(((row + 2) as u32, 0), *month)?;
        chart_data.set(((row + 2) as u32, 1), *sales_value)?;
    }

    let header_format = CellFormat {
        font: Some(CellFont {
            name: Some("Calibri".to_string()),
            size: Some(12.0),
            bold: true,
            ..CellFont::default()
        }),
        fill: Some(CellFill {
            pattern_type: CellFillPatternType::Solid,
            fg_color: Some("FFD3D3D3".to_string()),
            bg_color: None,
        }),
        ..CellFormat::default()
    };
    let title_alignment = Alignment::both(Horizontal::Center, Vertical::Center);

    let mut department = Validation::new(Source::Core, ValidationType::List, Sqref::parse("B2")?);
    department.set_formula1(Some(ListSource::QuotedList(
        "Engineering,Sales,Marketing,HR".to_string(),
    )))?;
    let validations = Collection::new(Source::Core, vec![department])?;
    let _typed_models = (header_format, title_alignment, validations);

    let workbook = edit.commit()?.into_workbook();
    workbook.save(&output_file)?;
    println!("Created transactional workbook at {output_file}");
    println!("Supported: cells, formulas, merges, dimensions, and tab names");
    println!("Typed model demonstrations: formatting, validation, and chart source data");
    Ok(())
}
