//! Comprehensive XLSB Writer Example
//!
//! Demonstrates all XLSB writing capabilities.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_writer --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::comments::Comment;
use litchi::ooxml::xlsb::hyperlinks::Hyperlink;
use litchi::ooxml::xlsb::merged_cells::MergedCell;
use litchi::ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() {
    println!("{}", "=".repeat(80));
    println!("XLSB Writer - Comprehensive Feature Test");
    println!("{}", "=".repeat(80));
    println!();

    let output_file = "xlsb_writer_output.xlsb";

    match create_workbook(output_file) {
        Ok(_) => {
            println!("\n{}", "=".repeat(80));
            println!("✓ Successfully created: {}", output_file);
            println!("{}", "=".repeat(80));
        },
        Err(e) => {
            eprintln!("\n✗ Error creating workbook: {}", e);
            std::process::exit(1);
        },
    }
}

fn create_workbook(filename: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = XlsbWorkbookWriter::new();
    workbook.set_date_system(false); // 1900 date system

    println!("Creating worksheets...\n");

    // Sheet 1: Basic data types
    create_basic_sheet(&mut workbook)?;

    // Sheet 2: Formulas
    create_formula_sheet(&mut workbook)?;

    // Sheet 3: Advanced features
    create_advanced_sheet(&mut workbook)?;

    // Save
    println!("Writing workbook to file...");
    let file = File::create(filename)?;
    workbook.save(file)?;

    Ok(())
}

fn create_basic_sheet(workbook: &mut XlsbWorkbookWriter) -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating Sheet 1: Basic Data Types...");

    let mut sheet = MutableXlsbWorksheet::new("BasicData");

    // Headers
    sheet.set_cell(0, 0, CellValue::String("Type".to_string()));
    sheet.set_cell(0, 1, CellValue::String("Value".to_string()));

    // Various types
    sheet.set_cell(1, 0, CellValue::String("Integer".to_string()));
    sheet.set_cell(1, 1, CellValue::Int(42));

    sheet.set_cell(2, 0, CellValue::String("Float".to_string()));
    sheet.set_cell(2, 1, CellValue::Float(std::f64::consts::PI));

    sheet.set_cell(3, 0, CellValue::String("String".to_string()));
    sheet.set_cell(3, 1, CellValue::String("Hello, World!".to_string()));

    sheet.set_cell(4, 0, CellValue::String("Boolean".to_string()));
    sheet.set_cell(4, 1, CellValue::Bool(true));

    sheet.set_cell(5, 0, CellValue::String("Date".to_string()));
    sheet.set_cell(5, 1, CellValue::DateTime(45604.0)); // Nov 7, 2024

    workbook.add_worksheet(sheet);
    println!("  ✓ Added 6 rows");

    Ok(())
}

fn create_formula_sheet(
    workbook: &mut XlsbWorkbookWriter,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating Sheet 2: Formulas...");

    let mut sheet = MutableXlsbWorksheet::new("Formulas");

    // Data
    sheet.set_cell(0, 0, CellValue::Float(10.0));
    sheet.set_cell(0, 1, CellValue::Float(20.0));
    sheet.set_cell(
        0,
        2,
        CellValue::Formula {
            formula: "A1+B1".to_string(),
            cached_value: Some(Box::new(CellValue::Float(30.0))),
            is_array: false,
            array_range: None,
        },
    );

    sheet.set_cell(1, 0, CellValue::Float(5.0));
    sheet.set_cell(1, 1, CellValue::Float(15.0));
    sheet.set_cell(
        1,
        2,
        CellValue::Formula {
            formula: "SUM(A2:B2)".to_string(),
            cached_value: Some(Box::new(CellValue::Float(20.0))),
            is_array: false,
            array_range: None,
        },
    );

    workbook.add_worksheet(sheet);
    println!("  ✓ Added formulas");

    Ok(())
}

fn create_advanced_sheet(
    workbook: &mut XlsbWorkbookWriter,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating Sheet 3: Advanced Features...");

    let mut sheet = MutableXlsbWorksheet::new("Advanced");

    // Title with merge
    sheet.set_cell(0, 0, CellValue::String("Sales Report".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 3));

    // Data
    sheet.set_cell(2, 0, CellValue::String("October".to_string()));
    sheet.set_cell(2, 1, CellValue::Float(125000.0));

    // Hyperlink to external details page
    let link = Hyperlink::new_external(3, 3, 0, 0, "https://example.com/details".to_string())
        .with_display("Details".to_string());
    sheet.set_cell(3, 0, CellValue::String("Link to details".to_string()));
    sheet.add_hyperlink(link);

    // Comment
    let comment = Comment::new(
        2,
        1,
        "Manager".to_string(),
        "Great performance!".to_string(),
    );
    sheet.add_comment(comment);

    workbook.add_worksheet(sheet);
    println!("  ✓ Added merged cells, hyperlinks, and comments");

    Ok(())
}
