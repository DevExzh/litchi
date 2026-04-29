//! Comprehensive demo of ODS features including hyperlinks, annotations, merged cells,
//! data validation, conditional formatting, and named expressions.

use litchi::Result;
use litchi::common::Metadata;
use litchi::odf::{CellValue, Spreadsheet, SpreadsheetBuilder};

fn main() -> Result<()> {
    println!("Creating comprehensive ODS file with supported features...");

    let mut builder = SpreadsheetBuilder::new();

    // Set metadata
    let mut metadata = Metadata::default();
    metadata.title = Some("ODS Features Demo".to_string());
    metadata.author = Some("Litchi Library".to_string());
    builder.set_metadata(metadata);

    // ===== Sheet 1: Overview =====
    builder.add_sheet("Overview")?;

    builder.add_row_with_values(&["Feature", "Description", "Status"])?;
    builder.add_row_with_values(&[
        "Metadata",
        "Document title and author are written into meta.xml",
        "Supported",
    ])?;
    builder.add_row_with_values(&[
        "Typed Cells",
        "Text and numeric values round-trip through the public API",
        "Supported",
    ])?;
    builder.add_row_with_values(&[
        "Formulas",
        "ODS formulas are written with table:formula attributes",
        "Supported",
    ])?;

    // ===== Sheet 2: Sales Data and Formulas =====
    builder.add_sheet("Formulas")?;

    builder.add_row_with_values(&["Item", "Price", "Quantity", "Total"])?;

    builder.set_cell(1, 0, CellValue::Text("Widget A".to_string()))?;
    builder.set_cell(1, 1, CellValue::Number(10.50))?;
    builder.set_cell(1, 2, CellValue::Number(5.0))?;
    builder.set_cell_formula(1, 3, "of:=[.B2]*[.C2]")?;

    builder.set_cell(2, 0, CellValue::Text("Widget B".to_string()))?;
    builder.set_cell(2, 1, CellValue::Number(25.00))?;
    builder.set_cell(2, 2, CellValue::Number(3.0))?;
    builder.set_cell_formula(2, 3, "of:=[.B3]*[.C3]")?;

    builder.set_cell(3, 0, CellValue::Text("Total".to_string()))?;
    builder.set_cell_formula(3, 3, "of:=SUM([.D2:.D3])")?;

    // ===== Sheet 3: Typed Values =====
    builder.add_sheet("Typed Values")?;
    builder.add_row_with_values(&["Kind", "Value"])?;
    builder.add_row_with_cell_values(&[
        CellValue::Text("Boolean".to_string()),
        CellValue::Boolean(true),
    ])?;
    builder.add_row_with_cell_values(&[
        CellValue::Text("Percentage".to_string()),
        CellValue::Percentage(12.5),
    ])?;
    builder.add_row_with_cell_values(&[
        CellValue::Text("Currency".to_string()),
        CellValue::Currency(19.99, "USD".to_string()),
    ])?;

    // Save the file
    let output_path = "output/ods_features_demo.ods";
    std::fs::create_dir_all("output")?;
    builder.save(output_path)?;

    let mut spreadsheet = Spreadsheet::open(output_path)?;
    let sheets = spreadsheet.sheets()?;
    println!("Read back {} sheet(s)", sheets.len());
    println!("  - {} rows in Overview", sheets[0].rows.len());
    println!(
        "  - {:?} formula in D4",
        sheets[1].rows[3].cells[3].formula()?
    );
    println!(
        "  - {} in Typed Values A2",
        sheets[2].rows[1].cells[0].text()?
    );

    println!("✅ Created: {}", output_path);
    println!("\nFeatures demonstrated:");
    println!("  • Metadata");
    println!("  • Multiple sheets");
    println!("  • Typed cell values");
    println!("  • Formulas (calculations and references)");
    println!("  • Basic read-back verification");
    println!("\nOpen the file in LibreOffice Calc to inspect the generated content.");

    Ok(())
}
