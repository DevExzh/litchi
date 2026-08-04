//! Build a small `.ods` spreadsheet, save it to a tempfile, then reopen it
//! and print the CSV representation.
//!
//! Run with:
//! ```bash
//! cargo run -p litchi-odf --example write_ods
//! ```

use litchi_ods::{Builder, CellValue, Spreadsheet};
use tempfile::NamedTempFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create a tempfile that will be cleaned up when this binding drops.
    let tmp = NamedTempFile::with_suffix(".ods")?;
    let path = tmp.path().to_path_buf();

    // ---- Build phase ----
    println!("Building ODS at: {}", path.display());
    let mut builder = Builder::new();
    builder.add_sheet("Demo")?;

    // Header row.
    builder.add_row_with_values(&["Item", "Quantity", "Price", "Total"])?;

    // Data rows. Row indices below are 0-based, so row 1 is the first data row
    // (spreadsheet row 2 in A1 notation).
    builder.add_row_with_cell_values(&[
        CellValue::Text("Apples".to_string()),
        CellValue::Number(10.0),
        CellValue::Number(0.5),
        CellValue::Empty,
    ])?;
    builder.add_row_with_cell_values(&[
        CellValue::Text("Bread".to_string()),
        CellValue::Number(2.0),
        CellValue::Number(2.25),
        CellValue::Empty,
    ])?;
    builder.add_row_with_cell_values(&[
        CellValue::Text("Cheese".to_string()),
        CellValue::Number(1.0),
        CellValue::Number(4.99),
        CellValue::Empty,
    ])?;

    // Per-row Total formulas.
    builder.set_cell_formula(1, 3, "of:=B2*C2")?;
    builder.set_cell_formula(2, 3, "of:=B3*C3")?;
    builder.set_cell_formula(3, 3, "of:=B4*C4")?;

    // Grand-total row using SUM.
    builder.add_row_with_values(&["Grand Total", "", "", ""])?;
    builder.set_cell_formula(4, 3, "of:=SUM(D2:D4)")?;

    builder.save(&path)?;
    println!(
        "Saved spreadsheet ({} bytes).",
        std::fs::metadata(&path)?.len()
    );

    // ---- Read-back phase ----
    println!("\nReopening for verification...");
    let mut sheet = Spreadsheet::open(&path)?;
    println!("Sheet count: {}", sheet.sheet_count()?);

    let csv = sheet.to_csv()?;
    println!("\n--- CSV output ---\n{}", csv);

    // tmp drops here -> file is deleted.
    Ok(())
}
