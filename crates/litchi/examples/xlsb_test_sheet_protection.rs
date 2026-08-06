//! Test XLSB sheet protection flags
//!
//! This example enables basic sheet protection on a single worksheet using the
//! XLSB writer-side `SheetProtection` structure. It does not set a password
//! hash, so protection can be cleared in Excel without a password.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_sheet_protection --features ooxml --no-default-features
//! ```

use litchi::sheet::CellValue;
use litchi::xlsb::writer::{MutableWorksheet, SheetProtection, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating XLSB with basic sheet protection enabled...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Protected");

    sheet.set_cell(
        0,
        0,
        CellValue::String("This sheet is protected (no password).".to_string()),
    );
    sheet.set_cell(
        2,
        0,
        CellValue::String("Try editing, inserting rows/columns, etc.".to_string()),
    );

    // Configure protection flags. We allow selecting locked and unlocked cells
    // but disallow most structural changes (formatting, inserting/deleting
    // rows/columns, sorting, etc.).
    let protection = SheetProtection {
        select_locked_cells: Some(true),
        select_unlocked_cells: Some(true),
        format_cells: Some(false),
        format_columns: Some(false),
        format_rows: Some(false),
        insert_columns: Some(false),
        insert_rows: Some(false),
        insert_hyperlinks: Some(false),
        delete_columns: Some(false),
        delete_rows: Some(false),
        sort: Some(false),
        auto_filter: Some(false),
        pivot_tables: Some(false),
        ..Default::default()
    };

    sheet.set_sheet_protection(Some(protection));

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_sheet_protection.xlsb")?;
    workbook.save(file)?;

    println!("\n[32m✓ Created xlsb_test_sheet_protection.xlsb[0m");
    Ok(())
}
