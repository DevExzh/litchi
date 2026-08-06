//! Test XLSB internal hyperlinks (within the workbook)
//!
//! This example creates a workbook with two worksheets. Cell A1 on the
//! "Source" sheet contains a hyperlink that jumps to cell B2 on the
//! "Target" sheet using an internal `BrtHLink` location string.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_internal_hyperlinks --features ooxml --no-default-features
//! ```

use litchi::sheet::CellValue;
use litchi::xlsb::hyperlinks::Hyperlink;
use litchi::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating XLSB with an internal hyperlink (Source!A1 -> Target!B2)...");

    let mut workbook = WorkbookWriter::new();

    // Source sheet with a clickable cell.
    let mut source = MutableWorksheet::new("Source");
    source.set_cell(0, 0, CellValue::String("Go to Target!B2".to_string()));

    // Internal hyperlink from A1 on "Source" to B2 on "Target". The writer
    // will encode the location in the BrtHLink record without creating an
    // external OPC relationship.
    let link = Hyperlink::new_internal(0, 0, 0, 0, "Target!B2".to_string())
        .with_tooltip("Jump to Target!B2".to_string());
    source.add_hyperlink(link);
    workbook.add_worksheet(source);

    // Target sheet with a visible marker at B2.
    let mut target = MutableWorksheet::new("Target");
    target.set_cell(1, 1, CellValue::String("You reached Target!B2".to_string()));
    workbook.add_worksheet(target);

    let file = File::create("xlsb_test_internal_hyperlinks.xlsb")?;
    workbook.save(file)?;

    println!("\n[32m✓ Created xlsb_test_internal_hyperlinks.xlsb[0m");
    Ok(())
}
