//! Test XLSB with a single hyperlink feature
//!
//! NOTE: This uses the current Hyperlink writer and is expected to help
//! diagnose hyperlink-related issues. It intentionally mirrors the
//! `xlsb_writer` advanced sheet pattern in a much smaller workbook.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_hyperlinks --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::hyperlinks::Hyperlink;
use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test H: Creating XLSB with one hyperlink...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Hyperlinks");

    // Visible cell text
    sheet.set_cell(0, 0, CellValue::String("Details".to_string()));

    // Hyperlink covering A1 pointing to an external URL. The writer will
    // create an external OPC relationship and assign a concrete rId
    // automatically before serializing the BrtHLink record.
    let link = Hyperlink::new_external(0, 0, 0, 0, "https://example.com/details".to_string())
        .with_display("Details".to_string());
    sheet.add_hyperlink(link);

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_hyperlinks.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_hyperlinks.xlsb");
    Ok(())
}
