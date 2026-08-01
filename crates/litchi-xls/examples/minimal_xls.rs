//! Minimal XLS writer example for Litchi
//!
//! Generates a tiny BIFF8 .xls workbook using the low-level OLE/XLS writer
//! and saves it as `minimal.xls` in the project root.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example minimal_xls --features ole --no-default-features
//! ```

use litchi_xls::XlsWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal.xls via XlsWriter...");

    // Create a new XLS writer
    let mut writer = XlsWriter::new();

    // Single worksheet with a few basic cells
    let sheet = writer.add_worksheet("Sheet1")?;

    // Header row
    writer.write_string(sheet, 0, 0, "Type")?;
    writer.write_string(sheet, 0, 1, "Value")?;

    // String cell
    writer.write_string(sheet, 1, 0, "String")?;
    writer.write_string(sheet, 1, 1, "Hello, XLS!")?;

    // Number cell
    writer.write_string(sheet, 2, 0, "Number")?;
    writer.write_number(sheet, 2, 1, 42.5)?;

    // Boolean cell
    writer.write_string(sheet, 3, 0, "Boolean")?;
    writer.write_boolean(sheet, 3, 1, true)?;

    // Save workbook into an OLE compound file
    let output = "minimal.xls";
    println!("Saving to {output}...");
    writer.save(output)?;

    println!("Done. Open '{output}' in Microsoft Excel to verify.");

    Ok(())
}
