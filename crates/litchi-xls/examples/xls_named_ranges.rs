//! Example demonstrating XLS named ranges with the Litchi library.
//!
//! This example generates a small `.xls` file with a few named ranges so
//! you can open it in Microsoft Excel and verify the definitions via the
//! **Formulas → Name Manager** UI.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example xls_named_ranges --features ole --no-default-features
//! ```

use litchi_xls::XlsWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating xls_named_ranges.xls with named ranges...");

    // Create a new XLS writer and a single worksheet.
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("NamedRanges")?;

    // Populate a small table so the named ranges have visible content.
    //
    // Layout (0-based indices shown as (row, col)):
    //   (0,0) = "Item"   (0,1) = "Value"
    //   (1,0) = "A"      (1,1) = 10
    //   (2,0) = "B"      (2,1) = 20
    //   (3,0) = "C"      (3,1) = 30
    writer.write_string(sheet, 0, 0, "Item")?;
    writer.write_string(sheet, 0, 1, "Value")?;

    writer.write_string(sheet, 1, 0, "A")?;
    writer.write_number(sheet, 1, 1, 10.0)?;

    writer.write_string(sheet, 2, 0, "B")?;
    writer.write_number(sheet, 2, 1, 20.0)?;

    writer.write_string(sheet, 3, 0, "C")?;
    writer.write_number(sheet, 3, 1, 30.0)?;

    // Define a sheet-scoped named range for the numeric values (B2:B4).
    //
    // Note: The public named-range API currently accepts only simple
    // A1-style references without sheet qualifiers (e.g., "B2:B4").
    writer.define_name_local("ValueRange", "B2:B4", sheet)?;

    // Define a workbook-scoped named range that covers the whole table
    // including headers (A1:B4).
    writer.define_name("ItemsTable", "A1:B4")?;

    // Define another workbook-scoped named range with a description.
    writer.define_name_with_comment(
        "ValueRangeWithComment",
        "B2:B4",
        "Example named range over the Value column",
    )?;

    // Save the workbook.
    let output_path = "xls_named_ranges.xls";
    println!("Saving to {output_path}...");
    writer.save(output_path)?;

    println!("✅ Done.");
    println!("Open '{output_path}' in Microsoft Excel and:");
    println!("  1. Go to the 'NamedRanges' sheet to see the sample data.");
    println!("  2. Use Formulas → Name Manager to inspect the defined names:");
    println!("     - 'ValueRange' (sheet-scoped) should highlight B2:B4.");
    println!("     - 'ItemsTable' should cover A1:B4.");
    println!("     - 'ValueRangeWithComment' should reuse B2:B4 with a comment.");

    Ok(())
}
