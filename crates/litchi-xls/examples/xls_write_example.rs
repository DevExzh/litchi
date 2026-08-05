//! Example demonstrating XLS file writing with the Litchi library
//!
//! NOTE: This example demonstrates the API but will not work until
//! the OLE2 writer infrastructure is complete. See OLE_WRITE_SUPPORT_STATUS.md
//! for implementation status.
//!
//! Run with: cargo run --example xls_write_example
use litchi_xls::Writer;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating a new XLS file...");

    // Create a new XLS writer
    let mut writer = Writer::new();

    // Add first worksheet
    let sheet1 = writer.add_worksheet("Sales Data")?;

    // Write headers
    writer.write_string(sheet1, 0, 0, "Product")?;
    writer.write_string(sheet1, 0, 1, "Quantity")?;
    writer.write_string(sheet1, 0, 2, "Price")?;
    writer.write_string(sheet1, 0, 3, "Total")?;

    // Write data rows
    writer.write_string(sheet1, 1, 0, "Widget A")?;
    writer.write_number(sheet1, 1, 1, 10.0)?;
    writer.write_number(sheet1, 1, 2, 25.50)?;
    writer.write_number(sheet1, 1, 3, 255.0)?;

    writer.write_string(sheet1, 2, 0, "Widget B")?;
    writer.write_number(sheet1, 2, 1, 5.0)?;
    writer.write_number(sheet1, 2, 2, 42.75)?;
    writer.write_number(sheet1, 2, 3, 213.75)?;

    writer.write_string(sheet1, 3, 0, "Widget C")?;
    writer.write_number(sheet1, 3, 1, 15.0)?;
    writer.write_number(sheet1, 3, 2, 18.00)?;
    writer.write_number(sheet1, 3, 3, 270.0)?;

    // Add a totals row
    writer.write_string(sheet1, 4, 0, "Total")?;
    writer.write_number(sheet1, 4, 1, 30.0)?;
    writer.write_number(sheet1, 4, 3, 738.75)?;

    // Add second worksheet with different data
    let sheet2 = writer.add_worksheet("Inventory")?;

    writer.write_string(sheet2, 0, 0, "Item")?;
    writer.write_string(sheet2, 0, 1, "In Stock")?;
    writer.write_string(sheet2, 0, 2, "Reorder?")?;

    writer.write_string(sheet2, 1, 0, "Widget A")?;
    writer.write_number(sheet2, 1, 1, 125.0)?;
    writer.write_boolean(sheet2, 1, 2, false)?;

    writer.write_string(sheet2, 2, 0, "Widget B")?;
    writer.write_number(sheet2, 2, 1, 8.0)?;
    writer.write_boolean(sheet2, 2, 2, true)?;

    writer.write_string(sheet2, 3, 0, "Widget C")?;
    writer.write_number(sheet2, 3, 1, 45.0)?;
    writer.write_boolean(sheet2, 3, 2, false)?;

    // NOTE: Formula support is coming soon
    // writer.write_formula(sheet1, 5, 1, "SUM(B2:B4)")?;

    // Save the file
    println!("Saving to output.xls...");
    writer.save("output.xls")?;

    println!("✅ XLS file created successfully!");
    println!("   - 2 worksheets created");
    println!("   - Multiple data types (strings, numbers, booleans)");
    println!("   - Automatic shared string table optimization");

    Ok(())
}
