//! Minimal XLSX writer example focusing on page setup features.
//!
//! This example creates a workbook that demonstrates:
//! - Print area (`set_print_area`)
//! - Repeating header rows (`set_repeating_rows`)
//! - Repeating header columns (`set_repeating_columns`)
//!
//! Usage:
//!
//! ```bash
//! cargo run --example xlsx_print_setup --features ooxml -- xlsx_print_setup.xlsx
//! ```
//!
//! Then open the generated file in Microsoft Excel and check:
//! - Page Layout → Print Titles (rows/columns to repeat)
//! - Page Layout → Print Area
//! - File → Print / Print Preview to confirm repeated titles.

use litchi::ooxml::xlsx::Workbook;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    // Determine output path from CLI or use a default.
    let args: Vec<String> = env::args().collect();
    let output_path = if args.len() > 1 {
        args[1].as_str()
    } else {
        "xlsx_print_setup.xlsx"
    };

    println!("Creating XLSX with print area and repeating rows/columns...");

    // Create a new workbook and use the default first worksheet.
    let mut workbook = Workbook::create()?;
    {
        let ws = workbook.worksheet_mut(0)?;
        ws.set_name("PrintDemo".to_string());

        // Header row that we want repeated on every printed page.
        ws.set_cell_value(1, 1, "ID");
        ws.set_cell_value(1, 2, "Name");
        ws.set_cell_value(1, 3, "Department");
        ws.set_cell_value(1, 4, "Value");

        // A few dozen rows to force multiple printed pages.
        for row in 2..=80 {
            ws.set_cell_value(row, 1, row - 1); // ID
            ws.set_cell_value(row, 2, format!("Employee {row}"));
            ws.set_cell_value(row, 3, if row % 2 == 0 { "Engineering" } else { "Sales" });
            ws.set_cell_value(row, 4, (row as i64 - 1) * 10);
        }

        // Define a print area that covers the header + data region.
        // Note: A1:D80 uses A1-style notation, matching Excel expectations.
        ws.set_print_area("A1:D80");

        // Repeat the first row on every printed page.
        // Excel will interpret this as the row range 1:1.
        ws.set_repeating_rows("$1:$1");

        // Also repeat the first column (ID) on each page for clarity.
        // Excel expects column ranges like $A:$A.
        ws.set_repeating_columns("$A:$A");
    }

    // Save the workbook.
    workbook.save(output_path)?;

    println!("Saved XLSX to: {output_path}");
    println!("Open this file in Excel and verify:");
    println!("  - Page Layout → Print Area shows A1:D80");
    println!("  - Page Layout → Print Titles shows rows $1:$1 and columns $A:$A");
    println!("  - File → Print / Print Preview repeats the header row and ID column.");

    Ok(())
}
