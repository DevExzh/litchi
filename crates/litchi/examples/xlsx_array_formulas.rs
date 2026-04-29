//! XLSX array formula example
//!
//! Run with:
//!
//! ```bash
//! cargo run --example xlsx_array_formulas --features ooxml -- xlsx_array_formulas.xlsx
//! ```
//!
//! Then open the generated file in Microsoft Excel and verify that:
//! - Cells C2:C4 contain a single array formula spanning the range.
//! - Selecting C2 shows an array/dynamic array formula in the formula bar
//!   (e.g., `{=A2:A4*B2:B4}` in older Excel, or a spilled formula in newer Excel).

use litchi::ooxml::xlsx::Workbook;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let args: Vec<String> = env::args().collect();
    let output = if args.len() > 1 {
        &args[1]
    } else {
        "xlsx_array_formulas.xlsx"
    };

    println!("Creating XLSX with array formulas: {}", output);

    // Create a new workbook and use the default first worksheet.
    let mut wb = Workbook::create()?;
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("ArrayFormulas".to_string());

        // Headers
        ws.set_cell_value(1, 1, "A");
        ws.set_cell_value(1, 2, "B");
        ws.set_cell_value(1, 3, "A*B (array)");

        // Input data in A2:A4 and B2:B4
        ws.set_cell_value(2, 1, 1);
        ws.set_cell_value(3, 1, 2);
        ws.set_cell_value(4, 1, 3);

        ws.set_cell_value(2, 2, 10);
        ws.set_cell_value(3, 2, 20);
        ws.set_cell_value(4, 2, 30);

        // Set an array formula over C2:C4 that multiplies A and B element-wise.
        // Excel should treat C2:C4 as a single array formula range.
        ws.set_array_formula(2, 3, 4, 3, "A2:A4*B2:B4");

        // A normal formula summing the array results, for comparison.
        ws.set_cell_value(6, 1, "Sum of A*B:");
        ws.set_cell_formula(6, 2, "SUM(C2:C4)");
    }

    wb.save(output)?;

    println!("Saved array formula example to: {}", output);
    println!("Open it in Excel and:");
    println!("  - Select C2:C4 to see the array formula region.");
    println!("  - Check the formula bar for the C2 formula.");
    println!("  - Verify that cells C2:C4 evaluate to 10, 40, 90 and that B6 shows their sum.");

    Ok(())
}
