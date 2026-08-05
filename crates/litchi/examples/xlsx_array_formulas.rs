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

use litchi::ooxml::xlsx::{Formula, Workbook};
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
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;
    {
        edit.tab(0)?
            .ok_or("default worksheet is missing")?
            .rename("ArrayFormulas")?;
        let mut ws = edit
            .sheet("ArrayFormulas")?
            .ok_or("renamed worksheet is missing")?;

        // Headers
        ws.set("A1", "A")?;
        ws.set("B1", "B")?;
        ws.set("C1", "A*B (array)")?;

        // Input data in A2:A4 and B2:B4
        ws.set("A2", 1_i32)?;
        ws.set("A3", 2_i32)?;
        ws.set("A4", 3_i32)?;

        ws.set("B2", 10_i32)?;
        ws.set("B3", 20_i32)?;
        ws.set("B4", 30_i32)?;

        // The canonical Formula representation stores the dynamic-array
        // anchor only; Excel spills its result from C2 through C4.
        ws.set("C2", Formula::new("A2:A4*B2:B4")?)?;

        // A normal formula summing the array results, for comparison.
        ws.set("A6", "Sum of A*B:")?;
        ws.set("B6", Formula::new("SUM(C2:C4)")?)?;
    }

    edit.commit()?.into_workbook().save(output)?;

    println!("Saved array formula example to: {}", output);
    println!("Open it in Excel and:");
    println!("  - Select C2:C4 to see the array formula region.");
    println!("  - Check the formula bar for the C2 formula.");
    println!("  - Verify that cells C2:C4 evaluate to 10, 40, 90 and that B6 shows their sum.");

    Ok(())
}
