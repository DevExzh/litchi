//! Open an `.xlsx` workbook, list its worksheets, and print a small grid of
//! cell values from each sheet.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-ooxml --example read_xlsx --all-features
//! cargo run -p litchi-ooxml --example read_xlsx --all-features -- path/to/file.xlsx
//! ```
//!
//! Default input: `test-data/ooxml/xlsx/ExcelTables.xlsx`.

use std::env;

use litchi_core::sheet::WorkbookTrait;
use litchi_ooxml::xlsx::Workbook;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() >= 2 {
        args[1].clone()
    } else {
        "test-data/ooxml/xlsx/ExcelTables.xlsx".to_string()
    };

    println!("Opening XLSX: {}", path);
    let wb = Workbook::open(&path).map_err(|e| -> Box<dyn std::error::Error> { e })?;

    let names = wb.worksheet_names();
    println!("Worksheet count: {}", wb.worksheet_count());
    println!("Worksheet names: {:?}", names);
    println!("Active sheet index: {}", wb.active_sheet_index());

    // Cap how many rows/cols we print so output stays bounded for any input.
    const MAX_ROWS: u32 = 5;
    const MAX_COLS: u32 = 6;

    for index in 0..wb.worksheet_count() {
        let ws = wb
            .worksheet_by_index(index)
            .map_err(|e| -> Box<dyn std::error::Error> { e })?;
        println!("\n--- Sheet [{}]: {:?} ---", index, ws.name());

        let dims = ws.dimensions();
        println!(
            "dimensions (min_row, min_col, max_row, max_col): {:?}",
            dims
        );

        let Some((min_r, min_c, max_r, max_c)) = dims else {
            println!("(empty sheet)");
            continue;
        };

        let row_end = max_r.min(min_r + MAX_ROWS - 1);
        let col_end = max_c.min(min_c + MAX_COLS - 1);

        for row in min_r..=row_end {
            for col in min_c..=col_end {
                match ws.cell_value(row, col) {
                    Ok(value) => print!("  ({},{})={:?}", row, col, *value),
                    Err(err) => print!("  ({},{})=<err: {}>", row, col, err),
                }
            }
            println!();
        }

        if row_end < max_r || col_end < max_c {
            println!(
                "  ... ({} more rows, {} more cols not shown)",
                max_r.saturating_sub(row_end),
                max_c.saturating_sub(col_end)
            );
        }
    }

    Ok(())
}
