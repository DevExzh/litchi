//! Open an `.xlsx` workbook, list its worksheets, and print a small grid of
//! cell values from each sheet.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-xlsx --example read_xlsx
//! cargo run -p litchi-xlsx --example read_xlsx -- path/to/file.xlsx
//! ```
//!
//! Default input: `test-data/ooxml/xlsx/ExcelTables.xlsx`.

use std::env;

use litchi_xlsx::{Workbook, WorksheetKind};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() >= 2 {
        args[1].clone()
    } else {
        "test-data/ooxml/xlsx/ExcelTables.xlsx".to_string()
    };

    println!("Opening XLSX: {}", path);
    let wb = Workbook::open(&path)?;

    let sheets = wb.sheets().collect::<Vec<_>>();
    let names = sheets.iter().map(|sheet| sheet.name()).collect::<Vec<_>>();
    println!("Worksheet count: {}", sheets.len());
    println!("Worksheet names: {:?}", names);
    println!(
        "Active sheet: {:?}",
        wb.active_sheet().map(|sheet| sheet.name().to_owned())
    );

    for (index, ws) in sheets.into_iter().enumerate() {
        println!("\n--- Sheet [{}]: {:?} ---", index, ws.name());

        if ws.kind() != WorksheetKind::Worksheet {
            println!("(non-worksheet sheet)");
            continue;
        }

        let extents = ws.extents()?;
        let dims = extents.used().or(extents.declared());
        println!("dimensions (used or declared): {:?}", dims);

        let Some(dims) = dims else {
            println!("(empty sheet)");
            continue;
        };

        for (address, cell) in ws.cells(dims)?.take(30) {
            println!(
                "  ({}, {}): {cell:?}",
                address.row().get() + 1,
                address.column().get() + 1
            );
        }
    }

    Ok(())
}
