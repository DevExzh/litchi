#![allow(
    clippy::print_stdout,
    reason = "This command-line example intentionally renders the semantic package projection."
)]

use std::error::Error;
use std::path::PathBuf;

use litchi_numbers::Package;

fn main() -> Result<(), Box<dyn Error>> {
    let path = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .ok_or("usage: read_numbers <spreadsheet.numbers>")?;
    let package = Package::open(&path)?;

    println!("{} rooted sheet(s)", package.sheets().len());
    for sheet in package.sheets() {
        println!("sheet {}: {}", sheet.index(), sheet.name());
        for table in sheet.tables() {
            println!(
                "  table {:?}: {}x{}, {} materialized cell(s)",
                table.name(),
                table.row_count(),
                table.column_count(),
                table.cell_count()
            );
        }
    }

    let compatibility = package.extract_structured_tables()?;
    println!("{} compatibility table(s)", compatibility.len());
    for (index, table) in compatibility.iter().enumerate() {
        println!("  compatibility table {index}: {:?}", table.name());
    }
    Ok(())
}
