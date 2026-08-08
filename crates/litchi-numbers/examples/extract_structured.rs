//! Extract globally ordered compatibility tables from an Apple Numbers file.
//!
//! ```text
//! cargo run -p litchi-numbers --example extract_structured -- workbook.numbers
//! ```

#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally renders table content"
)]

use std::{error::Error, path::PathBuf};

use litchi_numbers::Package;

fn main() -> Result<(), Box<dyn Error>> {
    let path = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .ok_or("usage: extract_structured <spreadsheet.numbers>")?;
    let package = Package::open(path)?;
    let tables = package.extract_structured_tables()?;

    println!("found {} compatibility table(s)", tables.len());
    let Some(first) = tables.first() else {
        println!("no tables found in this document");
        return Ok(());
    };

    println!(
        "--- table: {} ({} rows x {} cols) ---",
        first.name(),
        first.row_count(),
        first.column_count()
    );
    let csv = first.to_csv();
    if csv.is_empty() {
        println!("(table is empty)");
    } else {
        println!("{csv}");
    }

    if tables.len() > 1 {
        println!("... and {} more table(s) not shown", tables.len() - 1);
    }
    Ok(())
}
