//! Resize an existing Numbers table by workbook-wide table index.
//!
//! The focused resize transaction rewrites the trailing table extent and
//! reopens the saved output before this command reports success. Its inverse
//! is checked against the exact input package.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::env;
use std::error::Error;
use std::io;

use litchi_numbers::{Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or_else(|| invalid("missing input path"))?;
    let output = arguments
        .next()
        .ok_or_else(|| invalid("missing output path"))?;
    let table_index: usize = arguments
        .next()
        .ok_or_else(|| invalid("missing table index"))?
        .parse()
        .map_err(|_| invalid("table index must be a non-negative integer"))?;
    let rows: u32 = arguments
        .next()
        .ok_or_else(|| invalid("missing row count"))?
        .parse()
        .map_err(|_| invalid("row count must be a non-negative integer"))?;
    let columns: u32 = arguments
        .next()
        .ok_or_else(|| invalid("missing column count"))?
        .parse()
        .map_err(|_| invalid("column count must be a non-negative integer"))?;
    if arguments.next().is_some() {
        return Err(invalid("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    let (sheet_index, local_index) = locate_table(&package, table_index)?;
    let before = package.sheets()[sheet_index]
        .tables()
        .nth(local_index)
        .ok_or_else(|| invalid("table index was not found"))?;
    let original = exact_bytes(&package)?;
    let commit = package
        .edit_table_resize(
            SheetSelector::index(sheet_index),
            TableSelector::index(local_index),
        )?
        .resize(rows, columns)?
        .commit()?;
    let committed = commit.package();
    assert_size(
        committed,
        sheet_index,
        local_index,
        rows,
        columns,
        "committed",
    )?;

    let restored = committed.apply_table_resize(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != original {
        return Err(invalid(
            "inverse patch did not restore the exact input package",
        ));
    }

    committed.save(&output)?;
    let reopened = Package::open(&output)?;
    assert_size(
        &reopened,
        sheet_index,
        local_index,
        rows,
        columns,
        "reopened",
    )?;

    println!(
        "resized workbook table {} ({:?}) from {}x{} to {}x{}",
        table_index,
        before.name(),
        before.row_count(),
        before.column_count(),
        rows,
        columns,
    );
    Ok(())
}

fn locate_table(package: &Package, global_index: usize) -> Result<(usize, usize), Box<dyn Error>> {
    let mut offset = 0usize;
    for (sheet_index, sheet) in package.sheets().iter().enumerate() {
        let count = sheet.tables().count();
        if global_index < offset + count {
            return Ok((sheet_index, global_index - offset));
        }
        offset += count;
    }
    Err(invalid("table index was not found"))
}

fn assert_size(
    package: &Package,
    sheet_index: usize,
    table_index: usize,
    rows: u32,
    columns: u32,
    stage: &str,
) -> Result<(), Box<dyn Error>> {
    let table = package
        .sheets()
        .get(sheet_index)
        .and_then(|sheet| sheet.tables().nth(table_index))
        .ok_or_else(|| invalid(format!("{stage} package did not expose the resized table")))?;
    if table.row_count() != rows || table.column_count() != columns {
        return Err(invalid(format!(
            "{stage} package has {}x{} instead of {}x{}",
            table.row_count(),
            table.column_count(),
            rows,
            columns
        )));
    }
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(io::ErrorKind::InvalidInput, message.into()))
}
