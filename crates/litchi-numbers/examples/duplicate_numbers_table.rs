//! Duplicate one populated Numbers table with independent storage.
//!
//! The table index is the historical workbook-wide zero-based index. The
//! focused package API receives the corresponding sheet and sheet-local
//! selector, then the output is reopened and the copied semantic content is
//! checked.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::env;
use std::error::Error;
use std::io;

use litchi_numbers::cell::Value;
use litchi_numbers::{At, Dimensions, Package, SheetSelector, Table, TableSelector};

#[derive(Debug, PartialEq)]
struct TableSnapshot {
    dimensions: Dimensions,
    cells: Vec<(At, Value)>,
    column_headers: Vec<String>,
    row_headers: Vec<String>,
}

impl TableSnapshot {
    fn from_table(table: &Table) -> Self {
        Self {
            dimensions: table.dimensions(),
            cells: table
                .iter_cells()
                .map(|cell| (cell.position(), cell.value().clone()))
                .collect(),
            column_headers: table.column_headers().map(str::to_owned).collect(),
            row_headers: table.row_headers().map(str::to_owned).collect(),
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    if arguments.len() != 3 {
        return Err(invalid(
            "usage: duplicate_numbers_table <input.numbers> <output.numbers> <table-index>",
        ));
    }
    if arguments[0] == arguments[1] {
        return Err(invalid("input and output paths must differ"));
    }
    let table_index = arguments[2]
        .parse::<usize>()
        .map_err(|_| invalid("table index must be a non-negative integer"))?;

    let package = Package::open(&arguments[0])?;
    let (sheet_index, local_index) = locate_table(&package, table_index)?;
    let source = package
        .sheets()
        .get(sheet_index)
        .and_then(|sheet| sheet.tables().nth(local_index))
        .ok_or_else(|| invalid("table index was not found"))?;
    let expected = TableSnapshot::from_table(source);
    let before_count = package.sheets()[sheet_index].tables().count();
    let original = exact_bytes(&package)?;

    let commit = package
        .edit_table_lifecycle(SheetSelector::index(sheet_index))?
        .duplicate(TableSelector::index(local_index))?
        .commit()?;
    let committed = commit.package();
    let copied = committed
        .sheets()
        .get(sheet_index)
        .and_then(|sheet| sheet.tables().last())
        .ok_or_else(|| invalid("committed package did not expose the duplicate"))?;
    if committed.sheets()[sheet_index].tables().count() != before_count + 1
        || TableSnapshot::from_table(copied) != expected
    {
        return Err(invalid(
            "committed duplicate changed table content or count",
        ));
    }

    let restored = committed.apply_table_lifecycle(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != original {
        return Err(invalid(
            "inverse patch did not restore the exact input package",
        ));
    }

    committed.save(&arguments[1])?;
    let reopened = Package::open(&arguments[1])?;
    let reopened_sheet = reopened
        .sheets()
        .get(sheet_index)
        .ok_or_else(|| invalid("reopened package lost the source sheet"))?;
    let reopened_copy = reopened_sheet
        .tables()
        .last()
        .ok_or_else(|| invalid("reopened package did not expose the duplicate"))?;
    if reopened_sheet.tables().count() != before_count + 1
        || TableSnapshot::from_table(reopened_copy) != expected
    {
        return Err(invalid("reopened duplicate changed table content or count"));
    }

    println!(
        "duplicated workbook table {} on sheet {}: {} rows x {} columns, {} materialized cells",
        table_index,
        sheet_index,
        expected.dimensions.rows(),
        expected.dimensions.columns(),
        expected.cells.len(),
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

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(io::ErrorKind::InvalidInput, message.into()))
}
