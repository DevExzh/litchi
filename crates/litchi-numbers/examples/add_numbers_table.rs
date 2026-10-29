//! Add an independent empty table to an existing Numbers sheet.
//!
//! The command keeps the historical CLI shape while using the focused
//! selector-first table-lifecycle transaction. The committed package is
//! reopened from the requested output and the inverse patch is checked
//! against the exact input bytes.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::env;
use std::error::Error;
use std::io;

use litchi_numbers::{Package, SheetSelector};

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or_else(|| invalid("missing input path"))?;
    let output = arguments
        .next()
        .ok_or_else(|| invalid("missing output path"))?;
    let sheet_index: usize = arguments
        .next()
        .ok_or_else(|| invalid("missing sheet index"))?
        .parse()
        .map_err(|_| invalid("sheet index must be a non-negative integer"))?;
    let name = arguments
        .next()
        .ok_or_else(|| invalid("missing table name"))?;
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
    let before_count = package
        .sheets()
        .get(sheet_index)
        .ok_or_else(|| invalid("sheet index was not found"))?
        .tables()
        .count();
    let original = exact_bytes(&package)?;
    let commit = package
        .edit_table_lifecycle(SheetSelector::index(sheet_index))?
        .add(&name, rows, columns)?
        .commit()?;
    let committed = commit.package();
    let sheet = committed
        .sheets()
        .get(sheet_index)
        .ok_or_else(|| invalid("committed package lost the selected sheet"))?;
    let added = sheet
        .tables()
        .last()
        .ok_or_else(|| invalid("committed package did not expose the added table"))?;
    if sheet.tables().count() != before_count + 1
        || added.name() != name
        || added.row_count() != rows
        || added.column_count() != columns
    {
        return Err(invalid(
            "committed package did not preserve the requested table",
        ));
    }

    let restored = committed.apply_table_lifecycle(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != original {
        return Err(invalid(
            "inverse patch did not restore the exact input package",
        ));
    }

    committed.save(&output)?;
    let reopened = Package::open(&output)?;
    let reopened_sheet = reopened
        .sheets()
        .get(sheet_index)
        .ok_or_else(|| invalid("reopened package lost the selected sheet"))?;
    let reopened_added = reopened_sheet
        .tables()
        .last()
        .ok_or_else(|| invalid("reopened package did not expose the added table"))?;
    if reopened_sheet.tables().count() != before_count + 1
        || reopened_added.name() != name
        || reopened_added.row_count() != rows
        || reopened_added.column_count() != columns
    {
        return Err(invalid("reopened package did not preserve the added table"));
    }

    println!(
        "created table {:?}: {} rows x {} columns",
        name, rows, columns
    );
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!(
            "{}\n\nusage: add_numbers_table <input.numbers> <output.numbers> <sheet-index> <name> <rows> <columns>",
            message.into()
        ),
    ))
}
