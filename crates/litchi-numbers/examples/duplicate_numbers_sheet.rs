//! Duplicate one Numbers sheet by its zero-based semantic position.
//!
//! The focused transaction copies the sheet's supported table graph and
//! generates the same suffixed visible name as the leaf API. The example
//! verifies the committed and reopened sheet lists and exact inverse bytes.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::env;
use std::error::Error;
use std::io;

use litchi_numbers::{Package, SheetSelector};

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    if arguments.len() != 3 {
        return Err(invalid(
            "usage: duplicate_numbers_sheet <input.numbers> <output.numbers> <sheet-index>",
        ));
    }
    let sheet_index = arguments[2]
        .parse::<usize>()
        .map_err(|_| invalid("sheet index must be a non-negative integer"))?;
    if arguments[0] == arguments[1] {
        return Err(invalid("input and output paths must differ"));
    }

    let package = Package::open(&arguments[0])?;
    let source = package
        .sheets()
        .get(sheet_index)
        .ok_or_else(|| invalid("sheet index was not found"))?;
    let source_name = source.name().to_owned();
    let source_tables = source.table_count();
    let before = sheet_names(&package);
    let original = exact_bytes(&package)?;
    let commit = package
        .edit_sheet_lifecycle()
        .duplicate(SheetSelector::index(sheet_index))?
        .commit()?;
    let committed = commit.package();
    let committed_names = sheet_names(committed);
    let copied = committed
        .sheets()
        .get(sheet_index + 1)
        .ok_or_else(|| invalid("committed package did not expose the duplicate"))?;
    if committed_names.len() != before.len() + 1
        || copied.table_count() != source_tables
        || copied.name() == source_name
        || !copied.name().starts_with(&format!("{source_name}-"))
    {
        return Err(invalid(
            "committed package did not preserve the duplicated sheet",
        ));
    }

    let restored = committed.apply_sheet_lifecycle(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != original {
        return Err(invalid(
            "inverse patch did not restore the exact input package",
        ));
    }

    committed.save(&arguments[1])?;
    let reopened = Package::open(&arguments[1])?;
    let reopened_names = sheet_names(&reopened);
    if reopened_names != committed_names
        || reopened
            .sheets()
            .get(sheet_index + 1)
            .is_none_or(|sheet| sheet.table_count() != source_tables)
    {
        return Err(invalid("reopened package did not preserve the duplicate"));
    }

    println!(
        "duplicated sheet {:?} as {:?} with {} table(s)",
        source_name,
        copied.name(),
        source_tables
    );
    Ok(())
}

fn sheet_names(package: &Package) -> Vec<String> {
    package
        .sheets()
        .iter()
        .map(|sheet| sheet.name().to_owned())
        .collect()
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(io::ErrorKind::InvalidInput, message.into()))
}
