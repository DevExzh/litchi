//! Remove one Numbers sheet by its zero-based semantic position.
//!
//! The leaf route refuses the final sheet and sheets with unsupported
//! non-table drawables. This example keeps that typed refusal, then verifies
//! persisted names and exact inverse bytes for a successful removal.

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
    let sheet_index = arguments
        .next()
        .ok_or_else(|| invalid("missing sheet index"))?
        .parse::<usize>()
        .map_err(|_| invalid("sheet index must be a non-negative integer"))?;
    if arguments.next().is_some() {
        return Err(invalid("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    let before = sheet_names(&package);
    if before.get(sheet_index).is_none() {
        return Err(invalid("sheet index was not found"));
    }
    if before.len() <= 1 {
        return Err(invalid("cannot remove the workbook's final sheet"));
    }
    let original = exact_bytes(&package)?;
    let commit = package
        .edit_sheet_lifecycle()
        .remove(SheetSelector::index(sheet_index))?
        .commit()?;
    let committed = commit.package();
    let committed_names = sheet_names(committed);
    let mut expected = before.clone();
    expected.remove(sheet_index);
    if committed_names != expected {
        return Err(invalid("committed package did not preserve sheet order"));
    }

    let restored = committed.apply_sheet_lifecycle(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != original {
        return Err(invalid(
            "inverse patch did not restore the exact input package",
        ));
    }

    committed.save(&output)?;
    let reopened = Package::open(&output)?;
    if sheet_names(&reopened) != expected {
        return Err(invalid("reopened package did not preserve sheet removal"));
    }

    println!(
        "removed sheet {}: {} sheet(s) remain",
        sheet_index,
        expected.len()
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
