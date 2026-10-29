//! Append an empty sheet to an existing Numbers workbook.
//!
//! The example uses the selector-first leaf transaction, checks the committed
//! and reopened semantic result, and verifies that the inverse patch restores
//! the exact source artifact.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::env;
use std::error::Error;
use std::io;

use litchi_numbers::Package;

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or_else(|| invalid("missing input path"))?;
    let output = arguments
        .next()
        .ok_or_else(|| invalid("missing output path"))?;
    let name = arguments
        .next()
        .ok_or_else(|| invalid("missing sheet name"))?;
    if arguments.next().is_some() {
        return Err(invalid("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    let before = sheet_names(&package);
    let original = exact_bytes(&package)?;
    let commit = package.edit_sheet_lifecycle().add(&name)?.commit()?;
    let committed = commit.package();
    let committed_names = sheet_names(committed);
    if committed_names.len() != before.len() + 1
        || committed_names.last().map(String::as_str) != Some(name.as_str())
    {
        return Err(invalid(
            "committed package did not preserve the added sheet",
        ));
    }

    let restored = committed.apply_sheet_lifecycle(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != original {
        return Err(invalid(
            "inverse patch did not restore the exact input package",
        ));
    }

    committed.save(&output)?;
    let reopened = Package::open(&output)?;
    let reopened_names = sheet_names(&reopened);
    if reopened_names != committed_names {
        return Err(invalid("reopened package did not preserve the added sheet"));
    }

    println!(
        "created sheet {:?}: {} sheet(s)",
        name,
        reopened_names.len()
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
