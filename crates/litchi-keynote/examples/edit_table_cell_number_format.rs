//! Set the fixed decimal precision of one existing Keynote table cell.
//!
//! Slide/table selection and the A1 cell coordinate are semantic; native
//! object identifiers, BNC records, and format-list keys stay private.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::PathBuf;

use litchi_keynote::{CellPosition, DecimalPlaces, Package, SlideSelector, TableSelector};

const USAGE: &str = "usage: edit_table_cell_number_format \
                     <input.key> <output.key> <slide-index> <table-index> \
                     <cell-A1> <decimal-places>";

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let slide = parse_usize(
        required_argument(&mut arguments, "missing slide index")?,
        "slide index",
    )?;
    let table = parse_usize(
        required_argument(&mut arguments, "missing table index")?,
        "table index",
    )?;
    let cell = required_text(&mut arguments, "missing A1 cell coordinate")?;
    let position = CellPosition::from_a1(&cell)?;
    let decimal_places = parse_u8(
        required_argument(&mut arguments, "missing decimal places")?,
        "decimal places",
    )?;
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    let slide = SlideSelector::index(slide);
    let table = TableSelector::index(table);
    let before = package.slide_table_cell_number_format(slide, table, position)?;
    let requested = before
        .unwrap_or_default()
        .with_decimal_places(DecimalPlaces::fixed(decimal_places)?);
    let source_bytes = exact_bytes(&package)?;

    let commit = package
        .edit_slide_table_cell_number_format(slide, table, position)?
        .set(requested)
        .commit()?;
    if commit
        .package()
        .slide_table_cell_number_format(slide, table, position)?
        != Some(requested)
    {
        return Err(invalid_input(
            "committed package did not expose the requested Number format",
        ));
    }

    let restored = commit
        .package()
        .apply_slide_table_cell_number_format(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != source_bytes
        || restored
            .package()
            .slide_table_cell_number_format(slide, table, position)?
            != before
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and Number format",
        ));
    }

    commit.package().save(&output)?;
    let reopened = Package::open(&output)?;
    if reopened.slide_table_cell_number_format(slide, table, position)? != Some(requested) {
        return Err(invalid_input(
            "reopened output did not preserve the requested Number format",
        ));
    }

    println!(
        "table cell Number format: cell={}, before={before:?}, after={requested:?}, changed={}, touched_components={}, full_reparse={}",
        position,
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

fn required_text(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<String, Box<dyn Error>> {
    required_argument(arguments, message)?
        .into_string()
        .map_err(|_| invalid_input("cell coordinate must be valid UTF-8"))
}

fn parse_usize(argument: OsString, label: &str) -> Result<usize, Box<dyn Error>> {
    argument
        .into_string()
        .map_err(|_| invalid_input(format!("{label} must be valid UTF-8")))?
        .parse()
        .map_err(|_| invalid_input(format!("{label} must be a non-negative integer")))
}

fn parse_u8(argument: OsString, label: &str) -> Result<u8, Box<dyn Error>> {
    argument
        .into_string()
        .map_err(|_| invalid_input(format!("{label} must be valid UTF-8")))?
        .parse()
        .map_err(|_| invalid_input(format!("{label} must be an integer from 0 through 30")))
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
