//! Read, set, or reset one existing Numbers cell's Fraction format.
//!
//! The cell is selected by a sheet selector, a sheet-scoped table selector,
//! and a checked A1 position.  `clear` (and its `reset` alias) stages `None`,
//! meaning that the cell has no explicit Fraction format.  `set` stages an
//! explicit [`Fraction`] format with the requested denominator accuracy.  The
//! command verifies the semantic readback and exact inverse before publishing
//! through a sibling temporary file, so an existing destination is never
//! replaced.  A Fraction-format transaction changes display metadata only;
//! the scalar cell value is left untouched.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::cell::data_format::fraction::{Fraction, FractionAccuracy};
use litchi_numbers::{CellPosition, Package, SheetSelector, TableSelector};
use tempfile::NamedTempFile;

const USAGE: &str = concat!(
    "usage: edit_table_cell_fraction_format <input.numbers> <output.numbers> ",
    "<index:N|name:NAME> <index:N|name:NAME> <A1> ",
    "<clear|reset|set ACCURACY> [--inverse PATH]\n",
    "ACCURACY is up-to-one-digit, up-to-two-digits, up-to-three-digits, ",
    "halves, quarters, eighths, sixteenths, tenths, or hundredths; ",
    "clear and reset produce None (automatic/no explicit format)"
);

enum Selector {
    Index(usize),
    Name(String),
}

impl Selector {
    fn sheet(&self) -> SheetSelector<'_> {
        match self {
            Self::Index(index) => SheetSelector::index(*index),
            Self::Name(name) => SheetSelector::name(name),
        }
    }

    fn table(&self) -> TableSelector<'_> {
        match self {
            Self::Index(index) => TableSelector::index(*index),
            Self::Name(name) => TableSelector::name(name),
        }
    }
}

#[derive(Clone, Copy)]
enum Operation {
    Set(Fraction),
    Clear,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let sheet = parse_selector(required_text(&mut arguments, "missing sheet selector")?)?;
    let table = parse_selector(required_text(&mut arguments, "missing table selector")?)?;
    let position = CellPosition::from_a1(&required_text(&mut arguments, "missing cell address")?)
        .map_err(|error| invalid_input(format!("invalid A1 address: {error}")))?;
    let operation = parse_operation(&mut arguments)?;
    let inverse_output = parse_inverse_output(&mut arguments)?;

    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }
    if inverse_output
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "inverse path must differ from input and output paths",
        ));
    }

    let package = Package::open(&input)?;
    let sheet_selector = sheet.sheet();
    let table_selector = table.table();
    let before = package.table_cell_fraction_format(sheet_selector, table_selector, position)?;
    let source_bytes = exact_bytes(&package)?;

    // Read back the existing semantic value before staging any mutation.  A
    // clear/set of that same value must be a byte-exact no-op, which also
    // makes this command useful for checking an admitted source graph.
    let noop = match before {
        Some(format) => package
            .edit_table_cell_fraction_format(sheet.sheet(), table.table(), position)?
            .set(format)
            .commit()?,
        None => package
            .edit_table_cell_fraction_format(sheet.sheet(), table.table(), position)?
            .clear()
            .commit()?,
    };
    if !noop.patch().is_noop()
        || exact_bytes(noop.package())? != source_bytes
        || noop
            .package()
            .table_cell_fraction_format(sheet.sheet(), table.table(), position)?
            != before
    {
        return Err(invalid_input(
            "restaging the existing Fraction format was not an exact no-op",
        ));
    }

    let edit = package.edit_table_cell_fraction_format(sheet.sheet(), table.table(), position)?;
    let commit = match operation {
        Operation::Set(format) => edit.set(format).commit()?,
        Operation::Clear => edit.clear().commit()?,
    };
    let after =
        commit
            .package()
            .table_cell_fraction_format(sheet.sheet(), table.table(), position)?;
    let expected = match operation {
        Operation::Set(format) => Some(format),
        Operation::Clear => None,
    };
    if after != expected {
        return Err(invalid_input(
            "committed package did not expose the requested Fraction format",
        ));
    }

    let restored = commit
        .package()
        .apply_table_cell_fraction_format(&commit.patch().inverse())?;
    if restored
        .package()
        .table_cell_fraction_format(sheet.sheet(), table.table(), position)?
        != before
        || exact_bytes(restored.package())? != source_bytes
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and Fraction format",
        ));
    }

    save_new(&output, commit.package())?;
    if let Some(path) = inverse_output {
        save_new(&path, restored.package())?;
    }

    println!(
        "table cell Fraction format: cell={}, before={}, after={}, changed={}, touched_components={}, full_reparse={}, scalar_value=untouched",
        position,
        describe_format(before),
        describe_format(after),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match required_text(arguments, "missing Fraction format operation")?.as_str() {
        "clear" | "reset" => Ok(Operation::Clear),
        "set" => {
            let accuracy = parse_accuracy(required_text(
                arguments,
                "missing fraction accuracy after set",
            )?)?;
            Ok(Operation::Set(Fraction::new(accuracy)))
        },
        _ => Err(invalid_input(
            "Fraction format operation must be clear, reset, or set",
        )),
    }
}

fn parse_accuracy(value: String) -> Result<FractionAccuracy, Box<dyn Error>> {
    match value.as_str() {
        "up-to-one-digit" => Ok(FractionAccuracy::UpToOneDigit),
        "up-to-two-digits" => Ok(FractionAccuracy::UpToTwoDigits),
        "up-to-three-digits" => Ok(FractionAccuracy::UpToThreeDigits),
        "halves" => Ok(FractionAccuracy::Halves),
        "quarters" => Ok(FractionAccuracy::Quarters),
        "eighths" => Ok(FractionAccuracy::Eighths),
        "sixteenths" => Ok(FractionAccuracy::Sixteenths),
        "tenths" => Ok(FractionAccuracy::Tenths),
        "hundredths" => Ok(FractionAccuracy::Hundredths),
        _ => Err(invalid_input(
            "fraction accuracy must be up-to-one-digit, up-to-two-digits, up-to-three-digits, halves, quarters, eighths, sixteenths, tenths, or hundredths",
        )),
    }
}

fn parse_selector(value: String) -> Result<Selector, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        if index.is_empty() || !index.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(invalid_input(
                "selector index must be a non-negative decimal integer",
            ));
        }
        return index
            .parse()
            .map(Selector::Index)
            .map_err(|_| invalid_input("selector index is too large"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input("selector name must not be empty"));
        }
        return Ok(Selector::Name(name.to_owned()));
    }
    Err(invalid_input("selector must start with index: or name:"))
}

fn parse_inverse_output(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Option<PathBuf>, Box<dyn Error>> {
    let Some(flag) = arguments.next() else {
        return Ok(None);
    };
    if flag != OsStr::new("--inverse") {
        return Err(invalid_input(
            "unexpected trailing argument; expected --inverse PATH",
        ));
    }
    let path = PathBuf::from(required_argument(arguments, "missing --inverse path")?);
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    Ok(Some(path))
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: impl Into<String>,
) -> Result<OsString, Box<dyn Error>> {
    arguments
        .next()
        .ok_or_else(|| invalid_input(message.into()))
}

fn required_text(
    arguments: &mut impl Iterator<Item = OsString>,
    message: impl Into<String>,
) -> Result<String, Box<dyn Error>> {
    required_argument(arguments, message)?
        .into_string()
        .map_err(|_| invalid_input("selectors, addresses, and settings must be valid UTF-8"))
}

/// Publishes through a synchronized sibling temporary file without replacing
/// an existing destination.
fn save_new(path: &Path, package: &Package) -> Result<(), Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(temporary.as_file_mut())?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| Box::new(error.error))?;
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn describe_format(format: Option<Fraction>) -> String {
    match format {
        None => "none (automatic/no explicit format)".to_owned(),
        Some(format) => format!("explicit(accuracy={})", accuracy_name(format.accuracy())),
    }
}

fn accuracy_name(accuracy: FractionAccuracy) -> &'static str {
    match accuracy {
        FractionAccuracy::UpToOneDigit => "up-to-one-digit",
        FractionAccuracy::UpToTwoDigits => "up-to-two-digits",
        FractionAccuracy::UpToThreeDigits => "up-to-three-digits",
        FractionAccuracy::Halves => "halves",
        FractionAccuracy::Quarters => "quarters",
        FractionAccuracy::Eighths => "eighths",
        FractionAccuracy::Sixteenths => "sixteenths",
        FractionAccuracy::Tenths => "tenths",
        FractionAccuracy::Hundredths => "hundredths",
    }
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
