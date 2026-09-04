//! Read, set, or clear one existing Numbers cell's Text format.
//!
//! The cell is selected by semantic sheet/table selectors and a checked A1
//! position.  `set` stages the typed [`Text`] marker only; it does not change
//! the stored cell value or convert a numeric cell to text.  The output is
//! written through a sibling temporary file and is never replaced in place.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::cell::data_format::Text;
use litchi_numbers::{CellPosition, Package, SheetSelector, TableSelector};
use tempfile::NamedTempFile;

const USAGE: &str = concat!(
    "usage: edit_table_cell_text_format <input.numbers> <output.numbers> ",
    "<index:N|name:NAME> <index:N|name:NAME> <A1> <clear|set>"
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
    Set,
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
    let operation = parse_operation(required_text(&mut arguments, "missing Text operation")?)?;

    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }

    // Opening through the focused package owner retains the exact source
    // archive while keeping native IDs and wire records below this example's
    // semantic boundary.
    let package = Package::open(&input)?;
    let before = package.table_cell_text_format(sheet.sheet(), table.table(), position)?;

    let edit = package.edit_table_cell_text_format(sheet.sheet(), table.table(), position)?;
    let commit = match operation {
        Operation::Set => edit.set(Text).commit()?,
        Operation::Clear => edit.clear().commit()?,
    };
    let after = commit
        .package()
        .table_cell_text_format(sheet.sheet(), table.table(), position)?;
    let expected = match operation {
        Operation::Set => Some(Text),
        Operation::Clear => None,
    };
    if after != expected {
        return Err(invalid_input(
            "committed package did not expose the requested Text format",
        ));
    }

    save_new(&output, commit.package())?;
    let reopened = Package::open(&output)?;
    if reopened.table_cell_text_format(sheet.sheet(), table.table(), position)? != after {
        return Err(invalid_input(
            "saved package did not preserve the requested Text format",
        ));
    }

    println!(
        "table cell Text format: cell={position}, before={}, after={}, changed={}, touched_components={}, full_reparse={}",
        describe_format(before),
        describe_format(after),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_operation(value: String) -> Result<Operation, Box<dyn Error>> {
    match value.as_str() {
        "set" => Ok(Operation::Set),
        "clear" => Ok(Operation::Clear),
        _ => Err(invalid_input("Text operation must be clear or set")),
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
        .map_err(|_| invalid_input("selectors, addresses, and operations must be valid UTF-8"))
}

/// Publish through a synchronized sibling temporary file without replacing
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

fn describe_format(format: Option<Text>) -> &'static str {
    match format {
        None => "automatic/inherited",
        Some(Text) => "explicit Text",
    }
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
