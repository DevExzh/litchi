//! Move one Numbers table between two sheets through the focused package API.
//!
//! The input must contain both the source and destination sheets. The example
//! intentionally accepts the package path from the caller so it can be used
//! with an Apple-authored two-sheet fixture without manufacturing a package or
//! depending on the legacy `litchi-iwa` facade.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::cell::Value;
use litchi_numbers::{CellPosition, Dimensions, Package, SheetSelector, Table, TableSelector};

const USAGE: &str = "usage: move_table <input.numbers> <output.numbers> \\
                     <source-sheet-selector> <table-selector> \\
                     <destination-sheet-selector>\n\
selectors: index:N or name:NAME";

#[derive(Debug)]
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

#[derive(Debug, PartialEq)]
struct TableSnapshot {
    name: String,
    dimensions: Dimensions,
    cells: Vec<(CellPosition, Value)>,
    column_headers: Vec<String>,
    row_headers: Vec<String>,
}

impl TableSnapshot {
    fn from_table(table: &Table) -> Self {
        Self {
            name: table.name().to_owned(),
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
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let source_sheet = parse_selector(required_text(
        &mut arguments,
        "missing source-sheet selector",
    )?)?;
    let table = parse_selector(required_text(&mut arguments, "missing table selector")?)?;
    let destination_sheet = parse_selector(required_text(
        &mut arguments,
        "missing destination-sheet selector",
    )?)?;
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }

    ensure_numbers_path(&input, "input")?;
    ensure_numbers_path(&output, "output")?;
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    if package.sheet(destination_sheet.sheet())?.is_none() {
        return Err(invalid_input("destination sheet selector did not resolve"));
    }
    let source = package
        .table(source_sheet.sheet(), table.table())?
        .ok_or_else(|| invalid_input("source sheet/table selectors did not resolve"))?;
    let expected = TableSnapshot::from_table(source);

    let edit = package.edit_table_relocation(
        source_sheet.sheet(),
        table.table(),
        destination_sheet.sheet(),
    )?;
    let source_sheet_position = edit.source_sheet_position();
    let destination_sheet_position = edit.destination_sheet_position();
    let destination_table_position = edit.destination_table_position();
    let source_table_count = package.sheets()[source_sheet_position].tables().count();
    let destination_table_count = package.sheets()[destination_sheet_position]
        .tables()
        .count();
    let same_sheet = source_sheet_position == destination_sheet_position;
    let commit = edit.commit()?;
    let moved = commit.package();
    assert_relocation(
        moved,
        source_sheet_position,
        destination_sheet_position,
        destination_table_position,
        source_table_count,
        destination_table_count,
        same_sheet,
        &expected,
        "committed package",
    )?;

    // `Package::save` publishes the caller-selected path with its own
    // recoverable sibling staging file; this example creates no persistent
    // temporary artifact.
    moved.save(&output)?;
    let reopened = Package::open(&output)?;
    assert_relocation(
        &reopened,
        source_sheet_position,
        destination_sheet_position,
        destination_table_position,
        source_table_count,
        destination_table_count,
        same_sheet,
        &expected,
        "reopened package",
    )?;

    println!(
        "table moved: source={:?}, destination={:?}, name={:?}, rows={}, columns={}, materialized_cells={}",
        source_sheet,
        destination_sheet,
        expected.name,
        expected.dimensions.rows(),
        expected.dimensions.columns(),
        expected.cells.len(),
    );
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "the example verifies both sheet cardinalities and the exact destination position"
)]
fn assert_relocation(
    package: &Package,
    source_sheet: usize,
    destination_sheet: usize,
    destination_table: usize,
    source_table_count: usize,
    destination_table_count: usize,
    same_sheet: bool,
    expected: &TableSnapshot,
    stage: &str,
) -> Result<(), Box<dyn Error>> {
    let expected_source_count = if same_sheet {
        source_table_count
    } else {
        source_table_count
            .checked_sub(1)
            .ok_or_else(|| invalid_input("source table count underflowed"))?
    };
    let expected_destination_count = if same_sheet {
        destination_table_count
    } else {
        destination_table_count
            .checked_add(1)
            .ok_or_else(|| invalid_input("destination table count overflowed"))?
    };
    let actual_source_count = package
        .sheets()
        .get(source_sheet)
        .ok_or_else(|| invalid_input(format!("{stage} lost the source sheet")))?
        .tables()
        .count();
    let destination = package
        .sheets()
        .get(destination_sheet)
        .ok_or_else(|| invalid_input(format!("{stage} lost the destination sheet")))?;
    if actual_source_count != expected_source_count
        || destination.tables().count() != expected_destination_count
    {
        return Err(invalid_input(format!(
            "{stage} changed an unexpected table count"
        )));
    }
    let moved = destination
        .tables()
        .nth(destination_table)
        .ok_or_else(|| invalid_input(format!("{stage} did not expose the moved table")))?;
    assert_preserved(expected, moved, stage)
}

fn assert_preserved(
    expected: &TableSnapshot,
    actual: &Table,
    stage: &str,
) -> Result<(), Box<dyn Error>> {
    let actual = TableSnapshot::from_table(actual);
    if &actual != expected {
        return Err(invalid_input(format!(
            "{stage} changed the moved table's semantic content"
        )));
    }
    Ok(())
}

fn parse_selector(value: String) -> Result<Selector, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(Selector::Index)
            .map_err(|_| invalid_input("selector index must be a non-negative integer"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input("selector name must not be empty"));
        }
        return Ok(Selector::Name(name.to_owned()));
    }
    Err(invalid_input("selector must start with index: or name:"))
}

fn ensure_numbers_path(path: &Path, role: &str) -> Result<(), Box<dyn Error>> {
    if path
        .extension()
        .is_some_and(|extension| extension == "numbers")
    {
        Ok(())
    } else {
        Err(invalid_input(format!(
            "{role} path must have a .numbers extension"
        )))
    }
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
        .map_err(|_| invalid_input("selector arguments must be valid UTF-8"))
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
