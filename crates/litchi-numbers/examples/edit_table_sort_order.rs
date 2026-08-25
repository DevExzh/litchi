//! Set one Numbers table's persisted sort-rule configuration.
//!
//! This command edits only the table's stored sort order.  It never invokes
//! Numbers' physical "Sort Now" operation, so table rows and their related
//! native graphs remain in their original order.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};
use litchi_numbers::{Package, SheetSelector, TableSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_table_sort_order <input.numbers> <output.numbers> \\
                     <index:N|name:NAME> <index:N|name:NAME> \\
                     [--scope all|selected] <COLUMN:DIRECTION>... \\
                     [--inverse PATH]";

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

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let sheet = parse_selector(required_text(&mut arguments, "missing sheet selector")?)?;
    let table = parse_selector(required_text(&mut arguments, "missing table selector")?)?;
    let (scope, rules, inverse_output) = parse_operations(&mut arguments)?;

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

    let source_bytes = std::fs::read(&input)?;
    let package = Package::open(&input)?;
    if exact_bytes(&package)? != source_bytes {
        return Err(invalid_input(
            "input package did not preserve its exact source bytes",
        ));
    }
    let order =
        Order::with_scope(scope, rules).map_err(|error| -> Box<dyn Error> { Box::new(error) })?;
    let commit = package
        .edit_table_sort_order(sheet.sheet(), table.table())?
        .set(order.clone())
        .commit()?;
    if commit
        .package()
        .table_sort_order(sheet.sheet(), table.table())?
        .as_ref()
        != Some(&order)
    {
        return Err(invalid_input(
            "persisted sort order failed semantic readback",
        ));
    }

    let restored = commit
        .package()
        .apply_table_sort_order(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != source_bytes {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    save_new(&output, commit.package())?;
    let reopened = Package::open(&output)?;
    if reopened
        .table_sort_order(sheet.sheet(), table.table())?
        .as_ref()
        != Some(&order)
    {
        return Err(invalid_input(
            "written output did not preserve the persisted sort order",
        ));
    }
    if let Some(path) = inverse_output {
        save_new(&path, restored.package())?;
    }

    println!(
        "table sort order: scope={scope:?}, rules={}, changed={}, touched_components={}, deleted_previews={}",
        order.rules().len(),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
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
        .map_err(|_| invalid_input("selectors and sort arguments must be valid UTF-8"))
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

fn parse_operations(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<(Scope, Vec<Rule>, Option<PathBuf>), Box<dyn Error>> {
    let mut scope = Scope::EntireTable;
    let mut scope_seen = false;
    let mut rules = Vec::new();
    let mut inverse_output = None;

    while let Some(argument) = arguments.next() {
        let value = argument
            .into_string()
            .map_err(|_| invalid_input("sort arguments must be valid UTF-8"))?;
        match value.as_str() {
            "--scope" => {
                if scope_seen {
                    return Err(invalid_input("--scope may appear only once"));
                }
                scope_seen = true;
                scope = parse_scope(required_text(arguments, "missing --scope value")?)?;
            },
            "--inverse" => {
                let path = PathBuf::from(required_argument(arguments, "missing --inverse path")?);
                if arguments.next().is_some() {
                    return Err(invalid_input(
                        "--inverse PATH must be the final command-line option",
                    ));
                }
                inverse_output = Some(path);
                break;
            },
            _ => rules.push(parse_rule(&value)?),
        }
    }

    if rules.is_empty() {
        return Err(invalid_input(
            "at least one COLUMN:DIRECTION rule is required",
        ));
    }
    Ok((scope, rules, inverse_output))
}

fn parse_scope(value: String) -> Result<Scope, Box<dyn Error>> {
    match value.as_str() {
        "all" => Ok(Scope::EntireTable),
        "selected" => Ok(Scope::SelectedRows),
        _ => Err(invalid_input("scope must be all or selected")),
    }
}

fn parse_rule(value: &str) -> Result<Rule, Box<dyn Error>> {
    let (column, direction) = value
        .split_once(':')
        .ok_or_else(|| invalid_input("rule must have the form COLUMN:DIRECTION"))?;
    let column = column
        .parse::<usize>()
        .map_err(|_| invalid_input("sort column must be a non-negative integer"))?;
    let column = ColumnIndex::new(column).map_err(|error| -> Box<dyn Error> { Box::new(error) })?;
    let direction = match direction {
        "ascending" => Direction::Ascending,
        "descending" => Direction::Descending,
        _ => {
            return Err(invalid_input(
                "sort direction must be ascending or descending",
            ));
        },
    };
    Ok(Rule::new(column, direction))
}

/// Publishes through a sibling temporary file without overwriting an existing target.
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

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
