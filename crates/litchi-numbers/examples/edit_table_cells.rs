//! Atomically apply selector-first scalar cell changes to one Numbers table.
//!
//! Text inputs are accepted by the public API, although a changed text write
//! can still be refused for native storage profiles that have not yet been
//! admitted. Every successful artifact is published through a sibling
//! temporary file without clobbering an existing target.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports content-free commit diagnostics"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::{
    Package, SheetSelector, TableSelector,
    cell::Value,
    table::{
        CellPosition,
        cells::{Change, Input, Storage},
    },
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_table_cells <input.numbers> <output.numbers> \\
                     <index:N|name:NAME> <index:N|name:NAME> \\
                     <set A1 text TEXT|set A1 number NUMBER|set A1 boolean true|false|\
                      set A1 date SECONDS|set A1 duration SECONDS|clear A1>... \\
                     [--inverse PATH]";

enum Selector {
    Index(usize),
    Name(String),
}

impl Selector {
    fn sheet_selector(&self) -> SheetSelector<'_> {
        match self {
            Self::Index(index) => SheetSelector::index(*index),
            Self::Name(name) => SheetSelector::name(name),
        }
    }

    fn table_selector(&self) -> TableSelector<'_> {
        match self {
            Self::Index(index) => TableSelector::index(*index),
            Self::Name(name) => TableSelector::name(name),
        }
    }
}

enum Expected {
    Set(Input),
    Clear,
}

struct RequestedChange {
    change: Change,
    position: CellPosition,
    expected: Expected,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let sheet = parse_selector(required_text(&mut arguments, "missing sheet selector")?)?;
    let table = parse_selector(required_text(&mut arguments, "missing table selector")?)?;
    let (requested, inverse_output) = parse_changes(&mut arguments)?;

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
    let source = exact_bytes(&package)?;
    let changes = requested
        .iter()
        .map(|requested| requested.change.clone())
        .collect::<Vec<_>>();
    let commit = package
        .edit_table_cells(sheet.sheet_selector(), table.table_selector())?
        .extend(changes)?
        .commit()?;

    verify_readback(commit.package(), &sheet, &table, &requested)?;
    verify_commit_artifact(&source, &commit)?;

    let replay = package.apply_table_cells(commit.patch())?;
    if exact_bytes(replay.package())? != exact_bytes(commit.package())? {
        return Err(invalid_input(
            "patch replay did not reproduce the committed exact package",
        ));
    }
    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != source {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    save_new(&output, commit.package())?;
    if let Some(path) = inverse_output {
        save_new(&path, restored.package())?;
    }

    let diagnostics = commit.diagnostics();
    println!(
        "table cells: changed={}, requested_cells={}, changed_cells={}, touched_components={}, refreshed_formula_caches={}, deleted_previews={}",
        diagnostics.changed(),
        diagnostics.requested_cells(),
        diagnostics.changed_cells(),
        diagnostics.touched_components(),
        diagnostics.refreshed_formula_caches(),
        diagnostics.deleted_previews(),
    );
    Ok(())
}

fn parse_changes(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<(Vec<RequestedChange>, Option<PathBuf>), Box<dyn Error>> {
    let mut requested = Vec::new();
    let mut inverse = None;
    while let Some(operation) = arguments.next() {
        if operation == OsStr::new("--inverse") {
            if inverse.is_some() {
                return Err(invalid_input("--inverse may appear only once"));
            }
            inverse = Some(PathBuf::from(required_argument(
                arguments,
                "missing --inverse path",
            )?));
            if arguments.next().is_some() {
                return Err(invalid_input(
                    "unexpected trailing arguments after --inverse PATH",
                ));
            }
            break;
        }

        let operation = operation
            .into_string()
            .map_err(|_| invalid_input("cell operations must be valid UTF-8"))?;
        match operation.as_str() {
            "set" => {
                let address = required_text(arguments, "missing A1 address after set")?;
                let position = CellPosition::from_a1(&address)
                    .map_err(|error| invalid_input(format!("invalid A1 address: {error}")))?;
                let input = parse_input(arguments)?;
                requested.push(RequestedChange {
                    change: Change::set(position, input.clone()),
                    position,
                    expected: Expected::Set(input),
                });
            },
            "clear" => {
                let address = required_text(arguments, "missing A1 address after clear")?;
                let position = CellPosition::from_a1(&address)
                    .map_err(|error| invalid_input(format!("invalid A1 address: {error}")))?;
                requested.push(RequestedChange {
                    change: Change::clear(position),
                    position,
                    expected: Expected::Clear,
                });
            },
            _ => {
                return Err(invalid_input(
                    "operation must be set, clear, or --inverse PATH",
                ));
            },
        }
    }
    if requested.is_empty() {
        return Err(invalid_input(
            "at least one set or clear operation is required",
        ));
    }
    Ok((requested, inverse))
}

fn parse_input(arguments: &mut impl Iterator<Item = OsString>) -> Result<Input, Box<dyn Error>> {
    let kind = required_text(arguments, "missing scalar kind after set A1")?;
    let value = required_text(arguments, "missing scalar value after set A1 KIND")?;
    match kind.as_str() {
        "text" => Ok(Input::text(value)?),
        "number" => Ok(Input::number(parse_finite(&value, "number")?)?),
        "boolean" => match value.as_str() {
            "true" => Ok(Input::boolean(true)),
            "false" => Ok(Input::boolean(false)),
            _ => Err(invalid_input("boolean must be true or false")),
        },
        "date" => Ok(Input::date(parse_finite(&value, "date seconds")?)?),
        "duration" => Ok(Input::duration(parse_finite(&value, "duration seconds")?)?),
        _ => Err(invalid_input(
            "scalar kind must be text, number, boolean, date, or duration",
        )),
    }
}

fn parse_finite(value: &str, label: &str) -> Result<f64, Box<dyn Error>> {
    let value = value
        .parse::<f64>()
        .map_err(|_| invalid_input(format!("{label} must be a finite decimal number")))?;
    if !value.is_finite() {
        return Err(invalid_input(format!(
            "{label} must be a finite decimal number"
        )));
    }
    Ok(value)
}

fn verify_readback(
    package: &Package,
    sheet: &Selector,
    table: &Selector,
    requested: &[RequestedChange],
) -> Result<(), Box<dyn Error>> {
    for requested in requested {
        let state = package.table_cell(
            sheet.sheet_selector(),
            table.table_selector(),
            requested.position,
        )?;
        let matches = match &requested.expected {
            Expected::Set(input) => input_matches(input, state.storage()),
            Expected::Clear => matches!(
                state.storage(),
                Storage::Missing | Storage::Stored(Value::Empty)
            ),
        };
        if !matches {
            return Err(invalid_input(
                "committed cell values did not match the requested semantic batch",
            ));
        }
    }
    Ok(())
}

fn input_matches(input: &Input, storage: &Storage) -> bool {
    match (input, storage) {
        (Input::Text(expected), Storage::Stored(Value::Text(actual))) => expected == actual,
        (Input::Number(expected), Storage::Stored(Value::Number(actual))) => expected == actual,
        (Input::Boolean(expected), Storage::Stored(Value::Boolean(actual))) => expected == actual,
        (Input::Date(expected), Storage::Stored(Value::Date(actual))) => expected == actual,
        (Input::Duration(expected), Storage::Stored(Value::Duration(actual))) => expected == actual,
        _ => false,
    }
}

fn verify_commit_artifact(
    source: &[u8],
    commit: &litchi_numbers::table::cells::Commit,
) -> Result<(), Box<dyn Error>> {
    let target = exact_bytes(commit.package())?;
    if commit.diagnostics().changed() {
        if target == source || commit.patch().is_noop() {
            return Err(invalid_input(
                "changed diagnostics did not produce a changed exact package",
            ));
        }
    } else if target != source || !commit.patch().is_noop() {
        return Err(invalid_input(
            "no-op diagnostics did not preserve the exact input package",
        ));
    }
    Ok(())
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
        .map_err(|_| invalid_input("selectors, operations, and values must be valid UTF-8"))
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
