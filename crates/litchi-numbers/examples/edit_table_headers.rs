//! Replace one Numbers table's presence-preserving header and footer settings.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::{
    Package, SheetSelector, TableSelector,
    table::headers::{Count, Settings},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_table_headers <input.numbers> <output.numbers> \\
                     <index:N|name:NAME> <index:N|name:NAME> \\
                     <unset|1..5:header-rows> <unset|1..5:header-columns> \\
                     <unset|1..5:footer-rows> \\
                     <unset|true|false:freeze-rows> \\
                     <unset|true|false:freeze-columns> \\
                     <unset|true|false:repeat-rows> \\
                     <unset|true|false:repeat-columns> [--inverse PATH]";

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

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let sheet = parse_selector(required_text(&mut arguments, "missing sheet selector")?)?;
    let table = parse_selector(required_text(&mut arguments, "missing table selector")?)?;
    let settings = Settings {
        header_rows: parse_count(&mut arguments, "header rows")?,
        header_columns: parse_count(&mut arguments, "header columns")?,
        footer_rows: parse_count(&mut arguments, "footer rows")?,
        header_rows_frozen: parse_optional_bool(&mut arguments, "freeze rows")?,
        header_columns_frozen: parse_optional_bool(&mut arguments, "freeze columns")?,
        repeating_header_rows_enabled: parse_optional_bool(&mut arguments, "repeat rows")?,
        repeating_header_columns_enabled: parse_optional_bool(&mut arguments, "repeat columns")?,
    };
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
    let commit = package
        .edit_table_headers(sheet.sheet_selector(), table.table_selector())?
        .set(settings)
        .commit()?;

    let inverse = inverse_output
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_table_headers(&commit.patch().inverse())
        })
        .transpose()?;
    if let Some(restored) = inverse.as_ref() {
        if exact_bytes(restored.package())? != exact_bytes(&package)? {
            return Err(invalid_input(
                "inverse patch did not restore the exact input package",
            ));
        }
    }

    save_new(&output, commit.package())?;
    if let (Some(path), Some(restored)) = (inverse_output, inverse) {
        save_new(&path, restored.package())?;
    }

    println!(
        "table headers: changed={}, touched_components={}, deleted_previews={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
    );
    Ok(())
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
        .map_err(|_| invalid_input("selectors and settings arguments must be valid UTF-8"))
}

fn parse_selector(value: String) -> Result<Selector, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(Selector::Index)
            .map_err(|_| invalid_input("selector index must be a non-negative integer"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        return Ok(Selector::Name(name.to_owned()));
    }
    Err(invalid_input("selector must start with index: or name:"))
}

fn parse_count(
    arguments: &mut impl Iterator<Item = OsString>,
    label: &'static str,
) -> Result<Option<Count>, Box<dyn Error>> {
    let value = required_text(arguments, format!("missing {label}"))?;
    if value == "unset" {
        return Ok(None);
    }
    value
        .parse::<usize>()
        .map_err(|_| {
            invalid_input(format!(
                "{label} must be unset or an integer from 1 through 5"
            ))
        })
        .and_then(|count| {
            Count::new(count)
                .map(Some)
                .map_err(|error| -> Box<dyn Error> { Box::new(error) })
        })
}

fn parse_optional_bool(
    arguments: &mut impl Iterator<Item = OsString>,
    label: &'static str,
) -> Result<Option<bool>, Box<dyn Error>> {
    match required_text(arguments, format!("missing {label}"))?.as_str() {
        "unset" => Ok(None),
        "true" => Ok(Some(true)),
        "false" => Ok(Some(false)),
        _ => Err(invalid_input(format!(
            "{label} must be unset, true, or false"
        ))),
    }
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
