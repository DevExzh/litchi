//! Move one Numbers sheet through an exact immutable order transaction.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::{Package, SheetSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_sheet_order <input.numbers> <output.numbers> \\
                     <index:N|name:NAME> <index:N> [--inverse PATH]";

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
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let selector = parse_selector(required_text(&mut arguments, "missing sheet selector")?)?;
    let destination = parse_destination(required_text(&mut arguments, "missing destination")?)?;
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
    let before = sheet_names(&package);
    let source_position = resolve_source_position(&before, &selector)?;
    if destination >= before.len() {
        return Err(invalid_input(format!(
            "destination index {destination} is outside {} sheets",
            before.len()
        )));
    }
    let mut expected = before.clone();
    let moved = expected.remove(source_position);
    expected.insert(destination, moved);

    let commit = package
        .edit_sheet_order()
        .move_sheet(selector.sheet_selector(), destination)?
        .commit()?;
    if sheet_names(commit.package()) != expected {
        return Err(invalid_input(
            "committed sheet order did not match the requested move",
        ));
    }

    let inverse = inverse_output
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_sheet_order(&commit.patch().inverse())
        })
        .transpose()?;
    if let Some(restored) = inverse.as_ref() {
        if exact_bytes(restored.package())? != exact_bytes(&package)? {
            return Err(invalid_input(
                "inverse patch did not restore the exact input package",
            ));
        }
        if sheet_names(restored.package()) != before {
            return Err(invalid_input(
                "inverse patch did not restore the original sheet order",
            ));
        }
    }

    save_new(&output, commit.package())?;
    if let (Some(path), Some(restored)) = (inverse_output, inverse) {
        save_new(&path, restored.package())?;
    }

    println!(
        "sheet order: changed={}, touched_components={}, deleted_previews={}, full_reparse_performed={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_selector(value: String) -> Result<Selector, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(Selector::Index)
            .map_err(|_| invalid_input("sheet selector index must be a non-negative integer"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input("sheet selector name must not be empty"));
        }
        return Ok(Selector::Name(name.to_owned()));
    }
    Err(invalid_input(
        "sheet selector must start with index: or name:",
    ))
}

fn parse_destination(value: String) -> Result<usize, Box<dyn Error>> {
    let Some(index) = value.strip_prefix("index:") else {
        return Err(invalid_input("destination must start with index:"));
    };
    index
        .parse()
        .map_err(|_| invalid_input("destination index must be a non-negative integer"))
}

fn resolve_source_position(names: &[String], selector: &Selector) -> Result<usize, Box<dyn Error>> {
    match selector {
        Selector::Index(index) => names
            .get(*index)
            .map(|_| *index)
            .ok_or_else(|| invalid_input(format!("sheet index {index} does not exist"))),
        Selector::Name(name) => {
            let mut matches = names
                .iter()
                .enumerate()
                .filter_map(|(index, candidate)| (candidate == name).then_some(index));
            let Some(index) = matches.next() else {
                return Err(invalid_input("sheet selector did not resolve"));
            };
            if matches.next().is_some() {
                return Err(invalid_input("sheet selector name is ambiguous"));
            }
            Ok(index)
        },
    }
}

fn sheet_names(package: &Package) -> Vec<String> {
    package
        .sheets()
        .iter()
        .map(|sheet| sheet.name().to_owned())
        .collect()
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
        .map_err(|_| invalid_input("sheet-order arguments must be valid UTF-8"))
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
