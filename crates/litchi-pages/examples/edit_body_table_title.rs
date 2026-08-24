//! Edit one rooted Pages body table's title visibility and outline settings.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::io;
use std::path::{Path, PathBuf};

use litchi_pages::table::title::Settings;
use litchi_pages::{BodyTableSelector, Package};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_body_table_title <input.pages> <output.pages> \
                     <index:N|name:NAME> <absent|false|true> <absent|false|true> \
                     [inverse.pages]";

enum Selector {
    Index(usize),
    Name(String),
}

impl Selector {
    fn as_body_table_selector(&self) -> BodyTableSelector<'_> {
        match self {
            Self::Index(index) => BodyTableSelector::position((*index).into()),
            Self::Name(name) => BodyTableSelector::name(name),
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args().skip(1);
    let input = PathBuf::from(required(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required(&mut arguments, "missing output path")?);
    let selector = parse_selector(required(&mut arguments, "missing table selector")?)?;
    let visible = parse_optional_bool(required(&mut arguments, "missing visible value")?)?;
    let outlined = parse_optional_bool(required(&mut arguments, "missing outlined value")?)?;
    let inverse_output = arguments.next().map(PathBuf::from);
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    if input == output
        || inverse_output
            .as_deref()
            .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "input, output, and inverse paths must differ",
        ));
    }

    let package = Package::open(&input)?;
    let before = package.body_table_title_settings(selector.as_body_table_selector())?;
    let requested = Settings::new(visible, outlined);
    let commit = package
        .edit_body_table_title(selector.as_body_table_selector())?
        .set(requested)
        .commit()?;
    if commit
        .package()
        .body_table_title_settings(selector.as_body_table_selector())?
        != requested
    {
        return Err(invalid_input(
            "committed title settings did not match the request",
        ));
    }
    let restored = commit
        .package()
        .apply_body_table_title(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != exact_bytes(&package)? {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    save_new(&output, commit.package())?;
    if let Some(path) = inverse_output {
        save_new(&path, restored.package())?;
    }
    println!(
        "body-table title: {before:?} -> {requested:?}, changed={}, touched_components={}, deleted_previews={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
    );
    Ok(())
}

fn parse_selector(value: String) -> Result<Selector, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(Selector::Index)
            .map_err(|_| invalid_input("selector index must be a non-negative integer"));
    }
    value
        .strip_prefix("name:")
        .map(|name| Selector::Name(name.to_owned()))
        .ok_or_else(|| invalid_input("selector must start with index: or name:"))
}

fn parse_optional_bool(value: String) -> Result<Option<bool>, Box<dyn Error>> {
    match value.as_str() {
        "absent" => Ok(None),
        "false" => Ok(Some(false)),
        "true" => Ok(Some(true)),
        _ => Err(invalid_input("title values must be absent, false, or true")),
    }
}

fn required(
    arguments: &mut impl Iterator<Item = String>,
    message: &'static str,
) -> Result<String, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

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
