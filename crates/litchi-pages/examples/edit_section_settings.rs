//! Edit the four presence-preserving Boolean settings of one Pages section.
//!
//! The section's name and pagination are read from the immutable source and
//! carried into the aggregate replacement unchanged. Each requested artifact
//! is published independently through a sibling temporary file without
//! clobbering an existing path. If optional inverse publication fails after
//! the main output succeeds, the main output remains.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io::{self, Write as _};
use std::path::{Path, PathBuf};

use litchi_pages::{
    Package, SectionSelector,
    section::{Settings, settings::Commit},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_section_settings <input.pages> <output.pages> \\
                     <index:N|name:NAME> \\
                     <unset|true|false:inherit-previous> \\
                     <unset|true|false:first-page-different> \\
                     <unset|true|false:even-odd-different> \\
                     <unset|true|false:hide-first-page-header-footer> \\
                     [--inverse PATH]";

enum Selector {
    Index(usize),
    Name(String),
}

impl Selector {
    fn as_section_selector(&self) -> SectionSelector<'_> {
        match self {
            Self::Index(index) => SectionSelector::index(*index),
            Self::Name(name) => SectionSelector::name(name),
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let selector = parse_selector(required_text(&mut arguments, "missing section selector")?)?;
    let inherit_previous = parse_optional_bool(&mut arguments, "inherit-previous")?;
    let first_page_different = parse_optional_bool(&mut arguments, "first-page-different")?;
    let even_odd_different = parse_optional_bool(&mut arguments, "even-odd-different")?;
    let hide_first_page = parse_optional_bool(&mut arguments, "hide-first-page-header-footer")?;
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
    let before = package.section_settings(selector.as_section_selector())?;
    let mut requested: Settings = before.clone();
    requested.set_inherit_previous_header_footer(inherit_previous);
    requested.set_first_page_different(first_page_different);
    requested.set_even_odd_pages_different(even_odd_different);
    requested.set_first_page_hides_header_footer(hide_first_page);

    let commit: Commit = package
        .edit_section_settings(selector.as_section_selector())?
        .set(requested)?
        .commit()?;
    let readback = commit
        .package()
        .section_settings(selector.as_section_selector())?;
    if &readback != commit.patch().after() {
        return Err(invalid_input(
            "committed section settings did not match the requested semantic settings",
        ));
    }

    let restored = commit
        .package()
        .apply_section_settings(&commit.patch().inverse())?;
    if restored.package().source_bytes() != package.source_bytes() {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    save_new(&output, commit.package())?;
    if let Some(path) = inverse_output {
        save_new(&path, restored.package())?;
    }

    println!(
        "section settings: inherit-previous={}->{} first-page-different={}->{} \
         even-odd-different={}->{} hide-first-page-header-footer={}->{} \
         changed={}, touched_components={}, deleted_previews={}",
        optional_bool(before.inherit_previous_header_footer()),
        optional_bool(readback.inherit_previous_header_footer()),
        optional_bool(before.first_page_different()),
        optional_bool(readback.first_page_different()),
        optional_bool(before.even_odd_pages_different()),
        optional_bool(readback.even_odd_pages_different()),
        optional_bool(before.first_page_hides_header_footer()),
        optional_bool(readback.first_page_hides_header_footer()),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
    );
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
        return Ok(Selector::Name(name.to_owned()));
    }
    Err(invalid_input("selector must start with index: or name:"))
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
        .map_err(|_| invalid_input("selector and settings arguments must be valid UTF-8"))
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
    temporary.write_all(package.source_bytes())?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| Box::new(error.error))?;
    Ok(())
}

const fn optional_bool(value: Option<bool>) -> &'static str {
    match value {
        None => "unset",
        Some(false) => "false",
        Some(true) => "true",
    }
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
