//! Delete one Keynote slide through the selector-first immutable transaction.
//!
//! This structural operation preserves media/data payloads; it is not package
//! garbage collection. It admits only the supported flat native ownership
//! topology and fails closed on a surviving inbound owner.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{Package, SlideSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: remove_slide <input.key> <output.key> \\
                     <index:N|name:NAME> [--inverse PATH]";

enum SelectedSlide {
    Index(usize),
    Name(String),
}

impl SelectedSlide {
    fn selector(&self) -> SlideSelector<'_> {
        match self {
            Self::Index(index) => SlideSelector::index(*index),
            Self::Name(name) => SlideSelector::name(name),
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let selector = parse_selector(&required_text(&mut arguments, "missing slide selector")?)?;
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
    let mut edit = package.edit_slide_deletion();
    edit.remove_slide(selector.selector())?;
    let commit = edit.commit()?;

    let inverse = inverse_output
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_slide_deletion(&commit.patch().inverse())
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
        "slide removed: slides_removed={}, touched_components={}",
        commit.diagnostics().slides_removed(),
        commit.diagnostics().touched_components(),
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
        .map_err(|_non_utf8| invalid_input("text arguments must be valid UTF-8"))
}

fn parse_selector(value: &str) -> Result<SelectedSlide, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(SelectedSlide::Index)
            .map_err(|_parse| invalid_input("slide index must be a non-negative integer"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        return Ok(SelectedSlide::Name(name.to_owned()));
    }
    Err(invalid_input(
        "slide selector must start with index: or name:",
    ))
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

/// Publishes through a sibling temporary file without replacing an existing target.
///
/// This example does not provide the library's durable atomic-save contract.
fn save_new(path: &Path, package: &Package) -> Result<(), Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(&mut temporary)?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| -> Box<dyn Error> { Box::new(error.error) })?;
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
