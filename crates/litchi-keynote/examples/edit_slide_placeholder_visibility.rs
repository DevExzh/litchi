//! Show or hide one layout-provided Keynote title or body placeholder.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{
    Package, SlideSelector,
    slide::placeholder::{Kind, State},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_slide_placeholder_visibility <input.key> <output.key> \\
                     <index:N|name:NAME> <title|body> <show|hide> [--inverse PATH]";

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
    let slide = parse_selector(&required_text(&mut arguments, "missing slide selector")?)?;
    let kind = parse_kind(&required_text(&mut arguments, "missing placeholder kind")?)?;
    let state = parse_state(&required_text(&mut arguments, "missing placeholder state")?)?;
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
    let before = package
        .slide_placeholder_visibility(slide.selector(), kind)?
        .ok_or_else(|| invalid_input("selected slide has no placeholder for the requested kind"))?;
    let commit = package
        .edit_slide_placeholder_visibility(slide.selector(), kind)?
        .set(state)
        .commit()?;
    if commit
        .package()
        .slide_placeholder_visibility(slide.selector(), kind)?
        != Some(state)
    {
        return Err(invalid_input(
            "committed placeholder visibility did not match the requested state",
        ));
    }

    let inverse = inverse_output
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_slide_placeholder_visibility(&commit.patch().inverse())
        })
        .transpose()?;
    if let Some(restored) = inverse.as_ref() {
        if exact_bytes(restored.package())? != exact_bytes(&package)? {
            return Err(invalid_input(
                "inverse patch did not restore the exact input package",
            ));
        }
        if restored
            .package()
            .slide_placeholder_visibility(slide.selector(), kind)?
            != Some(before)
        {
            return Err(invalid_input(
                "inverse patch did not restore the original placeholder state",
            ));
        }
    }

    save_new(&output, commit.package())?;
    if let (Some(path), Some(restored)) = (inverse_output, inverse) {
        save_new(&path, restored.package())?;
    }

    println!(
        "slide {kind} placeholder: changed={}, touched_components={}, deleted_previews={}",
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
        .map_err(|_error| invalid_input("slide selector, kind, and state must be valid UTF-8"))
}

fn parse_selector(value: &str) -> Result<SelectedSlide, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(SelectedSlide::Index)
            .map_err(|_error| invalid_input("slide index must be a non-negative integer"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        return Ok(SelectedSlide::Name(name.to_owned()));
    }
    Err(invalid_input(
        "slide selector must start with index: or name:",
    ))
}

fn parse_kind(value: &str) -> Result<Kind, Box<dyn Error>> {
    match value {
        "title" => Ok(Kind::Title),
        "body" => Ok(Kind::Body),
        _ => Err(invalid_input("placeholder kind must be title or body")),
    }
}

fn parse_state(value: &str) -> Result<State, Box<dyn Error>> {
    match value {
        "show" => Ok(State::Visible),
        "hide" => Ok(State::Hidden),
        _ => Err(invalid_input("placeholder state must be show or hide")),
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
