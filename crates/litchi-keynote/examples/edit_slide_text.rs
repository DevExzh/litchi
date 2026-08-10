//! Edit an existing slide title or body without exposing Keynote internals.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io::{self, Write as _};
use std::path::{Path, PathBuf};

use litchi_keynote::{Package, SlideSelector, SlideTextRole, TextSpan};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_slide_text <input.key> <output.key> \\
                     <index:N|name:NAME> <title|body> \\
                     <set TEXT|clear|replace UTF16_START UTF16_END TEXT> \\
                     [--inverse PATH]";

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

enum Operation {
    Set(String),
    Clear,
    Replace { span: TextSpan, replacement: String },
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let slide_argument = required_text(&mut arguments, "missing slide selector")?;
    let slide = parse_selector(&slide_argument)?;
    let role_argument = required_text(&mut arguments, "missing text role")?;
    let role = parse_role(&role_argument)?;
    let operation = parse_operation(&mut arguments)?;
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
    let mut edit = package.edit_slide_text(slide.selector(), role)?;
    match operation {
        Operation::Set(text) => edit.set(&text)?,
        Operation::Clear => edit.clear()?,
        Operation::Replace { span, replacement } => edit.replace(span, &replacement)?,
    };
    let commit = edit.commit()?;

    let inverse = inverse_output
        .as_ref()
        .map(|_| commit.package().apply_slide_text(&commit.patch().inverse()))
        .transpose()?;
    if inverse
        .as_ref()
        .is_some_and(|restored| restored.package().source_bytes() != package.source_bytes())
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    save_new(&output, commit.package().source_bytes())?;
    if let (Some(path), Some(restored)) = (inverse_output, inverse) {
        save_new(&path, restored.package().source_bytes())?;
    }

    println!(
        "slide {role}: changed={}, touched_components={}",
        commit.diagnostics().changed(),
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

fn parse_role(value: &str) -> Result<SlideTextRole, Box<dyn Error>> {
    match value {
        "title" => Ok(SlideTextRole::Title),
        "body" => Ok(SlideTextRole::Body),
        _ => Err(invalid_input("text role must be title or body")),
    }
}

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match required_text(arguments, "missing text operation")?.as_str() {
        "set" => Ok(Operation::Set(required_text(
            arguments,
            "missing replacement text",
        )?)),
        "clear" => Ok(Operation::Clear),
        "replace" => {
            let start_argument = required_text(arguments, "missing UTF-16 start")?;
            let start = parse_index(&start_argument, "start")?;
            let end_argument = required_text(arguments, "missing UTF-16 end")?;
            let end = parse_index(&end_argument, "end")?;
            let replacement = required_text(arguments, "missing replacement text")?;
            Ok(Operation::Replace {
                span: TextSpan::from_utf16_indexes(start, end)?,
                replacement,
            })
        },
        _ => Err(invalid_input(
            "text operation must be set, clear, or replace",
        )),
    }
}

fn parse_index(value: &str, label: &str) -> Result<usize, Box<dyn Error>> {
    value
        .parse()
        .map_err(|_parse| invalid_input(format!("UTF-16 {label} must be a non-negative integer")))
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
fn save_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    temporary.write_all(bytes)?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| -> Box<dyn Error> { Box::new(error.error) })?;
    Ok(())
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
