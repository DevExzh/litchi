//! Set, clear, or replace a UTF-16 span in one slide's existing speaker notes.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::fs::OpenOptions;
use std::io::Write as _;
use std::path::{Path, PathBuf};

use litchi_keynote::{Package, SlideSelector, TextSpan};

const USAGE: &str = "usage: edit_slide_notes <input.key> <output.key> \
                     <index:N|name:NAME> \
                     <set NOTES|clear|replace UTF16_START UTF16_END NOTES> \
                     [inverse-output.key]";

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
    let input = PathBuf::from(arguments.next().ok_or(USAGE)?);
    let output = PathBuf::from(arguments.next().ok_or(USAGE)?);
    let selector = parse_selector(arguments.next())?;
    let operation = parse_operation(&mut arguments)?;
    let inverse_output = arguments.next().map(PathBuf::from);
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }
    if input == output {
        return Err("input and output paths must differ".into());
    }
    if inverse_output
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err("inverse-output path must differ from input and output".into());
    }

    let package = Package::open(&input)?;
    let mut edit = package.edit_slide_notes(selector.selector())?;
    match operation {
        Operation::Set(notes) => {
            edit.set(&notes)?;
        },
        Operation::Clear => {
            edit.clear()?;
        },
        Operation::Replace { span, replacement } => {
            edit.replace(span, &replacement)?;
        },
    }
    let commit = edit.commit()?;
    write_new(&output, commit.package().source_bytes())?;

    if let Some(inverse_path) = inverse_output {
        let restored = commit
            .package()
            .apply_slide_notes(&commit.patch().inverse())?;
        write_new(&inverse_path, restored.package().source_bytes())?;
    }

    println!(
        "slide notes: changed={}, touched_components={}, full_reparse={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_selector(argument: Option<OsString>) -> Result<SelectedSlide, Box<dyn Error>> {
    let value = text_argument(argument, "missing slide selector")?;
    if let Some(index) = value.strip_prefix("index:") {
        return Ok(SelectedSlide::Index(index.parse()?));
    }
    if let Some(name) = value.strip_prefix("name:") {
        return Ok(SelectedSlide::Name(name.to_owned()));
    }
    Err("selector must start with index: or name:".into())
}

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match text_argument(arguments.next(), "missing notes operation")?.as_str() {
        "set" => Ok(Operation::Set(text_argument(
            arguments.next(),
            "missing speaker notes",
        )?)),
        "clear" => Ok(Operation::Clear),
        "replace" => {
            let start = text_argument(arguments.next(), "missing UTF-16 start")?.parse()?;
            let end = text_argument(arguments.next(), "missing UTF-16 end")?.parse()?;
            let replacement = text_argument(arguments.next(), "missing replacement notes")?;
            Ok(Operation::Replace {
                span: TextSpan::from_utf16_indexes(start, end)?,
                replacement,
            })
        },
        _ => Err("notes operation must be set, clear, or replace".into()),
    }
}

fn text_argument(
    argument: Option<OsString>,
    missing: &'static str,
) -> Result<String, Box<dyn Error>> {
    argument.ok_or_else(|| missing.into()).and_then(|value| {
        value
            .into_string()
            .map_err(|_value| "argument is not valid UTF-8".into())
    })
}

fn write_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let mut destination = OpenOptions::new().write(true).create_new(true).open(path)?;
    destination.write_all(bytes)?;
    destination.sync_all()?;
    Ok(())
}
