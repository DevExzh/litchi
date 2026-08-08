//! Set, clear, or replace a UTF-16 span in one Pages section.

use std::error::Error;
use std::ffi::OsString;
use std::fs::OpenOptions;
use std::io::Write as _;
use std::path::{Path, PathBuf};

use litchi_pages::{Package, SectionSelector, TextSpan};

const USAGE: &str = "usage: edit_section_text <input.pages> <output.pages> \
                     <index SECTION_INDEX|name SECTION_NAME> \
                     <set TEXT|clear|range UTF16_START UTF16_END TEXT> \
                     [inverse-output.pages]";

enum SelectedSection {
    Index(usize),
    Name(String),
}

impl SelectedSection {
    fn selector(&self) -> SectionSelector<'_> {
        match self {
            Self::Index(index) => SectionSelector::index(*index),
            Self::Name(name) => SectionSelector::name(name),
        }
    }
}

enum Operation {
    Set(String),
    Clear,
    Range { span: TextSpan, replacement: String },
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(USAGE)?);
    let output = PathBuf::from(arguments.next().ok_or(USAGE)?);
    let selector = parse_selector(&mut arguments)?;
    let operation = parse_operation(&mut arguments)?;
    let inverse_output = arguments.next().map(PathBuf::from);
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }
    if output == input {
        return Err("input and output paths must differ".into());
    }
    if inverse_output
        .as_deref()
        .is_some_and(|path| path == output || path == input)
    {
        return Err("inverse-output path must differ from input and output".into());
    }

    let package = Package::open(&input)?;
    let mut edit = package.edit_section_text(selector.selector())?;
    match operation {
        Operation::Set(replacement) => edit.set(&replacement)?,
        Operation::Clear => edit.clear()?,
        Operation::Range { span, replacement } => edit.replace(span, &replacement)?,
    };
    let commit = edit.commit()?;
    write_new(&output, commit.package().source_bytes())?;

    if let Some(inverse_path) = inverse_output {
        let inverse = commit.patch().inverse();
        let restored = commit.package().apply_section_text(&inverse)?;
        write_new(&inverse_path, restored.package().source_bytes())?;
    }
    Ok(())
}

fn parse_selector(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<SelectedSection, Box<dyn Error>> {
    let kind = text_argument(arguments.next(), "missing section selector kind")?;
    let value = text_argument(arguments.next(), "missing section selector value")?;
    match kind.as_str() {
        "index" => Ok(SelectedSection::Index(value.parse()?)),
        "name" => Ok(SelectedSection::Name(value)),
        _ => Err("section selector must be index or name".into()),
    }
}

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match text_argument(arguments.next(), "missing edit mode")?.as_str() {
        "set" => Ok(Operation::Set(text_argument(
            arguments.next(),
            "missing replacement text",
        )?)),
        "clear" => Ok(Operation::Clear),
        "range" => {
            let start = text_argument(arguments.next(), "missing UTF-16 start")?.parse()?;
            let end = text_argument(arguments.next(), "missing UTF-16 end")?.parse()?;
            let replacement = text_argument(arguments.next(), "missing replacement text")?;
            Ok(Operation::Range {
                span: TextSpan::from_utf16_indexes(start, end)?,
                replacement,
            })
        },
        _ => Err("edit mode must be set, clear, or range".into()),
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
