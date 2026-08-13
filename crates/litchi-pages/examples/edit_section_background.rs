//! Set or clear one Pages section background through a semantic selector.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io::{self, Write as _};
use std::path::{Path, PathBuf};

use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_pages::{Package, SectionSelector, section::Background};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_section_background <input.pages> <output.pages> \\
                     <index:N|name:NAME> <clear|solid> \\
                     [RED GREEN BLUE ALPHA srgb|p3] [--inverse PATH]";

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

enum Operation {
    Clear,
    Set(Background),
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let selector = parse_selector(required_text(&mut arguments, "missing section selector")?)?;
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
    let before = package.section_background(selector.as_section_selector())?;
    let requested = match operation {
        Operation::Clear => Background::None,
        Operation::Set(background) => background,
    };
    let mut edit = package.edit_section_background(selector.as_section_selector())?;
    match requested {
        Background::None => edit.clear(),
        Background::Solid(color) => edit.set_solid(color)?,
        Background::Unsupported => {
            return Err(invalid_input(
                "the CLI can only author clear or semantic solid backgrounds",
            ));
        },
        _ => return Err(invalid_input("unknown section background representation")),
    };
    let commit = edit.commit()?;

    let readback = commit
        .package()
        .section_background(selector.as_section_selector())?;
    if readback != requested {
        return Err(invalid_input(
            "committed section background did not match the requested semantic background",
        ));
    }

    let restored = commit
        .package()
        .apply_section_background(&commit.patch().inverse())?;
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
        "section background: {} -> {}, changed={}, touched_components={}, deleted_previews={}",
        describe_background(&before),
        describe_background(&readback),
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

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match required_text(arguments, "missing background operation")?.as_str() {
        "clear" => Ok(Operation::Clear),
        "solid" => {
            let red = parse_channel(arguments, "red")?;
            let green = parse_channel(arguments, "green")?;
            let blue = parse_channel(arguments, "blue")?;
            let alpha = parse_channel(arguments, "alpha")?;
            let color_space = match required_text(arguments, "missing color space")?.as_str() {
                "srgb" => RgbColorSpace::Srgb,
                "p3" => RgbColorSpace::DisplayP3,
                _ => return Err(invalid_input("color space must be srgb or p3")),
            };
            Ok(Operation::Set(Background::Solid(Rgba::new(
                red,
                green,
                blue,
                alpha,
                color_space,
            )?)))
        },
        _ => Err(invalid_input("background operation must be clear or solid")),
    }
}

fn parse_channel(
    arguments: &mut impl Iterator<Item = OsString>,
    channel: &'static str,
) -> Result<f32, Box<dyn Error>> {
    required_text(arguments, format!("missing {channel} channel"))?
        .parse()
        .map_err(|_| invalid_input(format!("{channel} channel must be a finite number")))
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
        .map_err(|_| invalid_input("selector and background arguments must be valid UTF-8"))
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

fn describe_background(background: &Background) -> &'static str {
    match background {
        Background::None => "none",
        Background::Solid(_) => "solid",
        Background::Unsupported => "unsupported",
        _ => "unknown",
    }
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
