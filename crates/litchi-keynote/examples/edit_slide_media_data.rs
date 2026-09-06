//! Replace one existing Keynote slide movie or audio payload.
//!
//! The command accepts semantic slide and media positions.  Package metadata,
//! IWA object identifiers, and ZIP member names stay inside
//! [`litchi_keynote::Package`].  The output paths are required to be new files
//! so an invocation cannot accidentally overwrite a caller-owned artifact.
//!
//! ```text
//! edit_slide_media_data <input.key> <output.key> <index:N|name:TEXT> <movie-index>
//!                        <content|poster> <replacement-file> [--inverse PATH]
//! ```

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports semantic edit diagnostics"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::fs;
use std::path::PathBuf;

use litchi_keynote::{MediaPart, MovieSelector, Package, SlideSelector};

const USAGE: &str = "usage: edit_slide_media_data <input.key> <output.key> \
                     <index:N|name:TEXT> <movie-index> <content|poster> \
                     <replacement-file> [--inverse PATH]";

#[derive(Debug)]
enum SlideArgument {
    Index(usize),
    Name(String),
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let slide = parse_slide(required_text(&mut arguments, "missing slide selector")?)?;
    let movie = parse_index(required_text(&mut arguments, "missing movie position")?)?;
    let part = parse_part(required_text(&mut arguments, "missing media part")?)?;
    let replacement_path = PathBuf::from(required_argument(
        &mut arguments,
        "missing replacement path",
    )?);
    let inverse = parse_inverse(&mut arguments)?;

    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }
    if inverse
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "inverse path must differ from input and output paths",
        ));
    }
    refuse_existing(&output, "output")?;
    if let Some(path) = inverse.as_deref() {
        refuse_existing(path, "inverse output")?;
    }

    let replacement = fs::read(&replacement_path)?;
    let package = Package::open(&input)?;
    let edit = match &slide {
        SlideArgument::Index(index) => package.edit_slide_media_data(
            SlideSelector::index(*index),
            MovieSelector::index(movie),
            part,
        )?,
        SlideArgument::Name(name) => package.edit_slide_media_data(
            SlideSelector::name(name.as_str()),
            MovieSelector::index(movie),
            part,
        )?,
    };
    let commit = edit.set(&replacement)?.commit()?;

    // Publish the forward artifact only after all source and candidate checks
    // have succeeded.  The inverse is derived from the exact patch, so it can
    // be emitted without exposing or reconstructing native media references.
    commit.package().save(&output)?;
    if let Some(path) = inverse {
        let restored = commit
            .package()
            .apply_slide_media_data(&commit.patch().inverse())?;
        restored.package().save(path)?;
    }

    println!(
        "slide media: part={:?}, before_bytes={}, after_bytes={}, changed={}, touched_components={}, deleted_previews={}",
        commit.patch().part(),
        commit.patch().before_length(),
        commit.patch().after_length(),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
    );
    Ok(())
}

fn parse_slide(value: String) -> Result<SlideArgument, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return Ok(SlideArgument::Index(parse_index(index.to_owned())?));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input("slide name must not be empty"));
        }
        return Ok(SlideArgument::Name(name.to_owned()));
    }
    Err(invalid_input(
        "slide selector must use index:N or name:TEXT",
    ))
}

fn parse_part(value: String) -> Result<MediaPart, Box<dyn Error>> {
    match value.as_str() {
        "content" => Ok(MediaPart::Content),
        "poster" => Ok(MediaPart::Poster),
        _ => Err(invalid_input("media part must be content or poster")),
    }
}

fn parse_index(value: String) -> Result<usize, Box<dyn Error>> {
    value
        .parse::<usize>()
        .map_err(|_| invalid_input("positions must be non-negative integers"))
}

fn parse_inverse(
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
    let path = PathBuf::from(required_argument(arguments, "missing inverse path")?);
    if arguments.next().is_some() {
        return Err(invalid_input(
            "unexpected trailing argument after inverse path",
        ));
    }
    Ok(Some(path))
}

fn refuse_existing(path: &std::path::Path, label: &str) -> Result<(), Box<dyn Error>> {
    match fs::symlink_metadata(path) {
        Ok(_) => Err(invalid_input(format!(
            "{label} path already exists; refusing to clobber it"
        ))),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error.into()),
    }
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
        .map_err(|_| invalid_input("arguments must be valid UTF-8"))
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(std::io::Error::new(
        std::io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
