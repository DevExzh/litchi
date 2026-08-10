//! Set validated page dimensions and orientation without exposing Pages internals.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io::{self, Write as _};
use std::path::{Path, PathBuf};

use litchi_pages::{
    Package,
    page_layout::{Orientation, Result as LayoutResult},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_page_layout <input.pages> <output.pages> \\
                     <WIDTHxHEIGHT:portrait|landscape> [--inverse PATH]";

struct Mutation {
    width: f32,
    height: f32,
    orientation: Orientation,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let mutation = parse_mutation(required_text(&mut arguments, "missing layout mutation")?)?;
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
    let mut edit = package.edit_page_layout()?;
    let mut layout = edit.layout();
    apply_mutation(&mut layout, mutation)?;
    edit.set_layout(layout)?;
    let commit = edit.commit()?;

    let inverse = inverse_output
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_page_layout(&commit.patch().inverse())
        })
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
        "page layout: changed={}, touched_components={}, deleted_previews={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
    );
    Ok(())
}

fn apply_mutation(
    layout: &mut litchi_pages::page_layout::Layout,
    mutation: Mutation,
) -> LayoutResult<()> {
    layout.set_page_width(Some(mutation.width))?;
    layout.set_page_height(Some(mutation.height))?;
    layout.set_orientation(Some(mutation.orientation))
}

fn parse_mutation(value: String) -> Result<Mutation, Box<dyn Error>> {
    let (dimensions, orientation) = value
        .rsplit_once(':')
        .ok_or_else(|| invalid_input("layout mutation must contain :portrait or :landscape"))?;
    let (width, height) = dimensions
        .split_once('x')
        .ok_or_else(|| invalid_input("layout mutation must use WIDTHxHEIGHT"))?;
    let width = parse_dimension(width, "width")?;
    let height = parse_dimension(height, "height")?;
    let orientation = match orientation {
        "portrait" => Orientation::Portrait,
        "landscape" => Orientation::Landscape,
        _ => return Err(invalid_input("orientation must be portrait or landscape")),
    };
    Ok(Mutation {
        width,
        height,
        orientation,
    })
}

fn parse_dimension(value: &str, name: &str) -> Result<f32, Box<dyn Error>> {
    value
        .parse()
        .map_err(|_| invalid_input(format!("page {name} must be a finite number")))
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
        .map_err(|_| invalid_input("layout arguments must be valid UTF-8"))
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
        .map_err(|error| Box::new(error.error))?;
    Ok(())
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
