//! Read and replace one Pages body table's hidden rows and columns.
//!
//! The table is selected by its semantic name or body position.  Hidden axes
//! are supplied as zero-based semantic positions, for example
//! `row:2,column:1`; no Pages-native object identifier is required.
//!
//! Changed writes require an existing native hidden-state owner.  A table with
//! no owner can still be read as having no hidden axes, but this example will
//! refuse a non-empty request instead of synthesizing a Pages-internal graph.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};

use litchi_pages::table::hidden_axes::{AxisIndex, HiddenAxes};
use litchi_pages::{BodyTableSelector, Package};

const USAGE: &str = "usage: edit_body_table_hidden_axes <input.pages> <output.pages> \
                     <index:N|name:NAME> <none|row:N[,row:N...][,column:N...]> \
                     [inverse.pages]\n\n                     changed writes require an existing hidden-state owner";

enum Selector {
    Index(usize),
    Name(String),
}

impl Selector {
    fn as_body_table_selector(&self) -> BodyTableSelector<'_> {
        match self {
            Self::Index(index) => BodyTableSelector::index(*index),
            Self::Name(name) => BodyTableSelector::name(name),
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_os(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_os(&mut arguments, "missing output path")?);
    let selector = parse_selector(required_text(&mut arguments, "missing table selector")?)?;
    let requested = parse_hidden_axes(required_text(&mut arguments, "missing hidden-axis value")?)?;
    let inverse_output = arguments.next().map(PathBuf::from);
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    let output_aliases_input = paths_alias(&input, &output)?;
    let inverse_aliases_input = inverse_output
        .as_deref()
        .map(|path| paths_alias(&input, path))
        .transpose()?
        .unwrap_or(false);
    let inverse_aliases_output = inverse_output
        .as_deref()
        .map(|path| paths_alias(&output, path))
        .transpose()?
        .unwrap_or(false);
    if output_aliases_input || inverse_aliases_input || inverse_aliases_output {
        return Err(invalid_input(
            "input, output, and inverse paths must identify different files",
        ));
    }

    let package = Package::open(&input)?;
    let before = package.body_table_hidden_axes(selector.as_body_table_selector())?;
    let commit = package
        .edit_body_table_hidden_axes(selector.as_body_table_selector())?
        .set(requested.clone())
        .commit()?;
    let after = commit
        .package()
        .body_table_hidden_axes(selector.as_body_table_selector())?;
    if after != requested {
        return Err(invalid_input(
            "committed hidden rows and columns did not match the request",
        ));
    }

    // Apply the exact inverse in memory and prove that it restores every
    // source byte, including unsupported package members and previews.
    let restored = commit
        .package()
        .apply_body_table_hidden_axes(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != exact_bytes(&package)? {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    commit.package().save(&output)?;
    let reopened = Package::open(&output)?;
    if reopened.body_table_hidden_axes(selector.as_body_table_selector())? != requested {
        return Err(invalid_input(
            "saved package did not preserve hidden rows and columns",
        ));
    }
    if let Some(path) = inverse_output {
        restored.package().save(&path)?;
        let inverse_reopened = Package::open(&path)?;
        if inverse_reopened.body_table_hidden_axes(selector.as_body_table_selector())? != before {
            return Err(invalid_input(
                "saved inverse package did not restore hidden rows and columns",
            ));
        }
    }

    println!("body-table hidden axes: {before:?} -> {requested:?}");
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

fn parse_hidden_axes(value: String) -> Result<HiddenAxes, Box<dyn Error>> {
    if value == "none" {
        return Ok(HiddenAxes::empty());
    }
    if value.is_empty() {
        return Err(invalid_input(
            "hidden axes must be none or a comma-separated list",
        ));
    }

    let mut axes = Vec::new();
    for item in value.split(',') {
        let (kind, index) = item
            .split_once(':')
            .ok_or_else(|| invalid_input("hidden axes must use row:N or column:N"))?;
        let index = index
            .parse::<usize>()
            .map_err(|_| invalid_input("hidden-axis positions must be non-negative integers"))?;
        let axis = match kind {
            "row" => AxisIndex::row(index),
            "column" => AxisIndex::column(index),
            _ => return Err(invalid_input("hidden axes must use row:N or column:N")),
        };
        axes.push(axis);
    }
    HiddenAxes::new(axes).map_err(|error| invalid_input(format!("invalid hidden axes: {error}")))
}

fn required_os(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

fn required_text(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<String, Box<dyn Error>> {
    required_os(arguments, message)?.into_string().map_err(|_| {
        invalid_input("selector and hidden-axis values must be valid UTF-8; paths may be non-UTF-8")
    })
}

fn paths_alias(left: &Path, right: &Path) -> io::Result<bool> {
    if canonical_or_future(left)? == canonical_or_future(right)? {
        return Ok(true);
    }
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;

        if left.exists() && right.exists() {
            let left_metadata = fs::metadata(left)?;
            let right_metadata = fs::metadata(right)?;
            return Ok(left_metadata.dev() == right_metadata.dev()
                && left_metadata.ino() == right_metadata.ino());
        }
    }
    Ok(false)
}

fn canonical_or_future(path: &Path) -> io::Result<PathBuf> {
    if path.exists() {
        return fs::canonicalize(path);
    }
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let name = path
        .file_name()
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "path has no file name"))?;
    Ok(fs::canonicalize(parent)?.join(name))
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
