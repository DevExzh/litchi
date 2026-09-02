//! Execute a Keynote table sort through the selector-first physical owner.
//!
//! The command accepts a slide position or exact navigator name and a
//! zero-based table position.  Rows, when supplied, are body-relative and
//! half-open; headers and footers are never addressed by this interface.
//! It executes the table's existing persisted sort order; it does not create
//! or change that order.  The focused owner admits only package-backed tables
//! with present, homogeneous scalar keys (text, finite number, boolean, date,
//! or duration) and a supported row-local topology.  Unsupported formulas,
//! errors, rich text, missing keys, mixed key kinds, and unsupported
//! row-affine structures are rejected before output publication.  This
//! example demonstrates the Rust package contract and does not claim that
//! every producer-authored package is accepted by native Keynote.
//!
//! ```text
//! execute_slide_table_sort <input.key> <output.key> <index:N|name:NAME> <table-index>
//!     [--full | --rows START END] [--inverse PATH] [--dry-run]
//! ```

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports a content-free transaction diagnostic"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::fs;
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::slide::table::physical_sort::RowRange;
use litchi_keynote::{Package, SlideSelector, TableSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: execute_slide_table_sort <input.key> <output.key> \
                     <index:N|name:NAME> <table-index> \
                     [--full | --rows START END] [--inverse PATH] [--dry-run]";

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

#[derive(Default)]
struct OutputOptions {
    rows: Option<RowRange>,
    inverse: Option<PathBuf>,
    dry_run: bool,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let slide = parse_selector(required_text(&mut arguments, "missing slide selector")?)?;
    let table = parse_table(required_text(&mut arguments, "missing table index")?)?;
    let options = parse_options(&mut arguments)?;

    ensure_distinct_paths(&input, &output, "input and output paths must differ")?;
    if let Some(inverse) = options.inverse.as_deref() {
        ensure_distinct_paths(
            &input,
            inverse,
            "inverse path must differ from the input path",
        )?;
        ensure_distinct_paths(
            &output,
            inverse,
            "inverse path must differ from the output path",
        )?;
    }

    let package = Package::open(&input)?;
    let table = TableSelector::index(table);
    let commit = match options.rows {
        Some(rows) => {
            package.execute_slide_table_sort_order_to_rows(slide.selector(), table, rows)?
        },
        None => package.execute_slide_table_sort_order(slide.selector(), table)?,
    };

    // Validate the reversible physical transaction in memory before any
    // filesystem publication.  Comparing emitted bytes catches both missed
    // row-affine structures and non-lossless wire rewrites.
    let restored = commit
        .package()
        .apply_slide_table_physical_sort(&commit.patch().inverse())?;
    if exact_bytes(restored.package())? != exact_bytes(&package)? {
        return Err(invalid_input(
            "inverse physical sort did not restore the exact input package",
        ));
    }

    if !options.dry_run {
        save_new(&output, commit.package())?;
        if let Some(path) = options.inverse {
            save_new(&path, restored.package())?;
        }
    }

    println!(
        "slide table sort: changed={}, moved_rows={}, touched_components={}, deleted_previews={}, full_reparse={}, dry_run={}",
        commit.diagnostics().changed(),
        commit.diagnostics().moved_rows(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
        commit.diagnostics().full_reparse_performed(),
        options.dry_run,
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
        .map_err(|_| invalid_input("command-line arguments must be valid UTF-8"))
}

fn parse_selector(value: String) -> Result<SelectedSlide, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        return index
            .parse()
            .map(SelectedSlide::Index)
            .map_err(|_| invalid_input("slide index must be a non-negative integer"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input("slide name must not be empty"));
        }
        return Ok(SelectedSlide::Name(name.to_owned()));
    }
    Err(invalid_input(
        "slide selector must start with index: or name:",
    ))
}

fn parse_table(value: String) -> Result<usize, Box<dyn Error>> {
    value
        .parse()
        .map_err(|_| invalid_input("table index must be a non-negative integer"))
}

fn parse_options(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<OutputOptions, Box<dyn Error>> {
    let mut options = OutputOptions::default();
    let mut full = false;
    while let Some(argument) = arguments.next() {
        if argument == OsStr::new("--full") {
            if full || options.rows.is_some() {
                return Err(invalid_input("sort scope may be specified only once"));
            }
            full = true;
        } else if argument == OsStr::new("--rows") {
            if full || options.rows.is_some() {
                return Err(invalid_input("sort scope may be specified only once"));
            }
            let start = parse_row(
                required_text(arguments, "missing body-relative row start")?,
                "row start",
            )?;
            let end = parse_row(
                required_text(arguments, "missing body-relative row end")?,
                "row end",
            )?;
            options.rows = Some(
                RowRange::new(start, end)
                    .map_err(|_| invalid_input("body-relative row range must be non-empty"))?,
            );
        } else if argument == OsStr::new("--inverse") {
            if options.inverse.is_some() {
                return Err(invalid_input("--inverse may be specified only once"));
            }
            let path_argument = required_argument(arguments, "missing --inverse path")?;
            if path_argument
                .to_str()
                .is_some_and(|value| value.starts_with("--"))
            {
                return Err(invalid_input("missing --inverse path"));
            }
            let path = PathBuf::from(path_argument);
            if path.as_os_str().is_empty() {
                return Err(invalid_input("inverse path must not be empty"));
            }
            options.inverse = Some(path);
        } else if argument == OsStr::new("--dry-run") {
            if options.dry_run {
                return Err(invalid_input("--dry-run may be specified only once"));
            }
            options.dry_run = true;
        } else {
            return Err(invalid_input("unexpected command-line argument"));
        }
    }
    Ok(options)
}

fn parse_row(value: String, label: &str) -> Result<usize, Box<dyn Error>> {
    value
        .parse()
        .map_err(|_| invalid_input(format!("{label} must be a non-negative integer")))
}

fn ensure_distinct_paths(
    first: &Path,
    second: &Path,
    message: &'static str,
) -> Result<(), Box<dyn Error>> {
    if first == second || canonical_path(first)? == canonical_path(second)? {
        return Err(invalid_input(message));
    }
    Ok(())
}

fn canonical_path(path: &Path) -> Result<PathBuf, Box<dyn Error>> {
    if path.exists() {
        return Ok(fs::canonicalize(path)?);
    }
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    Ok(fs::canonicalize(parent)?.join(
        path.file_name()
            .ok_or_else(|| invalid_input("output path must name a file"))?,
    ))
}

/// Publish through a sibling temporary file without replacing an existing
/// destination.  The package library's durable replacement API is not used
/// here because this example intentionally demonstrates noclobber behavior.
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
