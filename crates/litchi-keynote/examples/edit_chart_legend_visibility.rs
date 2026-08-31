//! Edit one chart's legend visibility through the focused Keynote API.
//!
//! Selection stays semantic (`index:N` or `name:NAME`); native object
//! identifiers never enter this command-line interface or its diagnostics.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{ChartSelector, Package, SlideSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_chart_legend_visibility <input.key> <output.key> \
                     <true|false> [index:N|name:SLIDE] [index:N|name:CHART]";

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

enum SelectedChart {
    Index(usize),
    Name(String),
}

impl SelectedChart {
    fn selector(&self) -> ChartSelector<'_> {
        match self {
            Self::Index(index) => ChartSelector::index(*index),
            Self::Name(name) => ChartSelector::name(name),
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let visible = parse_visible(required_argument(&mut arguments, "missing visibility")?)?;
    let slide = parse_slide(arguments.next())?;
    let chart = parse_chart(arguments.next())?;
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    package.validate()?;
    let slide_selector = slide.selector();
    let chart_selector = chart.selector();
    let before = package.slide_chart_legend_visible(slide_selector, chart_selector)?;
    let source_bytes = exact_bytes(&package)?;

    // A requested value equal to the source value must remain an exact no-op:
    // no physical rewrite, and no change to the serialized package bytes.
    let noop = package
        .edit_slide_chart_legend(slide.selector(), chart.selector())?
        .set(before)
        .commit()?;
    if noop.diagnostics().changed()
        || !noop.patch().is_noop()
        || exact_bytes(noop.package())? != source_bytes
        || noop
            .package()
            .slide_chart_legend_visible(slide.selector(), chart.selector())?
            != before
    {
        return Err(invalid_input(
            "setting the existing legend visibility was not an exact no-op",
        ));
    }

    let commit = package
        .edit_slide_chart_legend(slide.selector(), chart.selector())?
        .set(visible)
        .commit()?;
    let committed = commit
        .package()
        .slide_chart_legend_visible(slide.selector(), chart.selector())?;
    if committed != visible {
        return Err(invalid_input(
            "committed package did not expose the requested legend visibility",
        ));
    }

    // Apply the exact inverse in memory and require both semantic and byte
    // equality with the source package, including when the requested edit is
    // itself a no-op.
    let restored = commit
        .package()
        .apply_slide_chart_legend(&commit.patch().inverse())?;
    if restored
        .package()
        .slide_chart_legend_visible(slide.selector(), chart.selector())?
        != before
        || exact_bytes(restored.package())? != source_bytes
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and legend visibility",
        ));
    }

    save_new(&output, commit.package())?;
    let reopened = Package::open(&output)?;
    reopened.validate()?;
    let after = reopened.slide_chart_legend_visible(slide_selector, chart_selector)?;
    if after != visible {
        return Err(invalid_input(
            "reopened output did not preserve requested legend visibility",
        ));
    }

    println!(
        "chart legend: visible_before={before}, visible_after={after}, changed={}, touched_components={}, deleted_previews={}, full_reparse={}, source_fingerprint={:016x}, target_fingerprint={:016x}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
        commit.diagnostics().full_reparse_performed(),
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint(),
    );
    Ok(())
}

fn parse_slide(argument: Option<OsString>) -> Result<SelectedSlide, Box<dyn Error>> {
    parse_selector(argument, "slide", SelectedSlide::Index, SelectedSlide::Name)
}

fn parse_chart(argument: Option<OsString>) -> Result<SelectedChart, Box<dyn Error>> {
    parse_selector(argument, "chart", SelectedChart::Index, SelectedChart::Name)
}

fn parse_visible(argument: OsString) -> Result<bool, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("visibility must be valid UTF-8 (true or false)"))?;
    match value.as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(invalid_input("visibility must be exactly true or false")),
    }
}

fn parse_selector<T>(
    argument: Option<OsString>,
    kind: &str,
    from_index: impl FnOnce(usize) -> T,
    from_name: impl FnOnce(String) -> T,
) -> Result<T, Box<dyn Error>> {
    let Some(argument) = argument else {
        return Ok(from_index(0));
    };
    let value = argument
        .into_string()
        .map_err(|_| invalid_input(format!("{kind} selector must be valid UTF-8")))?;
    if let Some(index) = value.strip_prefix("index:") {
        let index = index
            .parse()
            .map_err(|_| invalid_input(format!("{kind} index must be a non-negative integer")))?;
        return Ok(from_index(index));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input(format!("{kind} name cannot be empty")));
        }
        return Ok(from_name(name.to_owned()));
    }
    Err(invalid_input(format!(
        "{kind} selector must start with index: or name:"
    )))
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

/// Publish through a sibling temporary file without overwriting an existing
/// artifact. A failed write leaves the source and destination untouched.
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
