//! Set or clear one chart caption through the focused Keynote API.
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

const USAGE: &str = "usage: edit_chart_caption <input.key> <output.key> \
                     <set:CAPTION|clear> [index:N|name:SLIDE] [index:N|name:CHART]";

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

enum CaptionChange {
    Set(String),
    Clear,
}

impl CaptionChange {
    fn expected(&self) -> Option<&str> {
        match self {
            Self::Set(caption) => Some(caption),
            Self::Clear => None,
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let change = parse_change(required_argument(&mut arguments, "missing caption change")?)?;
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
    let before = package.slide_chart_caption(slide_selector, chart_selector)?;
    let source_bytes = exact_bytes(&package)?;
    let expected = change.expected().map(str::to_owned);

    let edit = package.edit_slide_chart_caption(slide.selector(), chart.selector())?;
    let staged = match &change {
        CaptionChange::Set(caption) => edit.set(caption).map_err(|error| {
            io::Error::other(format!("staging chart-caption set failed: {error}"))
        })?,
        CaptionChange::Clear => edit.clear().map_err(|error| {
            io::Error::other(format!("staging chart-caption clear failed: {error}"))
        })?,
    };
    let commit = staged
        .commit()
        .map_err(|error| io::Error::other(format!("chart-caption commit failed: {error}")))?;
    let committed = commit
        .package()
        .slide_chart_caption(slide.selector(), chart.selector())
        .map_err(|error| {
            io::Error::other(format!("reading committed chart caption failed: {error}"))
        })?;
    if committed.as_deref() != expected.as_deref() {
        return Err(invalid_input(
            "committed package did not expose the requested chart caption",
        ));
    }
    if commit.diagnostics().changed() != (before != committed) {
        return Err(invalid_input(
            "chart-caption diagnostics disagreed with the semantic change",
        ));
    }

    // Verify the exact inverse in memory before publishing the new artifact.
    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())
        .map_err(|error| io::Error::other(format!("chart-caption inverse failed: {error}")))?;
    let restored_caption = restored
        .package()
        .slide_chart_caption(slide.selector(), chart.selector())
        .map_err(|error| {
            io::Error::other(format!("reading inverse chart caption failed: {error}"))
        })?;
    let restored_bytes = exact_bytes(restored.package()).map_err(|error| {
        io::Error::other(format!("serializing inverse package failed: {error}"))
    })?;
    if restored_caption != before || restored_bytes != source_bytes {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and chart caption",
        ));
    }

    save_new(&output, commit.package())
        .map_err(|error| io::Error::other(format!("writing output package failed: {error}")))?;
    let reopened = Package::open(&output)
        .map_err(|error| io::Error::other(format!("reopening output package failed: {error}")))?;
    reopened.validate().map_err(|error| {
        io::Error::other(format!("validating reopened package failed: {error}"))
    })?;
    let after = reopened
        .slide_chart_caption(slide_selector, chart_selector)
        .map_err(|error| {
            io::Error::other(format!("reading reopened chart caption failed: {error}"))
        })?;
    if after.as_deref() != expected.as_deref() {
        return Err(invalid_input(
            "reopened output did not preserve the requested chart caption",
        ));
    }

    println!(
        "chart caption: before_present={}, after_present={}, after_bytes={}, changed={}, touched_components={}, deleted_previews={}, full_reparse={}, source_fingerprint={:016x}, target_fingerprint={:016x}",
        before.is_some(),
        after.is_some(),
        after.as_deref().map_or(0, str::len),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
        commit.diagnostics().full_reparse_performed(),
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint(),
    );
    Ok(())
}

fn parse_change(argument: OsString) -> Result<CaptionChange, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("caption change must be valid UTF-8"))?;
    if let Some(caption) = value.strip_prefix("set:") {
        return Ok(CaptionChange::Set(caption.to_owned()));
    }
    if value == "clear" {
        return Ok(CaptionChange::Clear);
    }
    Err(invalid_input("caption change must be set:CAPTION or clear"))
}

fn parse_slide(argument: Option<OsString>) -> Result<SelectedSlide, Box<dyn Error>> {
    parse_selector(argument, "slide", SelectedSlide::Index, SelectedSlide::Name)
}

fn parse_chart(argument: Option<OsString>) -> Result<SelectedChart, Box<dyn Error>> {
    parse_selector(argument, "chart", SelectedChart::Index, SelectedChart::Name)
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
        if index.is_empty() || !index.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(invalid_input(format!(
                "{kind} index must be a non-negative decimal integer"
            )));
        }
        let index = index
            .parse::<usize>()
            .map_err(|_| invalid_input(format!("{kind} index is too large for this platform")))?;
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

/// Publish through a sibling temporary file without replacing an existing
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
