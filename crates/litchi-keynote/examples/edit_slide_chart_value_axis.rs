//! Edit one chart's primary value-axis settings through the focused Keynote API.
//!
//! The example deliberately keeps selection and values semantic. It never
//! accepts or prints a native object identifier, archive name, or wire value.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::collections::hash_map::DefaultHasher;
use std::error::Error;
use std::ffi::OsString;
use std::hash::{Hash, Hasher};
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::chart::axis::{
    Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
};
use litchi_keynote::{ChartSelector, Package, SlideSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_slide_chart_value_axis <input.key> <output.key> \
                     [index:N|name:SLIDE] [index:N|name:CHART]";

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
    let before = package.slide_chart_value_axis_settings(slide_selector, chart_selector)?;

    // All three controls are staged in one aggregate transaction. The
    // constructors validate finite bounds and native step ranges before the
    // package editor is reached; positive bounds also keep logarithmic scale
    // meaningful for native Keynote inspectors.
    let requested = ValueAxisSettings::automatic()
        .with_bounds(Bounds::fixed(Bound::new(1.0)?, Bound::new(120.0)?)?)
        .with_steps(Steps::fixed(
            MajorStepCount::new(6)?,
            MinorStepCount::new(2)?,
        ))
        .with_scale(Scale::Logarithmic);

    // Exercise the exact no-op path without writing over the source artifact.
    // A no-op retains the original package bytes and does not reassemble it.
    let noop = package
        .edit_slide_chart_value_axis_settings(slide.selector(), chart.selector())?
        .set(before)?
        .commit()?;
    if noop.diagnostics().changed() || exact_bytes(noop.package())? != exact_bytes(&package)? {
        return Err(invalid_input(
            "setting the existing value-axis settings was not an exact no-op",
        ));
    }

    let commit = package
        .edit_slide_chart_value_axis_settings(slide.selector(), chart.selector())?
        .set(requested)?
        .commit()?;

    // Validate the reversible patch in memory. The source file is never
    // overwritten; only the newly requested output is published below.
    let restored = commit
        .package()
        .apply_slide_chart_value_axis_settings(&commit.patch().inverse())?;
    if restored
        .package()
        .slide_chart_value_axis_settings(slide.selector(), chart.selector())?
        != before
        || exact_bytes(restored.package())? != exact_bytes(&package)?
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    save_new(&output, commit.package())?;
    let reopened = Package::open(&output)?;
    reopened.validate()?;
    let after = reopened.slide_chart_value_axis_settings(slide_selector, chart_selector)?;
    if after != requested {
        return Err(invalid_input(
            "reopened output did not preserve requested value-axis settings",
        ));
    }

    println!(
        "chart value axis: changed={}, touched_components={}, deleted_previews={}, before_fingerprint={:016x}, after_fingerprint={:016x}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
        semantic_fingerprint(before),
        semantic_fingerprint(after),
    );
    Ok(())
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

fn semantic_fingerprint(settings: ValueAxisSettings) -> u64 {
    let mut hasher = DefaultHasher::new();
    if let Some(minimum) = settings.bounds().minimum() {
        1_u8.hash(&mut hasher);
        minimum.value().to_bits().hash(&mut hasher);
    } else {
        0_u8.hash(&mut hasher);
    }
    if let Some(maximum) = settings.bounds().maximum() {
        1_u8.hash(&mut hasher);
        maximum.value().to_bits().hash(&mut hasher);
    } else {
        0_u8.hash(&mut hasher);
    }
    settings
        .steps()
        .major()
        .map(MajorStepCount::value)
        .hash(&mut hasher);
    settings
        .steps()
        .minor()
        .map(MinorStepCount::value)
        .hash(&mut hasher);
    // Hash the public semantic scale directly. Native discriminants are an
    // adapter concern and must not leak into this user-facing example.
    settings.scale().hash(&mut hasher);
    hasher.finish()
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
